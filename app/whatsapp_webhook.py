"""
Receives Twilio's incoming-WhatsApp-message webhook (phase 3 stage 1,
2026-10-01) — this is the other half of cleaner_coordination.py: that
module only SENDS requests; this is what turns a cleaner's answer
(Ja / Nein / Vielleicht buttons) into state changes, a short German reply
to the cleaner, and — for Nein/Vielleicht — an immediate request to the
next person in the chain.

Must be reachable from the public internet for Twilio to call it — set
this route's full URL (https://.../webhooks/whatsapp) as the WhatsApp
Sandbox's "WHEN A MESSAGE COMES IN" webhook in the Twilio Console
(Messaging -> Try it out -> WhatsApp -> Sandbox settings). Only works
once this is deployed to the NAS (publicly reachable) — a local
`python run.py` can't receive it.

Security: every request is verified against Twilio's X-Twilio-Signature
header (RequestValidator, using TWILIO_AUTH_TOKEN) before anything in the
body is trusted — an unsigned/forged POST here could otherwise assign a
fake cleaner to a real Putzplan row.

Replies to the cleaner are plain TwiML <Message>s: that's free text sent
inside the 24h session the cleaner just opened by answering, so no
Content Template is needed for them.
"""
import datetime
import os
import re

from flask import Blueprint, request, Response
from twilio.request_validator import RequestValidator
from twilio.twiml.messaging_response import MessagingResponse

from . import cleaner_coordination as cc
from .cleaner_roster import load_roster
from .config import PROPERTY_LABELS
from . import owner_alerts
from .owner_alerts import notify_owner
from .tz import BERLIN

bp = Blueprint("whatsapp_webhook", __name__)

_EMPTY_TWIML = '<?xml version="1.0" encoding="UTF-8"?><Response></Response>'


def _twiml(text: str = None) -> Response:
    if not text:
        return Response(_EMPTY_TWIML, mimetype="text/xml")
    resp = MessagingResponse()
    resp.message(text)
    return Response(str(resp), mimetype="text/xml")


def _find_cleaner_ids_by_number(from_number: str, roster: dict) -> list:
    return [cid for cid, c in roster.items() if c.get("whatsapp_number") == from_number]


def _pick_cleaner_id(candidate_ids: list, state: dict, today: datetime.date):
    """Normally one roster entry per number. If several share a number
    (only in our own tests with Test1/Test2), pick the one with an open
    (unanswered) request, else the one most recently involved, else the first."""
    best = None
    for cid in candidate_ids:
        found = cc.find_row_for_reply(cid, state, today)
        if found is None:
            continue
        row_state = state["rows"][cc._state_key(*found)]
        mine = [a for a in row_state["attempts"] if a["cleaner_id"] == cid]
        rank = (mine[-1].get("response") is None, max(a.get("responded_at") or a["sent_at"] for a in mine))
        if best is None or rank > best[0]:
            best = (rank, cid)
    if best:
        return best[1]
    return candidate_ids[0] if candidate_ids else None


def _classify(*candidates: str):
    """Maps a button payload / typed text to "ja" | "nein" | "vielleicht",
    or None if it's none of those."""
    for text in candidates:
        t = (text or "").strip().lower()
        if t in ("ja", "yes"):
            return "ja"
        if t in ("nein", "no"):
            return "nein"
        if t == "vielleicht":
            return "vielleicht"
    return None


# Plain acknowledgements (German only — decided with Farzaneh 2026-10-08).
# A message counts only if EVERY word is in one of the sets below and at least
# one word is a CORE word; anything unknown ("Schlüssel", "krank", "später" on
# its own, "nicht", "kein", "ja", numbers, "?") makes it a real message that is
# forwarded to the owner as before.
_ACK_PHRASES = {  # multi-word forms that would otherwise contain unknown words
    "kein problem": "ok", "bis dann": "ok", "bis morgen": "ok", "bis bald": "ok", "bis später": "ok",
    "mache ich": "ok", "mach ich": "ok", "wird gemacht": "ok", "geht klar": "ok", "passt so": "ok",
}
_ACK_CORE = {
    # thanks
    "danke", "dankeschön", "dankeschoen", "dank", "merci", "thx",
    # understood / fine
    "ok", "okay", "oke", "klar", "gut", "bestens", "verstanden", "super", "prima", "perfekt", "passt",
    "ordnung", "sicher", "selbstverständlich", "natürlich",
    # informed / done
    "kenntnis", "genommen", "gelesen", "notiert", "bescheid", "informiert", "gemacht", "erledigt",
    # agreeing
    "gerne",
}
_ACK_FILLER = {
    "alles", "vielen", "herzlichen", "besten", "sehr", "schön", "schoen", "dir", "ihnen", "euch", "dich",
    "hallo", "hi", "hey", "moin", "guten", "tag", "morgen", "abend", "liebe", "lieber", "grüße", "grüsse",
    "gruß", "gruss", "viele", "schöne", "schönen", "feierabend", "lg", "mfg", "ciao",
    "ich", "habe", "es", "zur", "das", "ist", "in", "für", "die", "der", "info", "nachricht", "erinnerung",
    "wird", "weiß", "weiss", "bin", "so", "auch", "schon", "mal", "so",
}
_ACK_EMOJI = ("👍", "👌", "🙏", "✅", "😊", "🙂", "😀", "😃", "😄", "❤️", "❤", "💪", "🤝", "👏", "🙌")


def is_acknowledgement(text: str) -> bool:
    """True for a plain "thanks / ok / 👍 / zur Kenntnis genommen" — nothing
    the owner needs to act on, so no automatic "please use the buttons" reply
    and no 💬 forward (only a quiet line in the digest). Deliberately strict —
    see the word sets above."""
    t = (text or "").strip().lower()
    if not t or len(t) > 100 or "?" in t or any(ch.isdigit() for ch in t):
        return False
    for emoji in _ACK_EMOJI:
        t = t.replace(emoji, " ok ")
    t = re.sub(r"\s+", " ", t)
    for phrase, replacement in _ACK_PHRASES.items():
        t = t.replace(phrase, replacement)
    tokens = re.findall(r"[a-zäöüß]+", t)
    if not tokens or not any(tok in _ACK_CORE for tok in tokens):
        return False
    return all(tok in _ACK_CORE or tok in _ACK_FILLER for tok in tokens)


def _fmt(property_key: str, checkout_date: datetime.date) -> str:
    return f"{checkout_date.strftime('%d.%m.%Y')}, {PROPERTY_LABELS.get(property_key, property_key)}"


@bp.route("/webhooks/whatsapp", methods=["POST"])
def incoming_whatsapp():
    auth_token = os.environ.get("TWILIO_AUTH_TOKEN")
    validator = RequestValidator(auth_token)
    signature = request.headers.get("X-Twilio-Signature", "")
    # Twilio signs the exact public URL it called. Behind the QNAP reverse
    # proxy request.url is the internal http://localhost:8000/... one, so
    # the signature never matches — validate against the configured public
    # URL instead (TWILIO_WEBHOOK_URL), falling back to request.url.
    public_url = os.environ.get("TWILIO_WEBHOOK_URL") or request.url
    if not validator.validate(public_url, request.form, signature):
        return Response(status=403)

    from_raw = request.form.get("From", "")  # e.g. "whatsapp:+491511234567"
    from_number = from_raw.replace("whatsapp:", "").strip()
    response = _classify(request.form.get("ButtonPayload"), request.form.get("Body"))

    roster = load_roster()
    now = datetime.datetime.now(tz=BERLIN)
    with cc.STATE_LOCK:
        cleaner_id = _pick_cleaner_id(_find_cleaner_ids_by_number(from_number, roster), cc._load_state(), now.date())
        if cleaner_id is None:
            owner_number = os.environ.get("OWNER_WHATSAPP_NUMBER")
            if owner_number and from_number == owner_number:
                # The owner said hello: free-text live updates for the next 24h
                # (see owner_alerts). Important warnings never depend on this.
                owner_alerts.open_window(now)
                return _twiml("✓ Live-Updates sind jetzt 24 Stunden lang aktiv. "
                              "Wichtige Warnungen bekommst du immer, auch ohne Hallo.")
            # Not a number we recognize (a wrong number) — nothing to do, but
            # still 200 so Twilio doesn't retry.
            return _twiml()
        if response is None:
            # Free text, voice message, photo, ... — the automation can't read
            # those. Forward what we can to the owner (text; for media only a
            # note, the file itself can't go through a template), with the
            # sender's number so she can answer directly.
            body = (request.form.get("Body") or "").strip()
            media = int(request.form.get("NumMedia") or 0)
            if not media and is_acknowledgement(body):
                # a plain "danke / ok / 👍": log it quietly, no reply, no forward
                notify_owner(f"👍 {roster[cleaner_id]['name']}: „{body}“")
                return _twiml()
            note = f"„{body}“" if body else ""
            if media:
                note = (note + " " if note else "") + f"[{media} Anhang/Sprachnachricht — im Automaten nicht lesbar]"
            notify_owner(f"💬 Nachricht von {roster[cleaner_id]['name']} ({from_number}): {note or '(leer)'}")
            contact = os.environ.get("OWNER_CONTACT_NUMBER") or os.environ.get("OWNER_WHATSAPP_NUMBER") or ""
            where = f" direkt an Farzaneh: https://wa.me/{contact.lstrip('+')} ({contact})" if contact else " direkt an Farzaneh"
            return _twiml("Hinweis: Diese Nummer wird automatisch ausgewertet und Farzaneh sieht Nachrichten "
                          "hier nicht direkt. Text leite ich an sie weiter, Sprachnachrichten kann ich nicht "
                          f"weiterleiten — bitte schick diese{where}. "
                          "Für Anfragen nutze bitte die Buttons Ja / Nein / Vielleicht.")

        result = cc.apply_response(cleaner_id, response, now=now, roster=roster)

    kind = result["kind"]
    if kind == "no_row":
        return _twiml("Aktuell ist keine Anfrage für dich offen.")
    where = _fmt(result["property_key"], result["checkout_date"])
    if kind == "cancelled":
        return _twiml(f"Die Reinigung am {where} entfällt leider, die Reservierung wurde storniert.")
    if kind == "confirmed":
        from .day_before import guests_text

        guests = cc.guests_for_row(result["property_key"], result["checkout_date"])
        extra = ""
        if guests.get("adults") not in (None, ""):
            who, under3 = guests_text(guests)
            extra = f"\nNächste Gäste: {who}. {under3}"
        return _twiml(f"Danke, ist bestätigt: {where}{extra}")
    if kind == "still_confirmed":
        return _twiml(f"Du bist weiterhin bestätigt: {where}")
    if kind == "reserve":
        return _twiml(f"Leider schon vergeben ({where}) — du stehst aber als Reserve auf der Liste "
                      f"(Platz {result['position']}). Falls die Person absagt, melde ich mich sofort.")
    if kind in ("withdrawn_promoted", "withdrawn_reopened"):
        return _twiml(f"Alles klar, ich habe deine Bestätigung für {where} zurückgenommen und kümmere mich "
                      f"um Ersatz. Danke fürs Bescheid geben!")
    if kind == "declined":
        return _twiml("Alles klar, danke für die Antwort.")
    return _twiml("Danke! Bitte antworte so bald wie möglich mit Ja oder Nein. "
                  "In der Zwischenzeit frage ich schon mal die anderen.")
