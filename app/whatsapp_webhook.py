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

from flask import Blueprint, request, Response
from twilio.request_validator import RequestValidator
from twilio.twiml.messaging_response import MessagingResponse

from . import cleaner_coordination as cc
from .cleaner_roster import load_roster
from .config import PROPERTY_LABELS
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
            # Not a number we recognize (could be Farzaneh's own test replies
            # from earlier, or a wrong number) — nothing to do, but still 200
            # so Twilio doesn't retry.
            return _twiml()
        if response is None:
            # Free text, voice message, photo, ... — the automation can't read
            # those. Forward what we can to the owner (text; for media only a
            # note, the file itself can't go through a template), with the
            # sender's number so she can answer directly.
            body = (request.form.get("Body") or "").strip()
            media = int(request.form.get("NumMedia") or 0)
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
        return _twiml(f"Danke, ist bestätigt: {where}")
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
