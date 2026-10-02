"""
Owner notifications (phase 3, 2026-10-01; reworked same day after Meta put the
generic owner-notice template into the MARKETING category).

Three channels, by importance:
- IMPORTANT (text starts with "⚠️", or important=True): nobody available, everyone
  declined, wrong clock, contract ending, ... -> WhatsApp via the owner-notice
  Content Template AND an immediate e-mail. Only a few per month, so the
  template's price doesn't matter, and it never depends on a 24h window.
- ROUTINE (every request sent / answer received / reminder): collected in a
  daily digest e-mail (18:00, scheduler.py). Additionally, while the owner's
  24h WhatsApp window is open — the owner messaged the business number within
  the last 24h, see whatsapp_webhook — each line is also sent live as free
  text. Nothing breaks if the owner never says hello; the digest has it all.

The window can only be opened by the owner writing first (WhatsApp rule), it
cannot be automated.

State: data/owner_digest.json {"window_until", "entries": [...], "digested_until"}.
Never raises — a failed notice must not break the coordination run.
"""
import datetime
import json
import os
import smtplib
import threading
from email.mime.text import MIMEText

from .config import DATA_DIR
from .tz import BERLIN
from . import whatsapp_sender

_PATH = os.path.join(DATA_DIR, "data", "owner_digest.json")
_LOCK = threading.RLock()
WINDOW = datetime.timedelta(hours=24)
KEEP_DAYS = 7


def _now() -> datetime.datetime:
    return datetime.datetime.now(tz=BERLIN)


def _load() -> dict:
    if not os.path.exists(_PATH):
        return {"window_until": None, "entries": [], "digested_until": None}
    with open(_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(data: dict) -> None:
    os.makedirs(os.path.dirname(_PATH), exist_ok=True)
    tmp = f"{_PATH}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, _PATH)


def _clean(text: str) -> str:
    """WhatsApp template variables can't contain newlines/tabs or long runs of spaces."""
    return " / ".join(part.strip() for part in text.replace("\t", " ").splitlines() if part.strip())[:900]


def open_window(now: datetime.datetime = None) -> None:
    """The owner wrote to the business number: free-text updates are allowed
    for the next 24 hours."""
    now = now or _now()
    with _LOCK:
        data = _load()
        data["window_until"] = (now + WINDOW).isoformat()
        _save(data)


def window_open(now: datetime.datetime = None) -> bool:
    now = now or _now()
    until = _load().get("window_until")
    return bool(until) and now < datetime.datetime.fromisoformat(until)


def _recipient_email():
    return (os.environ.get("OWNER_DIGEST_EMAIL") or os.environ.get("NOTIFY_TO_EMAIL")
            or os.environ.get("GMAIL_USER"))


def _send_mail(subject: str, body: str, log=print) -> bool:
    user, password, to = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASSWORD"), _recipient_email()
    if not (user and password and to):
        log("    [owner-alert] GMAIL_USER/GMAIL_APP_PASSWORD/NOTIFY_TO_EMAIL fehlen — E-Mail übersprungen")
        return False
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"], msg["From"], msg["To"] = subject, user, to
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(user, password)
            server.sendmail(user, [to], msg.as_string())
    except Exception as exc:  # noqa: BLE001
        log(f"    [owner-alert] WARN — E-Mail fehlgeschlagen: {exc}")
        return False
    return True


def notify_owner(text: str, important: bool = False, log=print) -> bool:
    """Returns True if the notice was delivered through at least one durable
    channel (important: WhatsApp template or e-mail; routine: always — it is in
    the digest)."""
    text = _clean(text)
    important = important or text.startswith("⚠️")
    now = _now()
    with _LOCK:
        data = _load()
        data["entries"].append({"at": now.isoformat(), "text": text, "important": important})
        _save(data)

    if not important:
        if window_open(now):
            number = os.environ.get("OWNER_WHATSAPP_NUMBER")
            if number:
                try:
                    whatsapp_sender.send_whatsapp_freeform(number, text)
                except Exception as exc:  # noqa: BLE001
                    log(f"    [owner-alert] WARN — Live-Zeile fehlgeschlagen (steht im Tages-Digest): {exc}")
        return True

    delivered = False
    number = os.environ.get("OWNER_WHATSAPP_NUMBER")
    if number:
        try:
            whatsapp_sender.send_template(number, "owner_notice", {"1": text})
            delivered = True
        except Exception as exc:  # noqa: BLE001
            log(f"    [owner-alert] WARN — WhatsApp-Warnung fehlgeschlagen: {exc}")
    else:
        log(f"    [owner-alert] OWNER_WHATSAPP_NUMBER nicht gesetzt: {text}")
    if _send_mail("⚠️ Putzplan-Automat: " + text.lstrip("⚠️ ")[:80], text, log=log):
        delivered = True
    return delivered


def send_daily_digest(now: datetime.datetime = None, log=print) -> bool:
    """18:00: one e-mail with everything since the last digest. Sent only if
    there is something to report."""
    now = now or _now()
    with _LOCK:
        data = _load()
        since = data.get("digested_until")
        entries = [e for e in data["entries"] if not since or e["at"] > since]
        if not entries:
            log("    [owner-alert] Tages-Digest: nichts zu berichten — keine E-Mail")
            return False
        lines = []
        for e in entries:
            at = datetime.datetime.fromisoformat(e["at"]).astimezone(BERLIN)
            lines.append(f"{at:%d.%m. %H:%M}  {'!! ' if e['important'] else ''}{e['text']}")
        body = (f"Putzplan-Automat — Zusammenfassung bis {now:%d.%m.%Y %H:%M}\n\n" + "\n".join(lines)
                + "\n\n(Für Live-Updates in WhatsApp: schreib der Marktresidenz-Nummer einfach «Hallo» — "
                  "dann bekommst du 24 Stunden lang jede Zeile direkt.)")
        if not _send_mail(f"Putzplan-Tageszusammenfassung {now:%d.%m.%Y}", body, log=log):
            return False
        data["digested_until"] = now.isoformat()
        cutoff = (now - datetime.timedelta(days=KEEP_DAYS)).isoformat()
        data["entries"] = [e for e in data["entries"] if e["at"] > cutoff]
        _save(data)
        return True
