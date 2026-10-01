"""
Owner notifications over WhatsApp (phase 3, 2026-10-01): a short line for
every cleaner message sent / answer received (so Farzaneh can see the
automation working without opening the Twilio Console), and the real
"alert_owner" cases (nobody eligible, checkout passed unresolved...).

Goes to OWNER_WHATSAPP_NUMBER (.env, E.164). Uses the owner-notice
Content Template, since these are business-initiated messages. Never
raises — a failed notice must not break the coordination run itself.
"""
import os

from . import whatsapp_sender


def notify_owner(text: str, log=print) -> bool:
    # WhatsApp template variables can't contain newlines/tabs or long runs of spaces
    text = " / ".join(part.strip() for part in text.replace("	", " ").splitlines() if part.strip())[:900]
    number = os.environ.get("OWNER_WHATSAPP_NUMBER")
    if not number:
        log(f"    [owner-alert] OWNER_WHATSAPP_NUMBER nicht gesetzt — übersprungen: {text}")
        return False
    try:
        whatsapp_sender.send_template(number, "owner_notice", {"1": text})
    except Exception as exc:  # noqa: BLE001
        log(f"    [owner-alert] WARN — Nachricht an Inhaberin fehlgeschlagen: {exc}")
        return False
    return True
