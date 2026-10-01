"""
Thin wrapper around Twilio's WhatsApp API (phase 3 stage 1, 2026-09-30).

Credentials come ONLY from .env (TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN) —
never hardcode them, or any phone number, anywhere in this file or the
rest of the repo. Real WhatsApp numbers live in data/cleaners.json, which
is gitignored (see cleaner_roster.py).

TWILIO_WHATSAPP_FROM defaults to Twilio's WhatsApp Sandbox number, which
is the same fixed number for every trial account (+14155238886) — set it
in .env to override once a dedicated production number is registered
(see docs/STATUS.md section 14 for the stage plan).

Content Templates, not free text (2026-10-01 discovery): WhatsApp only
allows free-form message bodies as a REPLY inside an already-open 24h
conversation window (e.g. the recipient messaged you recently). Any
business-initiated message — which is exactly what the cleaner-
coordination "10 days before" first contact is — needs a pre-approved
Content Template (Twilio's Content API), referenced by ContentSid, with
named variables filled in via content_variables. Twilio's Content API
itself also requires an APPROVED Trust Hub primary compliance profile to
even create a template — see docs/STATUS.md section 14 for how that got
resolved (Individual profile, auto-approved; Business profile was
rejected — US-centric checks like EIN don't fit a German GbR).
TWILIO_CLEANER_TEMPLATE_SID is the cleaner-request template created that
way — a twilio/quick-reply template (2026-10-01, after Farzaneh's first
real test: a plain-text template only let her type "Ja" freehand, so it
was rebuilt with three tappable buttons, Ja/Nein/Vielleicht, instead —
also easier for the Stage-1-next reply webhook to parse reliably than
free text). Its SID isn't a secret (unlike the account credentials
above), but is still only ever read from .env/this constant, never
guessed at.
"""
import os

from twilio.rest import Client

TWILIO_ACCOUNT_SID = os.environ.get("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.environ.get("TWILIO_AUTH_TOKEN")
TWILIO_WHATSAPP_FROM = os.environ.get("TWILIO_WHATSAPP_FROM", "+14155238886")
TWILIO_CLEANER_TEMPLATE_SID = os.environ.get(
    "TWILIO_CLEANER_TEMPLATE_SID", "HX8953397e888345beaba2f36eb5f150d9"
)

_client = None


def _get_client() -> Client:
    global _client
    if not TWILIO_ACCOUNT_SID or not TWILIO_AUTH_TOKEN:
        raise RuntimeError("TWILIO_ACCOUNT_SID/TWILIO_AUTH_TOKEN not set in .env")
    if _client is None:
        _client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
    return _client


def send_cleaner_request(to_number: str, name: str, date_str: str, property_label: str) -> str:
    """Sends the cleaner-coordination request template (variables {{1}}
    name, {{2}} date, {{3}} property — see tools/setup_whatsapp_template.py
    for how TWILIO_CLEANER_TEMPLATE_SID was created). Works cold, with no
    prior conversation needed, unlike send_whatsapp_freeform below.
    to_number: E.164 format, e.g. '+491511234567' (no 'whatsapp:' prefix —
    added here). Returns the Twilio message SID; raises on any failure —
    callers must catch this (same defensive pattern as
    checkin_reminder.send_reminder_email)."""
    import json

    client = _get_client()
    message = client.messages.create(
        from_=f"whatsapp:{TWILIO_WHATSAPP_FROM}",
        to=f"whatsapp:{to_number}",
        content_sid=TWILIO_CLEANER_TEMPLATE_SID,
        content_variables=json.dumps({"1": name, "2": date_str, "3": property_label}),
    )
    return message.sid


def send_whatsapp_freeform(to_number: str, body: str) -> str:
    """Free-form text — only deliverable if the recipient has messaged
    this WhatsApp number within the last 24 hours (WhatsApp's session
    window). Not used by cleaner_coordination.py's cold first-contact
    messages; kept for future use (e.g. a reply within an open session)."""
    client = _get_client()
    message = client.messages.create(
        from_=f"whatsapp:{TWILIO_WHATSAPP_FROM}",
        to=f"whatsapp:{to_number}",
        body=body,
    )
    return message.sid
