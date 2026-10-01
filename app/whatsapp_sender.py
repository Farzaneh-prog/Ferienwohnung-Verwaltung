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

# Other German templates (created by tools/setup_whatsapp_templates.py,
# 2026-10-01). Same story as above: SIDs aren't secret, env can override.
TEMPLATE_SIDS = {
    "reminder_open": os.environ.get("TWILIO_TEMPLATE_REMINDER_OPEN", "HX96afc7964e8d69ad98ed4718f59d190c"),
    "urgent": os.environ.get("TWILIO_TEMPLATE_URGENT", "HXb5aad2c56e82bc85cd776ae00921c63e"),
    "taken": os.environ.get("TWILIO_TEMPLATE_TAKEN", "HX5abcf216022453b9e3c49c34c64e7294"),
    "promoted": os.environ.get("TWILIO_TEMPLATE_PROMOTED", "HXfc63f7d3907315836ef1338384e7665a"),
    "cancelled": os.environ.get("TWILIO_TEMPLATE_CANCELLED", "HX7b83ed2555eb8c5ac2b480b3126045a7"),
    "reminder_tomorrow": os.environ.get("TWILIO_TEMPLATE_REMINDER_TOMORROW", "HX9d7447ca4372c3ec648ede6e3dded71e"),
    "again": os.environ.get("TWILIO_TEMPLATE_AGAIN", "HXb5e8f2515c41bdd68e5a5b7347fd4f93"),
    "owner_notice": os.environ.get("TWILIO_TEMPLATE_OWNER_NOTICE", "HX825046c3748afed5a8097655a44467ac"),
}

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


def send_template(to_number: str, template: str, variables: dict) -> str:
    """Sends one of TEMPLATE_SIDS' templates (key in `template`) with
    `variables` ({"1": ..., "2": ...}). Raises on failure, like the
    other senders here."""
    import json

    message = _get_client().messages.create(
        from_=f"whatsapp:{TWILIO_WHATSAPP_FROM}",
        to=f"whatsapp:{to_number}",
        content_sid=TEMPLATE_SIDS[template],
        content_variables=json.dumps(variables),
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
