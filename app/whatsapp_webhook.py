"""
Receives Twilio's incoming-WhatsApp-message webhook (phase 3 stage 1,
2026-10-01) — this is the other half of cleaner_coordination.py: that
module only SENDS requests; this is what turns a cleaner's "Ja" reply
into an actual Putzplan column-A assignment.

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
"""
import os

from flask import Blueprint, request, Response
from twilio.request_validator import RequestValidator

from .cleaner_coordination import find_pending_row_for_cleaner, confirm_cleaner, _load_state
from .cleaner_roster import load_roster

bp = Blueprint("whatsapp_webhook", __name__)

_EMPTY_TWIML = Response('<?xml version="1.0" encoding="UTF-8"?><Response></Response>', mimetype="text/xml")


def _find_cleaner_id_by_number(from_number: str, roster: dict):
    for cleaner_id, cleaner in roster.items():
        if cleaner.get("whatsapp_number") == from_number:
            return cleaner_id
    return None


@bp.route("/webhooks/whatsapp", methods=["POST"])
def incoming_whatsapp():
    auth_token = os.environ.get("TWILIO_AUTH_TOKEN")
    validator = RequestValidator(auth_token)
    signature = request.headers.get("X-Twilio-Signature", "")
    # request.url reflects the scheme/host the app itself sees — behind
    # the QNAP reverse proxy that's already correctly https://..., same
    # URL Twilio was configured to call (no X-Forwarded-Proto juggling
    # needed here, unlike some reverse-proxy setups).
    if not validator.validate(request.url, request.form, signature):
        return Response(status=403)

    from_raw = request.form.get("From", "")  # e.g. "whatsapp:+491511234567"
    from_number = from_raw.replace("whatsapp:", "").strip()
    button_payload = (request.form.get("ButtonPayload") or "").strip().lower()
    body_text = (request.form.get("Body") or "").strip().lower()
    reply = button_payload or body_text

    roster = load_roster()
    cleaner_id = _find_cleaner_id_by_number(from_number, roster)
    if cleaner_id is None:
        # Not a number we recognize (could be Farzaneh's own test replies
        # from earlier, or a wrong number) — nothing to do, but still 200
        # so Twilio doesn't retry.
        return _EMPTY_TWIML

    if reply in ("ja", "yes"):
        state = _load_state()
        pending = find_pending_row_for_cleaner(cleaner_id, state)
        if pending is not None:
            property_key, checkout_date = pending
            confirm_cleaner(property_key, checkout_date, cleaner_id)
    # "nein"/"vielleicht"/anything else: no action needed — the cleaner
    # simply isn't marked confirmed, so the normal escalation (next
    # scheduled run_daily_check) moves on to whoever's next in the chain.

    return _EMPTY_TWIML
