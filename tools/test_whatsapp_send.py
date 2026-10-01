"""
One-off smoke test for phase 3 stage 1: sends ONE real WhatsApp message
via Twilio Sandbox to NOTIFY_TEST_WHATSAPP_NUMBER (.env) — meant to
verify the Twilio plumbing (app/whatsapp_sender.py) works before testing
against real cleaners. Run manually:

    python -m tools.test_whatsapp_send

Requires: TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, NOTIFY_TEST_WHATSAPP_NUMBER
in .env, and that number must have already sent the Sandbox "join <code>"
message from its own WhatsApp app (see docs/STATUS.md section 14).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv

load_dotenv()

from app import whatsapp_sender  # noqa: E402 — must load .env first

if __name__ == "__main__":
    to_number = os.environ.get("NOTIFY_TEST_WHATSAPP_NUMBER")
    if not to_number:
        print("NOTIFY_TEST_WHATSAPP_NUMBER not set in .env — nothing to send to.")
        raise SystemExit(1)

    sid = whatsapp_sender.send_cleaner_request(
        to_number, "Test", "01.01.2027", "Marktresidenz Karlstraße"
    )
    print(f"Gesendet (Content Template), message SID: {sid}")
