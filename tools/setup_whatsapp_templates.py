"""
One-off: creates the German WhatsApp Content Templates for the cleaner
coordination flow (besides the first-contact request template, which was
created earlier) and prints their ContentSids to paste into
app/whatsapp_sender.py / .env. Run manually, once:

    python -m tools.setup_whatsapp_templates

Templates are needed for every business-initiated message (see
app/whatsapp_sender.py docstring). In the Sandbox they work without Meta
review; a production sender needs each one approved (Twilio Console ->
Content Template Builder -> "Request WhatsApp approval").
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv

load_dotenv()

from twilio.rest import Client  # noqa: E402
from twilio.rest.content.v1.content import ContentList  # noqa: E402

BUTTONS = [("Ja", "ja"), ("Nein", "nein")]
BUTTONS_REQUEST = [("Ja", "ja"), ("Nein", "nein"), ("Vielleicht", "vielleicht")]

TEMPLATES = {
    # name: (body, variables, buttons)
    "putzplan_request_guests": (
        "Hallo {{1}}, am {{2}} wird eine Reinigung gebraucht ({{3}}). Nächste Gäste: {{4}}. {{5}} "
        "Kannst du das übernehmen?",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße", "4": "2 Erwachsene, 2 Kinder unter 18",
         "5": "Kinder unter 3 Jahren: keine Angabe."}, "request"),
    "putzplan_urgent_guests": (
        "Hallo {{1}}, für die Reinigung am {{2}} ({{3}}) ist noch niemand verfügbar. "
        "Nächste Gäste: {{4}}. {{5}} Kannst du sie bitte übernehmen?",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße", "4": "2 Erwachsene, 2 Kinder unter 18",
         "5": "Kinder unter 3 Jahren: keine Angabe."}, True),
    "putzplan_reminder_open_guests": (
        "Hallo {{1}}, für die Reinigung am {{2}} ({{3}}) haben wir noch niemanden gefunden. "
        "Nächste Gäste: {{4}}. {{5}} Wenn du kannst, antworte bitte bald.",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße", "4": "2 Erwachsene, 2 Kinder unter 18",
         "5": "Kinder unter 3 Jahren: keine Angabe."}, True),
    "putzplan_reminder_open": (
        "Hallo {{1}}, für die Reinigung am {{2}} ({{3}}) haben wir noch niemanden gefunden. "
        "Wenn du kannst, antworte bitte bald.",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, True),
    "putzplan_urgent": (
        "Hallo {{1}}, dringend: Für die Reinigung am {{2}} ({{3}}) suchen wir noch jemanden. "
        "Kannst du übernehmen?",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, True),
    "putzplan_taken": (
        "Hallo {{1}}, die Reinigung am {{2}} ({{3}}) ist inzwischen vergeben. Danke dir trotzdem!",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, False),
    "putzplan_promoted": (
        "Hallo {{1}}, gute Nachricht: Der Platz ist frei geworden — du bist jetzt für die Reinigung "
        "am {{2}} ({{3}}) bestätigt. Danke dir!",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, False),
    "putzplan_cancelled": (
        "Hallo {{1}}, leider wurde die Reservierung für den {{2}} ({{3}}) storniert. "
        "Die Reinigung entfällt, du musst nicht kommen. Danke dir!",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, False),
    "putzplan_reminder_tomorrow_v2": (
        "Hallo {{1}}, kurze Erinnerung: Morgen ({{2}}) ist die Reinigung in {{3}}. "
        "Nächste Gäste: {{4}}. {{5}} Bis dann!",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße", "4": "2 Erwachsene, 1 Kind unter 18",
         "5": "Kinder unter 3 Jahren: nein."}, False),
    "putzplan_again": (
        "Hallo {{1}}, doch noch: Am {{2}} ({{3}}) wird die Reinigung wieder gebraucht — "
        "es gibt eine neue Buchung. Kannst du?",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, True),
    # v2: Meta moved putzplan_urgent from UTILITY to MARKETING -> neutral wording, no "dringend"
    "putzplan_urgent_v2": (
        "Hallo {{1}}, für die Reinigung am {{2}} ({{3}}) ist noch niemand verfügbar. "
        "Kannst du sie bitte übernehmen?",
        {"1": "Name", "2": "01.01.2027", "3": "Karlstraße"}, True),
    # v2: Meta rejected v1 ("variables can't be at the start or end of the template")
    "putzplan_owner_notice_v2": (
        "Hinweis: {{1}} — Viele Grüße, Ihr Putzplan-Automat",
        {"1": "Beispieltext"}, False),
    "putzplan_owner_notice": (
        "Hinweis vom Putzplan-Automaten: {{1}}",
        {"1": "Beispieltext"}, False),
}


def main():
    client = Client(os.environ["TWILIO_ACCOUNT_SID"], os.environ["TWILIO_AUTH_TOKEN"])
    only = sys.argv[1:]  # optionally create just some templates: python -m tools.setup_whatsapp_templates putzplan_promoted
    for name, (body, variables, with_buttons) in TEMPLATES.items():
        if only and name not in only:
            continue
        if with_buttons:
            buttons = BUTTONS_REQUEST if with_buttons == "request" else BUTTONS
            types = {"twilio/quick-reply": {"body": body, "actions": [{"title": t, "id": i} for t, i in buttons]}}
        else:
            types = {"twilio/text": {"body": body}}
        req = ContentList.ContentCreateRequest({
            "friendly_name": name, "language": "de", "variables": variables, "types": types})
        content = client.content.v1.contents.create(content_create_request=req)
        print(f"{name}: {content.sid}")


if __name__ == "__main__":
    main()
