"""
One-off: submits our Content Templates to WhatsApp for approval (needed for
every business-initiated message through a real, non-Sandbox sender). Run
manually after the sender is registered:

    python -m tools.submit_whatsapp_templates            # all
    python -m tools.submit_whatsapp_templates --status   # only show the current status

Approved templates can't be edited — a change means a new template (new name
or "_v2") and a new SID in app/whatsapp_sender.py.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv

load_dotenv(".env")

from twilio.rest import Client  # noqa: E402
from twilio.rest.content.v1.content.approval_create import ApprovalCreateList  # noqa: E402

# friendly_name on the account -> SID (see tools/setup_whatsapp_templates.py)
TEMPLATES = {
    "cleaner_coordination_request_quickreply_de": "HX8953397e888345beaba2f36eb5f150d9",
    "putzplan_request_guests": "HXeb5e02f1420de73e3e83e981b8b7545f",
    "putzplan_urgent_guests": "HX59f9c6451dcf72891cf38795eee6e50e",
    "putzplan_reminder_open_guests": "HX9ec00bd264543913a7df79d430ac7840",
    "putzplan_reminder_open": "HX96afc7964e8d69ad98ed4718f59d190c",
    "putzplan_urgent_v2": "HX00f3a39268ae0fae30e4b5650faa168f",  # v1 was reclassified as MARKETING by Meta
    "putzplan_taken": "HX5abcf216022453b9e3c49c34c64e7294",
    "putzplan_promoted": "HXfc63f7d3907315836ef1338384e7665a",
    "putzplan_cancelled": "HX7b83ed2555eb8c5ac2b480b3126045a7",
    "putzplan_reminder_tomorrow_v2": "HX9d7447ca4372c3ec648ede6e3dded71e",
    "putzplan_again": "HXb5e8f2515c41bdd68e5a5b7347fd4f93",
    "putzplan_owner_notice_v2": "HX7d488d852c44b1e006c11553716158a4",  # v1 was rejected by Meta (variable at the end)
}


def main():
    client = Client(os.environ["TWILIO_ACCOUNT_SID"], os.environ["TWILIO_AUTH_TOKEN"])
    status_only = "--status" in sys.argv
    for name, sid in TEMPLATES.items():
        if not status_only:
            try:
                client.content.v1.contents(sid).approval_create.create(
                    ApprovalCreateList.ContentApprovalRequest({"name": name, "category": "UTILITY"}))
                print(f"submitted: {name}")
            except Exception as exc:  # noqa: BLE001 — keep going, report each
                print(f"FAILED   : {name}: {exc}")
        status = client.content.v1.contents(sid).approval_fetch().fetch()
        print(f"  {name}: {status.whatsapp}")


if __name__ == "__main__":
    main()
