# -*- coding: utf-8 -*-
"""
Fetches queued entries from the marktresidenz-eisenach.de staging inbox
(see wordpress-plugin/marktresidenz-incoming-api.php) and safely applies
them to the real GästeListe_*.xlsx files via app.xlsx_writer.process_batch
(the same safe-write logic used by the manual CSV/XLS import).

An incoming entry is only ack'd (removed from the queue) after its file
write has succeeded, so a crash mid-run just gets retried next time.

Run manually, or on a schedule, whenever the computer is on:
    python tools/apply_incoming.py
"""
import os
import sys

import requests
from dotenv import load_dotenv

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

from app.xlsx_writer import process_batch  # noqa: E402

API_BASE = os.environ.get("MRV_API_BASE")
READ_KEY = os.environ.get("MRV_READ_KEY")


def process_entry(entry):
    payload = entry["payload"]
    result = process_batch(payload.get("new_reservations", []), payload.get("cancellations", []))
    if payload.get("extra") or any(r.get("extra") for r in payload.get("new_reservations", [])):
        print("    NOTE: payload contained extra/unmapped fields — review manually, nothing was guessed.")
    return result


def main():
    if not API_BASE or not READ_KEY:
        print("MRV_API_BASE / MRV_READ_KEY missing from .env — nothing to do.")
        return

    resp = requests.get(f"{API_BASE}/incoming", headers={"X-MRV-Read-Key": READ_KEY}, timeout=30)
    resp.raise_for_status()
    entries = resp.json().get("entries", [])

    if not entries:
        print("No pending entries.")
        return

    for entry in entries:
        print(f"Processing entry id={entry['id']} run_date={entry['run_date']}")
        try:
            process_entry(entry)
        except Exception as exc:  # noqa: BLE001
            print(f"  FAILED: {exc} — leaving entry {entry['id']} un-acked for retry.")
            continue
        ack = requests.post(
            f"{API_BASE}/incoming/{entry['id']}/ack",
            headers={"X-MRV-Read-Key": READ_KEY},
            timeout=30,
        )
        ack.raise_for_status()
        print(f"  acked entry {entry['id']}")


if __name__ == "__main__":
    main()
