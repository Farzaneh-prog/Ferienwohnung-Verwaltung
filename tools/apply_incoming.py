# -*- coding: utf-8 -*-
"""
Fetches queued entries from the marktresidenz-eisenach.de staging inbox
(see wordpress-plugin/marktresidenz-incoming-api.php) and safely applies
them to the real GästeListe_*.xlsx files.

Safety rules (see README "Excel-Dateigröße" section and project history
for why these matter):
- Never inserts or deletes a row in the middle of a sheet (openpyxl does not
  rewrite other rows' formula references — see commit history). New
  reservations are only ever appended after the last used row of the
  correct quarter sheet.
- Cancellations are never deleted — the "Storniert" flag column (BB) is set
  to "ja" instead. The row stays fully intact; a human can safely delete it
  later from within Excel itself (Excel *does* fix up formula references on
  a native row delete).
- Every file is backed up (to backups/, gitignored) before being touched.
- An incoming entry is only ack'd (removed from the queue) after its file
  write has succeeded, so a crash mid-run just gets retried next time.

Run manually, or on a schedule, whenever the computer is on:
    python tools/apply_incoming.py
"""
import datetime
import os
import shutil
import sys

import openpyxl
import requests
from dotenv import load_dotenv

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

from app.config import PROPERTY_FILES, DATA_DIR  # noqa: E402
from app.excel_reader import FIELD_HEADERS, _build_header_map  # noqa: E402

API_BASE = os.environ.get("MRV_API_BASE")
READ_KEY = os.environ.get("MRV_READ_KEY")

# Column indices (1-based) with no header text — structural, not name-addressable.
COL_I_ADULT_NIGHTS = 9    # =F*G
COL_J_CHILD_NIGHTS = 10   # =H*F

WEEKDAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]

# Which of the "over-fetched" JSON fields we actually know how to place, and
# into which FIELD_HEADERS key. Anything else in the payload is logged but
# left untouched — deliberately conservative (see README/PII discussion:
# never guess where a number belongs).
KNOWN_RESERVATION_FIELDS = {
    "guest_name": "guest_name",
    "platform": "platform",
    "checkin": "checkin",
    "checkout": "checkout",
    "adults": "adults",
    "children": "children",
    "confirmation_code": "confirmation_code",
    "guest_email": "guest_email",
    "paid_amount": "paid",              # -> "Gezahlt"
    "cleaning_fee_charged": None,        # -> "Endreinigung" handled specially (see below)
    "platform_commission": None,         # -> "Payment Charge von Booking" (see below)
}


def parse_date(s):
    return datetime.datetime.strptime(s, "%Y-%m-%d").date()


def quarter_sheet_for(checkin_date) -> str:
    return str((checkin_date.month - 1) // 3 + 1)


def backup_file(path):
    os.makedirs(os.path.join(DATA_DIR, "backups"), exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = os.path.join(
        DATA_DIR, "backups",
        f"{os.path.splitext(os.path.basename(path))[0]}.pre-apply-{ts}.xlsx",
    )
    shutil.copy2(path, dest)
    return dest


def last_used_row(ws) -> int:
    """Last row (1-based) with anything in column A, scanning from the top —
    real sheets are always small (tens of rows), so this is fast and never
    trusts ws.max_row (which used to be bloated to ~1M, see shrink_xlsx.py)."""
    row = 1
    while ws.cell(row=row + 1, column=1).value not in (None, ""):
        row += 1
    return row


def append_reservation(ws, header_map, entry):
    row = last_used_row(ws) + 1

    def set_field(field_key, value):
        header = FIELD_HEADERS[field_key]
        col = header_map.get(header)
        if col is None:
            print(f"    WARN: column '{header}' not found, skipping {field_key}")
            return
        # NOTE: ws.cell(..., value=None) silently no-ops in openpyxl (None is
        # the "not given" sentinel) — always set .value directly, never via
        # the value= kwarg, or a None field gets silently skipped instead of
        # written. Harmless here (fresh row, nothing to overwrite) but keep
        # the pattern correct so it stays safe if this is ever reused to
        # update an existing row.
        ws.cell(row=row, column=col + 1).value = value  # header_map is 0-based

    checkin = parse_date(entry["checkin"])
    checkout = parse_date(entry["checkout"])
    nights = (checkout - checkin).days

    set_field("guest_name", entry.get("guest_name"))
    set_field("platform", entry.get("platform"))
    set_field("checkin", entry["checkin"])
    set_field("checkout", entry["checkout"])
    set_field("checkout_day", WEEKDAYS_DE[checkout.weekday()])
    set_field("nights", nights)
    set_field("adults", entry.get("adults"))
    set_field("children", entry.get("children", 0))
    set_field("confirmation_code", entry.get("confirmation_code"))
    if entry.get("guest_email"):
        set_field("guest_email", entry["guest_email"])
    # "Gezahlt" = what the guest paid in total. Prefer the unambiguous
    # guest_paid_total (from the platform's own payout breakdown); the older
    # generic paid_amount is kept as a fallback for entries sent before this
    # field existed, but it was found to sometimes actually be the host
    # payout (after platform fees), not the guest total — see cowork-command.md.
    guest_paid = entry.get("guest_paid_total")
    if guest_paid is None:
        guest_paid = entry.get("paid_amount")
    if guest_paid is not None:
        set_field("paid", guest_paid)
    if entry.get("cleaning_fee_charged") is not None:
        set_field("cleaning_fee_charged", entry["cleaning_fee_charged"])

    # Structural formula columns (I, J) — safe because this is a brand new
    # row referencing only its own cells; no other row is touched.
    ws.cell(row=row, column=COL_I_ADULT_NIGHTS, value=f"=F{row}*G{row}")
    ws.cell(row=row, column=COL_J_CHILD_NIGHTS, value=f"=H{row}*F{row}")

    print(f"    appended row {row}: {entry.get('guest_name')} ({entry['checkin']} -> {entry['checkout']})")


def apply_cancellation(wb, header_map_cache, entry):
    code = str(entry.get("confirmation_code", "")).strip()
    if not code:
        print("    WARN: cancellation entry missing confirmation_code, skipping")
        return False
    for sheet_name in ["1", "2", "3", "4"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_map = header_map_cache.setdefault(
            sheet_name, _build_header_map(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        )
        code_col = header_map.get(FIELD_HEADERS["confirmation_code"])
        flag_col = header_map.get(FIELD_HEADERS["cancelled"])
        if code_col is None or flag_col is None:
            continue
        row = 1
        while ws.cell(row=row + 1, column=1).value not in (None, ""):
            row += 1
            cell_val = ws.cell(row=row, column=code_col + 1).value
            if str(cell_val).strip() == code:
                ws.cell(row=row, column=flag_col + 1, value="ja")
                print(f"    marked row {row} (sheet {sheet_name}) as Storniert (code={code})")
                return True
    print(f"    WARN: confirmation_code {code} not found, cancellation not applied")
    return False


def process_entry(entry):
    payload = entry["payload"]
    new_reservations = payload.get("new_reservations", [])
    cancellations = payload.get("cancellations", [])

    touched_files = {}  # property -> (path, workbook)

    def get_workbook(property_key):
        if property_key not in touched_files:
            path = os.path.join(DATA_DIR, PROPERTY_FILES[property_key])
            backup_file(path)
            touched_files[property_key] = [path, openpyxl.load_workbook(path)]
        return touched_files[property_key][1]

    header_map_cache = {}

    for r in new_reservations:
        property_key = r.get("property")
        if property_key not in PROPERTY_FILES:
            print(f"    WARN: unknown property '{property_key}', skipping reservation")
            continue
        wb = get_workbook(property_key)
        sheet_name = quarter_sheet_for(parse_date(r["checkin"]))
        ws = wb[sheet_name]
        header_map = header_map_cache.setdefault(
            (property_key, sheet_name),
            _build_header_map(next(ws.iter_rows(min_row=1, max_row=1, values_only=True))),
        )
        append_reservation(ws, header_map, r)

    for c in cancellations:
        property_key = c.get("property")
        if property_key not in PROPERTY_FILES:
            print(f"    WARN: unknown property '{property_key}', skipping cancellation")
            continue
        wb = get_workbook(property_key)
        apply_cancellation(wb, {k[1]: v for k, v in header_map_cache.items() if k[0] == property_key}, c)

    for property_key, (path, wb) in touched_files.items():
        wb.save(path)
        print(f"    saved {path}")

    if payload.get("extra") or any(r.get("extra") for r in new_reservations):
        print("    NOTE: payload contained extra/unmapped fields — review manually, nothing was guessed.")


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
