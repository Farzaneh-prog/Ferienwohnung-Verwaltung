"""
Shared safe-write logic for the real GästeListe_*.xlsx files — used by both
the Cowork pipeline (tools/apply_incoming.py) and the manual CSV/XLS import
(app/routes_import.py), so there is exactly one place that knows how to
touch these files safely.

Safety rules (see README "Excel-Dateigröße" section and project history):
- Never inserts or deletes a row in the middle of a sheet (openpyxl does not
  rewrite other rows' formula references). New reservations are only ever
  appended after the last used row of the correct quarter sheet.
- Cancellations are never deleted — the "Storniert" flag column is set to
  "ja" instead. A human can safely delete the row later from within Excel
  itself (Excel *does* fix up formula references on a native row delete).
- Every file is backed up (to backups/, gitignored) before being touched.
"""
import datetime
import os
import shutil

import openpyxl

from .config import PROPERTY_FILES, DATA_DIR
from .excel_reader import FIELD_HEADERS, _build_header_map

# Column indices (1-based) with no header text — structural, not name-addressable.
COL_I_ADULT_NIGHTS = 9    # =F*G
COL_J_CHILD_NIGHTS = 10   # =H*F

WEEKDAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def parse_date(s):
    if isinstance(s, datetime.date):
        return s
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


def append_reservation(ws, header_map, entry, log=print):
    row = last_used_row(ws) + 1

    def set_field(field_key, value):
        header = FIELD_HEADERS[field_key]
        col = header_map.get(header)
        if col is None:
            log(f"    WARN: column '{header}' not found, skipping {field_key}")
            return
        # NOTE: ws.cell(..., value=None) silently no-ops in openpyxl (None is
        # the "not given" sentinel) — always set .value directly.
        ws.cell(row=row, column=col + 1).value = value  # header_map is 0-based

    checkin = parse_date(entry["checkin"])
    checkout = parse_date(entry["checkout"])
    nights = (checkout - checkin).days

    set_field("guest_name", entry.get("guest_name"))
    set_field("platform", entry.get("platform"))
    set_field("checkin", entry["checkin"] if isinstance(entry["checkin"], str) else checkin.isoformat())
    set_field("checkout", entry["checkout"] if isinstance(entry["checkout"], str) else checkout.isoformat())
    set_field("checkout_day", WEEKDAYS_DE[checkout.weekday()])
    set_field("nights", nights)
    set_field("adults", entry.get("adults"))
    set_field("children", entry.get("children", 0))
    set_field("confirmation_code", entry.get("confirmation_code"))
    if entry.get("guest_email"):
        set_field("guest_email", entry["guest_email"])

    # "Gezahlt" = what the guest paid in total. Prefer the unambiguous
    # guest_paid_total (from the platform's own breakdown); the older
    # generic paid_amount is kept as a fallback for entries sent before this
    # field existed, but it was found to sometimes actually be the host
    # payout (after platform fees), not the guest total.
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

    log(f"    appended row {row}: {entry.get('guest_name')} ({entry['checkin']} -> {entry['checkout']})")
    return row


def apply_cancellation(wb, header_map_cache, entry, log=print):
    code = str(entry.get("confirmation_code", "")).strip()
    if not code:
        log("    WARN: cancellation entry missing confirmation_code, skipping")
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
                log(f"    marked row {row} (sheet {sheet_name}) as Storniert (code={code})")
                return True
    log(f"    WARN: confirmation_code {code} not found, cancellation not applied")
    return False


def process_batch(new_reservations, cancellations, log=print):
    """Apply a batch of reservations/cancellations across however many
    properties they touch, backing up each touched file exactly once. Used
    by both the Cowork pipeline and the manual CSV/XLS import."""
    touched_files = {}  # property -> [path, workbook]
    header_map_cache = {}
    applied_rows = []
    skipped = []

    def get_workbook(property_key):
        if property_key not in touched_files:
            path = os.path.join(DATA_DIR, PROPERTY_FILES[property_key])
            backup_file(path)
            touched_files[property_key] = [path, openpyxl.load_workbook(path)]
        return touched_files[property_key][1]

    for r in new_reservations:
        property_key = r.get("property")
        if property_key not in PROPERTY_FILES:
            log(f"    WARN: unknown property '{property_key}', skipping reservation")
            skipped.append(r)
            continue
        wb = get_workbook(property_key)
        sheet_name = quarter_sheet_for(parse_date(r["checkin"]))
        ws = wb[sheet_name]
        header_map = header_map_cache.setdefault(
            (property_key, sheet_name),
            _build_header_map(next(ws.iter_rows(min_row=1, max_row=1, values_only=True))),
        )
        row = append_reservation(ws, header_map, r, log=log)
        applied_rows.append({"property": property_key, "sheet": sheet_name, "row": row, "entry": r})

    for c in cancellations:
        property_key = c.get("property")
        if property_key not in PROPERTY_FILES:
            log(f"    WARN: unknown property '{property_key}', skipping cancellation")
            skipped.append(c)
            continue
        wb = get_workbook(property_key)
        cache_for_property = {k[1]: v for k, v in header_map_cache.items() if k[0] == property_key}
        apply_cancellation(wb, cache_for_property, c, log=log)

    for property_key, (path, wb) in touched_files.items():
        wb.save(path)
        log(f"    saved {path}")

    return {"applied": applied_rows, "skipped": skipped, "files_touched": list(touched_files.keys())}


def update_reservation_fields(property_key, confirmation_code, updates: dict, log=print):
    """Locates a row by confirmation code and updates only the given fields
    (e.g. from the manual-completion form: real guest_name/guest_email/
    adults/children replacing a placeholder). Backs up the file first."""
    path = os.path.join(DATA_DIR, PROPERTY_FILES[property_key])
    backup_file(path)
    wb = openpyxl.load_workbook(path)
    code = str(confirmation_code).strip()
    found = False
    for sheet_name in ["1", "2", "3", "4"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_map = _build_header_map(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        code_col = header_map.get(FIELD_HEADERS["confirmation_code"])
        if code_col is None:
            continue
        row = 1
        while ws.cell(row=row + 1, column=1).value not in (None, ""):
            row += 1
            if str(ws.cell(row=row, column=code_col + 1).value).strip() == code:
                for field_key, value in updates.items():
                    header = FIELD_HEADERS.get(field_key)
                    col = header_map.get(header) if header else None
                    if col is None:
                        log(f"    WARN: column for '{field_key}' not found, skipping")
                        continue
                    ws.cell(row=row, column=col + 1).value = value
                found = True
                break
        if found:
            break
    if found:
        wb.save(path)
        log(f"    updated row for code={code} in {path}")
    else:
        log(f"    WARN: code={code} not found in {path}, nothing updated")
    return found
