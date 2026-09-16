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
import re
import shutil

import openpyxl
from openpyxl.utils import get_column_letter

from .config import PROPERTY_FILES, PROPERTY_FILE_YEAR, FUTURE_YEAR_SHEET, DATA_DIR
from .excel_reader import FIELD_HEADERS, _build_header_map

# Column indices (1-based) with no header text — structural, not name-addressable.
COL_I_ADULT_NIGHTS = 9    # =F*G
COL_J_CHILD_NIGHTS = 10   # =H*F

WEEKDAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]

# How many columns wide a row can be (A..BH ~ 60) — used when scanning a
# template row for formulas to clone.
MAX_COLUMN = 62


def format_date_de(d: datetime.date) -> str:
    """Every existing row stores dates as literal 'DD.MM.YYYY' text (not a
    real Excel date value) — match that exactly, or new rows visibly stick
    out and sort/filter differently in Excel."""
    return d.strftime("%d.%m.%Y")


def parse_date(s):
    if isinstance(s, datetime.date):
        return s
    return datetime.datetime.strptime(s, "%Y-%m-%d").date()


def target_sheet_for(checkin_date) -> str:
    """Quarter sheet ('1'-'4') for a check-in in the file's own year;
    'Muster' for any other year — that sheet is not actually a template
    (misleading name) but where reservations for a year that doesn't have
    its own GästeListe file yet are held (confirmed with Farzaneh
    2026-09-16), until moved by hand once that year's file exists."""
    if checkin_date.year != PROPERTY_FILE_YEAR:
        return FUTURE_YEAR_SHEET
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


def find_template_row(ws, before_row: int, min_formula_cells=3):
    """Find the nearest row at or above `before_row` that looks like a fully
    formed reservation row (has several formula cells), to clone formulas
    from. Scans upward from before_row down to row 2 first (closest/most
    recent row wins); if that finds nothing — e.g. the insertion point is
    early in a sheet with real data only much further down after a long
    gap, as happens in the 'Muster' sheet — falls back to scanning the
    whole sheet (rows 2-500) for the nearest usable row below instead. A
    half-filled earlier import (few/no formulas) is skipped either way, in
    favor of an older, fully formed row."""

    def has_formulas(r):
        return sum(
            1 for c in range(1, MAX_COLUMN + 1)
            if isinstance(ws.cell(row=r, column=c).value, str) and ws.cell(row=r, column=c).value.startswith("=")
        ) >= min_formula_cells

    for r in range(before_row, 1, -1):
        if has_formulas(r):
            return r
    for r in range(before_row + 1, 500):
        if ws.cell(row=r, column=1).value in (None, ""):
            continue
        if has_formulas(r):
            return r
    return None


_CELL_REF_RE = re.compile(r"([A-Z]{1,3})(\d+)")


def clone_formulas(ws, template_row: int, new_row: int, log=print):
    """Copy every formula cell from template_row into new_row, with all of
    that formula's OWN-ROW cell references (e.g. F26, AA26) rewritten to
    new_row. Never touches manual/literal values (K='ja', addresses, rates,
    etc.) — only cells whose value is a string starting with '='. Safe
    because every real formula in this sheet only ever references its own
    row (verified against the real files) — no cross-row references exist
    to get wrong."""
    if template_row is None:
        log("    WARN: no template row found to clone formulas from — only I/J will be set")
        return
    cloned = 0
    for c in range(1, MAX_COLUMN + 1):
        src = ws.cell(row=template_row, column=c)
        if not (isinstance(src.value, str) and src.value.startswith("=")):
            continue

        def repl(m):
            col_letters, ref_row = m.group(1), int(m.group(2))
            return f"{col_letters}{new_row}" if ref_row == template_row else m.group(0)

        new_formula = _CELL_REF_RE.sub(repl, src.value)
        ws.cell(row=new_row, column=c).value = new_formula
        cloned += 1
    log(f"    cloned {cloned} formulas from row {template_row} into row {new_row}")


def append_reservation(ws, header_map, entry, log=print):
    template_row = find_template_row(ws, last_used_row(ws))
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

    # Clone formulas FIRST so the plain-value fields below correctly
    # overwrite any formula that happens to live in the same column (e.g.
    # none currently do for our known fields, but this keeps the order safe
    # if that ever changes).
    clone_formulas(ws, template_row, row, log=log)

    set_field("guest_name", entry.get("guest_name"))
    set_field("platform", entry.get("platform"))
    set_field("checkin", format_date_de(checkin))
    set_field("checkout", format_date_de(checkout))
    set_field("checkout_day", WEEKDAYS_DE[checkout.weekday()])
    set_field("nights", nights)
    set_field("adults", entry.get("adults"))
    set_field("children", entry.get("children", 0))
    set_field("confirmation_code", entry.get("confirmation_code"))
    if entry.get("guest_email"):
        set_field("guest_email", entry["guest_email"])

    # "Gezahlt" (O) = Preis + Übernachtungssteuer (M) — confirmed by
    # Farzaneh 2026-09-16 against the real Booking.com 'detailed' export
    # ('Price' column) and the sheet's own M column. Written as a formula
    # referencing M{row} (not a second literal) so it stays correct if M
    # is edited later.
    guest_paid = entry.get("guest_paid_total")
    if guest_paid is None:
        guest_paid = entry.get("paid_amount")  # fallback for older Cowork-queue entries
    if guest_paid is not None:
        tourist_tax_col = header_map.get(FIELD_HEADERS["tourist_tax"])
        if tourist_tax_col is not None:
            m_letter = get_column_letter(tourist_tax_col + 1)
            set_field("paid", f"={guest_paid}+{m_letter}{row}")
        else:
            set_field("paid", guest_paid)

    if entry.get("cleaning_fee_charged") is not None:
        set_field("cleaning_fee_charged", entry["cleaning_fee_charged"])

    # AA (Nettobetrag) = the Booking.com 'detailed' export's pure Commission
    # amount. AX (Payment Charge von Booking) = the 'simple' export's
    # Commission (which combines that SAME commission with an extra payment
    # -processing fee) minus AA — isolating just the payment-processing
    # portion, which is what AX's name actually means. Both come from
    # Booking.com specifically; only written when we actually have the
    # matching figure(s) — never guessed, and not applied to Airbnb entries
    # (their fee breakdown means something different — see
    # docs/cowork-command.md history).
    commission_pure = entry.get("platform_fee_total")   # detailed export
    commission_combined = entry.get("platform_fee_combined")  # simple export
    if commission_pure is not None:
        set_field("nettobetrag", commission_pure)
    if commission_pure is not None and commission_combined is not None:
        set_field("platform_commission", round(commission_combined - commission_pure, 2))

    # Structural formula columns (I, J) — no header text, so set_field can't
    # reach them; safe regardless of template_row because they only ever
    # reference their own row.
    ws.cell(row=row, column=COL_I_ADULT_NIGHTS, value=f"=F{row}*G{row}")
    ws.cell(row=row, column=COL_J_CHILD_NIGHTS, value=f"=H{row}*F{row}")

    log(f"    appended row {row}: {entry.get('guest_name')} ({entry['checkin']} -> {entry['checkout']})")
    return row


def apply_cancellation(wb, header_map_cache, entry, log=print):
    code = str(entry.get("confirmation_code", "")).strip()
    if not code:
        log("    WARN: cancellation entry missing confirmation_code, skipping")
        return False
    for sheet_name in ["1", "2", "3", "4", FUTURE_YEAR_SHEET]:
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
        sheet_name = target_sheet_for(parse_date(r["checkin"]))
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
    for sheet_name in ["1", "2", "3", "4", FUTURE_YEAR_SHEET]:
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
