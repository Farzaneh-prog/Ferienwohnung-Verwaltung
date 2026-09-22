"""
Writes the cleaner schedule (Putzplan2026.xlsx) — one row per reservation's
Abreise (check-OUT) date, since cleaning happens right after a guest
leaves (confirmed with Farzaneh 2026-09-22 — an earlier version of this
module used Anreise/check-in, which was wrong; corrected same day before
any real data was touched by the wrong version).

Scope, by explicit decision (2026-09-22): only NEW reservations processed
by this app from now on get a Putzplan row. The existing 179 rows were
filled by hand over time and have drifted out of sync with GästeListe
(cancellations/edits since) — no attempt is made to reconcile or backfill
them.

Unlike GästeListe, this file has NO formulas anywhere (verified) — so,
unlike xlsx_writer's append-only-at-the-end rule (which exists solely to
avoid breaking formula row-references), a real physical row insertion at
the chronologically correct position is safe here and matches the sheet's
existing global date order (both properties interleaved by Datum).

NOT implemented here (2026-09-22, needs its own data source first):
Farzaneh described a staffing-gap escalation — if the next guest does
NOT arrive the same day (column F = Nein) AND no regular cleaner is
available that day, the room could be closed for one extra night and
handed to Ramic (or someone else who can come in the evening) instead.
This depends on a cleaner-availability roster that doesn't exist yet
anywhere in this app (column A / cleaner assignment is still fully
manual, phase 5 in the architecture doc) — revisit once that roster is
built, not before.
"""
import datetime
import os

import openpyxl

from .config import (
    DATA_DIR,
    PUTZPLAN_FILE,
    PUTZPLAN_SHEET,
    PUTZPLAN_WOHNUNG_LABELS,
    PUTZPLAN_CLEANING_WINDOW,
)
from .excel_reader import load_reservations
from .xlsx_writer import WEEKDAYS_DE, backup_file, format_date_de, parse_date

COL_WER = 1
COL_WOHNUNG = 2
COL_DATUM = 3
COL_TAG = 4
COL_WANN = 5
COL_GLEICHER_TAG = 6
COL_ERWACHSENE = 7
COL_KINDER_U18 = 8
COL_KINDER_U3 = 9
COL_CHECKIN_UHR = 10


def _putzplan_path():
    return os.path.join(DATA_DIR, PUTZPLAN_FILE)


def _parse_de_date(value):
    if not value:
        return None
    try:
        return datetime.datetime.strptime(str(value).strip(), "%d.%m.%Y").date()
    except ValueError:
        return None


def _find_insert_row(ws, target_date):
    """1-based row to insert at so column C (Datum) stays sorted. Scans the
    WHOLE sheet (not just until the first gap) because real gaps exist —
    empty rows and a stray free-text note row at the very end (row 179) —
    and stopping early would silently drop new rows into the wrong spot or
    lose the trailing junk rows this sheet already has."""
    last_dated_row = 1
    for row in range(2, ws.max_row + 2):
        d = _parse_de_date(ws.cell(row=row, column=COL_DATUM).value)
        if d is None:
            continue
        if d > target_date:
            return row
        last_dated_row = row
    return last_dated_row + 1


def _has_same_day_checkin(property_key, checkout_date, exclude_confirmation_code):
    """True if another (non-cancelled) reservation at the same property
    checks IN on this reservation's check-out date — a same-day turnover,
    which is what column F flags (confirmed 2026-09-22: does the NEXT
    guest also arrive the same day the departing guest leaves)."""
    for r in load_reservations(property_key):
        if str(r.get("confirmation_code", "")).strip() == str(exclude_confirmation_code).strip():
            continue
        if str(r.get("cancelled", "")).strip().lower() == "ja":
            continue
        checkin = r.get("checkin")
        if isinstance(checkin, (datetime.date, datetime.datetime)):
            checkin_date = checkin.date() if isinstance(checkin, datetime.datetime) else checkin
        else:
            checkin_date = _parse_de_date(checkin)
        if checkin_date == checkout_date:
            return True
    return False


def append_putzplan_row(entry, log=print):
    """entry is the same shape used by xlsx_writer.append_reservation
    (property, checkin, checkout, adults, children, confirmation_code).
    Best-effort: never raises on a data problem, just logs and skips —
    a missing cleaner-schedule row is not worth failing the whole import
    for."""
    property_key = entry.get("property")
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    if wohnung is None:
        log(f"    WARN: Putzplan skipped — unknown property '{property_key}'")
        return

    path = _putzplan_path()
    if not os.path.exists(path):
        log(f"    WARN: Putzplan skipped — file not found at {path}")
        return

    checkout_date = parse_date(entry["checkout"])
    same_day = _has_same_day_checkin(property_key, checkout_date, entry.get("confirmation_code"))

    backup_file(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[PUTZPLAN_SHEET]

    row = _find_insert_row(ws, checkout_date)
    ws.insert_rows(row)

    # A (Wer macht das) and I (Kinder unter 3) stay blank — cleaner
    # assignment is phase 5 (not built), under-3 age isn't in any export
    # (same decision as column L in GästeListe, 2026-09-16). J (Checkin
    # Uhr) is a free-text notes column in practice, not auto-filled.
    ws.cell(row=row, column=COL_WOHNUNG, value=wohnung)
    ws.cell(row=row, column=COL_DATUM, value=format_date_de(checkout_date))
    ws.cell(row=row, column=COL_TAG, value=WEEKDAYS_DE[checkout_date.weekday()])
    ws.cell(row=row, column=COL_WANN, value=PUTZPLAN_CLEANING_WINDOW)
    ws.cell(row=row, column=COL_GLEICHER_TAG, value="Ja" if same_day else "Nein")
    if entry.get("adults") is not None:
        ws.cell(row=row, column=COL_ERWACHSENE, value=entry["adults"])
    ws.cell(row=row, column=COL_KINDER_U18, value=entry.get("children", 0))

    wb.save(path)
    log(f"    Putzplan: inserted row {row} ({wohnung}, {format_date_de(checkout_date)}, gleicher Tag={same_day})")


def flag_putzplan_cancelled(property_key, checkout_date, adults, children, log=print):
    """Best-effort match on (Wohnung, Datum, Erwachsene, Kinder unter 18) —
    the same heuristic Farzaneh already uses by eye, since Putzplan has no
    confirmation-code column to match on exactly. Writes 'storniert' into
    column J (Checkin Uhr), reusing the convention already present in the
    real file (row 5 has this exact value) rather than inventing a new
    one. Silently does nothing if no matching row is found — this only
    covers reservations created by this app going forward (2026-09-22
    scope decision), so an unmatched cancellation is expected, not an
    error."""
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    if wohnung is None or checkout_date is None:
        return

    path = _putzplan_path()
    if not os.path.exists(path):
        return

    target_date_str = format_date_de(checkout_date) if isinstance(checkout_date, datetime.date) else str(checkout_date)

    backup_file(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[PUTZPLAN_SHEET]

    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=COL_WOHNUNG).value != wohnung:
            continue
        if ws.cell(row=row, column=COL_DATUM).value != target_date_str:
            continue
        row_adults = ws.cell(row=row, column=COL_ERWACHSENE).value
        row_children = ws.cell(row=row, column=COL_KINDER_U18).value or 0
        if row_adults != adults or row_children != (children or 0):
            continue
        ws.cell(row=row, column=COL_CHECKIN_UHR, value="storniert")
        wb.save(path)
        log(f"    Putzplan: marked row {row} storniert ({wohnung}, {target_date_str})")
        return

    log(f"    Putzplan: no matching row found to flag cancelled ({wohnung}, {target_date_str}) — skipped")
