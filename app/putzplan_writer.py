"""
Writes the cleaner schedule (Putzplan2026.xlsx) — one row per reservation's
Abreise (check-OUT) date, since cleaning happens right after a guest
leaves (confirmed with Farzaneh 2026-09-22).

Guest counts (columns G/H, Erwachsene/Kinder unter 18) show the NEXT
reservation's headcount, not the departing guests' — confirmed 2026-09-24.
The cleaner needs to know who's arriving next (how many beds/towels to
prepare), not who just left. An earlier version of this module wrote the
departing reservation's own counts, which was wrong; corrected same day.

Because "next reservation" isn't always known yet when a row is created
(the next guest might not be booked at all, or might get booked/completed
later), every place that learns a reservation's checkin+guest-count also
tries to refresh the PRECEDING departure's Putzplan row (see
_sync_previous_departure) — covers both "a new booking arrives after its
predecessor's row already exists" and "an existing reservation's counts
get filled in later via /complete".

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

Row matching (Wohnung, Datum) is treated as unique — each property here
is a single physical unit, so only one reservation can check out of it on
any given day. Simpler and more reliable than the earlier guest-count-
based heuristic this module used before 2026-09-24.

NOT implemented here (2026-09-22, needs its own data source first):
Farzaneh described a staffing-gap escalation — if the next guest does
NOT arrive the same day (column F = Nein) AND no regular cleaner is
available that day, the room could be closed for one extra night and
handed to Ramic (or someone else who can come in the evening) instead.
This depends on a cleaner-availability roster that doesn't exist yet
anywhere in this app (column A / cleaner assignment is still fully
manual, phase 5 in the architecture doc) — revisit once that roster is
built, not before.

NOT implemented here (2026-09-24, known limitation): if a reservation
that was some earlier departure's "next" gets cancelled, that earlier
row's guest counts go stale again (still show the now-cancelled
reservation's headcount) — cancellation doesn't currently trigger a
re-lookup of the new actual next. Low-frequency edge case; revisit if it
turns out to matter in practice.
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


def _to_date(value):
    """Normalize a GästeListe date cell (openpyxl datetime, or our own
    'DD.MM.YYYY' text) to a plain date, or None."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return _parse_de_date(value)


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


def _find_row_by_date(ws, wohnung, target_date_str):
    """(Wohnung, Datum) is unique — see module docstring."""
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=COL_WOHNUNG).value != wohnung:
            continue
        if ws.cell(row=row, column=COL_DATUM).value != target_date_str:
            continue
        return row
    return None


def _find_next_reservation(property_key, checkout_date, exclude_confirmation_code):
    """The reservation checking IN soonest at/after checkout_date, at the
    same property (excluding the departing reservation itself and any
    cancelled ones) — this is who Putzplan needs to describe."""
    candidates = []
    for r in load_reservations(property_key):
        if str(r.get("confirmation_code", "")).strip() == str(exclude_confirmation_code).strip():
            continue
        if str(r.get("cancelled", "")).strip().lower() == "ja":
            continue
        checkin_date = _to_date(r.get("checkin"))
        if checkin_date is None or checkin_date < checkout_date:
            continue
        candidates.append((checkin_date, r))
    if not candidates:
        return None
    candidates.sort(key=lambda pair: pair[0])
    return candidates[0][1]


def _find_previous_departure(property_key, checkin_date, exclude_confirmation_code):
    """Reverse of _find_next_reservation — the reservation checking OUT
    most recently at/before checkin_date, at the same property. Used to
    find which existing Putzplan row (dated at that departure) should now
    describe THIS reservation as its next guests."""
    candidates = []
    for r in load_reservations(property_key):
        if str(r.get("confirmation_code", "")).strip() == str(exclude_confirmation_code).strip():
            continue
        if str(r.get("cancelled", "")).strip().lower() == "ja":
            continue
        checkout_date = _to_date(r.get("checkout"))
        if checkout_date is None or checkout_date > checkin_date:
            continue
        candidates.append((checkout_date, r))
    if not candidates:
        return None
    candidates.sort(key=lambda pair: pair[0])
    return candidates[-1][1]


def _sync_previous_departure(property_key, checkin_date, adults, children, exclude_confirmation_code, log):
    """If some earlier departure's Putzplan row already exists and THIS
    reservation is (now) its next guest, refresh that row's guest counts
    and same-day flag. No-op if there's no such row (older/pre-automation
    data, or nothing preceding). Shared by append_putzplan_row (a new
    booking might complete an existing row) and the /complete-triggered
    sync (a reservation's counts becoming known later)."""
    prev = _find_previous_departure(property_key, checkin_date, exclude_confirmation_code)
    if prev is None:
        return
    prev_checkout = _to_date(prev.get("checkout"))
    if prev_checkout is None:
        return

    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    path = _putzplan_path()
    if not os.path.exists(path):
        return
    target_date_str = format_date_de(prev_checkout)

    backup_file(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[PUTZPLAN_SHEET]

    row = _find_row_by_date(ws, wohnung, target_date_str)
    if row is None:
        return

    if adults is not None:
        ws.cell(row=row, column=COL_ERWACHSENE, value=adults)
    ws.cell(row=row, column=COL_KINDER_U18, value=children or 0)
    same_day = checkin_date == prev_checkout
    ws.cell(row=row, column=COL_GLEICHER_TAG, value="Ja" if same_day else "Nein")
    wb.save(path)
    log(f"    Putzplan: refreshed row {row} ({wohnung}, {target_date_str}) next-guest counts -> adults={adults}, children={children}")


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

    checkin_date = parse_date(entry["checkin"])
    checkout_date = parse_date(entry["checkout"])
    code = entry.get("confirmation_code")

    next_res = _find_next_reservation(property_key, checkout_date, code)
    if next_res is not None:
        next_checkin = _to_date(next_res.get("checkin"))
        same_day = next_checkin == checkout_date
        next_adults = next_res.get("adults")
        next_children = next_res.get("children") or 0
    else:
        same_day = False
        next_adults = None
        next_children = 0

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
    if next_adults is not None:
        ws.cell(row=row, column=COL_ERWACHSENE, value=next_adults)
    ws.cell(row=row, column=COL_KINDER_U18, value=next_children)

    wb.save(path)
    log(f"    Putzplan: inserted row {row} ({wohnung}, {format_date_de(checkout_date)}), next guests: "
        f"adults={next_adults}, children={next_children}, gleicher Tag={same_day}")

    # This new reservation might itself be the "next" guest for an
    # earlier departure whose row already exists.
    _sync_previous_departure(property_key, checkin_date, entry.get("adults"), entry.get("children", 0), code, log)


def sync_putzplan_for_reservation(property_key, checkin_date, adults, children, exclude_confirmation_code, log=print):
    """Public entry point for xlsx_writer.update_reservation_fields — call
    when a reservation's adults/children becomes known/changes (e.g. via
    /complete or a detailed re-import filling in a placeholder), so the
    PRECEDING departure's Putzplan row (which shows THIS reservation's
    headcount as its 'next guests') gets refreshed too."""
    if checkin_date is None:
        return
    checkin_date = _to_date(checkin_date) if not isinstance(checkin_date, datetime.date) else checkin_date
    if checkin_date is None:
        return
    _sync_previous_departure(property_key, checkin_date, adults, children, exclude_confirmation_code, log)


def flag_putzplan_cancelled(property_key, checkout_date, log=print):
    """Marks the matching row (by Wohnung+Datum — see module docstring)
    'storniert' in column J (Checkin Uhr), reusing the convention already
    present in the real file (row 5 had this exact value) rather than
    inventing a new one. Silently does nothing if no matching row is
    found — this only covers reservations created by this app going
    forward (2026-09-22 scope decision), so an unmatched cancellation is
    expected, not an error.

    Known limitation (2026-09-24): if the cancelled reservation was
    itself some earlier departure's 'next guest', that earlier row's
    counts are now stale and aren't automatically recomputed — see module
    docstring."""
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

    row = _find_row_by_date(ws, wohnung, target_date_str)
    if row is None:
        log(f"    Putzplan: no matching row found to flag cancelled ({wohnung}, {target_date_str}) — skipped")
        return

    ws.cell(row=row, column=COL_CHECKIN_UHR, value="storniert")
    wb.save(path)
    log(f"    Putzplan: marked row {row} storniert ({wohnung}, {target_date_str})")
