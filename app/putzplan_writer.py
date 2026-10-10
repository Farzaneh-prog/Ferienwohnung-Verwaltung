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
from .xlsx_writer import WEEKDAYS_DE, backup_file, format_date_de, parse_date, save_workbook_atomic

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


def iter_open_rows():
    """Yields (property_key, checkout_date) for every Putzplan row that
    still needs a cleaner assigned — column A (Wer macht das) empty, and
    not already flagged 'storniert' in column J. Read-only. Used by
    app/cleaner_coordination.py (phase 3, 2026-09-30) to know which rows
    need a WhatsApp coordination attempt — column A assignment used to be
    fully manual (see module docstring history); phase 3 is what starts
    filling it in automatically."""
    path = _putzplan_path()
    if not os.path.exists(path):
        return
    label_to_property = {v: k for k, v in PUTZPLAN_WOHNUNG_LABELS.items()}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[PUTZPLAN_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            wer = row[COL_WER - 1]
            wohnung = row[COL_WOHNUNG - 1]
            datum = row[COL_DATUM - 1]
            checkin_uhr = row[COL_CHECKIN_UHR - 1]
            if wer:
                continue
            if str(checkin_uhr or "").strip().lower() == "storniert":
                continue
            property_key = label_to_property.get(wohnung)
            checkout_date = _parse_de_date(datum)
            if property_key is None or checkout_date is None:
                continue
            yield property_key, checkout_date
    finally:
        wb.close()


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


def _find_row_by_date(ws, wohnung, target_date_str, storniert=False):
    """Row for (Wohnung, Datum). Normally unique — see module docstring —
    except after a cancellation + replacement booking for the same day, but
    append_putzplan_row revives the storniert row in that case instead of
    inserting a second one, so there is still at most one. storniert=False
    (default) skips rows flagged 'storniert'; storniert=True looks for
    exactly those."""
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=COL_WOHNUNG).value != wohnung:
            continue
        if ws.cell(row=row, column=COL_DATUM).value != target_date_str:
            continue
        is_storniert = str(ws.cell(row=row, column=COL_CHECKIN_UHR).value or "").strip().lower() == "storniert"
        if is_storniert == storniert:
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
    save_workbook_atomic(wb, path)
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

    # A replacement booking for a date whose earlier reservation was just
    # cancelled: revive that row (clear 'storniert', refresh the guest
    # counts) instead of inserting a second row for the same day — column A
    # keeps the cleaner who was responsible, which is exactly what the
    # cancellation/replacement logic in cleaner_coordination needs.
    revived_row = _find_row_by_date(ws, wohnung, format_date_de(checkout_date), storniert=True)
    if revived_row is not None:
        cleaner_code = ws.cell(row=revived_row, column=COL_WER).value
        ws.cell(row=revived_row, column=COL_CHECKIN_UHR, value=None)
        ws.cell(row=revived_row, column=COL_GLEICHER_TAG, value="Ja" if same_day else "Nein")
        if next_adults is not None:
            ws.cell(row=revived_row, column=COL_ERWACHSENE, value=next_adults)
        ws.cell(row=revived_row, column=COL_KINDER_U18, value=next_children)
        save_workbook_atomic(wb, path)
        log(f"    Putzplan: revived storniert row {revived_row} ({wohnung}, {format_date_de(checkout_date)}), "
            f"cleaner in column A: {cleaner_code}")
        _sync_previous_departure(property_key, checkin_date, entry.get("adults"), entry.get("children", 0), code, log)
        return {"revived": True, "cleaner_code": cleaner_code}

    # A row for this date already exists (created by hand from the dashboard,
    # see register_manual_booking): refresh it, keep column A (cleaner).
    existing_row = _find_row_by_date(ws, wohnung, format_date_de(checkout_date))
    if existing_row is not None:
        cleaner_code = ws.cell(row=existing_row, column=COL_WER).value
        ws.cell(row=existing_row, column=COL_GLEICHER_TAG, value="Ja" if same_day else "Nein")
        if next_adults is not None:
            ws.cell(row=existing_row, column=COL_ERWACHSENE, value=next_adults)
        ws.cell(row=existing_row, column=COL_KINDER_U18, value=next_children)
        save_workbook_atomic(wb, path)
        log(f"    Putzplan: row {existing_row} ({wohnung}, {format_date_de(checkout_date)}) already existed — updated")
        _sync_previous_departure(property_key, checkin_date, entry.get("adults"), entry.get("children", 0), code, log)
        return {"revived": False, "cleaner_code": cleaner_code}

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

    save_workbook_atomic(wb, path)
    log(f"    Putzplan: inserted row {row} ({wohnung}, {format_date_de(checkout_date)}), next guests: "
        f"adults={next_adults}, children={next_children}, gleicher Tag={same_day}")

    # This new reservation might itself be the "next" guest for an
    # earlier departure whose row already exists.
    _sync_previous_departure(property_key, checkin_date, entry.get("adults"), entry.get("children", 0), code, log)
    return {"revived": False, "cleaner_code": None}


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
    docstring.

    Returns None if nothing was flagged, else {"cleaner_code": <column A or
    None>} — the row is NOT deleted and column A keeps the responsible
    cleaner, so cancellation + replacement booking can find them again."""
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    if wohnung is None or checkout_date is None:
        return None

    path = _putzplan_path()
    if not os.path.exists(path):
        return None

    target_date_str = format_date_de(checkout_date) if isinstance(checkout_date, datetime.date) else str(checkout_date)

    backup_file(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[PUTZPLAN_SHEET]

    row = _find_row_by_date(ws, wohnung, target_date_str)
    if row is None:
        log(f"    Putzplan: no matching row found to flag cancelled ({wohnung}, {target_date_str}) — skipped")
        return None

    cleaner_code = ws.cell(row=row, column=COL_WER).value  # stays in the row — who was responsible
    ws.cell(row=row, column=COL_CHECKIN_UHR, value="storniert")
    save_workbook_atomic(wb, path)
    log(f"    Putzplan: marked row {row} storniert ({wohnung}, {target_date_str}), cleaner: {cleaner_code}")
    return {"cleaner_code": cleaner_code}


def assign_cleaner(property_key, checkout_date, cleaner_code, log=print):
    """Writes the confirmed cleaner's code (e.g. 'Meh-Ü', matching
    app/cleaner_roster.py's 'code' field) into column A (Wer macht das) of
    the matching row — the automatic write-back phase 3's WhatsApp
    confirmation triggers (2026-10-01, see app/whatsapp_webhook.py).
    Silently does nothing if no matching row is found (shouldn't normally
    happen — iter_open_rows() is what found this row in the first place —
    but the row could have been manually edited in the meantime).
    cleaner_code=None clears the cell again (the confirmed cleaner backed
    out, 2026-10-01) — the row counts as open in iter_open_rows() again."""
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    if wohnung is None or checkout_date is None:
        return False

    path = _putzplan_path()
    if not os.path.exists(path):
        return False

    target_date_str = format_date_de(checkout_date) if isinstance(checkout_date, datetime.date) else str(checkout_date)

    backup_file(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[PUTZPLAN_SHEET]

    row = _find_row_by_date(ws, wohnung, target_date_str)
    if row is None:
        log(f"    Putzplan: no matching row found to assign cleaner ({wohnung}, {target_date_str}) — skipped")
        return False

    ws.cell(row=row, column=COL_WER, value=cleaner_code)
    save_workbook_atomic(wb, path)
    log(f"    Putzplan: {'cleared cleaner' if cleaner_code is None else 'assigned ' + cleaner_code} "
        f"in row {row} ({wohnung}, {target_date_str})")
    return True


def get_assigned_code(property_key, checkout_date):
    """Column A of the (non-storniert) row for this property/date, or None."""
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    path = _putzplan_path()
    if wohnung is None or checkout_date is None or not os.path.exists(path):
        return None
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    try:
        ws = wb[PUTZPLAN_SHEET]
        row = _find_row_by_date(ws, wohnung, format_date_de(checkout_date))
        return ws.cell(row=row, column=COL_WER).value if row else None
    finally:
        wb.close()


def iter_assigned_rows(checkout_date):
    """Yields (property_key, cleaner_code, guests) for every non-storniert
    row on `checkout_date` that already has a cleaner in column A — used by
    the day-before reminder (16:00) and cancellation (19:00) jobs. `guests`
    is the NEXT arrival's headcount as shown in the row: {"adults", "children"
    (under 18), "children_u3" (column I, filled by hand — usually blank)}.
    Read-only."""
    path = _putzplan_path()
    if not os.path.exists(path):
        return
    label_to_property = {v: k for k, v in PUTZPLAN_WOHNUNG_LABELS.items()}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[PUTZPLAN_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            wer = row[COL_WER - 1]
            if not wer:
                continue
            if str(row[COL_CHECKIN_UHR - 1] or "").strip().lower() == "storniert":
                continue
            property_key = label_to_property.get(row[COL_WOHNUNG - 1])
            if property_key is None or _parse_de_date(row[COL_DATUM - 1]) != checkout_date:
                continue
            yield property_key, str(wer).strip(), {
                "adults": row[COL_ERWACHSENE - 1],
                "children": row[COL_KINDER_U18 - 1],
                "children_u3": row[COL_KINDER_U3 - 1],
            }
    finally:
        wb.close()


def find_active_row(property_key, checkout_date):
    """None if there is no non-storniert Putzplan row for this property/date,
    else {"cleaner": <column A value or None>}. Read-only."""
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    path = _putzplan_path()
    if wohnung is None or checkout_date is None or not os.path.exists(path):
        return None
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    try:
        ws = wb[PUTZPLAN_SHEET]
        row = _find_row_by_date(ws, wohnung, format_date_de(checkout_date))
        if row is None:
            return None
        value = ws.cell(row=row, column=COL_WER).value
        return {"cleaner": str(value).strip() if value else None}
    finally:
        wb.close()


def register_manual_booking(property_key, checkout_date, adults, children, children_u3=None, checkin_date=None,
                            log=print):
    """Dashboard "quick alert" for a booking that has NOT been imported yet
    (2026-10-10). Does by hand what the import does in the Putzplan, so the
    cleaning messages already carry the right numbers:

      - the Putzplan row of the PREVIOUS cleaning (latest row before the new
        booking's check-in / checkout) now describes THIS booking as its next
        guests: columns G/H(/I) = adults, children (under 18), under-3;
      - the row for `checkout_date` (the new booking's own cleaning) is created
        if it doesn't exist, and gets the previous row's OLD numbers (G/H/I):
        those guests used to be the "next guests" of the previous cleaning and
        are now one cleaning later. Only if the previous row had no numbers is
        the guest list consulted (the booking checking in after that date).

    A later import of the booking finds the row and just refreshes it.
    Returns {"previous_date", "created_row", "next_guests"} (best effort)."""
    wohnung = PUTZPLAN_WOHNUNG_LABELS.get(property_key)
    path = _putzplan_path()
    if wohnung is None or checkout_date is None or not os.path.exists(path):
        return {"previous_date": None, "created_row": False, "next_guests": None}

    backup_file(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[PUTZPLAN_SHEET]

    prev_row, prev_date = None, None
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=COL_WOHNUNG).value != wohnung:
            continue
        if str(ws.cell(row=row, column=COL_CHECKIN_UHR).value or "").strip().lower() == "storniert":
            continue
        d = _parse_de_date(ws.cell(row=row, column=COL_DATUM).value)
        if d is None or d >= checkout_date or (checkin_date is not None and d > checkin_date):
            continue
        if prev_date is None or d > prev_date:
            prev_row, prev_date = row, d

    old = {"adults": None, "children": None, "children_u3": None}
    if prev_row is not None:
        old = {"adults": ws.cell(row=prev_row, column=COL_ERWACHSENE).value,
               "children": ws.cell(row=prev_row, column=COL_KINDER_U18).value,
               "children_u3": ws.cell(row=prev_row, column=COL_KINDER_U3).value}

    # the new booking's own cleaning row
    created = False
    next_guests = None
    own_row = _find_row_by_date(ws, wohnung, format_date_de(checkout_date))
    if own_row is None:
        nxt = _find_next_reservation(property_key, checkout_date, None)
        same_day = nxt is not None and _to_date(nxt.get("checkin")) == checkout_date
        if old["adults"] not in (None, ""):
            next_guests = dict(old)  # carried over from the previous row (the user's rule)
        elif nxt is not None and nxt.get("adults") not in (None, ""):
            next_guests = {"adults": nxt.get("adults"), "children": nxt.get("children") or 0, "children_u3": None}
        else:
            next_guests = None
        row = _find_insert_row(ws, checkout_date)
        ws.insert_rows(row)
        ws.cell(row=row, column=COL_WOHNUNG, value=wohnung)
        ws.cell(row=row, column=COL_DATUM, value=format_date_de(checkout_date))
        ws.cell(row=row, column=COL_TAG, value=WEEKDAYS_DE[checkout_date.weekday()])
        ws.cell(row=row, column=COL_WANN, value=PUTZPLAN_CLEANING_WINDOW)
        ws.cell(row=row, column=COL_GLEICHER_TAG, value="Ja" if same_day else "Nein")
        if next_guests:
            ws.cell(row=row, column=COL_ERWACHSENE, value=next_guests["adults"])
            ws.cell(row=row, column=COL_KINDER_U18, value=next_guests["children"] or 0)
            if next_guests.get("children_u3") not in (None, ""):
                ws.cell(row=row, column=COL_KINDER_U3, value=next_guests["children_u3"])
        created = True
        if prev_row is not None and row <= prev_row:  # never happens (dates sort), but keep indices honest
            prev_row += 1

    # the previous cleaning's row now shows THIS booking
    if prev_row is not None:
        ws.cell(row=prev_row, column=COL_ERWACHSENE, value=adults)
        ws.cell(row=prev_row, column=COL_KINDER_U18, value=children or 0)
        if children_u3 in ("Ja", "Nein"):
            ws.cell(row=prev_row, column=COL_KINDER_U3, value=children_u3)
        if checkin_date is not None:
            ws.cell(row=prev_row, column=COL_GLEICHER_TAG, value="Ja" if checkin_date == prev_date else "Nein")

    save_workbook_atomic(wb, path)
    log(f"    Putzplan: manual booking {wohnung} {format_date_de(checkout_date)}: previous row "
        f"{prev_date and format_date_de(prev_date)} -> adults={adults}, children={children}; "
        f"own row {'created' if created else 'already there'}, next guests {next_guests}")
    return {"previous_date": prev_date, "created_row": created, "next_guests": next_guests}
