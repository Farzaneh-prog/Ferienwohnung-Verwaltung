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
from openpyxl.utils import get_column_letter, column_index_from_string

from .config import PROPERTY_FILES, PROPERTY_FILE_YEAR, FUTURE_YEAR_SHEET, DATA_DIR
from .excel_reader import FIELD_HEADERS, _build_header_map
from . import known_codes

# Column indices (1-based) with no header text — structural, not name-addressable.
COL_I_ADULT_NIGHTS = 9    # =F*G
COL_J_CHILD_NIGHTS = 10   # =H*F

WEEKDAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]

# Business-rule constants confirmed by Farzaneh 2026-09-16 — not derived
# from any export, just fixed defaults she gave directly.
DEFAULT_VAT_RATE = 0.19  # Z / Umsatzsteuersatz — always 19%
# P / Putzarbeit (paid to the cleaner) — fixed rate per property, but only
# for Booking.com reservations; no rule given for Airbnb/Vrbo, so those
# stay blank rather than guessed.
CLEANING_COST_BOOKING = {"karlstrasse": 50, "eisenach": 65}

# Canonical formulas confirmed directly by Farzaneh (2026-09-16), addressed
# by their fixed real column LETTER rather than FIELD_HEADERS — R and AQ
# both happen to share the exact same header text ("Ausländische Firma mit
# steuer in Heimat"), so a name-based lookup can't tell them apart; these
# letters were verified against the real files repeatedly this session and
# have been stable. Applied AFTER clone_formulas, deliberately overriding
# whatever got cloned: several historical rows had these as one-off
# hardcoded/pasted numbers rather than live formulas (e.g. AF was a plain
# manual value in every example row checked), so blind cloning would have
# propagated a stale number instead of the real formula.
# {r} is substituted with the new row's own number.
EXPLICIT_FORMULAS = {
    "AK": '=IF(AQ{r}="ja",AJ{r},AJ{r}/1.07)',
    "AL": "=AK{r}*1.07",
    "AM": "=AJ{r}",
    "BC": "=AF{r}+AH{r}+AN{r}+AO{r}+AP{r}-AE{r}+AG{r}",
    "AV": (
        '=IF(B{r}="Airbnb","Ein Teil der Zahlung wird voraussichtlich einen Tag nach Ihrem '
        'Check-in über die Plattform Airbnb erfolgen. Bitte überweisen Sie die '
        'Tourismusförderabgabe separat auf das unten genannte Bankkonto.",'
        'IF(B{r}="booking","Ein Teil der Zahlung wird voraussichtlich einen Tag nach Ihrem '
        'Aufenthalt über die Plattform Booking.com erfolgen. Bitte überweisen Sie die '
        'Tourismusförderabgabe separat auf das unten genannte Bankkonto.", '
        'IF(AW{r}="ja","Die Zahlung für diese Unterkunft ist bereits erfolgt.",'
        'IF(B{r}="Vrbo","Die Zahlung erfolgt voraussichtlich zehn Tage nach Ihrem Aufenthalt '
        'über die Plattform Vrbo.", "Bitte überweisen Sie den Gesamtbetrag von "&'
        'TEXT(AE{r},"0.00")&" € binnen 14 Tagen auf das unten genannte Bankkonto."))))'
    ),
    # U (bezahlt zum putzfrau) — placeholder default, see EXPLICIT_LITERALS
    # comment below; real value once the actual cleaner/hours are known.
    "U": "=T{r}*12",
    # AN-BE (except AT, which is property-specific — see
    # EXPLICIT_FORMULAS_BY_PROPERTY below) — taken verbatim from the final
    # AK-BE table in نظافتچی_ها-و-فرمول_های-مالی.md (documented, not just
    # cloned from a neighboring row).
    "AN": "=ROUND(AK{r}*F{r},2)",
    "AO": '=IF(AQ{r}="ja","§13b UStG",ROUND((AN{r}+AF{r})*0.07,2))',
    "AP": '=IF(AQ{r}="ja","§13b UStG",ROUND(AG{r}*0.19,2))',
    "AR": "=AN{r}+AF{r}+AG{r}+AH{r}",
    "AS": "=ROUND(AK{r},2)",
    "AU": '=IF(B{r}="Airbnb","Airbnb Bestätigungscode: ",IF(B{r}="booking","Booking number: "," "))',
    "BE": "=Q{r}/F{r}",
    # Q, AB, AC, AE, AH, AI, AJ — confirmed by Farzaneh 2026-09-16. These
    # were the only 7 columns with NO documentation anywhere (not in either
    # architecture doc, not hardcoded) — everything else in the AK-BE range
    # was already covered by نظافتچی_ها-و-فرمول_های-مالی.md.
    "Q": '=IF(B{r}="booking",O{r}-U{r}-M{r}-AA{r}-AX{r}-AO{r},O{r}-U{r}-M{r}-AA{r}-AX{r}-AO{r}-AP{r})',
    "AB": "=AA{r}*0.19",
    "AC": "=AA{r}+AB{r}",
    "AE": "=O{r}",
    "AH": "=M{r}",
    "AI": '=IF(AQ{r}="ja",AE{r}-AF{r}-AG{r}-AH{r},AE{r}-AF{r}*1.07-AG{r}*1.19-AH{r})',
    "AJ": "=AI{r}/F{r}",
}
# AT (Stell / Stellplatz-Hinweistext) differs by property — Karlstraße says
# "exkl. Stellplatz" (parking not included), Eisenach says "inkl." (parking
# IS included in that property's price) — per
# نظافتچی_ها-و-فرمول_های-مالی.md. Nested IF (not OR(...)) to match the
# real formula text exactly, confirmed by Farzaneh 2026-09-16. Applied on
# top of EXPLICIT_FORMULAS, keyed by the reservation's own property, not
# the sheet/file it lands in.
EXPLICIT_FORMULAS_BY_PROPERTY = {
    "karlstrasse": {
        "AT": '=IF(B{r}="Airbnb",", exkl. Stellplatz.",IF(B{r}="booking",", exkl. Stellplatz.","."))',
    },
    "eisenach": {
        "AT": '=IF(B{r}="Airbnb",", inkl. Stellplatz.",IF(B{r}="booking",", inkl. Stellplatz.","."))',
    },
}
# AG (Stellplatz) = 0 is a permanent constant (never has a value in the
# real data). S/T/V are PLACEHOLDER defaults, not real values — Farzaneh
# (2026-09-16): fill these now so the row isn't blank, then replace with
# the actual figures once cleaning has happened, based on the (future)
# cleaning schedule / cleaner table (Phase 5). V='M-Ü' mirrors the
# cleaner-code short-labels already used in the real data (e.g. 'Jen-Ü',
# 'Meh-Ü') — a specific default cleaner code, not a generic placeholder.
EXPLICIT_LITERALS = {
    "AG": 0,
    "S": 0,      # selbstfahren zum putzen
    "T": 3.5,    # Putzstunden
    "V": "M-Ü",  # cleaner code
}


def set_explicit_cells(ws, row, property_key, log=print):
    formulas = dict(EXPLICIT_FORMULAS)
    formulas.update(EXPLICIT_FORMULAS_BY_PROPERTY.get(property_key, {}))
    for letter, template in formulas.items():
        col = column_index_from_string(letter)
        ws.cell(row=row, column=col).value = template.format(r=row)
    for letter, value in EXPLICIT_LITERALS.items():
        col = column_index_from_string(letter)
        ws.cell(row=row, column=col).value = value
    log(f"    set {len(formulas)} explicit formulas + {len(EXPLICIT_LITERALS)} literal(s) on row {row}")

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

    # M (Übernachtungssteuer) is always Erwachsene × Nächte.
    set_field("tourist_tax", f"=G{row}*F{row}")
    # Z (Umsatzsteuersatz) is always 19%.
    set_field("vat_rate", DEFAULT_VAT_RATE)
    # P (Putzarbeit): fixed per-property rate for Booking.com; for Airbnb,
    # its own export's Reinigungsgebühr goes here instead (confirmed by
    # Farzaneh 2026-09-16 — same field she'd previously called
    # cleaning_fee_charged, but it belongs in P, not AF, for Airbnb).
    if str(entry.get("platform", "")).strip().lower() == "booking":
        cleaning_cost = CLEANING_COST_BOOKING.get(entry.get("property"))
    else:
        cleaning_cost = entry.get("cleaning_fee_charged")
    if cleaning_cost is not None:
        set_field("cleaning_cost", cleaning_cost)

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

    # AF (Endreinigung — cleaning fee charged to the GUEST, distinct from P
    # /Putzarbeit paid TO the cleaner) = IF(R='ja', P, P/1.07), always, for
    # every platform (confirmed by Farzaneh 2026-09-16 — same formula
    # applies whether P came from a fixed Booking.com rate or Airbnb's own
    # Reinigungsgebühr). Set unconditionally, referencing P{row} directly.
    set_field("cleaning_fee_charged", f'=IF(R{row}="ja",P{row},P{row}/1.07)')

    # AA (Nettobetrag) = the platform fee: Booking.com detailed export's
    # pure "Commission amount", OR Airbnb's "Servicegebühr" (confirmed by
    # Farzaneh 2026-09-16 as "Booking, Airbnb gebühr" — applies to both,
    # not Booking-only as an earlier version of this comment incorrectly
    # claimed). AX (Payment Charge von Booking) = the 'simple' export's
    # Commission (which combines that SAME commission with an extra payment
    # -processing fee) minus AA — isolating just the payment-processing
    # portion, which is what AX's name actually means; AX is Booking.com-
    # only (Airbnb's simple export has no such combined-fee figure to
    # subtract from). Only written when we actually have the matching
    # figure(s) — never guessed.
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

    set_explicit_cells(ws, row, entry.get("property"), log=log)

    log(f"    appended row {row}: {entry.get('guest_name')} ({entry['checkin']} -> {entry['checkout']})")
    return row


def apply_cancellation(wb, header_map_cache, entry, log=print):
    """Returns None if no matching row was found, otherwise the cancelled
    row's own checkout date — used by process_batch to best-effort flag
    the matching Putzplan row too (see putzplan_writer; Putzplan is keyed
    on Abreise/checkout, not checkin — confirmed 2026-09-22)."""
    code = str(entry.get("confirmation_code", "")).strip()
    if not code:
        log("    WARN: cancellation entry missing confirmation_code, skipping")
        return None
    for sheet_name in ["1", "2", "3", "4", FUTURE_YEAR_SHEET]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_map = header_map_cache.setdefault(
            sheet_name, _build_header_map(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        )
        code_col = header_map.get(FIELD_HEADERS["confirmation_code"])
        flag_col = header_map.get(FIELD_HEADERS["cancelled"])
        checkout_col = header_map.get(FIELD_HEADERS["checkout"])
        if code_col is None or flag_col is None:
            continue
        # Scan every row up to the sheet's real extent — NOT "stop at the
        # first blank column-A cell" (that pattern is deliberate in
        # last_used_row, for finding an APPEND spot, but wrong here: a
        # manually-deleted-cell gap row would make every reservation
        # after it unreachable and silently "not found". Found 2026-09-24
        # when a real cancellation/update past such a gap failed.
        for row in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=code_col + 1).value
            if str(cell_val).strip() == code:
                ws.cell(row=row, column=flag_col + 1, value="ja")
                log(f"    marked row {row} (sheet {sheet_name}) as Storniert (code={code})")
                return ws.cell(row=row, column=checkout_col + 1).value if checkout_col is not None else None
    log(f"    WARN: confirmation_code {code} not found, cancellation not applied")
    return None


def process_batch(new_reservations, cancellations, log=print):
    """Apply a batch of reservations/cancellations across however many
    properties they touch, backing up each touched file exactly once. Used
    by both the Cowork pipeline and the manual CSV/XLS import.

    Also updates Putzplan2026.xlsx (the cleaner schedule) — a new
    reservation gets a Putzplan row, a cancellation gets its matching row
    (if found) flagged 'storniert'. Local import to avoid a circular
    import (putzplan_writer imports from this module). Deliberately only
    for reservations THIS app processes going forward (2026-09-22 scope
    decision, see putzplan_writer.py docstring) — done AFTER the
    GästeListe files are saved, so the same-day-checkout check in
    putzplan_writer reads fresh data including this batch's own rows."""
    from . import putzplan_writer

    touched_files = {}  # property -> [path, workbook]
    header_map_cache = {}
    applied_rows = []
    skipped = []
    cancelled_info = []  # (property_key, {checkin, adults, children})

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
        checkout_value = apply_cancellation(wb, cache_for_property, c, log=log)
        if checkout_value is not None:
            cancelled_info.append((property_key, checkout_value))

    for property_key, (path, wb) in touched_files.items():
        wb.save(path)
        log(f"    saved {path}")

    for applied in applied_rows:
        putzplan_writer.append_putzplan_row(applied["entry"], log=log)
    for property_key, checkout_value in cancelled_info:
        putzplan_writer.flag_putzplan_cancelled(property_key, checkout_value, log=log)

    # Record every code this batch touched — added OR cancelled, whether or
    # not the cancellation actually found a row to flag — in known_codes.py
    # so dedup still works even after Farzaneh manually deletes the row
    # (2026-09-23 decision, see known_codes.py docstring).
    codes_by_property = {}
    for applied in applied_rows:
        codes_by_property.setdefault(applied["property"], set()).add(applied["entry"].get("confirmation_code"))
    for c in cancellations:
        property_key = c.get("property")
        if property_key in PROPERTY_FILES:
            codes_by_property.setdefault(property_key, set()).add(c.get("confirmation_code"))
    for property_key, codes in codes_by_property.items():
        known_codes.record_codes(property_key, codes)

    return {"applied": applied_rows, "skipped": skipped, "files_touched": list(touched_files.keys())}


def update_reservation_fields(property_key, confirmation_code, updates: dict, log=print):
    """Locates a row by confirmation code and updates only the given fields
    (e.g. from the manual-completion form: real guest_name/guest_email/
    adults/children replacing a placeholder). Backs up the file first.

    If 'adults' or 'children' is among the updates, also refreshes
    whichever Putzplan row shows THIS reservation as its next guests
    (2026-09-24) — Putzplan displays the upcoming reservation's headcount,
    not the departing one's, so a reservation first imported with
    incomplete data (e.g. booking_simple, no guest counts) and later
    completed via /complete or a detailed re-import would otherwise leave
    that preceding departure's row showing the old/blank counts forever."""
    path = os.path.join(DATA_DIR, PROPERTY_FILES[property_key])
    backup_file(path)
    wb = openpyxl.load_workbook(path)
    code = str(confirmation_code).strip()
    found = False
    checkin_value = final_adults = final_children = None
    for sheet_name in ["1", "2", "3", "4", FUTURE_YEAR_SHEET]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_map = _build_header_map(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        code_col = header_map.get(FIELD_HEADERS["confirmation_code"])
        checkin_col = header_map.get(FIELD_HEADERS["checkin"])
        adults_col = header_map.get(FIELD_HEADERS["adults"])
        children_col = header_map.get(FIELD_HEADERS["children"])
        if code_col is None:
            continue
        # Scan every row up to the sheet's real extent, not "stop at the
        # first blank column-A cell" — see apply_cancellation for why
        # (a manually-deleted-cell gap row made rows after it
        # unreachable; found 2026-09-24 on a real row).
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row=row, column=code_col + 1).value).strip() == code:
                checkin_value = ws.cell(row=row, column=checkin_col + 1).value if checkin_col is not None else None
                for field_key, value in updates.items():
                    header = FIELD_HEADERS.get(field_key)
                    col = header_map.get(header) if header else None
                    if col is None:
                        log(f"    WARN: column for '{field_key}' not found, skipping")
                        continue
                    ws.cell(row=row, column=col + 1).value = value
                # Read back the FINAL values (not just what's in `updates`
                # this call) — a partial update (e.g. only adults, not
                # children) would otherwise wipe the other one when synced.
                final_adults = ws.cell(row=row, column=adults_col + 1).value if adults_col is not None else None
                final_children = ws.cell(row=row, column=children_col + 1).value if children_col is not None else None
                found = True
                break
        if found:
            break
    if found:
        wb.save(path)
        log(f"    updated row for code={code} in {path}")
        if ("adults" in updates or "children" in updates) and checkin_value is not None:
            from . import putzplan_writer  # local import to avoid a circular import

            putzplan_writer.sync_putzplan_for_reservation(
                property_key,
                checkin_value,
                final_adults,
                final_children,
                code,
                log=log,
            )
    else:
        log(f"    WARN: code={code} not found in {path}, nothing updated")
    return found
