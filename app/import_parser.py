"""
Parses the official reservation exports from Booking.com and Airbnb into a
common shape (matching the same fields used by tools/apply_incoming.py, so
both the Cowork pipeline and this manual-upload pipeline share one apply
step). No scraping, no AI reading a screenshot — these are the platforms'
own structured export files, downloaded by hand whenever convenient
(weekly, every few days, whatever). See docs/data-sources.md for how each
field was verified against a real reservation before being trusted.

None of these exports include the guest's name or email — confirmed absent
from all four real sample files during development. Those two fields always
need the manual-completion step (see app/routes_import.py).
"""
import csv
import datetime
import io

import xlrd

PLACEHOLDER_PREFIX = "Gast (Code "


def guest_placeholder(confirmation_code: str) -> str:
    return f"{PLACEHOLDER_PREFIX}{confirmation_code})"


def is_placeholder_name(name) -> bool:
    return isinstance(name, str) and name.startswith(PLACEHOLDER_PREFIX)


def detect_property(text: str):
    if not text:
        return None
    t = text.lower()
    if "karlstra" in t:
        return "karlstrasse"
    if "eisenach" in t:
        return "eisenach"
    # Booking.com's "Unit type" cell for Eisenach/Pfarrberg is a marketing
    # name with none of the words above in it — confirmed with Farzaneh
    # 2026-09-22 after 8 real reservation rows silently got skipped (no
    # error, just dropped) because this text matched neither check. Airbnb's
    # own listing names already contain "Eisenach" literally, so this is
    # Booking-specific.
    if "traumwebers r" in t and "feenparadies" in t:
        return "eisenach"
    return None


def _parse_money(value):
    """'122.09 EUR' -> 122.09; a plain number stays a number; blank -> None."""
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    s = str(value).strip().strip('"')
    if not s:
        return None
    s = s.split(" ")[0].replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def _xlrd_rows(path):
    wb = xlrd.open_workbook(path)
    sheet = wb.sheets()[0]
    header = [str(sheet.cell_value(0, c)).strip() for c in range(sheet.ncols)]
    for r in range(1, sheet.nrows):
        row = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
        if not any(str(v).strip() for v in row):
            continue
        yield dict(zip(header, row))


def parse_booking_simple(path):
    """Columns: Property name, Location, Booker name, Genius booker, Arrival,
    Departure, Booked on, Status, Total payment, Commission, Currency,
    Reservation number."""
    reservations, cancellations = [], []
    for row in _xlrd_rows(path):
        property_key = detect_property(row.get("Property name", ""))
        code = row.get("Reservation number")
        code = str(int(code)) if isinstance(code, float) else str(code).strip()
        if not property_key or not code:
            continue
        status = str(row.get("Status", "")).strip().lower()
        if "cancel" in status:
            cancellations.append({"property": property_key, "confirmation_code": code})
            continue
        arrival = datetime.datetime.strptime(row["Arrival"], "%d %B %Y").date()
        departure = datetime.datetime.strptime(row["Departure"], "%d %B %Y").date()
        # "Booker name" deliberately ignored (Farzaneh's call — not the field
        # she trusts/uses; only "Guest name(s)" from the detailed Booking.com
        # export and "Gast" from Airbnb count as a real guest name source).
        reservations.append({
            "property": property_key,
            "platform": "booking",
            "guest_name": guest_placeholder(code),
            "checkin": arrival.isoformat(),
            "checkout": departure.isoformat(),
            "adults": None,
            "children": None,
            "confirmation_code": code,
            "guest_paid_total": _parse_money(row.get("Total payment")),
            # "Commission" here is the COMBINED figure (commission + payment
            # processing fee) — confirmed by Farzaneh against the detailed
            # export's pure "Commission amount" for the same booking; kept
            # separate on purpose, see xlsx_writer.append_reservation.
            "platform_fee_combined": _parse_money(row.get("Commission")),
            "source_file": "booking_simple",
        })
    return reservations, cancellations


def parse_booking_detailed(path):
    """Columns: Book number, Booked by, Guest name(s), Check-in, Check-out,
    Booked on, Status, Rooms, Persons, Adults, Children, Children's age(s),
    Price, Commission %, Commission amount, Payment status, Payment method
    (...), Remarks, Booker country, Travel purpose, Device, Unit type,
    Duration (nights), Cancellation date, Address, Phone number.
    Phone number is deliberately never read (not needed, per Farzaneh)."""
    reservations, cancellations = [], []
    for row in _xlrd_rows(path):
        property_key = detect_property(row.get("Unit type", ""))
        code = row.get("Book number")
        code = str(int(code)) if isinstance(code, float) else str(code).strip()
        if not property_key or not code:
            continue
        status = str(row.get("Status", "")).strip().lower()
        cancelled = "cancel" in status or bool(str(row.get("Cancellation date", "")).strip())
        if cancelled:
            cancellations.append({"property": property_key, "confirmation_code": code})
            continue
        checkin = str(row["Check-in"]).strip()
        checkout = str(row["Check-out"]).strip()
        guest_name = str(row.get("Guest name(s)", "")).strip() or guest_placeholder(code)
        adults = row.get("Adults")
        children = row.get("Children")
        reservations.append({
            "property": property_key,
            "platform": "booking",
            "guest_name": guest_name,
            "checkin": checkin,
            "checkout": checkout,
            "adults": int(adults) if isinstance(adults, (int, float)) else None,
            "children": int(children) if isinstance(children, (int, float)) else 0,
            "confirmation_code": code,
            "guest_paid_total": _parse_money(row.get("Price")),
            # The PURE platform commission (unlike booking_simple's combined
            # "Commission", which also folds in a payment-processing fee) —
            # goes to Nettobetrag/AA. See xlsx_writer.append_reservation.
            "platform_fee_total": _parse_money(row.get("Commission amount")),
            "extra": {
                "commission_percent": row.get("Commission %"),
                "payment_method": row.get("Payment method (Payment Provider)"),
                "booker_country": row.get("Booker country"),
            },
            "source_file": "booking_detailed",
        })
    return reservations, cancellations


def _decode_bytes(raw):
    """Airbnb's export encoding changed 2026-09-22 (was cp1252, now UTF-8
    with a BOM — same day the delimiter also changed, see below) without
    warning; try UTF-8 first (a genuine cp1252 file will almost always
    contain a byte sequence invalid as UTF-8 and raise), fall back to
    cp1252 for older files so both still work."""
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252")


def _read_text_rows(path):
    """Airbnb also switched delimiter 2026-09-22 — the older export was
    tab-separated (no quoting seen), the new one is a real comma-separated
    CSV with quoted fields, some containing embedded commas AND embedded
    newlines (e.g. a multi-line note) — so this can't just split on a
    hardcoded character/line-by-line anymore. Sniff comma vs tab from the
    first physical line, then hand the WHOLE decoded text to the csv
    module via StringIO so it does its own (quote-aware) line splitting —
    pre-splitting on \\r\\n ourselves first, as an earlier version of this
    function did, breaks any field with a real newline inside its quotes."""
    with open(path, "rb") as f:
        raw = f.read()
    text = _decode_bytes(raw)
    first_line = text.splitlines()[0] if text else ""
    delimiter = "\t" if "\t" in first_line else ","
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    header = next(reader)
    for values in reader:
        if not any(v.strip() for v in values):
            continue
        values = [v.strip() for v in values]
        if len(values) < len(header):
            values += [""] * (len(header) - len(values))
        yield dict(zip(header, values))


def parse_airbnb_csv(path):
    """Handles both the 'past' export (with Ausgezahlt/payout columns) and
    the 'pending' export (upcoming bookings, no payout yet) — same core
    columns either way. Skips 'Payout' summary rows, only keeps 'Buchung'
    (booking) rows.

    Field mapping verified against a real Airbnb 'Einkünfte' breakdown
    (see docs/cowork-command.md history): Betrag = host payout,
    Servicegebühr = total platform fee (incl. VAT), Reinigungsgebühr =
    cleaning fee charged to guest. guest_paid_total is reconstructed as
    Betrag + Servicegebühr (verified to match exactly: 168.81 + 38.19 =
    207.00 against the real screenshot) — this needs the GROSS
    Servicegebühr, since the guest pays the gross amount.

    platform_fee_total (which feeds AA/Nettobetrag in xlsx_writer) is
    different: confirmed by Farzaneh 2026-09-24 that Servicegebühr is
    ALWAYS gross (incl. 19% VAT) — writing it straight into AA made
    xlsx_writer's own AB formula (=AA*0.19) double-charge the VAT that
    was already baked in (e.g. a real row: Servicegebühr 30.81 was
    written to AA, when the correct net figure is 30.81/1.19 = 25.89).
    So platform_fee_total is divided by 1.19 here; guest_paid_total above
    still uses the raw (gross) Servicegebühr, deliberately."""
    reservations = []
    for row in _read_text_rows(path):
        if row.get("Typ", "").strip() != "Buchung":
            continue
        property_key = detect_property(row.get("Inserat", ""))
        code = row.get("Bestätigungs-Code", "").strip()
        if not property_key or not code:
            continue
        checkin = row.get("Startdatum", "").strip()
        checkout = row.get("Enddatum", "").strip()
        try:
            checkin_iso = datetime.datetime.strptime(checkin, "%m/%d/%Y").date().isoformat()
            checkout_iso = datetime.datetime.strptime(checkout, "%m/%d/%Y").date().isoformat()
        except ValueError:
            continue
        guest_name = row.get("Gast", "").strip() or guest_placeholder(code)
        host_payout = _parse_money(row.get("Betrag"))
        platform_fee_gross = _parse_money(row.get("Servicegebühr"))
        platform_fee_total = round(platform_fee_gross / 1.19, 2) if platform_fee_gross is not None else None
        guest_paid_total = None
        if host_payout is not None and platform_fee_gross is not None:
            guest_paid_total = round(host_payout + platform_fee_gross, 2)
        reservations.append({
            "property": property_key,
            "platform": "Airbnb",
            "guest_name": guest_name,
            "checkin": checkin_iso,
            "checkout": checkout_iso,
            "adults": None,
            "children": None,
            "confirmation_code": code,
            "guest_paid_total": guest_paid_total,
            "host_payout": host_payout,
            "platform_fee_total": platform_fee_total,
            "cleaning_fee_charged": _parse_money(row.get("Reinigungsgebühr")),
            "extra": {
                "booking_date": row.get("Buchungsdatum", "").strip(),
                "gross_income_reported": row.get("Bruttoeinkünfte", "").strip(),
                "airbnb_collected_tax": row.get("Von Airbnb abgeführte Steuer", "").strip(),
            },
            "source_file": "airbnb_csv",
        })
    return reservations, []  # Airbnb export has no cancellation status column


def merge_reservation_entries(entries):
    """Field-level merge of multiple parsed entries for the same
    confirmation code — e.g. the same Booking.com reservation appears in
    both the 'simple' and 'detailed' exports, each with different real
    numbers (Commission vs. Commission amount are NOT the same figure).
    Picking one file over the other would silently drop real data, so this
    merges field-by-field: first non-empty value wins, except guest_name
    prefers a real name over a placeholder wherever it appears."""
    merged = {}
    for e in entries:
        for k, v in e.items():
            if k == "extra":
                merged.setdefault("extra", {}).update(v or {})
                continue
            if v in (None, ""):
                continue
            current = merged.get(k)
            if current in (None, ""):
                merged[k] = v
            elif k == "guest_name" and is_placeholder_name(current) and not is_placeholder_name(v):
                merged[k] = v
    return merged


def detect_and_parse(path, filename):
    """Sniff which of the 3 known export formats this is, parse it, and
    return (reservations, cancellations, kind). Raises ValueError with a
    clear message if the format isn't recognized — never guesses."""
    lower = filename.lower()
    if lower.endswith(".xls"):
        wb = xlrd.open_workbook(path)
        header = set(str(wb.sheets()[0].cell_value(0, c)).strip() for c in range(wb.sheets()[0].ncols))
        if "Reservation number" in header:
            res, can = parse_booking_simple(path)
            return res, can, "booking_simple"
        if "Book number" in header:
            res, can = parse_booking_detailed(path)
            return res, can, "booking_detailed"
        raise ValueError(f"Unrecognized .xls export format in {filename} (headers: {sorted(header)})")
    if lower.endswith(".csv"):
        with open(path, "rb") as f:
            first_line = _decode_bytes(f.readline())
        if "Bestätigungs-Code" in first_line and "Typ" in first_line:
            res, can = parse_airbnb_csv(path)
            return res, can, "airbnb_csv"
        raise ValueError(f"Unrecognized .csv export format in {filename}")
    raise ValueError(f"Unsupported file type: {filename} (expected .xls or .csv)")
