"""
Cleaner roster: who's in the WhatsApp coordination chain, which property
they cover, which weekdays they're available, their pay rate, contract
expiry, and their tier in the priority chain (سند-معماری, "الگوریتم
نهایی", بخش ۴: 1=جنیفر/مهرناز, 2=Manuela, 3=تهمینه, 4=Ramic).

Stored in data/cleaners.json (gitignored — real names/rates), so it can be
edited by hand without a code change. This is the phase-3 stage-0 MVP
(2026-09-30) of the architecture doc's "روستر قابل‌تنظیم، نه هاردکد"
principle — a real settings-dashboard page is future work (phase 1
frontend hasn't been built yet), a hand-editable JSON file is the
starting point.

available_days uses Python's date.weekday() convention (0=Montag,
6=Sonntag) — same as the rest of this app (see xlsx_writer.WEEKDAYS_DE).

DEFAULT_ROSTER seeds the file the first time it's read, using the rates
Farzaneh already confirmed in نظافتچی_ها-و-فرمول_های-مالی.md. WhatsApp
numbers are deliberately left blank (None) — nobody's real number has
been collected yet; Farzaneh fills those in by hand before phase 3 stage 1
(actual sending) can work.
"""
import json
import os

from .config import DATA_DIR

_PATH = os.path.join(DATA_DIR, "data", "cleaners.json")

ALL_DAYS = [0, 1, 2, 3, 4, 5, 6]
WEEKEND = [5, 6]  # Samstag, Sonntag
NOT_SUNDAY = [0, 1, 2, 3, 4, 5]

# Rates match نظافتچی_ها-و-فرمول_های-مالی.md's table exactly (2026-09-16
# confirmed version). Manuela's Eisenach rate is None — she only covers
# Karlstraße.
DEFAULT_ROSTER = {
    "jennifer": {
        "name": "Jennifer",
        "code": "Jen-Ü",
        "sheet_name": "Jennifer",
        "whatsapp_number": None,
        "properties": ["karlstrasse", "eisenach"],
        "available_days": ALL_DAYS,
        "tier": 1,
        "rate_karlstrasse": 55.26,
        "rate_eisenach": 72.78,
        "contract_until": None,
    },
    "mehrnaz": {
        "name": "Mehrnaz",
        "code": "Meh-Ü",
        "sheet_name": "Mehrnaz",
        "whatsapp_number": None,
        "properties": ["karlstrasse", "eisenach"],
        "available_days": ALL_DAYS,
        "tier": 1,
        "rate_karlstrasse": 51.22,
        "rate_eisenach": 65.58,
        "contract_until": "2026-10-31",
    },
    "manuela": {
        "name": "Manuela",
        "code": "M-Ü",
        "sheet_name": "Manuela",
        "whatsapp_number": None,
        "properties": ["karlstrasse"],
        "available_days": NOT_SUNDAY,
        "tier": 2,
        "rate_karlstrasse": 12.0,  # €/Stunde, variabel — kein fixer Pauschalbetrag
        "rate_eisenach": None,
        "contract_until": None,
    },
    "tahmine": {
        "name": "Tahmine",
        "code": "Tah-Ü",
        "sheet_name": "Tahmina",
        "whatsapp_number": None,
        # Rate table says "meistens Sa/So" but the architecture doc's
        # automated chain condition (بخش ۴) is stricter: "فقط اگه
        # روز=شنبه/یکشنبه" — following the algorithm's explicit rule here,
        # not the softer rate-table wording.
        "properties": ["karlstrasse", "eisenach"],
        "available_days": WEEKEND,
        "tier": 3,
        "rate_karlstrasse": 38.0,
        "rate_eisenach": 50.0,
        "contract_until": "2027-01-31",
    },
    "ramic": {
        "name": "Ramic (Ramesch)",
        "code": "R-Ü",
        "sheet_name": "Ramic",
        "whatsapp_number": None,
        "properties": ["karlstrasse", "eisenach"],
        "available_days": WEEKEND,
        "tier": 4,
        "rate_karlstrasse": 45.0,
        "rate_eisenach": 60.0,
        "contract_until": None,
    },
}


def load_roster() -> dict:
    """{cleaner_id: {...}}. Creates data/cleaners.json with DEFAULT_ROSTER
    the first time this is called, so there's always a real file to edit."""
    if not os.path.exists(_PATH):
        save_roster(DEFAULT_ROSTER)
    with open(_PATH, "r", encoding="utf-8") as f:
        roster = json.load(f)
    only = os.environ.get("SCHEDULER_ONLY_CLEANERS", "").strip()
    if only:  # TEST MODE (see cleaner_coordination.test_mode): hide everybody else
        allowed = {c.strip() for c in only.split(",") if c.strip()}
        roster = {cid: c for cid, c in roster.items() if cid in allowed}
    return roster


# Cleaning hours per property (karlstrasse, eisenach) from
# نظافتچی_ها-و-فرمول_های-مالی.md. The pay is a FIXED amount per property
# (rate_karlstrasse/rate_eisenach in the roster) — except Manuela, who is paid
# per hour (12 €) and whose hours vary: default 3.5, corrected by hand.
CLEANING_HOURS = {
    "jennifer": (3.016, 3.97),
    "mehrnaz": (3.016, 3.97),
    "tahmine": (3.016, 3.97),
    "ramic": (3.0, 4.0),
    "manuela": (3.5, None),
}


def cleaning_pay(cleaner_id: str, cleaner: dict, property_key: str):
    """(hours, amount, hourly_rate) for GästeListe columns T/U. Fixed-pay
    cleaners: amount = their rate for the property, hourly_rate None. Hourly
    (Manuela): amount None (U stays the formula =T*rate), hourly_rate = rate.
    Returns None if the roster has no rate/hours for this cleaner+property."""
    index = 0 if property_key == "karlstrasse" else 1
    rate = cleaner.get(f"rate_{property_key}")
    hours = cleaner.get(f"hours_{property_key}") or CLEANING_HOURS.get(cleaner_id, (None, None))[index]
    if rate is None or hours is None:
        return None
    if cleaner.get("hourly") or cleaner_id == "manuela":
        return hours, None, rate
    return hours, round(float(rate), 2), None


def sheet_label(cleaner: dict) -> str:
    """What goes into Putzplan column A for this cleaner. The real sheet uses
    first names (Mehrnaz, Manuela, Jennifer, Ramic, "Tahmina" — note the
    spelling), not the GästeListe codes (Jen-Ü, ...)."""
    return cleaner.get("sheet_name") or cleaner["name"]


def match_cleaner(roster: dict, text):
    """cleaner_id for a Putzplan column-A value (or None): case-insensitive,
    matched against id, name, sheet_name, code and optional "aliases". Free
    text that is no single name ("ich", "Mehrnaz-Jennifer", notes) -> None."""
    wanted = str(text or "").strip().lower()
    if not wanted:
        return None
    for cid, c in roster.items():
        names = {cid, c.get("name"), c.get("sheet_name"), c.get("code"), *c.get("aliases", [])}
        if wanted in {str(n).strip().lower() for n in names if n}:
            return cid
    return None


def save_roster(roster: dict) -> None:
    os.makedirs(os.path.dirname(_PATH), exist_ok=True)
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(roster, f, ensure_ascii=False, indent=2, sort_keys=True)
