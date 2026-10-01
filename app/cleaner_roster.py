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
        return json.load(f)


def save_roster(roster: dict) -> None:
    os.makedirs(os.path.dirname(_PATH), exist_ok=True)
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(roster, f, ensure_ascii=False, indent=2, sort_keys=True)
