"""
Central, code-free configuration for which Excel file belongs to which property.

Nothing here should ever need a code change just because a new year's file
shows up — only this dict (or the .env DATA_DIR override) changes.
"""
import os

# Folder that holds the real GästeListe_*.xlsx files.
# Defaults to the project root (this file's grandparent directory).
DATA_DIR = os.environ.get(
    "DATA_DIR",
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
)

# property_key -> filename (relative to DATA_DIR)
# Update this every year when a new GästeListe file is created — PROPERTY_FILE_YEAR
# too (it's what "belongs in Muster instead of a quarter sheet" is decided against).
PROPERTY_FILES = {
    "karlstrasse": "GästeListe_2026K.xlsx",
    "eisenach": "GästeListe_2026_Pf.xlsx",
}
PROPERTY_FILE_YEAR = 2026

PROPERTY_LABELS = {
    "karlstrasse": "Marktresidenz Karlstraße",
    "eisenach": "Marktresidenz Eisenach",
}

# Sheets that hold actual reservation rows (one per quarter).
QUARTER_SHEETS = ["1", "2", "3", "4"]

# "Muster" ("template") is a misleading name — Farzaneh confirmed it's
# actually where reservations for a *future* year get held (2026-09-16),
# since the next year's own GästeListe file doesn't exist yet when they're
# booked. It IS a real data sheet, not a template — must be read like the
# quarter sheets, not skipped.
FUTURE_YEAR_SHEET = "Muster"

# Sheets that are summaries, never reservation data — always skipped.
NON_DATA_SHEETS = {
    "Übersicht",
    "ÜbersichtSteuer",
    "zwischenjährlich",
    "Umsatzsteuer",
    "Tabelle3",
}


def resolve_property_path(property_key: str) -> str:
    """Return the absolute path to a property's Excel file."""
    filename = PROPERTY_FILES[property_key]
    return os.path.join(DATA_DIR, filename)
