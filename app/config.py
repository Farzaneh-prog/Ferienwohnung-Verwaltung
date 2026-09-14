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
# Update this every year when a new GästeListe file is created.
PROPERTY_FILES = {
    "karlstrasse": "GästeListe_2026K.xlsx",
    "eisenach": "GästeListe_2026_Pf.xlsx",
}

PROPERTY_LABELS = {
    "karlstrasse": "Marktresidenz Karlstraße",
    "eisenach": "Marktresidenz Eisenach",
}

# Sheets that hold actual reservation rows (one per quarter).
QUARTER_SHEETS = ["1", "2", "3", "4"]

# Sheets that are summaries/templates, never reservation data — always skipped.
NON_DATA_SHEETS = {
    "Übersicht",
    "ÜbersichtSteuer",
    "zwischenjährlich",
    "Umsatzsteuer",
    "Muster",
    "Tabelle3",
}


def resolve_property_path(property_key: str) -> str:
    """Return the absolute path to a property's Excel file."""
    filename = PROPERTY_FILES[property_key]
    return os.path.join(DATA_DIR, filename)
