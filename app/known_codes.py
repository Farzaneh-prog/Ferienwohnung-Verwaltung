"""
Persistent record of every confirmation code this app has ever processed
(added OR cancelled) — confirmed with Farzaneh 2026-09-23: she wants to be
able to physically delete a cancelled reservation's row from GästeListe
(so the yearly SUM() totals — which don't exclude Storniert=ja rows —
stay correct without needing ~271 formulas across both files rewritten
to SUMIFS) WITHOUT the deleted code ever looking "new" again if an old
export gets re-uploaded later.

Deleting the row erases the only place a confirmation code was tracked;
this file is a second, independent place that a manual Excel delete can
never touch, so dedup (excel_reader.get_existing_confirmation_codes)
keeps working even after the row is gone. Deliberately chosen over
fixing the SUM formulas — this touches no formula, works immediately on
already-deleted rows too, and carries no tax-reporting risk.
"""
import json
import os

from .config import DATA_DIR

_PATH = os.path.join(DATA_DIR, "data", "known_codes.json")


def _load() -> dict:
    if not os.path.exists(_PATH):
        return {}
    with open(_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(data: dict) -> None:
    os.makedirs(os.path.dirname(_PATH), exist_ok=True)
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)


def get_known_codes(property_key: str) -> set:
    return set(_load().get(property_key, []))


def record_codes(property_key: str, codes) -> None:
    """codes: any iterable of confirmation codes (None/empty entries are
    dropped). No-op if there's nothing real to add."""
    clean = {str(c).strip() for c in codes if c not in (None, "")}
    if not clean:
        return
    data = _load()
    existing = set(data.get(property_key, []))
    data[property_key] = sorted(existing | clean)
    _save(data)
