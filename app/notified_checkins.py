"""
Persistent record of which reservations have already gotten a check-in
reminder email, so the daily scheduler job (app/scheduler.py) never sends
the same reminder twice even if it runs more than once on the same day
(e.g. a container restart) or a reservation still matches "checkin is
tomorrow" on a later run.

Same JSON-file-next-to-known_codes.py pattern as app/known_codes.py.
"""
import json
import os

from .config import DATA_DIR

_PATH = os.path.join(DATA_DIR, "data", "notified_checkins.json")


def _load() -> dict:
    if not os.path.exists(_PATH):
        return {}
    with open(_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(data: dict) -> None:
    os.makedirs(os.path.dirname(_PATH), exist_ok=True)
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)


def was_notified(property_key: str, reservation_key: str) -> bool:
    return reservation_key in set(_load().get(property_key, []))


def record_notified(property_key: str, reservation_key: str) -> None:
    data = _load()
    existing = set(data.get(property_key, []))
    existing.add(reservation_key)
    data[property_key] = sorted(existing)
    _save(data)
