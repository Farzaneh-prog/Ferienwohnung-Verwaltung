"""
The decoupled "cleaner heads-up" list — a few-second, property+date-only
note so the (future, Phase 5) cleaner-scheduling module has something to
react to right away, before the full reservation gets entered via CSV
import. Deliberately NOT part of the real GästeListe_*.xlsx files — see
project discussion: keeping it as its own small list avoids needing to
touch the Excel file (and its formula-safety rules) for a 5-second note,
and avoids reconciliation headaches by just staying manually resolved once
the real row shows up.
"""
import datetime
import json
import os
import uuid

from .config import DATA_DIR

ALERTS_FILE = os.path.join(DATA_DIR, "data", "quick_alerts.json")


def _load():
    if not os.path.exists(ALERTS_FILE):
        return []
    with open(ALERTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(alerts):
    os.makedirs(os.path.dirname(ALERTS_FILE), exist_ok=True)
    with open(ALERTS_FILE, "w", encoding="utf-8") as f:
        json.dump(alerts, f, ensure_ascii=False, indent=2)


def list_alerts(include_resolved=False):
    alerts = _load()
    if not include_resolved:
        alerts = [a for a in alerts if not a.get("resolved")]
    return sorted(alerts, key=lambda a: a.get("date") or "")


def add_alert(property_key, date_str, note=""):
    alerts = _load()
    alert = {
        "id": uuid.uuid4().hex[:8],
        "property": property_key,
        "date": date_str,
        "note": note,
        "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "resolved": False,
    }
    alerts.append(alert)
    _save(alerts)
    return alert


def resolve_alert(alert_id):
    alerts = _load()
    for a in alerts:
        if a["id"] == alert_id:
            a["resolved"] = True
    _save(alerts)
