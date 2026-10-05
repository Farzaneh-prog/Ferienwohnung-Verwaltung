"""
Twilio costs (2026-10-05): a monthly cost report by e-mail and a low-balance
warning.

- check_balance(): daily (scheduler 16:00). Balance below TWILIO_LOW_BALANCE_USD
  (default 5 USD) -> important warning (WhatsApp + e-mail via owner_alerts),
  repeated every 3 days until the balance is back above the threshold. Twilio
  does not charge for reading the balance. NOTE: prepaid, auto-recharge is OFF
  by design (decision 2026-10-01) — when the balance hits zero messages fail.
- monthly_report(): on the 1st of the month (08:00 Berlin) an e-mail with last
  month's usage by category (Twilio fees; Meta's per-message fees for our
  template messages show up under "Channels") and the current balance.

Uses the Twilio REST API with the same credentials as the sender. Never raises.
"""
import datetime
import json
import os

from . import owner_alerts
from .config import DATA_DIR
from .tz import BERLIN

_STATE = os.path.join(DATA_DIR, "data", "billing_state.json")
REPEAT_DAYS = 3


def _client():
    from twilio.rest import Client

    return Client(os.environ["TWILIO_ACCOUNT_SID"], os.environ["TWILIO_AUTH_TOKEN"])


def get_balance():
    """(balance as float, currency) or None."""
    try:
        b = _client().api.v2010.accounts(os.environ["TWILIO_ACCOUNT_SID"]).balance.fetch()
        return float(b.balance), b.currency
    except Exception:  # noqa: BLE001
        return None


def _load() -> dict:
    if os.path.exists(_STATE):
        with open(_STATE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save(data: dict) -> None:
    os.makedirs(os.path.dirname(_STATE), exist_ok=True)
    with open(_STATE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1)


def check_balance(today: datetime.date = None, log=print) -> bool:
    """Returns True if the balance is fine (or unknown)."""
    today = today or datetime.datetime.now(tz=BERLIN).date()
    result = get_balance()
    if result is None:
        log("    [billing] Guthaben nicht lesbar")
        return True
    balance, currency = result
    threshold = float(os.environ.get("TWILIO_LOW_BALANCE_USD") or 5)
    log(f"    [billing] Twilio-Guthaben: {balance:.2f} {currency} (Warnschwelle {threshold:.2f})")
    state = _load()
    if balance >= threshold:
        if state.pop("last_low_alert", None):
            _save(state)
        return True
    last = state.get("last_low_alert")
    if last and (today - datetime.date.fromisoformat(last)).days < REPEAT_DAYS:
        return False
    owner_alerts.notify_owner(
        f"⚠️ Twilio-Guthaben niedrig: {balance:.2f} {currency} (Schwelle {threshold:.2f}). Bitte aufladen, "
        f"sonst werden keine WhatsApp-Nachrichten mehr zugestellt.", log=log)
    state["last_low_alert"] = today.isoformat()
    _save(state)
    return False


def monthly_report(period: str = "last_month", log=print) -> bool:
    """period: "last_month" (default, run on the 1st) or "this_month"."""
    try:
        records = getattr(_client().usage.records, period).list()
    except Exception as exc:  # noqa: BLE001
        log(f"    [billing] Nutzung nicht lesbar: {exc}")
        return False
    # Twilio lists parent categories too (they'd double-count): keep only the leaves
    parents = {"totalprice", "channels", "channels-messaging"}
    rows = [(r.description, int(float(r.count or 0)), float(r.price or 0), r.price_unit)
            for r in records if float(r.price or 0) > 0 and r.category not in parents]
    rows.sort(key=lambda x: -x[2])
    total_row = next((r for r in records if r.category == "totalprice"), None)
    total = float(total_row.price) if total_row else sum(r[2] for r in rows)
    unit = (total_row.price_unit if total_row else (rows[0][3] if rows else "usd")).upper()
    month = datetime.datetime.now(tz=BERLIN)
    if period == "last_month":
        month = (month.replace(day=1) - datetime.timedelta(days=1))
    lines = [f"Twilio-Kosten {month:%B %Y}" + (" (bisher)" if period == "this_month" else ""), ""]
    lines += [f"{count:>6}  {price:>7.2f} {u.upper()}  {desc}" for desc, count, price, u in rows] or ["(keine Kosten)"]
    lines += ["", f"Summe: {total:.2f} {unit}"]
    balance = get_balance()
    if balance:
        lines.append(f"Aktuelles Guthaben: {balance[0]:.2f} {balance[1]}")
    lines += ["", "Hinweis: Twilio berechnet 0,005 USD pro WhatsApp-Nachricht (gesendet und empfangen). "
                  "Die Meta-Gebühren für unsere Vorlagen-Nachrichten (Utility ca. 0,046 EUR) erscheinen "
                  "unter «Channels» und teils erst später."]
    return owner_alerts._send_mail(f"Twilio-Kostenbericht {month:%m/%Y}", "\n".join(lines), log=log)
