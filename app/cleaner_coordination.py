"""
Phase 3 — WhatsApp cleaner coordination. Stage 0 (2026-09-30) built and
tested the decision logic (WHO should be messaged WHEN for a given
Putzplan row) entirely without sending anything real. Stage 1
(2026-09-30, same day) plugged in the actual Twilio send
(app/whatsapp_sender.py) at run_daily_check's send_message/broadcast
branches — only fires when dry_run=False; everything else (roster,
eligibility, escalation chain, state) was unchanged by that switch, as
planned.

Algorithm per سند-معماری-سیستم-مهمانخانه.md, بخش ۴ ("الگوریتم نهایی"):
- Priority chain: tier 1 (Jennifer/Mehrnaz, fair rotation) -> tier 2
  (Manuela, Karlstraße + not Sunday) -> tier 3 (Tahmine, Sat/Sun) ->
  tier 4 (Ramic, Sat/Sun). Each tier is skipped if nobody in it is
  eligible for that property/weekday/contract-date/leave-week.
- Timeline (days remaining until the Putzplan row's Abreise/checkout
  date, since that's what Putzplan is keyed on):
    > 10 days: nothing automated (manual assignment only)
    <= 10, > 3 days: sequential chain, ~1h apart between attempts
    <= 3 days: "Sicherheitsnetz" — broadcast to everyone still eligible
    < 0 days, still unresolved: should not happen if broadcast worked;
      surfaced as an owner alert same as "nobody eligible at all".
  The doc's exact wording (10/7/2-3 day breakpoints for tier-by-tier
  attempts) is simplified here into two windows (sequential vs
  broadcast) — refine once stage 1 shows whether the finer-grained
  timing actually matters in practice.
- Quiet hours (Farzaneh, 2026-09-30): no outgoing message between 22:00
  and 08:00 Europe/Berlin, to anyone in the chain, at any tier. A
  send_message/broadcast that would otherwise fire during that window is
  held back — decide_next_action just returns "none" for it — until the
  next check at/after 08:00 actually sends it. Because record_attempt is
  only ever called with a real (already-≥08:00) `now`, the cooldown timer
  naturally starts counting from 08:00, not from whenever the decision
  would have fired overnight — no separate "activate the counter" step
  needed beyond simply not sending yet.

State lives in data/cleaner_coordination_state.json (gitignored):
{
  "rows": {"<property>|<checkout ISO date>": {
      "attempts": [{"cleaner_id", "tier", "stage", "sent_at"}],
      "confirmed_cleaner_id": None,
  }},
  "level1_rotation": {"last_first": "jennifer"},
  "leave_weeks": {"jennifer": [["2026-10-06", "2026-10-13"]], "mehrnaz": []},
}
"""
import datetime
import json
import os

from .config import DATA_DIR, PROPERTY_LABELS
from .cleaner_roster import load_roster
from .tz import BERLIN
from . import whatsapp_sender

_STATE_PATH = os.path.join(DATA_DIR, "data", "cleaner_coordination_state.json")

LEVEL1_START_DAYS = 10   # first automated attempt
BROADCAST_START_DAYS = 3  # "Sicherheitsnetz" — message everyone eligible at once
RETRY_COOLDOWN = datetime.timedelta(hours=1)

QUIET_HOURS_START = datetime.time(22, 0)  # Europe/Berlin
QUIET_HOURS_END = datetime.time(8, 0)


def _in_quiet_hours(now: datetime.datetime) -> bool:
    """22:00 -> 08:00 Europe/Berlin, wrapping past midnight. `now` must
    already be in Berlin local time (naive or BERLIN-aware — only the
    wall-clock time-of-day is used)."""
    t = now.timetz().replace(tzinfo=None) if now.tzinfo else now.time()
    return t >= QUIET_HOURS_START or t < QUIET_HOURS_END


def _state_key(property_key: str, checkout_date: datetime.date) -> str:
    return f"{property_key}|{checkout_date.isoformat()}"


def _load_state() -> dict:
    if not os.path.exists(_STATE_PATH):
        return {"rows": {}, "level1_rotation": {"last_first": None}, "leave_weeks": {}}
    with open(_STATE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_state(state: dict) -> None:
    os.makedirs(os.path.dirname(_STATE_PATH), exist_ok=True)
    with open(_STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2, sort_keys=True)


def _is_eligible(cleaner_id: str, cleaner: dict, property_key: str, date: datetime.date, state: dict) -> bool:
    if property_key not in cleaner["properties"]:
        return False
    if date.weekday() not in cleaner["available_days"]:
        return False
    contract_until = cleaner.get("contract_until")
    if contract_until and date > datetime.date.fromisoformat(contract_until):
        return False
    for start, end in state.get("leave_weeks", {}).get(cleaner_id, []):
        if datetime.date.fromisoformat(start) <= date <= datetime.date.fromisoformat(end):
            return False
    return True


def eligible_chain(property_key: str, date: datetime.date, roster: dict, state: dict) -> list:
    """Ordered list of cleaner_ids: eligible tier-1 person(s) first (fair
    rotation between Jennifer/Mehrnaz if both eligible), then tier 2, 3, 4
    in order."""
    eligible = [(cid, c) for cid, c in roster.items() if _is_eligible(cid, c, property_key, date, state)]
    tier1 = sorted(cid for cid, c in eligible if c["tier"] == 1)
    rest = sorted((cid for cid, c in eligible if c["tier"] != 1), key=lambda cid: roster[cid]["tier"])

    if len(tier1) == 2:
        last_first = state.get("level1_rotation", {}).get("last_first")
        if last_first in tier1:
            tier1 = [cid for cid in tier1 if cid != last_first] + [last_first]

    return tier1 + rest


def decide_next_action(property_key: str, checkout_date: datetime.date, roster: dict, state: dict,
                        today: datetime.date, now: datetime.datetime) -> dict:
    """Pure function — does NOT mutate state. Returns one of:
    {"action": "none", "reason": ...}
    {"action": "send_message", "cleaner_id", "tier"}
    {"action": "broadcast", "cleaner_ids": [...]}
    {"action": "alert_owner", "reason": ...}
    """
    row_key = _state_key(property_key, checkout_date)
    row_state = state["rows"].get(row_key, {"attempts": [], "confirmed_cleaner_id": None})

    if row_state.get("confirmed_cleaner_id"):
        return {"action": "none", "reason": "confirmed"}

    days_until = (checkout_date - today).days
    if days_until > LEVEL1_START_DAYS:
        return {"action": "none", "reason": "too_early"}

    chain = eligible_chain(property_key, checkout_date, roster, state)
    attempts = row_state.get("attempts", [])
    tried_ids = [a["cleaner_id"] for a in attempts]

    if days_until < 0:
        return {"action": "alert_owner", "reason": "checkout_passed_unresolved"}

    if days_until <= BROADCAST_START_DAYS:
        if not chain:
            return {"action": "alert_owner", "reason": "no_eligible_cleaner"}
        already_broadcast = any(a.get("stage") == "broadcast" for a in attempts)
        if already_broadcast:
            return {"action": "none", "reason": "broadcast_already_sent"}
        if _in_quiet_hours(now):
            return {"action": "none", "reason": "quiet_hours"}
        return {"action": "broadcast", "cleaner_ids": chain}

    untried = [cid for cid in chain if cid not in tried_ids]
    if not untried:
        if not chain:
            return {"action": "alert_owner", "reason": "no_eligible_cleaner"}
        return {"action": "none", "reason": "chain_exhausted_awaiting_broadcast_window"}

    if attempts:
        last_sent_at = datetime.datetime.fromisoformat(attempts[-1]["sent_at"])
        if now - last_sent_at < RETRY_COOLDOWN:
            return {"action": "none", "reason": "cooldown"}

    if _in_quiet_hours(now):
        return {"action": "none", "reason": "quiet_hours"}

    next_id = untried[0]
    return {"action": "send_message", "cleaner_id": next_id, "tier": roster[next_id]["tier"]}


def record_attempt(state: dict, row_key: str, cleaner_id: str, tier: int, sent_at: datetime.datetime,
                    stage: str = None) -> None:
    row_state = state["rows"].setdefault(row_key, {"attempts": [], "confirmed_cleaner_id": None})
    is_first_attempt = len(row_state["attempts"]) == 0
    row_state["attempts"].append({
        "cleaner_id": cleaner_id,
        "tier": tier,
        "stage": stage or ("level1" if tier == 1 else "chain"),
        "sent_at": sent_at.isoformat(),
    })
    if is_first_attempt and tier == 1:
        state.setdefault("level1_rotation", {})["last_first"] = cleaner_id


def find_pending_row_for_cleaner(cleaner_id: str, state: dict):
    """Returns (property_key, checkout_date) for the row this cleaner was
    most recently asked about that's still unconfirmed, or None. Used by
    app/whatsapp_webhook.py to figure out which row a 'Ja' reply is about
    — the reply itself doesn't carry that context, only who sent it.
    Picks the row with the EARLIEST checkout_date among matches (most
    urgent) in the rare case more than one is pending for the same
    cleaner at once."""
    candidates = []
    for row_key, row_state in state.get("rows", {}).items():
        if row_state.get("confirmed_cleaner_id"):
            continue
        if any(a["cleaner_id"] == cleaner_id for a in row_state.get("attempts", [])):
            property_key, date_str = row_key.split("|", 1)
            candidates.append((datetime.date.fromisoformat(date_str), property_key))
    if not candidates:
        return None
    candidates.sort()
    checkout_date, property_key = candidates[0]
    return property_key, checkout_date


def confirm_cleaner(property_key: str, checkout_date: datetime.date, cleaner_id: str, log=print) -> bool:
    """Call when a cleaner replies 'Ja' — marks the row confirmed in state
    AND writes their code into Putzplan column A (app/putzplan_writer.
    assign_cleaner). Returns False (no-op) if the row was already
    confirmed by someone else in the meantime (race between two tiers'
    replies — first one wins, matches the architecture doc's "به محض
    دریافت پاسخ «بله» از هرکس، چرخش/ارسال متوقف می‌شه")."""
    from . import putzplan_writer

    row_key = _state_key(property_key, checkout_date)
    state = _load_state()
    row_state = state["rows"].setdefault(row_key, {"attempts": [], "confirmed_cleaner_id": None})
    if row_state.get("confirmed_cleaner_id"):
        log(f"    [cleaner-coordination] {row_key}: 'Ja' von {cleaner_id} ignoriert — schon bestätigt "
            f"durch {row_state['confirmed_cleaner_id']}")
        return False

    row_state["confirmed_cleaner_id"] = cleaner_id
    _save_state(state)

    roster = load_roster()
    cleaner_code = roster.get(cleaner_id, {}).get("code", cleaner_id)
    putzplan_writer.assign_cleaner(property_key, checkout_date, cleaner_code, log=log)
    log(f"    [cleaner-coordination] {row_key}: bestätigt durch {cleaner_id} ({cleaner_code})")
    return True


def _send_request(cleaner: dict, property_key: str, checkout_date: datetime.date) -> str:
    """Sends the cleaner-request Content Template (required for a cold
    first contact — see whatsapp_sender.py's module docstring). Returns
    the Twilio message SID; raises on failure, same as
    whatsapp_sender.send_cleaner_request."""
    property_label = PROPERTY_LABELS.get(property_key, property_key)
    return whatsapp_sender.send_cleaner_request(
        cleaner["whatsapp_number"], cleaner["name"], checkout_date.strftime("%d.%m.%Y"), property_label
    )


def run_daily_check(today: datetime.date = None, now: datetime.datetime = None,
                     dry_run: bool = True, log=print) -> list:
    """Evaluates every still-open Putzplan row and decides what should
    happen now. dry_run=True (default) only logs — never touches state,
    so it's safe to run against the real files as often as you like while
    testing. dry_run=False records attempts and persists state — that's
    what stage 1 will flip on right where it inserts the real Twilio send
    call (marked below)."""
    from . import putzplan_writer

    today = today or datetime.datetime.now(tz=BERLIN).date()
    now = now or datetime.datetime.now(tz=BERLIN)
    roster = load_roster()
    state = _load_state()
    actions = []

    for property_key, checkout_date in putzplan_writer.iter_open_rows():
        row_key = _state_key(property_key, checkout_date)
        action = decide_next_action(property_key, checkout_date, roster, state, today, now)
        actions.append((row_key, action))

        if action["action"] == "send_message":
            cleaner = roster[action["cleaner_id"]]
            if dry_run:
                log(f"    [cleaner-coordination] {row_key}: WhatsApp an {cleaner['name']} "
                    f"(Tier {action['tier']}) — WÜRDE gesendet (dry-run)")
            else:
                try:
                    _send_request(cleaner, property_key, checkout_date)
                except Exception as exc:  # noqa: BLE001 — one bad send must not break the whole run
                    log(f"    [cleaner-coordination] {row_key}: WARN — WhatsApp an {cleaner['name']} "
                        f"fehlgeschlagen: {exc}")
                else:
                    log(f"    [cleaner-coordination] {row_key}: WhatsApp an {cleaner['name']} "
                        f"(Tier {action['tier']}) gesendet")
                    record_attempt(state, row_key, action["cleaner_id"], action["tier"], now)

        elif action["action"] == "broadcast":
            names = ", ".join(roster[cid]["name"] for cid in action["cleaner_ids"])
            if dry_run:
                log(f"    [cleaner-coordination] {row_key}: BROADCAST an alle verbleibenden "
                    f"({names}) — WÜRDE gesendet (dry-run)")
            else:
                for cid in action["cleaner_ids"]:
                    cleaner = roster[cid]
                    try:
                        _send_request(cleaner, property_key, checkout_date)
                    except Exception as exc:  # noqa: BLE001
                        log(f"    [cleaner-coordination] {row_key}: WARN — Broadcast an {cleaner['name']} "
                            f"fehlgeschlagen: {exc}")
                    else:
                        record_attempt(state, row_key, cid, cleaner["tier"], now, stage="broadcast")
                log(f"    [cleaner-coordination] {row_key}: BROADCAST an alle verbleibenden ({names}) gesendet")

        elif action["action"] == "alert_owner":
            log(f"    [cleaner-coordination] {row_key}: ALARM an Farzaneh nötig — {action['reason']}")

    if not dry_run:
        _save_state(state)

    return actions


if __name__ == "__main__":
    run_daily_check(dry_run=True)
