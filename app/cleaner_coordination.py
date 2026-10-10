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

Answers (2026-10-01, decided with Farzaneh after the first end-to-end test):
- "Nein" and "Vielleicht" both move to the next person in the chain
  IMMEDIATELY (the webhook calls advance_row right after recording the
  answer) — not after the 1h cooldown. No answer for 1h is treated like
  "Vielleicht" (the cooldown only applies while the LAST attempt is still
  unanswered).
- A later "Ja" from someone who said "Vielleicht"/didn't answer still
  counts (first "Ja" wins).
- Once everyone in the chain was asked and nobody confirmed: ONE reminder
  ("noch niemand gefunden") to everyone who didn't say "Nein"; after that
  it waits for the <=3-day broadcast window.
- Every message sent / answer received is mirrored to the owner's WhatsApp
  (owner_alerts.notify_owner), and so are the alert_owner cases (once per
  reason per row).

State lives in data/cleaner_coordination_state.json (gitignored):
{
  "rows": {"<property>|<checkout ISO date>": {
      "attempts": [{"cleaner_id", "tier", "stage", "sent_at",
                    "response": None|"ja"|"nein"|"vielleicht", "responded_at"}],
      "confirmed_cleaner_id": None,
      "reminder_sent": False, "alerts_sent": [],
  }},
  "level1_rotation": {"last_first": "jennifer"},
  "leave_weeks": {"jennifer": [["2026-10-01", "2026-10-11"]]},   # HARD: unreachable, never asked
  "soft_weeks": {"mehrnaz": [["2026-10-12", "2026-10-18"]]},      # SOFT: free week, asked last
}
"""
import datetime
import json
import os
import threading

from .config import DATA_DIR, PROPERTY_LABELS
from .cleaner_roster import load_roster, match_cleaner, sheet_label
from .tz import BERLIN
from . import whatsapp_sender
from .owner_alerts import notify_owner

# Webhook threads and the scheduled run both load-modify-save the same
# state file — serialize them (re-entrant: the webhook handler calls
# functions below that take the lock themselves).
STATE_LOCK = threading.RLock()

_STATE_PATH = os.path.join(DATA_DIR, "data", "cleaner_coordination_state.json")

LEVEL1_START_DAYS = 10   # first automated attempt
BROADCAST_START_DAYS = 3  # "Sicherheitsnetz" — message everyone eligible at once
# 1h by default; COORDINATION_COOLDOWN_MINUTES shortens it for live tests only.
RETRY_COOLDOWN = datetime.timedelta(minutes=int(os.environ.get("COORDINATION_COOLDOWN_MINUTES") or 60))


def test_mode() -> bool:
    """SCHEDULER_ONLY_CLEANERS set = live TEST MODE: only these (test) cleaners
    exist in the roster and only state rows flagged "test" are processed — real
    Putzplan rows and real cleaners are left completely alone."""
    return bool(os.environ.get("SCHEDULER_ONLY_CLEANERS", "").strip())

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


_BUSY_CACHE = {}


def _busy_cleaners(date: datetime.date, exclude_property: str, roster: dict) -> set:
    """Cleaner ids who already have a (non-storniert) Putzplan assignment on
    `date` at ANOTHER property (column A). Cached per Putzplan file version —
    eligible_chain runs for every open row on every check."""
    from . import putzplan_writer

    try:
        stamp = os.path.getmtime(putzplan_writer._putzplan_path())
    except OSError:
        return set()
    key = (date, stamp)
    if key not in _BUSY_CACHE:
        if len(_BUSY_CACHE) > 200:
            _BUSY_CACHE.clear()
        _BUSY_CACHE[key] = list(putzplan_writer.iter_assigned_rows(date))
    busy = set()
    for property_key, code, _guests in _BUSY_CACHE[key]:
        if property_key != exclude_property:
            cid = match_cleaner(roster, code)
            if cid:
                busy.add(cid)
    return busy


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

    chain = tier1 + rest
    # Soft "free week" (monthly, planned): that person drops to the END of
    # the chain — still used if nobody else is available. (Hard leave —
    # travelling, unreachable — already excluded them in _is_eligible.)
    def soft(cid):
        return any(datetime.date.fromisoformat(s) <= date <= datetime.date.fromisoformat(e)
                   for s, e in state.get("soft_weeks", {}).get(cid, []))

    # Already booked that day at the OTHER property (agreed with Farzaneh
    # 2026-10-10: one cleaner, one property per day): also moves to the END —
    # still asked if nobody else is left (and in the <=3-day broadcast).
    busy = _busy_cleaners(date, property_key, roster)
    # stable sort: normal (0) -> soft free week (1) -> busy elsewhere (2) -> both (3)
    return sorted(chain, key=lambda cid: (2 if cid in busy else 0) + (1 if soft(cid) else 0))


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

    if row_state.get("cancelled"):
        return {"action": "none", "reason": "cancelled"}
    if row_state.get("confirmed_cleaner_id"):
        return {"action": "none", "reason": "confirmed"}

    days_until = (checkout_date - today).days
    if days_until > LEVEL1_START_DAYS and not row_state.get("manual_start"):
        return {"action": "none", "reason": "too_early"}

    chain = eligible_chain(property_key, checkout_date, roster, state)
    attempts = row_state.get("attempts", [])
    tried_ids = [a["cleaner_id"] for a in attempts]

    # Cleaners already booked that day at the OTHER property are asked only
    # once everybody else has explicitly said Nein (decided with Farzaneh
    # 2026-10-10) — silence or "Vielleicht" does not open the door for them,
    # not in the sequential chain and not in the broadcast either.
    busy = _busy_cleaners(checkout_date, property_key, roster)
    if busy:
        latest_all = {a["cleaner_id"]: a.get("response") for a in attempts}
        out_all = set(row_state.get("withdrawn", []))
        normal = [cid for cid in chain if cid not in busy]
        if not all(latest_all.get(cid) == "nein" or cid in out_all for cid in normal):
            chain = normal

    if days_until < 0:
        return {"action": "alert_owner", "reason": "checkout_passed_unresolved"}

    if row_state.get("reopened"):
        # The confirmed cleaner backed out and no reserve was left: ask
        # everyone eligible again, except people who said Nein or just
        # backed out themselves.
        out = set(row_state.get("withdrawn", []))
        targets = [cid for cid in chain if cid not in out and not any(
            a["cleaner_id"] == cid and a.get("response") == "nein" for a in attempts)]
        if not targets:
            return {"action": "alert_owner", "reason": "everyone_declined"}
        if _in_quiet_hours(now):
            return {"action": "none", "reason": "quiet_hours"}
        return {"action": "broadcast" if days_until <= BROADCAST_START_DAYS else "remind_open",
                "cleaner_ids": targets, "reopen": True}

    if attempts and attempts[-1].get("stage") == "again" and _awaiting_answer(attempts, now):
        # replacement booking after a cancel notice: the previous cleaner was
        # asked first — give them the hour before anyone else (even a broadcast)
        return {"action": "none", "reason": "awaiting_again"}

    if days_until <= BROADCAST_START_DAYS:
        if not chain:
            return {"action": "alert_owner", "reason": "no_eligible_cleaner"}
        asked_in_broadcast = {a["cleaner_id"] for a in attempts if a.get("stage") == "broadcast"}
        declined = {a["cleaner_id"] for a in attempts if a.get("response") == "nein"}
        # not yet asked in a broadcast and not declined — a second wave happens
        # when the busy-elsewhere cleaners become eligible (all others said Nein)
        targets = [cid for cid in chain if cid not in declined and cid not in asked_in_broadcast]
        if not targets:
            if asked_in_broadcast and not all(cid in declined for cid in chain):
                return {"action": "none", "reason": "broadcast_already_sent"}
            return {"action": "alert_owner", "reason": "everyone_declined"}
        if _in_quiet_hours(now):
            return {"action": "none", "reason": "quiet_hours"}
        return {"action": "broadcast", "cleaner_ids": targets}

    untried = [cid for cid in chain if cid not in tried_ids]
    if not untried:
        if not chain:
            return {"action": "alert_owner", "reason": "no_eligible_cleaner"}
        latest = {a["cleaner_id"]: a.get("response") for a in attempts}  # last attempt per cleaner wins
        out = set(row_state.get("withdrawn", []))
        if all(latest.get(cid) == "nein" or cid in out for cid in chain):
            # everyone has declined / backed out — don't wait for the broadcast window
            return {"action": "alert_owner", "reason": "everyone_declined"}
        if not row_state.get("reminder_sent") and not _awaiting_answer(attempts, now):
            targets = [cid for cid in chain if not any(
                a["cleaner_id"] == cid and a.get("response") == "nein" for a in attempts)]
            if targets:
                if _in_quiet_hours(now):
                    return {"action": "none", "reason": "quiet_hours"}
                return {"action": "remind_open", "cleaner_ids": targets}
        return {"action": "none", "reason": "chain_exhausted_awaiting_broadcast_window"}

    if _awaiting_answer(attempts, now):
        return {"action": "none", "reason": "cooldown"}

    if _in_quiet_hours(now):
        return {"action": "none", "reason": "quiet_hours"}

    next_id = untried[0]
    return {"action": "send_message", "cleaner_id": next_id, "tier": roster[next_id]["tier"]}


# Set by scheduler.py: called after state changed so it can (re)schedule the
# one-shot wake-ups (1h answer timeout, 08:00 after quiet hours).
FOLLOWUP_HOOK = None


def _followup() -> None:
    if FOLLOWUP_HOOK is not None:
        try:
            FOLLOWUP_HOOK()
        except Exception as exc:  # noqa: BLE001
            print(f"    [cleaner-coordination] WARN — followup scheduling failed: {exc!r}", flush=True)


def pre_assign(property_key: str, checkout_date: datetime.date, cleaner_id: str) -> None:
    """Registers "this cleaner does this date" BEFORE the booking is in the
    system: when the Putzplan row appears (day_before.on_new_reservation) column
    A is filled automatically and the row never enters the search chain."""
    with STATE_LOCK:
        state = _load_state()
        state.setdefault("pre_assigned", {})[_state_key(property_key, checkout_date)] = cleaner_id
        _save_state(state)


def assign_by_owner(property_key: str, checkout_date: datetime.date, cleaner_id: str, log=print) -> str:
    """Dashboard "who does it" (owner decides before/instead of the search).
    - Putzplan row exists, column A empty -> written now, row confirmed, the
      search for it stops.
    - Row exists and someone is already in column A -> nothing changes.
    - No row yet (booking not imported) -> remembered (pre_assign) and applied
      the moment the row appears.
    Returns "assigned" | "already:<name>" | "pre_assigned"."""
    from . import putzplan_writer

    roster = load_roster()
    info = putzplan_writer.find_active_row(property_key, checkout_date)
    if info is None:
        pre_assign(property_key, checkout_date, cleaner_id)
        return "pre_assigned"
    if info["cleaner"]:
        return f"already:{info['cleaner']}"
    putzplan_writer.assign_cleaner(property_key, checkout_date, sheet_label(roster[cleaner_id]), log=log)
    with STATE_LOCK:
        state = _load_state()
        row_state = _row_state(state, _state_key(property_key, checkout_date))
        row_state["confirmed_cleaner_id"] = cleaner_id
        row_state["reopened"] = False
        # A search may already be running (booking imported a moment ago):
        # tell everyone who was asked and hasn't declined that it is taken.
        asked = {a["cleaner_id"] for a in row_state.get("attempts", [])
                 if a["cleaner_id"] != cleaner_id and a.get("response") not in ("nein",)}
        _save_state(state)
    for other_id in sorted(asked):
        other = roster.get(other_id)
        if other and other.get("whatsapp_number"):
            try:
                whatsapp_sender.send_template(
                    other["whatsapp_number"], "taken",
                    {"1": other["name"], "2": checkout_date.strftime("%d.%m.%Y"),
                     "3": PROPERTY_LABELS.get(property_key, property_key)})
            except Exception as exc:  # noqa: BLE001
                log(f"    [cleaner-coordination] WARN — 'vergeben' an {other['name']} fehlgeschlagen: {exc}")
    notify_owner(f"Von Hand zugewiesen: {roster[cleaner_id]['name']} → {_label(property_key, checkout_date)}", log=log)
    return "assigned"


def remove_pre_assignment(property_key: str, checkout_date: datetime.date) -> None:
    with STATE_LOCK:
        state = _load_state()
        state.get("pre_assigned", {}).pop(_state_key(property_key, checkout_date), None)
        _save_state(state)


def start_manual_search(property_key: str, checkout_date: datetime.date, now: datetime.datetime = None,
                        dry_run: bool = False, log=print) -> dict:
    """Dashboard's "find a cleaner for this date" (Putz-Alerts page): starts
    the chain right away even when the date is more than 10 days out."""
    now = now or datetime.datetime.now(tz=BERLIN)
    row_key = _state_key(property_key, checkout_date)
    if dry_run:
        log(f"    [cleaner-coordination] {row_key}: manuelle Suche — dry-run, nichts gestartet")
        return {"action": "none", "reason": "dry_run"}
    with STATE_LOCK:
        state = _load_state()
        rs = _row_state(state, row_key)
        rs["manual_start"] = True
        if test_mode():
            rs["test"] = True  # test mode only processes flagged rows
        _save_state(state)
    notify_owner(f"Suche gestartet → {_label(property_key, checkout_date)}", log=log)
    return advance_row(property_key, checkout_date, now=now, log=log)


def row_status(property_key: str, checkout_date: datetime.date, roster: dict = None, state: dict = None) -> str:
    """One-line German status for the dashboard."""
    roster = roster if roster is not None else load_roster()
    state = state if state is not None else _load_state()
    rs = state["rows"].get(_state_key(property_key, checkout_date))
    name = lambda cid: roster.get(cid, {}).get("name", cid)  # noqa: E731
    if rs is None:
        return "noch nicht gestartet"
    if rs.get("cancelled"):
        return "storniert" + (" (Info gesendet)" if rs.get("cancel_notified") else "")
    if rs.get("confirmed_cleaner_id"):
        extra = f", Reserve: {', '.join(name(c) for c in rs['reserves'])}" if rs.get("reserves") else ""
        return f"bestätigt: {name(rs['confirmed_cleaner_id'])}{extra}"
    attempts = rs.get("attempts", [])
    if not attempts:
        return "noch nicht gestartet"
    waiting = [name(a["cleaner_id"]) for a in attempts if a.get("response") is None]
    if waiting:
        return f"wartet auf Antwort: {waiting[-1]}"
    return "Antworten: " + ", ".join(f"{name(a['cleaner_id'])}: {a['response']}" for a in attempts)


def _awaiting_answer(attempts: list, now: datetime.datetime) -> bool:
    """True while the most recent request is unanswered AND younger than
    RETRY_COOLDOWN. An answered attempt (even Nein/Vielleicht) never
    blocks the next send; an unanswered one older than the cooldown counts
    as "Vielleicht"."""
    if not attempts or attempts[-1].get("response") is not None:
        return False
    last_sent_at = datetime.datetime.fromisoformat(attempts[-1]["sent_at"])
    return now - last_sent_at < RETRY_COOLDOWN


def record_attempt(state: dict, row_key: str, cleaner_id: str, tier: int, sent_at: datetime.datetime,
                    stage: str = None) -> None:
    row_state = state["rows"].setdefault(row_key, {"attempts": [], "confirmed_cleaner_id": None})
    is_first_attempt = len(row_state["attempts"]) == 0
    row_state["attempts"].append({
        "cleaner_id": cleaner_id,
        "tier": tier,
        "stage": stage or ("level1" if tier == 1 else "chain"),
        "sent_at": sent_at.isoformat(),
        "response": None,
        "responded_at": None,
    })
    if is_first_attempt and tier == 1:
        state.setdefault("level1_rotation", {})["last_first"] = cleaner_id


def _row_state(state: dict, row_key: str) -> dict:
    return state["rows"].setdefault(row_key, {"attempts": [], "confirmed_cleaner_id": None})


def find_row_for_reply(cleaner_id: str, state: dict, today: datetime.date):
    """Which row is a reply from this cleaner about? The reply itself
    carries no context, only the sender. Open requests first (their latest
    attempt for the row is still unanswered; earliest checkout wins), else
    the row they were most recently involved in — people may change their
    answer at any time ("Vielleicht" -> "Ja", even "Ja" -> "Nein"). Rows
    whose checkout date has passed are ignored. Returns
    (property_key, checkout_date) or None."""
    unanswered, recent = [], []
    for row_key, row_state in state.get("rows", {}).items():
        property_key, date_str = row_key.split("|", 1)
        checkout_date = datetime.date.fromisoformat(date_str)
        if checkout_date < today:
            continue
        mine = [a for a in row_state.get("attempts", []) if a["cleaner_id"] == cleaner_id]
        if not mine:
            continue
        if mine[-1].get("response") is None:
            unanswered.append((checkout_date, property_key))
        recent.append((max(a.get("responded_at") or a["sent_at"] for a in mine), checkout_date, property_key))
    if unanswered:
        checkout_date, property_key = min(unanswered)
        return property_key, checkout_date
    if recent:
        _, checkout_date, property_key = max(recent)
        return property_key, checkout_date
    return None


def record_response(state: dict, row_key: str, cleaner_id: str, response: str, now: datetime.datetime):
    """Stores the cleaner's answer on their LATEST attempt for the row,
    overwriting any earlier answer (last answer wins), and appends to the
    row's "history". Returns (found, previous_response)."""
    row_state = state["rows"].get(row_key, {})
    for attempt in reversed(row_state.get("attempts", [])):
        if attempt["cleaner_id"] != cleaner_id:
            continue
        previous = attempt.get("response")
        attempt["response"] = response
        attempt["responded_at"] = now.isoformat()
        if previous != response:
            row_state.setdefault("history", []).append(
                {"cleaner_id": cleaner_id, "response": response, "at": now.isoformat()})
        return True, previous
    return False, None


def _label(property_key: str, checkout_date: datetime.date) -> str:
    return f"{PROPERTY_LABELS.get(property_key, property_key)} {checkout_date.strftime('%d.%m.%Y')}"


def confirm_cleaner(property_key: str, checkout_date: datetime.date, cleaner_id: str, log=print,
                    now: datetime.datetime = None, notify_others: bool = True, promoted: bool = False) -> bool:
    """Marks the row confirmed for this cleaner in state AND writes their
    code into Putzplan column A (app/putzplan_writer.assign_cleaner).
    Returns False (no-op) if the row was already confirmed by someone else
    in the meantime (race between two tiers' replies — first one wins,
    matches the architecture doc's "به محض دریافت پاسخ «بله» از هرکس،
    چرخش/ارسال متوقف می‌شه").

    Side effects on success: unless notify_others=False, everyone else who
    was asked about this row and hasn't answered Nein or Ja gets the
    "schon vergeben" template (people who said Ja are Reserves — they were
    already told so); the owner gets a short confirmation line. With
    promoted=True (a Reserve moving up after the confirmed person backed
    out) the cleaner gets the "promoted" template instead, since there's
    no incoming message to reply to."""
    from . import putzplan_writer

    now = now or datetime.datetime.now(tz=BERLIN)
    row_key = _state_key(property_key, checkout_date)
    with STATE_LOCK:
        state = _load_state()
        row_state = _row_state(state, row_key)
        if row_state.get("confirmed_cleaner_id"):
            log(f"    [cleaner-coordination] {row_key}: 'Ja' von {cleaner_id} ignoriert — schon bestätigt "
                f"durch {row_state['confirmed_cleaner_id']}")
            return False

        record_response(state, row_key, cleaner_id, "ja", now)
        row_state["confirmed_cleaner_id"] = cleaner_id
        row_state["reserves"] = [c for c in row_state.get("reserves", []) if c != cleaner_id]
        row_state["reopened"] = False
        _save_state(state)

    roster = load_roster()
    cleaner = roster.get(cleaner_id, {})
    cleaner_code = sheet_label(cleaner) if cleaner else cleaner_id  # Putzplan column A uses first names
    putzplan_writer.assign_cleaner(property_key, checkout_date, cleaner_code, log=log)
    log(f"    [cleaner-coordination] {row_key}: bestätigt durch {cleaner_id} ({cleaner_code})")

    date_str = checkout_date.strftime("%d.%m.%Y")
    property_label = PROPERTY_LABELS.get(property_key, property_key)
    if promoted and cleaner.get("whatsapp_number"):
        try:
            whatsapp_sender.send_template(cleaner["whatsapp_number"], "promoted",
                                          {"1": cleaner["name"], "2": date_str, "3": property_label})
        except Exception as exc:  # noqa: BLE001
            log(f"    [cleaner-coordination] {row_key}: WARN — 'promoted' an {cleaner['name']} fehlgeschlagen: {exc}")
    if notify_others:
        others = {a["cleaner_id"] for a in row_state.get("attempts", [])
                  if a["cleaner_id"] != cleaner_id and a.get("response") not in ("nein", "ja")}
        for other_id in sorted(others):
            other = roster.get(other_id)
            if not other or not other.get("whatsapp_number"):
                continue
            try:
                whatsapp_sender.send_template(other["whatsapp_number"], "taken",
                                              {"1": other["name"], "2": date_str, "3": property_label})
            except Exception as exc:  # noqa: BLE001
                log(f"    [cleaner-coordination] {row_key}: WARN — 'vergeben' an {other['name']} "
                    f"fehlgeschlagen: {exc}")
    notify_owner(f"{cleaner.get('name', cleaner_id)}: "
                 f"{'rückt als Reserve nach ✓' if promoted else 'Ja ✓'} → {_label(property_key, checkout_date)} bestätigt",
                 log=log)
    return True


def apply_response(cleaner_id: str, response: str, now: datetime.datetime = None, roster: dict = None,
                   log=print) -> dict:
    """Central handler for an incoming Ja / Nein / Vielleicht
    (app/whatsapp_webhook.py just maps the result to a reply text).
    Answers can be changed at any time; the last one counts:

    - Ja on an open row -> confirmed (Putzplan column A written).
    - Ja on a row someone ELSE already confirmed -> recorded as a RESERVE
      (in order of arrival).
    - Nein/Vielleicht on an open row -> recorded, next person is asked right away.
    - Nein/Vielleicht from the CONFIRMED cleaner -> backs out: Putzplan
      column A is cleared; the first Reserve who still says Ja moves up
      ("promoted" template); with no Reserve left the row is reopened and
      everyone eligible (except those who said Nein) is asked again, and the
      owner gets an alert.

    Returns {"kind": ..., "property_key", "checkout_date", "position"?}.
    kind: no_row | confirmed | still_confirmed | reserve | declined | maybe
          | withdrawn_promoted | withdrawn_reopened"""
    from . import putzplan_writer

    now = now or datetime.datetime.now(tz=BERLIN)
    with STATE_LOCK:
        state = _load_state()
        roster = roster if roster is not None else load_roster()
        found = find_row_for_reply(cleaner_id, state, now.date())
        if found is None:
            return {"kind": "no_row"}
        property_key, checkout_date = found
        row_key = _state_key(property_key, checkout_date)
        row_state = _row_state(state, row_key)
        name = roster.get(cleaner_id, {}).get("name", cleaner_id)
        label = _label(property_key, checkout_date)
        info = {"property_key": property_key, "checkout_date": checkout_date}
        confirmed = row_state.get("confirmed_cleaner_id")
        if row_state.get("cancelled"):
            return dict(info, kind="cancelled")
        record_response(state, row_key, cleaner_id, response, now)

        if response == "ja":
            if confirmed == cleaner_id:
                _save_state(state)
                return dict(info, kind="still_confirmed")
            if confirmed:
                reserves = row_state.setdefault("reserves", [])
                if cleaner_id not in reserves:
                    reserves.append(cleaner_id)
                _save_state(state)
                notify_owner(f"{name}: Ja → {label}, aber schon vergeben — Reserve Nr. {reserves.index(cleaner_id) + 1}",
                             log=log)
                return dict(info, kind="reserve", position=reserves.index(cleaner_id) + 1)
            _save_state(state)
            confirm_cleaner(property_key, checkout_date, cleaner_id, log=log, now=now)
            return dict(info, kind="confirmed")

        # Nein / Vielleicht
        row_state["reserves"] = [c for c in row_state.get("reserves", []) if c != cleaner_id]
        word = "Nein" if response == "nein" else "Vielleicht"

        if confirmed == cleaner_id:
            row_state["confirmed_cleaner_id"] = None
            row_state.setdefault("withdrawn", []).append(cleaner_id)
            _save_state(state)
            putzplan_writer.assign_cleaner(property_key, checkout_date, None, log=log)
            replacement = next((c for c in row_state.get("reserves", [])
                                if c in roster and any(a["cleaner_id"] == c and a.get("response") == "ja"
                                                       for a in row_state["attempts"])), None)
            if replacement:
                notify_owner(f"{name}: {word} → {label}, Bestätigung zurückgenommen; "
                             f"Reserve {roster[replacement]['name']} rückt nach", log=log)
                confirm_cleaner(property_key, checkout_date, replacement, log=log, now=now,
                                notify_others=False, promoted=True)
                return dict(info, kind="withdrawn_promoted")
            state = _load_state()
            row_state = state["rows"][row_key]
            row_state["reopened"] = True
            row_state["reminder_sent"] = False
            _save_state(state)
            notify_owner(f"⚠️ {name}: {word} → {label}, Bestätigung zurückgenommen, keine Reserve — "
                         f"frage alle neu an", log=log)
            advance_row(property_key, checkout_date, now=now, roster=roster, log=log)
            return dict(info, kind="withdrawn_reopened")

        _save_state(state)
        notify_owner(f"{name}: {word} → {label}", log=log)
        advance_row(property_key, checkout_date, now=now, roster=roster, log=log)
        return dict(info, kind="declined" if response == "nein" else "maybe")


def _send_request(cleaner: dict, property_key: str, checkout_date: datetime.date, template: str = None) -> str:
    """Sends a cleaner-request template (required for a cold first contact —
    see whatsapp_sender.py's module docstring). template=None is the
    first-contact request; "reminder_open"/"urgent" are the follow-ups.
    Returns the Twilio message SID; raises on failure."""
    property_label = PROPERTY_LABELS.get(property_key, property_key)
    date_str = checkout_date.strftime("%d.%m.%Y")
    if template is None:
        return whatsapp_sender.send_cleaner_request(cleaner["whatsapp_number"], cleaner["name"], date_str,
                                                    property_label)
    return whatsapp_sender.send_template(
        cleaner["whatsapp_number"], template, {"1": cleaner["name"], "2": date_str, "3": property_label})


_ALERT_TEXT = {
    "no_eligible_cleaner": "Niemand ist für {label} verfügbar — bitte manuell zuweisen.",
    "everyone_declined": "Alle haben für {label} abgesagt — bitte manuell zuweisen.",
    "checkout_passed_unresolved": "{label}: Datum erreicht, keine Reinigung zugewiesen!",
}


def _execute_action(property_key: str, checkout_date: datetime.date, action: dict, roster: dict, state: dict,
                    now: datetime.datetime, dry_run: bool, log) -> None:
    """Performs one decide_next_action result against `state` (mutated in
    place — caller saves). Shared by run_daily_check and advance_row."""
    row_key = _state_key(property_key, checkout_date)
    label = _label(property_key, checkout_date)
    kind = action["action"]

    # remember rows whose next step is only waiting for 08:00 (quiet hours),
    # so the scheduler can wake up exactly then instead of polling
    held = kind == "none" and action.get("reason") == "quiet_hours"
    if not dry_run and (held or row_key in state["rows"]):
        state["rows"].setdefault(row_key, {"attempts": [], "confirmed_cleaner_id": None})["held_quiet"] = held

    if kind == "send_message":
        cleaner = roster[action["cleaner_id"]]
        if dry_run:
            log(f"    [cleaner-coordination] {row_key}: WhatsApp an {cleaner['name']} "
                f"(Tier {action['tier']}) — WÜRDE gesendet (dry-run)")
            return
        try:
            _send_request(cleaner, property_key, checkout_date)
        except Exception as exc:  # noqa: BLE001 — one bad send must not break the whole run
            log(f"    [cleaner-coordination] {row_key}: WARN — WhatsApp an {cleaner['name']} "
                f"fehlgeschlagen: {exc}")
        else:
            log(f"    [cleaner-coordination] {row_key}: WhatsApp an {cleaner['name']} "
                f"(Tier {action['tier']}) gesendet")
            record_attempt(state, row_key, action["cleaner_id"], action["tier"], now)
            notify_owner(f"Anfrage an {cleaner['name']} → {label}", log=log)

    elif kind in ("broadcast", "remind_open"):
        template, stage = ("urgent", "broadcast") if kind == "broadcast" else ("reminder_open", None)
        names = ", ".join(roster[cid]["name"] for cid in action["cleaner_ids"])
        what = "BROADCAST an alle verbleibenden" if kind == "broadcast" else "ERINNERUNG 'noch niemand gefunden' an"
        if dry_run:
            log(f"    [cleaner-coordination] {row_key}: {what} ({names}) — WÜRDE gesendet (dry-run)")
            return
        sent = []
        for cid in action["cleaner_ids"]:
            cleaner = roster[cid]
            try:
                _send_request(cleaner, property_key, checkout_date, template=template)
            except Exception as exc:  # noqa: BLE001
                log(f"    [cleaner-coordination] {row_key}: WARN — {template} an {cleaner['name']} "
                    f"fehlgeschlagen: {exc}")
            else:
                sent.append(cleaner["name"])
                if kind == "broadcast" or action.get("reopen"):
                    # a fresh unanswered attempt, so the reply maps to this new request
                    record_attempt(state, row_key, cid, cleaner["tier"], now,
                                   stage=stage or "reopen")
        if sent:
            row_state = state["rows"][row_key]
            if kind == "remind_open":
                row_state["reminder_sent"] = True
            if action.get("reopen"):
                row_state["reopened"] = False
        if sent:
            log(f"    [cleaner-coordination] {row_key}: {what} ({', '.join(sent)}) gesendet")
            notify_owner(("Broadcast" if kind == "broadcast" else "Erinnerung") + f" an {', '.join(sent)} → {label}",
                         log=log)

    elif kind == "alert_owner":
        row_state = state["rows"].setdefault(row_key, {"attempts": [], "confirmed_cleaner_id": None})
        already = row_state.setdefault("alerts_sent", [])
        log(f"    [cleaner-coordination] {row_key}: ALARM an Farzaneh nötig — {action['reason']}")
        if dry_run or action["reason"] in already:
            return
        text = _ALERT_TEXT.get(action["reason"], "{label}: " + action["reason"]).format(label=label)
        if notify_owner("⚠️ " + text, log=log):
            already.append(action["reason"])


def advance_row(property_key: str, checkout_date: datetime.date, now: datetime.datetime = None,
                today: datetime.date = None, roster: dict = None, log=print) -> dict:
    """Re-evaluates ONE row right now and acts on it — called by the
    webhook right after a Nein/Vielleicht so the next person is asked
    immediately instead of waiting for the next scheduled run."""
    now = now or datetime.datetime.now(tz=BERLIN)
    today = today or now.date()
    with STATE_LOCK:
        roster = roster if roster is not None else load_roster()
        state = _load_state()
        action = decide_next_action(property_key, checkout_date, roster, state, today, now)
        _execute_action(property_key, checkout_date, action, roster, state, now, False, log)
        _save_state(state)
    _followup()
    return action


def run_daily_check(today: datetime.date = None, now: datetime.datetime = None,
                     dry_run: bool = True, log=print) -> list:
    """Evaluates every still-open Putzplan row and decides what should
    happen now. dry_run=True (default) only logs — never touches state,
    so it's safe to run against the real files as often as you like while
    testing. dry_run=False sends the WhatsApp messages, records attempts
    and persists state."""
    from . import putzplan_writer

    today = today or datetime.datetime.now(tz=BERLIN).date()
    now = now or datetime.datetime.now(tz=BERLIN)
    actions = []

    with STATE_LOCK:
        roster = load_roster()
        state = _load_state()
        if test_mode():
            rows = [(k.split("|", 1)[0], datetime.date.fromisoformat(k.split("|", 1)[1]))
                    for k, rs in state["rows"].items() if rs.get("test")]
        else:
            rows = putzplan_writer.iter_open_rows()
        for property_key, checkout_date in rows:
            action = decide_next_action(property_key, checkout_date, roster, state, today, now)
            actions.append((_state_key(property_key, checkout_date), action))
            _execute_action(property_key, checkout_date, action, roster, state, now, dry_run, log)
        if not dry_run:
            _save_state(state)

    if not dry_run:
        _followup()
    return actions


if __name__ == "__main__":
    run_daily_check(dry_run=True)
