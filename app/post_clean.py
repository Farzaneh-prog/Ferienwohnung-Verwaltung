"""
After the cleaning date (phase 3 stage 4, 2026-10-02): fills the real cleaner
code, hours and pay into the GästeListe row of the guest who just checked out
(columns V, T, U), replacing the placeholder defaults every new row gets
(xlsx_writer.EXPLICIT_LITERALS: V='M-Ü', T=3.5, U='=T*12').

Rules (agreed with Farzaneh):
- Who cleaned = Putzplan column A of the (non-storniert) row for that
  property/date, matched to the roster (match_cleaner). Manual entries count
  too. Free text that is no single person never occurs; if it does, she gets
  an alert and fills the row by hand.
- Fixed pay per cleaner/property from the roster (cleaner_roster.cleaning_pay);
  Manuela is paid per hour: hours default to 3.5, U stays =T*12 — correct the
  hours by hand if they differ.
- Only dates from POST_CLEAN_SINCE on ("from now on", default 2026-10-02) and
  only rows that still hold the untouched placeholders — anything edited by
  hand is left alone. Each property/date is handled once (state flag
  "excel_filled").
- Runs daily (scheduler.py, 09:00) for every date from the cutoff up to
  yesterday that hasn't been handled yet.
"""
import datetime
import os

from . import cleaner_coordination as cc
from . import putzplan_writer, xlsx_writer
from .cleaner_roster import cleaning_pay, load_roster, match_cleaner
from .config import PROPERTY_FILES
from .owner_alerts import notify_owner
from .tz import BERLIN

DEFAULT_SINCE = "2026-10-02"
MAX_BACKFILL_DAYS = 14


def _since() -> datetime.date:
    return datetime.date.fromisoformat(os.environ.get("POST_CLEAN_SINCE") or DEFAULT_SINCE)


def run_post_clean(now: datetime.datetime = None, dry_run: bool = False, log=print) -> list:
    """Returns [(property, date, outcome)] for what was looked at."""
    now = now or datetime.datetime.now(tz=BERLIN)
    today = now.date()
    first = max(_since(), today - datetime.timedelta(days=MAX_BACKFILL_DAYS))
    roster = load_roster()
    results = []
    d = first
    while d < today:
        for property_key in PROPERTY_FILES:
            with cc.STATE_LOCK:
                state = cc._load_state()
                key = cc._state_key(property_key, d)
                if state["rows"].get(key, {}).get("excel_filled"):
                    continue
            outcome = _handle(property_key, d, roster, dry_run, log)
            results.append((property_key, d, outcome))
            if outcome != "dry_run" and outcome != "retry":
                with cc.STATE_LOCK:
                    state = cc._load_state()
                    cc._row_state(state, key)["excel_filled"] = True
                    cc._save_state(state)
        d += datetime.timedelta(days=1)
    return results


def _handle(property_key: str, d: datetime.date, roster: dict, dry_run: bool, log) -> str:
    label = cc._label(property_key, d)
    code = putzplan_writer.get_assigned_code(property_key, d)
    if not code:
        # nothing was cleaned/planned for that day (or nobody entered) — nothing to fill
        return "no_cleaner_in_putzplan"
    cid = match_cleaner(roster, code)
    cleaner = roster.get(cid) if cid else None
    pay = cleaning_pay(cid, cleaner, property_key) if cleaner else None
    if pay is None:
        text = (f"⚠️ {label}: Putzplan nennt «{code}» — daraus lässt sich kein Stundenlohn ableiten; "
                f"GästeListe-Spalten T/U/V bitte von Hand ausfüllen.")
        log(f"    [post-clean] {text}")
        if not dry_run:
            notify_owner(text, log=log)
        return "unknown_cleaner"
    hours, amount, hourly_rate = pay
    if dry_run:
        log(f"    [post-clean] {label}: würde V={cleaner['code']}, T={hours}, "
            f"U={amount if amount is not None else f'=T*{hourly_rate:g}'} eintragen (dry-run)")
        return "dry_run"
    try:
        result = xlsx_writer.fill_cleaning_cells(property_key, d, cleaner["code"], hours, amount, hourly_rate, log=log)
    except Exception as exc:  # noqa: BLE001 — e.g. file locked by an open Excel; try again tomorrow
        log(f"    [post-clean] WARN — {label}: {exc!r}")
        return "retry"
    if result == "filled":
        notify_owner(f"GästeListe {label}: {cleaner['name']}, {hours} h, "
                     f"{amount if amount is not None else f'{hours * hourly_rate:.2f}'} € eingetragen", log=log)
    return result
