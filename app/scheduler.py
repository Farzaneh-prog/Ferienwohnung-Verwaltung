"""
Background scheduler (phase 3 stage 2/3, 2026-10-01) — event driven, no
blind polling:

- EVENTS: after an import, a manual "find a cleaner" or a cancellation the
  routes call trigger_check() — a coordination check runs right away.
  Cleaner replies are handled by the webhook itself.
- ONE-SHOT TIMERS (schedule_followups, rebuilt from the state file after
  every change and at startup, so a restart loses nothing):
    * a request nobody answered -> check again 1h after it was sent
      (silence counts like "Vielleicht", the next person is asked)
    * a decision held back by quiet hours (22:00-08:00) -> check at 08:00
- Mondays 16:00 post-clean: real cleaner/hours/pay into GästeListe T/U/V (post_clean.py)
- 1st of month 08:00: Twilio cost report e-mail; daily 16:00: low-balance warning (billing.py)
- 18:00 daily digest e-mail to the owner (owner_alerts.send_daily_digest).
- DAILY JOBS (Europe/Berlin): 16:00 check + reminders for tomorrow's
  cleanings + contract warnings; 19:00 cancellations nobody replaced + late
  reminders. The daily 16:00 check is also what starts rows crossing the
  10-day / 3-day marks (no event announces that) and the safety net for
  manual Excel edits.

Safe by default — nothing runs unless you opt in via .env:
  SCHEDULER_ENABLED=1   start the scheduler at all (default: off, so
                        `python run.py` on the dev machine never messages anyone)
  SCHEDULER_DRY_RUN=1|0 1 (default): only LOG what would be sent; 0: really send
"""
import datetime
import os
import threading

import pytz

_scheduler = None
_lock = threading.Lock()
BERLIN_PYTZ = pytz.timezone("Europe/Berlin")


def _env_flag(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None or value.strip() == "":
        return default
    return value.strip().lower() in ("1", "true", "yes", "on")


def is_enabled() -> bool:
    return _env_flag("SCHEDULER_ENABLED", False)


def is_dry_run() -> bool:
    return _env_flag("SCHEDULER_DRY_RUN", True)


def _log(message: str) -> None:
    print(message, flush=True)


def run_check_job(reason: str = "") -> list:
    """One coordination check over all open rows. Never raises (an
    exception inside a scheduled job would otherwise just vanish)."""
    from .cleaner_coordination import run_daily_check
    from . import clock_guard

    dry_run = is_dry_run()
    if not dry_run and not clock_guard.ok_to_send():
        _log(f"[scheduler] check{f' ({reason})' if reason else ''} SKIPPED — system clock is off, see clock-guard")
        clock_guard.check(notify=True, log=_log)
        return []
    try:
        actions = run_daily_check(dry_run=dry_run, log=_log)
    except Exception as exc:  # noqa: BLE001
        _log(f"[scheduler] ERROR in cleaner coordination check: {exc!r}")
        return []
    interesting = [(key, a["action"]) for key, a in actions if a["action"] != "none"]
    _log(f"[scheduler] {datetime.datetime.now():%Y-%m-%d %H:%M} check{f' ({reason})' if reason else ''} done "
         f"({'dry-run' if dry_run else 'LIVE'}): {len(actions)} open rows, actions: {interesting or 'none'}")
    return actions


def run_evening_job(hour: int) -> None:
    """16:00 / 19:00: check, then reminders/cancellations for tomorrow."""
    from . import clock_guard, day_before

    clock_guard.check(notify=not is_dry_run(), log=_log)
    if hour < 19 and not is_dry_run():
        from . import billing

        try:
            billing.check_balance(log=_log)
        except Exception as exc:  # noqa: BLE001
            _log(f"[scheduler] ERROR in balance check: {exc!r}")
    if not is_dry_run() and not clock_guard.ok_to_send():
        _log(f"[scheduler] evening job {hour}:00 SKIPPED — system clock is off")
        return
    if hour < 19:
        run_check_job(f"daily {hour}:00")
    try:
        day_before.run_evening_job(hour, log=_log)
    except Exception as exc:  # noqa: BLE001
        _log(f"[scheduler] ERROR in evening job {hour}:00: {exc!r}")
    _log(f"[scheduler] evening job {hour}:00 done ({'dry-run' if is_dry_run() else 'LIVE'})")


def run_digest_job() -> None:
    """18:00: daily e-mail with everything the automation did/noticed."""
    from . import owner_alerts

    try:
        sent = owner_alerts.send_daily_digest(log=_log)
        _log(f"[scheduler] daily digest {'sent' if sent else 'not sent (nothing to report or e-mail failed)'}")
    except Exception as exc:  # noqa: BLE001
        _log(f"[scheduler] ERROR in daily digest: {exc!r}")


def run_post_clean_job() -> None:
    """Mondays 16:00: GästeListe columns V/T/U for the cleanings that already happened."""
    from . import clock_guard, post_clean

    if not is_dry_run() and not clock_guard.ok_to_send():
        _log("[scheduler] post-clean SKIPPED — system clock is off")
        return
    try:
        results = post_clean.run_post_clean(dry_run=is_dry_run(), log=_log)
        _log(f"[scheduler] post-clean done ({'dry-run' if is_dry_run() else 'LIVE'}): "
             f"{[(p, str(d), o) for p, d, o in results if o != 'no_cleaner_in_putzplan'] or 'nothing to fill'}")
    except Exception as exc:  # noqa: BLE001
        _log(f"[scheduler] ERROR in post-clean job: {exc!r}")


def run_monthly_report_job() -> None:
    from . import billing

    try:
        sent = billing.monthly_report(log=_log)
        _log(f"[scheduler] monthly Twilio cost report {'sent' if sent else 'NOT sent'}")
    except Exception as exc:  # noqa: BLE001
        _log(f"[scheduler] ERROR in monthly report: {exc!r}")


def trigger_check(reason: str = "event") -> None:
    """Event hook for routes (import, cancellation, manual search): run a
    check right away in a background thread. No-op if the scheduler is off."""
    if not is_enabled():
        return
    threading.Thread(target=run_check_job, args=(reason,), daemon=True).start()


def _next_0800(now: datetime.datetime) -> datetime.datetime:
    target = BERLIN_PYTZ.localize(datetime.datetime.combine(now.date(), datetime.time(8, 0)))
    if target <= now:
        target += datetime.timedelta(days=1)
    return target


def schedule_followups() -> None:
    """(Re)creates the one-shot wake-ups from the state file — see module
    docstring. Idempotent; safe to call after every state change."""
    scheduler = _scheduler
    if scheduler is None or is_dry_run():
        return
    from . import cleaner_coordination as cc

    now = datetime.datetime.now(BERLIN_PYTZ)
    with cc.STATE_LOCK:
        rows = cc._load_state().get("rows", {})
    held_quiet = False
    wanted = {}
    for key, rs in rows.items():
        if rs.get("cancelled") or rs.get("confirmed_cleaner_id"):
            continue
        attempts = rs.get("attempts", [])
        if attempts and attempts[-1].get("response") is None:
            due = datetime.datetime.fromisoformat(attempts[-1]["sent_at"]) + cc.RETRY_COOLDOWN
            # Only FUTURE timeouts get a job. A past-due one was either just
            # handled by the check that fired at `due` (nothing left to do —
            # re-scheduling it would loop forever, found in the first live
            # test) or is picked up by the startup check after a restart.
            if due > now:
                wanted[f"followup:{key}"] = due
        if rs.get("held_quiet"):
            held_quiet = True
    if held_quiet:
        wanted["quiet-release"] = _next_0800(now)
    for job in scheduler.get_jobs():
        if (job.id.startswith("followup:") or job.id == "quiet-release") and job.id not in wanted:
            job.remove()
    for job_id, run_at in wanted.items():
        scheduler.add_job(run_check_job, "date", run_date=run_at, kwargs={"reason": job_id},
                          id=job_id, replace_existing=True, misfire_grace_time=3600)


def start_scheduler():
    """Starts the scheduler once per process, if SCHEDULER_ENABLED. Returns
    the scheduler or None."""
    global _scheduler
    if not is_enabled():
        return None
    with _lock:
        if _scheduler is not None:
            return _scheduler
        from apscheduler.schedulers.background import BackgroundScheduler
        from apscheduler.triggers.cron import CronTrigger
        from . import cleaner_coordination as cc

        scheduler = BackgroundScheduler(timezone=BERLIN_PYTZ)
        for hour in (16, 19):
            scheduler.add_job(run_evening_job, CronTrigger(hour=hour, minute=0, timezone=BERLIN_PYTZ),
                              args=[hour], id=f"evening-{hour}", max_instances=1, coalesce=True,
                              misfire_grace_time=1800)
        scheduler.add_job(run_post_clean_job, CronTrigger(day_of_week="mon", hour=16, minute=0, timezone=BERLIN_PYTZ),
                          id="post-clean-weekly", max_instances=1, coalesce=True, misfire_grace_time=3600)
        scheduler.add_job(run_monthly_report_job, CronTrigger(day=1, hour=8, minute=0, timezone=BERLIN_PYTZ),
                          id="billing-monthly", max_instances=1, coalesce=True, misfire_grace_time=6 * 3600)
        scheduler.add_job(run_digest_job, CronTrigger(hour=18, minute=0, timezone=BERLIN_PYTZ),
                          id="digest-18", max_instances=1, coalesce=True, misfire_grace_time=1800)
        scheduler.add_job(lambda: __import__("app.clock_guard", fromlist=["check"]).check(
            notify=not is_dry_run(), log=_log), "date", id="startup-clock-check",
            run_date=datetime.datetime.now(BERLIN_PYTZ) + datetime.timedelta(seconds=20))
        # recovery after a restart: one full check shortly after startup
        scheduler.add_job(run_check_job, "date", run_date=datetime.datetime.now(BERLIN_PYTZ)
                          + datetime.timedelta(seconds=45), kwargs={"reason": "startup"}, id="startup-check")
        scheduler.start()
        _scheduler = scheduler
        cc.FOLLOWUP_HOOK = schedule_followups
        schedule_followups()
        _log(f"[scheduler] started: daily 16:00 + 19:00, event/timer driven — "
             f"{'DRY-RUN (log only)' if is_dry_run() else 'LIVE (sends WhatsApp!)'}")
        return scheduler
