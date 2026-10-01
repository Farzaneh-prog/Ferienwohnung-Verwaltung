"""
System clock guard (2026-10-01): the NAS clock was once 1h37 behind (found
during the first live test). Every time-based rule here — the 16:00/19:00
jobs, quiet hours 22:00-08:00, the 19:00 cancellation deadline, the 1h
answer timeout — silently goes wrong with a wrong clock.

Compares the system clock with the HTTP "Date" header of well-known HTTPS
servers (needs nothing but outbound HTTPS, accurate to ~1-2 s).
- drift > WARN_SECONDS (15 min): warn in the log and, once per day, on the
  owner's WhatsApp (not in dry-run).
- drift > BLOCK_SECONDS (also 15 min): in LIVE mode scheduler.run_check_job skips the
  check — better to send nothing than to send at the wrong time.
- Fewer than two reachable servers: unknown drift, nothing blocked.
"""
import datetime
import email.utils
import time

import statistics

import requests

WARN_SECONDS = 900   # 15 min (chosen with Farzaneh: 2 min was too close to normal jitter)
BLOCK_SECONDS = 900
_SOURCES = ("https://www.google.com", "https://www.cloudflare.com", "https://www.microsoft.com", "https://www.apple.com")
_CACHE_SECONDS = 600

_cache = {"at": 0.0, "drift": None}
_last_alert_day = None


def drift_seconds(use_cache: bool = True):
    """system clock minus real time, in seconds (positive = system is ahead),
    or None if no reference could be reached. Uses the MEDIAN over all
    reachable sources: some servers answer HEAD requests with a stale cached
    Date header (api.twilio.com was 5 minutes off in the first test), a single
    source would give false alarms."""
    if use_cache and time.time() - _cache["at"] < _CACHE_SECONDS:
        return _cache["drift"]
    drifts = []
    for url in _SOURCES:
        try:
            before = time.time()
            response = requests.head(url, timeout=8, allow_redirects=False)
            after = time.time()
            date_header = response.headers.get("Date")
            if date_header:
                server = email.utils.parsedate_to_datetime(date_header).timestamp()
                drifts.append((before + after) / 2 - server)
        except Exception:  # noqa: BLE001 — try the next source
            continue
    drift = statistics.median(drifts) if len(drifts) >= 2 else None  # one source alone isn't trustworthy
    _cache.update(at=time.time(), drift=drift)
    return drift


def _format(drift: float) -> str:
    minutes, seconds = divmod(abs(int(drift)), 60)
    return f"{minutes} min {seconds} s {'vor' if drift > 0 else 'nach'}"


def check(notify: bool, log=print) -> bool:
    """Logs the result and, if notify, alerts the owner (once per day) when the
    drift is above WARN_SECONDS. Returns True if the clock looks fine or
    can't be verified."""
    global _last_alert_day
    drift = drift_seconds(use_cache=False)
    if drift is None:
        log("[clock-guard] keine Referenzzeit erreichbar — Uhr nicht geprüft")
        return True
    if abs(drift) <= WARN_SECONDS:
        log(f"[clock-guard] Uhr in Ordnung (Abweichung {drift:+.1f} s)")
        return True
    log(f"[clock-guard] WARNUNG: Systemuhr geht {_format(drift)} (Abweichung {drift:+.0f} s)")
    today = datetime.date.today()
    if notify and _last_alert_day != today:
        from .owner_alerts import notify_owner

        if notify_owner(f"⚠️ Die Uhr des NAS geht {_format(drift)}! Zeitgesteuerte Nachrichten sind unzuverlässig — "
                        f"bitte Zeitserver/Zeitzone im QNAP prüfen.", log=log):
            _last_alert_day = today
    return False


def ok_to_send() -> bool:
    """False only if the (cached) drift is large enough to block sending."""
    drift = drift_seconds()
    return drift is None or abs(drift) <= BLOCK_SECONDS
