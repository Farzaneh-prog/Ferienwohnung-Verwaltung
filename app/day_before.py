"""
The day before a cleaning (phase 3 stage 3, 2026-10-01): standard reminder,
cancellations, replacement bookings, contract-expiry alerts.

Rules (architecture doc §4, refined with Farzaneh):
- 16:00 the evening before: standard reminder ("morgen ist die Reinigung")
  to the cleaner in column A of every non-storniert Putzplan row for
  tomorrow. Also, once a day, an owner alert for cleaners whose contract
  ends soon.
- Cancellation (import, or the dashboard's "Stornierung melden"): the row
  is flagged 'storniert' in Putzplan but NOT deleted, and column A keeps the
  cleaner who was responsible (that's how we know who to tell). The cleaner
  is told "entfällt" at 19:00 the evening before, unless a replacement
  booking for the same date came in meanwhile (then nothing changes and the
  16:00/19:00 reminder goes out as usual).
- Entered after 19:00 the evening before (or on the day itself): the cancel
  notice goes out IMMEDIATELY, even during quiet hours — the only exception
  to the 22:00-08:00 rule, because the cleaner may already be on the way.
- Replacement booking after the cancel notice was sent: the same cleaner is
  asked first ("doch noch" template, Ja/Nein); on Nein/silence the normal
  chain starts over (without them). Replacement before the notice: they
  just stay.

Nothing here sends unless the scheduler is enabled (see scheduler.py), and
only logs while SCHEDULER_DRY_RUN is on.
"""
import datetime

from . import cleaner_coordination as cc
from . import putzplan_writer, whatsapp_sender
from .cleaner_roster import load_roster, match_cleaner, sheet_label
from .config import PROPERTY_LABELS
from .owner_alerts import notify_owner
from .tz import BERLIN

CANCEL_FINAL_HOUR = 19
CONTRACT_WARNING_DAYS = 30


def _enabled_and_dry():
    from . import scheduler

    return scheduler.is_enabled(), scheduler.is_dry_run()


def _cid_by_code(roster: dict, code):
    """Putzplan column A value -> cleaner id (see cleaner_roster.match_cleaner)."""
    return match_cleaner(roster, code)


def _vars(cleaner: dict, property_key: str, d: datetime.date) -> dict:
    return {"1": cleaner["name"], "2": d.strftime("%d.%m.%Y"), "3": PROPERTY_LABELS.get(property_key, property_key)}


def _plural(n: int, one: str, many: str) -> str:
    return f"{n} {one if n == 1 else many}"


def guests_text(guests: dict):
    """(variable 4, variable 5) for the reminder: who arrives next and whether
    there are children under 3. Column I (under 3) is filled by hand — blank
    means we simply don't know, and the message says so instead of guessing."""
    adults = guests.get("adults") if guests else None
    if adults in (None, ""):
        who = "noch keine nächste Buchung bekannt"
    else:
        kids = int(guests.get("children") or 0)
        who = _plural(int(adults), "Erwachsener", "Erwachsene") + ", " + (
            _plural(kids, "Kind unter 18", "Kinder unter 18") if kids else "keine Kinder")
    u3 = str(guests.get("children_u3") if guests else "").strip().lower()
    if u3 in ("", "none"):
        under3 = "Kinder unter 3 Jahren: keine Angabe."
    elif u3 in ("nein", "0", "no"):
        under3 = "Kinder unter 3 Jahren: nein."
    else:
        under3 = "Kinder unter 3 Jahren: ja" + (f" ({u3})." if u3.isdigit() else ".")
    return who, under3


def _send(cleaner: dict, template: str, property_key: str, d: datetime.date, log, extra: dict = None) -> bool:
    if not cleaner.get("whatsapp_number"):
        log(f"    [day-before] {cleaner['name']}: keine WhatsApp-Nummer — übersprungen")
        return False
    try:
        whatsapp_sender.send_template(cleaner["whatsapp_number"], template,
                                      dict(_vars(cleaner, property_key, d), **(extra or {})))
    except Exception as exc:  # noqa: BLE001
        log(f"    [day-before] WARN — {template} an {cleaner['name']} fehlgeschlagen: {exc}")
        return False
    return True


def cancel_deadline(checkout_date: datetime.date) -> datetime.datetime:
    """From this moment on a newly entered cancellation is announced at once."""
    return datetime.datetime.combine(checkout_date - datetime.timedelta(days=1),
                                     datetime.time(CANCEL_FINAL_HOUR, 0), tzinfo=BERLIN)


def _send_cancel_notices(property_key: str, checkout_date: datetime.date, roster: dict, state: dict,
                         dry_run: bool, log) -> None:
    """Tells the responsible cleaner — and everyone still waiting on this
    row (open request / Vielleicht / Reserve) — that the cleaning is off."""
    key = cc._state_key(property_key, checkout_date)
    rs = cc._row_state(state, key)
    ids = []
    if rs.get("cancelled_cleaner_id"):
        ids.append(rs["cancelled_cleaner_id"])
    for a in rs.get("attempts", []):
        if a.get("response") != "nein" and a["cleaner_id"] not in ids:
            ids.append(a["cleaner_id"])
    label = cc._label(property_key, checkout_date)
    if dry_run:
        log(f"    [day-before] {key}: Stornierungs-Info an {[roster[c]['name'] for c in ids if c in roster]} "
            f"— WÜRDE gesendet (dry-run)")
        return
    told = [roster[c]["name"] for c in ids if c in roster and _send(roster[c], "cancelled", property_key, checkout_date, log)]
    rs["cancel_notified"] = True
    log(f"    [day-before] {key}: Stornierungs-Info gesendet an {told}")
    notify_owner(f"Stornierung {label}: " + (f"Info an {', '.join(told)} gesendet" if told else "niemand zu informieren"),
                 log=log)


def on_cancellation(property_key: str, checkout_date, cleaner_code, now: datetime.datetime = None, log=print) -> None:
    """Called by xlsx_writer.process_batch (and the dashboard) right after
    the Putzplan row was flagged 'storniert'. `cleaner_code` is what column
    A held (None if nobody was assigned)."""
    enabled, dry_run = _enabled_and_dry()
    if not enabled or not isinstance(checkout_date, datetime.date):
        return
    now = now or datetime.datetime.now(tz=BERLIN)
    roster = load_roster()
    with cc.STATE_LOCK:
        state = cc._load_state()
        key = cc._state_key(property_key, checkout_date)
        if cc.test_mode() and not state["rows"].get(key, {}).get("test"):
            return  # test mode: real rows are not touched
        rs = cc._row_state(state, key)
        cid = _cid_by_code(roster, cleaner_code) or rs.get("confirmed_cleaner_id")
        label = cc._label(property_key, checkout_date)
        who = roster[cid]["name"] if cid in roster else "niemand zugewiesen"
        if dry_run:
            log(f"    [day-before] {key}: Stornierung erkannt (zuständig: {who}) — dry-run, nichts gespeichert/gesendet")
            return
        rs.update({"cancelled": True, "cancelled_cleaner_id": cid, "cancel_notified": False,
                   "cancelled_at": now.isoformat()})
        notify_owner(f"Stornierung → {label} (zuständig: {who})", log=log)
        if now >= cancel_deadline(checkout_date) and checkout_date >= now.date():
            _send_cancel_notices(property_key, checkout_date, roster, state, False, log)
        cc._save_state(state)


def on_new_reservation(property_key: str, checkout_date, revived_cleaner_code, now: datetime.datetime = None,
                       log=print) -> None:
    """Called after a new Putzplan row was added/revived for this date. Only
    matters if an earlier reservation for the same date was cancelled."""
    enabled, dry_run = _enabled_and_dry()
    if not enabled or not isinstance(checkout_date, datetime.date):
        return
    now = now or datetime.datetime.now(tz=BERLIN)
    roster = load_roster()
    with cc.STATE_LOCK:
        state = cc._load_state()
        key = cc._state_key(property_key, checkout_date)
        pre = state.get("pre_assigned", {}).get(key)
        if pre and pre in roster and not (cc.test_mode() and not state["rows"].get(key, {}).get("test")):
            # Pre-assigned by the owner before the booking existed: write the
            # name into Putzplan column A right away (data, not a message — so
            # also in dry-run) and keep the row out of the search chain.
            putzplan_writer.assign_cleaner(property_key, checkout_date, sheet_label(roster[pre]), log=log)
            cc._row_state(state, key)["confirmed_cleaner_id"] = pre
            state["pre_assigned"].pop(key)
            cc._save_state(state)
            notify_owner(f"Vorab-Zuweisung angewendet: {roster[pre]['name']} → {cc._label(property_key, checkout_date)}",
                         log=log)
            return
        rs = state["rows"].get(key)
        if not rs or not rs.get("cancelled") or (cc.test_mode() and not rs.get("test")):
            return
        cid = rs.get("cancelled_cleaner_id") or _cid_by_code(roster, revived_cleaner_code)
        name = roster[cid]["name"] if cid in roster else None
        label = cc._label(property_key, checkout_date)
        notified = rs.get("cancel_notified")
        if dry_run:
            log(f"    [day-before] {key}: Ersatzbuchung erkannt (zuständig war: {name}, "
                f"{'schon informiert → neu anfragen' if notified else 'bleibt'}) — dry-run")
            return
        rs["cancelled"] = False
        if not notified:
            if cid:
                rs["confirmed_cleaner_id"] = cid
            cc._save_state(state)
            notify_owner(f"Ersatzbuchung → {label}: {name or 'bisherige Zuweisung'} bleibt", log=log)
            return
        # Cleaner was already told "entfällt": start again, asking them first.
        rs.update({"attempts": [], "confirmed_cleaner_id": None, "reserves": [], "reopened": False,
                   "reminder_sent": False, "withdrawn": [], "cancel_notified": False,
                   "day_before_reminder": None})
        putzplan_writer.assign_cleaner(property_key, checkout_date, None, log=log)
        if cid in roster and _send(roster[cid], "again", property_key, checkout_date, log):
            cc.record_attempt(state, key, cid, roster[cid]["tier"], now, stage="again")
            notify_owner(f"Ersatzbuchung → {label}: frage {name} nochmal an", log=log)
        else:
            notify_owner(f"Ersatzbuchung → {label}: Suche startet neu", log=log)
        cc._save_state(state)
    cc.advance_row(property_key, checkout_date, now=now, roster=roster, log=log)


def run_evening_job(hour: int, now: datetime.datetime = None, log=print) -> None:
    """16:00 and 19:00 job (see scheduler.py) for TOMORROW's cleanings."""
    enabled, dry_run = _enabled_and_dry()
    if not enabled:
        return
    now = now or datetime.datetime.now(tz=BERLIN)
    tomorrow = now.date() + datetime.timedelta(days=1)
    roster = load_roster()
    with cc.STATE_LOCK:
        state = cc._load_state()

        # standard reminder for every active row (the 19:00 run also catches
        # rows that were just revived by a replacement booking)
        if cc.test_mode():  # test rows live only in state: treat "confirmed" as the Putzplan column A
            assigned = [(k.split("|", 1)[0], roster[rs["confirmed_cleaner_id"]]["name"], rs.get("test_guests") or {})
                        for k, rs in state["rows"].items()
                        if rs.get("test") and rs.get("confirmed_cleaner_id") in roster
                        and k.split("|", 1)[1] == tomorrow.isoformat()]
        else:
            assigned = putzplan_writer.iter_assigned_rows(tomorrow)
        for property_key, code, guests in assigned:
            cid = _cid_by_code(roster, code)
            if cid is None:
                log(f"    [day-before] {property_key} {tomorrow}: Kürzel '{code}' nicht im Roster — keine Erinnerung")
                continue
            key = cc._state_key(property_key, tomorrow)
            rs = state["rows"].get(key, {})
            if rs.get("cancelled") or rs.get("day_before_reminder") == tomorrow.isoformat():
                continue
            if dry_run:
                log(f"    [day-before] {key}: Erinnerung an {roster[cid]['name']} — WÜRDE gesendet (dry-run)")
                continue
            who, under3 = guests_text(guests)
            if _send(roster[cid], "reminder_tomorrow", property_key, tomorrow, log, extra={"4": who, "5": under3}):
                cc._row_state(state, key)["day_before_reminder"] = tomorrow.isoformat()
                notify_owner(f"Erinnerung an {roster[cid]['name']} → {cc._label(property_key, tomorrow)}", log=log)

        # cancellations nobody replaced by 19:00 -> tell the cleaner
        if hour >= CANCEL_FINAL_HOUR:
            for key, rs in list(state["rows"].items()):
                property_key, date_str = key.split("|", 1)
                if cc.test_mode() and not rs.get("test"):
                    continue
                if (rs.get("cancelled") and not rs.get("cancel_notified")
                        and datetime.date.fromisoformat(date_str) == tomorrow):
                    _send_cancel_notices(property_key, tomorrow, roster, state, dry_run, log)

        if hour < CANCEL_FINAL_HOUR:
            _contract_alerts(roster, state, now.date(), dry_run, log)
        if not dry_run:
            cc._save_state(state)


def _contract_alerts(roster: dict, state: dict, today: datetime.date, dry_run: bool, log) -> None:
    """Once per cleaner/contract date: owner alert when a contract ends
    within CONTRACT_WARNING_DAYS (the cleaner is dropped from the chain
    automatically afterwards — a replacement has to be arranged)."""
    sent = state.setdefault("alerts_sent_global", [])
    for cid, cleaner in roster.items():
        until = cleaner.get("contract_until")
        if not until:
            continue
        days = (datetime.date.fromisoformat(until) - today).days
        marker = f"contract:{cid}:{until}"
        if days <= CONTRACT_WARNING_DAYS and marker not in sent:
            text = (f"⚠️ Vertrag von {cleaner['name']} endet am "
                    f"{datetime.date.fromisoformat(until).strftime('%d.%m.%Y')} — Ersatz nötig.")
            if dry_run:
                log(f"    [day-before] Vertragswarnung — WÜRDE gesendet (dry-run): {text}")
            elif notify_owner(text, log=log):
                sent.append(marker)
