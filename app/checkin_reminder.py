"""
Phase 2 (2026-09-25 decision, see docs/STATUS.md section 10.2): whenever a
new reservation gets added, email Farzaneh a reminder that also drops a
real alarm onto her iPhone for the day before that guest's check-in.

How the "email -> iPhone alarm" part works: the email carries a .ics
calendar attachment with a VALARM. Opening it in iOS Mail offers "Add to
Calendar"; once added, iOS itself fires a real alarm/notification at the
event's start time — no push-notification service or paid API needed.

Decisions Farzaneh made explicitly (do not change without asking her):
- Sent from her own Gmail (far.samsami@gmail.com) via an App Password.
- **Not** a daily cron/scheduler check — she revised this 2026-09-25 while
  phase 2 was being built: the reminder email goes out immediately when a
  reservation is written (see the hook in xlsx_writer.process_batch), not
  at some fixed time of day. `send_reminders_for_all_upcoming` below is the
  one-off backfill for bookings that already existed before this feature
  shipped.
- The calendar event/alarm inside the .ics is set for 16:00 Europe/Berlin,
  the day before the guest's Anreise (check-in) date, regardless of when
  the email itself happens to be sent/opened.
"""
import datetime
import os
import smtplib
import uuid
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders

from .config import PROPERTY_LABELS
from . import notified_checkins
from .tz import BERLIN
ALARM_HOUR = 16  # 16:00 Europe/Berlin, the day before check-in

GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
# Defaults to sending the reminder to Farzaneh herself (same account it's sent from).
NOTIFY_TO_EMAIL = os.environ.get("NOTIFY_TO_EMAIL") or GMAIL_USER


def parse_reservation_date(value):
    """Mirrors excel_reader._sort_key's date handling for a single field:
    openpyxl usually hands back a real date/datetime, but some manually
    edited cells are plain strings."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def reservation_key(reservation) -> str:
    """Stable identity for the notified_checkins.json dedup record. Prefers
    the confirmation code (matches known_codes.py's convention); falls back
    to guest name + checkin date for manually-entered rows without one."""
    code = reservation.get("confirmation_code")
    if code not in (None, ""):
        return f"code:{str(code).strip()}"
    checkin = parse_reservation_date(reservation.get("checkin"))
    return f"manual:{reservation.get('guest_name')}:{checkin}"


def _format_ics_datetime_utc(local_dt: datetime.datetime) -> str:
    return local_dt.astimezone(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _ics_escape(text: str) -> str:
    return str(text).replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", "\\n")


def _ics_uid(reservation, property_key: str) -> str:
    """Same UID for a reservation's reminder AND its later cancellation (if
    any) — required so a CANCEL .ics can even in principle refer back to
    the original event."""
    return f"checkin-reminder-{property_key}-{uuid.uuid5(uuid.NAMESPACE_URL, reservation_key(reservation))}@ferienwohnung-automat"


def build_ics(reservation, property_key: str, cancelled: bool = False) -> bytes:
    checkin_date = parse_reservation_date(reservation.get("checkin"))
    alarm_date = checkin_date - datetime.timedelta(days=1)
    start_local = datetime.datetime.combine(alarm_date, datetime.time(ALARM_HOUR, 0), tzinfo=BERLIN)
    end_local = start_local + datetime.timedelta(minutes=30)

    property_label = PROPERTY_LABELS.get(property_key, property_key)
    guest_name = reservation.get("guest_name") or "?"
    prefix = "STORNIERT — " if cancelled else ""
    summary = f"{prefix}Morgen kommt Gast: {guest_name} ({property_label})"
    description_lines = [
        f"Gast: {guest_name}",
        f"Objekt: {property_label}",
        f"Anreise: {checkin_date.strftime('%d.%m.%Y')}",
    ]
    if reservation.get("adults") is not None:
        description_lines.append(f"Erwachsene: {reservation.get('adults')}")
    if reservation.get("children") is not None:
        description_lines.append(f"Kinder: {reservation.get('children')}")
    if reservation.get("platform"):
        description_lines.append(f"Plattform: {reservation.get('platform')}")
    description = "\\n".join(_ics_escape(line) for line in description_lines)

    uid = _ics_uid(reservation, property_key)
    now_utc = _format_ics_datetime_utc(datetime.datetime.now(tz=BERLIN))

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Ferienwohnung Automat//Checkin Reminder//DE",
        "CALSCALE:GREGORIAN",
        f"METHOD:{'CANCEL' if cancelled else 'PUBLISH'}",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{now_utc}",
        f"DTSTART:{_format_ics_datetime_utc(start_local)}",
        f"DTEND:{_format_ics_datetime_utc(end_local)}",
        f"SEQUENCE:{1 if cancelled else 0}",
    ]
    if cancelled:
        lines.append("STATUS:CANCELLED")
    lines += [
        f"SUMMARY:{_ics_escape(summary)}",
        f"DESCRIPTION:{description}",
    ]
    if not cancelled:
        lines += [
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            f"DESCRIPTION:{_ics_escape(summary)}",
            "TRIGGER:PT0M",
            "END:VALARM",
        ]
    lines += [
        "END:VEVENT",
        "END:VCALENDAR",
        "",
    ]
    return "\r\n".join(lines).encode("utf-8")


def _send_email(subject: str, body_lines: list, ics_bytes: bytes, ics_method: str) -> None:
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        raise RuntimeError(
            "GMAIL_USER/GMAIL_APP_PASSWORD not set in .env — cannot send checkin reminder emails."
        )

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = GMAIL_USER
    msg["To"] = NOTIFY_TO_EMAIL
    msg.attach(MIMEText("\n".join(body_lines), "plain", "utf-8"))

    attachment = MIMEBase("text", "calendar", method=ics_method, name="erinnerung.ics")
    attachment.set_payload(ics_bytes)
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename="erinnerung.ics")
    msg.attach(attachment)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, [NOTIFY_TO_EMAIL], msg.as_string())


def send_reminder_email(reservation, property_key: str) -> None:
    property_label = PROPERTY_LABELS.get(property_key, property_key)
    guest_name = reservation.get("guest_name") or "?"
    checkin_date = parse_reservation_date(reservation.get("checkin"))

    body_lines = [
        f"Morgen ({checkin_date.strftime('%d.%m.%Y')}) kommt ein Gast an:",
        "",
        f"Gast: {guest_name}",
        f"Objekt: {property_label}",
    ]
    if reservation.get("adults") is not None:
        body_lines.append(f"Erwachsene: {reservation.get('adults')}")
    if reservation.get("children") is not None:
        body_lines.append(f"Kinder: {reservation.get('children')}")
    if reservation.get("platform"):
        body_lines.append(f"Plattform: {reservation.get('platform')}")
    body_lines += [
        "",
        "Im Anhang ein Kalender-Termin (.ics) — auf dem iPhone öffnen und "
        "zum Kalender hinzufügen, dann kommt der Alarm automatisch von iOS.",
    ]

    _send_email(
        subject=f"Morgen Ankunft: {guest_name} ({property_label})",
        body_lines=body_lines,
        ics_bytes=build_ics(reservation, property_key),
        ics_method="PUBLISH",
    )


def send_cancellation_email(reservation, property_key: str) -> None:
    property_label = PROPERTY_LABELS.get(property_key, property_key)
    guest_name = reservation.get("guest_name") or "?"
    checkin_date = parse_reservation_date(reservation.get("checkin"))
    alarm_date = checkin_date - datetime.timedelta(days=1)

    body_lines = [
        f"Diese Reservierung wurde storniert: {guest_name} ({property_label}), "
        f"Anreise war {checkin_date.strftime('%d.%m.%Y')}.",
        "",
        f"Bitte den Kalender-Eintrag vom {alarm_date.strftime('%d.%m.%Y')} um "
        f"{ALARM_HOUR}:00 Uhr ('Morgen kommt Gast: {guest_name}') manuell aus dem "
        "Kalender löschen — manche Kalender-Apps entfernen ihn nicht automatisch, "
        "auch wenn dieser Mail ein Storno-Termin beiliegt.",
    ]

    _send_email(
        subject=f"STORNIERT: {guest_name} ({property_label}) — Kalendereintrag löschen",
        body_lines=body_lines,
        ics_bytes=build_ics(reservation, property_key, cancelled=True),
        ics_method="CANCEL",
    )


def maybe_send_reminder(reservation, property_key: str, log=print) -> None:
    """Send the reminder for one reservation unless it's already been sent
    (notified_checkins.json) or its check-in is today/in the past (no
    "day before" left to alarm on). Never raises — a broken mailbox/App
    Password must not break the import that's writing real guest data;
    it only logs a warning so the failure is visible in the app's log
    output/Docker logs."""
    checkin_date = parse_reservation_date(reservation.get("checkin"))
    if checkin_date is None:
        return
    if checkin_date <= datetime.date.today():
        return  # nothing to alarm "the day before" for a same-day/past checkin
    if reservation.get("cancelled") == "ja":
        return

    key = reservation_key(reservation)
    if notified_checkins.was_notified(property_key, key):
        return

    try:
        send_reminder_email(reservation, property_key)
    except Exception as exc:  # noqa: BLE001 — must never break the caller's write flow
        log(f"    WARN: checkin reminder email failed for {reservation.get('guest_name')}: {exc}")
        return

    notified_checkins.record_notified(property_key, key)
    log(f"    sent checkin reminder email for {reservation.get('guest_name')} ({property_key})")


def maybe_send_cancellation_email(reservation, property_key: str, log=print) -> None:
    """Tells Farzaneh to delete the checkin-reminder calendar entry for a
    reservation that just got cancelled (app/xlsx_writer.apply_cancellation)
    — only if a reminder was actually sent for it (nothing to delete
    otherwise). `reservation` here is apply_cancellation's return dict:
    confirmation_code/checkin/guest_name read straight from the cancelled
    row, which is enough for reservation_key() to find the matching record.
    Never raises, same reasoning as maybe_send_reminder."""
    checkin_date = parse_reservation_date(reservation.get("checkin"))
    if checkin_date is None:
        return

    key = reservation_key(reservation)
    if not notified_checkins.was_notified(property_key, key):
        return  # no reminder was ever sent for this one, nothing to undo

    try:
        send_cancellation_email(reservation, property_key)
    except Exception as exc:  # noqa: BLE001 — must never break the caller's write flow
        log(f"    WARN: cancellation email failed for {reservation.get('guest_name')}: {exc}")
        return

    log(f"    sent cancellation email for {reservation.get('guest_name')} ({property_key})")


def send_reminders_for_all_upcoming(log=print) -> None:
    """One-off backfill (run via `python -m app.checkin_reminder`): sends
    the reminder for every currently-upcoming reservation in both property
    files that hasn't gotten one yet. Safe to re-run — already-notified
    reservations are skipped via notified_checkins.json."""
    from .config import PROPERTY_FILES
    from . import excel_reader

    for property_key in PROPERTY_FILES:
        for reservation in excel_reader.load_reservations(property_key):
            maybe_send_reminder(reservation, property_key, log=log)


if __name__ == "__main__":
    send_reminders_for_all_upcoming()
