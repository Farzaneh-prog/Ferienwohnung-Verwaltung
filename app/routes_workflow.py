"""
Routes for the manual CSV/XLS import pipeline, the incomplete-row completion
form, and the quick cleaner-heads-up alert list. See app/import_parser.py
and app/xlsx_writer.py for the underlying logic — this file is just the web
glue (upload, preview, confirm) around it.
"""
import datetime
import json
import os
import uuid

from flask import Blueprint, render_template, request, redirect, url_for, flash

from .auth import login_required
from .config import PROPERTY_LABELS, PROPERTY_FILES, DATA_DIR
from .excel_reader import get_existing_confirmation_codes, find_incomplete_reservations
from .import_parser import detect_and_parse, merge_reservation_entries
from .quick_alerts import list_alerts, add_alert, resolve_alert
from .xlsx_writer import process_batch, update_reservation_fields
from . import cleaner_coordination as cc, day_before, putzplan_writer, scheduler
from .cleaner_roster import load_roster

bp = Blueprint("workflow", __name__)

STAGING_DIR = os.path.join(DATA_DIR, "data", "import_staging")
UPLOAD_DIR = os.path.join(DATA_DIR, "data", "uploads")


# ---------------------------------------------------------------- import ---

@bp.route("/import", methods=["GET"])
@login_required
def import_form():
    return render_template("import_form.html")


@bp.route("/import", methods=["POST"])
@login_required
def import_preview():
    files = request.files.getlist("files")
    files = [f for f in files if f and f.filename]
    if not files:
        flash("هیچ فایلی انتخاب نشده بود.")
        return redirect(url_for("workflow.import_form"))

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    entries_by_code = {}  # confirmation_code -> list of entries (one per file it appeared in)
    all_cancellations = []
    errors = []
    file_summaries = []

    for f in files:
        saved_path = os.path.join(UPLOAD_DIR, f.filename)
        f.save(saved_path)
        try:
            res, can, kind = detect_and_parse(saved_path, f.filename)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{f.filename}: {exc}")
            continue
        file_summaries.append({"filename": f.filename, "kind": kind, "count": len(res) + len(can)})
        for r in res:
            entries_by_code.setdefault(r["confirmation_code"], []).append(r)
        all_cancellations.extend(can)

    # Same booking in more than one uploaded file (e.g. Booking.com's
    # 'simple' and 'detailed' exports) — merge field-by-field rather than
    # picking one, since they carry different real numbers (see
    # xlsx_writer.append_reservation for how AA/AX get combined from them).
    all_reservations = {code: merge_reservation_entries(es) for code, es in entries_by_code.items()}

    existing_codes = {}
    to_add, duplicates = [], []
    for code, entry in all_reservations.items():
        prop = entry["property"]
        if prop not in existing_codes:
            existing_codes[prop] = get_existing_confirmation_codes(prop)
        if code in existing_codes[prop]:
            duplicates.append(entry)
        else:
            to_add.append(entry)

    token = uuid.uuid4().hex
    os.makedirs(STAGING_DIR, exist_ok=True)
    with open(os.path.join(STAGING_DIR, f"{token}.json"), "w", encoding="utf-8") as fh:
        json.dump({"new_reservations": to_add, "cancellations": all_cancellations}, fh, ensure_ascii=False)

    return render_template(
        "import_preview.html",
        token=token,
        file_summaries=file_summaries,
        errors=errors,
        to_add=to_add,
        duplicates=duplicates,
        cancellations=all_cancellations,
        property_labels=PROPERTY_LABELS,
    )


@bp.route("/import/confirm", methods=["POST"])
@login_required
def import_confirm():
    token = request.form.get("token", "")
    staging_path = os.path.join(STAGING_DIR, f"{token}.json")
    if not os.path.exists(staging_path):
        flash("این پیش‌نمایش منقضی شده — دوباره آپلود کنید.")
        return redirect(url_for("workflow.import_form"))

    with open(staging_path, "r", encoding="utf-8") as fh:
        batch = json.load(fh)

    log_lines = []
    # Confirm may run twice for the same preview (double click, retry after an
    # error) — never append a reservation whose code is already in the file.
    existing = {}
    fresh = []
    for r in batch["new_reservations"]:
        prop = r["property"]
        if prop not in existing:
            existing[prop] = get_existing_confirmation_codes(prop)
        if str(r.get("confirmation_code", "")).strip() in existing[prop]:
            log_lines.append(f"    skipped duplicate {r.get('confirmation_code')} ({prop}): already in the file")
        else:
            fresh.append(r)
    result = process_batch(fresh, batch["cancellations"], log=log_lines.append)
    os.remove(staging_path)
    scheduler.trigger_check("import")  # new/cancelled bookings -> look for cleaners right away

    return render_template(
        "import_result.html",
        applied=result["applied"],
        skipped=result["skipped"],
        log_lines=log_lines,
        property_labels=PROPERTY_LABELS,
    )


# -------------------------------------------------------------- complete ---

@bp.route("/complete", methods=["GET"])
@login_required
def complete_list():
    incomplete = {}
    for key in PROPERTY_FILES:
        incomplete[key] = find_incomplete_reservations(key)
    return render_template("complete_list.html", incomplete=incomplete, property_labels=PROPERTY_LABELS)


@bp.route("/complete/<property_key>/<code>", methods=["POST"])
@login_required
def complete_submit(property_key, code):
    updates = {}
    guest_name = request.form.get("guest_name", "").strip()
    guest_email = request.form.get("guest_email", "").strip()
    adults = request.form.get("adults", "").strip()
    children = request.form.get("children", "").strip()
    if guest_name:
        updates["guest_name"] = guest_name
    if guest_email:
        updates["guest_email"] = guest_email
    if adults:
        updates["adults"] = int(adults)
    if children:
        updates["children"] = int(children)
    if updates:
        update_reservation_fields(property_key, code, updates)
        flash(f"رزرو {code} به‌روزرسانی شد.")
    return redirect(url_for("workflow.complete_list"))


# ---------------------------------------------------------------- alerts ---
# "Putz-Alerts": heads-up for a date + manual "find a cleaner" trigger,
# cancellation reporting, and leave (Urlaub) — all feed cleaner_coordination.

def _parse_form_date(value):
    try:
        return datetime.date.fromisoformat((value or "").strip())
    except ValueError:
        return None


@bp.route("/alerts", methods=["GET"])
@login_required
def alerts_list():
    alerts = list_alerts()
    roster, state = load_roster(), cc._load_state()
    statuses = {}
    for a in alerts:
        d = _parse_form_date(a.get("date"))
        if d is None:
            continue
        status = cc.row_status(a["property"], d, roster, state)
        code = putzplan_writer.get_assigned_code(a["property"], d)
        if code and status in ("noch nicht gestartet",):
            status = f"im Putzplan eingetragen: {code}"
        statuses[a["id"]] = status
    pre = []
    for key, cid in sorted(state.get("pre_assigned", {}).items()):
        property_key, date_str = key.split("|", 1)
        pre.append({"property": property_key, "date": date_str, "cleaner": roster.get(cid, {}).get("name", cid)})
    return render_template("alerts.html", alerts=alerts, statuses=statuses, property_labels=PROPERTY_LABELS,
                           automation_on=scheduler.is_enabled(), dry_run=scheduler.is_dry_run(),
                           cleaners={cid: c["name"] for cid, c in roster.items()}, pre_assigned=pre)


@bp.route("/alerts", methods=["POST"])
@login_required
def alerts_add():
    property_key = request.form.get("property")
    date_str = request.form.get("date", "").strip()
    note = request.form.get("note", "").strip()
    d = _parse_form_date(date_str)
    if property_key in PROPERTY_LABELS and d:
        add_alert(property_key, date_str, note)
        if request.form.get("adults", "").strip().isdigit():
            # guests of THIS (not yet imported) booking -> Putzplan, like an import would
            try:
                result = putzplan_writer.register_manual_booking(
                    property_key, d, int(request.form["adults"]), int(request.form.get("children", "").strip() or 0),
                    {"ja": "Ja", "nein": "Nein"}.get(request.form.get("children_u3", "")),
                    _parse_form_date(request.form.get("checkin")))
                if result.get("previous_date"):
                    flash(f"Putzplan: Gästezahl bei der Reinigung am {result['previous_date']:%d.%m.%Y} eingetragen"
                          f"{' (Zeile für ' + d.strftime('%d.%m.%Y') + ' angelegt)' if result.get('created_row') else ''}.")
            except Exception as exc:  # noqa: BLE001
                flash(f"Gästezahl konnte nicht in den Putzplan geschrieben werden: {exc}")
        if scheduler.is_enabled():
            try:
                cc.start_manual_search(property_key, d, dry_run=scheduler.is_dry_run())
                flash("Suche nach Putzkraft gestartet." if not scheduler.is_dry_run() else "Dry-Run: Suche würde starten.")
            except Exception as exc:  # noqa: BLE001
                flash(f"Suche konnte nicht gestartet werden: {exc}")
    return redirect(url_for("workflow.alerts_list"))


@bp.route("/alerts/assign", methods=["POST"])
@login_required
def assign_cleaner_form():
    """Owner already knows who does a (new) booking: tell the system before it
    starts searching. Works before the booking is imported (remembered) and
    after (written to Putzplan at once)."""
    property_key = request.form.get("property")
    d = _parse_form_date(request.form.get("date"))
    cleaner_id = request.form.get("cleaner")
    roster = load_roster()
    if property_key in PROPERTY_LABELS and d and cleaner_id in roster:
        result = cc.assign_by_owner(property_key, d, cleaner_id)
        name = roster[cleaner_id]["name"]
        if result == "assigned":
            flash(f"{name} ist für {d:%d.%m.%Y} eingetragen (Putzplan Spalte A). Keine Suche nötig.")
        elif result == "pre_assigned":
            flash(f"Gespeichert: {name} für {d:%d.%m.%Y}. Sobald die Buchung importiert ist, wird sie eingetragen "
                  f"und es gibt keine Suche.")
        else:
            flash(f"Dieser Tag hat schon eine Putzkraft im Putzplan ({result.split(':', 1)[1]}) — nichts geändert.")
    else:
        flash("Ungültige Eingabe.")
    return redirect(url_for("workflow.alerts_list"))


@bp.route("/alerts/assign/delete", methods=["POST"])
@login_required
def assign_cleaner_delete():
    property_key = request.form.get("property")
    d = _parse_form_date(request.form.get("date"))
    if property_key in PROPERTY_LABELS and d:
        cc.remove_pre_assignment(property_key, d)
    return redirect(url_for("workflow.alerts_list"))


@bp.route("/alerts/<alert_id>/resolve", methods=["POST"])
@login_required
def alerts_resolve(alert_id):
    resolve_alert(alert_id)
    return redirect(url_for("workflow.alerts_list"))


@bp.route("/alerts/cancel", methods=["POST"])
@login_required
def cancellation_report():
    """Owner reports a cancelled date by hand (before/without an import)."""
    property_key = request.form.get("property")
    d = _parse_form_date(request.form.get("date"))
    if property_key in PROPERTY_LABELS and d:
        flagged = putzplan_writer.flag_putzplan_cancelled(property_key, d)
        if flagged is None:
            flash("Kein passender Putzplan-Eintrag gefunden — nichts geändert.")
        else:
            day_before.on_cancellation(property_key, d, flagged.get("cleaner_code"))
            flash("Stornierung eingetragen.")
    return redirect(url_for("workflow.alerts_list"))


# ----------------------------------------------------------------- leave ---

@bp.route("/leave", methods=["GET"])
@login_required
def leave_list():
    state = cc._load_state()
    return render_template("leave.html", roster=load_roster(), hard=state.get("leave_weeks", {}),
                           soft=state.get("soft_weeks", {}))


@bp.route("/leave/add", methods=["POST"])
@login_required
def leave_add():
    cid = request.form.get("cleaner")
    kind = "soft_weeks" if request.form.get("kind") == "soft" else "leave_weeks"
    start, end = _parse_form_date(request.form.get("start")), _parse_form_date(request.form.get("end"))
    if cid in load_roster() and start and end and start <= end:
        with cc.STATE_LOCK:
            state = cc._load_state()
            state.setdefault(kind, {}).setdefault(cid, []).append([start.isoformat(), end.isoformat()])
            cc._save_state(state)
    else:
        flash("Ungültige Eingabe (Person oder Datum).")
    return redirect(url_for("workflow.leave_list"))


@bp.route("/leave/delete", methods=["POST"])
@login_required
def leave_delete():
    cid, kind = request.form.get("cleaner"), request.form.get("kind")
    if kind in ("leave_weeks", "soft_weeks"):
        with cc.STATE_LOCK:
            state = cc._load_state()
            ranges = state.get(kind, {}).get(cid, [])
            idx = int(request.form.get("index", -1))
            if 0 <= idx < len(ranges):
                ranges.pop(idx)
                cc._save_state(state)
    return redirect(url_for("workflow.leave_list"))
