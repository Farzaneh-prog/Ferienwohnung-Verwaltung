"""
Routes for the manual CSV/XLS import pipeline, the incomplete-row completion
form, and the quick cleaner-heads-up alert list. See app/import_parser.py
and app/xlsx_writer.py for the underlying logic — this file is just the web
glue (upload, preview, confirm) around it.
"""
import json
import os
import uuid

from flask import Blueprint, render_template, request, redirect, url_for, flash

from .auth import login_required
from .config import PROPERTY_LABELS, PROPERTY_FILES, DATA_DIR
from .excel_reader import get_existing_confirmation_codes, find_incomplete_reservations
from .import_parser import detect_and_parse
from .quick_alerts import list_alerts, add_alert, resolve_alert
from .xlsx_writer import process_batch, update_reservation_fields

bp = Blueprint("workflow", __name__)

STAGING_DIR = os.path.join(DATA_DIR, "data", "import_staging")
UPLOAD_DIR = os.path.join(DATA_DIR, "data", "uploads")


def _richness(entry):
    """Higher = more complete, used to pick the best candidate when the same
    confirmation code shows up in more than one uploaded file."""
    score = 0
    if not str(entry.get("guest_name", "")).startswith("Gast (Code"):
        score += 2
    if entry.get("adults") is not None:
        score += 1
    if entry.get("guest_email"):
        score += 1
    return score


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
    all_reservations = {}  # confirmation_code -> best entry
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
            code = r["confirmation_code"]
            if code not in all_reservations or _richness(r) > _richness(all_reservations[code]):
                all_reservations[code] = r
        all_cancellations.extend(can)

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
    result = process_batch(batch["new_reservations"], batch["cancellations"], log=log_lines.append)
    os.remove(staging_path)

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

@bp.route("/alerts", methods=["GET"])
@login_required
def alerts_list():
    return render_template("alerts.html", alerts=list_alerts(), property_labels=PROPERTY_LABELS)


@bp.route("/alerts", methods=["POST"])
@login_required
def alerts_add():
    property_key = request.form.get("property")
    date_str = request.form.get("date", "").strip()
    note = request.form.get("note", "").strip()
    if property_key in PROPERTY_LABELS and date_str:
        add_alert(property_key, date_str, note)
    return redirect(url_for("workflow.alerts_list"))


@bp.route("/alerts/<alert_id>/resolve", methods=["POST"])
@login_required
def alerts_resolve(alert_id):
    resolve_alert(alert_id)
    return redirect(url_for("workflow.alerts_list"))
