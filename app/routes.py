from flask import Blueprint, render_template, request, redirect, url_for, session

from .auth import check_credentials, login_required
from .config import PROPERTY_LABELS
from .excel_reader import load_reservations

bp = Blueprint("main", __name__)


@bp.route("/", methods=["GET"])
def index():
    return redirect(url_for("main.reservations"))


@bp.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if check_credentials(username, password):
            session["logged_in"] = True
            next_url = request.args.get("next") or url_for("main.reservations")
            return redirect(next_url)
        error = "Benutzername oder Passwort ist falsch."
    return render_template("login.html", error=error)


@bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("main.login"))


@bp.route("/reservations")
@bp.route("/reservations/<property_key>")
@login_required
def reservations(property_key=None):
    if property_key not in PROPERTY_LABELS:
        property_key = next(iter(PROPERTY_LABELS))
    reservation_rows = load_reservations(property_key)
    return render_template(
        "reservations.html",
        reservations=reservation_rows,
        properties=PROPERTY_LABELS,
        current_property=property_key,
        property_label=PROPERTY_LABELS[property_key],
    )
