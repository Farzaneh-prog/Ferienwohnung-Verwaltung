"""Simple session-based single-user login (no user table, no signup)."""
import functools
import os

from flask import session, redirect, url_for, request
from werkzeug.security import check_password_hash


def check_credentials(username: str, password: str) -> bool:
    expected_username = os.environ.get("ADMIN_USERNAME", "")
    expected_hash = os.environ.get("ADMIN_PASSWORD_HASH", "")
    if not expected_username or not expected_hash:
        return False
    if username != expected_username:
        return False
    return check_password_hash(expected_hash, password)


def login_required(view):
    @functools.wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("main.login", next=request.path))
        return view(*args, **kwargs)

    return wrapped_view
