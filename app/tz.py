"""Shared Europe/Berlin timezone object — every module that needs local
wall-clock time (checkin_reminder.py, cleaner_coordination.py) imports
BERLIN from here instead of repeating the Python-3.8-compat import dance."""
try:
    from zoneinfo import ZoneInfo  # stdlib on Python 3.9+
except ImportError:  # Python 3.8 (e.g. local dev machines) needs the backport
    from backports.zoneinfo import ZoneInfo

BERLIN = ZoneInfo("Europe/Berlin")
