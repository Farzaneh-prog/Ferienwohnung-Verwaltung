"""Production entrypoint (used by waitress in the Docker container).

Unlike run.py, this never enables Flask debug mode — the app will be
reachable from the public internet once the container is deployed, and the
Werkzeug debugger allows arbitrary code execution if left on.
"""
from app import create_app

app = create_app()
