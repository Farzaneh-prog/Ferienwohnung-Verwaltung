import os

from dotenv import load_dotenv
from flask import Flask

load_dotenv()


def create_app():
    app = Flask(__name__)
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-insecure-key")

    from .routes import bp
    from .routes_workflow import bp as workflow_bp

    app.register_blueprint(bp)
    app.register_blueprint(workflow_bp)
    return app
