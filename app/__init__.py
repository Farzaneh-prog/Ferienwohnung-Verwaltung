import os

from dotenv import load_dotenv
from flask import Flask

load_dotenv()


def create_app():
    app = Flask(__name__)
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-insecure-key")

    from .routes import bp

    app.register_blueprint(bp)
    return app
