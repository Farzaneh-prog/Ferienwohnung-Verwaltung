from app import create_app

app = create_app()

if __name__ == "__main__":
    # use_reloader=False: avoids an incompatible watchdog/werkzeug combo in some
    # local Python environments; debug mode (tracebacks) still works fine.
    app.run(debug=True, use_reloader=False)
