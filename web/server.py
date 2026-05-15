"""Web server entry point for Onurion OMR Studio.

Adds the project root to sys.path so the shared `app/` package is importable,
then delegates to app.server.create_app().

Run locally:
    python web/server.py

Deploy (gunicorn):
    gunicorn web.server:application
"""
import os
import sys

# Add project root to path so 'app' package is found
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)

from app.server import create_app  # noqa: E402

application = create_app()  # WSGI name used by gunicorn

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    application.run(host="0.0.0.0", port=port, debug=False)
