"""Entry point: run the local web app.

Usage:
    python run.py [--host 127.0.0.1] [--port 5000] [--no-browser]

Opens your default browser to http://HOST:PORT after starting.
"""

from __future__ import annotations
import argparse
import os
import sys
import threading
import time
import webbrowser

from app.server import create_app


def main():
    parser = argparse.ArgumentParser(description="MCQ Shuffler — local web app")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=5000, help="Bind port (default: 5000)")
    parser.add_argument("--no-browser", action="store_true",
                        help="Don't auto-open a browser tab")
    parser.add_argument("--debug", action="store_true",
                        help="Run Flask in debug mode (auto-reload)")
    args = parser.parse_args()

    app = create_app()

    url = f"http://{args.host}:{args.port}/"
    if not args.no_browser:
        # Give Flask a moment to bind before opening the tab.
        def _open():
            time.sleep(0.8)
            try:
                webbrowser.open(url)
            except Exception:
                pass
        threading.Thread(target=_open, daemon=True).start()

    print(f"\n  MCQ Shuffler is starting at: {url}")
    print("  Press Ctrl+C to stop.\n")

    # use_reloader=False even in debug, because the reloader spawns a child
    # process which would double-trigger the browser-open thread.
    app.run(host=args.host, port=args.port, debug=args.debug, use_reloader=False)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nGoodbye.")
        sys.exit(0)
