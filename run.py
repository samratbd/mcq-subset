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


def check_dependencies() -> None:
    """Verify critical deps work; print actionable errors if not.

    Most local-install failures present as a 500 error on the first request.
    Catching them here gives a clear message at startup instead.
    """
    problems = []
    try:
        import cv2  # noqa
    except ImportError:
        problems.append(
            "  - OpenCV (cv2) not installed.\n"
            "    Fix: pip install opencv-python-headless"
        )
    except OSError as e:
        msg = str(e)
        if "libGL" in msg or "libgl" in msg.lower():
            problems.append(
                "  - OpenCV is installed but a system library is missing.\n"
                "    On Linux: sudo apt install libgl1 libglib2.0-0\n"
                "    On macOS: brew install pkg-config\n"
                f"    Underlying error: {e}"
            )
        else:
            problems.append(f"  - OpenCV import failed: {e}")
    try:
        import numpy  # noqa
        import flask  # noqa
        import openpyxl  # noqa
        import docx  # noqa
    except ImportError as e:
        problems.append(
            f"  - Missing Python package: {e}.\n"
            "    Fix: pip install -r requirements.txt"
        )

    if problems:
        sys.stderr.write(
            "\n❌ MCQ Shuffler cannot start because of dependency problems:\n\n"
        )
        for p in problems:
            sys.stderr.write(p + "\n\n")
        sys.stderr.write(
            "See README.md → 'Troubleshooting' for full setup instructions.\n\n"
        )
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="MCQ Shuffler — local web app")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=5000, help="Bind port (default: 5000)")
    parser.add_argument("--no-browser", action="store_true",
                        help="Don't auto-open a browser tab")
    parser.add_argument("--debug", action="store_true",
                        help="Run Flask in debug mode (auto-reload)")
    parser.add_argument("--skip-checks", action="store_true",
                        help="Skip startup dependency checks")
    args = parser.parse_args()

    if not args.skip_checks:
        check_dependencies()

    from app.server import create_app
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
