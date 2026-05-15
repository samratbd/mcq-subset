#!/usr/bin/env python3
"""Desktop OMR Scanner — launcher.

Run this script to start the desktop OMR scanner. It checks for required
libraries and prints helpful install instructions if anything is missing.

Run with: python run_desktop.py
"""

from __future__ import annotations

import importlib
import platform
import shutil
import subprocess
import sys
import textwrap


REQUIRED = [
    ("cv2", "opencv-python-headless", "OpenCV (image processing)"),
    ("numpy", "numpy", "NumPy"),
    ("PIL", "Pillow", "Pillow (image I/O)"),
    ("openpyxl", "openpyxl", "openpyxl (Excel output — optional)"),
]


def _is_installed(module_name: str) -> bool:
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False


def _check_dependencies() -> list[tuple[str, str]]:
    missing = []
    for mod, pkg, name in REQUIRED:
        if not _is_installed(mod):
            missing.append((pkg, name))
    return missing


def _check_tkinter() -> bool:
    try:
        import tkinter  # noqa: F401
        return True
    except ImportError:
        return False


def _print_install_hint(missing_pkgs: list[tuple[str, str]]) -> None:
    print("\nMissing libraries:")
    for pkg, name in missing_pkgs:
        print(f"  - {name} (pip package: {pkg})")
    print(
        f"\nInstall them with:\n"
        f"\n    {sys.executable} -m pip install "
        + " ".join(p for p, _ in missing_pkgs)
        + "\n"
    )


def _print_tk_hint() -> None:
    system = platform.system()
    if system == "Linux":
        print(
            textwrap.dedent("""
            Tkinter is missing.

            On Ubuntu / Debian:
                sudo apt install python3-tk

            On Fedora / RHEL:
                sudo dnf install python3-tkinter

            On Arch:
                sudo pacman -S tk
            """).strip() + "\n"
        )
    elif system == "Darwin":
        print(
            textwrap.dedent("""
            Tkinter is missing. The easiest fix is to install Python from
            python.org — its installer includes Tkinter:

                https://www.python.org/downloads/macos/

            Then run this launcher with that Python interpreter.
            """).strip() + "\n"
        )
    else:  # Windows
        print(
            textwrap.dedent("""
            Tkinter is missing. Reinstall Python from python.org with the
            "tcl/tk and IDLE" option ticked:

                https://www.python.org/downloads/windows/
            """).strip() + "\n"
        )


def main() -> int:
    print("OMR Scanner Desktop launcher\n" + "=" * 32)
    print(f"Python: {sys.version}")
    print(f"Platform: {platform.system()} {platform.release()}\n")

    if not _check_tkinter():
        _print_tk_hint()
        return 1

    missing = _check_dependencies()
    # openpyxl is optional — we degrade to CSV if missing
    hard_missing = [(p, n) for p, n in missing if p != "openpyxl"]
    if hard_missing:
        _print_install_hint(hard_missing)
        return 1
    if missing:
        # Just openpyxl missing — warn but continue
        print("Note: openpyxl missing — Excel output will fall back to CSV.")
        print(f"      Install it for XLSX output: "
              f"{sys.executable} -m pip install openpyxl\n")

    # All good — launch GUI
    print("All dependencies OK. Launching application…\n")
    import desktop_omr
    desktop_omr.main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
