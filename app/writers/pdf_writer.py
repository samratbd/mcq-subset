"""PDF writer — converts DOCX to PDF using the best available engine.

Priority:
  1. docx2pdf  — uses Microsoft Word via COM (Windows only, best quality).
     Requires:  pip install docx2pdf  AND  Microsoft Word installed.
  2. LibreOffice headless — free, cross-platform fallback.
     Requires:  LibreOffice installed (https://www.libreoffice.org).

If neither works, raises RuntimeError with install instructions.
"""

from __future__ import annotations
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from typing import List, Optional

from ..models import Question
from .docx_writer import write_docx_normal, write_docx_database

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


# ---------------------------------------------------------------------------
# Engine detection
# ---------------------------------------------------------------------------

def _docx2pdf_available() -> bool:
    """True only on Windows with docx2pdf + Word installed."""
    if sys.platform != "win32":
        return False  # docx2pdf is Windows-only
    try:
        import docx2pdf  # noqa: F401
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def _libreoffice_binary() -> Optional[str]:
    for name in ("libreoffice", "soffice"):
        path = shutil.which(name)
        if path:
            return path
    if sys.platform == "win32":
        for root in (
            r"C:\Program Files\LibreOffice\program",
            r"C:\Program Files (x86)\LibreOffice\program",
            r"C:\Program Files\LibreOffice 7\program",
        ):
            candidate = os.path.join(root, "soffice.exe")
            if os.path.exists(candidate):
                return candidate
    return None


def libreoffice_available() -> bool:
    return _libreoffice_binary() is not None


def pdf_engine_available() -> bool:
    return _docx2pdf_available() or libreoffice_available()


def pdf_engine_name() -> str:
    if _docx2pdf_available():
        return "Microsoft Word (docx2pdf)"
    if libreoffice_available():
        return "LibreOffice"
    return "none — install LibreOffice or (Windows) pip install docx2pdf"


# ---------------------------------------------------------------------------
# Converter
# ---------------------------------------------------------------------------

def docx_bytes_to_pdf_bytes(docx_bytes: bytes, *, timeout: int = 120) -> bytes:
    """Convert a DOCX blob to PDF using the best available engine."""
    if _docx2pdf_available():
        return _convert_via_docx2pdf(docx_bytes)
    binary = _libreoffice_binary()
    if binary:
        return _convert_via_libreoffice(docx_bytes, binary, timeout)
    raise RuntimeError(
        "No PDF engine found.\n\n"
        "Windows: pip install docx2pdf  (requires Microsoft Word)\n"
        "All platforms: install LibreOffice from https://www.libreoffice.org\n\n"
        "Restart the app after installing."
    )


def _convert_via_docx2pdf(docx_bytes: bytes) -> bytes:
    import docx2pdf
    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "in.docx")
        pdf_path = os.path.join(td, "in.pdf")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        try:
            docx2pdf.convert(docx_path, pdf_path)
        except Exception as e:
            raise RuntimeError(
                f"Word conversion failed: {e}\n"
                "Ensure Microsoft Word is installed and can open .docx files."
            )
        if not os.path.exists(pdf_path):
            raise RuntimeError("docx2pdf ran but produced no PDF.")
        with open(pdf_path, "rb") as f:
            return f.read()


def _convert_via_libreoffice(docx_bytes: bytes,
                              binary: str,
                              timeout: int) -> bytes:
    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "in.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        profile = os.path.join(td, f"lo_{uuid.uuid4().hex}")
        cmd = [
            binary,
            f"-env:UserInstallation=file://{profile}",
            "--headless", "--convert-to", "pdf",
            "--outdir", td, docx_path,
        ]
        try:
            r = subprocess.run(
                cmd, capture_output=True, timeout=timeout, check=False,
                creationflags=_NO_WINDOW,
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError(f"LibreOffice timed out after {timeout}s.")
        if r.returncode != 0:
            raise RuntimeError(
                "LibreOffice failed: "
                + (r.stderr.decode("utf-8", errors="replace")[:400]
                   or r.stdout.decode("utf-8", errors="replace")[:400])
            )
        pdf_path = os.path.join(td, "in.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError("LibreOffice ran but produced no PDF.")
        with open(pdf_path, "rb") as f:
            return f.read()


# Aliases used by the web server
def write_pdf_normal(questions: List[Question], *,
                     title: str = "", math_mode: str = "equation",
                     header_image: Optional[bytes] = None) -> bytes:
    return docx_bytes_to_pdf_bytes(
        write_docx_normal(questions, title=title,
                          math_mode=math_mode, header_image=header_image)
    )


def write_pdf_database(questions: List[Question], *,
                        title: str = "", math_mode: str = "equation",
                        header_image: Optional[bytes] = None) -> bytes:
    return docx_bytes_to_pdf_bytes(
        write_docx_database(questions, title=title,
                             math_mode=math_mode, header_image=header_image)
    )
