"""PDF writer.

We generate PDF by rendering the matching DOCX layout, then converting it
with LibreOffice in headless mode. This gives perfect fidelity to the Word
output because we're literally rendering the same file.

LibreOffice is heavy (~600 MB on Linux), so the function checks for the
binary at call time and raises a clear error if it's not installed.

Concurrency note: LibreOffice's "user profile" is a global lock by default
— two simultaneous invocations would fight over `~/.config/libreoffice`.
We sidestep that by giving each invocation its own throwaway profile dir
via `-env:UserInstallation`.
"""

from __future__ import annotations
import os
import shutil
import subprocess
import tempfile
import uuid
from typing import List, Optional

from ..models import Question
from .docx_writer import write_docx_normal, write_docx_database


def _libreoffice_binary() -> Optional[str]:
    """Find a LibreOffice binary: ``libreoffice``, then ``soffice``."""
    return shutil.which("libreoffice") or shutil.which("soffice")


def libreoffice_available() -> bool:
    return _libreoffice_binary() is not None


def docx_bytes_to_pdf_bytes(docx_bytes: bytes, *, timeout: int = 90) -> bytes:
    """Convert a docx blob to PDF using LibreOffice headless mode."""
    binary = _libreoffice_binary()
    if not binary:
        raise RuntimeError(
            "LibreOffice is not installed on this server. "
            "PDF output requires libreoffice; install with `apt install "
            "libreoffice` (or use the project's Dockerfile, which bundles it). "
            "All other formats (Word, Excel, CSV) work without it."
        )

    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "in.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        # Per-invocation user profile dir avoids concurrent-call lock issues.
        profile_dir = os.path.join(td, f"lo_profile_{uuid.uuid4().hex}")
        cmd = [
            binary,
            f"-env:UserInstallation=file://{profile_dir}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", td,
            docx_path,
        ]
        try:
            result = subprocess.run(
                cmd, capture_output=True, timeout=timeout, check=False
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError(
                f"LibreOffice timed out after {timeout}s converting the doc."
            )

        if result.returncode != 0:
            raise RuntimeError(
                "LibreOffice failed to convert the document: "
                + (result.stderr.decode("utf-8", errors="replace")[:500]
                   or result.stdout.decode("utf-8", errors="replace")[:500])
            )

        pdf_path = os.path.join(td, "in.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError(
                "LibreOffice ran but didn't produce a PDF — likely a "
                "rendering failure inside the document."
            )
        with open(pdf_path, "rb") as f:
            return f.read()


def write_pdf_normal(questions: List[Question],
                     *,
                     title: str = "",
                     math_mode: str = "equation",
                     header_image: Optional[bytes] = None) -> bytes:
    docx = write_docx_normal(
        questions, title=title, math_mode=math_mode, header_image=header_image,
    )
    return docx_bytes_to_pdf_bytes(docx)


def write_pdf_database(questions: List[Question],
                       *,
                       title: str = "",
                       math_mode: str = "equation",
                       header_image: Optional[bytes] = None) -> bytes:
    docx = write_docx_database(
        questions, title=title, math_mode=math_mode, header_image=header_image,
    )
    return docx_bytes_to_pdf_bytes(docx)
