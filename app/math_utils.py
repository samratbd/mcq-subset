"""KaTeX helpers.

Two responsibilities:

1. Detect & split text containing inline KaTeX expressions delimited by `$...$`
   (single-dollar inline math, matching the user's CSV format). We do NOT support
   display math `$$...$$` because the input data is row-based.

2. Convert a KaTeX expression to an OMML <m:oMath> XML fragment, using pandoc
   as a subprocess when available. Results are cached per expression to avoid
   repeated subprocess invocations.

If pandoc isn't installed, `katex_to_omml_xml` returns None and the docx writer
falls back to plain text.
"""

from __future__ import annotations
import functools
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from typing import List, Optional, Tuple

from lxml import etree


# Single-dollar inline math. The negative lookbehind/lookahead guard against
# `$$` (display math, which we don't emit) and against escaped `\$`.
_INLINE_MATH_RE = re.compile(
    r"(?<!\\)\$(?!\$)(?P<expr>.+?)(?<!\\)\$(?!\$)",
    re.DOTALL,
)


def has_katex(text: str) -> bool:
    if not text:
        return False
    return _INLINE_MATH_RE.search(text) is not None


def split_text(text: str) -> List[Tuple[str, str]]:
    """Split a string into [(kind, value)] segments.

    kind is "text" or "math". A returned text segment never contains a math
    delimiter; a math segment is the inner LaTeX without the surrounding `$`.

    Order is preserved so callers can reassemble the line by walking segments.
    """
    if not text:
        return [("text", "")]
    out: List[Tuple[str, str]] = []
    pos = 0
    for m in _INLINE_MATH_RE.finditer(text):
        if m.start() > pos:
            out.append(("text", text[pos:m.start()]))
        out.append(("math", m.group("expr")))
        pos = m.end()
    if pos < len(text):
        out.append(("text", text[pos:]))
    if not out:
        out.append(("text", ""))
    return out


# ---------------------------------------------------------------------------
# pandoc-backed KaTeX → OMML conversion
# ---------------------------------------------------------------------------

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@functools.lru_cache(maxsize=1)
def _pandoc_path() -> Optional[str]:
    return shutil.which("pandoc")


def pandoc_available() -> bool:
    return _pandoc_path() is not None


@functools.lru_cache(maxsize=4096)
def katex_to_omml_xml(expr: str) -> Optional[str]:
    """Convert a single KaTeX expression to an OMML <m:oMath> XML fragment.

    Returns the OMML element as a UTF-8 XML string (no XML declaration),
    or None if conversion fails or pandoc isn't installed.
    """
    pandoc = _pandoc_path()
    if not pandoc:
        return None

    # Markdown source: a paragraph with a single inline-math expression.
    # Pandoc converts $...$ to a real OMML equation in the resulting docx.
    md = f"x ${expr}$ x\n"

    with tempfile.TemporaryDirectory() as td:
        md_path = os.path.join(td, "in.md")
        dx_path = os.path.join(td, "out.docx")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md)
        try:
            subprocess.run(
                [pandoc, md_path, "-o", dx_path, "--from", "markdown", "--to", "docx"],
                check=True, capture_output=True, timeout=15,
            )
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            return None

        try:
            with zipfile.ZipFile(dx_path) as z:
                doc_xml = z.read("word/document.xml")
        except Exception:
            return None

    try:
        root = etree.fromstring(doc_xml)
    except etree.XMLSyntaxError:
        return None

    omath = root.find(f".//{{{_M_NS}}}oMath")
    if omath is None:
        return None

    # Strip pandoc-set namespaces from the root attribute set; we'll
    # rely on the host docx declaring m: when we inject.
    return etree.tostring(omath, encoding="unicode")
