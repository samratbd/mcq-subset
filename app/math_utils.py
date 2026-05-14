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


# ---------------------------------------------------------------------------
# Reverse direction: OMML → LaTeX/KaTeX (so Word source equations can be
# stored as KaTeX strings in our model and exported to CSV/XLSX).
# ---------------------------------------------------------------------------

# Minimal docx wrapper used to hand a single equation to pandoc. We slot a
# given OMML fragment into the body's first paragraph, then pandoc reads it
# and emits inline LaTeX.
_PROBE_DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:body>{PARA}<w:sectPr/></w:body></w:document>'
)
_PROBE_PARA_BEFORE = (
    '<w:p><w:r><w:t xml:space="preserve">LEAD </w:t></w:r>'
)
_PROBE_PARA_AFTER = (
    '<w:r><w:t xml:space="preserve"> TAIL</w:t></w:r></w:p>'
)
_PROBE_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
_PROBE_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)


@functools.lru_cache(maxsize=4096)
def omml_to_latex(omml_xml: str) -> Optional[str]:
    """Convert an OMML <m:oMath> XML string to an inline LaTeX expression.

    Returns the inner LaTeX (without surrounding $...$), or None if pandoc
    isn't available or conversion fails. Caller wraps with $...$ when
    embedding into a KaTeX-using text field.
    """
    pandoc = _pandoc_path()
    if not pandoc:
        return None
    if not omml_xml or "<m:oMath" not in omml_xml:
        return None

    para = f"{_PROBE_PARA_BEFORE}{omml_xml}{_PROBE_PARA_AFTER}"
    doc_xml = _PROBE_DOCUMENT_XML.format(PARA=para)

    with tempfile.TemporaryDirectory() as td:
        dx = os.path.join(td, "probe.docx")
        with zipfile.ZipFile(dx, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _PROBE_CONTENT_TYPES)
            z.writestr("_rels/.rels", _PROBE_RELS)
            z.writestr("word/document.xml", doc_xml)
        try:
            r = subprocess.run(
                [pandoc, dx, "-f", "docx", "-t", "markdown"],
                capture_output=True, text=True, timeout=15,
            )
            if r.returncode != 0:
                return None
        except subprocess.TimeoutExpired:
            return None

    # Extract the single $...$ from the line "LEAD $...$ TAIL"
    m = re.search(r"LEAD\s*\$(.+?)\$\s*TAIL", r.stdout, re.DOTALL)
    if not m:
        return None
    return m.group(1).strip()


# ---------------------------------------------------------------------------
# KaTeX → Unicode (best-effort plain text). Pandoc renders simple super/
# subscripts as Unicode (e.g. x^2 → x²); for complex cases (fractions,
# square roots) it keeps the LaTeX source. We expose what pandoc gives us.
# ---------------------------------------------------------------------------

@functools.lru_cache(maxsize=4096)
def katex_to_unicode(latex: str) -> str:
    """Best-effort conversion of a KaTeX expression to plain Unicode text.

    Falls back to "$latex$" verbatim if pandoc isn't available or can't
    handle it — so the source is never silently lost.
    """
    pandoc = _pandoc_path()
    if not pandoc:
        return f"${latex}$"

    md = f"LEAD ${latex}$ TAIL\n"
    try:
        r = subprocess.run(
            [pandoc, "-f", "markdown", "-t", "plain", "--wrap=none"],
            input=md, capture_output=True, text=True, timeout=10,
        )
        if r.returncode != 0:
            return f"${latex}$"
    except subprocess.TimeoutExpired:
        return f"${latex}$"

    out = r.stdout
    m = re.search(r"LEAD\s+(.*?)\s+TAIL", out, re.DOTALL)
    if not m:
        return f"${latex}$"
    val = m.group(1).strip()
    # If pandoc gave back the literal $...$ source, there was nothing it
    # could simplify — still return that as a Unicode fallback (callers can
    # decide whether to keep the dollars).
    return val


def render_text(text: str, mode: str) -> str:
    """Render a string containing `$...$` math segments into pure text in
    the requested mode.

      mode = "katex"   → return as-is ($...$ preserved)
      mode = "unicode" → replace each $expr$ with katex_to_unicode(expr)

    Used by data writers (CSV / XLSX) to optionally strip KaTeX.
    """
    if not text or mode == "katex":
        return text or ""
    if mode != "unicode":
        raise ValueError(f"unknown text-math mode: {mode!r}")
    out_parts = []
    for kind, val in split_text(text):
        if kind == "text":
            out_parts.append(val)
        else:
            out_parts.append(katex_to_unicode(val))
    return "".join(out_parts)
