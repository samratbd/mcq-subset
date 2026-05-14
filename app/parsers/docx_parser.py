"""DOCX parser.

Auto-detects two layouts:

* **Normal**: 2 tables.
    - Table 0: 2 columns × N rows. col-0 = "SL." string; col-1 contains a
      *nested* table whose paragraphs are
          [question, "A.", optA, "B.", optB, "C.", optC, "D.", optD]
      in that order.
    - Table 1: 3 columns × (N+1) rows. Header [Q No., Ans, Explanation]
      then one row per question: [SL., letter, explanation].

* **Database**: 1 table, 8 columns × N rows (no header), where each row is
      [blank, question, optA, optB, optC, optD, blank, "Letter; explanation"]

Detection is by table count and column shape — never by string sniffing.

KaTeX inside docx is rare (the format is usually rendered Unicode math), but
if it's present as `$...$` literal text we preserve it.
"""

from __future__ import annotations
import io
from typing import List

from docx import Document
from lxml import etree

from ..models import Question, letter_to_idx


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _omath_to_katex_string(omath_el) -> str:
    """Convert an <m:oMath> element to a `$...$` KaTeX string when possible.

    Bulletproof: any exception (no pandoc, subprocess crash, malformed XML,
    None return from a downstream helper, etc.) is swallowed and we fall
    through to the Unicode-text join of inner <m:t> nodes — so the math
    text is never lost and the parse never crashes on an upload.

    Returns a string. Never returns None.
    """
    # 1) Try the pandoc-backed OMML → LaTeX path.
    try:
        from ..math_utils import omml_to_latex
        xml = etree.tostring(omath_el, encoding="unicode")
        if isinstance(xml, str):
            latex = omml_to_latex(xml)
            if latex:
                return f"${latex}$"
    except Exception:
        # pandoc not installed, broken, timeout, or anything else — fall through.
        pass

    # 2) Fallback: concatenate <m:t> text nodes.
    try:
        return "".join(
            (t.text or "") for t in omath_el.iter(f"{{{_M_NS}}}t")
        )
    except Exception:
        return ""


def _paragraph_text(p_element) -> str:
    """Concatenate every text node inside a paragraph, in document order.

    Picks up:
      - normal <w:t> runs
      - math equation objects (<m:oMath>): rendered as `$LATEX$` when pandoc
        is available, otherwise as the concatenation of inner <m:t> nodes.

    Equations are *captured as units* so KaTeX delimiters survive. Naive
    iteration over <w:t> + <m:t> would lose the math boundary, leaving
    LaTeX fragments inside what looks like prose.
    """
    parts = []
    w_t = f"{{{_W_NS}}}t"
    m_omath = f"{{{_M_NS}}}oMath"

    def walk(node):
        try:
            children = list(node)
        except Exception:
            return
        for child in children:
            tag = getattr(child, "tag", None)
            if not isinstance(tag, str):
                # Comments, processing instructions, etc. — recurse over them
                # in case they contain elements (rare, but safe).
                try:
                    walk(child)
                except Exception:
                    pass
                continue
            if tag == w_t:
                if child.text:
                    parts.append(child.text)
            elif tag == m_omath:
                parts.append(_omath_to_katex_string(child) or "")
            else:
                walk(child)

    walk(p_element)
    return "".join(p for p in parts if p)


def _cell_paragraph_texts(tc_element) -> List[str]:
    """All paragraph texts inside a cell, including nested tables, in document order.

    Empty paragraphs are kept (so callers can use positional indexing) — but
    callers usually filter them out.
    """
    return [_paragraph_text(p) for p in tc_element.iter(f"{{{_W_NS}}}p")]


def _non_empty(paragraphs: List[str]) -> List[str]:
    return [p.strip() for p in paragraphs if p and p.strip()]


def parse_docx_bytes(data: bytes) -> List[Question]:
    doc = Document(io.BytesIO(data))
    tables = doc.tables
    if not tables:
        raise ValueError("DOCX has no tables")

    # Detection: prefer "Normal" if there's a 2-col first table + 3-col second table.
    is_normal = (
        len(tables) >= 2
        and len(tables[0].columns) == 2
        and len(tables[1].columns) == 3
    )
    is_database = len(tables[0].columns) == 8

    if is_normal:
        return _parse_normal(tables[0], tables[1])
    if is_database:
        return _parse_database(tables[0])

    raise ValueError(
        f"Unrecognised DOCX layout: {len(tables)} tables, "
        f"first table has {len(tables[0].columns)} columns. "
        "Expected Normal (2 cols + 3-col answer sheet) or Database (8 cols)."
    )


# --- Normal layout -----------------------------------------------------------

def _strip_answer_letter_prefix(expl: str) -> str:
    """Strip a redundant leading 'X;' / 'X)' / 'X.' prefix from an explanation.

    Both source layouts (Normal and Database) tend to start the explanation
    cell with the answer letter — e.g. "D; Is-377; ...". That letter is
    redundant with the dedicated answer column / output prefix. When the
    paper is shuffled, the redundant copy would otherwise stay frozen at
    the *original* letter and contradict the shuffled answer. We strip it
    here so the model holds just the body of the explanation.
    """
    import re
    if not expl:
        return ""
    m = re.match(r"^\s*([A-Da-d])\s*[;:.)]\s*(.*)$", expl, re.DOTALL)
    if m:
        return m.group(2).strip()
    return expl.strip()


def _parse_normal(q_table, ans_table) -> List[Question]:
    """Walk a 2-col question table + 3-col answer sheet."""
    # Build SL → (letter, explanation) map from the answer sheet.
    answers = {}
    # Skip the header row (row 0). It has been observed to contain the header
    # text "Q No. | Ans | Explanation".
    for row in ans_table.rows[1:]:
        cells = row.cells
        if len(cells) < 3:
            continue
        sl_raw = _cell_combined_text(cells[0]._tc).strip().rstrip(".")
        letter = _cell_combined_text(cells[1]._tc).strip().upper()
        expl = _strip_answer_letter_prefix(_cell_combined_text(cells[2]._tc))
        if not sl_raw or not letter:
            continue
        try:
            sl_int = int(sl_raw)
        except ValueError:
            continue
        answers[sl_int] = (letter, expl)

    questions: List[Question] = []
    for row in q_table.rows:
        cells = row.cells
        if len(cells) < 2:
            continue
        sl_raw = _cell_combined_text(cells[0]._tc).strip().rstrip(".")
        if not sl_raw:
            continue
        try:
            sl_int = int(sl_raw)
        except ValueError:
            continue

        paragraphs = _non_empty(_cell_paragraph_texts(cells[1]._tc))
        q_text, opts = _split_question_and_options(paragraphs, sl_int)

        if sl_int not in answers:
            raise ValueError(
                f"Question SL={sl_int} has no entry in the answer sheet"
            )
        letter, expl = answers[sl_int]
        try:
            ans_idx = letter_to_idx(letter)
        except ValueError as e:
            raise ValueError(f"Answer sheet row for SL={sl_int}: {e}") from None

        questions.append(Question(
            sl=sl_int,
            question=q_text,
            options=opts,
            answer_index=ans_idx,
            explanation=expl,
        ))

    if not questions:
        raise ValueError("Normal DOCX layout: no questions parsed")
    return questions


def _split_question_and_options(paragraphs: List[str], sl_int: int):
    """Given the ordered non-empty paragraphs inside a Normal-layout question cell,
    extract (question_text, [4 options]).

    The reliable shape is:
      paragraphs[0]   = full question text
      then markers "A.", "B.", "C.", "D." separating each option's text.

    Strategy:
      1. Strict mode: find paragraphs that are *exactly* "A.", "B.", "C.", "D."
         (in order). This handles edge cases like a question that itself
         starts with "A.C. Dynamo …" — that paragraph isn't *only* "A.",
         so it won't be mistaken for a marker.
      2. Lenient fallback: regex match for marker-fused-with-text, when the
         source isn't well-structured enough for strict mode.
    """
    if not paragraphs:
        raise ValueError(f"SL={sl_int}: question cell is empty")

    import re
    strict_marker_re = re.compile(r"^\s*([A-D])\s*[\.\)]\s*$")
    lenient_marker_re = re.compile(r"^\s*([A-Da-d])\s*[\.\)]\s*(.+)$")

    # --- Strict mode -------------------------------------------------------
    marker_positions = {}  # letter -> index in paragraphs
    for i, para in enumerate(paragraphs):
        m = strict_marker_re.match(para)
        if m:
            letter = m.group(1).upper()
            # Only keep the first occurrence per letter, in left-to-right order.
            if letter not in marker_positions:
                marker_positions[letter] = i

    if set(marker_positions) == {"A", "B", "C", "D"}:
        positions = [marker_positions[L] for L in ("A", "B", "C", "D")]
        if positions == sorted(positions):
            # Question text is everything before the "A." marker.
            q_text = " ".join(
                paragraphs[i].strip() for i in range(positions[0])
                if paragraphs[i].strip()
            ).strip()
            if not q_text:
                raise ValueError(f"SL={sl_int}: question text is empty")
            # Each option's text spans (marker+1) up to next marker (or end).
            opts = []
            bounds = positions + [len(paragraphs)]
            for idx in range(4):
                start = bounds[idx] + 1
                end = bounds[idx + 1]
                opt = " ".join(
                    paragraphs[i].strip() for i in range(start, end)
                    if paragraphs[i].strip()
                ).strip()
                if not opt:
                    raise ValueError(
                        f"SL={sl_int}: option {chr(ord('A')+idx)} is empty"
                    )
                opts.append(opt)
            return q_text, opts

    # --- Lenient fallback --------------------------------------------------
    q_parts: List[str] = []
    current_letter: str = ""
    buckets = {"A": [], "B": [], "C": [], "D": []}

    for para in paragraphs:
        m = lenient_marker_re.match(para)
        if m:
            current_letter = m.group(1).upper()
            tail = m.group(2).strip()
            if tail:
                buckets[current_letter].append(tail)
        else:
            if current_letter:
                buckets[current_letter].append(para.strip())
            else:
                q_parts.append(para.strip())

    question_text = " ".join(p for p in q_parts if p).strip()
    if not question_text:
        raise ValueError(f"SL={sl_int}: could not extract question text")

    opts = []
    for L in ("A", "B", "C", "D"):
        joined = " ".join(p for p in buckets[L] if p).strip()
        if not joined:
            raise ValueError(f"SL={sl_int}: option {L} is empty or missing")
        opts.append(joined)

    return question_text, opts


# --- Database layout ---------------------------------------------------------

def _cell_combined_text(tc_element) -> str:
    """Total text content of a cell.

    Same equation-aware walker as _paragraph_text — equations are emitted
    as `$LATEX$` strings, with plain `<w:t>` runs as Unicode text.

    Bulletproof: comments / PIs / unexpected node types are skipped silently
    so a malformed cell can't crash the whole parse.
    """
    parts = []
    w_t = f"{{{_W_NS}}}t"
    m_omath = f"{{{_M_NS}}}oMath"

    def walk(node):
        try:
            children = list(node)
        except Exception:
            return
        for child in children:
            tag = getattr(child, "tag", None)
            if not isinstance(tag, str):
                try:
                    walk(child)
                except Exception:
                    pass
                continue
            if tag == w_t:
                if child.text:
                    parts.append(child.text)
            elif tag == m_omath:
                parts.append(_omath_to_katex_string(child) or "")
            else:
                walk(child)

    walk(tc_element)
    return "".join(p for p in parts if p)


def _parse_database(table) -> List[Question]:
    """Walk an 8-column table where each row is one question."""
    questions: List[Question] = []
    for r_idx, row in enumerate(table.rows, start=1):
        cells = row.cells
        if len(cells) < 8:
            continue
        q_text = _cell_combined_text(cells[1]._tc).strip()
        if not q_text:
            # Likely a spacer/header row.
            continue
        opts = [_cell_combined_text(cells[i]._tc).strip() for i in (2, 3, 4, 5)]
        last_cell = _cell_combined_text(cells[7]._tc).strip()
        if not last_cell:
            raise ValueError(f"Database row {r_idx}: answer cell is empty")

        # Last cell format: "A; Is-377; (...) explanation..."
        # The leading letter is redundant with the separate answer column we
        # store in the model — strip it via the shared helper so it doesn't
        # contradict the shuffled answer in output.
        letter_part, sep, _rest = last_cell.partition(";")
        letter = letter_part.strip().upper()
        if letter not in ("A", "B", "C", "D"):
            raise ValueError(
                f"Database row {r_idx}: cannot find A/B/C/D letter "
                f"at start of answer cell (got {letter!r})"
            )
        try:
            ans_idx = letter_to_idx(letter)
        except ValueError as e:
            raise ValueError(f"Database row {r_idx}: {e}") from None
        explanation = _strip_answer_letter_prefix(last_cell)

        if any(o == "" for o in opts):
            raise ValueError(f"Database row {r_idx}: one of the options is empty")

        questions.append(Question(
            sl=r_idx,
            question=q_text,
            options=opts,
            answer_index=ans_idx,
            explanation=explanation,
        ))

    if not questions:
        raise ValueError("Database DOCX layout: no questions parsed")
    return questions
