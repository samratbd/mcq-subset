"""DOCX writer.

Two output layouts, both matching the shape of the source files the user
uploaded:

* **normal**: 2 tables.
    1. Question table: 2 cols × N rows.
       Col 0 = SL ("01."). Col 1 = nested 4-column table:
           Row 0:  question (spans all 4 cols)
           Row 1:  [A.] [option A] [B.] [option B]
           Row 2:  [C.] [option C] [D.] [option D]
       The 2×2 option grid mirrors the source visually. The whole
       questions section is rendered in a **two-column page layout**, the
       way the source paginates questions side-by-side.
    2. Answer sheet: 3 cols × (N+1) rows in a single-column section.
       Explanation does *not* repeat the answer letter — it's already in
       column 2, and the redundant copy in the source would otherwise
       contradict the shuffled answer.

* **database**: single 8-column table (single-column page layout).

Math handling (per `math_mode`):
  - "equation": $...$ KaTeX → real Word equations via OMML.
  - "text":     $...$ kept verbatim as text.
  - "unicode":  best-effort conversion (e.g. `x^2` → `x²`).
"""

from __future__ import annotations
import copy
import io
from typing import List

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm
from docx.text.paragraph import Paragraph
from lxml import etree

from ..models import Question
from ..math_utils import split_text, katex_to_omml_xml, katex_to_unicode


# ---------------------------------------------------------------------------
# Rich-text helper: writes a string into a paragraph, rendering inline KaTeX
# according to math_mode.
# ---------------------------------------------------------------------------

def _add_rich(paragraph, text: str, *,
              bold: bool = False, size_pt: int = 11,
              math_mode: str = "equation"):
    if text is None:
        text = ""
    if math_mode not in ("equation", "text", "unicode"):
        raise ValueError(f"unknown math_mode: {math_mode!r}")

    for kind, val in split_text(text):
        if kind == "text":
            if val:
                run = paragraph.add_run(val)
                run.bold = bold
                run.font.size = Pt(size_pt)
            continue
        # kind == "math"
        if math_mode == "equation":
            omml = katex_to_omml_xml(val)
            if omml:
                try:
                    el = etree.fromstring(omml)
                    paragraph._p.append(el)
                    continue
                except etree.XMLSyntaxError:
                    pass
            literal = f"${val}$"
        elif math_mode == "unicode":
            literal = katex_to_unicode(val)
        else:  # "text"
            literal = f"${val}$"
        run = paragraph.add_run(literal)
        run.bold = bold
        run.font.size = Pt(size_pt)


def _set_cell_borders(cell, *, color: str = "cccccc"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = etree.SubElement(tcPr, qn("w:tcBorders"))
    for side in ("top", "left", "bottom", "right"):
        b = etree.SubElement(tcBorders, qn(f"w:{side}"))
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), color)


def _new_document(title: str | None = None) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    if title:
        h = doc.add_paragraph()
        run = h.add_run(title)
        run.bold = True
        run.font.size = Pt(14)
    return doc


def _set_section_columns(section, num_cols: int, space_dxa: int = 360):
    """Set the number of text columns on a section.

    `num_cols`=1 → single column (default); `num_cols`=2 → side-by-side.
    """
    sectPr = section._sectPr
    # Remove any existing cols element
    for existing in sectPr.findall(qn("w:cols")):
        sectPr.remove(existing)
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), str(num_cols))
    cols.set(qn("w:space"), str(space_dxa))
    sectPr.append(cols)


def _start_new_section(doc: Document, *, num_cols: int) -> None:
    """Insert a continuous section break and set its column count.

    Called once between the questions block and the answer sheet so the
    questions can be 2-column while the answer sheet stays single-column.
    """
    new_section = doc.add_section(WD_SECTION.CONTINUOUS)
    # Inherit page size/margins from the previous section by default.
    new_section.left_margin = Cm(1.5)
    new_section.right_margin = Cm(1.5)
    _set_section_columns(new_section, num_cols)


# ---------------------------------------------------------------------------
# Layout 1: Normal — 2x2 option grid + answer sheet
# ---------------------------------------------------------------------------

_LETTERS = ["A", "B", "C", "D"]

_W_NS_DECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
)


def _nested_table_xml(col_widths_dxa: list[int]) -> str:
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths_dxa)
    total = sum(col_widths_dxa)
    q_row = (
        '<w:tr>'
        '<w:tc>'
        '<w:tcPr><w:gridSpan w:val="4"/></w:tcPr>'
        '<w:p/>'
        '</w:tc>'
        '</w:tr>'
    )
    opt_cell = '<w:tc><w:tcPr/><w:p/></w:tc>'
    opt_row = '<w:tr>' + (opt_cell * 4) + '</w:tr>'
    return (
        f'<w:tbl {_W_NS_DECL}>'
        f'<w:tblPr>'
        f'<w:tblW w:w="{total}" w:type="dxa"/>'
        f'<w:tblInd w:w="180" w:type="dxa"/>'  # small left indent for visual breathing room
        f'<w:tblBorders>'
        f'<w:top w:val="nil"/><w:left w:val="nil"/>'
        f'<w:bottom w:val="nil"/><w:right w:val="nil"/>'
        f'<w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        f'</w:tblBorders>'
        f'</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{q_row}{opt_row}{opt_row}'
        f'</w:tbl>'
    )


def _fill_marker(tc_el, letter: str):
    p = tc_el.find(qn("w:p"))
    if p is None:
        p = etree.SubElement(tc_el, qn("w:p"))
    r = etree.SubElement(p, qn("w:r"))
    rPr = etree.SubElement(r, qn("w:rPr"))
    etree.SubElement(rPr, qn("w:b"))
    t = etree.SubElement(r, qn("w:t"))
    t.text = f"{letter}."


def _fill_paragraph(p_el, text: str, *, bold: bool = False,
                    math_mode: str = "equation"):
    para = Paragraph(p_el, parent=None)
    _add_rich(para, text, bold=bold, size_pt=10, math_mode=math_mode)


def write_docx_normal(questions: List[Question],
                      *,
                      title: str = "Question Paper",
                      math_mode: str = "equation") -> bytes:
    doc = _new_document(title)

    # Two-column page layout for the questions section. Questions flow
    # left-column then right-column, matching how the source paginates.
    _set_section_columns(doc.sections[0], num_cols=2, space_dxa=360)

    # Inner column widths for the nested 4-col option table. Total ~4400 DXA,
    # which fits within one of the two page columns at A4 1.5cm margins.
    # Marker columns must be wide enough that "A." doesn't wrap to a second
    # line at 10pt; ~700 DXA (~0.5") is comfortable.
    nested_widths = [700, 1500, 700, 1500]

    for q in questions:
        # One paragraph per question: SL inline, then the nested 2×2 grid
        # appears as a block immediately after. Keeping it all under one
        # outer "question block" makes the 2-column page layout flow nicely.
        sl_para = doc.add_paragraph()
        sl_run = sl_para.add_run(f"{q.sl:02d}. ")
        sl_run.bold = True
        sl_run.font.size = Pt(10)
        _add_rich(sl_para, q.question, bold=True, size_pt=10,
                  math_mode=math_mode)

        nested = etree.fromstring(_nested_table_xml(nested_widths))
        sl_para._p.addnext(nested)

        trs = nested.findall(qn("w:tr"))
        # Row 0 was reserved for the question; we already wrote the question
        # in sl_para, so collapse this row by leaving it empty.
        # (Some renderers won't accept a zero-row table, so we keep the
        # placeholder but make it visually negligible.)
        q_row_tc = trs[0].find(qn("w:tc"))
        # Make it 0 height so it doesn't visually duplicate the question.
        trPr = OxmlElement("w:trPr")
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), "0")
        h.set(qn("w:hRule"), "atLeast")
        trs[0].insert(0, trPr)

        # Row 1: A | optA | B | optB
        tcs = trs[1].findall(qn("w:tc"))
        _fill_marker(tcs[0], "A")
        _fill_paragraph(tcs[1].find(qn("w:p")), q.options[0],
                        math_mode=math_mode)
        _fill_marker(tcs[2], "B")
        _fill_paragraph(tcs[3].find(qn("w:p")), q.options[1],
                        math_mode=math_mode)

        # Row 2: C | optC | D | optD
        tcs = trs[2].findall(qn("w:tc"))
        _fill_marker(tcs[0], "C")
        _fill_paragraph(tcs[1].find(qn("w:p")), q.options[2],
                        math_mode=math_mode)
        _fill_marker(tcs[2], "D")
        _fill_paragraph(tcs[3].find(qn("w:p")), q.options[3],
                        math_mode=math_mode)

    # End the 2-column section before the answer sheet so the table renders
    # full-width across the page.
    _start_new_section(doc, num_cols=1)

    head = doc.add_paragraph()
    hr = head.add_run("Answer Sheet")
    hr.bold = True
    hr.font.size = Pt(13)

    atbl = doc.add_table(rows=1, cols=3)
    atbl.autofit = False
    hdr = atbl.rows[0].cells
    for i, label in enumerate(("Q No.", "Ans", "Explanation")):
        p = hdr[i].paragraphs[0]
        r = p.add_run(label)
        r.bold = True
        r.font.size = Pt(11)
        _set_cell_borders(hdr[i])

    for q in questions:
        row = atbl.add_row()
        c = row.cells
        c[0].paragraphs[0].add_run(f"{q.sl:02d}.")
        ar = c[1].paragraphs[0].add_run(q.answer_letter)
        ar.bold = True
        _add_rich(c[2].paragraphs[0], q.explanation,
                  size_pt=10, math_mode=math_mode)
        for cell in c:
            _set_cell_borders(cell)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Layout 2: Database (8-column single table)
# ---------------------------------------------------------------------------

def write_docx_database(questions: List[Question],
                        *,
                        title: str | None = None,
                        math_mode: str = "equation") -> bytes:
    doc = _new_document(title)

    tbl = doc.add_table(rows=0, cols=8)
    tbl.autofit = True

    for q in questions:
        row = tbl.add_row()
        c = row.cells
        c[0].paragraphs[0].add_run(f"{q.sl:02d}").bold = True
        _add_rich(c[1].paragraphs[0], q.question, bold=True, size_pt=10,
                  math_mode=math_mode)
        for i in range(4):
            _add_rich(c[2 + i].paragraphs[0], q.options[i], size_pt=10,
                      math_mode=math_mode)
        last = c[7].paragraphs[0]
        prefix = last.add_run(f"{q.answer_letter}; ")
        prefix.bold = True
        _add_rich(last, q.explanation, size_pt=10, math_mode=math_mode)
        for cell in c:
            _set_cell_borders(cell)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

