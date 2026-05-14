"""Output writers — dispatch by format name."""

from __future__ import annotations
from typing import List, Tuple

from ..models import Question
from .csv_writer import write_csv
from .xlsx_writer import write_xlsx
from .docx_writer import write_docx_normal, write_docx_database


def write_set(questions: List[Question], fmt: str, *,
              title: str = "Question Paper",
              math_in_docx: str = "equation",
              math_in_data: str = "katex") -> Tuple[bytes, str]:
    """Render a question list in the requested format.

    fmt:           "csv" | "xlsx" | "docx_normal" | "docx_database"
    math_in_docx:  "equation" | "text" | "unicode" — math rendering for Word
    math_in_data:  "katex" | "unicode" — math representation for CSV/XLSX
    Returns (bytes, suggested_extension).
    """
    fmt = (fmt or "").lower()
    if fmt == "csv":
        return write_csv(questions, math_mode=math_in_data), "csv"
    if fmt == "xlsx":
        return write_xlsx(questions, math_mode=math_in_data), "xlsx"
    if fmt == "docx_normal":
        return write_docx_normal(questions, title=title,
                                 math_mode=math_in_docx), "docx"
    if fmt == "docx_database":
        return write_docx_database(questions, title=title,
                                   math_mode=math_in_docx), "docx"
    raise ValueError(f"Unknown output format: {fmt!r}")
