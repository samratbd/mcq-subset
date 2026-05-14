"""Output writers — dispatch by format name."""

from __future__ import annotations
from typing import List, Tuple

from ..models import Question
from .csv_writer import write_csv
from .xlsx_writer import write_xlsx
from .docx_writer import write_docx_normal, write_docx_database


def write_set(questions: List[Question], fmt: str, *,
              title: str = "Question Paper") -> Tuple[bytes, str]:
    """Render a question list in the requested format.

    Returns (bytes, suggested_extension).
    fmt: "csv" | "xlsx" | "docx_normal" | "docx_database"
    """
    fmt = (fmt or "").lower()
    if fmt == "csv":
        return write_csv(questions), "csv"
    if fmt == "xlsx":
        return write_xlsx(questions), "xlsx"
    if fmt == "docx_normal":
        return write_docx_normal(questions, title=title), "docx"
    if fmt == "docx_database":
        return write_docx_database(questions, title=title), "docx"
    raise ValueError(f"Unknown output format: {fmt!r}")
