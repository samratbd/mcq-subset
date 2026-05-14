"""XLSX writer (same schema as the CSV writer)."""

from __future__ import annotations
import io
from typing import List

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

from ..models import Question


HEADER = ["title", "type", "option1", "option2", "option3", "option4",
          "answer", "explanation"]


def write_xlsx(questions: List[Question]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Questions"

    ws.append(HEADER)
    # Header styling: bold, top-aligned
    for col_idx in range(1, len(HEADER) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="top")

    for q in questions:
        ws.append([
            q.question,
            "MCQ",
            q.options[0], q.options[1], q.options[2], q.options[3],
            q.answer_index + 1,
            q.explanation,
        ])

    # Reasonable column widths; the content can still wrap.
    widths = [60, 8, 28, 28, 28, 28, 8, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    # Enable wrapping on body rows for the long text columns.
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
