"""Dispatch the right parser based on file extension."""

from __future__ import annotations
import os
from typing import List

from ..models import Question
from .csv_parser import parse_csv_bytes
from .xlsx_parser import parse_xlsx_bytes
from .docx_parser import parse_docx_bytes


def parse_upload(filename: str, data: bytes) -> List[Question]:
    ext = os.path.splitext(filename or "")[1].lower()
    if ext == ".csv":
        return parse_csv_bytes(data)
    if ext == ".xlsx":
        return parse_xlsx_bytes(data)
    if ext == ".docx":
        return parse_docx_bytes(data)
    raise ValueError(
        f"Unsupported file type: {ext!r}. Use .csv, .xlsx or .docx."
    )
