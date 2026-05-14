"""CSV parser.

Schema (case-insensitive headers, BOM tolerated, comma-separated):
    title, type, option1, option2, option3, option4, answer, explanation

`answer` is the 1-based index (1..4) of the correct option.
`type` is informational and ignored (we assume MCQ).
KaTeX in any text field is preserved verbatim.
"""

from __future__ import annotations
import csv
import io
from typing import List

from ..models import Question, num_to_idx


_HEADERS = ["title", "type", "option1", "option2", "option3", "option4",
            "answer", "explanation"]


def _normalize_header(s: str) -> str:
    return (s or "").strip().lstrip("\ufeff").lower()


def parse_csv_bytes(data: bytes) -> List[Question]:
    # Try utf-8-sig first (strips BOM), then plain utf-8 as fallback.
    for enc in ("utf-8-sig", "utf-8"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV is not valid UTF-8")

    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        raise ValueError("CSV is empty")

    header = [_normalize_header(c) for c in rows[0]]
    # Build column index map; require all expected columns to exist.
    col_idx = {}
    for h in _HEADERS:
        if h not in header:
            raise ValueError(
                f"CSV missing required column '{h}'. Found: {header}"
            )
        col_idx[h] = header.index(h)

    questions: List[Question] = []
    for raw_sl, row in enumerate(rows[1:], start=1):
        if not row or all((c or "").strip() == "" for c in row):
            continue
        # Pad short rows defensively (trailing empty cells get dropped by Excel sometimes).
        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))

        title = (row[col_idx["title"]] or "").strip()
        opts = [
            (row[col_idx["option1"]] or "").strip(),
            (row[col_idx["option2"]] or "").strip(),
            (row[col_idx["option3"]] or "").strip(),
            (row[col_idx["option4"]] or "").strip(),
        ]
        ans_raw = (row[col_idx["answer"]] or "").strip()
        expl = (row[col_idx["explanation"]] or "").strip()

        if not title or any(o == "" for o in opts):
            raise ValueError(
                f"CSV row {raw_sl + 1}: title or one of option1..4 is empty"
            )

        try:
            answer_idx = num_to_idx(ans_raw)
        except ValueError as e:
            raise ValueError(f"CSV row {raw_sl + 1}: {e}") from None

        questions.append(Question(
            sl=raw_sl,
            question=title,
            options=opts,
            answer_index=answer_idx,
            explanation=expl,
        ))

    if not questions:
        raise ValueError("CSV had a header but no data rows")
    return questions
