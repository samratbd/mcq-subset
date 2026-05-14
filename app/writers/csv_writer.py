"""CSV writer.

Mirrors the input schema exactly:
    title, type, option1, option2, option3, option4, answer, explanation
- `answer` is 1..4
- KaTeX text passes through unchanged.
- Written with UTF-8 BOM so Excel opens it without mojibake.
"""

from __future__ import annotations
import csv
import io
from typing import List

from ..models import Question


HEADER = ["title", "type", "option1", "option2", "option3", "option4",
          "answer", "explanation"]


def write_csv(questions: List[Question]) -> bytes:
    buf = io.StringIO()
    # quoting=csv.QUOTE_MINIMAL → only quote when necessary; preserves clean look.
    w = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
    w.writerow(HEADER)
    for q in questions:
        w.writerow([
            q.question,
            "MCQ",
            q.options[0], q.options[1], q.options[2], q.options[3],
            str(q.answer_index + 1),
            q.explanation,
        ])
    # UTF-8 BOM so Excel auto-detects encoding (Bengali + KaTeX both render).
    return "\ufeff".encode("utf-8") + buf.getvalue().encode("utf-8")
