"""CSV writer.

Mirrors the input schema exactly:
    title, type, option1, option2, option3, option4, answer, explanation
- `answer` is 1..4
- KaTeX text passes through unchanged when math_mode="katex" (default),
  or is best-effort converted to Unicode when math_mode="unicode".
- Written with UTF-8 BOM so Excel opens it without mojibake.
"""

from __future__ import annotations
import csv
import io
from typing import List

from ..models import Question
from ..math_utils import render_text


HEADER = ["title", "type", "option1", "option2", "option3", "option4",
          "answer", "explanation"]


def write_csv(questions: List[Question], *, math_mode: str = "katex") -> bytes:
    if math_mode not in ("katex", "unicode"):
        raise ValueError(f"unknown math_mode for CSV: {math_mode!r}")

    def m(s: str) -> str:
        return render_text(s, math_mode)

    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
    w.writerow(HEADER)
    for q in questions:
        w.writerow([
            m(q.question),
            "MCQ",
            m(q.options[0]), m(q.options[1]), m(q.options[2]), m(q.options[3]),
            str(q.answer_index + 1),
            m(q.explanation),
        ])
    return "\ufeff".encode("utf-8") + buf.getvalue().encode("utf-8")
