"""Downloadable sample/template files.

These are generated on demand from a small built-in question set so they
always reflect the current writer schemas exactly. Updating a writer's
column order automatically updates the samples — no static files to keep
in sync.

The samples cover every input format the parsers accept:

  - sample_questions.csv   — 5 rows; columns: title, type, option1..option4,
                             answer (1..4), explanation
  - sample_questions.xlsx  — same schema as CSV but as a real Excel file
  - sample_questions_normal.docx     — Normal Word layout (2-col + answer sheet)
  - sample_questions_database.docx   — Database Word layout (8-col table)

A user downloads one, replaces the content with their own questions while
keeping the structure, and uploads — round-trip safe.
"""

from __future__ import annotations
from typing import List, Tuple

from .models import Question
from .writers import write_csv, write_xlsx
from .writers.docx_writer import write_docx_normal, write_docx_database


def build_sample_paper() -> List[Question]:
    """Five mixed sample questions covering plain text, math, and English/Bengali.

    Designed so a user can open the file and immediately recognise what
    each column does without needing the README.
    """
    return [
        Question(
            sl=1,
            question="What is the SI unit of electric current?",
            options=["Volt", "Ampere", "Ohm", "Watt"],
            answer_index=1,  # B
            explanation="The ampere (A) is the SI base unit of electric current.",
        ),
        Question(
            sl=2,
            question="Energy has the dimension $ML^2T^{-2}$. Which equation is correct?",
            options=[
                "$E = mc^2$",
                "$E = \\dfrac{1}{2}mv^2$",
                "Both A and B",
                "Neither A nor B",
            ],
            answer_index=2,  # C — both
            explanation=(
                "Both $E=mc^2$ (rest energy) and $E=\\dfrac{1}{2}mv^2$ "
                "(kinetic energy) have the dimension $ML^2T^{-2}$."
            ),
        ),
        Question(
            sl=3,
            question="তড়িৎ প্রবাহের SI একক কোনটি?",
            options=["ভোল্ট", "অ্যাম্পিয়ার", "ওহম", "ওয়াট"],
            answer_index=1,  # B
            explanation="তড়িৎ প্রবাহের SI মৌলিক একক হলো অ্যাম্পিয়ার (A)।",
        ),
        Question(
            sl=4,
            question="A wire carries 4 A. After shuffling, the option positions "
                     "may change but the answer always matches.",
            options=["2 A", "4 A", "6 A", "8 A"],
            answer_index=1,  # B
            explanation=(
                "Tip: in shuffled sets the answer LETTER changes, but the "
                "option text at that letter is always the correct one. "
                "Verify this by comparing two generated sets side-by-side."
            ),
        ),
        Question(
            sl=5,
            question="Replace these rows with your own questions. The header "
                     "row and column count must stay the same.",
            options=["Plain text works", "$LaTeX$ math works",
                     "বাংলা works", "All of the above"],
            answer_index=3,  # D
            explanation=(
                "Keep the same column order. For CSV/Excel: title, type, "
                "option1..option4, answer (1-4), explanation. For Word "
                "files, use one of the two table layouts shown in the "
                ".docx sample."
            ),
        ),
    ]


# Sample registry: (filename) → (mimetype, builder_callable)
def _build_csv() -> bytes:
    return write_csv(build_sample_paper(), math_mode="katex")


def _build_xlsx() -> bytes:
    return write_xlsx(build_sample_paper(), math_mode="katex")


def _build_docx_normal() -> bytes:
    return write_docx_normal(
        build_sample_paper(),
        title="Sample question paper — Normal layout",
        math_mode="equation",
    )


def _build_docx_database() -> bytes:
    return write_docx_database(
        build_sample_paper(),
        title="Sample question paper — Database layout",
        math_mode="equation",
    )


_SAMPLES = {
    "sample_questions.csv": (
        "text/csv; charset=utf-8",
        "CSV template — open in Excel, replace rows, save. Math may be written as $LaTeX$.",
        _build_csv,
    ),
    "sample_questions.xlsx": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Excel template — same schema as the CSV, but with column widths and wrapping.",
        _build_xlsx,
    ),
    "sample_questions_normal.docx": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Word template (Normal layout): question table + answer sheet.",
        _build_docx_normal,
    ),
    "sample_questions_database.docx": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Word template (Database layout): single 8-column table.",
        _build_docx_database,
    ),
}


def sample_manifest():
    """Return a list of {filename, description} for the UI to render."""
    return [
        {"filename": name, "description": desc}
        for name, (_mime, desc, _build) in _SAMPLES.items()
    ]


def write_sample_to_bytes(filename: str) -> Tuple[bytes, str]:
    """Build a sample file on demand. Returns (bytes, mimetype)."""
    if filename not in _SAMPLES:
        raise KeyError(filename)
    mimetype, _desc, builder = _SAMPLES[filename]
    return builder(), mimetype
