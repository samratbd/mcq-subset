"""Canonical question model used across all parsers, writers, and the shuffler.

The model is intentionally minimal:
  - `options` is a fixed-length list of 4 strings.
  - `answer_index` is 0..3 (0 = option A / 1, 3 = option D / 4).
  - All strings (question, options, explanation) may contain KaTeX delimited
    by single dollars: `$...$`. We preserve them verbatim.

Conversions:
  - answer_index ↔ letter: idx_to_letter / letter_to_idx
  - answer_index ↔ 1-based number: idx + 1 / num - 1
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import List


@dataclass
class Question:
    sl: int                       # 1-based position in source
    question: str
    options: List[str]            # always length 4
    answer_index: int             # 0..3
    explanation: str = ""

    def __post_init__(self):
        if len(self.options) != 4:
            raise ValueError(
                f"Question SL={self.sl} must have exactly 4 options, "
                f"got {len(self.options)}"
            )
        if not (0 <= self.answer_index <= 3):
            raise ValueError(
                f"Question SL={self.sl} answer_index must be 0..3, "
                f"got {self.answer_index}"
            )

    @property
    def correct_option_text(self) -> str:
        """The text of the correct option — the invariant that must survive shuffling."""
        return self.options[self.answer_index]

    @property
    def answer_letter(self) -> str:
        return idx_to_letter(self.answer_index)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Question":
        return cls(
            sl=int(d["sl"]),
            question=d["question"],
            options=list(d["options"]),
            answer_index=int(d["answer_index"]),
            explanation=d.get("explanation", "") or "",
        )


# --- conversions -------------------------------------------------------------

_LETTERS = ["A", "B", "C", "D"]


def idx_to_letter(idx: int) -> str:
    if not (0 <= idx <= 3):
        raise ValueError(f"answer index {idx} out of range 0..3")
    return _LETTERS[idx]


def letter_to_idx(letter: str) -> int:
    letter = (letter or "").strip().upper()
    if letter not in _LETTERS:
        raise ValueError(f"answer letter {letter!r} not in A/B/C/D")
    return _LETTERS.index(letter)


def num_to_idx(n) -> int:
    """1/2/3/4 → 0/1/2/3."""
    try:
        i = int(str(n).strip())
    except (TypeError, ValueError):
        raise ValueError(f"answer number {n!r} is not 1..4")
    if not (1 <= i <= 4):
        raise ValueError(f"answer number {n!r} not in 1..4")
    return i - 1
