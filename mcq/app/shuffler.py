"""Seeded shuffler for question papers.

Two independent toggles:
  - shuffle_questions: reorder the questions in the paper
  - shuffle_options:   reorder the 4 options inside each question
                       (answer_index is remapped to follow the correct text)

The shuffle is reproducible: given the same (paper_id, set_number, mode)
the output is byte-identical. We use a seeded random.Random for that.

After shuffling, `verify_set` re-checks that the correct option text still
sits at the new answer_index — if not, that's a bug we want to fail loudly.
"""

from __future__ import annotations
import copy
import hashlib
import random
from dataclasses import replace
from typing import List

from .models import Question


def _seed(paper_id: str, set_number: int, mode_tag: str) -> int:
    raw = f"{paper_id}|{set_number}|{mode_tag}".encode("utf-8")
    digest = hashlib.sha256(raw).digest()
    # 64-bit seed is plenty for random.Random()
    return int.from_bytes(digest[:8], "big")


def make_set(
    questions: List[Question],
    paper_id: str,
    set_number: int,
    shuffle_questions: bool,
    shuffle_options: bool,
) -> List[Question]:
    """Return a new list of Question objects representing one shuffled set.

    The input list is never mutated; every Question is deep-copied so callers
    can't accidentally share state across sets.
    """
    if not (1 <= set_number <= 99):
        raise ValueError("set_number must be between 1 and 99")

    mode_tag = f"q{int(shuffle_questions)}o{int(shuffle_options)}"
    rng = random.Random(_seed(paper_id, set_number, mode_tag))

    work = [copy.deepcopy(q) for q in questions]

    if shuffle_options:
        for q in work:
            perm = list(range(4))
            rng.shuffle(perm)
            # remember the correct text so we can find its new home
            correct_text = q.correct_option_text
            q.options = [q.options[i] for i in perm]
            # locate where the correct option ended up
            try:
                q.answer_index = q.options.index(correct_text)
            except ValueError:  # should be impossible
                raise RuntimeError(
                    f"Internal shuffle error: lost correct option for SL={q.sl}"
                )

    if shuffle_questions:
        rng.shuffle(work)

    # Renumber so SL reflects the new order (1-based)
    for new_sl, q in enumerate(work, start=1):
        q.sl = new_sl

    return work


def verify_set(original: List[Question], shuffled: List[Question]) -> None:
    """Validate that the shuffled set is internally consistent and matches the source.

    Raises AssertionError on any mismatch. Caller should regenerate or surface
    the error — we never silently return bad output.
    """
    if len(original) != len(shuffled):
        raise AssertionError(
            f"length mismatch: {len(original)} vs {len(shuffled)}"
        )

    # Build a multiset of (question, correct_option_text, sorted(options))
    # from the source. After shuffling, the same multiset must reappear.
    def signature(q: Question):
        return (q.question, q.correct_option_text, tuple(sorted(q.options)))

    src_sigs = sorted(signature(q) for q in original)
    out_sigs = sorted(signature(q) for q in shuffled)
    if src_sigs != out_sigs:
        raise AssertionError("question signatures differ — content was altered")

    # Per-question: the option at answer_index must be a real option
    for q in shuffled:
        if not (0 <= q.answer_index <= 3):
            raise AssertionError(f"SL={q.sl} bad answer_index {q.answer_index}")
        if q.options[q.answer_index] != q.correct_option_text:
            raise AssertionError(
                f"SL={q.sl} answer position points at wrong text"
            )

    # SL numbering is 1..N
    expected_sls = list(range(1, len(shuffled) + 1))
    actual_sls = [q.sl for q in shuffled]
    if actual_sls != expected_sls:
        raise AssertionError(
            f"SL renumbering broken: {actual_sls} vs {expected_sls}"
        )
