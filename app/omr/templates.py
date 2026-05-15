"""Bubble grid templates for each OMR sheet type.

Measured by pixel-projection on the user's blank templates and verified
against the actual sheets. Templates define every bubble position in
canonical (post-fiducial-warp) coordinates.

Reference frames:
  omr_50  → 1400 × 2080 (portrait)
  omr_100 → 2500 × 1700 (landscape)
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Tuple


@dataclass
class SheetTemplate:
    name: str
    canonical_w: int
    canonical_h: int
    n_questions: int
    n_digits: int = 6
    set_letters: str = "ABCDEF"
    bubble_radius: int = 22
    snap_search_radius: int = 12  # search ±this many px for the true bubble centre

    # roll_bubbles[d][digit_value]   → (x, y) for digit-column d, digit 0..9
    # set_bubbles[letter_idx]        → (x, y)
    # answer_bubbles[q][opt]         → (x, y), q 0..N-1, opt 0..3 (A..D)
    roll_bubbles: List[List[Tuple[int, int]]] = field(default_factory=list)
    set_bubbles: List[Tuple[int, int]] = field(default_factory=list)
    answer_bubbles: List[List[Tuple[int, int]]] = field(default_factory=list)


def _build_50q() -> SheetTemplate:
    """50-question template (portrait 1400 × 2080).

    All bubble coordinates measured against the user's blank
    (MCQ050202605150002.bmp) to pixel precision. Each printed bubble has
    a template circle centred to within ±2 px.
    """
    t = SheetTemplate(
        name="omr_50",
        canonical_w=1400,
        canonical_h=2080,
        n_questions=50,
        bubble_radius=22,
        snap_search_radius=18,
    )

    # ROLL NUMBER — 6 columns × 10 rows.
    # X stride ~83, Y stride ~78.
    roll_xs = [81, 165, 247, 330, 413, 496]
    roll_ys = [199, 277, 355, 434, 512, 590, 668, 747, 825, 903]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — 1 column at X=578, same Y stride as roll
    t.set_bubbles = [(578, y) for y in roll_ys[:6]]

    # Question rows — 25 rows from Y=118 to Y=1996, stride ~78.25
    q_y_start, q_y_end = 118, 1996
    q_ys = [int(round(q_y_start + i * (q_y_end - q_y_start) / 24)) for i in range(25)]

    # Q1-25 (left column): 4 options at stride ~80
    q1_25_xs = [711, 791, 870, 951]
    # Q26-50 (right column): same stride, shifted right by ~361
    q26_50_xs = [1072, 1152, 1232, 1312]

    answer_bubbles = []
    for y in q_ys:
        answer_bubbles.append([(x, y) for x in q1_25_xs])
    for y in q_ys:
        answer_bubbles.append([(x, y) for x in q26_50_xs])
    t.answer_bubbles = answer_bubbles

    return t


def _build_100q() -> SheetTemplate:
    """100-question template (landscape 2500 × 1700).

    All bubble coordinates measured against the user's blank
    (MCQ100202605150002.bmp). Pixel-perfect alignment verified.

    Layout:
      Roll: 6 columns × 10 rows
      SET:  1 column × 6 rows (A-F)
      Q01-100: 5 blocks of 4 columns × 20 rows
    """
    t = SheetTemplate(
        name="omr_100",
        canonical_w=2500,
        canonical_h=1700,
        n_questions=100,
        bubble_radius=22,
        snap_search_radius=18,
    )

    # ROLL — 6 columns × 10 rows. X stride ~74-75, Y stride ~75.6
    roll_xs = [103, 178, 255, 323, 398, 473]
    roll_y_start, roll_y_end = 198, 879
    roll_ys = [int(round(roll_y_start + i * (roll_y_end - roll_y_start) / 9))
               for i in range(10)]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — 1 column at X=580, first 6 roll Y values
    t.set_bubbles = [(580, y) for y in roll_ys[:6]]

    # Question rows — 20 rows from Y=154 to Y=1584
    q_y_start, q_y_end = 154, 1584
    q_ys = [int(round(q_y_start + i * (q_y_end - q_y_start) / 19))
            for i in range(20)]

    # 5 question blocks. X positions measured precisely per block.
    block_xs = [
        [731, 795, 863, 937],      # Q01-20
        [1091, 1156, 1223, 1298],  # Q21-40
        [1454, 1517, 1586, 1659],  # Q41-60
        [1813, 1878, 1946, 2021],  # Q61-80
        [2174, 2239, 2307, 2393],  # Q81-100
    ]

    answer_bubbles = []
    for block_x in block_xs:
        for y in q_ys:
            answer_bubbles.append([(x, y) for x in block_x])
    t.answer_bubbles = answer_bubbles

    return t


TEMPLATES = {
    "omr_50": _build_50q(),
    "omr_100": _build_100q(),
}


def get_template(sheet_type: str) -> SheetTemplate:
    if sheet_type not in TEMPLATES:
        raise ValueError(
            f"Unknown OMR sheet type: {sheet_type!r}. "
            f"Known: {sorted(TEMPLATES)}"
        )
    return TEMPLATES[sheet_type]
