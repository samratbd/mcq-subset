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

    Verified pixel-perfect against the user's blank template — every bubble
    aligns within ±5 px of its actual centre.
    """
    t = SheetTemplate(
        name="omr_50",
        canonical_w=1400,
        canonical_h=2080,
        n_questions=50,
        bubble_radius=22,
        snap_search_radius=12,
    )

    roll_xs = [84, 167, 250, 333, 416, 499]
    roll_ys = [200, 278, 356, 434, 512, 590, 668, 746, 824, 902]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    t.set_bubbles = [(580, y) for y in roll_ys[:6]]

    q_y_start, q_y_stride = 119, 78
    q_ys = [q_y_start + i * q_y_stride for i in range(25)]
    q1_25_xs = [712, 792, 872, 952]
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

    All coordinates measured by pixel-projection on the user's blank
    template (MCQ100202605150001.bmp), then verified against filled scans:

      - Roll number: 6 columns × 10 rows (digits 0-9)
      - SET:         1 column × 6 rows (letters A-F)
      - 5 question blocks (Q01-20, Q21-40, ...), each 4 columns × 20 rows
    """
    t = SheetTemplate(
        name="omr_100",
        canonical_w=2500,
        canonical_h=1700,
        n_questions=100,
        bubble_radius=22,
        snap_search_radius=15,
    )

    # ROLL NUMBER — 6 columns × 10 digit rows (Y stride 75, matching Q rows)
    roll_xs = [106, 178, 250, 323, 395, 477]
    roll_ys = [205, 280, 355, 430, 505, 580, 655, 730, 805, 880]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — 1 column × 6 letter rows (A-F), same Y as first 6 roll rows
    t.set_bubbles = [(580, y) for y in roll_ys[:6]]

    # Question rows — 20 rows shared across all 5 blocks
    q_y_start, q_y_stride = 156, 75
    q_ys = [q_y_start + i * q_y_stride for i in range(20)]

    # 5 question blocks. X positions measured directly from the blank.
    # Note the strides within a block are slightly uneven (64-68-74 px)
    # — this matches the actual printed bubble positions, not a uniform grid.
    block_xs = [
        [731, 795, 863, 937],      # Q01-20:   A B C D
        [1092, 1156, 1224, 1298],  # Q21-40
        [1454, 1517, 1586, 1660],  # Q41-60
        [1814, 1879, 1946, 2021],  # Q61-80
        [2175, 2239, 2307, 2399],  # Q81-100
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
