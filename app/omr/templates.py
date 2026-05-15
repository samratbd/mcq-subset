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

    Coordinates measured against `MCQ050202605150002.bmp` using the
    OUTER-CORNER fiducial warp (each fiducial's outermost corner maps to
    the canonical image corner, giving zero margin).

    Bubble radius = 18 px (actual printed bubble diameter ≈ 37 px).
    """
    t = SheetTemplate(
        name="omr_50",
        canonical_w=1400,
        canonical_h=2080,
        n_questions=50,
        bubble_radius=18,
        snap_search_radius=15,
    )

    # ROLL NUMBER — 6 columns × 10 digit rows.
    # X measured at digit-0 row: 62, 148, 233, 318, 404, 489
    # Y from 181 to ~. Stride = roll stride (matches first 10 of SET-like grid).
    roll_xs = [62, 148, 233, 318, 404, 489]
    # 10 Y values: from 181 to 909, evenly spaced (stride ~81)
    roll_y_start, roll_y_end = 181, 909
    roll_ys = [int(round(roll_y_start + i * (roll_y_end - roll_y_start) / 9))
               for i in range(10)]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — 1 column at X=573, same Y values as first 6 roll rows
    t.set_bubbles = [(573, y) for y in roll_ys[:6]]

    # Question rows — 25 rows from Y=99 to Y=2018, stride ~80.0
    q_y_start, q_y_end = 99, 2018
    q_ys = [int(round(q_y_start + i * (q_y_end - q_y_start) / 24))
            for i in range(25)]

    # Q1-25 X: 710, 793, 873, 957  (stride ~82)
    q1_25_xs = [710, 793, 873, 957]
    # Q26-50 X: 1083, 1165, 1248, 1330
    q26_50_xs = [1083, 1165, 1248, 1330]

    answer_bubbles = []
    for y in q_ys:
        answer_bubbles.append([(x, y) for x in q1_25_xs])
    for y in q_ys:
        answer_bubbles.append([(x, y) for x in q26_50_xs])
    t.answer_bubbles = answer_bubbles

    return t


def _build_100q() -> SheetTemplate:
    """100-question template (landscape 2500 × 1700).

    Coordinates measured against `MCQ100202605150002.bmp` using the
    OUTER-CORNER fiducial warp. Bubble radius = 18 px.
    """
    t = SheetTemplate(
        name="omr_100",
        canonical_w=2500,
        canonical_h=1700,
        n_questions=100,
        bubble_radius=18,
        snap_search_radius=15,
    )

    # ROLL NUMBER — 6 columns × 10 rows.
    # Measured at digit-0 row (Y≈179): X = 81, 158, 236, 306, 384, 458
    # NOTE: col2-col3 stride (70) is smaller than others (~78); this matches
    # the actual printed layout (slight non-uniform spacing).
    roll_xs = [81, 158, 236, 306, 384, 458]
    # Y from 179 to 877, stride ~77.5
    roll_y_start, roll_y_end = 179, 877
    roll_ys = [int(round(roll_y_start + i * (roll_y_end - roll_y_start) / 9))
               for i in range(10)]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — 1 column at X=568, Y from 177 to 564
    set_y_start, set_y_end = 177, 564
    set_ys = [int(round(set_y_start + i * (set_y_end - set_y_start) / 5))
              for i in range(6)]
    t.set_bubbles = [(568, y) for y in set_ys]

    # Question rows — 20 rows from Y=135 to Y=1604
    q_y_start, q_y_end = 135, 1604
    q_ys = [int(round(q_y_start + i * (q_y_end - q_y_start) / 19))
            for i in range(20)]

    # 5 blocks of 4 columns each, measured precisely
    block_xs = [
        [721, 787, 855, 931],      # Q01-20
        [1087, 1154, 1223, 1299],  # Q21-40
        [1458, 1522, 1592, 1667],  # Q41-60
        [1823, 1890, 1959, 2035],  # Q61-80
        [2190, 2257, 2327, 2413],  # Q81-100
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
