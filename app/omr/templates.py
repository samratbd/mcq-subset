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

    Coordinates measured by Hough Circle detection on the user's blank
    `MCQ050202605150002.bmp`. Hough finds the GEOMETRIC center of each
    printed pink bubble ring (not biased by the letter glyph inside).

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

    # ROLL NUMBER — 6 columns × 10 rows. Hough X centers: 61, 147, 231, 317, 403, 487
    roll_xs = [61, 147, 231, 317, 403, 487]
    # Y stride matches roll = 80.4. 10 rows from Y=180 to Y=904
    roll_y_start, roll_y_end = 180, 904
    roll_ys = [int(round(roll_y_start + i * (roll_y_end - roll_y_start) / 9))
               for i in range(10)]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — 1 column at X=572, same Y as first 6 roll rows
    t.set_bubbles = [(572, y) for y in roll_ys[:6]]

    # Question rows — 25 rows from Y=99 to Y=2018, stride ~80
    q_y_start, q_y_end = 99, 2018
    q_ys = [int(round(q_y_start + i * (q_y_end - q_y_start) / 24))
            for i in range(25)]

    # Q1-25 X (Hough centers): 700, 785, 870, 956
    q1_25_xs = [700, 785, 870, 956]
    # Q26-50 X: 1083, 1167, 1251, 1337
    q26_50_xs = [1083, 1167, 1251, 1337]

    answer_bubbles = []
    for y in q_ys:
        answer_bubbles.append([(x, y) for x in q1_25_xs])
    for y in q_ys:
        answer_bubbles.append([(x, y) for x in q26_50_xs])
    t.answer_bubbles = answer_bubbles

    return t


def _build_100q() -> SheetTemplate:
    """100-question template (landscape 2500 × 1700).

    Coordinates measured by Hough Circle detection on the user's blank
    `MCQ100202605150002.bmp`. Bubble radius = 18 px.
    """
    t = SheetTemplate(
        name="omr_100",
        canonical_w=2500,
        canonical_h=1700,
        n_questions=100,
        bubble_radius=18,
        snap_search_radius=15,
    )

    # ROLL — 6 columns × 10 rows. Hough X centers: 90, 164, 237, 310, 384, 458
    roll_xs = [90, 164, 237, 310, 384, 458]
    roll_y_start, roll_y_end = 175, 873
    roll_ys = [int(round(roll_y_start + i * (roll_y_end - roll_y_start) / 9))
               for i in range(10)]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    # SET — column at X=568, with measured (slightly non-uniform) Y values
    t.set_bubbles = [
        (568, 176), (568, 251), (568, 329),
        (568, 407), (568, 484), (568, 564),
    ]

    # Question rows — 20 rows from Y=135 to Y=1602
    q_y_start, q_y_end = 135, 1602
    q_ys = [int(round(q_y_start + i * (q_y_end - q_y_start) / 19))
            for i in range(20)]

    # 5 blocks. Hough-measured X positions:
    block_xs = [
        [715, 788, 862, 935],      # Q01-20
        [1082, 1156, 1230, 1303],  # Q21-40
        [1451, 1524, 1597, 1671],  # Q41-60
        [1818, 1892, 1965, 2039],  # Q61-80
        [2186, 2259, 2333, 2406],  # Q81-100
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
