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

    5 blocks of 20 questions; each block has 4 evenly-spaced answer bubbles.
    Y values measured by pixel projection through a clean bubble column.
    """
    t = SheetTemplate(
        name="omr_100",
        canonical_w=2500,
        canonical_h=1700,
        n_questions=100,
        bubble_radius=22,
        snap_search_radius=15,
    )

    roll_xs = [106, 178, 250, 323, 395, 477]
    # 10 digit rows, Y stride 84 (measured precisely)
    roll_ys = [154, 238, 322, 406, 490, 574, 658, 742, 826, 910]
    t.roll_bubbles = [[(x, y) for y in roll_ys] for x in roll_xs]

    t.set_bubbles = [(626, y) for y in roll_ys[:6]]

    # 20 rows, Y start 157, stride 75 (measured against blank to ±2 px)
    q_y_start, q_y_stride = 157, 75
    q_ys = [q_y_start + i * q_y_stride for i in range(20)]
    # 5 blocks, each 4 bubble columns at stride 75, block-to-block offset 345
    block_xs = [
        [790, 865, 940, 1015],     # Q01-20
        [1135, 1210, 1285, 1360],  # Q21-40
        [1480, 1555, 1630, 1705],  # Q41-60
        [1825, 1900, 1975, 2050],  # Q61-80
        [2170, 2245, 2320, 2395],  # Q81-100
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
