"""Bubble grid templates for each OMR sheet type.

A template defines, in the CANONICAL coordinate space (post perspective
correction), the (x, y) centre of every bubble plus the bubble radius:

  - roll_bubbles   : 6 digit columns × 10 rows (0..9)
  - set_bubbles    : 1 letter column × 6 rows (A..F)
  - answer_bubbles : N questions, each with 4 option bubbles (A, B, C, D)

The provisional coordinates below were derived by clustering filled-bubble
positions across the user's sample sheets, then fitting a regular grid.
They will be replaced with exact values once the user uploads a blank
(unfilled) template — the rest of the pipeline does not change.

Reference frame:
  50-question sheet  → canonical 1000 × 1500 (portrait)
  100-question sheet → canonical 1500 × 1000 (landscape)
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Tuple


@dataclass
class BubbleGrid:
    """A list of bubble centres and a common bubble radius."""
    positions: List[Tuple[int, int]]  # (x, y) in canonical pixels
    radius: int                        # bubble radius in pixels


@dataclass
class SheetTemplate:
    name: str                          # 'omr_50' or 'omr_100'
    canonical_w: int                   # canonical image width
    canonical_h: int                   # canonical image height
    n_questions: int                   # 50 or 100
    n_digits: int = 6                  # roll number digits
    set_letters: str = "ABCDEF"        # SET options
    bubble_radius: int = 22            # default radius

    # Each of these is a list aligned by question index / digit index / letter.
    # roll_bubbles[d][digit]   → (x, y) for digit-column d (0..5), digit value 0..9
    # set_bubbles[letter_idx]  → (x, y)
    # answer_bubbles[q][opt]   → (x, y) where q is 0..(N-1) and opt is 0..3 (A..D)
    roll_bubbles: List[List[Tuple[int, int]]] = field(default_factory=list)
    set_bubbles: List[Tuple[int, int]] = field(default_factory=list)
    answer_bubbles: List[List[Tuple[int, int]]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Provisional 50-question template (portrait, canonical 1000 × 1500)
# ---------------------------------------------------------------------------
# Question rows: Y starting at 160, stride 56, 25 rows.
# Answer-bubble columns:
#   Left  (Q1-25)  : X = 504, 560, 616, 672   (A, B, C, D)
#   Right (Q26-50) : X = 752, 808, 864, 920   (A, B, C, D)
# These match the cluster centroids found across 24 sample sheets.
#
# Roll-number grid: 6 digit columns; provisional coordinates — will be
# refined once a blank template is provided.

def _build_50q() -> SheetTemplate:
    t = SheetTemplate(
        name="omr_50",
        canonical_w=1000,
        canonical_h=1500,
        n_questions=50,
        bubble_radius=22,
    )

    # --- Answer bubbles ------------------------------------------------------
    # Bubble-column X positions (canonical coords). Derived from clustering
    # the centroids of filled bubbles across 24 sample sheets.
    LEFT_COL_X = [504, 560, 616, 672]    # A, B, C, D for Q01-25
    RIGHT_COL_X = [752, 808, 864, 920]   # A, B, C, D for Q26-50

    # Row Y positions. The empirical data shows 56 px stride but with a
    # couple of small drifts around Q9 and Q17 — likely physical paper
    # alignment noise in the source sheet. Listing each row explicitly is
    # more accurate than `Y0 + n*stride`.
    ROW_Y = [
        160, 216, 272, 328, 384, 440, 496, 552, 600, 656,
        712, 768, 824, 880, 936, 992, 1040, 1096, 1152, 1208,
        1264, 1320, 1368, 1424, 1480,
    ]
    assert len(ROW_Y) == 25

    answer_bubbles: List[List[Tuple[int, int]]] = []
    # Q1..Q25 → left block
    for y in ROW_Y:
        answer_bubbles.append([(x, y) for x in LEFT_COL_X])
    # Q26..Q50 → right block (same Y values)
    for y in ROW_Y:
        answer_bubbles.append([(x, y) for x in RIGHT_COL_X])
    t.answer_bubbles = answer_bubbles

    # --- Roll number & SET bubbles -------------------------------------------
    # PROVISIONAL placement. Replaced once the blank template is provided.
    roll_x = [80, 140, 200, 260, 320, 380]
    roll_y0 = 80
    roll_stride = 20
    t.roll_bubbles = [
        [(x, roll_y0 + d * roll_stride) for d in range(10)]
        for x in roll_x
    ]

    set_x = 450
    set_y0 = 80
    set_stride = 20
    t.set_bubbles = [(set_x, set_y0 + i * set_stride) for i in range(6)]

    return t


# ---------------------------------------------------------------------------
# Provisional 100-question template (landscape, canonical 1500 × 1000)
# ---------------------------------------------------------------------------
# 5 columns of 20 questions each:
#   Q1-20    : column 0  (Q label X≈250, bubbles to its right)
#   Q21-40   : column 1
#   Q41-60   : column 2
#   Q61-80   : column 3
#   Q81-100  : column 4
# Provisional coordinates; row stride ≈ 45 px, column stride ≈ 250 px.

def _build_100q() -> SheetTemplate:
    t = SheetTemplate(
        name="omr_100",
        canonical_w=1500,
        canonical_h=1000,
        n_questions=100,
        bubble_radius=20,
    )

    # 5 question blocks. Each block: bubble columns A, B, C, D positioned
    # to the right of a Q-number label.
    BLOCK_FIRST_BUBBLE_X = [310, 560, 810, 1060, 1310]
    BUBBLE_STRIDE_X = 40  # within a block, A→B→C→D
    ROW_Y0 = 80
    ROW_STRIDE = 45

    answer_bubbles: List[List[Tuple[int, int]]] = []
    for block in range(5):
        for row in range(20):
            y = ROW_Y0 + row * ROW_STRIDE
            xs = [BLOCK_FIRST_BUBBLE_X[block] + i * BUBBLE_STRIDE_X
                  for i in range(4)]
            answer_bubbles.append([(x, y) for x in xs])
    t.answer_bubbles = answer_bubbles

    # --- Roll number bubbles (PROVISIONAL) -----------------------------------
    roll_x = [80, 130, 180, 230, 280]  # 5 columns visible; 6th may overlap header
    # Adjust to 6 columns
    roll_x = [60, 105, 150, 195, 240, 285]
    roll_y0 = 90
    roll_stride = 18
    t.roll_bubbles = [
        [(x, roll_y0 + d * roll_stride) for d in range(10)]
        for x in roll_x
    ]

    # SET bubbles to the right of roll number
    set_x = 340
    set_y0 = 90
    set_stride = 18
    t.set_bubbles = [(set_x, set_y0 + i * set_stride) for i in range(6)]

    return t


TEMPLATES = {
    "omr_50": _build_50q(),
    "omr_100": _build_100q(),
}


def get_template(sheet_type: str) -> SheetTemplate:
    if sheet_type not in TEMPLATES:
        raise ValueError(
            f"Unknown OMR sheet type: {sheet_type!r}. "
            f"Known types: {sorted(TEMPLATES)}"
        )
    return TEMPLATES[sheet_type]
