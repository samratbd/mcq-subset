"""OMR scanner — read an image, return a structured result.

The pipeline:

  1. Load the image (BMP/PNG/JPEG/etc.) into grayscale.
  2. Detect 4 corner fiducial markers.
  3. Perspective-warp to a canonical reference frame.
  4. Auto-detect sheet type (50- vs 100-question) from aspect ratio,
     or honour the caller's hint.
  5. For every bubble position in the template:
        a. Sample a circular patch around the centre.
        b. Compute "fill fraction" = ink coverage inside the circle.
  6. Classify each question / digit / set letter using configurable
     thresholds:
        empty     : fill < LOW_THRESHOLD   (treat as blank)
        ambiguous : LOW < fill < HIGH      (flag for review)
        filled    : fill ≥ HIGH_THRESHOLD
  7. Compute per-question confidence and aggregate per-sheet stats.

Public API:
  - scan_omr(image_bytes, sheet_type='auto') → OmrResult
  - render_review_image(image_bytes, result) → annotated PNG bytes
"""

from __future__ import annotations
import io
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Tuple

import cv2
import numpy as np

from .fiducial import detect_fiducials, warp_to_canonical
from .templates import SheetTemplate, get_template, TEMPLATES


# --- Tuneable thresholds -----------------------------------------------------

# Fraction of bubble area that must be dark for a bubble to be "filled".
LOW_THRESHOLD = 0.25   # Below this → empty.
HIGH_THRESHOLD = 0.55  # Above this → filled.
# Between LOW and HIGH → ambiguous (flag for review).

# How much darker a filled bubble must be than the mean empty bubble on the
# same sheet, before we count it as filled. Defends against over-darkening
# due to a poorly-exposed scan.
RELATIVE_FILL_MARGIN = 0.20

# Question-level confidence margin: difference between the most-filled and
# the second-most-filled option needed to call a question "confident".
CONFIDENT_OPTION_MARGIN = 0.30


# --- Public dataclasses ------------------------------------------------------

@dataclass
class OmrResult:
    """The data extracted from one OMR sheet."""
    sheet_type: str                       # "omr_50" or "omr_100"
    roll_number: str                      # 6-char digits string; '?' for unread
    set_letter: str                       # 'A'..'F' or '?'
    answers: List[str]                    # one entry per question
                                          # '' (blank), 'A'..'D', or 'A,C'
    # Diagnostics
    confidence: float                     # 0..1 — average per-question confidence
    needs_review: bool                    # any cell flagged ambiguous
    review_items: List[str]               # ['Q3', 'Q17', 'roll_d2', 'set'] etc.
    fill_fractions: List[List[float]]     # per-question: [fill_A, fill_B, fill_C, fill_D]
    # If processing failed, error contains the message and other fields may
    # be empty/zero. The caller decides whether to ship partial results.
    error: Optional[str] = None

    def as_dict(self) -> dict:
        d = asdict(self)
        # Reduce fill_fractions to ints for compactness
        d["fill_fractions"] = [
            [round(f * 100, 1) for f in row] for row in self.fill_fractions
        ]
        return d


# --- Image loading -----------------------------------------------------------

def _load_gray(image_bytes: bytes) -> np.ndarray:
    arr = np.frombuffer(image_bytes, dtype=np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_UNCHANGED)
    if img is None:
        raise ValueError("Could not decode the image. Supported: BMP, PNG, JPEG.")
    if img.ndim == 3:
        if img.shape[2] == 4:
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
        else:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    elif img.dtype != np.uint8:
        img = img.astype(np.uint8)
    return img


def _detect_sheet_type(img: np.ndarray) -> str:
    """Pick 50- vs 100-question template based on aspect ratio."""
    H, W = img.shape[:2]
    aspect = W / H
    return "omr_100" if aspect > 1.0 else "omr_50"


# --- Bubble sampling ---------------------------------------------------------

def _sample_bubble(warped: np.ndarray, cx: int, cy: int, r: int) -> float:
    """Return the fraction of dark pixels inside a circular bubble area.

    Threshold is local (Otsu over the patch) so we don't depend on global
    contrast. For the 1-bit BMP scans this collapses to a simple count of
    black pixels, which is exact.
    """
    H, W = warped.shape[:2]
    # Clip to image bounds.
    x0 = max(0, cx - r)
    y0 = max(0, cy - r)
    x1 = min(W, cx + r + 1)
    y1 = min(H, cy + r + 1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    patch = warped[y0:y1, x0:x1]

    # Build a circular mask.
    yy, xx = np.mgrid[0:patch.shape[0], 0:patch.shape[1]]
    cy_local = cy - y0
    cx_local = cx - x0
    mask = (xx - cx_local) ** 2 + (yy - cy_local) ** 2 <= r ** 2
    if not mask.any():
        return 0.0

    pixels = patch[mask]
    # Bitonal patch → already 0/255. Grayscale → Otsu threshold.
    if len(np.unique(pixels)) <= 2:
        dark = (pixels < 128).sum()
    else:
        try:
            t, _ = cv2.threshold(pixels.reshape(-1, 1), 0, 255,
                                 cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            dark = (pixels < t).sum()
        except cv2.error:
            dark = (pixels < 128).sum()
    return float(dark) / float(mask.sum())


def _sample_grid(warped: np.ndarray,
                 positions: List[Tuple[int, int]],
                 radius: int) -> List[float]:
    return [_sample_bubble(warped, x, y, radius) for x, y in positions]


# --- Classification logic ----------------------------------------------------

def _select_filled(fills: List[float],
                   sheet_avg_empty: float
                   ) -> Tuple[List[int], str]:
    """Given fill fractions for a single group (e.g. 4 options for a Q,
    or 10 digit rows for a roll-number column), return:

        (selected_indices, classification)

    where classification ∈ {"empty", "single", "multi", "ambiguous"}.

    Logic:
      * "Filled" = fill ≥ HIGH_THRESHOLD AND ≥ sheet_avg_empty + RELATIVE_FILL_MARGIN.
      * "Definitely empty" = fill < LOW_THRESHOLD.
      * Anything else is ambiguous.

      0 filled         → "empty" if all definitely-empty, else "ambiguous".
      1 filled, others empty   → "single".
      ≥2 filled, others empty  → "multi".
      Otherwise        → "ambiguous".
    """
    threshold_high = max(HIGH_THRESHOLD,
                         sheet_avg_empty + RELATIVE_FILL_MARGIN)
    filled = [i for i, f in enumerate(fills) if f >= threshold_high]
    empty_def = all(f < LOW_THRESHOLD for f in fills)

    if not filled:
        if empty_def:
            return [], "empty"
        return [], "ambiguous"

    # Some option(s) crossed the high bar — check the rest are clearly empty
    non_filled = [i for i in range(len(fills)) if i not in filled]
    if all(fills[i] < LOW_THRESHOLD for i in non_filled):
        return filled, ("single" if len(filled) == 1 else "multi")

    # Some "non-filled" are actually in the grey zone → ambiguous
    return filled, "ambiguous"


def _question_confidence(fills: List[float]) -> float:
    """Per-question confidence (0..1) based on the gap between the most-
    filled and second-most-filled option."""
    if not fills:
        return 0.0
    s = sorted(fills, reverse=True)
    top = s[0]
    second = s[1] if len(s) > 1 else 0.0
    margin = top - second
    if top < LOW_THRESHOLD:
        # All-empty case is high confidence (it's clearly a "blank")
        return 1.0
    return min(1.0, max(0.0, margin / CONFIDENT_OPTION_MARGIN))


def _opt_letter(i: int) -> str:
    return "ABCD"[i]


# --- Top-level scan ----------------------------------------------------------

def scan_omr(image_bytes: bytes,
             sheet_type: str = "auto") -> OmrResult:
    """Scan one OMR image and return the extracted data + diagnostics.

    sheet_type: 'auto', 'omr_50', or 'omr_100'.
    """
    try:
        gray = _load_gray(image_bytes)
    except Exception as e:
        return _error_result(sheet_type, str(e))

    # Auto-detect if requested
    if sheet_type == "auto":
        sheet_type = _detect_sheet_type(gray)

    template = get_template(sheet_type)

    # Find fiducials and warp
    try:
        fids = detect_fiducials(gray)
    except Exception as e:
        return _error_result(sheet_type,
                             f"Fiducial detection failed: {e}")

    warped = warp_to_canonical(
        gray, fids, template.canonical_w, template.canonical_h
    )

    # --- Sample every bubble -------------------------------------------------
    r = template.bubble_radius
    # 1) Answers — list of [4 fills] per question
    answer_fills: List[List[float]] = []
    for q_positions in template.answer_bubbles:
        answer_fills.append(_sample_grid(warped, q_positions, r))

    # 2) Roll digits — list of [10 fills] per digit column
    roll_fills: List[List[float]] = []
    for col_positions in template.roll_bubbles:
        roll_fills.append(_sample_grid(warped, col_positions, r))

    # 3) Set letters — single column of 6 fills
    set_fills: List[float] = _sample_grid(warped, template.set_bubbles, r)

    # --- Estimate the "empty bubble baseline" --------------------------------
    # Most option bubbles on the sheet are unfilled. Use the median of the
    # bottom 60% of all answer fills as a per-sheet baseline.
    all_fills = [f for row in answer_fills for f in row]
    if all_fills:
        sorted_fills = sorted(all_fills)
        n_low = max(1, int(len(sorted_fills) * 0.6))
        sheet_avg_empty = float(np.median(sorted_fills[:n_low]))
    else:
        sheet_avg_empty = 0.05

    # --- Classify answers ----------------------------------------------------
    answers: List[str] = []
    review_items: List[str] = []
    confidences: List[float] = []

    for q_idx, fills in enumerate(answer_fills):
        selected, kind = _select_filled(fills, sheet_avg_empty)
        if kind == "empty":
            answers.append("")
        elif kind == "single":
            answers.append(_opt_letter(selected[0]))
        elif kind == "multi":
            answers.append(",".join(_opt_letter(i) for i in selected))
        else:  # ambiguous
            # Output the best guess (the most-filled option), but flag it.
            best = int(np.argmax(fills))
            answers.append(_opt_letter(best))
            review_items.append(f"Q{q_idx + 1}")
        confidences.append(_question_confidence(fills))

    # --- Classify roll number digits -----------------------------------------
    roll_chars: List[str] = []
    for d_idx, fills in enumerate(roll_fills):
        selected, kind = _select_filled(fills, sheet_avg_empty)
        if kind == "single":
            roll_chars.append(str(selected[0]))
        elif kind == "empty":
            roll_chars.append("?")
            review_items.append(f"roll_d{d_idx + 1}")
        else:
            # ambiguous or multi → take best, flag.
            best = int(np.argmax(fills))
            roll_chars.append(str(best))
            review_items.append(f"roll_d{d_idx + 1}")
    roll_number = "".join(roll_chars)

    # --- Classify SET letter -------------------------------------------------
    selected, kind = _select_filled(set_fills, sheet_avg_empty)
    if kind == "single":
        set_letter = template.set_letters[selected[0]]
    elif kind == "empty":
        set_letter = "?"
        review_items.append("set")
    else:
        best = int(np.argmax(set_fills))
        set_letter = template.set_letters[best]
        review_items.append("set")

    overall_conf = float(np.mean(confidences)) if confidences else 0.0

    return OmrResult(
        sheet_type=sheet_type,
        roll_number=roll_number,
        set_letter=set_letter,
        answers=answers,
        confidence=overall_conf,
        needs_review=bool(review_items),
        review_items=review_items,
        fill_fractions=answer_fills,
    )


def _error_result(sheet_type: str, msg: str) -> OmrResult:
    return OmrResult(
        sheet_type=sheet_type if sheet_type != "auto" else "omr_50",
        roll_number="?",
        set_letter="?",
        answers=[],
        confidence=0.0,
        needs_review=True,
        review_items=["scan_failed"],
        fill_fractions=[],
        error=msg,
    )


# --- Annotated review image --------------------------------------------------

def render_review_image(image_bytes: bytes,
                        result: OmrResult,
                        max_height: int = 1400) -> bytes:
    """Render a debug image: warped sheet with bubble positions overlaid,
    each colour-coded by classification.

        green  — clearly empty
        red    — clearly filled (selected as the answer)
        orange — ambiguous (flagged for review)
    """
    try:
        gray = _load_gray(image_bytes)
        fids = detect_fiducials(gray)
        template = get_template(result.sheet_type)
        warped = warp_to_canonical(
            gray, fids, template.canonical_w, template.canonical_h
        )
    except Exception:
        return b""

    canvas = cv2.cvtColor(warped, cv2.COLOR_GRAY2BGR)

    # Draw every answer bubble
    r = template.bubble_radius
    for q_idx, q_positions in enumerate(template.answer_bubbles):
        fills = result.fill_fractions[q_idx] if q_idx < len(result.fill_fractions) else []
        flagged = f"Q{q_idx + 1}" in result.review_items
        selected_letters = set()
        if q_idx < len(result.answers):
            for letter in result.answers[q_idx].split(","):
                if letter:
                    selected_letters.add(letter)
        for opt_idx, (x, y) in enumerate(q_positions):
            letter = _opt_letter(opt_idx)
            if flagged:
                color = (0, 140, 255)   # orange (BGR)
            elif letter in selected_letters:
                color = (0, 0, 255)     # red — selected
            else:
                color = (0, 200, 0)     # green — empty
            cv2.circle(canvas, (x, y), r, color, 2)

    # Resize down for display
    H, W = canvas.shape[:2]
    if H > max_height:
        new_w = int(W * max_height / H)
        canvas = cv2.resize(canvas, (new_w, max_height))

    ok, png = cv2.imencode(".png", canvas)
    return png.tobytes() if ok else b""
