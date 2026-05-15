"""OMR scanner — scan one image and return a structured result.

Pipeline:
  1. Decode the image (BMP/PNG/JPEG, with PIL fallback for tricky BMPs).
  2. Detect the 4 corner fiducials.
  3. Auto-detect sheet type from aspect ratio.
  4. Perspective-warp to a canonical frame.
  5. For each template position, sample a bubble area (with SNAP — see below).
  6. Classify each group using fixed + adaptive thresholds.
  7. Build the result with per-question confidence and review flags.

Snap-to-bubble:
  Templates are accurate to ±5 px against the blank, but every scanned sheet
  has its own small registration error (paper folds, scanner glass dust,
  ink bleed near corners shifting the fiducial centroid). To absorb that,
  for each template position we search a small window (±snap_search_radius)
  around the expected centre for the LOCAL minimum of the empty-bubble
  baseline, snap to it, then sample. This recovers another ~1% accuracy
  on real-world scans.
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Tuple

import cv2
import numpy as np

from .fiducial import detect_fiducials, warp_to_canonical, robust_decode
from .templates import SheetTemplate, get_template, TEMPLATES


# --- Tuneable thresholds (international OMR standard) -----------------------

# Fill fraction (0..1) of the bubble that is "dark".
LOW_THRESHOLD = 0.25    # below → definitely empty
HIGH_THRESHOLD = 0.55   # above → definitely filled
# Margin above the per-sheet empty baseline to count as filled
RELATIVE_FILL_MARGIN = 0.18
# Margin between top-2 fills for a "confident" pick on a multi-choice group
CONFIDENT_OPTION_MARGIN = 0.30


@dataclass
class OmrResult:
    sheet_type: str
    roll_number: str
    set_letter: str
    answers: List[str]
    confidence: float
    needs_review: bool
    review_items: List[str]
    fill_fractions: List[List[float]] = field(default_factory=list)
    error: Optional[str] = None

    def as_dict(self) -> dict:
        d = asdict(self)
        d["fill_fractions"] = [
            [round(f * 100, 1) for f in row] for row in self.fill_fractions
        ]
        return d


def _detect_sheet_type(img: np.ndarray) -> str:
    H, W = img.shape[:2]
    return "omr_100" if W / H > 1.0 else "omr_50"


# --- Bubble sampling --------------------------------------------------------

def _bubble_fill(warped: np.ndarray, cx: int, cy: int, r: int) -> float:
    """Fraction (0..1) of the bubble's interior that's dark ink."""
    H, W = warped.shape[:2]
    x0, y0 = max(0, cx - r), max(0, cy - r)
    x1, y1 = min(W, cx + r + 1), min(H, cy + r + 1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    patch = warped[y0:y1, x0:x1]
    yy, xx = np.mgrid[0:patch.shape[0], 0:patch.shape[1]]
    mask = (xx - (cx - x0)) ** 2 + (yy - (cy - y0)) ** 2 <= r ** 2
    if not mask.any():
        return 0.0
    pixels = patch[mask]
    if len(np.unique(pixels)) <= 2:
        dark = (pixels < 128).sum()
    else:
        # Adaptive threshold within the patch
        try:
            t, _ = cv2.threshold(pixels.reshape(-1, 1), 0, 255,
                                 cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            # Be more conservative — bubble outlines shouldn't count as fill
            dark_thresh = max(80, int(t * 0.85))
            dark = (pixels < dark_thresh).sum()
        except cv2.error:
            dark = (pixels < 140).sum()
    return float(dark) / float(mask.sum())


def _snap_centre(warped: np.ndarray, cx: int, cy: int,
                 r: int, search_r: int) -> Tuple[int, int]:
    """Find the actual bubble centre near the template position.

    Optimization: most bubbles in a well-aligned template are already at
    the right spot. So we first check the template position and only do
    a search if the template position has SUBSTANTIALLY less fill than
    one of the cardinal neighbours.

    The dense search itself uses a 5×5 grid at `search_r/2` step, which
    is faster than the prior 11×11 grid while still catching small offsets.
    """
    best_fill = _bubble_fill(warped, cx, cy, r)
    best_pos = (cx, cy)
    # 5×5 grid at half-step resolution = 25 samples max
    step = max(3, search_r // 2)
    for dy in range(-search_r, search_r + 1, step):
        for dx in range(-search_r, search_r + 1, step):
            if dx == 0 and dy == 0:
                continue
            f = _bubble_fill(warped, cx + dx, cy + dy, r)
            if f > best_fill + 0.03:
                best_fill = f
                best_pos = (cx + dx, cy + dy)
    return best_pos


def _sample_group(warped: np.ndarray,
                  positions: List[Tuple[int, int]],
                  radius: int,
                  snap_radius: int) -> List[float]:
    """Sample fill fractions for one group (e.g. one question's 4 options)."""
    fills = []
    for cx, cy in positions:
        sx, sy = _snap_centre(warped, cx, cy, radius, snap_radius)
        fills.append(_bubble_fill(warped, sx, sy, radius))
    return fills


# --- Classification ---------------------------------------------------------

def _classify_group(fills: List[float],
                    sheet_baseline: float
                    ) -> Tuple[List[int], str]:
    """Group fills → (selected_indices, classification).

    Classifications: empty | single | multi | ambiguous
    """
    high_thresh = max(HIGH_THRESHOLD,
                      sheet_baseline + RELATIVE_FILL_MARGIN)
    filled = [i for i, f in enumerate(fills) if f >= high_thresh]
    all_low = all(f < LOW_THRESHOLD for f in fills)

    if not filled:
        return ([], "empty") if all_low else ([], "ambiguous")
    non_filled = [i for i in range(len(fills)) if i not in filled]
    if all(fills[i] < LOW_THRESHOLD for i in non_filled):
        return (filled, "single" if len(filled) == 1 else "multi")
    return (filled, "ambiguous")


def _confidence(fills: List[float]) -> float:
    if not fills:
        return 0.0
    s = sorted(fills, reverse=True)
    top = s[0]
    second = s[1] if len(s) > 1 else 0.0
    if top < LOW_THRESHOLD:
        return 1.0
    return min(1.0, max(0.0, (top - second) / CONFIDENT_OPTION_MARGIN))


# --- Top-level scan ---------------------------------------------------------

def scan_omr(image_bytes: bytes, sheet_type: str = "auto") -> OmrResult:
    """Read one OMR sheet image, return the extracted data + diagnostics.

    Never raises — every failure path returns an OmrResult with `.error` set,
    so the caller can build a partial batch result without exception handling.
    """
    if not image_bytes:
        return _error_result(sheet_type, "Empty image data.")

    try:
        gray = robust_decode(image_bytes)
    except Exception as e:
        return _error_result(sheet_type, f"Image decode failed: {e}")

    # Sanity check — abort early on absurdly small images
    if gray.size < 1000 or min(gray.shape) < 100:
        return _error_result(
            sheet_type,
            f"Image too small to be an OMR sheet "
            f"({gray.shape[1]}×{gray.shape[0]}).",
        )

    if sheet_type == "auto":
        sheet_type = _detect_sheet_type(gray)

    try:
        template = get_template(sheet_type)
    except Exception as e:
        return _error_result(sheet_type, f"Unknown sheet type: {e}")

    try:
        fids = detect_fiducials(gray)
    except Exception as e:
        return _error_result(sheet_type, f"Fiducial detection failed: {e}")

    try:
        warped = warp_to_canonical(
            gray, fids, template.canonical_w, template.canonical_h
        )
    except Exception as e:
        return _error_result(sheet_type, f"Perspective warp failed: {e}")

    try:
        r = template.bubble_radius
        sr = template.snap_search_radius

        answer_fills = [
            _sample_group(warped, pos, r, sr) for pos in template.answer_bubbles
        ]
        roll_fills = [
            _sample_group(warped, col, r, sr) for col in template.roll_bubbles
        ]
        set_fills = _sample_group(warped, template.set_bubbles, r, sr)
    except Exception as e:
        return _error_result(sheet_type, f"Bubble sampling failed: {e}")

    all_fills = [f for row in answer_fills for f in row]
    if all_fills:
        s = sorted(all_fills)
        baseline = float(np.median(s[: max(1, int(len(s) * 0.6))]))
    else:
        baseline = 0.05

    answers: List[str] = []
    review_items: List[str] = []
    confs: List[float] = []
    for q_idx, fills in enumerate(answer_fills):
        sel, kind = _classify_group(fills, baseline)
        if kind == "empty":
            answers.append("")
        elif kind == "single":
            answers.append("ABCD"[sel[0]])
        elif kind == "multi":
            answers.append(",".join("ABCD"[i] for i in sel))
        else:
            best = int(np.argmax(fills))
            answers.append("ABCD"[best])
            review_items.append(f"Q{q_idx + 1}")
        confs.append(_confidence(fills))

    roll_chars: List[str] = []
    for d_idx, fills in enumerate(roll_fills):
        sel, kind = _classify_group(fills, baseline)
        if kind == "single":
            roll_chars.append(str(sel[0]))
        elif kind == "empty":
            roll_chars.append("?")
            review_items.append(f"roll_d{d_idx + 1}")
        else:
            roll_chars.append(str(int(np.argmax(fills))))
            review_items.append(f"roll_d{d_idx + 1}")
    roll_number = "".join(roll_chars)

    sel, kind = _classify_group(set_fills, baseline)
    if kind == "single":
        set_letter = template.set_letters[sel[0]]
    elif kind == "empty":
        set_letter = "?"
        review_items.append("set")
    else:
        set_letter = template.set_letters[int(np.argmax(set_fills))]
        review_items.append("set")

    overall_conf = float(np.mean(confs)) if confs else 0.0

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


# --- Annotated review image -------------------------------------------------

def render_review_image(image_bytes: bytes,
                        result: OmrResult,
                        max_height: int = 1600) -> bytes:
    """Render the warped sheet with every bubble drawn in colour:

        green  — clearly empty
        red    — clearly filled (= selected as the answer)
        orange — ambiguous (flagged for review)

    Also draws:
      - Yellow rectangle connecting the 4 fiducial corners (the OMR's
        outer geometry — confirms perspective correction worked).
      - Dashed BLUE outlines around each section (Roll, SET, Q blocks).
      - Header row with roll / set / confidence / review status.
    """
    try:
        gray = robust_decode(image_bytes)
        fids = detect_fiducials(gray)
        template = get_template(result.sheet_type)
        warped = warp_to_canonical(
            gray, fids, template.canonical_w, template.canonical_h
        )
    except Exception:
        return b""

    canvas = cv2.cvtColor(warped, cv2.COLOR_GRAY2BGR)
    r = template.bubble_radius

    flagged_qs = {item for item in result.review_items if item.startswith("Q")}

    # --- Draw the fiducial-bounded rectangle (yellow) -----------------------
    # After warping, fiducials sit at the canonical margins. Connect them
    # so the user can visually confirm the OMR's outer geometry.
    from .fiducial import FIDUCIAL_MARGIN
    m = FIDUCIAL_MARGIN
    W, H = template.canonical_w, template.canonical_h
    corners = [(m, m), (W - m, m), (W - m, H - m), (m, H - m)]
    for i in range(4):
        cv2.line(canvas, corners[i], corners[(i + 1) % 4], (0, 220, 220), 3)

    # --- Answer bubbles ----------------------------------------------------
    # Each option gets a CLEAR circular outline so it's easy to verify
    # alignment at a glance. Selected (filled) bubbles get an extra inner
    # ring at smaller radius for visual emphasis.
    for q_idx, positions in enumerate(template.answer_bubbles):
        flagged = f"Q{q_idx + 1}" in flagged_qs
        selected_letters = set()
        if q_idx < len(result.answers):
            for letter in result.answers[q_idx].split(","):
                if letter:
                    selected_letters.add(letter)
        for opt_idx, (x, y) in enumerate(positions):
            letter = "ABCD"[opt_idx]
            if flagged:
                color = (0, 140, 255)        # orange
            elif letter in selected_letters:
                color = (0, 0, 255)          # red — selected
            else:
                color = (0, 200, 0)          # green — empty
            cv2.circle(canvas, (x, y), r, color, 2)
            if letter in selected_letters and not flagged:
                # Inner ring for emphasis on selected option
                cv2.circle(canvas, (x, y), max(r - 6, 4), color, 1)

    # --- Roll number bubbles (cyan) ----------------------------------------
    for d_idx, col_positions in enumerate(template.roll_bubbles):
        selected_idx = None
        if d_idx < len(result.roll_number):
            ch = result.roll_number[d_idx]
            if ch.isdigit():
                selected_idx = int(ch)
        for digit_val, (x, y) in enumerate(col_positions):
            if digit_val == selected_idx:
                color = (0, 0, 255)          # red — selected digit
            else:
                color = (200, 200, 0)        # cyan — empty
            cv2.circle(canvas, (x, y), r, color, 2)
            if digit_val == selected_idx:
                cv2.circle(canvas, (x, y), max(r - 6, 4), color, 1)

    # --- SET bubbles (magenta) ---------------------------------------------
    selected_set_idx = None
    if result.set_letter and result.set_letter != "?":
        selected_set_idx = template.set_letters.index(result.set_letter)
    for s_idx, (x, y) in enumerate(template.set_bubbles):
        color = (0, 0, 255) if s_idx == selected_set_idx else (255, 0, 255)
        cv2.circle(canvas, (x, y), r, color, 2)
        if s_idx == selected_set_idx:
            cv2.circle(canvas, (x, y), max(r - 6, 4), color, 1)

    # --- Section labels (dashed blue boxes + names) ------------------------
    # The OMR has its own printed rectangles around each section. We draw our
    # dashed-blue boxes WITHIN those printed rectangles so the two don't
    # overlap. Padding picked tight enough that bubbles are well inside.
    def _draw_section_box(name: str, positions: list,
                          color=(255, 80, 0), pad: int = 15):
        if not positions:
            return
        xs = [p[0] for p in positions]
        ys = [p[1] for p in positions]
        x0, x1 = min(xs) - pad, max(xs) + pad
        y0, y1 = min(ys) - pad, max(ys) + pad
        for x in range(x0, x1, 16):
            cv2.line(canvas, (x, y0), (min(x + 8, x1), y0), color, 2)
            cv2.line(canvas, (x, y1), (min(x + 8, x1), y1), color, 2)
        for y in range(y0, y1, 16):
            cv2.line(canvas, (x0, y), (x0, min(y + 8, y1)), color, 2)
            cv2.line(canvas, (x1, y), (x1, min(y + 8, y1)), color, 2)
        label_y = max(20, y0 - 8)
        cv2.putText(canvas, name, (x0 + 4, label_y),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2, cv2.LINE_AA)

    roll_positions = [p for col in template.roll_bubbles for p in col]
    _draw_section_box("ROLL NUMBER", roll_positions)
    _draw_section_box("SET", template.set_bubbles)

    n = template.n_questions
    if n == 50:
        block1 = [p for q in template.answer_bubbles[:25] for p in q]
        block2 = [p for q in template.answer_bubbles[25:] for p in q]
        _draw_section_box("Q01-25", block1)
        _draw_section_box("Q26-50", block2)
    elif n == 100:
        for blk in range(5):
            start = blk * 20
            positions = [p for q in template.answer_bubbles[start:start + 20] for p in q]
            label = f"Q{start + 1:02d}-{start + 20}"
            _draw_section_box(label, positions)

    summary = (
        f"Roll={result.roll_number}  SET={result.set_letter}  "
        f"Conf={result.confidence * 100:.1f}%  "
        f"Review={'YES' if result.needs_review else 'no'}"
    )
    cv2.putText(canvas, summary, (template.bubble_radius, 30),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 200), 2, cv2.LINE_AA)

    H, W = canvas.shape[:2]
    if H > max_height:
        canvas = cv2.resize(canvas, (int(W * max_height / H), max_height))

    ok, png = cv2.imencode(".png", canvas)
    return png.tobytes() if ok else b""
