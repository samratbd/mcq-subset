"""Fiducial-marker detection and perspective normalization.

Every OMR sheet in this project has four black square fiducial markers at
its four corners. We detect them, then warp the image so they sit at fixed
canonical positions. After warping, every bubble lives at a predictable
pixel coordinate regardless of how the sheet was scanned (tilt, slight
scale variation, modest perspective distortion).

Public API:
  - detect_fiducials(gray) → dict with TL, TR, BL, BR (each = (x, y) float)
  - warp_to_canonical(gray, fiducials, target_w, target_h) → warped image

The "canonical" image is plain grayscale (0..255), with low values = ink.
"""

from __future__ import annotations
from typing import Dict, List, Tuple, Optional

import cv2
import numpy as np


# Canonical frame margins from the page edge to the fiducial centroids.
# Tuned to match the fiducial positions in the user's sample sheets.
FIDUCIAL_MARGIN = 50  # pixels in canonical frame


def _binarize(gray: np.ndarray) -> np.ndarray:
    """Return an INVERTED binary image where ink = 255, paper = 0.

    Uses Otsu when the image isn't already bitonal; bitonal scans pass
    through a fixed threshold (which is fast and exact).
    """
    if len(np.unique(gray)) <= 2:
        # Already bitonal (1-bit BMP). Just invert.
        return cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY_INV)[1]
    # Otsu for variable-contrast scans
    _, bw = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU
    )
    return bw


def detect_fiducials(gray: np.ndarray) -> Dict[str, Tuple[float, float]]:
    """Detect the 4 corner squares of an OMR sheet.

    Returns {'TL', 'TR', 'BL', 'BR'} → (x, y) centroid in the original
    image's coordinate space.

    Strategy:
      1. Threshold to binary.
      2. Find all connected components.
      3. Filter to "square-ish" shapes (aspect within 0.7..1.4) of plausible
         size (15..80 px). These are bubble-or-fiducial-sized blobs.
      4. For each of the 4 image corners, pick the candidate closest to it.
      5. Sanity-check: the four chosen blobs must form a roughly-rectangular
         quadrilateral, otherwise raise an error.

    Raises ValueError if fiducials cannot be located reliably.
    """
    H, W = gray.shape[:2]
    bw = _binarize(gray)
    n, _labels, stats, centroids = cv2.connectedComponentsWithStats(bw, 8)

    # Build a list of candidate squares: roughly-square dark blobs, not tiny,
    # not enormous. Both fiducials and filled bubbles fit; we'll discriminate
    # by corner proximity in step 4.
    candidates: List[Tuple[float, float, int, int, int]] = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if w < 15 or h < 15 or w > 90 or h > 90:
            continue
        aspect = w / h
        if not (0.7 <= aspect <= 1.4):
            continue
        if area < 200:
            continue
        cx, cy = float(centroids[i][0]), float(centroids[i][1])
        candidates.append((cx, cy, w, h, area))

    if len(candidates) < 4:
        raise ValueError(
            f"Could not find enough fiducial candidates "
            f"(found {len(candidates)}, need ≥ 4)."
        )

    # For each corner, find the candidate whose centroid is closest to it
    # (within a generous radius so we don't accidentally pick a bubble).
    corner_targets = {
        "TL": (0, 0),
        "TR": (W, 0),
        "BL": (0, H),
        "BR": (W, H),
    }

    fiducials: Dict[str, Tuple[float, float]] = {}
    remaining = list(candidates)
    # Process corners in order; remove the chosen candidate so the same blob
    # can't double-count.
    for name, (tx, ty) in corner_targets.items():
        # Restrict the search to the candidate's actual quadrant —
        # otherwise on extreme skew the wrong blob can win for a far corner.
        in_quadrant = [
            c for c in remaining
            if ((c[0] < W / 2) == (tx < W / 2))
            and ((c[1] < H / 2) == (ty < H / 2))
        ]
        pool = in_quadrant or remaining
        best = min(pool, key=lambda c: (c[0] - tx) ** 2 + (c[1] - ty) ** 2)
        # Must actually be near the corner (within 12% of image diagonal)
        dist = np.hypot(best[0] - tx, best[1] - ty)
        if dist > 0.18 * np.hypot(W, H):
            raise ValueError(
                f"No fiducial near {name} corner "
                f"(closest blob is {dist:.0f}px away)."
            )
        fiducials[name] = (best[0], best[1])
        remaining.remove(best)

    # Sanity check the resulting quadrilateral:
    # opposite-side lengths must be similar (≤ 20% difference)
    tl, tr = fiducials["TL"], fiducials["TR"]
    bl, br = fiducials["BL"], fiducials["BR"]
    top = np.hypot(tr[0] - tl[0], tr[1] - tl[1])
    bot = np.hypot(br[0] - bl[0], br[1] - bl[1])
    left = np.hypot(bl[0] - tl[0], bl[1] - tl[1])
    right = np.hypot(br[0] - tr[0], br[1] - tr[1])
    if max(top, bot) / max(min(top, bot), 1) > 1.25:
        raise ValueError(
            f"Fiducial quadrilateral is not rectangular enough "
            f"(top={top:.0f}, bottom={bot:.0f})."
        )
    if max(left, right) / max(min(left, right), 1) > 1.25:
        raise ValueError(
            f"Fiducial quadrilateral is not rectangular enough "
            f"(left={left:.0f}, right={right:.0f})."
        )

    return fiducials


def warp_to_canonical(
    gray: np.ndarray,
    fiducials: Dict[str, Tuple[float, float]],
    target_w: int,
    target_h: int,
    margin: int = FIDUCIAL_MARGIN,
) -> np.ndarray:
    """Perspective-warp the image so the 4 fiducials sit at fixed positions.

    The canonical output has the fiducial centroids at:
        TL = (margin, margin)
        TR = (target_w - margin, margin)
        BL = (margin, target_h - margin)
        BR = (target_w - margin, target_h - margin)

    Any rotation, scale, or modest perspective distortion is removed.
    """
    src = np.float32([
        fiducials["TL"], fiducials["TR"],
        fiducials["BL"], fiducials["BR"],
    ])
    dst = np.float32([
        [margin, margin],
        [target_w - margin, margin],
        [margin, target_h - margin],
        [target_w - margin, target_h - margin],
    ])
    M = cv2.getPerspectiveTransform(src, dst)
    warped = cv2.warpPerspective(
        gray, M, (target_w, target_h),
        flags=cv2.INTER_LINEAR,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=255,
    )
    return warped
