"""Fiducial-marker detection and perspective normalization.

Detects the 4 corner squares on every OMR sheet, then warps the image
to a canonical reference frame so every bubble lands at a known coordinate.

Handles both grayscale and color images, and uses a PIL fallback for BMP
variants that OpenCV can't decode directly.
"""

from __future__ import annotations
from typing import Dict, Tuple

import io
import cv2
import numpy as np
from PIL import Image


FIDUCIAL_MARGIN = 40


def robust_decode(image_bytes: bytes) -> np.ndarray:
    """Decode bytes to a grayscale ndarray.

    OpenCV's imdecode fails on some 8-bit indexed-palette BMPs. PIL handles
    those fine, so we fall back automatically.
    """
    arr = np.frombuffer(image_bytes, dtype=np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_UNCHANGED)
    if img is None:
        pil = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        img = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
    if img.ndim == 3:
        if img.shape[2] == 4:
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
        else:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    if img.dtype != np.uint8:
        img = img.astype(np.uint8)
    return img


def _binarize_for_fiducials(gray: np.ndarray) -> np.ndarray:
    """Binary mask where the fiducial squares are white (255)."""
    _, bw = cv2.threshold(gray, 80, 255, cv2.THRESH_BINARY_INV)
    if bw.sum() / 255 < 200:
        _, bw = cv2.threshold(gray, 0, 255,
                              cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    return bw


def detect_fiducials(gray: np.ndarray) -> Dict[str, Tuple[float, float]]:
    """Find the 4 corner square markers.

    Returns {'TL', 'TR', 'BL', 'BR'} → centroid (x, y).
    """
    H, W = gray.shape[:2]
    bw = _binarize_for_fiducials(gray)
    n, _labels, stats, centroids = cv2.connectedComponentsWithStats(bw, 8)

    candidates = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if not (20 <= w <= 100 and 20 <= h <= 100):
            continue
        aspect = w / max(h, 1)
        if not (0.65 <= aspect <= 1.5):
            continue
        if area < 300:
            continue
        cx, cy = float(centroids[i][0]), float(centroids[i][1])
        candidates.append((cx, cy, w, h, area))

    if len(candidates) < 4:
        raise ValueError(
            f"Could not find 4 fiducial candidates "
            f"(found {len(candidates)} square shapes)."
        )

    corner_targets = {
        "TL": (0, 0), "TR": (W, 0), "BL": (0, H), "BR": (W, H),
    }
    fiducials: Dict[str, Tuple[float, float]] = {}
    remaining = list(candidates)
    for name, (tx, ty) in corner_targets.items():
        in_quadrant = [
            c for c in remaining
            if ((c[0] < W / 2) == (tx < W / 2))
            and ((c[1] < H / 2) == (ty < H / 2))
        ]
        pool = in_quadrant or remaining
        best = min(pool, key=lambda c: (c[0] - tx) ** 2 + (c[1] - ty) ** 2)
        dist = np.hypot(best[0] - tx, best[1] - ty)
        if dist > 0.20 * np.hypot(W, H):
            raise ValueError(
                f"No fiducial near {name} corner "
                f"(closest is {dist:.0f}px from the corner)."
            )
        fiducials[name] = (best[0], best[1])
        remaining.remove(best)

    # Sanity-check that the 4 detected corners form a roughly-rectangular
    # quadrilateral. Off-axis scans get heavily warped without this check.
    tl, tr = fiducials["TL"], fiducials["TR"]
    bl, br = fiducials["BL"], fiducials["BR"]
    top = np.hypot(tr[0] - tl[0], tr[1] - tl[1])
    bot = np.hypot(br[0] - bl[0], br[1] - bl[1])
    left = np.hypot(bl[0] - tl[0], bl[1] - tl[1])
    right = np.hypot(br[0] - tr[0], br[1] - tr[1])
    if max(top, bot) / max(min(top, bot), 1) > 1.30:
        raise ValueError(
            f"Fiducial quadrilateral isn't rectangular "
            f"(top={top:.0f}, bot={bot:.0f})."
        )
    if max(left, right) / max(min(left, right), 1) > 1.30:
        raise ValueError(
            f"Fiducial quadrilateral isn't rectangular "
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
    return cv2.warpPerspective(
        gray, M, (target_w, target_h),
        flags=cv2.INTER_LINEAR,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=255,
    )
