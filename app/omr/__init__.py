"""OMR (Optical Mark Recognition) scanner.

Scans OMR answer sheets (50 or 100 question variants), extracting roll
number, set letter, and per-question answers. Output formats: CSV, XLSX,
JSON.

Public entry points:
  - scan_omr(image_bytes, sheet_type='auto') → OmrResult
  - render_review_image(image_bytes, result) → annotated PNG bytes
  - write_csv / write_xlsx / write_json: format a batch of results
  - TEMPLATES: dictionary of known sheet types
"""

from .scanner import scan_omr, render_review_image, scan_and_render, OmrResult
from .output import write_csv, write_xlsx, write_json
from .templates import TEMPLATES, get_template, SheetTemplate

__all__ = [
    "scan_omr",
    "scan_and_render",
    "render_review_image",
    "OmrResult",
    "write_csv", "write_xlsx", "write_json",
    "TEMPLATES", "get_template", "SheetTemplate",
]
