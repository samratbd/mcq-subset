# Onurion OMR Studio

Optical Mark Recognition (OMR) scanner + MCQ paper shuffler. Two apps, one shared engine.

```
onurion-omr-studio/
├── app/            ← Shared core engine (OMR scanner, MCQ shuffler, parsers, writers)
├── web/            ← Flask web app  (cloud deployment, browser UI)
├── desktop/        ← Tkinter desktop app  (Windows/Mac/Linux, no limits)
└── docs/           ← Architecture notes, API docs
```

## Quick start

### Desktop app (recommended for large batches)

```bash
pip install -r desktop/requirements.txt
python desktop/mcq_studio.py
```

Or double-click `desktop/build_installer.bat` to produce a standalone `.exe`.

### Web app (local or cloud)

```bash
pip install -r web/requirements.txt
python web/server.py
# → open http://localhost:5000
```

### Deploy to DigitalOcean

```bash
cd web/
git push  # DigitalOcean App Platform auto-deploys from the Dockerfile
```

---

## Architecture

```
app/                    ← Python package, importable by both apps
│
├── omr/                ← OMR scanning engine
│   ├── fiducial.py     ← Detect the 4 corner squares, perspective-warp
│   ├── templates.py    ← 50-mark / 100-mark bubble grid coordinates
│   └── scanner.py      ← Sample bubbles, classify, render review images
│
├── parsers/            ← Question bank parsers
│   ├── csv_parser.py
│   ├── xlsx_parser.py
│   └── docx_parser.py
│
├── writers/            ← Output writers
│   ├── docx_writer.py  ← Word document (2-column Normal + Database layouts)
│   ├── pdf_writer.py   ← PDF via Word COM (Windows) or LibreOffice
│   ├── xlsx_writer.py
│   └── csv_writer.py
│
├── shuffler.py         ← Randomize question/option order, assign Set letters
├── models.py           ← Question dataclass
├── server.py           ← Flask routes (used by web/server.py)
└── math_utils.py       ← KaTeX → OMML conversion via pandoc
```

---

## Adding features

### Improve OMR accuracy
Edit `app/omr/templates.py` to adjust bubble coordinates, or `app/omr/scanner.py`
for the detection thresholds. Both apps pick up the change automatically.

### Add a new question format
Add a new parser in `app/parsers/` following the pattern of `csv_parser.py`,
then register it in `app/parsers/__init__.py`.

### Add an LMS API integration
Create `app/api/` with a connector (e.g. `moodle.py`, `google_classroom.py`).
The web app can expose it as a REST endpoint, the desktop can call it too.

```python
# Example future structure
app/
└── api/
    ├── __init__.py
    ├── moodle.py           ← POST results to Moodle gradebook
    ├── google_classroom.py ← Sync with Google Classroom
    └── rest.py             ← Generic REST API for any LMS
```

### Build another frontend
Any new frontend (mobile app, CLI tool, VS Code extension) just does:
```python
sys.path.insert(0, '/path/to/onurion-omr-studio')
from app.omr import scan_omr
from app.shuffler import make_set
```

---

## Developer info

**Developer:** S M Samrat  
**Company:** Onurion.com

---

## PDF output

| Platform | Engine | How to enable |
|---|---|---|
| Windows | Microsoft Word | `pip install docx2pdf` (Word must be installed) |
| Windows / Linux / macOS | LibreOffice | Install from https://www.libreoffice.org |

The app auto-detects which engine is available and uses the best one.

---

## OMR accuracy notes

- Templates calibrated by Hough Circle Transform on physical blank sheets
- Fiducial detection uses **outer corner** of each marker square (more stable than centroid)
- Snap-to-bubble: ±15 px search window corrects for scan registration errors
- Auto-detect: 50-mark (portrait) vs 100-mark (landscape) by aspect ratio
- Review flagging: ambiguous bubbles (fill between 25% and 55%) are flagged but still answered
- Typical accuracy: 100% on roll number, 99%+ on answers
