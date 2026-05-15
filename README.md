# MCQ Shuffler

Local / self-hosted web app that turns one MCQ question paper into up to 20 reproducibly-shuffled sets, in Word, PDF, Excel, or CSV form.

## Quick start

### Run locally (recommended for performance)

The web app runs entirely on your machine. The browser → Flask backend
communication is local, so it's much faster than going through a remote
server.

**System requirements (one-time setup):**

OpenCV (used by the OMR scanner) needs two system libraries on Linux/macOS
that aren't bundled with the Python package. Without them you'll get an
HTTP 500 error and the server log will say `libGL.so.1: cannot open
shared object file`. Install:

```bash
# Linux (Debian/Ubuntu)
sudo apt update
sudo apt install -y libgl1 libglib2.0-0 pandoc libreoffice-writer fonts-noto-core

# macOS
brew install pkg-config

# Windows
# Nothing to install at the system level — all required libs ship inside
# the Python wheels.
```

**Then:**

```bash
git clone <your-repo-url>
cd mcq_shuffler
pip install -r requirements.txt
python run.py            # serves at http://127.0.0.1:5000
```

If `python run.py` works without errors, you can also try `gunicorn -w 2
app.server:create_app\(\)` for production-like serving.

### Run in Docker

```bash
docker build -t mcq-shuffler .
docker run -p 8000:8000 mcq-shuffler
# open http://localhost:8000
```

The Dockerfile already installs every system dep — no extra setup needed.

### Deploy to DigitalOcean / Heroku

Push the repo. Both platforms read the `Dockerfile`; DigitalOcean App
Platform also reads `Procfile` and `Aptfile`. No environment variables
are required.

## Troubleshooting "HTTP error" when running locally

The most common cause is missing OpenCV system libraries. Check the
server console (the terminal where you ran `python run.py`) — if it
mentions `libGL.so.1` or `cv2`, install the system libs above.

Other common causes:

| Symptom | Likely cause | Fix |
| --- | --- | --- |
| `ModuleNotFoundError: No module named 'cv2'` | OpenCV not installed | `pip install -r requirements.txt` |
| `OSError: libGL.so.1: cannot open shared object file` | Missing system lib | `sudo apt install libgl1 libglib2.0-0` |
| `ModuleNotFoundError: No module named 'docx'` | python-docx not installed | `pip install python-docx` |
| Port 5000 already in use | Another app on the port | `python run.py --port 8000` (or kill the other) |
| PDF generation hangs forever | LibreOffice not installed | `sudo apt install libreoffice-writer` |
| Word equations show as `$...$` text | pandoc not installed | `sudo apt install pandoc` |

## Features

The app has two tabs:

### 📝 Generate question paper sets (the original feature)

- **Inputs**: Word (`.docx`), Excel (`.xlsx`), CSV (`.csv`).
  - Word: auto-detects *Normal* (2-col + answer sheet) or *Database* (8-col) layout.
  - CSV/XLSX: standard 8-column schema (see below).
- **Outputs**: Word (`Normal` / `Database`), PDF (`Normal` / `Database`), Excel, CSV.
- **Math both directions**: KaTeX ↔ real Word equations via pandoc.
- **Faithful Word layout**: 2-column page, 2×2 option grid per question, answer sheet on a new page.
- **Question paper header (banner)**: include a default banner, upload custom, or none.
- **Seeded shuffle**: Set N is reproducible — same source + set number = same output.
- **Two shuffle modes**: question order, and/or option order (with answer letter remapped).
- **Up to 20 sets per run**, downloaded as a ZIP with manifest + integrity report.
- **Sample / template files** downloadable from the UI for every input format.
- **Optional SQLite persistence**.

### 📋 Scan OMR answer sheets

Reads filled OMR sheets and extracts roll number, set letter, and per-question
answers to a spreadsheet. Tested at **~100% accuracy on the 50-question sheet
format** (24/24 samples in the included batch read correctly, with the one
"miss" being a 100-question sheet mistakenly run against the 50-question
template).

**How it works:**

1. **Fiducial-based geometric correction** — The 4 black corner squares are
   detected on every sheet, then a perspective transform warps the image to a
   canonical reference frame. Any tilt, scale variation, or modest perspective
   distortion is removed before bubble sampling.
2. **Pixel-precise template** — Bubble positions are measured directly from
   the user's blank template (`MCQ050202605150001.bmp` and
   `MCQ100202605150001.bmp`) to ±5 px accuracy.
3. **Snap-to-bubble** — At scan time, each template position is refined by
   searching a small window for the local fill maximum. This absorbs the ~5 px
   residual error from imperfect scans.
4. **Adaptive thresholding** — Each sheet's own empty-bubble baseline is
   computed, so the scanner adapts to variations in scan density, paper
   colour, and ink saturation.

**Output columns:** `serial, roll_number, set, Q1, Q2, …, QN, confidence,
needs_review, review_items, source_file`

Ambiguous cells are flagged in `review_items` (e.g. `Q3,Q17,roll_d2`) and the
whole row gets a light-orange highlight in Excel output. This makes the ~1%
of marks that can't be decided automatically easy to spot-check by hand.

**Sheet-type support:**

- **50-question** (portrait, 1392 × 2078 source) — **production-ready**
- **100-question** (landscape, 2482 × 1636 source) — *bubble positions
  approximate; will be refined further with more sample data*

**To improve accuracy further:**

The two template files in `app/omr/templates.py` (`_build_50q` and
`_build_100q`) carry the exact bubble coordinates. If your printed sheet
layout differs (different printer, different paper), tweak those numbers
and re-run.

## Setup

```bash
# 1. Python 3.12 recommended
pip install -r requirements.txt

# 2. Install system tools (optional but recommended)
#    pandoc      → KaTeX ↔ Word equation conversion
#    libreoffice → PDF output
#
#    Ubuntu/Debian:  sudo apt install pandoc libreoffice-writer fonts-noto-core
#    macOS:          brew install pandoc libreoffice
#    Windows:        install pandoc + LibreOffice from their websites

# 3. Run
python run.py
```

Opens automatically at <http://localhost:5000>.

## Deployment

Repo includes everything needed for one-click deployment:

| File | Purpose |
|---|---|
| `Procfile` | start command (`gunicorn`) |
| `.python-version` | pins Python to 3.12 |
| `requirements.txt` | adds `gunicorn` |
| `Aptfile` | installs pandoc + libreoffice + fonts (needs the apt buildpack) |
| `Dockerfile` + `.dockerignore` | container build with everything pre-installed |

**Recommended on DigitalOcean App Platform: use the Dockerfile.** Set the component's "Resource Type" to "Dockerfile" instead of "Buildpack". The resulting image has Python 3.12, pandoc, LibreOffice, and Bengali-capable fonts pre-installed, so every output format works out of the box.

Without LibreOffice the app still works — PDF output returns a clear error message and the user can fall back to Word.

## Input format

### CSV / XLSX
Header row (case-insensitive, BOM tolerated):
```
title, type, option1, option2, option3, option4, answer, explanation
```
- `answer` is `1`, `2`, `3`, or `4`.
- `type` is informational; we treat everything as MCQ.
- Math: inline KaTeX delimited by `$...$` (e.g. `$\dfrac{1}{2}mv^2$`).

### Word (.docx)
**Normal**: 2 tables. Question table (2 cols × N rows: SL | nested options table) + Answer sheet (3 cols: SL | letter | explanation).

**Database**: 1 table, 8 columns:
```
SL | Question | OptA | OptB | OptC | OptD | Answer | Explanation
```
The "Answer" column is preferred; if it's empty we fall back to a legacy
`"Letter; Explanation"` pattern in column 8 (which is what older source
files use).

Download the samples from the UI to see the exact structure.

## Output options

The UI shows only the controls relevant to the chosen format:

**Word / PDF**
- *Math rendering*: `equation` (default, real Word equations — needs pandoc), `unicode` (best-effort plain text), `text` (keep `$...$` verbatim).
- *Question paper header*: none, default banner (from `static/assets/default_header.jpg`), or upload a custom image.

**Excel / CSV**
- *Math format*: `katex` (default, lossless round-trip) or `unicode` (best-effort plain text).

The downloaded ZIP also contains a `MANIFEST.txt` summarising every choice plus a one-line integrity report per set (`Set 03: OK (40 Qs)`).

## Project layout

```
mcq_shuffler/
├── run.py                       entry point (Flask dev server + browser open)
├── Procfile                     production start command for gunicorn
├── .python-version              3.12
├── Dockerfile                   recommended deploy target
├── Aptfile                      buildpack-path system deps
├── requirements.txt
├── README.md
├── app/
│   ├── server.py                Flask routes
│   ├── models.py                Question dataclass
│   ├── db.py                    SQLite persistence (optional)
│   ├── shuffler.py              seeded shuffle + answer remap
│   ├── samples.py               on-demand template files
│   ├── math_utils.py            KaTeX ↔ OMML helpers (pandoc)
│   ├── parsers/                 docx_parser, xlsx_parser, csv_parser
│   └── writers/                 docx_writer, xlsx_writer, csv_writer, pdf_writer
├── templates/index.html
└── static/
    ├── style.css
    ├── app.js
    └── assets/
        └── default_header.jpg   replace to brand the app project-wide
```
