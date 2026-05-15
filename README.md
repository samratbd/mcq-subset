# MCQ Shuffler

Local / self-hosted web app that turns one MCQ question paper into up to 20 reproducibly-shuffled sets, in Word, PDF, Excel, or CSV form.

## Features

- **Inputs**: Word (`.docx`), Excel (`.xlsx`), CSV (`.csv`).
  - Word: auto-detects *Normal* (2-col + answer sheet) or *Database* (8-col) layout.
  - CSV/XLSX: standard 8-column schema (see below).
- **Outputs**: Word (`Normal` / `Database`), PDF (`Normal` / `Database`), Excel, CSV.
  - PDF is rendered from the corresponding Word file via LibreOffice for perfect fidelity.
- **Math handling, both directions**:
  - KaTeX in CSV/Excel → real Word equations on output (via pandoc).
  - OMML equations in source Word → KaTeX `$...$` strings on Excel/CSV output.
  - Best-effort Unicode rendering (`x^2` → `x²`) as a middle ground.
- **Faithful Word layout**: 2-column page section, 2×2 option grid per question, compact paragraph spacing matching typical printed question papers. Answer sheet starts on a new page.
- **Question paper header (banner)**: include a default banner image (shipped in `static/assets/default_header.jpg` — replace the file to brand the app), upload a custom one per generation, or skip.
- **Seeded shuffle**: Set N is reproducible — same source + set number = same output, every time.
- **Two shuffle modes**: question order, and/or option order (with answer letter remapped so the correct option still wins).
- **Up to 20 sets per run**, downloaded as a ZIP with a `MANIFEST.txt` and per-set integrity report.
- **Integrity checks** before every download: question multiset preserved, correct option text preserved at the new answer position, SLs are 1..N. Failures abort rather than ship a wrong paper.
- **Sample / template files** downloadable from the UI for every supported input format.
- **Optional SQLite persistence**: keep papers and regenerate sets later, or run one-shot.

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
