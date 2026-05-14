# MCQ Shuffler

Local web app for generating shuffled question-paper sets from an MCQ source file.

## Features

- **Upload**: Word (.docx), Excel (.xlsx), or CSV (.csv).
- **Auto-detect** both Word layouts: *Normal* (2-col question table + answer sheet) and *Database* (8-col single table).
- **Math handling — both directions**:
  - **KaTeX → Word**: `$\dfrac{\mu_0 i}{2\pi r}$` in CSV/Excel becomes a real Word equation in `.docx`.
  - **Word → KaTeX**: OMML equations in a source Word doc are converted *back* to `$...$` strings when exporting to CSV/Excel.
  - **Best-effort Unicode**: e.g. `$x^2$` → `x²` for cleaner reading in either direction.
- **Per-output math toggle** in the UI: equation / unicode / verbatim text for Word; KaTeX / Unicode for CSV/Excel.
- **Faithful Word layout**: output uses a 2-column page section with a 2×2 option grid per question, matching the look of typical source papers.
- **Seeded shuffle**: Set N is reproducible — same input + same set number = same output, every time.
- **Two shuffle modes**: question order only, or question order + options (with answer letter remapped so the new letter still points to the correct option).
- **Up to 20 sets** generated at once, downloaded as a ZIP with a `MANIFEST.txt` recording every choice and a per-set integrity report.
- **Optional SQLite persistence**: keep papers and regenerate sets later, or run one-shot.
- **Integrity checks**: every output is verified — the option text at the new answer position must equal the original correct option text, the multiset of all options must match the source, and SLs must be 1..N. Failures abort generation rather than ship a wrong paper.

## Setup

```bash
# 1. Python 3.9+ recommended
pip install -r requirements.txt

# 2. (Optional but recommended) install pandoc for native Word math
#    Ubuntu/Debian:  sudo apt install pandoc
#    macOS:          brew install pandoc
#    Windows:        https://pandoc.org/installing.html

# 3. Run
python run.py
```

The app opens automatically in your browser at <http://localhost:5000>.

## Deployment

The repo includes everything needed for one-click deployment on DigitalOcean App Platform, Heroku, or any buildpack-based PaaS:

- **`Procfile`** — start command (`gunicorn`).
- **`.python-version`** — pins Python to 3.12.
- **`requirements.txt`** — includes `gunicorn`.
- **`Aptfile`** — installs `pandoc` (requires the apt buildpack to be enabled on your platform).
- **`Dockerfile`** + **`.dockerignore`** — alternative if you prefer container deploys; bundles pandoc automatically and is the most reliable path.

**Recommended for DigitalOcean**: use the Dockerfile path. Create the App, point it at your repo, set "Resource Type → Dockerfile", and the build will produce an image with pandoc preinstalled. App Platform sets `$PORT` automatically; the container reads it.

**Buildpack path (no pandoc)**: if you use the default Python buildpack without the apt buildpack, the app still works — KaTeX in CSV/Excel passes through as text, but Word output won't render `$...$` as native equations and Word source OMML won't be converted to KaTeX. Everything else (shuffling, integrity, layout, downloads) is identical.

## Input formats

### CSV / XLSX
Header row (case-insensitive, BOM tolerated):
```
title, type, option1, option2, option3, option4, answer, explanation
```
- `answer` is `1`, `2`, `3`, or `4` (the index of the correct option).
- `type` is informational (always `MCQ` in practice).
- Math expressions use KaTeX delimited by `$...$` (e.g. `$\dfrac{1}{2}mv^2$`).

### Word (.docx)
Two layouts, auto-detected:

**Normal layout** — 2 tables:
1. Question table: 2 cols, one row per question. Col 0 is SL (`01.`); col 1 contains a nested table with question text and `A./B./C./D.` options.
2. Answer sheet: 3 cols — `Q No. | Ans | Explanation`.

**Database layout** — 1 table, 8 cols:
`(blank) | Question | Opt A | Opt B | Opt C | Opt D | (blank) | "Letter; Explanation"`

## Output

The UI exposes two math-rendering toggles, automatically shown only when relevant:

**Word output (`Word — Normal` or `Word — Database`)**
- `equation` (default): KaTeX → real Word equations. Needs pandoc.
- `unicode`: best-effort plain text (e.g. `x^2` → `x²`). Complex expressions like `\dfrac{a}{b}` stay as `$...$` if pandoc can't simplify them.
- `text`: keep `$...$` source verbatim — no rendering at all.

**Excel / CSV output**
- `katex` (default): preserve `$...$` so the file round-trips losslessly into Word later.
- `unicode`: best-effort plain text. Useful when you want a "clean" spreadsheet to read directly.

**Layout fidelity**:
- *DOCX Normal*: 2-column page section, with each question shown as `NN.` heading followed by a small 2×2 option table (`A. | B.` / `C. | D.`). The answer sheet at the end is single-column.
- *DOCX Database*: single 8-column table.
- *CSV / XLSX*: 8-column schema, identical to the input format.

The downloadable ZIP also contains a `MANIFEST.txt` summarising the run and a one-line integrity report per set (`Set 03: OK (40 Qs)`).

## Project layout

```
mcq_shuffler/
├── run.py                     entry point
├── app/
│   ├── server.py              Flask routes
│   ├── models.py              Question dataclass
│   ├── db.py                  SQLite persistence (optional)
│   ├── shuffler.py            seeded shuffle + answer remap
│   ├── math_utils.py          KaTeX → OMML (pandoc) + helpers
│   ├── parsers/{docx,xlsx,csv}_parser.py
│   └── writers/{docx,xlsx,csv}_writer.py
├── templates/index.html
└── static/{style.css, app.js}
```
