# MCQ Shuffler

Local web app for generating shuffled question-paper sets from an MCQ source file.

## Features

- **Upload**: Word (.docx), Excel (.xlsx), or CSV (.csv).
- **Auto-detect** both Word layouts: *Normal* (2-col question table + answer sheet) and *Database* (8-col single table).
- **KaTeX-aware**: math like `$\dfrac{\mu_0 i}{2\pi r}$` round-trips through CSV/XLSX as KaTeX, and renders as real Word equations in .docx output (requires `pandoc`).
- **Seeded shuffle**: Set N is reproducible — same input + same set number = same output, every time.
- **Two shuffle modes**: question order only, or question order + options (with answer remapped automatically so it stays correct).
- **Up to 20 sets** generated at once, downloaded as a ZIP.
- **Optional SQLite persistence**: keep papers and regenerate sets later, or run one-shot.
- **Integrity checks**: every output is verified — the answer at the new position must equal the original correct option text, or the set is rejected and rebuilt.

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

- **CSV / XLSX**: KaTeX preserved verbatim (round-trips losslessly).
- **DOCX Normal** and **DOCX Database**: same shapes as input; math from KaTeX is rendered as real Word equations when pandoc is installed, otherwise as Unicode text.

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
