# OMR Scanner — Desktop application

A standalone GUI to scan OMR sheets locally — no web server, no upload size
limits, no time-outs. Built for batches of 50 / 100 / 1000+ sheets.

![OMR Scanner Desktop GUI](omr_scanner.png)

## Features

* **Batch processing** — pick a folder, hit Start, walk away
* **Live progress** — % done, sheets/sec, ETA, per-sheet result log
* **Parallel** — uses all CPU cores (4-8× faster than the web app)
* **Output formats** — Excel (.xlsx), CSV, or JSON
* **Annotated review images** — see exactly which bubble was detected for each option
* **Auto-detect** sheet type (50-mark vs 100-mark) or force one
* **Stop** in the middle of a long run — partial results preserved

## Quick start (any OS)

```bash
# 1. unzip the project
unzip mcq_shuffler.zip
cd mcq_shuffler

# 2. install dependencies
pip install -r requirements.txt

# 3. run the desktop app
python run_desktop.py
```

`run_desktop.py` checks that everything's installed and prints
platform-specific hints if anything is missing (especially Tkinter on Linux,
where it sometimes needs to be installed separately).

## Building a standalone .exe / .app / Linux binary

For users who shouldn't have to install Python:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed \
    --add-data "app:app" \
    --name "OMR Scanner" \
    desktop_omr.py
```

The result lands in `dist/`:

* Windows: `dist/OMR Scanner.exe`
* macOS:   `dist/OMR Scanner.app`
* Linux:   `dist/OMR Scanner`

Distribute that single file — recipients don't need Python or pip.

## Using the GUI

1. **Input folder** — folder containing your scanned OMR sheets (.bmp,
   .png, .jpg, .jpeg, .tif, .tiff, .gif).
2. **Output folder** — where results and review images get written.
   Defaults to `<input>/results_<timestamp>/`.
3. **Sheet type** — leave on "Auto-detect" unless your folder is mixed.
4. **Output format** — Excel for easy review; CSV for piping into other
   tools; JSON for programmatic consumption.
5. **Save annotated review images** — useful for verifying alignment.
   Adds a `review/` subfolder with one PNG per sheet.
6. **Search subfolders recursively** — turn on if your sheets are
   organised into per-section folders.

Click **Start Scanning**. The progress bar advances live; the log
streams each sheet's result as soon as it's done. Click **Stop** to
abort early — output is still written for the sheets already done.

## Output files

Every run produces:

```
output_folder/
├── omr_results.xlsx     (or .csv / .json depending on choice)
├── SUMMARY.txt           (counts, average confidence, per-sheet summary)
└── review/               (only if "Save review images" is on)
    ├── 160201MCQ_…_review.png
    ├── 160202MCQ_…_review.png
    └── ...
```

The Excel file colour-codes rows: **orange** = needs review, **red** =
failed scan. Each sheet contributes one row with `roll`, `set`, `Q1` …
`Qn`, `confidence`, `needs_review`, `review_items`, and the source
filename.

## Speed

Measured on a 4-core laptop with mixed 50-mark and 100-mark scans:

| Configuration             | Throughput        |
|---------------------------|-------------------|
| Single-threaded (web app) | ~3-5 sheets/sec   |
| Desktop (4 workers)       | ~15-20 sheets/sec |
| Desktop (8 workers)       | ~25-35 sheets/sec |

So **500 sheets ≈ 30 seconds** on a modern laptop.

## Troubleshooting

* **"No module named tkinter"** (Linux): `sudo apt install python3-tk`
* **"libGL.so.1: cannot open shared object"** (Linux):
  `sudo apt install libgl1 libglib2.0-0`
* **All sheets show "review=YES"**: the input folder probably contains
  a mix of 50-mark and 100-mark sheets but you forced one of them.
  Switch to "Auto-detect".
* **Slow on large batches**: check Task Manager / `htop` — if Python
  isn't pinning all cores, your system probably doesn't have enough
  RAM. Each worker uses ~150 MB. Reduce `max_workers` in
  `desktop_omr.py` if needed.
