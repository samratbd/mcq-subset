# Building MCQ Studio for Windows

## The easy way — just double-click

In the `mcq_shuffler` folder, double-click:

    build_installer.bat

That's it. It installs everything needed and builds `dist\MCQ_Studio.exe`.
When done it opens the `dist\` folder automatically so you can see the file.

---

## What it does step by step

1. Installs Python packages (`pip install -r requirements.txt`)
2. Installs PyInstaller (`pip install pyinstaller`)
3. Runs `python -m PyInstaller --noconfirm mcq_studio.spec`
4. Produces `dist\MCQ_Studio.exe` — one self-contained file

---

## Requirements

- **Python 3.11 or 3.12** (recommended)
  - Download: https://www.python.org/downloads/
  - During install: tick **"Add Python to PATH"**
  - ⚠️ Python 3.14 (latest as of 2025) may not work with PyInstaller yet.
    If you have 3.14 and see errors, install Python 3.11 alongside it and
    use that version to build.

---

## If `pyinstaller` command is not found

This happens with some Python installs (the Scripts folder isn't on PATH).
The `build_installer.bat` already handles this by using
`python -m PyInstaller` instead of the bare `pyinstaller` command.

If you want to run it manually:

```powershell
# Instead of this (may fail):
pyinstaller --noconfirm mcq_studio.spec

# Use this (always works):
python -m PyInstaller --noconfirm mcq_studio.spec
```

---

## Output

After a successful build:

```
dist\
  MCQ_Studio.exe        ← the ONE file you need (80-120 MB)
  MCQ_Studio\           ← build artefact folder, ignore this
```

`MCQ_Studio.exe` is completely self-contained. Copy it to any
Windows 10/11 PC — Python does NOT need to be installed there.

---

## Optional: build a proper Windows installer

If you want a double-click installer with Start Menu shortcuts:

1. Install Inno Setup 6: https://jrsoftware.org/isdl.php
2. Open `installer\mcq_studio.iss` in Inno Setup Compiler
3. Press F9 (Compile)
4. Output: `installer\Output\MCQStudio-Setup-1.0.exe`

---

## Troubleshooting

| Error | Fix |
|---|---|
| `pyinstaller not recognized` | Use `python -m PyInstaller` (the .bat does this) |
| `cd path\to\mcq_shuffler` | That was just an example — stay in your actual folder |
| Python 3.14 build errors | Install Python 3.11 from python.org and use that |
| `ModuleNotFoundError` at runtime | Rebuild; the hidden imports list may need updating |
| Antivirus blocks the .exe | Click "More info → Run anyway" in Windows Defender |
| First launch is slow | Normal — unpacks ~80 MB to temp on first run (~5 sec) |
