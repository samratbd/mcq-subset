@echo off
:: Onurion OMR Studio — one-click builder
:: Run this from the PROJECT ROOT folder (one level ABOVE desktop/)
:: Or double-click it from the desktop/ folder — the script handles both.
title Onurion OMR Studio Builder

echo ================================================
echo  Onurion OMR Studio — Windows Executable Builder
echo ================================================
echo.

:: Navigate to project root (parent of this file's folder)
pushd "%~dp0.."

:: Confirm we're in the right place
if not exist app\omr\scanner.py (
    echo ERROR: Cannot find the app/ folder.
    echo Make sure this .bat file is inside the desktop/ subfolder.
    popd
    pause
    exit /b 1
)

echo Project root: %cd%
echo.

echo [1/3] Installing dependencies...
python -m pip install -r desktop\requirements.txt --quiet
if errorlevel 1 ( echo ERROR: pip failed & pause & exit /b 1 )

echo [2/3] Installing PyInstaller...
python -m pip install "pyinstaller>=6.0" --quiet
if errorlevel 1 ( echo ERROR: PyInstaller failed & pause & exit /b 1 )

echo [3/3] Building Onurion_OMR_Studio.exe ...
echo       (takes 2-5 minutes)
echo.
python -m PyInstaller --noconfirm desktop\mcq_studio.spec

if errorlevel 1 (
    echo.
    echo BUILD FAILED. Common fixes:
    echo  - Python 3.14: install Python 3.11 from python.org instead
    echo  - Missing deps: python -m pip install -r desktop\requirements.txt
    popd & pause & exit /b 1
)

echo.
echo ================================================
echo  BUILD COMPLETE
echo ================================================
echo  File: dist\Onurion_OMR_Studio.exe
echo  This single .exe works on any Windows 10/11 PC.
echo ================================================
echo.
explorer dist
popd
pause
