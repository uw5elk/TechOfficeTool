@echo off
chcp 65001 >nul
title Build EXE - TechOfficeTool (UW5ELK)
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
echo [1/3] Installing dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: could not install dependencies!
    pause
    exit /b 1
)
echo [2/3] Installing pyinstaller...
python -m pip install pyinstaller
echo [3/3] Building EXE...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "TechOfficeTool" --icon="icon.ico" ^
    --add-data "icon.ico;." ^
    --hidden-import techoffice_core ^
    --hidden-import win32api ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import win32com.client.dynamic ^
    --collect-submodules win32com ^
    --collect-submodules send2trash ^
    DocxImageCompressor.py
echo.
echo Finished! EXE is here: %CD%\dist\TechOfficeTool.exe
pause
