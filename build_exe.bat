@echo off
chcp 65001 >nul
title Build EXE - TechOfficeTool (UW5ELK)
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
echo [1/4] Checking Pillow...
python -m pip install pillow
echo [2/4] Checking pywin32...
python -m pip install pywin32
echo [3/4] Installing pyinstaller...
python -m pip install pyinstaller
echo [4/4] Building EXE...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "TechOfficeTool" --icon="icon.ico" ^
    --hidden-import win32api ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import win32com.client.dynamic ^
    --collect-submodules win32com ^
    DocxImageCompressor.py
echo.
echo Finished! EXE is here: %CD%\dist\TechOfficeTool.exe
pause
