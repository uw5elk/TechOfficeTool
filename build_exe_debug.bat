@echo off
chcp 65001 >nul
title Build DEBUG EXE - Word Compressor (UW5ELK)
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --console --name "TechOfficeTool_debug" --icon="icon.ico" ^
    --hidden-import win32api ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --collect-submodules win32com ^
    DocxImageCompressor.py
echo.
echo Finished! EXE is here: %CD%\dist\WordCompressor_debug.exe
pause
