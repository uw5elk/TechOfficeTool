@echo off
chcp 65001 >nul
title Build DEBUG EXE - TechOfficeTool (UW5ELK)
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --console --name "TechOfficeTool_debug" --icon="icon.ico" ^
    --add-data "icon.ico;." ^
    --hidden-import techoffice_core ^
    --hidden-import win32api ^
    --hidden-import pywintypes ^
    --hidden-import pythoncom ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --collect-submodules win32com ^
    --collect-submodules send2trash ^
    DocxImageCompressor.py
echo.
echo Finished! EXE is here: %CD%\dist\TechOfficeTool_debug.exe
echo Log file: %LOCALAPPDATA%\TechOfficeTool\TechOfficeTool.log
pause
