@echo off
chcp 65001 >nul
title Setup - TechOfficeTool (UW5ELK)
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
echo Installing dependencies...
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo ERROR: could not install dependencies!
    pause
    exit /b 1
)
set "DESKTOP=%USERPROFILE%\Desktop"
if not exist "%DESKTOP%" set "DESKTOP=%USERPROFILE%\OneDrive\Desktop"
(
echo Set oWS = WScript.CreateObject("WScript.Shell")
echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\TechOfficeTool.lnk"
echo Set oLink = oWS.CreateShortcut(sLinkFile)
echo oLink.TargetPath = "pythonw.exe"
echo oLink.Arguments = """%~dp0DocxImageCompressor.py"""
echo oLink.WorkingDirectory = "%~dp0"
echo oLink.Save
) > "%TEMP%\create_shortcut.vbs"
cscript //nologo "%TEMP%\create_shortcut.vbs"
del "%TEMP%\create_shortcut.vbs" >nul 2>&1
echo Setup complete! Shortcut created on Desktop.
pause
