@echo off
chcp 65001 >nul
title Setup - Word Compressor (UW5ELK)
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
echo Installing Pillow and pywin32...
python -m pip install pillow pywin32
set "DESKTOP=%USERPROFILE%\Desktop"
if not exist "%DESKTOP%" set "DESKTOP=%USERPROFILE%\OneDrive\Desktop"
(
echo Set oWS = WScript.CreateObject("WScript.Shell")
echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\Word Compressor.lnk"
echo Set oLink = oWS.CreateShortcut(sLinkFile)
echo oLink.TargetPath = "pythonw.exe"
echo oLink.Arguments = """%CD%\DocxImageCompressor.py"""
echo oLink.WorkingDirectory = "%CD%"
echo oLink.Save
) > "%TEMP%\create_shortcut.vbs"
cscript //nologo "%TEMP%\create_shortcut.vbs"
del "%TEMP%\create_shortcut.vbs" >nul 2>&1
echo Setup complete! Shortcut created on Desktop.
pause
