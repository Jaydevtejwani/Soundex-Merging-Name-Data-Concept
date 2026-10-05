@echo off
cd /d "%~dp0"
python soundex_name_matching.py
if errorlevel 1 (
    echo.
    echo The program ended with an error.
    pause
)
