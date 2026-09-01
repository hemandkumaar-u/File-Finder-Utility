@echo off
title Word Image Matcher & Copier
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python Virtual Environment...
    python -m venv .venv
    echo Installing required packages...
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\pip.exe install -r requirements.txt
)

echo Starting Word Image Matcher GUI...
.venv\Scripts\python.exe main.py
pause
