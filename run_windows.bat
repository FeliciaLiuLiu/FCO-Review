@echo off
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
    echo Python launcher not found. Install Python from python.org with the launcher enabled.
    pause
    exit /b 1
)
py -3 -m pip install -r requirements.txt
if errorlevel 1 (
    echo Dependency installation failed. Check your network or contact IT.
    pause
    exit /b 1
)
py -3 convert.py
if errorlevel 1 (
    echo Conversion failed. See the error above.
) else (
    echo Done. Your Excel file is in the output folder.
)
pause
