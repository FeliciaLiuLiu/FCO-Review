@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Dependency installation failed. Verify Python and network access.
    pause
    exit /b 1
)
python map_address_details.py
if errorlevel 1 (
    echo Mapping failed. See the error above.
) else (
    echo Done. Mapped workbook is in output.
)
pause
