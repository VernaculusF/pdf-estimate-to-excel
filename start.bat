@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "input" mkdir "input"
if not exist "output" mkdir "output"

if not exist "venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv venv
)

echo Installing dependencies...
"venv\Scripts\python.exe" -m pip install -r requirements.txt

echo.
echo Put PDF files into the input folder. Conversion is starting now.
echo Results will be saved into the output folder.
echo.

"venv\Scripts\python.exe" main.py --input input --output output

echo.
echo Done. Check the output folder.
pause
