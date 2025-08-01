@echo off
SETLOCAL

REM Step 1: Create virtual environment if it doesn't exist
IF NOT EXIST venv (
    echo Creating virtual environment...
    python -m venv venv
)

REM Step 2: Activate the venv
call venv\Scripts\activate.bat

REM Step 3: Install dependencies (only needed once, but safe to re-run)
echo Installing dependencies...
pip install -r requirements.txt

REM Step 4: Run your main program
echo Running the invoice scanner...
python pdfplumber_test.py

pause

