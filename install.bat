@echo off
REM Create a virtual environment named 'venv'
python -m venv venv

REM Activate the virtual environment
call venv\Scripts\activate

REM Install packages from requirements.txt
python -m pip install -r requirements.txt

REM Run the main.py script
python scripts\main.py

REM Deactivate the virtual environment
call venv\Scripts\deactivate

echo Installation and execution completed.
pause