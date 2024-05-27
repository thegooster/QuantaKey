@echo off
REM Activate the virtual environment
call venv\Scripts\activate

REM Run the main.py script
python \scripts\main.py

REM Deactivate the virtual environment
call venv\Scripts\deactivate

echo Execution completed.
pause