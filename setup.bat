@echo off
echo =======================================
echo Setting up Python environment (Windows)
echo =======================================

echo.
echo [1/3] Creating virtual environment (.venv)...
python -m venv .venv

echo.
echo [2/3] Activating virtual environment...
call .venv\Scripts\activate

echo.
echo [3/3] Installing necessary libraries...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo =======================================
echo Setup Complete!
echo =======================================
echo Remember to always activate your environment before running the tool:
echo.
echo     .venv\Scripts\activate
echo.
pause
