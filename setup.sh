#!/bin/bash

echo "======================================="
echo "Setting up Python environment (Mac/Linux)"
echo "======================================="

echo ""
echo "[1/3] Creating virtual environment (.venv)..."
# Check if python3 is available, otherwise use python
if command -v python3 &>/dev/null; then
    python3 -m venv .venv
else
    python -m venv .venv
fi

echo ""
echo "[2/3] Activating virtual environment..."
source .venv/bin/activate

echo ""
echo "[3/3] Installing necessary libraries..."
python -m pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "======================================="
echo "Setup Complete!"
echo "======================================="
echo "Remember to always activate your environment before running the tool:"
echo ""
echo "    source .venv/bin/activate"
echo ""
