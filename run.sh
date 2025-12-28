#!/bin/bash
# Cross-platform script to run the Streamlit app

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Change to script directory
cd "$SCRIPT_DIR"

# Check if Python is available
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "[ERROR] Python not found! Please install Python 3.9 or higher."
    exit 1
fi

echo "========================================"
echo "  AI OCR & Document Editor"
echo "========================================"
echo ""
echo "Starting Streamlit app..."
echo ""

# Run the Streamlit app
$PYTHON_CMD -m streamlit run app.py


