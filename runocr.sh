#!/bin/bash
# Cross-platform script to run OCR processing (API mode)

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
echo "  OCR Processing (API Typhoon)"
echo "========================================"
echo ""

# Run the OCR script with arguments
$PYTHON_CMD Extract_Inv.py "$@"

EXIT_CODE=$?

echo ""
echo "========================================"
if [ $EXIT_CODE -eq 0 ]; then
    echo "[SUCCESS] OCR process completed!"
else
    echo "[ERROR] OCR process failed with return code: $EXIT_CODE"
fi
echo "========================================"

exit $EXIT_CODE


