#!/bin/bash
# Cross-platform script to run Local OCR processing (Ollama)

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
echo "  OCR Local Processing (Typhoon/Ollama)"
echo "========================================"
echo ""

# Check if Ollama is running
if ! curl -s http://localhost:11434/api/tags > /dev/null 2>&1; then
    echo "[WARNING] Ollama might not be running. Starting check..."
    echo "Please ensure Ollama is installed and running."
    echo "Install: https://ollama.ai/"
    echo "Run: ollama serve"
    echo ""
fi

echo "Starting OCR process..."
echo ""

# Run the OCR script with arguments
$PYTHON_CMD Extract_Inv_local.py "$@"

EXIT_CODE=$?

echo ""
echo "========================================"
if [ $EXIT_CODE -eq 0 ]; then
    echo "[SUCCESS] OCR process completed!"
else
    echo "[ERROR] OCR process failed with return code: $EXIT_CODE"
    echo "Check output above for details"
fi
echo "========================================"

exit $EXIT_CODE


