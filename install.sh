#!/bin/bash
# Cross-platform installation script for OCR App

echo "========================================"
echo "Installing Python Libraries for OCR App"
echo "========================================"
echo ""

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Check if Python is available
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
    PIP_CMD="pip3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
    PIP_CMD="pip"
else
    echo "[ERROR] Python not found! Please install Python 3.9 or higher."
    exit 1
fi

# Check Python version
PYTHON_VERSION=$($PYTHON_CMD -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "Found Python version: $PYTHON_VERSION"

# Upgrade pip
echo ""
echo "[1/2] Upgrading pip..."
$PYTHON_CMD -m pip install --upgrade pip

# Install requirements
echo ""
echo "[2/2] Installing dependencies from requirements.txt..."
$PIP_CMD install -r requirements.txt

echo ""
echo "========================================"
echo "Python Installation Complete!"
echo "========================================"
echo ""

# Check OS and provide additional instructions
OS_TYPE=$(uname -s)

case "$OS_TYPE" in
    Linux*)
        echo "Detected OS: Linux"
        echo ""
        echo "Additional Setup for Linux:"
        echo "========================================"
        echo ""
        echo "1. Install Poppler (for PDF processing):"
        echo "   Ubuntu/Debian: sudo apt-get install poppler-utils"
        echo "   CentOS/RHEL:   sudo yum install poppler-utils"
        echo "   Fedora:        sudo dnf install poppler-utils"
        echo ""
        echo "2. Install Tesseract OCR (optional, for text highlighting):"
        echo "   Ubuntu/Debian: sudo apt-get install tesseract-ocr tesseract-ocr-tha"
        echo "   CentOS/RHEL:   sudo yum install tesseract tesseract-langpack-tha"
        echo ""
        echo "3. Install Ollama (optional, for local OCR):"
        echo "   curl -fsSL https://ollama.ai/install.sh | sh"
        echo "   ollama pull scb10x/typhoon-ocr1.5-3b:latest"
        ;;
    Darwin*)
        echo "Detected OS: macOS"
        echo ""
        echo "Additional Setup for macOS:"
        echo "========================================"
        echo ""
        echo "1. Install Homebrew (if not installed):"
        echo "   /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
        echo ""
        echo "2. Install Poppler (for PDF processing):"
        echo "   brew install poppler"
        echo ""
        echo "3. Install Tesseract OCR (optional, for text highlighting):"
        echo "   brew install tesseract"
        echo "   brew install tesseract-lang  # for Thai language support"
        echo ""
        echo "4. Install Ollama (optional, for local OCR):"
        echo "   brew install ollama"
        echo "   ollama pull scb10x/typhoon-ocr1.5-3b:latest"
        ;;
    MINGW*|CYGWIN*|MSYS*)
        echo "Detected OS: Windows (Git Bash/MSYS)"
        echo ""
        echo "Additional Setup for Windows:"
        echo "========================================"
        echo ""
        echo "1. Install Poppler:"
        echo "   Download: https://github.com/oschwartz10612/poppler-windows/releases/"
        echo "   Extract to: C:\\poppler\\"
        echo ""
        echo "2. Install Tesseract OCR:"
        echo "   Download: https://github.com/UB-Mannheim/tesseract/wiki"
        echo "   Install to: C:\\Program Files\\Tesseract-OCR\\"
        echo "   Select Thai language pack during installation!"
        echo ""
        echo "3. Install Ollama (optional):"
        echo "   Download: https://ollama.ai/"
        echo "   After install: ollama pull scb10x/typhoon-ocr1.5-3b:latest"
        ;;
    *)
        echo "Unknown OS: $OS_TYPE"
        echo "Please install poppler-utils and tesseract-ocr manually."
        ;;
esac

echo ""
echo "========================================"
echo "To start the application, run:"
echo "  ./run.sh  (Linux/macOS)"
echo "  runapp.bat  (Windows)"
echo "========================================"


