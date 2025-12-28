@echo off
echo ========================================
echo Installing Python Libraries for OCR App
echo ========================================
echo.

REM Core libraries
echo [1/10] Installing Streamlit...
pip install streamlit

echo [2/10] Installing Pandas...
pip install pandas

echo [3/10] Installing OpenPyXL...
pip install openpyxl

echo [4/10] Installing Requests...
pip install requests

echo [5/10] Installing PyPDF...
pip install pypdf

echo [6/10] Installing pdf2image...
pip install pdf2image

echo [7/10] Installing Pillow...
pip install Pillow

echo [8/10] Installing pytesseract...
pip install pytesseract

echo [9/10] Installing PyMuPDF (fitz)...
pip install PyMuPDF

echo [10/10] Installing streamlit-pdf-viewer...
pip install streamlit-pdf-viewer

echo.
echo ========================================
echo Installation Complete!
echo ========================================
echo.
echo IMPORTANT: Additional Software Required
echo ========================================
echo.
echo 1. TESSERACT OCR (for accurate text positioning)
echo    Download: https://github.com/UB-Mannheim/tesseract/wiki
echo    Install to: C:\Program Files\Tesseract-OCR\
echo    Don't forget to select Thai language pack during installation!
echo.
echo 2. POPPLER (for PDF to image conversion)
echo    Download: https://github.com/oschwartz10612/poppler-windows/releases/
echo    Extract to: C:\poppler\
echo    Or update POPPLER_PATH in app.py
echo.
echo 3. OLLAMA (for Local Typhoon OCR - optional)
echo    Download: https://ollama.ai/
echo    After install, run: ollama pull typhoon-vision
echo.
echo ========================================
echo Press any key to exit...
pause >nul

