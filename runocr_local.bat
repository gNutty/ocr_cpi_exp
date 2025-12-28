@echo off
echo ========================================
echo   OCR Local Processing (Typhoon)
echo ========================================
echo.

REM ตรวจสอบว่า Python ติดตั้งแล้วหรือยัง
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found! Please install Python first.
    echo.
    pause
    exit /b 1
)

echo Starting OCR process...
echo.

REM รัน Python script
python Extract_Inv_local.py %1 %2 %3

REM เก็บ exit code
set EXIT_CODE=%ERRORLEVEL%

echo.
echo ========================================
if %EXIT_CODE% EQU 0 (
    echo [SUCCESS] OCR process completed!
) else (
    echo [ERROR] OCR process failed with return code: %EXIT_CODE%
    echo Check output above for details
)
echo ========================================
echo.

REM ไม่ pause เมื่อรันผ่าน subprocess (app.py)
REM ถ้าต้องการ pause เมื่อรันโดยตรง ให้ uncomment บรรทัดด้านล่าง
REM pause

REM ส่ง exit code กลับ
exit /b %EXIT_CODE%
