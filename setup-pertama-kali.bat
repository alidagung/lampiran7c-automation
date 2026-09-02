@echo off
REM ============================================================
REM  SETUP PERTAMA KALI - Aplikasi Lampiran 7C
REM  Jalankan file ini SEKALI SAJA saat pertama kali memakai
REM  aplikasi (untuk menginstall komponen yang dibutuhkan).
REM ============================================================

echo ============================================================
echo   SETUP APLIKASI LAMPIRAN 7C
echo   Menginstall komponen yang dibutuhkan...
echo ============================================================
echo.

REM Pindah ke folder tempat file .bat ini berada
cd /d "%~dp0"

REM Cek apakah Python tersedia
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan!
    echo.
    echo Silakan install Python terlebih dahulu dari:
    echo   https://www.python.org/downloads/
    echo PENTING: centang "Add Python to PATH" saat menginstall.
    echo.
    pause
    exit /b 1
)

echo Python terdeteksi. Menginstall dependencies...
echo.
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Gagal menginstall dependencies.
    echo Periksa koneksi internet Anda, lalu coba jalankan lagi.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   SETUP SELESAI!
echo   Sekarang Anda bisa menjalankan aplikasi dengan
echo   double-click file: jalankan-aplikasi.bat
echo ============================================================
echo.
pause
