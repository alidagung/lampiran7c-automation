@echo off
REM ============================================================
REM  JALANKAN APLIKASI LAMPIRAN 7C (web lokal)
REM  Double-click file ini untuk membuka aplikasi di browser.
REM  Aplikasi berjalan 100%% di komputer ini - data tidak
REM  dikirim ke internet.
REM ============================================================

REM Pindah ke folder tempat file .bat ini berada
cd /d "%~dp0"

REM Cek Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan!
    echo Jalankan "setup-pertama-kali.bat" terlebih dahulu,
    echo atau install Python dari https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo   APLIKASI LAMPIRAN 7C sedang dijalankan...
echo.
echo   Browser akan terbuka otomatis di:
echo       http://localhost:8501
echo.
echo   Jika browser tidak terbuka, buka alamat di atas
echo   secara manual.
echo.
echo   UNTUK MENGHENTIKAN aplikasi: tutup jendela ini
echo   atau tekan Ctrl + C.
echo ============================================================
echo.

python -m streamlit run app.py

REM Jika streamlit gagal jalan, tampilkan pesan
if errorlevel 1 (
    echo.
    echo [ERROR] Aplikasi gagal dijalankan.
    echo Pastikan Anda sudah menjalankan "setup-pertama-kali.bat".
    echo.
    pause
)
