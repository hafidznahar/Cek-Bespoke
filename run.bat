@echo off
REM Klik 2x file ini untuk menjalankan aplikasi (Windows).
cd /d "%~dp0"
if not exist .venv (
    echo Pertama kali jalan: menyiapkan environment, mohon tunggu...
    python -m venv .venv
    call .venv\Scripts\activate
    python -m pip install --upgrade pip
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate
)
echo Membuka aplikasi di browser...
streamlit run app.py
pause
