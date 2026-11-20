@echo off
setlocal
title CHECK FS - Kiem tra Bao cao tai chinh

rem Luon chay tu thu muc chua file .bat de app.py va .streamlit\config.toml duoc nhan dung.
cd /d "%~dp0"

rem Uu tien moi truong ao cua du an; neu khong co thi dung Python trong PATH.
set "PYTHON_EXE="
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"
if not defined PYTHON_EXE if exist "venv\Scripts\python.exe" set "PYTHON_EXE=venv\Scripts\python.exe"
if not defined PYTHON_EXE (
    where python >nul 2>&1 && set "PYTHON_EXE=python"
)
if not defined PYTHON_EXE (
    where py >nul 2>&1 && set "PYTHON_EXE=py"
)
if not defined PYTHON_EXE (
    echo [LOI] Khong tim thay Python. Vui long cai Python 3 va thu lai.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo [LOI] Chua cai thu vien Streamlit.
    echo Chay lenh sau trong thu muc ung dung roi mo lai file nay:
    echo     "%PYTHON_EXE%" -m pip install -r requirements.txt
    pause
    exit /b 1
)

echo Dang khoi dong ung dung... Trinh duyet se tu mo sau vai giay.
echo Dong cua so nay de dung ung dung.
echo.
"%PYTHON_EXE%" -m streamlit run app.py
if errorlevel 1 (
    echo.
    echo [LOI] Ung dung dung bat thuong. Xem thong bao phia tren de biet nguyen nhan.
    pause
)
endlocal
