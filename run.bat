@echo off
setlocal
cd /d "%~dp0"
title FormaFlow Local v0.6.0

where python >nul 2>&1
if not errorlevel 1 (
  set "PYTHON=python"
) else (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python was not found. Install Python 3.11 or newer, then run this file again.
    pause
    exit /b 1
  )
  set "PYTHON=py -3"
)

echo Checking required Python packages...
%PYTHON% -c "import fastapi,uvicorn,openpyxl,docx,xlrd,multipart,bs4,pdfplumber" >nul 2>&1
if errorlevel 1 (
  echo Installing required Python packages...
  %PYTHON% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo Installation failed. Check Python and internet access.
    pause
    exit /b 1
  )
)

echo Starting FormaFlow at http://127.0.0.1:8000
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:8000"
%PYTHON% app.py
pause
