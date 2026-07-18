#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

if ! python3 -c "import fastapi,uvicorn,openpyxl,docx,xlrd,multipart,bs4,pdfplumber" >/dev/null 2>&1; then
  python3 -m pip install -r requirements.txt
fi

python3 app.py
