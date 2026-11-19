@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import openpyxl, reportlab, pypdf" 2>nul || (
  echo 必要なライブラリをインストールします...
  python -m pip install openpyxl reportlab pypdf certifi
)
start "" pythonw "%~dp0payroll_app.py"
