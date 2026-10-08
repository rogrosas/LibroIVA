@echo off
cd /d "%~dp0"
python -c "import openpyxl" 2>nul || (
    echo Instalando dependencias...
    python -m pip install -r requirements.txt
)
start "" pythonw main.py
