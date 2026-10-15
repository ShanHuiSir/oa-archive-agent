@echo off
cd /d "%~dp0"
.venv\Scripts\python -m pip install -r requirements-dev.txt
.venv\Scripts\python -m PyInstaller --noconfirm --clean --name OAArchive --onedir --add-data "oa/web;oa/web" --collect-all rapidocr_onnxruntime --collect-all onnxruntime main.py
pause
