@echo off
cd /d "%~dp0"
if exist "dist\EditorInformativos.exe" (
    start "" "dist\EditorInformativos.exe"
) else (
    ".venv\Scripts\python.exe" app.py
)
