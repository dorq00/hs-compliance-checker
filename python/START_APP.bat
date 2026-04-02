@echo off
title HS Checker — Web UI
cd /d "%~dp0"
echo.
echo  ==========================================
echo   HS CHECKER — Web UI
echo  ==========================================
echo.
echo  Starting server...
echo  Browser will open at http://localhost:8501
echo.
echo  Press Ctrl+C to stop.
echo.
python -m streamlit run app.py
pause
