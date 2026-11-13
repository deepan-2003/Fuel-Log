@echo off
title Fleet Fuel Log
cd /d "%~dp0"

rem Use "python"; fall back to the "py" launcher if python is not on PATH
set PY=python
%PY% --version >nul 2>&1 || set PY=py

rem First run only: install the required packages
%PY% -c "import streamlit, pandas, openpyxl" >nul 2>&1 || (
    echo Installing required packages, please wait...
    %PY% -m pip install -r requirements.txt
)

echo Starting Fleet Fuel Log... your browser will open shortly.
echo Keep this window open while using the app. Close it to stop the app.
%PY% -m streamlit run app.py

echo.
echo The app stopped. Press any key to close this window.
pause >nul
