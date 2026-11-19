@echo off
REM Drag an Excel/CSV/TXT file of URLs onto this file (or double-click it and drag the file in).
python -m pip install --quiet --disable-pip-version-check --user openpyxl playwright dnspython >nul 2>&1
python "%~dp0url_checker.py" %*
pause
