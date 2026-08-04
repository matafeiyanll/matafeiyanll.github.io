@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel% equ 0 (
  py -3 update-bug-data.py
) else (
  python update-bug-data.py
)

if errorlevel 1 (
  echo.
  echo Update failed. Install Python 3 or run the script from a Python environment.
)

echo.
pause
