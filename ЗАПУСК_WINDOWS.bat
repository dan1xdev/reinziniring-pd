@echo off
setlocal
cd /d "%~dp0"
title Reinziniring PD - local server
chcp 65001 >nul
echo ==============================================
echo   REINZINIRING PD - START
 echo ==============================================
where py >nul 2>nul
if not errorlevel 1 (
  py -3.12 -c "import sys" >nul 2>nul
  if not errorlevel 1 set "PY=py -3.12"
)
if not defined PY (
  where python >nul 2>nul
  if errorlevel 1 goto no_python
  set "PY=python"
)
if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating virtual environment...
  %PY% -m venv .venv
  if errorlevel 1 goto failed
)
echo [2/3] Installing required packages...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo [3/3] Starting server on an available local port...
echo Browser should open automatically. Do not close this window.
echo If an error occurs, take a screenshot of this window.
".venv\Scripts\python.exe" -m ou_processor.web --port 0
if errorlevel 1 goto failed
echo Server stopped.
pause
exit /b 0
:no_python
echo ERROR: Python not found. Install Python 3.12 from python.org.
goto failed
:failed
echo.
echo ERROR: Startup failed. Copy the error above and send it to support.
pause
exit /b 1
