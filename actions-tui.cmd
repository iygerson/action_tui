@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "LOCAL_VENV_PY=%SCRIPT_DIR%.venv\Scripts\python.exe"

if exist "%LOCAL_VENV_PY%" (
  "%LOCAL_VENV_PY%" "%SCRIPT_DIR%actions_tui.py" %*
  exit /b %errorlevel%
)

py -3 "%SCRIPT_DIR%actions_tui.py" %*
