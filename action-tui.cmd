@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "VENV_PY=%SCRIPT_DIR%..\notes_tool\.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
  "%VENV_PY%" "%SCRIPT_DIR%actions_tui.py" %*
) else (
  py -3 "%SCRIPT_DIR%actions_tui.py" %*
)
