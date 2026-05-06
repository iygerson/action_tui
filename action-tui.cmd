@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "LOCAL_VENDOR=%SCRIPT_DIR%vendor"
set "LOCAL_DEPS=%SCRIPT_DIR%.deps"
set "SHARED_VENV_PY=%SCRIPT_DIR%..\..\notes_tool\.venv\Scripts\python.exe"
set "LOCAL_VENV_PY=%SCRIPT_DIR%.venv\Scripts\python.exe"
if exist "%SHARED_VENV_PY%" (
  "%SHARED_VENV_PY%" "%SCRIPT_DIR%actions_tui.py" %*
  exit /b %errorlevel%
) else if exist "%LOCAL_VENV_PY%" (
  if exist "%LOCAL_VENDOR%" (
    set "PYTHONPATH=%LOCAL_VENDOR%;%PYTHONPATH%"
  ) else if exist "%LOCAL_DEPS%" (
    set "PYTHONPATH=%LOCAL_DEPS%;%PYTHONPATH%"
  )
  "%LOCAL_VENV_PY%" "%SCRIPT_DIR%actions_tui.py" %*
  exit /b %errorlevel%
) else (
  if exist "%LOCAL_VENDOR%" (
    set "PYTHONPATH=%LOCAL_VENDOR%;%PYTHONPATH%"
  ) else if exist "%LOCAL_DEPS%" (
    set "PYTHONPATH=%LOCAL_DEPS%;%PYTHONPATH%"
  )
  py -3 "%SCRIPT_DIR%actions_tui.py" %*
)
