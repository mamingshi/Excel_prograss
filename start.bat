@echo off
setlocal
pushd "%~dp0"
if errorlevel 1 exit /b 1

rem Prefer the project environment so installing and running use the same Python.
if exist ".venv\Scripts\python.exe" (
    set "TASK_PYTHON=%~dp0.venv\Scripts\python.exe"
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo Python was not found. Install Python 3.11+ and add it to PATH.
        pause
        popd
        exit /b 1
    )
    set "TASK_PYTHON=python"
)

set "TASK_LOG=%~dp0startup.log"
echo Starting Excel preview assistant...
"%TASK_PYTHON%" -X utf8 -c "import sys; print('Python:', sys.executable); print('Version:', sys.version)" > "%TASK_LOG%" 2>&1
"%TASK_PYTHON%" -X utf8 -X faulthandler -u main.py >> "%TASK_LOG%" 2>&1
set "TASK_EXIT_CODE=%ERRORLEVEL%"

if not "%TASK_EXIT_CODE%"=="0" (
    echo.
    echo Startup failed. Error details:
    type "%TASK_LOG%"
    echo.
    echo Log saved to: "%TASK_LOG%"
    pause
)

popd
exit /b %TASK_EXIT_CODE%
