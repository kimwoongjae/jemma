@echo off
setlocal

rem Always run from this batch file's folder, even when launched by double-click.
cd /d "%~dp0"
if errorlevel 1 goto folder_error

where py >nul 2>&1
if not errorlevel 1 goto use_py

where python >nul 2>&1
if errorlevel 1 goto python_error

python upload.py
goto finished

:use_py
py -3 upload.py
goto finished

:folder_error
echo [ERROR] Could not open the project folder: %~dp0
set "UPLOAD_EXIT_CODE=1"
goto pause_exit

:python_error
echo [ERROR] Python was not found. Install Python or add it to PATH.
set "UPLOAD_EXIT_CODE=1"
goto pause_exit

:finished
set "UPLOAD_EXIT_CODE=%ERRORLEVEL%"
if "%UPLOAD_EXIT_CODE%"=="0" (
    echo.
    echo Upload completed and verified.
) else (
    echo.
    echo Upload failed. Read the error above; GitHub was not confirmed updated.
)

:pause_exit
echo.
pause
exit /b %UPLOAD_EXIT_CODE%
