@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "ROOT=%CD%"
set "PYTHON_EXE=%ROOT%\portable_python\python.exe"

echo [1/4] Preparing portable Python runtime on this VM...
call "%ROOT%\scripts\prepare_portable_python.bat"
if errorlevel 1 (
    echo VM runtime setup failed.
    pause
    exit /b 1
)

echo [2/4] Checking Microsoft Edge...
set "EDGE_FOUND="
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" set "EDGE_FOUND=1"
if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" set "EDGE_FOUND=1"
where msedge.exe >nul 2>nul
if not errorlevel 1 set "EDGE_FOUND=1"
if not defined EDGE_FOUND (
    echo Microsoft Edge was not found.
    echo Install or enable Microsoft Edge on the VM before running this program.
    pause
    exit /b 1
)

echo [3/4] Verifying Python packages...
"%PYTHON_EXE%" -c "import tkinter, pandas, openpyxl, customtkinter, PIL; from playwright.sync_api import sync_playwright"
if errorlevel 1 (
    echo Runtime package verification failed.
    pause
    exit /b 1
)

echo [4/4] Verifying application imports...
"%PYTHON_EXE%" -c "import gui_launcher; print('application OK')"
if errorlevel 1 (
    echo Application import check failed.
    pause
    exit /b 1
)

echo.
echo VM setup passed.
echo Use run_gui.bat to start the application.
pause

endlocal
