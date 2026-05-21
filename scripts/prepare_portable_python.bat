@echo off
setlocal EnableExtensions

cd /d "%~dp0\.."

if not defined PYTHON_VERSION set "PYTHON_VERSION=3.13.5"
if not defined PYTHON_INSTALLER_URL set "PYTHON_INSTALLER_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe"

set "ROOT=%CD%"
set "PYTHON_DIR=%ROOT%\portable_python"
set "DOWNLOAD_DIR=%ROOT%\.setup_downloads"
set "PYTHON_INSTALLER=%DOWNLOAD_DIR%\python-%PYTHON_VERSION%-amd64.exe"
set "PYTHON_INSTALL_LOG=%DOWNLOAD_DIR%\python-install.log"
set "PYTHON_READY="

echo [1/6] Checking runtime requirements...
if not exist "%ROOT%\requirements-runtime.txt" (
    echo Missing requirements-runtime.txt
    exit /b 1
)

mkdir "%DOWNLOAD_DIR%" 2>nul

echo [2/6] Downloading official Python installer...
if not exist "%PYTHON_INSTALLER%" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%PYTHON_INSTALLER_URL%' -OutFile '%PYTHON_INSTALLER%'"
    if errorlevel 1 exit /b 1
) else (
    echo Using cached %PYTHON_INSTALLER%
)

echo [3/6] Installing Python into portable_python...
if exist "%PYTHON_DIR%\python.exe" (
    echo Using existing %PYTHON_DIR%
    "%PYTHON_DIR%\python.exe" -c "import tkinter; print('tkinter OK')" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_READY=1"
    ) else (
        echo Existing portable_python is missing tkinter support.
        call :backup_existing_python
        if errorlevel 1 exit /b 1
    )
)

if not defined PYTHON_READY (
    "%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 TargetDir="%PYTHON_DIR%" DefaultJustForMeTargetDir="%PYTHON_DIR%" Include_launcher=0 Shortcuts=0 PrependPath=0 Include_test=0 Include_doc=0 Include_pip=1 Include_tcltk=1 Include_symbols=0 Include_debug=0 /log "%PYTHON_INSTALL_LOG%"
    if errorlevel 1 exit /b 1
    call :wait_for_python_exe
    if errorlevel 1 exit /b 1
)

if not exist "%PYTHON_DIR%\python.exe" (
    echo Python executable was not found:
    echo %PYTHON_DIR%\python.exe
    echo Installer log:
    echo %PYTHON_INSTALL_LOG%
    exit /b 1
)

echo Installed Python path:
"%PYTHON_DIR%\python.exe" -c "import sys; print(sys.executable)"
if errorlevel 1 exit /b 1

echo Installed Python version:
"%PYTHON_DIR%\python.exe" -V
if errorlevel 1 exit /b 1

echo Checking tkinter files...
if not exist "%PYTHON_DIR%\DLLs\_tkinter.pyd" (
    echo Missing %PYTHON_DIR%\DLLs\_tkinter.pyd
)
if not exist "%PYTHON_DIR%\tcl" (
    echo Missing %PYTHON_DIR%\tcl
)

echo [4/6] Verifying tkinter support...
"%PYTHON_DIR%\python.exe" -c "import tkinter; print('tkinter OK')"
if errorlevel 1 (
    echo tkinter is required for the GUI but was not found.
    echo Recreate portable_python with Include_tcltk=1.
    echo Installer log:
    echo %PYTHON_INSTALL_LOG%
    exit /b 1
)

echo [5/6] Upgrading pip...
"%PYTHON_DIR%\python.exe" -m pip install --upgrade pip
if errorlevel 1 exit /b 1

echo [6/6] Installing runtime packages from the internet...
"%PYTHON_DIR%\python.exe" -m pip install -r "%ROOT%\requirements-runtime.txt"
if errorlevel 1 exit /b 1

echo Verifying portable runtime imports...
"%PYTHON_DIR%\python.exe" -c "import pandas, openpyxl, customtkinter, PIL; from playwright.sync_api import sync_playwright; print('portable runtime OK')"
if errorlevel 1 exit /b 1

echo.
echo VM runtime is ready.
echo Use run_gui.bat to start the application.

endlocal
exit /b 0

:wait_for_python_exe
for /l %%I in (1,1,30) do (
    if not exist "%PYTHON_DIR%\python.exe" (
        timeout /t 1 /nobreak >nul
    ) else (
        exit /b 0
    )
)
echo Python installer finished, but python.exe was not created in:
echo %PYTHON_DIR%
echo Installer log:
echo %PYTHON_INSTALL_LOG%
exit /b 1

:backup_existing_python
set "BACKUP_NAME=portable_python_broken"
:choose_backup_name
if exist "%ROOT%\%BACKUP_NAME%" set "BACKUP_NAME=portable_python_broken_%RANDOM%" & goto choose_backup_name
echo Backing up existing portable_python to %ROOT%\%BACKUP_NAME%
ren "%PYTHON_DIR%" "%BACKUP_NAME%" >nul 2>nul
if errorlevel 1 exit /b 1
if exist "%PYTHON_DIR%" (
    echo Failed to move existing portable_python. Close any running app or terminal using it, then retry.
    exit /b 1
)
exit /b 0
