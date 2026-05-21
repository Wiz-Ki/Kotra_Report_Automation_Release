@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=%CD%\portable_python\python.exe"
set "PYTHONW_EXE=%CD%\portable_python\pythonw.exe"

if not exist "%PYTHONW_EXE%" (
    echo portable_python\pythonw.exe was not found.
    echo Run install_offline.bat first.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -c "import tkinter; import gui_launcher" >nul 2>nul
if errorlevel 1 (
    echo Application import check failed.
    echo Run install_offline.bat and check the error message.
    pause
    exit /b 1
)

start "" "%PYTHONW_EXE%" "%CD%\gui_launcher.py"

endlocal
