@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "RUNTIME_FILE=%CD%\.runtime_python.bat"
set "PYTHON_EXE=%CD%\portable_python\python.exe"
set "PYTHONW_EXE=%CD%\portable_python\pythonw.exe"
if exist "%RUNTIME_FILE%" call "%RUNTIME_FILE%"
if not exist "%PYTHONW_EXE%" if exist "%PYTHON_EXE%" set "PYTHONW_EXE=%PYTHON_EXE%"

if not exist "%PYTHONW_EXE%" (
    echo Python runtime was not found.
    echo Run install_vm.bat first.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -c "import tkinter; import gui_launcher" >nul 2>nul
if errorlevel 1 (
    echo Application import check failed.
    echo Run install_vm.bat and check the error message.
    pause
    exit /b 1
)

start "" "%PYTHONW_EXE%" "%CD%\gui_launcher.py"

endlocal
