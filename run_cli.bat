@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "RUNTIME_FILE=%CD%\.runtime_python.bat"
set "PYTHON_EXE=%CD%\portable_python\python.exe"
if exist "%RUNTIME_FILE%" call "%RUNTIME_FILE%"

if not exist "%PYTHON_EXE%" (
    echo Python runtime was not found.
    echo Run install_vm.bat first.
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%CD%\main.py" %*

endlocal
