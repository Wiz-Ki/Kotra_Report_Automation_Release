@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=%CD%\portable_python\python.exe"

if not exist "%PYTHON_EXE%" (
    echo portable_python\python.exe was not found.
    echo Run install_offline.bat first.
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%CD%\main.py" %*

endlocal
