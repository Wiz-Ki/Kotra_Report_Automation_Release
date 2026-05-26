@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "ROOT=%~dp0.."
set "RUNTIME_FILE=%ROOT%\.runtime_python.bat"
set "PYTHON_EXE=%ROOT%\portable_python\python.exe"
if exist "%RUNTIME_FILE%" call "%RUNTIME_FILE%"

if not exist "%PYTHON_EXE%" (
    echo Python runtime was not found.
    echo Run install_vm.bat in the main program folder first.
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%~dp0check_reports_against_input.py" --folder "%~dp0" %*
if errorlevel 1 (
    echo.
    pause
    exit /b 1
)

echo.
pause
endlocal
