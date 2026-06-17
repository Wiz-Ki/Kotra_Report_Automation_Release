@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "RUNTIME_FILE=%CD%\.runtime_python.bat"
set "PYTHON_EXE=%CD%\portable_python\python.exe"
set "PYTHONW_EXE=%CD%\portable_python\pythonw.exe"

if exist "%RUNTIME_FILE%" call "%RUNTIME_FILE%"

if not exist "%PYTHON_EXE%" (
    echo Python runtime was not found.
    echo Run install_vm.bat first, then open dev_shell.bat again.
    pause
    exit /b 1
)

for %%I in ("%PYTHON_EXE%") do set "PYTHON_DIR=%%~dpI"
set "PYTHON_DIR=%PYTHON_DIR:~0,-1%"

if exist "%PYTHON_DIR%\Scripts" (
    set "PATH=%PYTHON_DIR%;%PYTHON_DIR%\Scripts;%PATH%"
) else (
    set "PATH=%PYTHON_DIR%;%PATH%"
)

echo.
echo Development shell ready.
echo Python: %PYTHON_EXE%
echo.
echo You can now use:
echo   python -V
echo   pip install jupyterlab
echo   jupyter lab
echo.

endlocal & set "PATH=%PATH%" & set "PYTHON_EXE=%PYTHON_EXE%" & set "PYTHONW_EXE=%PYTHONW_EXE%" & cmd /k
