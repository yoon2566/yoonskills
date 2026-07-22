@echo off
setlocal

cd /d "%~dp0.."
set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
set "SCRIPT=%CD%\examples\button_led_rgb_cycle.py"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Project Python was not found:
    echo %PYTHON_EXE%
    echo Create the project .venv with Python 3.8.10 first.
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo [ERROR] Example script was not found:
    echo %SCRIPT%
    pause
    exit /b 1
)

echo Starting MODI+ button RGB cycle...
echo Press the button: red -^> blue -^> green -^> red -^> ...
echo Press Ctrl+C to stop.
echo.

"%PYTHON_EXE%" "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo Program exit code: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
