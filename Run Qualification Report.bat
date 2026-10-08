@echo off
setlocal
title Qualification Coverage Report

cd /d "%~dp0"

echo ================================================
echo Qualification Coverage Report
echo ================================================
echo.
echo Generating report...
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The Python virtual environment was not found.
    echo Expected location:
    echo %CD%\.venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" "Scripts\report_generator.py"

if errorlevel 1 (
    echo.
    echo ================================================
    echo REPORT GENERATION FAILED
    echo ================================================
    echo Review the error message above.
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================
echo REPORT COMPLETED SUCCESSFULLY
echo ================================================
echo.
echo Opening the Output folder...

start "" "%CD%\Output"

echo.
echo This window will close in 5 seconds.
timeout /t 5 /nobreak >nul

endlocal
exit /b 0