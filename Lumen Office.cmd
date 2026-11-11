@echo off
rem Starts the Lumen Office local server and opens the start page in its own window.
rem Usage: "Lumen Office.cmd" [write|cells|slides]
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw "tools\serve.py" %1 & exit /b)
where python >nul 2>nul && (start "Lumen Office server" /min python "tools\serve.py" %1 & exit /b)
echo Python 3 is needed. Install it from https://www.python.org/downloads/ and run this again.
pause
