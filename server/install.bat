@echo off
cd /d "%~dp0"
where npm >nul 2>&1
if errorlevel 1 (
    echo npm not found. Please install Node.js first.
    pause
    exit /b 1
)
if not exist "package.json" (
    echo package.json not found in %CD%
    pause
    exit /b 1
)
call npm install
if errorlevel 1 goto failed
echo Installation complete.
pause
exit /b 0
:failed
echo Installation failed. Review the error above.
pause
exit /b 1