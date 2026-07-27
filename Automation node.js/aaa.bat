@echo off
cd /d "%~dp0"

echo ============================================
echo   Excel CRM
echo ============================================

:: Check Node.js
node -v >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found.
    echo Please install from https://nodejs.org and run this again.
    pause
    exit /b 1
)

:: Install packages if node_modules missing
if not exist "node_modules" (
    echo Installing packages...
    npm install
    if errorlevel 1 ( echo [ERROR] npm install failed. & pause & exit /b 1 )
)

:: Install Playwright browser if missing
if not exist "%USERPROFILE%\AppData\Local\ms-playwright" (
    echo Installing Playwright browser...
    node node_modules\playwright\cli.js install chromium
    if errorlevel 1 ( echo [ERROR] Playwright install failed. & pause & exit /b 1 )
)

echo.
echo Starting server...

:: ✅ Start server in background
start "CRM Server" /min node server.js

:: ✅ Wait 2 seconds for server to boot, then open browser
timeout /t 2 /nobreak >nul
start http://127.0.0.1:5000

echo Server is running. Close this window to stop.
pause