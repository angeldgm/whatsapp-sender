@echo off
setlocal

echo ========================================
echo   Node.js / Express Server Startup
echo ========================================

REM Check if Node.js is installed
where node >nul 2>nul

if %ERRORLEVEL% EQU 0 (
    echo Node.js is already installed.
    node --version
) else (
    echo Node.js is not installed.
    echo Installing Node.js LTS using winget...

    where winget >nul 2>nul

    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: winget is not available.
        echo Please install Node.js manually from https://nodejs.org/
        pause
        exit /b 1
    )

    winget install OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements

    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Failed to install Node.js.
        pause
        exit /b 1
    )

    REM Refresh PATH so the newly installed node can be found
    set "PATH=%ProgramFiles%\nodejs;%PATH%"

    where node >nul 2>nul

    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Node.js was installed but could not be found.
        echo Please close and reopen this terminal, then run this script again.
        pause
        exit /b 1
    )

    echo Node.js installed successfully.
    node --version
)

echo.
echo Installing npm dependencies...
call npm install

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: npm install failed.
    pause
    exit /b 1
)

echo.
echo Opening web browser and starting server...
echo.

start "" http://localhost:3000

node server.js

pause