@echo off
title Electrical Node Designer - Server
echo ==========================================
echo   Electrical Node Designer - Running
echo ==========================================
echo.

node --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Node.js is not installed.
    echo Please install from https://nodejs.org
    echo.
    start https://nodejs.org
    pause
    exit /b 1
)

echo Server starting on https://localhost:3001
echo Keep this window open while using Excel.
echo.
node "%~dp0server.js"
pause
