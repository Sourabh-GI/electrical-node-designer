@echo off
title Electrical Node Designer - Installer
echo ==========================================
echo   Electrical Node Designer - Installer
echo ==========================================
echo.
echo This installer requires Administrator permission.
echo Please click YES when Windows asks.
echo.
timeout /t 2 >nul

:: Run the full installation as Administrator via PowerShell
PowerShell -NoProfile -ExecutionPolicy Bypass ^
  -Command "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0install.ps1""' -Verb RunAs -Wait"

echo.
echo Installation complete.
echo.
echo NEXT STEPS:
echo   1. Double-click Start-AddIn.bat ^(keep window open^)
echo   2. Open Excel
echo   3. Insert - Add-ins - My Add-ins - Shared Folder
echo   4. Select Electrical Node Designer - Add
echo.
pause
