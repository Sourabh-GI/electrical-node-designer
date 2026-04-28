@echo off
echo Stopping Electrical Node Designer server...
taskkill /f /im node.exe >nul 2>&1
echo Server stopped.
echo You can close Excel now.
pause
