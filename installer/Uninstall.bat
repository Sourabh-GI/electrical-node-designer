@echo off
echo Uninstalling Electrical Node Designer...

:: Stop server
taskkill /f /im node.exe >nul 2>&1

:: Remove catalog folder
rd /s /q "%APPDATA%\ElectricalNodeDesignerCatalog" 2>nul

:: Remove registry entry
PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$keyPath = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{12345678-1234-1234-1234-123456789ABC}';" ^
  "if (Test-Path $keyPath) { Remove-Item -Path $keyPath -Force };" ^
  "Write-Host 'Registry entry removed'"

:: Clear Office cache
rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" 2>nul

echo Uninstall complete.
pause
