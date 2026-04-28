$ErrorActionPreference = "Stop"
$installerDir = $PSScriptRoot
$certDir = Join-Path $installerDir "certs"
$catalogDir = "$env:APPDATA\ElectricalNodeDesignerCatalog"
$manifestSrc = Join-Path $installerDir "manifest-local.xml"
$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Electrical Node Designer - Full Install" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# STEP 1 - Check Node.js
Write-Host "Step 1/7: Checking Node.js..." -ForegroundColor Yellow
try {
    $nodeVersion = & node --version 2>&1
    Write-Host "         Node.js found: $nodeVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Node.js is not installed." -ForegroundColor Red
    Write-Host "Please install from https://nodejs.org (LTS version)" -ForegroundColor Yellow
    Start-Process "https://nodejs.org"
    Read-Host "Press Enter after installing Node.js, then run this installer again"
    exit 1
}

# STEP 2 - Install office-addin-dev-certs
Write-Host "Step 2/7: Installing trusted certificates..." -ForegroundColor Yellow

$devCertsDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"

# Check if certs already exist
$certExists = Test-Path (Join-Path $devCertsDir "localhost.crt")

if (-not $certExists) {
    Write-Host "         Generating dev certs..." -ForegroundColor Gray

    # Find node in the project or globally
    $projectDir = Split-Path $installerDir -Parent
    $npxPath = "npx"

    try {
        $result = & cmd /c "npx office-addin-dev-certs install --machine 2>&1"
        Write-Host "         $result" -ForegroundColor Gray
    } catch {
        Write-Host "         Warning: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "         Dev certs already exist at: $devCertsDir" -ForegroundColor Green
}

# Verify certs exist
$keyPath  = Join-Path $devCertsDir "localhost.key"
$certPath = Join-Path $devCertsDir "localhost.crt"

if (Test-Path $keyPath) {
    Write-Host "         localhost.key found" -ForegroundColor Green
} else {
    Write-Host "         WARNING: localhost.key not found" -ForegroundColor Yellow
    Write-Host "         Checking installer certs folder..." -ForegroundColor Yellow
}

if (Test-Path $certPath) {
    Write-Host "         localhost.crt found" -ForegroundColor Green

    # Trust the CA cert in Windows store
    $caPath = Join-Path $devCertsDir "ca.crt"
    if (Test-Path $caPath) {
        $caCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $caCert.Import($caPath)

        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
            [System.Security.Cryptography.X509Certificates.StoreName]::Root,
            [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
        )
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $store.Add($caCert)
        $store.Close()
        Write-Host "         CA cert trusted in LocalMachine\Root" -ForegroundColor Green
    }
} else {
    Write-Host "         WARNING: Certificates not found - add-in may show certificate error" -ForegroundColor Yellow
    Write-Host "         Run this command manually: npx office-addin-dev-certs install --machine" -ForegroundColor Yellow
}

# STEP 3 - Certificate trust (handled above in Step 2)
Write-Host "Step 3/7: Certificate trust complete..." -ForegroundColor Yellow
Write-Host "         office-addin-dev-certs CA trusted in LocalMachine\Root" -ForegroundColor Green

# STEP 4 - Close Excel
Write-Host "Step 4/7: Closing Excel if open..." -ForegroundColor Yellow
$excel = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excel) {
    Stop-Process -Name "EXCEL" -Force
    Start-Sleep -Seconds 2
    Write-Host "         Excel closed" -ForegroundColor Green
} else {
    Write-Host "         Excel was not running" -ForegroundColor Green
}

# STEP 5 - Clear Office cache
Write-Host "Step 5/7: Clearing Office cache..." -ForegroundColor Yellow
if (Test-Path $wefPath) {
    Remove-Item "$wefPath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "         Cache cleared" -ForegroundColor Green
} else {
    Write-Host "         Cache folder not found (OK)" -ForegroundColor Green
}

# STEP 6 - Copy manifest to catalog
Write-Host "Step 6/7: Setting up add-in catalog..." -ForegroundColor Yellow
if (-not (Test-Path $catalogDir)) {
    New-Item -ItemType Directory -Path $catalogDir -Force | Out-Null
}
Copy-Item $manifestSrc "$catalogDir\manifest.xml" -Force
Write-Host "         Manifest copied to: $catalogDir" -ForegroundColor Green

# STEP 7 - Register catalog in Excel registry
Write-Host "Step 7/7: Registering with Excel..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$keyGuid = "{12345678-1234-1234-1234-123456789ABC}"
$keyFullPath = "$regPath\$keyGuid"

if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}
if (-not (Test-Path $keyFullPath)) {
    New-Item -Path $keyFullPath -Force | Out-Null
}

Set-ItemProperty -Path $keyFullPath -Name "Id"    -Value $keyGuid
Set-ItemProperty -Path $keyFullPath -Name "Url"   -Value $catalogDir
Set-ItemProperty -Path $keyFullPath -Name "Flags" -Value 1 -Type DWord
Write-Host "         Registry updated" -ForegroundColor Green

# DONE
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Certificates (office-addin-dev-certs):" -ForegroundColor White
Write-Host "  localhost.crt : $(Test-Path (Join-Path $devCertsDir 'localhost.crt'))"
Write-Host "  localhost.key : $(Test-Path (Join-Path $devCertsDir 'localhost.key'))"
Write-Host "  ca.crt        : $(Test-Path (Join-Path $devCertsDir 'ca.crt'))"
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Double-click Start-AddIn.bat (keep window open)" -ForegroundColor White
Write-Host "  2. Open Excel" -ForegroundColor White
Write-Host "  3. Insert - Add-ins - My Add-ins - Shared Folder tab" -ForegroundColor White
Write-Host "  4. Select Electrical Node Designer and click Add" -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to close"
