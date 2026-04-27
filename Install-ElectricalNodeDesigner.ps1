# Electrical Node Designer - Excel Add-in Installer
# Double-click to install, or right-click -> Run with PowerShell

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Electrical Node Designer - Excel Add-in  " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check if Excel is running and warn user
$excelRunning = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelRunning) {
    Write-Host "WARNING: Excel is currently open." -ForegroundColor Yellow
    Write-Host "Please close Excel before continuing." -ForegroundColor Yellow
    Write-Host ""
    $response = Read-Host "Press Enter after closing Excel, or type 'skip' to continue anyway"
}

# Define paths
$manifestUrl = "https://goodlyinsights.sharepoint.com/sites/ElectricalNodeDesigner/AddInFiles/manifest.xml"
$catalogPath = "$env:APPDATA\Microsoft\Excel\XLSTART"
$manifestLocalPath = "$env:TEMP\ElectricalNodeDesigner-manifest.xml"
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"

Write-Host "Step 1: Downloading manifest..." -ForegroundColor Green

# Download manifest from SharePoint
try {
    Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestLocalPath -UseDefaultCredentials
    Write-Host "         Manifest downloaded successfully." -ForegroundColor Green
} catch {
    Write-Host "         Could not download from SharePoint. Using local manifest..." -ForegroundColor Yellow

    # Fallback: look for manifest in same folder as script
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $localManifest = Join-Path $scriptDir "manifest.xml"
    if (Test-Path $localManifest) {
        Copy-Item $localManifest $manifestLocalPath
        Write-Host "         Local manifest found and copied." -ForegroundColor Green
    } else {
        Write-Host "ERROR: No manifest found. Please ensure manifest.xml is in the same folder as this script." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""
Write-Host "Step 2: Setting up trusted catalog..." -ForegroundColor Green

# Create catalog folder if it doesn't exist
$catalogFolder = "$env:TEMP\ElectricalNodeDesignerCatalog"
if (-not (Test-Path $catalogFolder)) {
    New-Item -ItemType Directory -Path $catalogFolder -Force | Out-Null
}

# Copy manifest to catalog folder
Copy-Item $manifestLocalPath "$catalogFolder\manifest.xml" -Force

# Share the folder (needed for Excel to trust it)
$shareName = "ElectricalNodeDesignerCatalog"
$existingShare = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if (-not $existingShare) {
    try {
        New-SmbShare -Name $shareName -Path $catalogFolder -FullAccess "Everyone" | Out-Null
        Write-Host "         Shared catalog folder created." -ForegroundColor Green
    } catch {
        Write-Host "         Note: Could not create network share (may need admin rights)." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Step 3: Registering in Excel trusted catalogs..." -ForegroundColor Green

# Add to Excel trusted catalogs via registry
$guid = [System.Guid]::NewGuid().ToString("B")
$catalogKey = "$registryPath\$guid"

try {
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    New-Item -Path $catalogKey -Force | Out-Null
    Set-ItemProperty -Path $catalogKey -Name "Id" -Value $guid
    Set-ItemProperty -Path $catalogKey -Name "Url" -Value "\\localhost\$shareName"
    Set-ItemProperty -Path $catalogKey -Name "Flags" -Value 1 -Type DWord

    Write-Host "         Registry entry created successfully." -ForegroundColor Green
} catch {
    Write-Host "         Registry update failed: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Step 4: Clearing Excel add-in cache..." -ForegroundColor Green

$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefPath) {
    try {
        Remove-Item "$wefPath\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "         Cache cleared successfully." -ForegroundColor Green
    } catch {
        Write-Host "         Could not clear cache - please clear manually if needed." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Open Excel" -ForegroundColor White
Write-Host "  2. Go to Insert -> Add-ins -> My Add-ins" -ForegroundColor White
Write-Host "  3. Click 'Shared Folder' tab" -ForegroundColor White
Write-Host "  4. Select 'Electrical Node Designer'" -ForegroundColor White
Write-Host "  5. Click Add" -ForegroundColor White
Write-Host ""
Write-Host "The add-in will appear in your Excel Home tab ribbon." -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
