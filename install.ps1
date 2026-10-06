<#
.SYNOPSIS
    SMILES2IMG Excel Add-in Automatic Installer
.DESCRIPTION
    Automatically detects Excel bitness (32-bit vs 64-bit),
    copies or downloads the packed .xll to %APPDATA%\Microsoft\AddIns,
    unblocks the file, and registers it in Excel's add-in registry.
#>
[CmdletBinding()]
param(
    [switch]$Force,
    [string]$Version = $(if ($env:SMILES2IMG_VERSION) { $env:SMILES2IMG_VERSION } else { 'latest' })
)

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

$ErrorActionPreference = 'Stop'

Write-Host "=== SMILES2IMG Installer / Updater ===" -ForegroundColor Cyan

# 1. Check if Excel is running
$excelProc = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Warning "Excel is currently running. Please save your work and close Excel so registration can complete."
    if (-not $Force) {
        $confirm = Read-Host "Close Excel now? (y/N)"
        if ($confirm -match '^[yY]') {
            $excelProc | Stop-Process -Force
            Start-Sleep -Seconds 1
        } else {
            Write-Host "Installation aborted. Please close Excel and run the installer again." -ForegroundColor Yellow
            return
        }
    } else {
        $excelProc | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
}

# 2. Detect Excel Bitness & Version
$bitness = 'x64'
$officeVersion = '16.0'

$c2r = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
if ($c2r -and $c2r.Platform) {
    $bitness = if ($c2r.Platform -eq 'x86') { 'x86' } else { 'x64' }
} else {
    try {
        $excel = New-Object -ComObject Excel.Application
        $bitness = if ($excel.OperatingSystem -match '64-bit') { 'x64' } else { 'x86' }
        if ($excel.Version) { $officeVersion = $excel.Version }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        $bitness = if ([Environment]::Is64BitOperatingSystem) { 'x64' } else { 'x86' }
    }
}

Write-Host "[+] Detected Excel Architecture: $bitness (Office $officeVersion)" -ForegroundColor Green

# 3. Destination Directory
$targetDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

$xllName = if ($bitness -eq 'x64') { 'Smiles2Img-AddIn64-packed.xll' } else { 'Smiles2Img-AddIn-packed.xll' }
$targetPath = Join-Path $targetDir $xllName
$isUpdate = Test-Path -LiteralPath $targetPath

if ($isUpdate) {
    Write-Host "[*] Existing installation found. Updating to the latest version..." -ForegroundColor Yellow
} else {
    Write-Host "[*] Fresh installation starting..." -ForegroundColor Cyan
}

# 4. Copy from local dist if available (latest only), otherwise download packed XLL from GitHub Releases
$tag = if ($Version -eq 'latest' -or [string]::IsNullOrWhiteSpace($Version)) { 'latest' } else { if ($Version -notmatch '^v') { "v$Version" } else { $Version } }
$localDistPath = if ($PSScriptRoot -and $tag -eq 'latest') { Join-Path $PSScriptRoot "dist\$bitness\$xllName" } else { $null }
if ($localDistPath -and (Test-Path -LiteralPath $localDistPath)) {
    Write-Host "[+] Installing from local build: $localDistPath" -ForegroundColor Cyan
    Copy-Item -LiteralPath $localDistPath -Destination $targetPath -Force
} else {
    $downloadUrl = if ($tag -eq 'latest') {
        "https://github.com/naramdash/func_SMILES2IMG/releases/latest/download/$xllName"
    } else {
        "https://github.com/naramdash/func_SMILES2IMG/releases/download/$tag/$xllName"
    }
    Write-Host "[+] Downloading release ($tag): $downloadUrl" -ForegroundColor Cyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $targetPath -UseBasicParsing
}

# 5. Unblock file
Write-Host "[+] Unblocking file (removing Mark of the Web)..." -ForegroundColor Gray
Unblock-File -LiteralPath $targetPath

# 6. Register in Excel Options
$regPath = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel\Options"
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

$props = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
$openKeyToUse = $null

$openProps = $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN(\d+)?$' }
foreach ($p in $openProps) {
    if ($p.Value -match 'Smiles2Img') {
        $openKeyToUse = $p.Name
        break
    }
}

if (-not $openKeyToUse) {
    if (-not $props.OPEN) {
        $openKeyToUse = 'OPEN'
    } else {
        $idx = 1
        while ($props.("OPEN$idx")) {
            $idx++
        }
        $openKeyToUse = "OPEN$idx"
    }
}

$regValue = "/R `"$targetPath`""
Set-ItemProperty -Path $regPath -Name $openKeyToUse -Value $regValue
Write-Host "[+] Registered in Excel Options ($openKeyToUse): $regValue" -ForegroundColor Green

if ($isUpdate) {
    Write-Host "`n[SUCCESS] SMILES2IMG has been successfully updated to the latest version!" -ForegroundColor Cyan
} else {
    Write-Host "`n[SUCCESS] SMILES2IMG has been successfully installed and registered!" -ForegroundColor Cyan
}
Write-Host "Open Excel and enter: =SMILES2IMG(""CCO"")" -ForegroundColor White

# 7. Check Windows 11 Smart App Control (SAC)
$sacPolicy = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy" -ErrorAction SilentlyContinue
if ($sacPolicy -and $sacPolicy.VerifiedAndReputablePolicyState -eq 1) {
    Write-Host "`n----------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[!] Notice: Windows 11 Smart App Control (SAC) is enabled on this system." -ForegroundColor Yellow
    Write-Host "    If Excel shows '#NAME?' or 'file format and extension do not match' warnings:" -ForegroundColor Gray
    Write-Host "    1. Open Windows Security > App & browser control > Smart App Control settings" -ForegroundColor Gray
    Write-Host "    2. Set it to 'Off' (Real-time Microsoft Defender Antivirus remains 100% active)" -ForegroundColor Gray
    Write-Host "    3. Restart Excel to use SMILES2IMG normally." -ForegroundColor Gray
    Write-Host "----------------------------------------------------------------------" -ForegroundColor Yellow
}
