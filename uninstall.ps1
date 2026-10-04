<#
.SYNOPSIS
    SMILES2IMG Excel Add-in Uninstaller
.DESCRIPTION
    Removes SMILES2IMG from Excel's add-in registry and deletes the .xll from %APPDATA%\Microsoft\AddIns.
#>
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

Write-Host "=== SMILES2IMG Uninstaller ===" -ForegroundColor Cyan

# 1. Check if Excel is running
$excelProc = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Warning "Excel is currently running. Please save your work and close Excel so uninstallation can complete."
    if (-not $Force) {
        $confirm = Read-Host "Close Excel now? (y/N)"
        if ($confirm -match '^[yY]') {
            $excelProc | Stop-Process -Force
            Start-Sleep -Seconds 1
        } else {
            Write-Host "Uninstallation aborted." -ForegroundColor Yellow
            return
        }
    } else {
        $excelProc | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
}

# 2. Remove registry keys
$officeVersion = '16.0'
$regPath = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel\Options"
if (Test-Path $regPath) {
    $props = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
    $openProps = $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN(\d+)?$' }
    foreach ($p in $openProps) {
        if ($p.Value -match 'Smiles2Img') {
            Remove-ItemProperty -Path $regPath -Name $p.Name -ErrorAction SilentlyContinue
            Write-Host "[+] Removed registry key: $($p.Name)" -ForegroundColor Green
        }
    }
}

# 3. Remove .xll files from AddIns directory
$targetDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
@('Smiles2Img-AddIn64-packed.xll', 'Smiles2Img-AddIn-packed.xll') | ForEach-Object {
    $file = Join-Path $targetDir $_
    if (Test-Path $file) {
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        Write-Host "[+] Deleted: $file" -ForegroundColor Gray
    }
}

Write-Host "`n[SUCCESS] SMILES2IMG has been successfully uninstalled." -ForegroundColor Cyan
