<#
.SYNOPSIS
    func_SMILES2IMG Uninstaller (Excel, Word, PowerPoint)
.DESCRIPTION
    Removes func_SMILES2IMG add-in registry entries from Excel, Word, and PowerPoint,
    unregisters COM classes from HKCU, and removes installed files from %APPDATA%\Microsoft\AddIns.
#>
[CmdletBinding()]
param(
    [switch]$Force
)

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

$ErrorActionPreference = 'Stop'

Write-Host "=== func_SMILES2IMG Uninstaller ===" -ForegroundColor Cyan

# 1. Check if Office applications are running
$officeProcs = Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue
if ($officeProcs) {
    $procNames = ($officeProcs | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
    Write-Warning "The following Office application(s) are currently running: $procNames."
    Write-Warning "Please save your work and close them so uninstallation can complete."
    if (-not $Force) {
        $confirm = Read-Host "Close Office applications now? (y/N)"
        if ($confirm -match '^[yY]') {
            $officeProcs | Stop-Process -Force
            Start-Sleep -Seconds 1
        } else {
            Write-Host "Uninstallation aborted." -ForegroundColor Yellow
            return
        }
    } else {
        $officeProcs | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
}

# 2. Remove Excel registry entries
$officeVersions = @('16.0', '15.0', '14.0')
foreach ($ver in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $regPath) {
        $props = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
        $openProps = $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN(\d+)?$' }
        foreach ($p in $openProps) {
            if ($p.Value -match 'Smiles2Img') {
                Remove-ItemProperty -Path $regPath -Name $p.Name -ErrorAction SilentlyContinue
                Write-Host "[+] Removed Excel registry key: $($p.Name) ($ver)" -ForegroundColor Green
            }
        }
    }
}

# 3. Remove Office Addin registration keys
$officeAddinKeys = @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\Smiles2Img.Word.AddIn",
    "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Smiles2Img.PowerPoint.AddIn"
)
foreach ($key in $officeAddinKeys) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "[+] Removed Office Add-in key: $key" -ForegroundColor Green
    }
}

# 4. Remove COM registrations (HKCU and Wow6432Node)
$roots = @("HKCU:\Software\Classes", "HKCU:\Software\Classes\Wow6432Node")
$progIds = @("Smiles2Img.Word.AddIn", "Smiles2Img.PowerPoint.AddIn")
$clsids = @(
    "{41A5C1F0-559A-46E0-B9C0-8B35A0F1D910}",
    "{A755973C-21FB-494A-A3C5-086AFFB7C2AC}"
)

foreach ($root in $roots) {
    foreach ($progId in $progIds) {
        $pKey = "$root\$progId"
        if (Test-Path $pKey) {
            Remove-Item -Path $pKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "[+] Removed COM ProgID: $pKey" -ForegroundColor Green
        }
    }

    foreach ($clsid in $clsids) {
        $cKey = "$root\CLSID\$clsid"
        if (Test-Path $cKey) {
            Remove-Item -Path $cKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "[+] Removed COM CLSID: $cKey" -ForegroundColor Green
        }
    }
}

# 5. Remove installed files
$addInsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$xllFiles = @('Smiles2Img-AddIn64-packed.xll', 'Smiles2Img-AddIn-packed.xll')
foreach ($f in $xllFiles) {
    $xllPath = Join-Path $addInsDir $f
    if (Test-Path $xllPath) {
        Remove-Item -LiteralPath $xllPath -Force -ErrorAction SilentlyContinue
        Write-Host "[+] Deleted Excel Add-in: $xllPath" -ForegroundColor Gray
    }
}

$suiteDir = Join-Path $addInsDir 'Smiles2Img'
if (Test-Path $suiteDir) {
    Remove-Item -LiteralPath $suiteDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "[+] Deleted func_SMILES2IMG directory: $suiteDir" -ForegroundColor Gray
}

# 6. Remove cache directory
$cacheDir = Join-Path $env:LOCALAPPDATA 'Smiles2Img'
if (Test-Path $cacheDir) {
    Remove-Item -LiteralPath $cacheDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "[+] Cleaned cache directory: $cacheDir" -ForegroundColor Gray
}

Write-Host "`n[SUCCESS] func_SMILES2IMG has been successfully uninstalled." -ForegroundColor Cyan
