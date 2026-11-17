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

Write-Host "=== func_SMILES2IMG Uninstaller (Excel, Word, PowerPoint) ===" -ForegroundColor Cyan

# Version extractor helper
function Get-InstalledFileVersion($filePath) {
    if (-not $filePath -or -not (Test-Path -LiteralPath $filePath)) { return $null }
    try {
        $info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($filePath)
        $rawVer = if ($info.ProductVersion) { $info.ProductVersion } elseif ($info.FileVersion) { $info.FileVersion } else { $null }
        if ($rawVer -and $rawVer -match '^(\d+\.\d+\.\d+(?:-[A-Za-z0-9.]+)?)') {
            return "v$($matches[1])"
        }
        return $rawVer
    } catch {
        return $null
    }
}

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

# Detect previously installed version before removing files
$addInsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$suiteDir  = Join-Path $addInsDir 'Smiles2Img'
$cacheDir  = Join-Path $env:LOCALAPPDATA 'Smiles2Img'

$candidateFiles = @(
    (Join-Path $suiteDir 'Smiles2Img.Core.dll'),
    (Join-Path $suiteDir 'Smiles2Img.dll'),
    (Join-Path $suiteDir 'Smiles2Img-AddIn64.xll'),
    (Join-Path $suiteDir 'Smiles2Img-AddIn.xll'),
    (Join-Path $addInsDir 'Smiles2Img-AddIn64-packed.xll'),
    (Join-Path $addInsDir 'Smiles2Img-AddIn-packed.xll')
)

$removedVersion = $null
$wasInstalled = $false
foreach ($candidate in $candidateFiles) {
    if (Test-Path -LiteralPath $candidate) {
        $wasInstalled = $true
        $ver = Get-InstalledFileVersion $candidate
        if ($ver) {
            $removedVersion = $ver
            break
        }
    }
}

$unregisteredApps = [System.Collections.Generic.List[string]]::new()

# 2. Remove Excel registry entries
$excelUnregistered = $false
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
                $excelUnregistered = $true
                $wasInstalled = $true
            }
        }
    }
}
if ($excelUnregistered) { $unregisteredApps.Add("Excel") }

# 3. Remove Office Addin registration keys
$officeAddinTargets = @(
    @{ App = "Word";       Key = "HKCU:\Software\Microsoft\Office\Word\Addins\Smiles2Img.Word.AddIn" },
    @{ App = "PowerPoint"; Key = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Smiles2Img.PowerPoint.AddIn" }
)
foreach ($target in $officeAddinTargets) {
    if (Test-Path $target.Key) {
        Remove-Item -Path $target.Key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "[+] Removed Office Add-in key: $($target.Key)" -ForegroundColor Green
        $unregisteredApps.Add($target.App)
        $wasInstalled = $true
    }
}

# 4. Remove COM registrations (HKCU and Wow6432Node) & Toast AUMID
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
            $wasInstalled = $true
        }
    }

    foreach ($clsid in $clsids) {
        $cKey = "$root\CLSID\$clsid"
        if (Test-Path $cKey) {
            Remove-Item -Path $cKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "[+] Removed COM CLSID: $cKey" -ForegroundColor Green
            $wasInstalled = $true
        }
    }
}

$aumidKey = "HKCU:\Software\Classes\AppUserModelId\func_SMILES2IMG"
if (Test-Path $aumidKey) {
    Remove-Item -Path $aumidKey -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "[+] Removed Toast Notification AUMID: $aumidKey" -ForegroundColor Green
}

# 5. Remove installed files
$suiteRemoved = $false
$xllFiles = @('Smiles2Img-AddIn64-packed.xll', 'Smiles2Img-AddIn-packed.xll', 'Smiles2Img-AddIn64.xll', 'Smiles2Img-AddIn.xll')
foreach ($f in $xllFiles) {
    $xllPath = Join-Path $addInsDir $f
    if (Test-Path -LiteralPath $xllPath) {
        Remove-Item -LiteralPath $xllPath -Force -ErrorAction SilentlyContinue
        Write-Host "[+] Deleted Excel Add-in: $xllPath" -ForegroundColor Gray
        $suiteRemoved = $true
        $wasInstalled = $true
    }
}

if (Test-Path -LiteralPath $suiteDir) {
    Remove-Item -LiteralPath $suiteDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "[+] Deleted func_SMILES2IMG directory: $suiteDir" -ForegroundColor Gray
    $suiteRemoved = $true
    $wasInstalled = $true
}

# 6. Remove cache directory
$cacheRemoved = $false
if (Test-Path -LiteralPath $cacheDir) {
    Remove-Item -LiteralPath $cacheDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "[+] Cleaned cache directory: $cacheDir" -ForegroundColor Gray
    $cacheRemoved = $true
}

# 7. Uninstallation Summary Output
$statusLabel = if ($wasInstalled) { "$([char]0xC81C)$([char]0xAC70) $([char]0xC644)$([char]0xB8CC) (Uninstalled)" } else { "$([char]0xC124)$([char]0xCE58) $([char]0xB0B4)$([char]0xC5ED) $([char]0xC5C6)$([char]0xC74C) (Already Clean)" }
$removedVerLabel = if ($removedVersion) { $removedVersion } elseif ($wasInstalled) { "$([char]0xC124)$([char]0xCE58)$([char]0xB428) ($([char]0xBC84)$([char]0xC804) $([char]0xC815)$([char]0xBCF4) $([char]0xC5C6)$([char]0xC74C))" } else { "$([char]0xC5C6)$([char]0xC74C)" }
$unregisteredAppsText = if ($unregisteredApps.Count -gt 0) { ($unregisteredApps -join ', ') } else { "$([char]0xC5C6)$([char]0xC74C)" }
$suiteStatusText = if ($suiteRemoved) { "$([char]0xC0AD)$([char]0xC81C) $([char]0xC644)$([char]0xB8CC) ($suiteDir)" } else { "$([char]0xD574)$([char]0xB2F9) $([char]0xC5C6)$([char]0xC74C)" }
$cacheStatusText = if ($cacheRemoved) { "$([char]0xC815)$([char]0xB9AC) $([char]0xC644)$([char]0xB8CC) ($cacheDir)" } else { "$([char]0xD574)$([char]0xB2F9) $([char]0xC5C6)$([char]0xC74C)" }

Write-Host ""
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host " [SUMMARY] func_SMILES2IMG $([char]0xC81C)$([char]0xAC70) $([char]0xC694)$([char]0xC57D)" -ForegroundColor Cyan
Write-Host "----------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  * $([char]0xC81C)$([char]0xAC70) $([char]0xC0C1)$([char]0xD0DC) (Status):        $statusLabel" -ForegroundColor $(if ($wasInstalled) { "Green" } else { "Yellow" })
Write-Host "  * $([char]0xC81C)$([char]0xAC70)$([char]0xB41C) $([char]0xBC84)$([char]0xC804) (Removed Ver): $removedVerLabel" -ForegroundColor White
Write-Host "  * $([char]0xB4F1)$([char]0xB85D) $([char]0xD574)$([char]0xC81C)$([char]0xB41C) $([char]0xD504)$([char]0xB85C)$([char]0xADF8)$([char]0xB7A8):      $unregisteredAppsText" -ForegroundColor White
Write-Host "  * $([char]0xCD94)$([char]0xAC00) $([char]0xAE30)$([char]0xB2A5) $([char]0xD30C)$([char]0xC77C) $([char]0xC815)$([char]0xB9AC):       $suiteStatusText" -ForegroundColor Gray
Write-Host "  * $([char]0xB85C)$([char]0xCEEC) $([char]0xCE90)$([char]0xC2DC) $([char]0xD3F4)$([char]0xB354) $([char]0xC815)$([char]0xB9AC):       $cacheStatusText" -ForegroundColor Gray
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[SUCCESS] func_SMILES2IMG $([char]0xC81C)$([char]0xAC70)$([char]0xAC00) $([char]0xC644)$([char]0xB8CC)$([char]0xB418)$([char]0xC5C8)$([char]0xC2B5)$([char]0xB2C8)$([char]0xB2E4)." -ForegroundColor Cyan
