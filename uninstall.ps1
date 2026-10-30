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
        if ($rawVer -and $rawVer -match '^(\d+\.\d+\.\d+)') {
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
$statusLabel = if ($wasInstalled) { "제거 완료 (Uninstalled)" } else { "설치 내역 없음 (Already Clean)" }
$removedVerLabel = if ($removedVersion) { $removedVersion } elseif ($wasInstalled) { "설치됨 (버전 정보 없음)" } else { "없음" }
$unregisteredAppsText = if ($unregisteredApps.Count -gt 0) { ($unregisteredApps -join ', ') } else { "없음" }
$suiteStatusText = if ($suiteRemoved) { "삭제 완료 ($suiteDir)" } else { "해당 없음" }
$cacheStatusText = if ($cacheRemoved) { "정리 완료 ($cacheDir)" } else { "해당 없음" }

Write-Host ""
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host " [SUMMARY] func_SMILES2IMG 제거 요약" -ForegroundColor Cyan
Write-Host "----------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  * 제거 상태 (Status):        $statusLabel" -ForegroundColor $(if ($wasInstalled) { "Green" } else { "Yellow" })
Write-Host "  * 제거된 버전 (Removed Ver): $removedVerLabel" -ForegroundColor White
Write-Host "  * 등록 해제된 프로그램:      $unregisteredAppsText" -ForegroundColor White
Write-Host "  * 추가 기능 파일 정리:       $suiteStatusText" -ForegroundColor Gray
Write-Host "  * 로컬 캐시 폴더 정리:       $cacheStatusText" -ForegroundColor Gray
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[SUCCESS] func_SMILES2IMG 제거가 완료되었습니다." -ForegroundColor Cyan
