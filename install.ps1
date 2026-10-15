<#
.SYNOPSIS
    SMILES for Microsoft Office Suite Automatic Installer (Excel, Word, PowerPoint)
.DESCRIPTION
    Automatically detects Office bitness (32-bit vs 64-bit),
    installs Excel Add-in (.xll) and Word/PowerPoint COM Add-ins (.dll),
    unblocks all files, and registers them in Office and Windows COM registry (HKCU, no admin required).
#>
[CmdletBinding()]
param(
    [switch]$Force,
    [string]$Version = $(
        if ($v) { $v }
        elseif ($env:SMILES2IMG_VERSION) { $env:SMILES2IMG_VERSION }
        else { 'latest' }
    )
)

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

$ErrorActionPreference = 'Stop'

Write-Host "=== SMILES for Microsoft Office Suite Installer (Excel, Word, PowerPoint) ===" -ForegroundColor Cyan

# 1. Check if Office applications are running
$officeProcs = Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue
if ($officeProcs) {
    $procNames = ($officeProcs | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
    Write-Warning "The following Office application(s) are currently running: $procNames."
    Write-Warning "Please save your work and close them so add-in registration can complete."
    if (-not $Force) {
        $confirm = Read-Host "Close Office applications now? (y/N)"
        if ($confirm -match '^[yY]') {
            $officeProcs | Stop-Process -Force
            Start-Sleep -Seconds 1
        } else {
            Write-Host "Installation aborted. Please close Office apps and run the installer again." -ForegroundColor Yellow
            return
        }
    } else {
        $officeProcs | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
}

# 2. Detect Office Bitness & Version
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

Write-Host "[+] Detected Office Architecture: $bitness (Office $officeVersion)" -ForegroundColor Green

# 3. Target Directories
$addInsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$suiteDir = Join-Path $addInsDir 'Smiles2Img'
if (-not (Test-Path $addInsDir)) { New-Item -ItemType Directory -Path $addInsDir -Force | Out-Null }
if (-not (Test-Path $suiteDir))  { New-Item -ItemType Directory -Path $suiteDir -Force | Out-Null }

$tag = if ($Version -eq 'latest' -or [string]::IsNullOrWhiteSpace($Version)) { 'latest' } else { if ($Version -notmatch '^v') { "v$Version" } else { $Version } }

# -------------------------------------------------------------
# 4. Deploy Office Suite Files (%APPDATA%\Microsoft\AddIns\Smiles2Img)
# -------------------------------------------------------------
$unpackedXllName = if ($bitness -eq 'x64') { 'Smiles2Img-AddIn64.xll' } else { 'Smiles2Img-AddIn.xll' }
$dnaName         = if ($bitness -eq 'x64') { 'Smiles2Img-AddIn64.dna' } else { 'Smiles2Img-AddIn.dna' }
$packedXllName   = if ($bitness -eq 'x64') { 'Smiles2Img-AddIn64-packed.xll' } else { 'Smiles2Img-AddIn-packed.xll' }

$localWordDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot "dist\$bitness\Word" } else { $null }
$localPptDir  = if ($PSScriptRoot) { Join-Path $PSScriptRoot "dist\$bitness\PowerPoint" } else { $null }
$localBinDir  = if ($PSScriptRoot) { Join-Path $PSScriptRoot "src\Smiles2Img.Excel\bin\Release\net48" } else { $null }

if ($localWordDir -and (Test-Path $localWordDir) -and $localPptDir -and (Test-Path $localPptDir) -and $localBinDir -and (Test-Path $localBinDir)) {
    Write-Host "[+] Installing Office Suite files from local build..." -ForegroundColor Cyan
    Copy-Item -Path "$localWordDir\*.dll" -Destination $suiteDir -Force
    Copy-Item -Path "$localPptDir\*.dll" -Destination $suiteDir -Force
    Copy-Item -Path "$localBinDir\*.dll" -Destination $suiteDir -Force
    Copy-Item -LiteralPath (Join-Path $localBinDir $unpackedXllName) -Destination (Join-Path $suiteDir $unpackedXllName) -Force
    Copy-Item -LiteralPath (Join-Path $localBinDir $dnaName) -Destination (Join-Path $suiteDir $dnaName) -Force
} else {
    # Remote download suite zip
    $zipName = "Smiles2Img-Office-$bitness.zip"
    $zipUrl = if ($tag -eq 'latest') {
        "https://github.com/naramdash/func_SMILES2IMG/releases/latest/download/$zipName"
    } else {
        "https://github.com/naramdash/func_SMILES2IMG/releases/download/$tag/$zipName"
    }
    $tempZip = Join-Path $env:TEMP $zipName
    try {
        Write-Host "[+] Downloading Office Suite package ($tag): $zipUrl" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing
        Expand-Archive -LiteralPath $tempZip -DestinationPath $suiteDir -Force
    } catch {
        Write-Warning "Office Suite zip package not found on release ($($_.Exception.Message))."
        Write-Warning "Falling back to standalone packed Excel Add-in..."
    } finally {
        if (Test-Path $tempZip) { Remove-Item $tempZip -Force -ErrorAction SilentlyContinue }
    }
}

# -------------------------------------------------------------
# 5. Register Excel Add-in (.xll)
# -------------------------------------------------------------
$targetXllPath = Join-Path $suiteDir $unpackedXllName
if (-not (Test-Path -LiteralPath $targetXllPath)) {
    # Fallback to packed XLL
    $targetXllPath = Join-Path $addInsDir $packedXllName
    $localPacked = if ($PSScriptRoot -and $tag -eq 'latest') { Join-Path $PSScriptRoot "dist\$bitness\$packedXllName" } else { $null }
    if ($localPacked -and (Test-Path -LiteralPath $localPacked)) {
        Copy-Item -LiteralPath $localPacked -Destination $targetXllPath -Force
    } else {
        $downloadUrl = if ($tag -eq 'latest') {
            "https://github.com/naramdash/func_SMILES2IMG/releases/latest/download/$packedXllName"
        } else {
            "https://github.com/naramdash/func_SMILES2IMG/releases/download/$tag/$packedXllName"
        }
        Invoke-WebRequest -Uri $downloadUrl -OutFile $targetXllPath -UseBasicParsing
    }
} else {
    # Clean up legacy slow packed XLL from AddIns directory if present
    $oldPacked = Join-Path $addInsDir $packedXllName
    if (Test-Path -LiteralPath $oldPacked) {
        Remove-Item -LiteralPath $oldPacked -Force -ErrorAction SilentlyContinue
    }
}

Unblock-File -LiteralPath $targetXllPath

# Register in Excel Options
$excelRegPath = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel\Options"
if (-not (Test-Path $excelRegPath)) { New-Item -Path $excelRegPath -Force | Out-Null }

$props = Get-ItemProperty $excelRegPath -ErrorAction SilentlyContinue
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
        while ($props.("OPEN$idx")) { $idx++ }
        $openKeyToUse = "OPEN$idx"
    }
}
$regValue = "/R `"$targetXllPath`""
Set-ItemProperty -Path $excelRegPath -Name $openKeyToUse -Value $regValue
Write-Host "[+] Excel Add-in registered ($openKeyToUse): $regValue" -ForegroundColor Green

# -------------------------------------------------------------
# 6. Register Word & PowerPoint COM Add-ins (.dll)
# -------------------------------------------------------------

# Unblock all DLLs in suite directory
Get-ChildItem -Path $suiteDir -Filter "*.dll" | ForEach-Object {
    Unblock-File -LiteralPath $_.FullName
}

# Helper to register COM AddIn in HKCU (Zero Admin Required)
function Register-OfficeComAddIn {
    param(
        [string]$ProgId,
        [string]$Clsid,
        [string]$ClassName,
        [string]$AssemblyString,
        [string]$DllPath,
        [string]$OfficeApp,
        [string]$FriendlyName,
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $DllPath)) {
        Write-Warning "File not found: $DllPath. Skipping $OfficeApp add-in registration."
        return
    }

    $codeBaseUri = ([uri]$DllPath).AbsoluteUri
    $roots = @("HKCU:\Software\Classes", "HKCU:\Software\Classes\Wow6432Node")

    foreach ($root in $roots) {
        # ProgID
        $pKey = "$root\$ProgId"
        New-Item -Path $pKey -Force | Out-Null
        Set-ItemProperty -Path $pKey -Name "(Default)" -Value $ClassName
        New-Item -Path "$pKey\CLSID" -Force | Out-Null
        Set-ItemProperty -Path "$pKey\CLSID" -Name "(Default)" -Value $Clsid

        # CLSID
        $cKey = "$root\CLSID\$Clsid"
        New-Item -Path $cKey -Force | Out-Null
        Set-ItemProperty -Path $cKey -Name "(Default)" -Value $ClassName

        # InprocServer32
        $sKey = "$cKey\InprocServer32"
        New-Item -Path $sKey -Force | Out-Null
        Set-ItemProperty -Path $sKey -Name "(Default)" -Value "mscoree.dll"
        Set-ItemProperty -Path $sKey -Name "ThreadingModel" -Value "Both"
        Set-ItemProperty -Path $sKey -Name "Class" -Value $ClassName
        Set-ItemProperty -Path $sKey -Name "Assembly" -Value $AssemblyString
        Set-ItemProperty -Path $sKey -Name "RuntimeVersion" -Value "v4.0.30319"
        Set-ItemProperty -Path $sKey -Name "CodeBase" -Value $codeBaseUri

        # ProgId subkey under CLSID
        New-Item -Path "$cKey\ProgId" -Force | Out-Null
        Set-ItemProperty -Path "$cKey\ProgId" -Name "(Default)" -Value $ProgId

        # Implemented Categories
        New-Item -Path "$cKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null
    }

    # Office App Addin registration
    $addInKey = "HKCU:\Software\Microsoft\Office\$OfficeApp\Addins\$ProgId"
    New-Item -Path $addInKey -Force | Out-Null
    Set-ItemProperty -Path $addInKey -Name "FriendlyName" -Value $FriendlyName
    Set-ItemProperty -Path $addInKey -Name "Description" -Value $Description
    Set-ItemProperty -Path $addInKey -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $addInKey -Name "CommandLineSafe" -Value 0 -Type DWord

    Write-Host "[+] $OfficeApp COM Add-in registered: $FriendlyName" -ForegroundColor Green
}

# Register Word COM Add-in
$wordDllPath = Join-Path $suiteDir 'Smiles2Img.Word.dll'
if (Test-Path $wordDllPath) {
    $wordAssembly = try { [System.Reflection.AssemblyName]::GetAssemblyName($wordDllPath).FullName } catch { "Smiles2Img.Word, Version=1.5.0.0, Culture=neutral, PublicKeyToken=null" }
    Register-OfficeComAddIn `
        -ProgId "Smiles2Img.Word.AddIn" `
        -Clsid "{41A5C1F0-559A-46E0-B9C0-8B35A0F1D910}" `
        -ClassName "Smiles2Img.Word.WordAddIn" `
        -AssemblyString $wordAssembly `
        -DllPath $wordDllPath `
        -OfficeApp "Word" `
        -FriendlyName "SMILES for Word" `
        -Description "SMILES 화학 구조식 렌더러 및 Ketcher 에디터"
}

# Register PowerPoint COM Add-in
$pptDllPath = Join-Path $suiteDir 'Smiles2Img.PowerPoint.dll'
if (Test-Path $pptDllPath) {
    $pptAssembly = try { [System.Reflection.AssemblyName]::GetAssemblyName($pptDllPath).FullName } catch { "Smiles2Img.PowerPoint, Version=1.5.0.0, Culture=neutral, PublicKeyToken=null" }
    Register-OfficeComAddIn `
        -ProgId "Smiles2Img.PowerPoint.AddIn" `
        -Clsid "{A755973C-21FB-494A-A3C5-086AFFB7C2AC}" `
        -ClassName "Smiles2Img.PowerPoint.PowerPointAddIn" `
        -AssemblyString $pptAssembly `
        -DllPath $pptDllPath `
        -OfficeApp "PowerPoint" `
        -FriendlyName "SMILES for PowerPoint" `
        -Description "SMILES 화학 구조식 렌더러 및 Ketcher 에디터"
}

# Register Windows Toast Notification AUMID
try {
    $aumidKey = "HKCU:\Software\Classes\AppUserModelId\func_SMILES2IMG"
    New-Item -Path $aumidKey -Force | Out-Null
    Set-ItemProperty -Path $aumidKey -Name "DisplayName" -Value "func_SMILES2IMG"
} catch { }

# -------------------------------------------------------------
# 6. Success Output
# -------------------------------------------------------------
Write-Host "`n[SUCCESS] SMILES for Microsoft Office Suite has been successfully installed!" -ForegroundColor Cyan
Write-Host "  * Excel:      Type =SMILES2IMG(""CCO"") in any cell" -ForegroundColor White
Write-Host "  * Word:       Click 'Insert Structure' in the [SMILES] ribbon tab" -ForegroundColor White
Write-Host "  * PowerPoint: Click 'Insert Structure' in the [SMILES] ribbon tab" -ForegroundColor White
Write-Host "  * All apps:   Double-click any molecule image to edit in Ketcher!" -ForegroundColor Yellow

# -------------------------------------------------------------
# 7. Check Windows 11 Smart App Control (SAC)
# -------------------------------------------------------------
$sacPolicy = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy" -ErrorAction SilentlyContinue
if ($sacPolicy -and $sacPolicy.VerifiedAndReputablePolicyState -eq 1) {
    Write-Host "`n----------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[!] Notice: Windows 11 Smart App Control (SAC) is enabled on this system." -ForegroundColor Yellow
    Write-Host "    If Excel shows '#NAME?' or 'file format and extension do not match' warnings:" -ForegroundColor Gray
    Write-Host "    1. Open Windows Security > App & browser control > Smart App Control settings" -ForegroundColor Gray
    Write-Host "    2. Set it to 'Off' (Real-time Microsoft Defender Antivirus remains 100% active)" -ForegroundColor Gray
    Write-Host "    3. Restart Office to use SMILES2IMG normally." -ForegroundColor Gray
    Write-Host "----------------------------------------------------------------------" -ForegroundColor Yellow
}
