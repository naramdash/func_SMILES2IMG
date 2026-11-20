<#
.SYNOPSIS
    Configures Microsoft Office UI preferences for English (en-US / 1033) and,
    when executed as Administrator, downloads & installs the Microsoft 365 en-US language pack via ODT.
#>
param(
    [switch]$RegistryOnly
)

$ErrorActionPreference = 'Stop'

Write-Host "=== func_SMILES2IMG Office English (en-US) Setup ===" -ForegroundColor Cyan

# 1. Set HKCU Office UI Language preferences (no admin needed)
$langReg = 'HKCU:\Software\Microsoft\Office\16.0\Common\LanguageResources'
if (-not (Test-Path $langReg)) { New-Item -Path $langReg -Force | Out-Null }
Set-ItemProperty -Path $langReg -Name 'UILanguage' -Value 1033 -Type DWord
Set-ItemProperty -Path $langReg -Name 'HelpLanguage' -Value 1033 -Type DWord
Set-ItemProperty -Path $langReg -Name 'InstallLanguage' -Value 1033 -Type DWord
Set-ItemProperty -Path $langReg -Name 'UILanguageTag' -Value 'en-us' -Type String
Set-ItemProperty -Path $langReg -Name 'HelpLanguageTag' -Value 'en-us' -Type String
Set-ItemProperty -Path $langReg -Name 'PreferredUILanguage' -Value 'en-US' -Type String
Set-ItemProperty -Path $langReg -Name 'PreferredEditingLanguage' -Value 'en-US' -Type String
# Note: in Microsoft 365, 'x-none' means "Match Microsoft Windows"; 'en-us' MUST come first!
Set-ItemProperty -Path $langReg -Name 'UIFallbackLanguages' -Value 'en-us;x-none;ko-kr' -Type String
Set-ItemProperty -Path $langReg -Name 'HelpFallbackLanguages' -Value 'en-us;x-none;ko-kr' -Type String
Set-ItemProperty -Path $langReg -Name 'UISnapshotLanguages' -Value 'en-us' -Type String
Set-ItemProperty -Path $langReg -Name 'FollowSystemUILanguage' -Value 0 -Type DWord
Set-ItemProperty -Path $langReg -Name 'FollowSystemUI' -Value 0 -Type DWord
Write-Host "[+] Updated HKCU Office LanguageResources to en-US (1033)" -ForegroundColor Green

$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$officeRoot = 'C:\Program Files\Microsoft Office\Root\Office16\1033'
$hasFullEnUs = (Test-Path (Join-Path $officeRoot 'WWLIB.DLL')) -or (Test-Path (Join-Path $officeRoot 'XLINTL32.DLL'))
if ($hasFullEnUs) {
    Write-Host "[+] Full Office en-US UI resources are already installed in $officeRoot" -ForegroundColor Green
    return
}

if ($RegistryOnly) {
    Write-Warning "Full Office en-US Click-to-Run pack is not yet installed in $officeRoot. Run this script as Administrator without -RegistryOnly to install it."
    return
}

# 2. Check elevation for Click-to-Run language pack installation
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Installing Microsoft 365 en-US language pack requires an elevated Administrator PowerShell window. Right-click PowerShell -> Run as Administrator, then run: powershell -ExecutionPolicy Bypass -File scripts\Install-OfficeEnglishPack.ps1"
}

$arch = if ($c2r -and $c2r.Platform -eq 'x86') { '32' } else { '64' }
$prodIds = if ($c2r -and $c2r.ProductReleaseIds) { ($c2r.ProductReleaseIds -split ',')[0].Trim() } else { 'O365HomePremRetail' }

$workDir = Join-Path $env:TEMP 'odt_en_us'
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

Write-Host "[+] Installing Office Deployment Tool via winget or Microsoft download..." -ForegroundColor Cyan
$odtExe = Get-Command setup.exe -ErrorAction SilentlyContinue | Where-Object { $_.Source -match 'OfficeDeploymentTool|odt' } | Select-Object -First 1 -ExpandProperty Source
if (-not $odtExe) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        winget install --id Microsoft.OfficeDeploymentTool --silent --accept-package-agreements --accept-source-agreements | Out-Null
    }
    $odtExe = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages","C:\Program Files*","C:\ODT*" -Recurse -Filter setup.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $odtExe) {
    throw "Could not locate Office Deployment Tool setup.exe. Download from https://www.microsoft.com/en-us/download/details.aspx?id=49117 and extract to C:\ODT."
}

$xmlPath = Join-Path $workDir 'install-en-us.xml'
$xml = @"
<Configuration>
  <Add OfficeClientEdition="$arch" Channel="Current">
    <Product ID="$prodIds">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
[System.IO.File]::WriteAllText($xmlPath, $xml, [System.Text.Encoding]::ASCII)

Write-Host "[+] Running ODT to add en-us language pack to $prodIds ($arch-bit)..." -ForegroundColor Cyan
$proc = Start-Process -FilePath $odtExe -ArgumentList "/configure `"$xmlPath`"" -Wait -PassThru
if ($proc.ExitCode -ne 0) {
    throw "Office Deployment Tool exited with code $($proc.ExitCode)."
}
Write-Host "[+] English (en-us) Office UI pack installed successfully! Now re-run scripts\Capture-LiveDocsMedia.ps1." -ForegroundColor Green
