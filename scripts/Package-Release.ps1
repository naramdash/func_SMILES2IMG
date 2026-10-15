<#
.SYNOPSIS
    Packages dist output for GitHub Releases
#>
$ErrorActionPreference = 'Stop'

foreach ($arch in @('x64', 'x86')) {
    $distArch = Join-Path $PSScriptRoot "..\dist\$arch"
    if (-not (Test-Path $distArch)) { continue }

    $tempDir = Join-Path $env:TEMP "Smiles2Img-Office-$arch"
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $wordDir = Join-Path $distArch "Word"
    $pptDir  = Join-Path $distArch "PowerPoint"
    $excelDir = Join-Path $PSScriptRoot "..\src\Smiles2Img.Excel\bin\Release\net48"
    $xllName = if ($arch -eq 'x64') { "Smiles2Img-AddIn64.xll" } else { "Smiles2Img-AddIn.xll" }
    $dnaName = if ($arch -eq 'x64') { "Smiles2Img-AddIn64.dna" } else { "Smiles2Img-AddIn.dna" }

    if (Test-Path $wordDir) { Copy-Item -Path "$wordDir\*.dll" -Destination $tempDir -Force }
    if (Test-Path $pptDir)  { Copy-Item -Path "$pptDir\*.dll" -Destination $tempDir -Force }
    if (Test-Path (Join-Path $excelDir $xllName)) {
        Copy-Item (Join-Path $excelDir $xllName) -Destination $tempDir -Force
        Copy-Item (Join-Path $excelDir $dnaName) -Destination $tempDir -Force
        Copy-Item -Path "$excelDir\*.dll" -Destination $tempDir -Force
    }

    $zipPath = Join-Path $distArch "Smiles2Img-Office-$arch.zip"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

    Compress-Archive -Path "$tempDir\*.*" -DestinationPath $zipPath -Force
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "[+] Created $zipPath ($([math]::Round((Get-Item $zipPath).Length / 1MB, 2)) MB)" -ForegroundColor Green
}
