$ErrorActionPreference = 'Stop'

Write-Host "=== Installing SMILES2IMG Excel Add-In on Current System ===" -ForegroundColor Cyan

# 1. Check Running Excel Processes
$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "Closing running Excel instances..." -ForegroundColor Yellow
    $excelProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
}

# 2. Determine Excel Bitness via COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$is64Bit = $false
try {
    # Check OperatingSystem / Bitness
    $proc = Get-Process -Id ([System.Diagnostics.Process]::GetCurrentProcess().Id)
    # Check Excel process bitness
    $excelProc = Get-Process excel | Select-Object -First 1
    # Check OperatingSystem property if available or check via DLL load
    $addin64 = (Resolve-Path "dist/x64/Smiles2Img-AddIn64-packed.xll").Path
    $addin32 = (Resolve-Path "dist/x86/Smiles2Img-AddIn-packed.xll").Path
    
    $registered64 = $excel.RegisterXLL($addin64)
    if ($registered64) {
        $is64Bit = $true
        Write-Host "Detected Excel Bitness: 64-bit (x64)" -ForegroundColor Green
    } else {
        $registered32 = $excel.RegisterXLL($addin32)
        if ($registered32) {
            $is64Bit = $false
            Write-Host "Detected Excel Bitness: 32-bit (x86)" -ForegroundColor Green
        } else {
            throw "Neither 64-bit nor 32-bit XLL could be registered."
        }
    }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 3. Target Paths
$addInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path -LiteralPath $addInsDir)) {
    New-Item -ItemType Directory -Path $addInsDir -Force | Out-Null
}

$sourceXll = if ($is64Bit) {
    (Resolve-Path "dist/x64/Smiles2Img-AddIn64-packed.xll").Path
} else {
    (Resolve-Path "dist/x86/Smiles2Img-AddIn-packed.xll").Path
}

$destFileName = if ($is64Bit) { "Smiles2Img-AddIn64.xll" } else { "Smiles2Img-AddIn.xll" }
$destPath = Join-Path $addInsDir $destFileName

Write-Host "Copying XLL to user's official Add-Ins folder..." -ForegroundColor Cyan
Write-Host "Source: $sourceXll" -ForegroundColor Gray
Write-Host "Dest  : $destPath" -ForegroundColor Gray

Copy-Item -LiteralPath $sourceXll -Destination $destPath -Force
Unblock-File -LiteralPath $destPath

# 4. Register Permanently in Excel AddIns Collection
Write-Host "`nRegistering Add-In permanently in Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Add to Excel AddIns
    $addin = $excel.AddIns.Add($destPath, $false)
    $addin.Installed = $true
    Write-Host "AddIn '$($addin.Name)' Installed state set to: $($addin.Installed)" -ForegroundColor Green
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 5. Verify Installation in a fresh, clean Excel session
Write-Host "`n=== Verifying Fresh Excel Auto-Load & Function Availability ===" -ForegroundColor Cyan
$testExcel = New-Object -ComObject Excel.Application
$testExcel.Visible = $false
$testExcel.DisplayAlerts = $false

try {
    # Check if add-in is in installed list
    $found = $false
    foreach ($a in $testExcel.AddIns) {
        if ($a.Name -like "Smiles2Img*" -and $a.Installed) {
            Write-Host "Verified installed in Excel AddIns: $($a.Name)" -ForegroundColor Green
            $found = $true
            break
        }
    }
    if (-not $found) {
        Write-Host "Warning: Add-in not explicitly listed as installed, checking function execution directly..." -ForegroundColor Yellow
    }

    # Test executing function in fresh workbook WITHOUT RegisterXLL
    $wb = $testExcel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Range("A1").Value2 = "CCO"
    $ws.Range("B1").Formula2 = "=SMILES2IMG(A1)"
    $testExcel.Calculate()
    Start-Sleep -Seconds 2

    $formula = $ws.Range("B1").Formula2
    $val = $ws.Range("B1").Text
    Write-Host "B1 Formula2  : $formula" -ForegroundColor Gray
    Write-Host "B1 Evaluation: $val" -ForegroundColor Gray

    if ($val -match '^#(NAME\?|REF!)') {
        throw "Function not recognized in fresh Excel session: $val"
    }

    Write-Host "`n🎉 SUCCESS: SMILES2IMG Add-In is permanently installed and verified on this computer!" -ForegroundColor Green
    Write-Host "Installed Path: $destPath" -ForegroundColor Cyan
    $wb.Close($false)
} finally {
    $testExcel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testExcel) | Out-Null
}
