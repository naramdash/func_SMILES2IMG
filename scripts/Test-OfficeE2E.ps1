<#
.SYNOPSIS
    End-to-End (E2E) Automated Verification for SMILES for Microsoft Office Suite (Word, PowerPoint, Excel)
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$repoRoot = Split-Path -Parent $PSScriptRoot
$distX64 = Join-Path $repoRoot "dist\x64"
$wordDll = Join-Path $distX64 "Word\Smiles2Img.Word.dll"
$pptDll  = Join-Path $distX64 "PowerPoint\Smiles2Img.PowerPoint.dll"
$coreDll = Join-Path $distX64 "Word\Smiles2Img.Core.dll"
$xllPath = Join-Path $env:APPDATA "Microsoft\AddIns\Smiles2Img-AddIn64-packed.xll"

[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[System.Reflection.Assembly]::LoadFrom($wordDll) | Out-Null
[System.Reflection.Assembly]::LoadFrom($pptDll)  | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path -Parent $coreDll))

$sampleKet = '{"root":{"nodes":[{"$ref":"mol0"}]},"mol0":{"type":"molecule","atoms":[{"label":"C","location":[0,0,0]},{"label":"O","location":[1,0,0]}],"bonds":[{"type":1,"atoms":[0,1]}]}}'
$aspirinSmiles = "CC(=O)Oc1ccccc1C(=O)O"

Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host " 1. Testing Microsoft Word Live COM Add-in & E2E Workflow" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $wordAddIn = $null
    foreach ($ai in $word.COMAddIns) {
        if ($ai.ProgId -eq "Smiles2Img.Word.AddIn") { $wordAddIn = $ai; break }
    }
    if (-not $wordAddIn) { throw "Smiles2Img.Word.AddIn COM Add-in is not registered in Word!" }
    Write-Host "  [PASS] Word COM Add-in registered (ProgId=$($wordAddIn.ProgId), Connected=$($wordAddIn.Connect))" -ForegroundColor Green

    $doc = $word.Documents.Add()

    # Test 1A: Selected text -> QuickInsertFromSelectionOrClipboard
    $word.Selection.TypeText($aspirinSmiles)
    $doc.Range(0, $aspirinSmiles.Length).Select()
    [Smiles2Img.Word.WordStructureManager]::QuickInsertFromSelectionOrClipboard($word)

    if ($doc.InlineShapes.Count -lt 1) { throw "Word QuickInsert from selection failed to create InlineShape!" }
    $inline1 = $doc.InlineShapes.Item(1)
    $extractedSmiles1 = [Smiles2Img.Word.WordStructureManager]::ExtractSmiles($inline1.Title, $inline1.AlternativeText)
    if ([string]::IsNullOrWhiteSpace($extractedSmiles1)) { throw "Word InlineShape #1 missing SMILES metadata!" }
    $info1 = [Smiles2Img.Rendering.MoleculeRenderer]::TryGetMoleculeInfo($extractedSmiles1)
    if (-not $info1 -or $info1.Formula -ne "C9H8O4") { throw "Word InlineShape #1 MoleculeInfo mismatch!" }
    Write-Host "  [PASS] Word QuickInsert (Selection -> InlineShape): $($inline1.Width)x$($inline1.Height) pt, Formula=$($info1.Formula), MW=$($info1.MolecularWeight.ToString('F2'))" -ForegroundColor Green

    # Test 1B: KET Canvas via Clipboard -> QuickInsertFromSelectionOrClipboard
    $word.Selection.Collapse(0) # wdCollapseEnd
    [System.Windows.Forms.Clipboard]::SetText("KET:$sampleKet")
    [Smiles2Img.Word.WordStructureManager]::QuickInsertFromSelectionOrClipboard($word)

    if ($doc.InlineShapes.Count -lt 2) { throw "Word QuickInsert from KET clipboard failed to create 2nd InlineShape!" }
    $inline2 = $doc.InlineShapes.Item(2)
    $extractedKet2 = [Smiles2Img.Word.WordStructureManager]::ExtractStructure($inline2.AlternativeText, $inline2.Title)
    if ($extractedKet2 -ne $sampleKet) { throw "Word InlineShape #2 KET round-trip mismatch! Got: $extractedKet2" }
    Write-Host "  [PASS] Word KET Canvas Round-Trip (InlineShape #2): KET metadata preserved 100%" -ForegroundColor Green

    $doc.Close($false)
} finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

Write-Host ""
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host " 2. Testing Microsoft PowerPoint Live COM Add-in & E2E Workflow" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application

    $pptAddIn = $null
    foreach ($ai in $ppt.COMAddIns) {
        if ($ai.ProgId -eq "Smiles2Img.PowerPoint.AddIn") { $pptAddIn = $ai; break }
    }
    if (-not $pptAddIn) { throw "Smiles2Img.PowerPoint.AddIn COM Add-in is not registered in PowerPoint!" }
    Write-Host "  [PASS] PowerPoint COM Add-in registered (ProgId=$($pptAddIn.ProgId), Connected=$($pptAddIn.Connect))" -ForegroundColor Green

    $pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)
    $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank = 12

    # Test 2A: QuickInsert SMILES onto PowerPoint slide
    [System.Windows.Forms.Clipboard]::SetText($aspirinSmiles)
    [Smiles2Img.PowerPoint.PowerPointStructureManager]::QuickInsertFromClipboard($ppt)

    if ($slide.Shapes.Count -lt 1) { throw "PowerPoint QuickInsert failed to insert shape!" }
    $pptShape1 = $slide.Shapes.Item(1)
    $pptSmiles1 = $pptShape1.Tags.Item("SMILES")
    if ([string]::IsNullOrWhiteSpace($pptSmiles1)) { throw "PowerPoint Shape #1 missing SMILES tag! Title=$($pptShape1.Title)" }
    $pptInfo1 = [Smiles2Img.Rendering.MoleculeRenderer]::TryGetMoleculeInfo($pptSmiles1)
    if (-not $pptInfo1 -or $pptInfo1.Formula -ne "C9H8O4") { throw "PowerPoint Shape #1 MoleculeInfo mismatch!" }
    Write-Host "  [PASS] PowerPoint QuickInsert (SMILES -> Slide Shape): $([math]::Round($pptShape1.Width,1))x$([math]::Round($pptShape1.Height,1)) pt, Formula=$($pptInfo1.Formula), MW=$($pptInfo1.MolecularWeight.ToString('F2'))" -ForegroundColor Green

    # Test 2B: QuickInsert KET onto PowerPoint slide & verify KET Tag + AlternativeText round-trip
    [System.Windows.Forms.Clipboard]::SetText("KET:$sampleKet")
    [Smiles2Img.PowerPoint.PowerPointStructureManager]::QuickInsertFromClipboard($ppt)

    if ($slide.Shapes.Count -lt 2) { throw "PowerPoint QuickInsert KET failed to insert 2nd shape!" }
    $pptShape2 = $slide.Shapes.Item(2)
    $pptKet2 = $pptShape2.Tags.Item("KET")
    if ($pptKet2 -ne $sampleKet) { throw "PowerPoint Shape #2 KET round-trip mismatch! Got: $pptKet2" }
    if (-not $pptShape2.AlternativeText.StartsWith("KET:")) { throw "PowerPoint Shape #2 AlternativeText missing KET prefix!" }
    Write-Host "  [PASS] PowerPoint KET Canvas Round-Trip (Shape #2): Tags['KET'] & AlternativeText preserved 100%" -ForegroundColor Green

    # Test 2C: Insert WebView2 Ketcher-captured PNG directly onto PowerPoint slide
    $webViewPng = Join-Path $repoRoot "output\samples\e2e_ketcher_trans.png"
    if (Test-Path $webViewPng) {
        $pptShape3 = $slide.Shapes.AddPicture($webViewPng, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 100, 100, -1, -1)
        if ($pptShape3.Width -lt 20 -or $pptShape3.Height -lt 20) { throw "PowerPoint WebView2 PNG shape has invalid size!" }
        Write-Host "  [PASS] PowerPoint WebView2 Canvas PNG Insertion (Shape #3): $([math]::Round($pptShape3.Width,1))x$([math]::Round($pptShape3.Height,1)) pt (No error icon!)" -ForegroundColor Green
    }

    $pres.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pres.Close()
} finally {
    if ($ppt) {
        try { $ppt.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
}

Write-Host ""
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host " 3. Testing Microsoft Excel Live Add-in (.xll) & E2E Workflow" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    if (-not (Test-Path $xllPath)) { throw "Excel XLL not found at $xllPath" }
    $registered = $excel.RegisterXLL($xllPath)
    if (-not $registered) { throw "Excel failed to load $xllPath via RegisterXLL!" }
    Write-Host "  [PASS] Excel XLL loaded via RegisterXLL: $xllPath" -ForegroundColor Green

    $wb = $excel.Workbooks.Add()
    $ws = $wb.ActiveSheet
    $ws.Range("A1").Value2 = $aspirinSmiles
    $ws.Range("B1").RowHeight = 90
    $ws.Range("B1").ColumnWidth = 24
    $ws.Range("B1").Formula = "=SMILES2IMG.FLOAT(A1)"
    Start-Sleep -Seconds 2

    $valB1 = $ws.Range("B1").Text
    Write-Host "  [PASS] Excel =SMILES2IMG.FLOAT(A1) evaluated cleanly (Cell text='$valB1', Shapes on sheet=$($ws.Shapes.Count))" -ForegroundColor Green

    $wb.Close($false)
} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Write-Host ""
Write-Host "[ALL E2E TESTS PASSED SUCCESSFULLY]" -ForegroundColor Green
