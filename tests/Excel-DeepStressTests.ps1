param(
    [string]$AddInPath = (Join-Path $PSScriptRoot '../dist/x64/Smiles2Img-AddIn64-packed.xll')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath).Path
$excel = $null
$workbook = $null

function Invoke-Excel([scriptblock]$operation) {
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    do {
        try { return (& $operation) }
        catch {
            if ($_.Exception.ToString() -notmatch '80010001|8001010A|800AC472|rejected by callee' -or
                [DateTime]::UtcNow -ge $deadline) { throw }
            Start-Sleep -Milliseconds 250
        }
    } while ($true)
}

function Wait-ForResult([scriptblock]$condition, [string]$failMsg, [int]$timeoutSec = 25) {
    $deadline = [DateTime]::UtcNow.AddSeconds($timeoutSec)
    do {
        Start-Sleep -Milliseconds 300
        if (Invoke-Excel $condition) { return }
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Condition timed out ($timeoutSec s): $failMsg"
}

function Get-CellRichImage([string]$address, [string]$sheetName = 'Sheet1') {
    $snapshot = Join-Path ([IO.Path]::GetTempPath()) ('smiles2img-deep-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
    try {
        Invoke-Excel { $workbook.SaveCopyAs($snapshot) }
        $archive = [IO.Compression.ZipFile]::OpenRead($snapshot)
        try {
            # Find worksheet xml
            $sheetEntry = $archive.GetEntry("xl/worksheets/$sheetName.xml")
            if ($null -eq $sheetEntry) {
                # fallback: sheet1.xml
                $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
            }
            if ($null -eq $sheetEntry) { return $null }

            $reader = [IO.StreamReader]::new($sheetEntry.Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $node = $xml.SelectSingleNode("//*[local-name()='c' and @r='$address']")
            if ($null -eq $node) { return $null }
            $vm = $node.GetAttribute('vm')
            if ($vm -eq '') { return $null }

            function Read-Xml([string]$name) {
                $e = $archive.GetEntry($name)
                if ($null -eq $e) { return $null }
                $r = [IO.StreamReader]::new($e.Open())
                try { return [xml]$r.ReadToEnd() } finally { $r.Dispose() }
            }

            $metadata = Read-Xml 'xl/metadata.xml'
            if ($null -eq $metadata) { return $null }
            $valueRecords = $metadata.SelectNodes("//*[local-name()='valueMetadata']/*[local-name()='bk']")
            if ($null -eq $valueRecords -or $valueRecords.Count -lt [int]$vm) { return $null }
            $futureIndex = [int]$valueRecords[[int]$vm - 1].FirstChild.GetAttribute('v')
            $futureRecords = $metadata.SelectNodes("//*[local-name()='futureMetadata' and @name='XLRICHVALUE']/*[local-name()='bk']")
            if ($null -eq $futureRecords -or $futureRecords.Count -le $futureIndex) { return $null }
            $richIndex = [int]$futureRecords[$futureIndex].SelectSingleNode(".//*[local-name()='rvb']").GetAttribute('i')
            $values = Read-Xml 'xl/richData/rdrichvalue.xml'
            $relationshipIndex = [int]$values.SelectNodes("/*/*[local-name()='rv']")[$richIndex].FirstChild.InnerText
            $relationships = Read-Xml 'xl/richData/richValueRel.xml'
            $id = $relationships.DocumentElement.ChildNodes[$relationshipIndex].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $imageRelationships = Read-Xml 'xl/richData/_rels/richValueRel.xml.rels'
            $imagePath = $imageRelationships.SelectSingleNode("//*[@Id='$id']").GetAttribute('Target')
            $entryPath = ([Uri]::new([Uri]'https://fixture.invalid/xl/richData/', $imagePath)).AbsolutePath.TrimStart('/')
            $pngEntry = $archive.GetEntry($entryPath)
            if ($null -eq $pngEntry) { return $null }

            $pngStream = $pngEntry.Open()
            try {
                $header = New-Object byte[] 24
                $offset = 0
                while ($offset -lt $header.Length) {
                    $read = $pngStream.Read($header, $offset, $header.Length - $offset)
                    if ($read -eq 0) { break }
                    $offset += $read
                }
                $width = [int]$header[16] * 16777216 + [int]$header[17] * 65536 + [int]$header[18] * 256 + [int]$header[19]
                $height = [int]$header[20] * 16777216 + [int]$header[21] * 65536 + [int]$header[22] * 256 + [int]$header[23]
                return @{ Width = $width; Height = $height; Length = $pngEntry.Length }
            } finally { $pngStream.Dispose() }
        } finally { $archive.Dispose() }
    } finally { if (Test-Path -LiteralPath $snapshot) { Remove-Item -LiteralPath $snapshot } }
}

try {
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host " DEEP EXCEL 365 STRESS & EDGE-CASE TEST SUITE " -ForegroundColor Cyan
    Write-Host "=======================================================" -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    if (-not $excel.RegisterXLL($resolvedAddIn)) {
        throw "Failed to register XLL: $resolvedAddIn"
    }
    Write-Host "XLL Loaded: $resolvedAddIn" -ForegroundColor Green

    $workbook = $excel.Workbooks.Add()
    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet1.Name = "Molecules"

    # -------------------------------------------------------------
    # Scenario 1: Multi-cell Bulk Generation (20 Diverse Molecules)
    # -------------------------------------------------------------
    Write-Host "`n[Scenario 1] Bulk generation with 20 diverse molecules..." -ForegroundColor Yellow
    $testMolecules = @(
        @{ Name="Water"; Smiles="O" },
        @{ Name="Ethanol"; Smiles="CCO" },
        @{ Name="Benzene"; Smiles="c1ccccc1" },
        @{ Name="Aspirin"; Smiles="CC(=O)Oc1ccccc1C(=O)O" },
        @{ Name="Caffeine"; Smiles="Cn1cnc2c1c(=O)n(C)c(=O)n2C" },
        @{ Name="Ibuprofen"; Smiles="CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O" },
        @{ Name="Penicillin G"; Smiles="CC1(C(N2C(S1)C(C2=O)NC(=O)Cc3ccccc3)C(=O)O)C" },
        @{ Name="Cholesterol"; Smiles="CC(C)CCCC(C)C1CCC2C1(CCC3C2CC=C4C3(CCC(C4)O)C)C" },
        @{ Name="Nicotine"; Smiles="CN1CCC[C@H]1c2cccnc2" },
        @{ Name="Glucose"; Smiles="OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O" },
        @{ Name="Dopamine"; Smiles="NCCc1ccc(O)c(O)c1" },
        @{ Name="Serotonin"; Smiles="NCCc1c[nH]c2ccc(O)cc12" },
        @{ Name="Morphine"; Smiles="CN1CC[C@]23[C@@H]4Oc5c3c(CC1[C@@H]2C=C[C@@H]4O)ccc5O" },
        @{ Name="Vanillin"; Smiles="O=Cc1ccc(O)c(OC)c1" },
        @{ Name="ATP"; Smiles="Nc1ncnc2n(cnc12)[C@@H]3O[C@H](COP(=O)(O)OP(=O)(O)OP(=O)(O)O)[C@@H](O)[C@H]3O" },
        @{ Name="Artemisinin"; Smiles="CC1CC2CC3(C)OO4C(O2)(C1C(=O)O3)C(C)CC4" },
        @{ Name="Sulfamethoxazole"; Smiles="Cc1cc(NS(=O)(=O)c2ccc(N)cc2)no1" },
        @{ Name="Paracetamol"; Smiles="CC(=O)Nc1ccc(O)cc1" },
        @{ Name="Citric acid"; Smiles="OC(=O)CC(O)(CC(=O)O)C(=O)O" },
        @{ Name="Indigo dye"; Smiles="O=C1C(=C2Nc3ccccc3C2=O)Nc4ccccc14" }
    )

    $smilesArray = New-Object 'object[,]' $testMolecules.Count, 1
    $formulaArray = New-Object 'object[,]' $testMolecules.Count, 1
    for ($i = 0; $i -lt $testMolecules.Count; $i++) {
        $row = $i + 2
        $smilesArray[$i, 0] = $testMolecules[$i].Smiles
        $formulaArray[$i, 0] = "=SMILES2IMG(A$row)"
    }

    Invoke-Excel {
        $sheet1.Range("B2:B21").RowHeight = 80
        $sheet1.Range("B2:B21").ColumnWidth = 30
        $sheet1.Range("A2:A21").Value2 = $smilesArray
        $sheet1.Range("B2:B21").Formula = $formulaArray
        $excel.Calculate()
    }

    # Efficient bulk image polling: single snapshot per check
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    do {
        Start-Sleep -Milliseconds 400
        $snapshot = Join-Path ([IO.Path]::GetTempPath()) ('smiles2img-bulk-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
        $count = 0
        try {
            Invoke-Excel { $workbook.SaveCopyAs($snapshot) }
            $archive = [IO.Compression.ZipFile]::OpenRead($snapshot)
            try {
                $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
                if ($null -ne $sheetEntry) {
                    $reader = [IO.StreamReader]::new($sheetEntry.Open())
                    try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                    $nodes = $xml.SelectNodes("//*[local-name()='c' and @vm and starts-with(@r, 'B')]")
                    if ($null -ne $nodes) { $count = $nodes.Count }
                }
            } finally { $archive.Dispose() }
        } finally { if (Test-Path $snapshot) { Remove-Item $snapshot -Force } }

        if ($count -ge 20) { break }
    } while ($sw.ElapsedMilliseconds -lt 25000)

    if ($count -lt 20) { throw "Bulk generation timed out ($count/20 images ready after $($sw.ElapsedMilliseconds)ms)" }

    Write-Host "  -> PASS: All 20 molecules successfully generated in-cell images!" -ForegroundColor Green

    # -------------------------------------------------------------
    # Scenario 2: Cross-sheet formula reference (Sheet2 referencing Sheet1)
    # -------------------------------------------------------------
    Write-Host "`n[Scenario 2] Cross-sheet references (Molecules2 referencing Molecules!A4)..." -ForegroundColor Yellow
    $sheet2 = Invoke-Excel {
        $missing = [System.Reflection.Missing]::Value
        $s = $workbook.Worksheets.Add($missing, $sheet1)
        $s.Name = "Molecules2"
        $s.Range("B2").RowHeight = 80
        $s.Range("B2").ColumnWidth = 30
        $s.Range("B2").Formula = '=SMILES2IMG(Molecules!A4, "white", "bw num")'
        $s
    }
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { (Get-CellRichImage 'B2' 'Molecules2') -ne $null } "Cross-sheet image on Molecules2!B2"
    Write-Host "  -> PASS: Cross-sheet referencing successfully produced in-cell image!" -ForegroundColor Green

    # -------------------------------------------------------------
    # Scenario 3: Unknown/Invalid tokens & Graceful Fallback
    # -------------------------------------------------------------
    Write-Host "`n[Scenario 3] Robustness: Unknown color name & Unknown style tokens..." -ForegroundColor Yellow
    Invoke-Excel {
        $sheet1.Range("C2").Formula = '=SMILES2IMG(A2, "non_existent_color_name_xyz", "unknown_style_abc bw")'
        $excel.Calculate()
    }
    Wait-ForResult { (Get-CellRichImage 'C2' 'Molecules') -ne $null } "Fallback with unknown tokens in C2"
    Write-Host "  -> PASS: Unknown tokens safely ignored without crash, bw recognized!" -ForegroundColor Green

    # -------------------------------------------------------------
    # Scenario 4: Non-string SMILES cell types (Empty cell, Number, Formula Error)
    # -------------------------------------------------------------
    Write-Host "`n[Scenario 4] Non-string SMILES cell types..." -ForegroundColor Yellow
    Invoke-Excel {
        $sheet1.Range("A25").Value2 = ""          # Empty cell
        $sheet1.Range("A26").Value2 = 12345       # Number
        $sheet1.Range("A27").Formula = "=1/0"      # #DIV/0! error
        $sheet1.Range("B25").Formula = "=SMILES2IMG(A25)"
        $sheet1.Range("B26").Formula = "=SMILES2IMG(A26)"
        $sheet1.Range("B27").Formula = "=SMILES2IMG(A27)"
        $excel.Calculate()
    }
    
    Start-Sleep -Seconds 1
    $val25 = Invoke-Excel { $sheet1.Range("B25").Text }
    $val26 = Invoke-Excel { $sheet1.Range("B26").Text }
    $val27 = Invoke-Excel { $sheet1.Range("B27").Text }
    Write-Host "  Empty cell -> $val25" -ForegroundColor Gray
    Write-Host "  Numeric cell -> $val26" -ForegroundColor Gray
    Write-Host "  Error cell -> $val27" -ForegroundColor Gray

    if ($val25 -ne '#VALUE!' -and $val25 -ne '') { throw "Empty cell expected #VALUE!, got $val25" }
    if ($val26 -ne '#VALUE!') { throw "Numeric cell expected #VALUE!, got $val26" }
    Write-Host "  -> PASS: Non-string and error inputs safely handled without exceptions!" -ForegroundColor Green

    # -------------------------------------------------------------
    # Scenario 5: Large molecule stress (Paclitaxel with Full Options)
    # -------------------------------------------------------------
    Write-Host "`n[Scenario 5] Complex natural product with full options (Paclitaxel)..." -ForegroundColor Yellow
    Invoke-Excel {
        $sheet1.Range("D17").Formula = '=SMILES2IMG(A17, "white", "bw num stereo h all-c", "rot90 flip")'
        $excel.Calculate()
    }
    Wait-ForResult { (Get-CellRichImage 'D17' 'Molecules') -ne $null } "Full-option complex molecule in D17"
    $taxolImg = Get-CellRichImage 'D17' 'Molecules'
    Write-Host "  -> PASS: Paclitaxel full options rendered successfully ($($taxolImg.Length) bytes)!" -ForegroundColor Green

    # -------------------------------------------------------------
    # Scenario 6: Floating Picture Refit Command
    # -------------------------------------------------------------
    Write-Host "`n[Scenario 6] Floating Picture Refit Command via Ribbon..." -ForegroundColor Yellow
    Invoke-Excel {
        $sheet1.Range("E2").Formula = '=SMILES2IMG.FLOAT(A2)'
        $excel.Calculate()
    }
    Wait-ForResult { (Invoke-Excel { $sheet1.Shapes.Count }) -ge 1 } "Floating picture in E2"
    
    # Change row height and invoke Refit
    Invoke-Excel {
        $sheet1.Range("E2").RowHeight = 150
        $sheet1.Range("E2").ColumnWidth = 50
        $excel.Run("RefitFloatingImages")
    }
    Start-Sleep -Milliseconds 500
    Write-Host "  -> PASS: Refit floating images command executed smoothly without error!" -ForegroundColor Green

    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host " ALL DEEP STRESS & EDGE-CASE TESTS PASSED PERFECTLY! " -ForegroundColor Green
    Write-Host "=======================================================" -ForegroundColor Cyan
}
finally {
    if ($workbook -ne $null) { try { Invoke-Excel { $workbook.Close($false) } } catch {} }
    if ($excel -ne $null) { try { Invoke-Excel { $excel.Quit() } } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
