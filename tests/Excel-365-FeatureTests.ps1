param(
    [string]$AddInPath = (Join-Path $PSScriptRoot '../dist/x64/Smiles2Img-AddIn64-packed.xll')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath).Path
$excel = $null
$workbook = $null
$sheet = $null

function Invoke-Excel([scriptblock]$operation) {
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    do {
        try { return (& $operation) }
        catch {
            if ($_.Exception.ToString() -notmatch '80010001|8001010A|800AC472|rejected by callee' -or
                [DateTime]::UtcNow -ge $deadline) { throw }
            Start-Sleep -Milliseconds 200
        }
    } while ($true)
}

function Wait-ForResult([scriptblock]$condition, [string]$failMsg) {
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 250
        if (Invoke-Excel $condition) { return }
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Condition timed out: $failMsg"
}

function Get-CellRichImage([string]$address) {
    $snapshot = Join-Path ([IO.Path]::GetTempPath()) ('smiles2img-feat-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
    try {
        Invoke-Excel { $workbook.SaveCopyAs($snapshot) }
        $archive = [IO.Compression.ZipFile]::OpenRead($snapshot)
        try {
            $sheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
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
            $valueRecords = $metadata.SelectNodes("//*[local-name()='valueMetadata']/*[local-name()='bk']")
            $futureIndex = [int]$valueRecords[[int]$vm - 1].FirstChild.GetAttribute('v')
            $futureRecords = $metadata.SelectNodes("//*[local-name()='futureMetadata' and @name='XLRICHVALUE']/*[local-name()='bk']")
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
    Write-Host "=== Starting Excel Desktop 365 Real Add-In Verification ===" -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Excel Version: $($excel.Version)" -ForegroundColor Gray
    if (-not $excel.RegisterXLL($resolvedAddIn)) {
        throw "Failed to register XLL: $resolvedAddIn"
    }
    Write-Host "XLL registered successfully: $resolvedAddIn" -ForegroundColor Green

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Range('B2:E10').RowHeight = 100
    $sheet.Range('B2:E10').ColumnWidth = 35

    # Test Case 1: Standard In-cell SMILES2IMG with space-separated styles
    Write-Host "`n[Test 1] Space-delimited styles: =SMILES2IMG(A2, 'trans', 'bw num')" -ForegroundColor Yellow
    $sheet.Range('A2').Value2 = 'CC(=O)Oc1ccccc1C(=O)O'  # Aspirin
    $sheet.Range('B2').Formula = '=SMILES2IMG(A2, "trans", "bw num")'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { (Get-CellRichImage 'B2') -ne $null } "In-cell image for B2 (space delimited style)"
    $img1 = Get-CellRichImage 'B2'
    Write-Host "  -> PASS: Generated in-cell image (${($img1.Width)}x${($img1.Height)}, $($img1.Length) bytes)" -ForegroundColor Green

    # Test Case 2: Pipe-delimited styles
    Write-Host "`n[Test 2] Pipe-delimited styles: =SMILES2IMG(A3, , 'bw|stereo|num')" -ForegroundColor Yellow
    $sheet.Range('A3').Value2 = 'C[C@H](O)C(=O)O'  # Lactic acid
    $sheet.Range('B3').Formula = '=SMILES2IMG(A3, , "bw|stereo|num")'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { (Get-CellRichImage 'B3') -ne $null } "In-cell image for B3 (pipe delimited style)"
    $img2 = Get-CellRichImage 'B3'
    Write-Host "  -> PASS: Generated in-cell image (${($img2.Width)}x${($img2.Height)}, $($img2.Length) bytes)" -ForegroundColor Green

    # Test Case 3: The Ultimate Full-Option Formula
    Write-Host "`n[Test 3] Ultimate Full-Option Formula: =SMILES2IMG(A4, 'white', 'bw num stereo h all-c', 'flip rot90')" -ForegroundColor Yellow
    $sheet.Range('A4').Value2 = 'Cn1cnc2c1c(=O)n(C)c(=O)n2C'  # Caffeine
    $sheet.Range('B4').Formula = '=SMILES2IMG(A4, "white", "bw num stereo h all-c", "flip rot90")'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { (Get-CellRichImage 'B4') -ne $null } "In-cell image for B4 (Ultimate Full-Option)"
    $img3 = Get-CellRichImage 'B4'
    Write-Host "  -> PASS: Generated in-cell image (${($img3.Width)}x${($img3.Height)}, $($img3.Length) bytes)" -ForegroundColor Green

    # Test Case 4: Full-Option Dark Mode
    Write-Host "`n[Test 4] Full-Option Dark Mode: =SMILES2IMG(A5, '#1a1a1a', 'white stereo num h', 'flip')" -ForegroundColor Yellow
    $sheet.Range('A5').Value2 = 'C1=CC=C2C(=C1)C(=O)NS2(=O)=O'  # Saccharin
    $sheet.Range('B5').Formula = '=SMILES2IMG(A5, "#1a1a1a", "white stereo num h", "flip")'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { (Get-CellRichImage 'B5') -ne $null } "In-cell image for B5 (Dark mode full option)"
    $img4 = Get-CellRichImage 'B5'
    Write-Host "  -> PASS: Generated in-cell image (${($img4.Width)}x${($img4.Height)}, $($img4.Length) bytes)" -ForegroundColor Green

    # Test Case 5: SMILES2IMG.FLOAT with Ultimate Full-Option Formula
    Write-Host "`n[Test 5] SMILES2IMG.FLOAT with Full-Option Formula" -ForegroundColor Yellow
    $sheet.Range('A6').Value2 = 'CC(=O)NC1=CC=C(O)C=C1'  # Acetaminophen
    $sheet.Range('C6').Formula = '=SMILES2IMG.FLOAT(A6, "white", "bw num stereo h all-c", "flip rot90")'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { $sheet.Shapes.Count -ge 1 } "Floating shape for C6"
    $shape = $sheet.Shapes.Item(1)
    Write-Host "  -> PASS: Floating shape created (Name: $($shape.Name), Title: $($shape.Title), TopLeftCell: $($shape.TopLeftCell.Address()))" -ForegroundColor Green
    if ($shape.Placement -ne 1) { throw "Expected Placement = xlMoveAndSize (1)" }
    Write-Host "  -> PASS: Shape is correctly configured with xlMoveAndSize (Placement: 1)" -ForegroundColor Green

    # Test Case 6: Save workbook, verify no file corruption, and reload
    Write-Host "`n[Test 6] Saving workbook with embedded molecular rich-data and reopening" -ForegroundColor Yellow
    $testSavePath = Join-Path ([IO.Path]::GetTempPath()) ("Smiles2Img_FullTest_" + [Guid]::NewGuid().ToString('N') + ".xlsx")
    Invoke-Excel { $workbook.SaveAs($testSavePath, 51) }
    Invoke-Excel { $workbook.Close($false) }
    $workbook = $excel.Workbooks.Open($testSavePath, 0)
    $sheet = $workbook.Worksheets.Item(1)
    Invoke-Excel { $excel.CalculateFull() }
    
    # Verify formulas and images survived reload
    if ($sheet.Range('B4').Formula -ne '=SMILES2IMG(A4, "white", "bw num stereo h all-c", "flip rot90")') {
        throw "Formula in B4 did not match after reopen."
    }
    $reloadedImg = Get-CellRichImage 'B4'
    if ($null -eq $reloadedImg) {
        throw "In-cell image did not survive file reopen."
    }
    Write-Host "  -> PASS: Saved workbook successfully reopened with intact formulas & RichValue images!" -ForegroundColor Green

    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host " ALL EXCEL DESKTOP 365 TESTS PASSED SUCCESSFULLY! " -ForegroundColor Green
    Write-Host "=======================================================" -ForegroundColor Cyan
}
finally {
    if ($workbook -ne $null) { try { Invoke-Excel { $workbook.Close($false) } } catch {} }
    if ($excel -ne $null) { try { Invoke-Excel { $excel.Quit() } } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($testSavePath -ne $null -and (Test-Path $testSavePath)) {
        Remove-Item -LiteralPath $testSavePath -Force -ErrorAction SilentlyContinue
    }
}
