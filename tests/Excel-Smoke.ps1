param(
    [string]$AddInPath = (Join-Path $PSScriptRoot '../dist/x64/Smiles2Img-AddIn64-packed.xll')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelSmokeWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
}
'@
$resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath).Path
$excel = $null
$workbook = $null
$sheet = $null
$ownedExcelId = 0
$reopenPath = $null
$existingExcelIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

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
function Wait-ForResult([scriptblock]$condition) {
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 200
        if (Invoke-Excel $condition) { return }
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Result timed out. B2: $($sheet.Range('B2').Text); shapes: $($sheet.Shapes.Count)"
}
function Wait-ForImage {
    Wait-ForResult { $sheet.Shapes.Count -eq 0 -and (Get-ImageMetadata 'B2') -ne '' }
}
function Get-ImageMetadata([string]$address) {
    $snapshot = Join-Path ([IO.Path]::GetTempPath()) ('smiles2img-smoke-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
    try {
        Invoke-Excel { $workbook.SaveCopyAs($snapshot) }
        $archive = [IO.Compression.ZipFile]::OpenRead($snapshot)
        try {
            $reader = [IO.StreamReader]::new($archive.GetEntry('xl/worksheets/sheet1.xml').Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $node = $xml.SelectSingleNode("//*[local-name()='c' and @r='$address']")
            if ($null -eq $node) { return '' }
            $vm = $node.GetAttribute('vm')
            if ($vm -eq '') { return '' }
            function Read-PackageXml([string]$name) {
                $xmlReader = [IO.StreamReader]::new($archive.GetEntry($name).Open())
                try { return [xml]$xmlReader.ReadToEnd() } finally { $xmlReader.Dispose() }
            }
            $metadata = Read-PackageXml 'xl/metadata.xml'
            $valueRecords = $metadata.SelectNodes("//*[local-name()='valueMetadata']/*[local-name()='bk']")
            $futureIndex = [int]$valueRecords[[int]$vm - 1].FirstChild.GetAttribute('v')
            $futureRecords = $metadata.SelectNodes("//*[local-name()='futureMetadata' and @name='XLRICHVALUE']/*[local-name()='bk']")
            $richIndex = [int]$futureRecords[$futureIndex].SelectSingleNode(".//*[local-name()='rvb']").GetAttribute('i')
            $values = Read-PackageXml 'xl/richData/rdrichvalue.xml'
            $relationshipIndex = [int]$values.SelectNodes("/*/*[local-name()='rv']")[$richIndex].FirstChild.InnerText
            $relationships = Read-PackageXml 'xl/richData/richValueRel.xml'
            $id = $relationships.DocumentElement.ChildNodes[$relationshipIndex].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $imageRelationships = Read-PackageXml 'xl/richData/_rels/richValueRel.xml.rels'
            $imagePath = $imageRelationships.SelectSingleNode("//*[@Id='$id']").GetAttribute('Target')
            $entryPath = ([Uri]::new([Uri]'https://fixture.invalid/xl/richData/', $imagePath)).AbsolutePath.TrimStart('/')
            $pngStream = $archive.GetEntry($entryPath).Open()
            try {
                $header = New-Object byte[] 24
                $offset = 0
                while ($offset -lt $header.Length) {
                    $read = $pngStream.Read($header, $offset, $header.Length - $offset)
                    if ($read -eq 0) { throw 'Incomplete PNG header.' }
                    $offset += $read
                }
                $width = [int]$header[16] * 16777216 + [int]$header[17] * 65536 + [int]$header[18] * 256 + [int]$header[19]
                $height = [int]$header[20] * 16777216 + [int]$header[21] * 65536 + [int]$header[22] * 256 + [int]$header[23]
                if ($width -lt 200 -or $height -lt 100) { throw "Unexpected embedded PNG resolution: ${width}x${height}" }
            } finally { $pngStream.Dispose() }
            $imageStream = $archive.GetEntry($entryPath).Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try { return -join ($sha.ComputeHash($imageStream) | ForEach-Object { $_.ToString('x2') }) }
            finally { $sha.Dispose(); $imageStream.Dispose() }
        } finally { $archive.Dispose() }
    } finally { if (Test-Path -LiteralPath $snapshot) { Remove-Item -LiteralPath $snapshot } }
}
try {
    # Use a separate Excel instance and a new workbook; never attach to user documents.
    $excel = New-Object -ComObject Excel.Application
    [uint32]$instanceId = 0
    [void][ExcelSmokeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$instanceId)
    if ($existingExcelIds -contains [int]$instanceId) { throw 'Refusing to use an existing Excel process.' }
    $ownedExcelId = [int]$instanceId
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    if (-not $excel.RegisterXLL($resolvedAddIn)) { throw 'Excel could not register the XLL.' }
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    [void]$sheet.Range('E5').Select()
    $sheet.Range('B2').RowHeight = 120
    $sheet.Range('B2').ColumnWidth = 40
    $sheet.Range('A2').Value2 = 'CCO'
    $sheet.Range('B2').Formula = '=SMILES2IMG(A2)'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForImage
    if (-not $sheet.Range('B2').HasFormula) { throw 'The image replaced the formula.' }
    if ($excel.ActiveSheet.Name -ne $sheet.Name -or $excel.Selection.Address() -ne '$E$5') { throw 'User selection was not restored.' }
    $firstImage = Get-ImageMetadata 'B2'
    if ($firstImage -eq '') { throw 'B2 is not an actual in-cell image.' }
    $cache = $workbook.Worksheets.Item(2)
    if ($cache.Visible -ne 2 -or $cache.Cells.Item(2, 3).Value2 -ne 'CCO') { throw 'Image cache is not private or contains the wrong molecule.' }
    Write-Output 'PASS SMILES2IMG registration, image output, formula and selection preservation'

    $sheet.Range('A2').Value2 = 'c1ccccc1'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForImage
    Wait-ForResult { $cache.Cells.Item(3, 3).Value2 -eq 'c1ccccc1' }
    $nextImage = Get-ImageMetadata 'B2'
    if ($nextImage -eq '' -or $nextImage -eq $firstImage) { throw 'In-cell image was not updated.' }
    Write-Output 'PASS molecule update without floating or duplicate pictures'

    $sheet.Range('A2').Value2 = 'not-a-smiles'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { $sheet.Shapes.Count -eq 0 -and $sheet.Range('B2').Text -eq '#VALUE!' }
    if ((Get-ImageMetadata 'B2') -ne '') { throw 'Stale image metadata remains.' }
    Write-Output 'PASS invalid input and stale image removal'

    $sheet.Range('A2').Value2 = 'CCO'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForImage
    $sheet.Range('C2').Formula = '=SMILES2IMG(A2)'
    Invoke-Excel { $excel.Calculate() }
    if ((Get-ImageMetadata 'C2') -eq '') { throw 'Repeated SMILES did not return an in-cell image.' }
    if ($cache.UsedRange.Rows.Count -ne 3) { throw 'Repeated SMILES created duplicate cache images.' }
    $sheet.Range('C2').ClearContents()
    Write-Output 'PASS native image reference and cache reuse in another cell'
    $reopenPath = Join-Path ([IO.Path]::GetTempPath()) ('smiles2img reopen ' + [Guid]::NewGuid().ToString('N') + '.xlsx')
    Invoke-Excel { $workbook.SaveAs($reopenPath, 51) }
    Invoke-Excel { $workbook.Close($false) }
    $workbook = $excel.Workbooks.Open($reopenPath, 0)
    $sheet = $workbook.Worksheets.Item(1)
    $cache = $workbook.Worksheets.Item(2)
    Invoke-Excel { $excel.CalculateFullRebuild() }
    Wait-ForImage
    if (-not $sheet.Range('B2').HasFormula -or (Get-ImageMetadata 'B2') -ne $firstImage) { throw 'Saved image or formula did not survive reopening.' }
    if ($workbook.Worksheets.Count -ne 2 -or $cache.UsedRange.Rows.Count -ne 3) { throw 'Reopening created a duplicate cache.' }
    Write-Output 'PASS saving, reopening and reuse of the existing image cache'
    $sheet.Range('B2').ClearContents()
    Wait-ForResult { $sheet.Shapes.Count -eq 0 -and $sheet.Range('B2').Value2 -eq $null }
    if ((Get-ImageMetadata 'B2') -ne '') { throw 'Image remains after deleting the formula.' }
    Write-Output 'PASS image cleanup when the formula is deleted'

    # Verify SMILES2IMG.FLOAT (Excel 2016/2019/2021 compatible floating shape mode)
    $sheet.Range('A2').Value2 = 'CCO'
    $sheet.Range('B2').Formula = '=SMILES2IMG.FLOAT(A2)'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { $sheet.Shapes.Count -eq 1 }
    $floatShape = $sheet.Shapes.Item(1)
    if ($floatShape.Placement -ne 1) { throw 'Floating picture is not xlMoveAndSize (Placement 1).' }
    if (-not $floatShape.Title.StartsWith('SMILES2IMG.FLOAT|')) { throw 'Floating picture Title metadata missing.' }
    if ($floatShape.AlternativeText -ne 'CCO') { throw 'Floating picture alt text does not match molecule.' }
    Write-Output 'PASS SMILES2IMG.FLOAT shape creation and xlMoveAndSize placement'

    $sheet.Range('A2').Value2 = 'c1ccccc1'
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { $sheet.Shapes.Count -eq 1 -and $sheet.Shapes.Item(1).AlternativeText -eq 'c1ccccc1' }
    Write-Output 'PASS SMILES2IMG.FLOAT molecule update and shape replacement'

    $sheet.Range('B2').ClearContents()
    Invoke-Excel { $excel.Calculate() }
    Wait-ForResult { $sheet.Shapes.Count -eq 0 }
    Write-Output 'PASS SMILES2IMG.FLOAT shape cleanup on formula deletion'
} finally {
    try { if ($workbook -ne $null) { Invoke-Excel { $workbook.Close($false) } } }
    finally { if ($excel -ne $null) { Invoke-Excel { $excel.Quit() } } }
    foreach ($ownedObject in @($sheet, $workbook, $excel)) {
        if ($ownedObject -ne $null -and [Runtime.InteropServices.Marshal]::IsComObject($ownedObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ownedObject)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($ownedExcelId -ne 0) {
        $ownedExcel = Get-Process -Id $ownedExcelId -ErrorAction SilentlyContinue
        if ($null -ne $ownedExcel -and -not $ownedExcel.WaitForExit(2000) -and $ownedExcel.MainWindowTitle -eq '') {
            Stop-Process -Id $ownedExcelId
        }
    }
    if ($reopenPath -ne $null -and (Test-Path -LiteralPath $reopenPath)) { Remove-Item -LiteralPath $reopenPath }
}
