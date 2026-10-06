$filePath = (Resolve-Path ".\SMILES2IMG_Comprehensive_Showcase.xlsx").Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open($filePath)
    Write-Host "Total Sheets: $($wb.Worksheets.Count)"
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
        $ws = $wb.Worksheets.Item($i)
        $shapes = $ws.Shapes.Count
        $usedRows = $ws.UsedRange.Rows.Count
        Write-Host "Sheet $($i): $($ws.Name) | Shapes: $shapes | UsedRows: $usedRows"
    }

    # Check sample cell values in Sheet 1
    $s1 = $wb.Worksheets.Item("SMILES2IMG (In-Cell 365)")
    Write-Host "`n--- Sheet 1 Samples ---"
    Write-Host "G4 Formula: $($s1.Range('G4').Formula)"
    Write-Host "G4 Text/Value: $($s1.Range('G4').Text)"
    Write-Host "G29 (Ultimate Full Option) Formula: $($s1.Range('G29').Formula)"
    Write-Host "G29 Text/Value: $($s1.Range('G29').Text)"

    # Check Sheet 3 Samples
    $s3 = $wb.Worksheets.Item("Complex Natural Products")
    Write-Host "`n--- Sheet 3 Samples ---"
    Write-Host "G4 (Artemisinin) Formula: $($s3.Range('G4').Formula)"
    Write-Host "G5 (Paclitaxel) Formula: $($s3.Range('G5').Formula)"

    # Check Sheet 4 Samples
    $s4 = $wb.Worksheets.Item("Style & Transform Matrix")
    Write-Host "`n--- Sheet 4 Samples ---"
    Write-Host "F4 Formula: $($s4.Range('F4').Formula)"
    Write-Host "F17 (Full option pipe) Formula: $($s4.Range('F17').Formula)"

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
