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
        $usedCols = $ws.UsedRange.Columns.Count
        Write-Host "Sheet $($i): $($ws.Name) | Shapes: $shapes | Used: ${usedRows}x${usedCols}"
    }

    # Check sample cell values in Sheet 1 (0. Quick Guide)
    $s1 = $wb.Worksheets.Item("0. Quick Guide")
    Write-Host "`n--- Sheet 1 (0. Quick Guide) Samples ---"
    Write-Host "E4 (Aspirin Name2Img) Formula: $($s1.Range('E4').Formula)"
    Write-Host "E6 (SMILES2Img) Formula: $($s1.Range('E6').Formula)"

    # Check Sheet 2 (1. Structure to Image)
    $s2 = $wb.Worksheets.Item("1. Structure to Image")
    Write-Host "`n--- Sheet 2 (1. Structure to Image) Samples ---"
    Write-Host "E4 (Name2Img Aspirin) Formula: $($s2.Range('E4').Formula)"
    Write-Host "F4 (Name2Smiles Aspirin) Value: $($s2.Range('F4').Text)"

    # Check Sheet 3 (2. Molecular Properties)
    $s3 = $wb.Worksheets.Item("2. Molecular Properties")
    Write-Host "`n--- Sheet 3 (2. Molecular Properties) Samples ---"
    Write-Host "E4 (Formula): $($s3.Range('E4').Text) | F4 (MW): $($s3.Range('F4').Text) | M4 (Lipinski): $($s3.Range('M4').Text)"
    Write-Host "E13 (Paclitaxel Formula): $($s3.Range('E15').Text) | M15 (Paclitaxel Lipinski): $($s3.Range('M15').Text)"

    # Check Sheet 4 (3. Substructure & Similarity)
    $s4 = $wb.Worksheets.Item("3. Substructure & Similarity")
    Write-Host "`n--- Sheet 4 (3. Substructure & Similarity) Samples ---"
    Write-Host "D5 (Aspirin Match): $($s4.Range('D5').Text) | F5 (Aspirin Similarity): $($s4.Range('F5').Text)"
    Write-Host "D8 (Pyridine Match): $($s4.Range('D8').Text)"

    # Check Sheet 5 (4. Drawing Styles & Options)
    $s5 = $wb.Worksheets.Item("4. Drawing Styles & Options")
    Write-Host "`n--- Sheet 5 (4. Drawing Styles & Options) Samples ---"
    Write-Host "F5 (Default Formula): $($s5.Range('F5').Formula)"
    Write-Host "F7 (B&W Formula): $($s5.Range('F7').Formula)"
    Write-Host "F20 (All-in-One Formula): $($s5.Range('F20').Formula)"

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
