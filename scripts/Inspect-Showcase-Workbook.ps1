$filePath = (Resolve-Path ".\examples\func_SMILES2IMG_Excel_Playbook.xlsx").Path
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

    # Check sample cell values in Sheet 1 (0. Start Here)
    $s1 = $wb.Worksheets.Item("0. Start Here")
    Write-Host "`n--- Sheet 1 (0. Start Here) Samples ---"
    Write-Host "E6 (Name2Img) Formula: $($s1.Range('E6').Formula)"
    Write-Host "E8 (SMILES2Img) Formula: $($s1.Range('E8').Formula)"

    # Check Sheet 2 (1. Structure Rendering)
    $s2 = $wb.Worksheets.Item("1. Structure Rendering")
    Write-Host "`n--- Sheet 2 (1. Structure Rendering) Samples ---"
    Write-Host "E4 (Name2Img Aspirin) Formula: $($s2.Range('E4').Formula)"
    Write-Host "F4 (Name2Smiles Aspirin) Value: $($s2.Range('F4').Text)"

    # Check Sheet 3 (2. Properties & 3D)
    $s3 = $wb.Worksheets.Item("2. Properties & 3D")
    Write-Host "`n--- Sheet 3 (2. Properties & 3D) Samples ---"
    Write-Host "E4 (Formula): $($s3.Range('E4').Text) | F4 (MW): $($s3.Range('F4').Text) | H4 (LogP): $($s3.Range('H4').Text) | I4 (TPSA): $($s3.Range('I4').Text)"
    Write-Host "E15 (Paclitaxel Formula): $($s3.Range('E15').Text) | H15 (LogP): $($s3.Range('H15').Text) | I15 (TPSA): $($s3.Range('I15').Text)"

    # Check Sheet 4 (3. SAR & SMARTS)
    $s4 = $wb.Worksheets.Item("3. SAR & SMARTS")
    Write-Host "`n--- Sheet 4 (3. SAR & SMARTS) Samples ---"
    Write-Host "D5 (Aspirin Match): $($s4.Range('D5').Text) | F5 (InChI Match): $($s4.Range('F5').Text)"
    Write-Host "D11 (Pyridine Match): $($s4.Range('D11').Text)"

    # Check Sheet 6 (5. Styles & Transforms)
    $s5 = $wb.Worksheets.Item("5. Styles & Transforms")
    Write-Host "`n--- Sheet 6 (5. Styles & Transforms) Samples ---"
    Write-Host "F4 (Default Formula): $($s5.Range('F4').Formula)"
    Write-Host "F6 (B&W Formula): $($s5.Range('F6').Formula)"
    Write-Host "F19 (All-in-One Formula): $($s5.Range('F19').Formula)"

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
