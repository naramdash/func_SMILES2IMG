$filePath = (Resolve-Path ".\SMILES2IMG_Comprehensive_Showcase.xlsx").Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open($filePath)
    $errorCount = 0
    Write-Host "Total Worksheets: $($wb.Worksheets.Count)"
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
        $ws = $wb.Worksheets.Item($i)
        $used = $ws.UsedRange
        $sheetErrors = 0
        for ($r = 1; $r -le $used.Rows.Count; $r++) {
            for ($c = 1; $c -le $used.Columns.Count; $c++) {
                $cell = $used.Cells.Item($r, $c)
                $txt = $cell.Text
                if ($txt -match '^#(VALUE!|NAME\?|REF!|DIV/0!|NUM!)') {
                    Write-Host "ERROR at Sheet [$($ws.Name)] Row $r Col $c ($($cell.Address)) : $txt | Formula: $($cell.Formula)" -ForegroundColor Red
                    $errorCount++
                    $sheetErrors++
                }
            }
        }
        Write-Host "Sheet $i [$($ws.Name)]: $($used.Rows.Count) rows, $($used.Columns.Count) cols | Errors: $sheetErrors" -ForegroundColor $(if ($sheetErrors -eq 0) { "Green" } else { "Red" })
    }
    if ($errorCount -eq 0) {
        Write-Host "`nSUCCESS: Zero formula errors across all sheets!" -ForegroundColor Green
    } else {
        Write-Host "`nFAIL: Found $errorCount errors total" -ForegroundColor Red
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
