$filePath = (Resolve-Path ".\SMILES2IMG_Comprehensive_Showcase.xlsx").Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open($filePath)
    $errorCount = 0
    for ($i = 1; $i -le 4; $i++) {
        $ws = $wb.Worksheets.Item($i)
        $used = $ws.UsedRange
        for ($r = 1; $r -le $used.Rows.Count; $r++) {
            for ($c = 1; $c -le $used.Columns.Count; $c++) {
                $cell = $used.Cells.Item($r, $c)
                $txt = $cell.Text
                if ($txt -match '^#(VALUE!|NAME\?|REF!|N/A|DIV/0!|NUM!)') {
                    Write-Host "ERROR at $($ws.Name) Row $r Col $c : $txt" -ForegroundColor Red
                    $errorCount++
                }
            }
        }
    }
    if ($errorCount -eq 0) {
        Write-Host "SUCCESS: Zero formula errors across all 4 sheets!" -ForegroundColor Green
    } else {
        Write-Host "FAIL: Found $errorCount errors" -ForegroundColor Red
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
