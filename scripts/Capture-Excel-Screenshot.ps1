param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot '../SMILES2IMG_Comprehensive_Showcase.xlsx'),
    [string]$TableImage = (Join-Path $PSScriptRoot '../assets/excel-showcase.png'),
    [string]$MatrixImage = (Join-Path $PSScriptRoot '../assets/excel-matrix-showcase.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$resolvedWb = (Resolve-Path $WorkbookPath).Path
$tablePath = [System.IO.Path]::GetFullPath($TableImage)
$matrixPath = [System.IO.Path]::GetFullPath($MatrixImage)

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 400

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($resolvedWb)

    # 1. Capture Sheet 1: In-Cell 365 Showcase (A1:G15)
    $sheet1 = $wb.Worksheets.Item("SMILES2IMG (In-Cell 365)")
    $tableRange = $sheet1.Range("A1:G15")
    $tableRange.CopyPicture(1, 2)
    Start-Sleep -Milliseconds 400
    $imgTable = [System.Windows.Forms.Clipboard]::GetImage()
    if ($imgTable) {
        $imgTable.Save($tablePath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host "Saved In-Cell Table Showcase: $tablePath ($($imgTable.Width)x$($imgTable.Height))" -ForegroundColor Green
    }

    # 2. Capture Sheet 4: Option Matrix Showcase (A1:F17)
    $sheet4 = $wb.Worksheets.Item("Style & Transform Matrix")
    $matrixRange = $sheet4.Range("A1:G18")
    $matrixRange.CopyPicture(1, 2)
    Start-Sleep -Milliseconds 400
    $imgMatrix = [System.Windows.Forms.Clipboard]::GetImage()
    if ($imgMatrix) {
        $imgMatrix.Save($matrixPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host "Saved Option Matrix Showcase: $matrixPath ($($imgMatrix.Width)x$($imgMatrix.Height))" -ForegroundColor Green
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "All screenshots captured successfully!" -ForegroundColor Green
