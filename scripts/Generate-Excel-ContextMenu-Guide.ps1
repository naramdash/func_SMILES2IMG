param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/excel-context-menu-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
. (Join-Path $PSScriptRoot 'Draw-ContextMenuFlyout.ps1')

$width = 920
$height = 580
$bmp = New-Object System.Drawing.Bitmap($width, $height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

$excelGreen = [System.Drawing.Color]::FromArgb(16, 124, 65)
$brushDark = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 41, 59))
$fontTitle = New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$fontBar = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$fontFormula = New-Object System.Drawing.Font('Consolas', 9.0, [System.Drawing.FontStyle]::Regular)
$fontTh = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)

# 1. Excel Header & Formula Bar
$g.FillRectangle((New-Object System.Drawing.SolidBrush($excelGreen)), 0, 0, $width, 28)
$g.DrawString('SMILES2IMG_Showcase.xlsx - Excel', $fontTitle, [System.Drawing.Brushes]::White, 16, 5)

$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(243, 244, 246))), 0, 28, $width, 26)
$g.DrawLine((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 0, 54, $width, 54)
$g.FillRectangle([System.Drawing.Brushes]::White, 12, 31, 46, 20)
$g.DrawRectangle((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 12, 31, 46, 20)
$g.DrawString('D2', $fontBar, $brushDark, 24, 33)
$g.DrawString('fx', (New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Italic)), (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100, 116, 139))), 68, 32)
$g.DrawLine((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 86, 30, 86, 51)
$g.DrawString('=SMILES2IMG(E2)', $fontFormula, $brushDark, 96, 33)

# 2. Spreadsheet Grid
$g.FillRectangle([System.Drawing.Brushes]::White, 0, 55, $width, ($height - 55))
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(243, 244, 246))), 0, 55, 36, ($height - 55))
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(243, 244, 246))), 36, 55, ($width - 36), 22)
$g.DrawLine((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 36, 55, 36, $height)
$g.DrawLine((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 36, 77, $width, 77)

$cols = @(@{Name='A';W=56}, @{Name='B';W=105}, @{Name='C';W=95}, @{Name='D';W=185}, @{Name='E';W=235}, @{Name='F';W=110}, @{Name='G';W=98})
$cx = 36
foreach ($c in $cols) {
    $g.DrawLine((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), $cx, 55, $cx, $height)
    $g.DrawString($c.Name, $fontBar, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(75, 85, 99))), ($cx + $c.W/2 - 5), 58)
    $cx += $c.W
}
$ry = 77; $rowHeights = @(24, 130, 130, 90, 90)
for ($i=0; $i -lt $rowHeights.Count; $i++) {
    $ry += $rowHeights[$i]
    $g.DrawLine((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(226, 232, 240), 1)), 0, $ry, $width, $ry)
    $g.DrawString("$($i+1)", $fontBar, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(75, 85, 99))), 12, ($ry - $rowHeights[$i] + 4))
}

$g.DrawString('No.', $fontTh, $brushDark, 48, 82)
$g.DrawString('Compound', $fontTh, $brushDark, 102, 82)
$g.DrawString('Formula', $fontTh, $brushDark, 210, 82)
$g.DrawString('2D Structure (=SMILES2IMG)', $fontTh, $brushDark, 304, 82)
$g.DrawString('SMILES', $fontTh, $brushDark, 492, 82)
$g.DrawString('MW (g/mol)', $fontTh, $brushDark, 725, 82)

$g.DrawString('1', $fontBar, $brushDark, 52, 140)
$g.DrawString('Aspirin', $fontTh, $brushDark, 102, 140)
$g.DrawString('C9H8O4', $fontBar, $brushDark, 210, 140)
$g.DrawString('CC(=O)Oc1ccccc1C(=O)O', $fontFormula, $brushDark, 492, 140)
$g.DrawString('180.16', $fontBar, $brushDark, 738, 140)

$showcasePath = Resolve-Path (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
$sb = [System.Drawing.Bitmap]::FromFile($showcasePath.Path)
$g.DrawImage($sb, (New-Object System.Drawing.Rectangle(324, 108, 114, 114)), (New-Object System.Drawing.Rectangle(98, 368, 372, 418)), [System.Drawing.GraphicsUnit]::Pixel)
$g.DrawImage($sb, (New-Object System.Drawing.Rectangle(324, 238, 114, 114)), (New-Object System.Drawing.Rectangle(550, 368, 372, 418)), [System.Drawing.GraphicsUnit]::Pixel)
$sb.Dispose()

$d2 = New-Object System.Drawing.Rectangle(292, 101, 185, 130)
$g.DrawRectangle((New-Object System.Drawing.Pen($excelGreen, 2)), $d2)

# 3. Cascading Excel Right-Click Context Menus (matches BuildExcelRangeContextMenu)
$mainItems = @(
    @{ Label = 'Cut'; Icon = 'cut'; Shortcut = 'Ctrl+X' },
    @{ Label = 'Copy'; Icon = 'copy'; Shortcut = 'Ctrl+C' },
    @{ Label = 'Paste Options:'; Muted = $true },
    @{ Kind = 'sep' },
    @{ Label = 'Chemical Structure'; Icon = 'chem'; Highlight = $true; HasSub = $true; Key = 'chem' },
    @{ Kind = 'sep' },
    @{ Label = 'Paste Special...' },
    @{ Label = 'Insert...' },
    @{ Label = 'Delete...' },
    @{ Label = 'Clear Contents' },
    @{ Kind = 'sep' },
    @{ Label = 'Filter'; HasSub = $true },
    @{ Label = 'Sort'; HasSub = $true },
    @{ Label = 'Format Cells...' }
)
$res1 = Draw-OfficeContextMenuPanel -g $g -X 46 -Y 86 -Width 224 -Items $mainItems -AccentColor $excelGreen -MiniToolbar @{ Mode = 'text' } -ShowSearchBox $true

$blue = [System.Drawing.Color]::FromArgb(37, 99, 235)
$teal = [System.Drawing.Color]::FromArgb(13, 148, 136)
$chemItems = @(
    @{ Label = 'Insert / Edit Structure (SMILES2IMG)...'; Icon = 'chem'; IconColor = $excelGreen },
    @{ Label = 'Insert / Edit Structure (INCHI2IMG)...'; Icon = 'fx'; IconColor = $blue },
    @{ Label = 'Insert / Edit Structure (KET2IMG)...'; Icon = 'ket'; IconColor = $teal },
    @{ Label = '3D Structure (Mol*)...'; Icon = 'cube'; IconColor = $blue },
    @{ Label = 'Paste Structure'; Icon = 'paste'; IconColor = [System.Drawing.Color]::FromArgb(202, 138, 4) },
    @{ Kind = 'sep' },
    @{ Label = 'Copy SMILES'; Icon = 'copy'; IconColor = $blue },
    @{ Label = 'Copy InChI'; Icon = 'inchi'; IconColor = $blue },
    @{ Label = 'Copy Canvas Data (KET)'; Icon = 'copy'; IconColor = $teal },
    @{ Label = 'Copy Vector SVG'; Icon = 'svg'; IconColor = $blue },
    @{ Kind = 'sep' },
    @{ Label = 'Copy Molecule Info'; Icon = 'info'; Highlight = $true; HasSub = $true; Key = 'info' }
)
$sub1Y = [math]::Max(130, $res1.Anchors['chem'].Y - 28)
$res2 = Draw-OfficeContextMenuPanel -g $g -X 268 -Y $sub1Y -Width 294 -Items $chemItems -AccentColor $excelGreen

$infoItems = Get-SharedMoleculeInfoContextItems
$sub2Y = [math]::Min(($height - 315), [math]::Max(200, $res2.Anchors['info'].Y - 220))
[void](Draw-OfficeContextMenuPanel -g $g -X 560 -Y $sub2Y -Width 268 -Items $infoItems -AccentColor $excelGreen)

$g.Dispose()
foreach ($dest in @($OutputPath, (Join-Path $PSScriptRoot '../docs/public/excel-context-menu-guide.png'), (Join-Path $PSScriptRoot '../assets/context-menu-guide.png'), (Join-Path $PSScriptRoot '../docs/public/context-menu-guide.png'))) {
    $dir = Split-Path $dest -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $bmp.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
}
$bmp.Dispose()
Write-Host "[+] Generated v1.8.0 Excel context menu guide at $OutputPath" -ForegroundColor Green
