param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/excel-ribbon-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$width = 760
$height = 360
$bmp = New-Object System.Drawing.Bitmap ($width, $height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

# 1. Excel Green Header (#107C41)
$excelGreen = [System.Drawing.Color]::FromArgb(16, 124, 65)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ($excelGreen)), 0, 0, $width, 58)

# Title Bar
$fontTitle = New-Object System.Drawing.Font ('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$fontTab = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Regular)
$fontTabBold = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$brushWhite = [System.Drawing.Brushes]::White
$brushTabDim = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 245, 238))

$g.DrawString('AutoSave', $fontTitle, $brushWhite, 16, 6)
$g.DrawString('Book1 - Excel', $fontTitle, $brushWhite, 340, 6)

# Ribbon Tabs
$tabs = @('File', 'Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Automate')
$xTab = 16
foreach ($t in $tabs) {
    $g.DrawString($t, $fontTab, $brushTabDim, $xTab, 32)
    $sz = $g.MeasureString($t, $fontTab)
    $xTab += [int]$sz.Width + 14
}

# Active SMILES Tab
$smilesTabRect = New-Object System.Drawing.Rectangle ($xTab, 27, 68, 31)
$g.FillRectangle([System.Drawing.Brushes]::White, $smilesTabRect)
$g.DrawString('SMILES', $fontTabBold, (New-Object System.Drawing.SolidBrush ($excelGreen)), ($xTab + 10), 32)

# 2. Ribbon Surface
$ribbonRect = New-Object System.Drawing.Rectangle (0, 58, $width, 94)
$ribbonBg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(249, 250, 251))
$g.FillRectangle($ribbonBg, $ribbonRect)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(229, 231, 235), 1)), 0, 152, $width, 152)

# Group Dividers
$penDivider = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1)
$g.DrawLine($penDivider, 232, 64, 232, 140)
$g.DrawLine($penDivider, 464, 64, 464, 140)
$g.DrawLine($penDivider, 574, 64, 574, 140)

# Group Labels
$fontGrp = New-Object System.Drawing.Font ('Segoe UI', 8.0, [System.Drawing.FontStyle]::Regular)
$brushGrp = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 116, 139))
$g.DrawString('Insert Structure', $fontGrp, $brushGrp, 76, 136)
$g.DrawString('Structure Tools', $fontGrp, $brushGrp, 306, 136)
$g.DrawString('Diagnostics', $fontGrp, $brushGrp, 492, 136)

# Button Text & Icons
$fontBtn = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$fontBtnBold = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
$fontIconSymbol = New-Object System.Drawing.Font ('Cambria Math', 17.0, [System.Drawing.FontStyle]::Bold)
$fontIconStd = New-Object System.Drawing.Font ('Segoe UI Symbol', 14.0, [System.Drawing.FontStyle]::Regular)
$brushDark = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))

# Helper to draw button
function Draw-Ribbon-Button($x, $y, $w, $h, $isHighlighted) {
    if ($isHighlighted) {
        $btnRect = New-Object System.Drawing.Rectangle ($x, $y, $w, $h)
        $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 247, 238))), $btnRect)
        $g.DrawRectangle((New-Object System.Drawing.Pen ($excelGreen, 1)), $btnRect)
    }
}

# 1. Insert / Edit (SMILES) (Active / Hovered)
Draw-Ribbon-Button 14 62 80 70 $true
$g.DrawString('π', $fontIconSymbol, (New-Object System.Drawing.SolidBrush ($excelGreen)), 42, 64)
$g.DrawString('Insert / Edit', $fontBtnBold, $brushDark, 20, 96)
$g.DrawString('(SMILES)', $fontBtnBold, $brushDark, 28, 110)

# 2. Insert / Edit (KET)
Draw-Ribbon-Button 98 62 66 70 $false
$penKet = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(16, 124, 65), 1.5)
$g.DrawRectangle($penKet, 120, 68, 18, 18)
$g.DrawLine($penKet, 124, 74, 134, 74)
$g.DrawLine($penKet, 124, 78, 134, 78)
$g.DrawString('Insert / Edit', $fontBtn, $brushDark, 100, 96)
$g.DrawString('(KET)', $fontBtn, $brushDark, 116, 110)

# 3. Paste Structure
Draw-Ribbon-Button 166 62 64 70 $false
# Clipboard icon
$penClip = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(202, 138, 4), 1.5)
$g.DrawRectangle($penClip, 188, 70, 16, 18)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(202, 138, 4))), 192, 67, 8, 4)
$g.DrawString('Paste', $fontBtn, $brushDark, 181, 96)
$g.DrawString('Structure', $fontBtn, $brushDark, 172, 110)

# 4. Copy SMILES
Draw-Ribbon-Button 242 62 70 70 $false
$penCopy = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(37, 99, 235), 1.4)
$g.DrawRectangle($penCopy, 268, 69, 13, 15)
$g.DrawRectangle($penCopy, 265, 73, 13, 15)
$g.DrawString('Copy', $fontBtn, $brushDark, 262, 96)
$g.DrawString('SMILES', $fontBtn, $brushDark, 256, 110)

# 5. Copy KET
Draw-Ribbon-Button 316 62 68 70 $false
$penKetCopy = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(5, 150, 105), 1.4)
$g.DrawRectangle($penKetCopy, 342, 69, 13, 15)
$g.DrawRectangle($penKetCopy, 339, 73, 13, 15)
$g.DrawString('Copy', $fontBtn, $brushDark, 336, 96)
$g.DrawString('KET', $fontBtn, $brushDark, 340, 110)

# 6. Copy Molecule Info
Draw-Ribbon-Button 388 62 72 70 $false
$g.DrawString('ℹ', (New-Object System.Drawing.Font ('Segoe UI', 14.0, [System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(217, 119, 6))), 415, 66)
$g.DrawString('Copy Info', $fontBtn, $brushDark, 396, 96)
$g.DrawString('(Formula/MW)', (New-Object System.Drawing.Font ('Segoe UI', 7.5, [System.Drawing.FontStyle]::Regular)), $brushDark, 390, 110)

# 7. Diagnostics
Draw-Ribbon-Button 474 62 92 70 $false
$penShield = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(234, 88, 12), 1.5)
[System.Drawing.Point[]]$shieldPoints = @(
    (New-Object System.Drawing.Point (520, 68)),
    (New-Object System.Drawing.Point (510, 86)),
    (New-Object System.Drawing.Point (530, 86))
)
$g.DrawPolygon($penShield, $shieldPoints)
$g.DrawString('!', (New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(234, 88, 12))), 518, 73)
$g.DrawString('Settings &', $fontBtn, $brushDark, 489, 96)
$g.DrawString('Diagnostics', $fontBtn, $brushDark, 485, 110)

# 3. Excel Formula Bar (#F9FAFB)
$fBarRect = New-Object System.Drawing.Rectangle (0, 153, $width, 27)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), $fBarRect)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 0, 180, $width, 180)

# Name box
$g.FillRectangle([System.Drawing.Brushes]::White, 16, 157, 50, 19)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 16, 157, 50, 19)
$fontFBar = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$g.DrawString('C3', $fontFBar, $brushDark, 30, 159)

# fx icon
$fontFx = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Italic)
$g.DrawString('fx', $fontFx, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 116, 139))), 78, 158)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 94, 156, 94, 177)

# Formula text
$fontFormula = New-Object System.Drawing.Font ('Consolas', 9.0, [System.Drawing.FontStyle]::Regular)
$g.DrawString('=SMILES2IMG("CC(=O)Oc1ccccc1C(=O)O", "white", "mw")', $fontFormula, $brushDark, 102, 159)

# 4. Excel Worksheet Grid
$gridRect = New-Object System.Drawing.Rectangle (0, 181, $width, ($height - 181))
$g.FillRectangle([System.Drawing.Brushes]::White, $gridRect)

# Row header column (left)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), 0, 181, 40, ($height - 181))
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 40, 181, 40, $height)

# Column header row
$colHeaders = @(
    @{ Name = 'A'; W = 80 },
    @{ Name = 'B'; W = 110 },
    @{ Name = 'C'; W = 160 },
    @{ Name = 'D'; W = 160 },
    @{ Name = 'E'; W = 110 },
    @{ Name = 'F'; W = 100 }
)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), 40, 181, ($width - 40), 22)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 40, 203, $width, 203)

$curColX = 40
foreach ($col in $colHeaders) {
    $g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), $curColX, 181, $curColX, $height)
    $g.DrawString($col.Name, $fontFBar, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(75, 85, 99))), ($curColX + $col.W / 2 - 5), 184)
    $curColX += $col.W
}

# Horizontal Grid Lines
$rowHeights = @(22, 28, 120)
$curRowY = 203
$rowIdx = 1
foreach ($rh in $rowHeights) {
    $curRowY += $rh
    $g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1)), 0, $curRowY, $width, $curRowY)
    $g.DrawString("$rowIdx", $fontFBar, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(75, 85, 99))), 16, ($curRowY - $rh + 4))
    $rowIdx++
}

# Table headers in Row 1 & 2
$fontTh = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Compound', $fontTh, $brushDark, 48, 208)
$g.DrawString('Formula', $fontTh, $brushDark, 128, 208)
$g.DrawString('2D Structure (=SMILES2IMG)', $fontTh, $brushDark, 238, 208)
$g.DrawString('SMILES Code', $fontTh, $brushDark, 398, 208)

# Row 2 values
$g.DrawString('Aspirin', $fontFBar, $brushDark, 48, 236)
$g.DrawString('C9H8O4', $fontFBar, $brushDark, 128, 236)
$g.DrawString('CC(=O)Oc1ccccc1C(=O)O', $fontFormula, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))), 398, 236)

# Molecule in Cell C3 (Row 2 structure cell)
$showcasePath = Resolve-Path (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
$showcaseBmp = [System.Drawing.Bitmap]::FromFile($showcasePath.Path)
$aspirinRectSrc = New-Object System.Drawing.Rectangle (98, 368, 372, 418)
$cellPicRect = New-Object System.Drawing.Rectangle (248, 230, 115, 115)
$g.DrawImage($showcaseBmp, $cellPicRect, $aspirinRectSrc, [System.Drawing.GraphicsUnit]::Pixel)
$showcaseBmp.Dispose()

# Cell C3 Active Selection (Green Excel Border #107C41)
$c3Rect = New-Object System.Drawing.Rectangle (230, 225, 160, 128)
$penCellSelect = New-Object System.Drawing.Pen ($excelGreen, 2)
$g.DrawRectangle($penCellSelect, $c3Rect)
# Small green fill handle at bottom right
$g.FillRectangle((New-Object System.Drawing.SolidBrush ($excelGreen)), ($c3Rect.Right - 3), ($c3Rect.Bottom - 3), 6, 6)

# 5. Screentip / Supertip Tooltip (Hovering under Insert / Edit SMILES)
$tipX = 14
$tipY = 138
$tipW = 390
$tipH = 80

# Tooltip Drop Shadow
for ($t = 6; $t -ge 1; $t--) {
    $alpha = [int](18 - $t * 2)
    $tBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
    $g.FillRectangle($tBrush, ($tipX - $t + 2), ($tipY - $t + 3), ($tipW + $t * 2), ($tipH + $t * 2))
}

$tipRect = New-Object System.Drawing.Rectangle ($tipX, $tipY, $tipW, $tipH)
$g.FillRectangle([System.Drawing.Brushes]::White, $tipRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 1)), $tipRect)

$fontTipTitle = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$fontTipBody = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$g.DrawString('Insert / Edit Structure (SMILES2IMG) (Ctrl+Shift+K)', $fontTipTitle, $brushDark, ($tipX + 12), ($tipY + 10))
$tipText = "Opens Ketcher editor to sketch chemical structures and insert them into the active cell as formulas (=SMILES2IMG)."
$g.DrawString($tipText, $fontTipBody, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))), (New-Object System.Drawing.RectangleF (($tipX + 12), ($tipY + 30), ($tipW - 24), 44)))

$g.Dispose()

# Save
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "[+] Generated Excel ribbon guide at $OutputPath" -ForegroundColor Green
