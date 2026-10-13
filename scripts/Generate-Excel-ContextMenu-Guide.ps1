param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/excel-context-menu-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$width = 840
$height = 560
$bmp = New-Object System.Drawing.Bitmap ($width, $height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

$excelGreen = [System.Drawing.Color]::FromArgb(16, 124, 65)

# 1. Excel Green Header Sliver
$g.FillRectangle((New-Object System.Drawing.SolidBrush ($excelGreen)), 0, 0, $width, 28)
$fontPptTitle = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Book1 - Excel', $fontPptTitle, [System.Drawing.Brushes]::White, 16, 5)

# 2. Formula Bar (#F3F4F6)
$fBarRect = New-Object System.Drawing.Rectangle (0, 28, $width, 26)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), $fBarRect)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 0, 54, $width, 54)

# Name box
$g.FillRectangle([System.Drawing.Brushes]::White, 12, 31, 46, 20)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 12, 31, 46, 20)
$fontFBar = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$brushDark = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))
$g.DrawString('D2', $fontFBar, $brushDark, 24, 33)

# fx icon
$fontFx = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Italic)
$g.DrawString('fx', $fontFx, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 116, 139))), 68, 32)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 86, 30, 86, 51)

# Formula text
$fontFormula = New-Object System.Drawing.Font ('Consolas', 9.0, [System.Drawing.FontStyle]::Regular)
$g.DrawString('=SMILES2IMG("CC(=O)Oc1ccccc1C(=O)O", "white", "mw")', $fontFormula, $brushDark, 94, 33)

# 3. Excel Spreadsheet Grid
$gridRect = New-Object System.Drawing.Rectangle (0, 55, $width, ($height - 55))
$g.FillRectangle([System.Drawing.Brushes]::White, $gridRect)

# Row header column (left)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), 0, 55, 36, ($height - 55))
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 36, 55, 36, $height)

# Column headers
$colHeaders = @(
    @{ Name = 'A'; W = 60 },
    @{ Name = 'B'; W = 110 },
    @{ Name = 'C'; W = 95 },
    @{ Name = 'D'; W = 180 },
    @{ Name = 'E'; W = 210 },
    @{ Name = 'F'; W = 90 },
    @{ Name = 'G'; W = 90 }
)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), 36, 55, ($width - 36), 22)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), 36, 77, $width, 77)

$curColX = 36
foreach ($col in $colHeaders) {
    $g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), $curColX, 55, $curColX, $height)
    $g.DrawString($col.Name, $fontFBar, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(75, 85, 99))), ($curColX + $col.W / 2 - 5), 58)
    $curColX += $col.W
}

# Rows
$rowHeights = @(24, 130, 130, 80, 80, 80)
$curRowY = 77
$rowIdx = 1
foreach ($rh in $rowHeights) {
    $curRowY += $rh
    $g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1)), 0, $curRowY, $width, $curRowY)
    $g.DrawString("$rowIdx", $fontFBar, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(75, 85, 99))), 12, ($curRowY - $rh + 4))
    $rowIdx++
}

# Table Data
$fontTh = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
$g.DrawString('No.', $fontTh, $brushDark, 48, 82)
$g.DrawString('Compound', $fontTh, $brushDark, 108, 82)
$g.DrawString('Formula', $fontTh, $brushDark, 218, 82)
$g.DrawString('2D Structure (=SMILES2IMG)', $fontTh, $brushDark, 312, 82)
$g.DrawString('SMILES Code', $fontTh, $brushDark, 492, 82)

# Row 2 (Aspirin)
$g.DrawString('1', $fontFBar, $brushDark, 52, 140)
$g.DrawString('Aspirin', $fontTh, $brushDark, 108, 140)
$g.DrawString('C9H8O4', $fontFBar, $brushDark, 218, 140)
$g.DrawString('CC(=O)Oc1ccccc1C(=O)O', $fontFormula, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))), 492, 140)

# Row 3 (Caffeine)
$g.DrawString('2', $fontFBar, $brushDark, 52, 270)
$g.DrawString('Caffeine', $fontTh, $brushDark, 108, 270)
$g.DrawString('C8H10N4O2', $fontFBar, $brushDark, 218, 270)
$g.DrawString('Cn1cnc2c1c(=O)n(C)c(=O)n2C', $fontFormula, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))), 492, 270)

# Draw molecules into cells D2 & D3
$showcasePath = Resolve-Path (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
$showcaseBmp = [System.Drawing.Bitmap]::FromFile($showcasePath.Path)
$aspirinRectSrc = New-Object System.Drawing.Rectangle (98, 368, 372, 418)
$caffeineRectSrc = New-Object System.Drawing.Rectangle (550, 368, 372, 418)

$cellD2PicRect = New-Object System.Drawing.Rectangle (330, 108, 115, 115)
$g.DrawImage($showcaseBmp, $cellD2PicRect, $aspirinRectSrc, [System.Drawing.GraphicsUnit]::Pixel)

$cellD3PicRect = New-Object System.Drawing.Rectangle (330, 238, 115, 115)
$g.DrawImage($showcaseBmp, $cellD3PicRect, $caffeineRectSrc, [System.Drawing.GraphicsUnit]::Pixel)
$showcaseBmp.Dispose()

# Cell D2 Active Cell Green Border
$d2Rect = New-Object System.Drawing.Rectangle (301, 101, 180, 130)
$penGreenCell = New-Object System.Drawing.Pen ($excelGreen, 2)
$g.DrawRectangle($penGreenCell, $d2Rect)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ($excelGreen)), ($d2Rect.Right - 3), ($d2Rect.Bottom - 3), 6, 6)

# 4. Excel Main Context Menu (Cell right-click)
$menuX = 140
$menuY = 88
$menuW = 250
$menuH = 440

# Drop Shadow for Main Menu
for ($i = 7; $i -ge 1; $i--) {
    $alpha = [int](20 - $i * 2.2)
    $sBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
    $g.FillRectangle($sBrush, ($menuX - $i + 3), ($menuY - $i + 4), ($menuW + $i * 2), ($menuH + $i * 2))
}

$menuRect = New-Object System.Drawing.Rectangle ($menuX, $menuY, $menuW, $menuH)
$g.FillRectangle([System.Drawing.Brushes]::White, $menuRect)
$menuBorderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)
$g.DrawRectangle($menuBorderPen, $menuRect)

# Mini Formatting Toolbar
$miniBarRect = New-Object System.Drawing.Rectangle ($menuX, $menuY, $menuW, 46)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(243, 244, 246))), $miniBarRect)
$g.DrawLine($menuBorderPen, $menuX, ($menuY + 46), ($menuX + $menuW), ($menuY + 46))

# Font Box & Size
$fontMini = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$fontMiniBold = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
$fontMiniItalic = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Italic)

$g.FillRectangle([System.Drawing.Brushes]::White, ($menuX + 10), ($menuY + 6), 65, 18)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), ($menuX + 10), ($menuY + 6), 65, 18)
$g.DrawString('Calibri', $fontMini, $brushDark, ($menuX + 13), ($menuY + 7))

$g.FillRectangle([System.Drawing.Brushes]::White, ($menuX + 78), ($menuY + 6), 28, 18)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(209, 213, 219), 1)), ($menuX + 78), ($menuY + 6), 28, 18)
$g.DrawString('11', $fontMini, $brushDark, ($menuX + 84), ($menuY + 7))

$g.DrawString('B', $fontMiniBold, $brushDark, ($menuX + 120), ($menuY + 7))
$g.DrawString('I', $fontMiniItalic, $brushDark, ($menuX + 140), ($menuY + 7))
$g.DrawString('A', $fontMiniBold, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 38, 38))), ($menuX + 162), ($menuY + 7))
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 38, 38))), ($menuX + 162), ($menuY + 21), 10, 2)

# Row 2 of mini toolbar: Fill bucket, borders, align
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(250, 204, 21))), ($menuX + 10), ($menuY + 38), 12, 2)
$g.DrawString('▨', $fontMini, $brushDark, ($menuX + 10), ($menuY + 26))
$g.DrawString('田', $fontMini, $brushDark, ($menuX + 35), ($menuY + 26))
$g.DrawString('≡', $fontMini, $brushDark, ($menuX + 60), ($menuY + 26))
$g.DrawString('$', $fontMiniBold, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(16, 124, 65))), ($menuX + 85), ($menuY + 26))
$g.DrawString('%', $fontMiniBold, $brushDark, ($menuX + 105), ($menuY + 26))

# Search box
$searchRect = New-Object System.Drawing.Rectangle (($menuX + 8), ($menuY + 54), ($menuW - 16), 24)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(249, 250, 251))), $searchRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(229, 231, 235), 1)), $searchRect)

# Vector Magnifying Glass
$penMag = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(148, 163, 184), 1.5)
$g.DrawEllipse($penMag, ($menuX + 18), ($menuY + 60), 7, 7)
$g.DrawLine($penMag, ($menuX + 24), ($menuY + 66), ($menuX + 27), ($menuY + 69))
$fontSearch = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$brushSearch = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(148, 163, 184))
$g.DrawString('Search the menus', $fontSearch, $brushSearch, ($menuX + 32), ($menuY + 58))

# Menu Items
$fontMenu = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Regular)
$fontMenuBold = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$brushMenu = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(31, 41, 55))
$brushMutedMenu = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(107, 114, 128))
$penSep = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(229, 231, 235), 1)

$curY = $menuY + 86

# Helpers for icons
function Draw-Icon-Copy($ix, $iy, $color) {
    $p = New-Object System.Drawing.Pen ($color, 1.2)
    $g.DrawRectangle($p, ($ix + 3), ($iy + 1), 8, 10)
    $g.FillRectangle([System.Drawing.Brushes]::White, ($ix + 1), ($iy + 4), 8, 10)
    $g.DrawRectangle($p, ($ix + 1), ($iy + 4), 8, 10)
}
function Draw-Icon-Scissors($ix, $iy) {
    $p = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(100, 116, 139), 1.2)
    $g.DrawEllipse($p, ($ix + 2), ($iy + 9), 4, 4)
    $g.DrawEllipse($p, ($ix + 2), ($iy + 2), 4, 4)
    $g.DrawLine($p, ($ix + 5), ($iy + 9), ($ix + 12), ($iy + 3))
    $g.DrawLine($p, ($ix + 5), ($iy + 4), ($ix + 12), ($iy + 11))
}

# Cut
Draw-Icon-Scissors ($menuX + 14) ($curY + 4)
$g.DrawString('Cut', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$g.DrawString('Ctrl+X', $fontMini, $brushMutedMenu, ($menuX + $menuW - 50), ($curY + 5))
$curY += 24

# Copy
Draw-Icon-Copy ($menuX + 14) ($curY + 4) ([System.Drawing.Color]::FromArgb(100, 116, 139))
$g.DrawString('Copy', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$g.DrawString('Ctrl+C', $fontMini, $brushMutedMenu, ($menuX + $menuW - 50), ($curY + 5))
$curY += 24

# Paste Options
$g.DrawString('Paste Options:', $fontMenu, $brushMutedMenu, ($menuX + 36), ($curY + 4))
$pasteBtnRect = New-Object System.Drawing.Rectangle (($menuX + 124), ($curY + 2), 20, 20)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(241, 245, 249))), $pasteBtnRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 1)), $pasteBtnRect)
Draw-Icon-Copy ($menuX + 127) ($curY + 5) ([System.Drawing.Color]::FromArgb(100, 116, 139))
$curY += 26

# Separator
$g.DrawLine($penSep, ($menuX + 36), $curY, ($menuX + $menuW - 10), $curY)
$curY += 6

# *** Chemical Structure Submenu (Hovered / Active State) ***
$chemSubmenuRect = New-Object System.Drawing.Rectangle (($menuX + 2), $curY, ($menuW - 4), 26)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 247, 238))), $chemSubmenuRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(187, 247, 208), 1)), $chemSubmenuRect)

# Equation / Beaker icon
$g.DrawString('π', (New-Object System.Drawing.Font ('Cambria Math', 11.0, [System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush ($excelGreen)), ($menuX + 14), ($curY + 3))
$g.DrawString('Chemical Structure', $fontMenuBold, (New-Object System.Drawing.SolidBrush ($excelGreen)), ($menuX + 36), ($curY + 4))
$g.DrawString('›', $fontMenuBold, (New-Object System.Drawing.SolidBrush ($excelGreen)), ($menuX + $menuW - 18), ($curY + 3))
$curY += 28

# More standard Excel items
$g.DrawString('Paste Special...', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

$g.DrawString('Insert...', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

$g.DrawString('Delete...', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

$g.DrawString('Clear Contents', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

# Separator
$g.DrawLine($penSep, ($menuX + 36), $curY, ($menuX + $menuW - 10), $curY)
$curY += 6

$g.DrawString('Quick Analysis', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

$g.DrawString('Filter', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$g.DrawString('›', $fontMenuBold, $brushMutedMenu, ($menuX + $menuW - 18), ($curY + 3))
$curY += 24

$g.DrawString('Sort', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$g.DrawString('›', $fontMenuBold, $brushMutedMenu, ($menuX + $menuW - 18), ($curY + 3))
$curY += 24

$g.DrawString('Format Cells...', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

$g.DrawString('Define Name...', $fontMenu, $brushMenu, ($menuX + 36), ($curY + 4))
$curY += 24

# 5. Flyout Submenu for Chemical Structure!
$subX = $menuX + $menuW - 2
$subY = $chemSubmenuRect.Y
$subW = 310
$subH = 196

# Drop Shadow for Flyout Submenu
for ($j = 7; $j -ge 1; $j--) {
    $alpha = [int](22 - $j * 2.3)
    $sBrush2 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
    $g.FillRectangle($sBrush2, ($subX - $j + 3), ($subY - $j + 4), ($subW + $j * 2), ($subH + $j * 2))
}

$flyoutRect = New-Object System.Drawing.Rectangle ($subX, $subY, $subW, $subH)
$g.FillRectangle([System.Drawing.Brushes]::White, $flyoutRect)
$g.DrawRectangle($menuBorderPen, $flyoutRect)

$curSubY = $subY + 8

# Helper for pen icon
function Draw-Icon-Pen($ix, $iy) {
    $p = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(16, 124, 65), 1.5)
    $g.DrawLine($p, ($ix + 2), ($iy + 11), ($ix + 11), ($iy + 2))
    $g.DrawLine($p, ($ix + 11), ($iy + 2), ($ix + 13), ($iy + 4))
    $g.DrawLine($p, ($ix + 13), ($iy + 4), ($ix + 4), ($iy + 13))
    $g.DrawLine($p, ($ix + 2), ($iy + 11), ($ix + 1), ($iy + 14))
    $g.DrawLine($p, ($ix + 1), ($iy + 14), ($ix + 4), ($iy + 13))
}

# 1. Insert / Edit Structure (SMILES2IMG)... [Highlighted]
$subHoverRect = New-Object System.Drawing.Rectangle (($subX + 2), $curSubY, ($subW - 4), 26)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 247, 238))), $subHoverRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(187, 247, 208), 1)), $subHoverRect)

Draw-Icon-Pen ($subX + 12) ($curSubY + 6)
$g.DrawString('Insert / Edit Structure (SMILES2IMG)...', $fontMenuBold, (New-Object System.Drawing.SolidBrush ($excelGreen)), ($subX + 36), ($curSubY + 5))
$curSubY += 28

# 2. Insert / Edit Structure (KET2IMG)...
$penKet = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(16, 124, 65), 1.4)
$g.DrawRectangle($penKet, ($subX + 14), ($curSubY + 6), 12, 12)
$g.DrawLine($penKet, ($subX + 16), ($curSubY + 10), ($subX + 23), ($curSubY + 10))
$g.DrawString('Insert / Edit Structure (KET2IMG)...', $fontMenu, $brushMenu, ($subX + 36), ($curSubY + 5))
$curSubY += 26

# 3. Paste Structure
$penClip = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(202, 138, 4), 1.4)
$g.DrawRectangle($penClip, ($subX + 14), ($curSubY + 5), 11, 14)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(202, 138, 4))), ($subX + 16), ($curSubY + 3), 7, 3)
$g.DrawString('Paste Structure', $fontMenu, $brushMenu, ($subX + 36), ($curSubY + 5))
$curSubY += 28

# Separator
$g.DrawLine($penSep, ($subX + 36), $curSubY, ($subX + $subW - 10), $curSubY)
$curSubY += 6

# 4. Copy SMILES
Draw-Icon-Copy ($subX + 12) ($curSubY + 5) ([System.Drawing.Color]::FromArgb(37, 99, 235))
$g.DrawString('Copy SMILES', $fontMenu, $brushMenu, ($subX + 36), ($curSubY + 5))
$curSubY += 26

# 5. Copy Canvas Data (KET)
Draw-Icon-Copy ($subX + 12) ($curSubY + 5) ([System.Drawing.Color]::FromArgb(5, 150, 105))
$g.DrawString('Copy Canvas Data (KET)', $fontMenu, $brushMenu, ($subX + 36), ($curSubY + 5))
$curSubY += 26

# 6. Copy Molecule Info (Formula & MW)
$fontInfoIcon = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$g.DrawString('ℹ', $fontInfoIcon, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(217, 119, 6))), ($subX + 14), ($curSubY + 3))
$g.DrawString('Copy Molecule Info (Formula & MW)', $fontMenu, $brushMenu, ($subX + 36), ($curSubY + 5))
$curSubY += 26

$g.Dispose()

# Save
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "[+] Generated Excel context menu guide at $OutputPath" -ForegroundColor Green
