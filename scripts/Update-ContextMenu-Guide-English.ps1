param(
    [string]$ImagePath = (Join-Path $PSScriptRoot '../assets/context-menu-guide.png'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/context-menu-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$bmp = [System.Drawing.Bitmap]::FromFile((Resolve-Path $ImagePath).Path)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

# Brushes
$bgMenu = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(189, 189, 189))
$textBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(38, 38, 38))
$mutedBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(115, 115, 115))
$paperBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(224, 224, 224))

# Fonts (System Segoe UI matching Office UI specifications)
$fontMenu = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Regular)
$fontMenuMuted = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$fontBold = New-Object System.Drawing.Font ('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$fontItalic = New-Object System.Drawing.Font ('Segoe UI', 9.5, [System.Drawing.FontStyle]::Italic)
$fontMiniA = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
$fontMiniSmall = New-Object System.Drawing.Font ('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)

# 1. Mini formatting toolbar updates
# Font dropdown (맑은 고딕 -> Calibri):
$g.FillRectangle($bgMenu, 144, 98, 64, 20)
$g.DrawString('Calibri', $fontMenu, $textBrush, 146, 101)

# Grow font & shrink font:
$g.FillRectangle($bgMenu, 258, 98, 48, 20)
$g.DrawString('A', $fontMiniA, $textBrush, 266, 100)
$g.DrawString('^', $fontMiniSmall, $textBrush, 276, 98)
$g.DrawString('A', $fontMiniSmall, $textBrush, 290, 102)

# Bold (가 -> B):
$g.FillRectangle($bgMenu, 142, 122, 22, 20)
$g.DrawString('B', $fontBold, $textBrush, 148, 123)

# Italic (가 -> I):
$g.FillRectangle($bgMenu, 168, 122, 22, 20)
$g.DrawString('I', $fontItalic, $textBrush, 174, 123)

# Font color button (가 -> A above red color bar):
$g.FillRectangle($bgMenu, 250, 120, 24, 16)
$g.DrawString('A', $fontBold, $textBrush, 257, 120)

# 2. Search box (메뉴 검색 -> Search the menus):
$g.FillRectangle($bgMenu, 148, 160, 227, 25)
$g.DrawString('Search the menus', $fontMenu, $mutedBrush, 170, 166)

# 3. Context Menu Items:
function Update-Row ($midY, $text, $isMuted = $false, $hasIcon = $true, $hasRightArrow = $false) {
    $xStart = if ($hasIcon) { 172 } else { 164 }
    $xEnd = if ($hasRightArrow) { 364 } else { 375 }
    $width = $xEnd - $xStart
    $rect = New-Object System.Drawing.Rectangle ($xStart, ($midY - 9), $width, 18)
    $g.FillRectangle($bgMenu, $rect)
    $f = if ($isMuted) { $fontMenuMuted } else { $fontMenu }
    $b = if ($isMuted) { $mutedBrush } else { $textBrush }
    $g.DrawString($text, $f, $b, 204, ($midY - 7))
}

Update-Row 208 'Cut' $false $true
Update-Row 238 'Copy' $false $true
Update-Row 270 'Paste Options:' $true $true

# Paste button paper icon letter '가' inside clipboard paper:
$g.FillRectangle($paperBrush, 180, 302, 16, 20)

# Chemical Structure is at midY = 334 (kept 100% authentic real screenshot!)

Update-Row 362 'Paste Special...' $false $false
Update-Row 393 'Insert...' $false $false
Update-Row 420 'Delete...' $false $false
Update-Row 449 'Clear Contents' $false $false
Update-Row 480 'Quick Analysis' $false $true
Update-Row 508 'Filter' $false $false $true
Update-Row 536 'Sort' $false $false $true
Update-Row 566 'Get Data from Table/Range...' $false $true
Update-Row 598 'New Comment' $false $true
Update-Row 626 'New Note' $false $true
Update-Row 657 'Format Cells...' $false $true
Update-Row 684 'Pick from Drop-down List...' $false $false

# Define Name area (clearing all Asian phonetic text and 내천 icon):
$g.FillRectangle($bgMenu, 162, 642, 214, 126)
$g.DrawString('Define Name...', $fontMenu, $textBrush, 204, 706)

Update-Row 772 'Link' $false $true $true
Update-Row 798 'Open Hyperlink' $false $false
Update-Row 830 'Link to this range' $false $true

$g.Dispose()

# Crop top 46 px (removes Korean ribbon sliver completely, keeping formula bar and grid)
$rectResult = New-Object System.Drawing.Rectangle (0, 46, $bmp.Width, ($bmp.Height - 46))
$croppedResult = $bmp.Clone($rectResult, $bmp.PixelFormat)
$bmp.Dispose()

# Save output
$tmpPath = [System.IO.Path]::GetTempFileName() + ".png"
$croppedResult.Save($tmpPath, [System.Drawing.Imaging.ImageFormat]::Png)
$croppedResult.Dispose()

Move-Item -Path $tmpPath -Destination $OutputPath -Force
Write-Host "[+] Updated Context Menu Guide with 100% English UI: $OutputPath" -ForegroundColor Green

# Sync to docs/public
$docsPublicPath = (Join-Path $PSScriptRoot '../docs/public/context-menu-guide.png')
Copy-Item -Path $OutputPath -Destination $docsPublicPath -Force
Write-Host "[+] Synced Context Menu Guide to docs/public: $docsPublicPath" -ForegroundColor Cyan
