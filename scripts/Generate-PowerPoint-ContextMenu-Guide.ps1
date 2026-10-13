param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/powerpoint-context-menu-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$width = 860
$height = 580
$bmp = New-Object System.Drawing.Bitmap ($width, $height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

# 1. Background: PowerPoint Workspace (Dark theme workspace `#262626` or `#282828`)
$workspaceBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(40, 40, 40))
$g.FillRectangle($workspaceBrush, 0, 0, $width, $height)

# PowerPoint Top Window / Ribbon Header bar sliver (Orange PowerPoint branding)
$pptOrange = [System.Drawing.Color]::FromArgb(196, 62, 28) # PowerPoint #C43E1C
$g.FillRectangle((New-Object System.Drawing.SolidBrush ($pptOrange)), 0, 0, $width, 32)
$fontPptTitle = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$fontPptSub = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$brushWhite = [System.Drawing.Brushes]::White
$g.DrawString('Presentation1 - PowerPoint', $fontPptTitle, $brushWhite, 16, 7)
$g.DrawString('Slide 1 of 4  •  16:9 Widescreen', $fontPptSub, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(254, 215, 215))), 660, 8)

# Left Slide Thumbnail Sidebar (like real PowerPoint UI!)
$sideRect = New-Object System.Drawing.Rectangle (0, 32, 100, ($height - 32))
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(49, 49, 49))), $sideRect)
$g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60, 60, 60), 1)), 100, 32, 100, $height)

# Slide thumbnail 1 (Active, orange border)
$thumbRect = New-Object System.Drawing.Rectangle (16, 48, 72, 42)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(15, 23, 42))), $thumbRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ($pptOrange, 2)), $thumbRect)
$fontThumbNum = New-Object System.Drawing.Font ('Segoe UI', 8.0, [System.Drawing.FontStyle]::Bold)
$g.DrawString('1', $fontThumbNum, $brushWhite, 6, 48)

# Slide thumbnail 2 (Inactive)
$thumb2Rect = New-Object System.Drawing.Rectangle (16, 102, 72, 42)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(24, 24, 27))), $thumb2Rect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(75, 85, 99), 1)), $thumb2Rect)
$g.DrawString('2', $fontThumbNum, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(156, 163, 175))), 6, 102)

# Main Slide Canvas (Centered with drop shadow)
$slideX = 120
$slideY = 46
$slideW = 716
$slideH = 508

# Slide drop shadow
for ($s = 8; $s -ge 1; $s--) {
    $alpha = [int](25 - $s * 2.5)
    $shBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
    $g.FillRectangle($shBrush, ($slideX - $s + 4), ($slideY - $s + 5), ($slideW + $s * 2), ($slideH + $s * 2))
}

$slideRect = New-Object System.Drawing.Rectangle ($slideX, $slideY, $slideW, $slideH)
$slideBg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(15, 23, 42)) # Slate 900
$g.FillRectangle($slideBg, $slideRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(51, 65, 85), 1)), $slideRect)

# Slide Title & Subtitle
$fontSlideTitle = New-Object System.Drawing.Font ('Segoe UI', 13.5, [System.Drawing.FontStyle]::Bold)
$fontSlideSub = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$brushTitle = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(248, 250, 252))
$brushSub = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(148, 163, 184))

$g.DrawString('Lead Compound Candidates — Bioactive Screening Deck', $fontSlideTitle, $brushTitle, ($slideX + 24), ($slideY + 16))
$g.DrawString('SMILES for PowerPoint • EPAM Indigo Vector Renderer & Ketcher 2D Editor', $fontSlideSub, $brushSub, ($slideX + 24), ($slideY + 39))

# Crop molecules from powerpoint-showcase.png
$showcasePath = Resolve-Path (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
$showcaseBmp = [System.Drawing.Bitmap]::FromFile($showcasePath.Path)
$aspirinRectSrc = New-Object System.Drawing.Rectangle (98, 368, 372, 418)
$caffeineRectSrc = New-Object System.Drawing.Rectangle (550, 368, 372, 418)

# Card 1: Aspirin (Left)
$card1Rect = New-Object System.Drawing.Rectangle (($slideX + 24), ($slideY + 68), 316, 415)
$cardBg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))
$cardBorder = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(71, 85, 105), 1)
$g.FillRectangle($cardBg, $card1Rect)
$g.DrawRectangle($cardBorder, $card1Rect)

$fontCardTitle = New-Object System.Drawing.Font ('Segoe UI', 10.5, [System.Drawing.FontStyle]::Bold)
$brushGold = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(250, 204, 21))
$g.DrawString('Aspirin', $fontCardTitle, $brushGold, ($slideX + 38), ($slideY + 80))
$g.DrawString('C9H8O4 • Analgesic & Anti-inflammatory', $fontSlideSub, $brushSub, ($slideX + 38), ($slideY + 100))

# Aspirin picture box (White box inside card)
$pic1Rect = New-Object System.Drawing.Rectangle (($slideX + 38), ($slideY + 124), 288, 280)
$g.FillRectangle([System.Drawing.Brushes]::White, $pic1Rect)
$g.DrawImage($showcaseBmp, $pic1Rect, $aspirinRectSrc, [System.Drawing.GraphicsUnit]::Pixel)

$fontCode = New-Object System.Drawing.Font ('Consolas', 8.0, [System.Drawing.FontStyle]::Regular)
$g.DrawString('SMILES: CC(=O)Oc1ccccc1C(=O)O', $fontCode, $brushSub, ($slideX + 38), ($slideY + 414))
$g.DrawString('MW: 180.16 g/mol  •  Exact Mass: 180.04 g/mol', $fontCode, $brushSub, ($slideX + 38), ($slideY + 432))
$g.DrawString('* Right-click or double-click to edit structure in Ketcher', $fontSlideSub, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(96, 165, 250))), ($slideX + 38), ($slideY + 454))

# Card 2: Caffeine (Right)
$card2Rect = New-Object System.Drawing.Rectangle (($slideX + 364), ($slideY + 68), 328, 415)
$g.FillRectangle($cardBg, $card2Rect)
$g.DrawRectangle($cardBorder, $card2Rect)

$g.DrawString('Caffeine', $fontCardTitle, $brushGold, ($slideX + 378), ($slideY + 80))
$g.DrawString('C8H10N4O2 • Central Nervous System Stimulant', $fontSlideSub, $brushSub, ($slideX + 378), ($slideY + 100))

$pic2Rect = New-Object System.Drawing.Rectangle (($slideX + 378), ($slideY + 124), 300, 280)
$g.FillRectangle([System.Drawing.Brushes]::White, $pic2Rect)
$g.DrawImage($showcaseBmp, $pic2Rect, $caffeineRectSrc, [System.Drawing.GraphicsUnit]::Pixel)

$g.DrawString('SMILES: Cn1cnc2c1c(=O)n(C)c(=O)n2C', $fontCode, $brushSub, ($slideX + 378), ($slideY + 414))
$g.DrawString('MW: 194.19 g/mol  •  Exact Mass: 194.08 g/mol', $fontCode, $brushSub, ($slideX + 378), ($slideY + 432))
$g.DrawString('Formula: C8H10N4O2 • Xanthine Class', $fontSlideSub, $brushSub, ($slideX + 378), ($slideY + 454))

$showcaseBmp.Dispose()

# PowerPoint Picture Selection Handles around Aspirin picture
$penSelect = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(0, 120, 212), 1.5)
$g.DrawRectangle($penSelect, ($pic1Rect.X - 1), ($pic1Rect.Y - 1), ($pic1Rect.Width + 2), ($pic1Rect.Height + 2))

function Draw-Handle($hx, $hy) {
    $rect = New-Object System.Drawing.Rectangle (($hx - 4), ($hy - 4), 8, 8)
    $g.FillRectangle([System.Drawing.Brushes]::White, $rect)
    $g.DrawRectangle($penSelect, $rect)
}
Draw-Handle $pic1Rect.Left $pic1Rect.Top
Draw-Handle ($pic1Rect.Left + $pic1Rect.Width / 2) $pic1Rect.Top
Draw-Handle $pic1Rect.Right $pic1Rect.Top
Draw-Handle $pic1Rect.Left ($pic1Rect.Top + $pic1Rect.Height / 2)
Draw-Handle $pic1Rect.Right ($pic1Rect.Top + $pic1Rect.Height / 2)
Draw-Handle $pic1Rect.Left $pic1Rect.Bottom
Draw-Handle ($pic1Rect.Left + $pic1Rect.Width / 2) $pic1Rect.Bottom
Draw-Handle $pic1Rect.Right $pic1Rect.Bottom

# Rotation handle at top
$midTopX = [int]($pic1Rect.Left + $pic1Rect.Width / 2)
$g.DrawLine($penSelect, $midTopX, $pic1Rect.Top, $midTopX, ($pic1Rect.Top - 14))
$rotCircleRect = New-Object System.Drawing.Rectangle (($midTopX - 5), ($pic1Rect.Top - 23), 10, 10)
$g.FillEllipse([System.Drawing.Brushes]::White, $rotCircleRect)
$g.DrawEllipse($penSelect, $rotCircleRect)

# 2. PowerPoint Picture Context Menu (Positioned naturally over right side of selection)
$menuX = $pic1Rect.Left + 140
$menuY = $pic1Rect.Top - 20
$menuW = 280
$menuH = 410

# Multi-layered smooth shadow
for ($i = 8; $i -ge 1; $i--) {
    $alpha = [int](22 - $i * 2.2)
    $sBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
    $sRect = New-Object System.Drawing.Rectangle (($menuX - $i + 4), ($menuY - $i + 6), ($menuW + $i * 2), ($menuH + $i * 2))
    $g.FillRectangle($sBrush, $sRect)
}

# Context Menu Box
$menuRect = New-Object System.Drawing.Rectangle ($menuX, $menuY, $menuW, $menuH)
$menuBgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 255))
$menuBorderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 1)
$g.FillRectangle($menuBgBrush, $menuRect)
$g.DrawRectangle($menuBorderPen, $menuRect)

# Mini Formatting Toolbar at top of Context Menu
$miniBarRect = New-Object System.Drawing.Rectangle ($menuX, $menuY, $menuW, 32)
$miniBarBg = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(248, 250, 252))
$g.FillRectangle($miniBarBg, $miniBarRect)
$g.DrawLine($menuBorderPen, $menuX, ($menuY + 32), ($menuX + $menuW), ($menuY + 32))

# Mini toolbar buttons
$fontMini = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$brushMiniText = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))
$g.DrawString('Style', $fontMini, $brushMiniText, ($menuX + 14), ($menuY + 8))
$g.DrawString('Crop', $fontMini, $brushMiniText, ($menuX + 68), ($menuY + 8))
$g.DrawString('Arrange', $fontMini, $brushMiniText, ($menuX + 120), ($menuY + 8))
$g.DrawString('Format', $fontMini, $brushMiniText, ($menuX + 192), ($menuY + 8))

# Search the menus
$searchRect = New-Object System.Drawing.Rectangle (($menuX + 10), ($menuY + 40), ($menuW - 20), 24)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(248, 250, 252))), $searchRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1)), $searchRect)

# Draw Vector Magnifying Glass
$penMag = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(148, 163, 184), 1.5)
$g.DrawEllipse($penMag, ($menuX + 20), ($menuY + 46), 7, 7)
$g.DrawLine($penMag, ($menuX + 26), ($menuY + 52), ($menuX + 29), ($menuY + 55))

$fontSearch = New-Object System.Drawing.Font ('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular)
$brushSearch = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(148, 163, 184))
$g.DrawString('Search the menus', $fontSearch, $brushSearch, ($menuX + 34), ($menuY + 44))

# Menu Items
$fontMenu = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Regular)
$fontMenuBold = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$brushMenu = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))
$brushMutedMenu = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 116, 139))
$penSep = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1)

$curY = $menuY + 72

# Helper for Drawing Icons
function Draw-Icon-Pen($ix, $iy) {
    # Small pencil / edit icon
    $p = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(37, 99, 235), 1.6)
    $g.DrawLine($p, ($ix + 2), ($iy + 11), ($ix + 11), ($iy + 2))
    $g.DrawLine($p, ($ix + 11), ($iy + 2), ($ix + 13), ($iy + 4))
    $g.DrawLine($p, ($ix + 13), ($iy + 4), ($ix + 4), ($iy + 13))
    $g.DrawLine($p, ($ix + 2), ($iy + 11), ($ix + 1), ($iy + 14))
    $g.DrawLine($p, ($ix + 1), ($iy + 14), ($ix + 4), ($iy + 13))
}

function Draw-Icon-Copy($ix, $iy, $color) {
    $p = New-Object System.Drawing.Pen ($color, 1.2)
    # Background page
    $g.DrawRectangle($p, ($ix + 4), ($iy + 1), 8, 10)
    # Foreground page
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

# 1. Edit Structure (Ketcher)... [Highlighted Hover State]
$hoverRect = New-Object System.Drawing.Rectangle (($menuX + 3), $curY, ($menuW - 6), 28)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(239, 246, 255))), $hoverRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(191, 219, 254), 1)), $hoverRect)

Draw-Icon-Pen ($menuX + 14) ($curY + 6)
$g.DrawString('Edit Structure (Ketcher)...', $fontMenuBold, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(29, 78, 216))), ($menuX + 38), ($curY + 6))
$curY += 29

# 2. Copy SMILES
Draw-Icon-Copy ($menuX + 14) ($curY + 6) ([System.Drawing.Color]::FromArgb(37, 99, 235))
$g.DrawString('Copy SMILES', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 6))
$curY += 26

# 3. Copy Canvas Data (KET)
Draw-Icon-Copy ($menuX + 14) ($curY + 6) ([System.Drawing.Color]::FromArgb(37, 99, 235))
$g.DrawString('Copy Canvas Data (KET)', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 6))
$curY += 26

# 4. Copy Molecule Info (Formula & MW)
$fontInfoIcon = New-Object System.Drawing.Font ('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$g.DrawString('ℹ', $fontInfoIcon, (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(37, 99, 235))), ($menuX + 14), ($curY + 4))
$g.DrawString('Copy Molecule Info (Formula & MW)', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 6))
$curY += 28

# Separator Line
$g.DrawLine($penSep, ($menuX + 38), $curY, ($menuX + $menuW - 12), $curY)
$curY += 6

# Standard PowerPoint Picture Items
# Cut
Draw-Icon-Scissors ($menuX + 14) ($curY + 5)
$g.DrawString('Cut', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 5))
$g.DrawString('Ctrl+X', $fontMini, $brushMutedMenu, ($menuX + $menuW - 54), ($curY + 6))
$curY += 24

# Copy
Draw-Icon-Copy ($menuX + 14) ($curY + 5) ([System.Drawing.Color]::FromArgb(100, 116, 139))
$g.DrawString('Copy', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 5))
$g.DrawString('Ctrl+C', $fontMini, $brushMutedMenu, ($menuX + 54 + 160), ($curY + 6))
$curY += 24

# Paste Options
$g.DrawString('Paste Options:', $fontMenu, $brushMutedMenu, ($menuX + 38), ($curY + 5))
# Paste clipboard mini button
$pasteBtnRect = New-Object System.Drawing.Rectangle (($menuX + 130), ($curY + 2), 22, 22)
$g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(241, 245, 249))), $pasteBtnRect)
$g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 1)), $pasteBtnRect)
Draw-Icon-Copy ($menuX + 133) ($curY + 5) ([System.Drawing.Color]::FromArgb(100, 116, 139))
$curY += 26

# Separator
$g.DrawLine($penSep, ($menuX + 38), $curY, ($menuX + $menuW - 12), $curY)
$curY += 6

# Bring to Front
$g.DrawString('Bring to Front', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 5))
$g.DrawString('›', $fontMenuBold, $brushMutedMenu, ($menuX + $menuW - 22), ($curY + 3))
$curY += 24

# Send to Back
$g.DrawString('Send to Back', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 5))
$g.DrawString('›', $fontMenuBold, $brushMutedMenu, ($menuX + $menuW - 22), ($curY + 3))
$curY += 24

# Save as Picture...
$g.DrawString('Save as Picture...', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 5))
$curY += 24

# Format Picture...
$g.DrawString('Format Picture...', $fontMenu, $brushMenu, ($menuX + 38), ($curY + 5))
$curY += 24

$g.Dispose()

# Save
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "[+] Generated PowerPoint context menu guide at $OutputPath" -ForegroundColor Green
