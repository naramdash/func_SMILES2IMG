param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/workflow-ketcher.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$width = 1200
$height = 420
$bmp = New-Object System.Drawing.Bitmap $width, $height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

# Background
$g.Clear([System.Drawing.Color]::FromArgb(248, 250, 252)) # slate-50

# Header Bar (Dark Slate)
$headerRect = New-Object System.Drawing.Rectangle 0, 0, $width, 80
$headerBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(15, 23, 42)) # slate-900
$g.FillRectangle($headerBrush, $headerRect)

$titleFont = New-Object System.Drawing.Font ("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$subTitleFont = New-Object System.Drawing.Font ("Segoe UI", 10.5, [System.Drawing.FontStyle]::Regular)
$whiteBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$subTitleBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(148, 163, 184)) # slate-400

$g.DrawString("SMILES for Office — 4-Step Interactive Workflow", $titleFont, $whiteBrush, 32, 14)
$g.DrawString("Draw in one click, modify anytime with a double click — seamless chemical structure workflow", $subTitleFont, $subTitleBrush, 34, 48)

# Cards layout
$cardY = 105
$cardW = 260
$cardH = 285
$startX = 30
$gap = 33

$steps = @(
    @{
        Step = "STEP 1"
        Title = "Launch & Trigger"
        Color = [System.Drawing.Color]::FromArgb(37, 99, 235) # Blue-600
        BorderColor = [System.Drawing.Color]::FromArgb(191, 219, 254) # Blue-200
        Desc1 = "In Excel cells, Word docs, or PPT:"
        Desc2 = "Press Ctrl+Shift+K or click the"
        Desc3 = "[SMILES] ribbon tab button"
    },
    @{
        Step = "STEP 2"
        Title = "2D Molecule Sketch"
        Color = [System.Drawing.Color]::FromArgb(13, 148, 136) # Teal-600
        BorderColor = [System.Drawing.Color]::FromArgb(153, 246, 228) # Teal-200
        Desc1 = "In embedded EPAM Ketcher:"
        Desc2 = "Draw structure with mouse freely"
        Desc3 = "or paste SMILES / MOL / KET"
    },
    @{
        Step = "STEP 3"
        Title = "Instant Insert"
        Color = [System.Drawing.Color]::FromArgb(22, 163, 74) # Green-600
        BorderColor = [System.Drawing.Color]::FromArgb(187, 247, 208) # Green-200
        Desc1 = "Click [Apply] button (Ctrl+Enter):"
        Desc2 = "Crisp vector molecular graphics"
        Desc3 = "inserted into active document"
    },
    @{
        Step = "STEP 4"
        Title = "Double-Click Edit"
        Color = [System.Drawing.Color]::FromArgb(147, 51, 234) # Purple-600
        BorderColor = [System.Drawing.Color]::FromArgb(233, 213, 255) # Purple-200
        Desc1 = "Need modifications anytime?"
        Desc2 = "Double-click the molecule image"
        Desc3 = "to re-open and update in-place"
    }
)

$stepFont = New-Object System.Drawing.Font ("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$cardTitleFont = New-Object System.Drawing.Font ("Segoe UI", 13.5, [System.Drawing.FontStyle]::Bold)
$descFont1 = New-Object System.Drawing.Font ("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
$descFont2 = New-Object System.Drawing.Font ("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$arrowFont = New-Object System.Drawing.Font ("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)

$darkTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))
$mutedTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 116, 139))
$arrowBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(148, 163, 184))
$divPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(241, 245, 249), 1.5)

$centerSf = New-Object System.Drawing.StringFormat
$centerSf.Alignment = [System.Drawing.StringAlignment]::Center
$centerSf.LineAlignment = [System.Drawing.StringAlignment]::Center

for ($i = 0; $i -lt $steps.Count; $i++) {
    $s = $steps[$i]
    $cx = $startX + $i * ($cardW + $gap)

    # Card background (White with clean shadow border)
    $cardRect = New-Object System.Drawing.Rectangle $cx, $cardY, $cardW, $cardH
    $g.FillRectangle([System.Drawing.Brushes]::White, $cardRect)

    $cardPen = New-Object System.Drawing.Pen ($s.BorderColor, 1.5)
    $g.DrawRectangle($cardPen, $cardRect)
    $cardPen.Dispose()

    # Step Badge Pill
    $badgeRect = New-Object System.Drawing.Rectangle ($cx + 18), ($cardY + 20), 74, 26
    $badgeBrush = New-Object System.Drawing.SolidBrush ($s.Color)
    $g.FillRectangle($badgeBrush, $badgeRect)
    $g.DrawString($s.Step, $stepFont, $whiteBrush, [System.Drawing.RectangleF]$badgeRect, $centerSf)
    $badgeBrush.Dispose()

    # Card Title
    $g.DrawString($s.Title, $cardTitleFont, $darkTextBrush, ($cx + 18), ($cardY + 58))

    # Divider line
    $g.DrawLine($divPen, ($cx + 18), ($cardY + 98), ($cx + $cardW - 18), ($cardY + 98))

    # Descriptions
    $g.DrawString($s.Desc1, $descFont1, $mutedTextBrush, ($cx + 18), ($cardY + 115))
    $g.DrawString($s.Desc2, $descFont2, $darkTextBrush, ($cx + 18), ($cardY + 145))
    $g.DrawString($s.Desc3, $descFont2, $darkTextBrush, ($cx + 18), ($cardY + 172))

    # Arrow to next step
    if ($i -lt ($steps.Count - 1)) {
        $arrowX = $cx + $cardW + 4
        $arrowRect = New-Object System.Drawing.Rectangle $arrowX, ($cardY + 120), 25, 30
        $g.DrawString([char]0x2192, $arrowFont, $arrowBrush, [System.Drawing.RectangleF]$arrowRect, $centerSf)
    }
}

# Save Image
$exportPath = [System.IO.Path]::GetFullPath($OutputPath)
$bmp.Save($exportPath, [System.Drawing.Imaging.ImageFormat]::Png)
Write-Host "[+] Generated English Workflow Graphic: $exportPath" -ForegroundColor Green

$docsPublicPath = (Join-Path $PSScriptRoot '../docs/public/workflow-ketcher.png')
if (Test-Path $exportPath) {
    Copy-Item -Path $exportPath -Destination $docsPublicPath -Force
    Write-Host "[+] Synced Workflow Graphic to docs/public: $docsPublicPath" -ForegroundColor Cyan
}

$g.Dispose()
$bmp.Dispose()
