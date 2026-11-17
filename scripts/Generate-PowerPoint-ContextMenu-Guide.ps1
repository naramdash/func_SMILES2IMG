param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/powerpoint-context-menu-guide.png')
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

$pptOrange = [System.Drawing.Color]::FromArgb(196, 62, 28)

# 1. PowerPoint Workspace & Dark Mode Slide
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(40, 40, 40))), 0, 0, $width, $height)
$g.FillRectangle((New-Object System.Drawing.SolidBrush($pptOrange)), 0, 0, $width, 32)
$g.DrawString('Screening_Deck.pptx - PowerPoint', (New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)), [System.Drawing.Brushes]::White, 16, 7)

$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(49, 49, 49))), 0, 32, 92, ($height - 32))
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(15, 23, 42))), 14, 48, 66, 38)
$g.DrawRectangle((New-Object System.Drawing.Pen($pptOrange, 2)), 14, 48, 66, 38)

$slideX = 114; $slideY = 48; $slideW = 780; $slideH = 508
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(15, 23, 42))), $slideX, $slideY, $slideW, $slideH)
$g.DrawRectangle((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(51, 65, 85), 1)), $slideX, $slideY, $slideW, $slideH)

$g.DrawString('Lead Compound Candidates — Bioactive Screening Deck', (New-Object System.Drawing.Font('Segoe UI', 13.5, [System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 250, 252))), ($slideX + 24), ($slideY + 18))
$g.DrawString('func_SMILES2IMG • EPAM Indigo Vector Renderer & Ketcher 2D Editor', (New-Object System.Drawing.Font('Segoe UI', 8.5)), (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(148, 163, 184))), ($slideX + 24), ($slideY + 42))

$showcasePath = Resolve-Path (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
$sb = [System.Drawing.Bitmap]::FromFile($showcasePath.Path)
$cards = @(
    @{ X = ($slideX + 24); Title = 'Aspirin'; Sub = 'C9H8O4 • Analgesic'; Src = (New-Object System.Drawing.Rectangle(98, 368, 372, 418)) },
    @{ X = ($slideX + 400); Title = 'Caffeine'; Sub = 'C8H10N4O2 • CNS Stimulant'; Src = (New-Object System.Drawing.Rectangle(550, 368, 372, 418)) }
)
foreach ($c in $cards) {
    $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 41, 59))), $c.X, ($slideY + 72), 352, 408)
    $g.DrawRectangle((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(71, 85, 105), 1)), $c.X, ($slideY + 72), 352, 408)
    $g.DrawString($c.Title, (New-Object System.Drawing.Font('Segoe UI', 11.0, [System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(250, 204, 21))), ($c.X + 16), ($slideY + 86))
    $g.DrawString($c.Sub, (New-Object System.Drawing.Font('Segoe UI', 8.5)), (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(148, 163, 184))), ($c.X + 16), ($slideY + 108))
    $g.FillRectangle([System.Drawing.Brushes]::White, ($c.X + 16), ($slideY + 134), 320, 280)
    $g.DrawImage($sb, (New-Object System.Drawing.Rectangle(($c.X + 36), ($slideY + 144), 280, 260)), $c.Src, [System.Drawing.GraphicsUnit]::Pixel)
}
$sb.Dispose()

# 2. Cascading PowerPoint Picture Context Menus (matches BuildObjectContextMenu -> insertBeforeMso="Cut")
$mainItems = @(
    @{ Label = 'Chemical Structure'; Icon = 'chem'; Highlight = $true; HasSub = $true; Key = 'chem' },
    @{ Kind = 'sep' },
    @{ Label = 'Cut'; Icon = 'cut'; Shortcut = 'Ctrl+X' },
    @{ Label = 'Copy'; Icon = 'copy'; Shortcut = 'Ctrl+C' },
    @{ Label = 'Paste Options:'; Muted = $true },
    @{ Kind = 'sep' },
    @{ Label = 'Bring to Front'; HasSub = $true },
    @{ Label = 'Send to Back'; HasSub = $true },
    @{ Label = 'Save as Picture...' },
    @{ Label = 'Format Picture...' }
)
$res1 = Draw-OfficeContextMenuPanel -g $g -X 84 -Y 108 -Width 216 -Items $mainItems -AccentColor $pptOrange -MiniToolbar @{ Mode = 'picture' } -ShowSearchBox $true

$blue = [System.Drawing.Color]::FromArgb(37, 99, 235)
$teal = [System.Drawing.Color]::FromArgb(13, 148, 136)
$chemItems = @(
    @{ Label = 'Insert / Edit Structure...'; Icon = 'pen'; IconColor = $pptOrange },
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
$sub1Y = [math]::Max(120, $res1.Anchors['chem'].Y - 6)
$res2 = Draw-OfficeContextMenuPanel -g $g -X 298 -Y $sub1Y -Width 248 -Items $chemItems -AccentColor $pptOrange

$infoItems = Get-SharedMoleculeInfoContextItems
$sub2Y = [math]::Min(($height - 315), [math]::Max(200, $res2.Anchors['info'].Y - 190))
[void](Draw-OfficeContextMenuPanel -g $g -X 544 -Y $sub2Y -Width 268 -Items $infoItems -AccentColor $pptOrange)

$g.Dispose()
foreach ($dest in @($OutputPath, (Join-Path $PSScriptRoot '../docs/public/powerpoint-context-menu-guide.png'))) {
    $dir = Split-Path $dest -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $bmp.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
}
$bmp.Dispose()
Write-Host "[+] Generated v1.8.0 PowerPoint context menu guide at $OutputPath" -ForegroundColor Green
