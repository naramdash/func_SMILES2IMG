param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/word-context-menu-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
. (Join-Path $PSScriptRoot 'Draw-ContextMenuFlyout.ps1')

$width = 920
$height = 560
$bmp = New-Object System.Drawing.Bitmap($width, $height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

$wordBlue = [System.Drawing.Color]::FromArgb(24, 90, 189)
$brushDark = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 41, 59))
$brushMuted = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100, 116, 139))

# 1. Word Top Header & Document Surface
$g.FillRectangle((New-Object System.Drawing.SolidBrush($wordBlue)), 0, 0, $width, 30)
$g.DrawString('Compound_Summary.docx - Word', (New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)), [System.Drawing.Brushes]::White, 16, 6)

$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(241, 245, 249))), 0, 30, $width, ($height - 30))

# White page sheet
$pageRect = New-Object System.Drawing.Rectangle(34, 44, ($width - 68), ($height - 58))
$g.FillRectangle([System.Drawing.Brushes]::White, $pageRect)
$g.DrawRectangle((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(203, 213, 225), 1)), $pageRect)

$g.DrawString('Chemical Compound Research Summary (func_SMILES2IMG)', (New-Object System.Drawing.Font('Segoe UI', 13.0, [System.Drawing.FontStyle]::Bold)), $brushDark, 64, 64)
$g.DrawString('EPAM Indigo Vector Renderer & Ketcher 2D Editor Integration', (New-Object System.Drawing.Font('Segoe UI', 9.0)), $brushMuted, 64, 88)

# Word Table behind context menus
$tblBorder = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(203, 213, 225), 1)
$hdrRect = New-Object System.Drawing.Rectangle(64, 114, 792, 28)
$g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 250, 252))), $hdrRect)
$g.DrawRectangle($tblBorder, 64, 114, 792, 300)
$g.DrawLine($tblBorder, 64, 142, 856, 142)
$g.DrawLine($tblBorder, 64, 276, 856, 276)
$g.DrawLine($tblBorder, 200, 114, 200, 414)
$g.DrawLine($tblBorder, 320, 114, 320, 414)
$g.DrawLine($tblBorder, 540, 114, 540, 414)

$fTh = New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
$fCode = New-Object System.Drawing.Font('Consolas', 8.5)
$g.DrawString('Compound', $fTh, $brushDark, 76, 120)
$g.DrawString('Formula', $fTh, $brushDark, 212, 120)
$g.DrawString('2D Structure (Double-Click)', $fTh, $brushDark, 336, 120)
$g.DrawString('SMILES Code', $fTh, $brushDark, 556, 120)

$g.DrawString('Aspirin', $fTh, $brushDark, 76, 200)
$g.DrawString('C9H8O4', (New-Object System.Drawing.Font('Segoe UI', 9.0)), $brushDark, 212, 200)
$g.DrawString('CC(=O)Oc1ccccc1C(=O)O', $fCode, $brushMuted, 556, 200)
$g.DrawString('Caffeine', $fTh, $brushDark, 76, 338)
$g.DrawString('C8H10N4O2', (New-Object System.Drawing.Font('Segoe UI', 9.0)), $brushDark, 212, 338)
$g.DrawString('Cn1cnc2c1c(=O)n(C)c(=O)n2C', $fCode, $brushMuted, 556, 338)

$showcasePath = Resolve-Path (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
$sb = [System.Drawing.Bitmap]::FromFile($showcasePath.Path)
$g.DrawImage($sb, (New-Object System.Drawing.Rectangle(372, 152, 114, 114)), (New-Object System.Drawing.Rectangle(98, 368, 372, 418)), [System.Drawing.GraphicsUnit]::Pixel)
$g.DrawImage($sb, (New-Object System.Drawing.Rectangle(372, 288, 114, 114)), (New-Object System.Drawing.Rectangle(550, 368, 372, 418)), [System.Drawing.GraphicsUnit]::Pixel)
$sb.Dispose()

# Selection rectangle on Aspirin image
$g.DrawRectangle((New-Object System.Drawing.Pen($wordBlue, 1.5)), 372, 152, 114, 114)

# 2. Cascading Word Context Menus (matches BuildObjectContextMenu -> insertBeforeMso="Cut")
$mainItems = @(
    @{ Label = 'Chemical Structure'; Icon = 'chem'; Highlight = $true; HasSub = $true; Key = 'chem' },
    @{ Kind = 'sep' },
    @{ Label = 'Cut'; Icon = 'cut'; Shortcut = 'Ctrl+X' },
    @{ Label = 'Copy'; Icon = 'copy'; Shortcut = 'Ctrl+C' },
    @{ Label = 'Paste Options:'; Muted = $true },
    @{ Kind = 'sep' },
    @{ Label = 'Wrap Text'; HasSub = $true },
    @{ Label = 'Edit Picture...' },
    @{ Label = 'Format Picture...' }
)
$res1 = Draw-OfficeContextMenuPanel -g $g -X 62 -Y 108 -Width 216 -Items $mainItems -AccentColor $wordBlue -MiniToolbar @{ Mode = 'picture' } -ShowSearchBox $true

$blue = [System.Drawing.Color]::FromArgb(37, 99, 235)
$teal = [System.Drawing.Color]::FromArgb(13, 148, 136)
$chemItems = @(
    @{ Label = 'Insert / Edit Structure...'; Icon = 'pen'; IconColor = $wordBlue },
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
$res2 = Draw-OfficeContextMenuPanel -g $g -X 276 -Y $sub1Y -Width 248 -Items $chemItems -AccentColor $wordBlue

$infoItems = Get-SharedMoleculeInfoContextItems
$sub2Y = [math]::Min(($height - 315), [math]::Max(190, $res2.Anchors['info'].Y - 190))
[void](Draw-OfficeContextMenuPanel -g $g -X 522 -Y $sub2Y -Width 268 -Items $infoItems -AccentColor $wordBlue)

$g.Dispose()
foreach ($dest in @($OutputPath, (Join-Path $PSScriptRoot '../docs/public/word-context-menu-guide.png'))) {
    $dir = Split-Path $dest -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $bmp.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
}
$bmp.Dispose()
Write-Host "[+] Generated v1.8.0 Word context menu guide at $OutputPath" -ForegroundColor Green
