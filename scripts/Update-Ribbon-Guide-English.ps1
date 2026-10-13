param(
    [string]$ImagePath = (Join-Path $PSScriptRoot '../assets/ribbon-guide.png'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/ribbon-guide.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$bmp = [System.Drawing.Bitmap]::FromFile((Resolve-Path $ImagePath).Path)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

# 1. Background color for title and tab bar
$bgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(90, 81, 75))

# Clear title bar text area (Y: 0 to 24)
$g.FillRectangle($bgBrush, 24, 0, 80, 24)   # "AutoSave" area
$g.FillRectangle($bgBrush, 140, 0, 360, 24) # Title area (clean all Korean artifacts)

# Clear tab row (Y: 25 to 52, X: 12 to 555)
$g.FillRectangle($bgBrush, 12, 25, 545, 27)

# Fonts & Brushes
$titleFont = New-Object System.Drawing.Font ("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
$tabFont = New-Object System.Drawing.Font ("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular)
$textBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 235, 235))
$mutedBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(195, 195, 195))

# Draw English Title Bar text
$g.DrawString("AutoSave", $titleFont, $textBrush, 30, 6)
$g.DrawString("Book1 - Excel", $titleFont, $mutedBrush, 245, 6)

# Draw English Tabs with balanced spacing
$tabs = @(
    @{ Name = "File"; X = 18 },
    @{ Name = "Home"; X = 54 },
    @{ Name = "Insert"; X = 102 },
    @{ Name = "Draw"; X = 150 },
    @{ Name = "Page Layout"; X = 194 },
    @{ Name = "Formulas"; X = 278 },
    @{ Name = "Data"; X = 344 },
    @{ Name = "Review"; X = 386 },
    @{ Name = "View"; X = 442 },
    @{ Name = "Automate"; X = 486 }
)

foreach ($t in $tabs) {
    $g.DrawString($t.Name, $tabFont, $textBrush, $t.X, 33)
}

$g.Dispose()

# Save to destination
$tmpPath = [System.IO.Path]::GetTempFileName() + ".png"
$bmp.Save($tmpPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

Move-Item -Path $tmpPath -Destination $OutputPath -Force
Write-Host "[+] Updated Ribbon Guide with English UI: $OutputPath" -ForegroundColor Green

# Sync to docs/public
$docsPublicPath = (Join-Path $PSScriptRoot '../docs/public/ribbon-guide.png')
Copy-Item -Path $OutputPath -Destination $docsPublicPath -Force
Write-Host "[+] Synced Ribbon Guide to docs/public: $docsPublicPath" -ForegroundColor Cyan
