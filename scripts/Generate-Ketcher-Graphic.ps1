Add-Type -AssemblyName System.Drawing

$width = 1200
$height = 680
$bmp = New-Object System.Drawing.Bitmap $width, $height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

# Outer Background
$g.Clear([System.Drawing.Color]::FromArgb(241, 245, 249))

# Window Container
$winRect = New-Object System.Drawing.Rectangle 30, 25, 1140, 630
$winBorderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 2)
$g.FillRectangle([System.Drawing.Brushes]::White, $winRect)
$g.DrawRectangle($winBorderPen, $winRect)

# Window Title Bar
$titleRect = New-Object System.Drawing.Rectangle 30, 25, 1140, 50
$titleBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))
$g.FillRectangle($titleBrush, $titleRect)

$winTitleFont = New-Object System.Drawing.Font ("Malgun Gothic", 12, [System.Drawing.FontStyle]::Bold)
$g.DrawString("🧬 EPAM Ketcher 2D 화학 분자 에디터 — [ Excel · Word · PowerPoint 공통 ]", $winTitleFont, [System.Drawing.Brushes]::White, 45, 38)

# Top Action Bar inside Ketcher
$actionBarRect = New-Object System.Drawing.Rectangle 30, 75, 1140, 60
$actionBgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(248, 250, 252))
$actionDivPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1.5)
$g.FillRectangle($actionBgBrush, $actionBarRect)
$g.DrawLine($actionDivPen, 30, 135, 1170, 135)

# Callout / Help text in action bar
$guideFont = New-Object System.Drawing.Font ("Malgun Gothic", 10, [System.Drawing.FontStyle]::Bold)
$subGuideFont = New-Object System.Drawing.Font ("Malgun Gothic", 9, [System.Drawing.FontStyle]::Regular)
$blueTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(37, 99, 235))
$mutedTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 116, 139))
$g.DrawString("💡 2D 캔버스에 구조를 그리거나 SMILES/MOL을 붙여넣은 후, 우측 상단의 [적용]을 누르세요!", $guideFont, $blueTextBrush, 50, 85)
$g.DrawString("※ 100% 로컬 PC에서 오프라인 구동되며 기업 기밀 화학식이 외부로 전송되지 않습니다.", $subGuideFont, $mutedTextBrush, 52, 108)

# Big Green [적용] Button
$btnApplyRect = New-Object System.Drawing.Rectangle 920, 84, 140, 42
$btnApplyBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(22, 163, 74))
$g.FillRectangle($btnApplyBrush, $btnApplyRect)
$btnFont = New-Object System.Drawing.Font ("Malgun Gothic", 11, [System.Drawing.FontStyle]::Bold)
$sf = New-Object System.Drawing.StringFormat
$sf.Alignment = [System.Drawing.StringAlignment]::Center
$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
$g.DrawString("✅ 문서에 적용", $btnFont, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]$btnApplyRect, $sf)

# Gray [취소] Button
$btnCancelRect = New-Object System.Drawing.Rectangle 1075, 84, 80, 42
$btnCancelBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(226, 232, 240))
$btnCancelTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))
$g.FillRectangle($btnCancelBrush, $btnCancelRect)
$g.DrawString("취소", $btnFont, $btnCancelTextBrush, [System.Drawing.RectangleF]$btnCancelRect, $sf)

# Left Toolbar (Atoms)
$leftToolRect = New-Object System.Drawing.Rectangle 30, 135, 60, 520
$toolbarBgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(248, 250, 252))
$toolDivPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(226, 232, 240), 1)
$toolBorderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 1)
$g.FillRectangle($toolbarBgBrush, $leftToolRect)
$g.DrawLine($toolDivPen, 90, 135, 90, 655)

$atoms = @("H", "C", "N", "O", "S", "P", "F", "Cl", "Br", "I")
$atomFont = New-Object System.Drawing.Font ("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$darkBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 41, 59))

for ($i = 0; $i -lt $atoms.Count; $i++) {
    $ay = 150 + $i * 45
    $aRect = New-Object System.Drawing.Rectangle 40, $ay, 40, 35
    $g.FillRectangle([System.Drawing.Brushes]::White, $aRect)
    $g.DrawRectangle($toolBorderPen, $aRect)
    $g.DrawString($atoms[$i], $atomFont, $darkBrush, [System.Drawing.RectangleF]$aRect, $sf)
}

# Top Toolbar (Bonds & Templates)
$topToolRect = New-Object System.Drawing.Rectangle 90, 135, 1080, 45
$g.FillRectangle($toolbarBgBrush, $topToolRect)
$g.DrawLine($toolDivPen, 90, 180, 1170, 180)

$tools = @("단일결합 ─", "이중결합 ═", "삼중결합 ≡", "입체결합 ◢", "점선결합 ▰", "벤젠고리 ⬡", "사이클로펜탄 ⬠", "지우개 ⌫", "되돌리기 ↩")
$toolFont = New-Object System.Drawing.Font ("Malgun Gothic", 9.5, [System.Drawing.FontStyle]::Regular)
for ($t = 0; $t -lt $tools.Count; $t++) {
    $tx = 105 + $t * 115
    $tRect = New-Object System.Drawing.Rectangle $tx, 142, 105, 30
    $g.FillRectangle([System.Drawing.Brushes]::White, $tRect)
    $g.DrawRectangle($toolBorderPen, $tRect)
    $g.DrawString($tools[$t], $toolFont, $darkBrush, [System.Drawing.RectangleF]$tRect, $sf)
}

# Center Canvas: Render Molecule (Aspirin)
$canvasRect = New-Object System.Drawing.Rectangle 90, 180, 1080, 440
$g.FillRectangle([System.Drawing.Brushes]::White, $canvasRect)

# Grid Dots
$dotBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(241, 245, 249))
for ($gx = 110; $gx -lt 1150; $gx += 40) {
    for ($gy = 200; $gy -lt 600; $gy += 40) {
        $g.FillEllipse($dotBrush, $gx, $gy, 3, 3)
    }
}

# Render Molecule in center using Smiles2Img.Core
$coreDll = (Resolve-Path (Join-Path $PSScriptRoot "../dist/x64/Word/Smiles2Img.Core.dll")).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))
$bytes = [Smiles2Img.Rendering.MoleculeRenderer]::Render("CC(=O)Oc1ccccc1C(=O)O")
$ms = New-Object System.IO.MemoryStream (,$bytes)
$molImg = [System.Drawing.Image]::FromStream($ms)
$destX = 400
$destY = 230
$destW = 440
$destH = 330
$g.DrawImage($molImg, $destX, $destY, $destW, $destH)
$molImg.Dispose()
$ms.Dispose()

# Label on Molecule
$labelFont = New-Object System.Drawing.Font ("Malgun Gothic", 11, [System.Drawing.FontStyle]::Bold)
$g.DrawString("Aspirin (아스피린) — CC(=O)Oc1ccccc1C(=O)O", $labelFont, $mutedTextBrush, 460, 580)

# Bottom Status Footer inside Window
$footerRect = New-Object System.Drawing.Rectangle 30, 620, 1140, 35
$footerBgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(241, 245, 249))
$footerFont = New-Object System.Drawing.Font ("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$g.FillRectangle($footerBgBrush, $footerRect)
$g.DrawString("Ready | Smiles2Img Office Add-In v1.3.0 | EPAM Ketcher v3.1.0 Offline Virtual Host", $footerFont, $mutedTextBrush, 45, 630)

$outPath = (Join-Path $PSScriptRoot "../assets/ketcher-guide.png")
$bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()

Write-Host "[+] Generated ketcher-guide.png: $outPath"
