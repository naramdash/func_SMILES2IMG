param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/powerpoint-showcase.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 1. Load Smiles2Img.Core for high-res molecule rendering
$coreDll = (Resolve-Path (Join-Path $PSScriptRoot "../dist/x64/Word/Smiles2Img.Core.dll")).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))

# Render sample molecules to temp files
$samples = @(
    @{ Name = "Aspirin"; Formula = "C9H8O4"; Smiles = "CC(=O)Oc1ccccc1C(=O)O"; Desc = "Analgesic • Anti-inflammatory" },
    @{ Name = "Caffeine"; Formula = "C8H10N4O2"; Smiles = "Cn1cnc2c1c(=O)n(C)c(=O)n2C"; Desc = "CNS Stimulant" },
    @{ Name = "Nicotine"; Formula = "C10H14N2"; Smiles = "c1cncc(c1)[C@@H]2CCCN2C"; Desc = "Pyridine-Pyrrolidine Alkaloid" },
    @{ Name = "Paracetamol"; Formula = "C8H9NO2"; Smiles = "CC(=O)Nc1ccc(O)cc1"; Desc = "Antipyretic Analgesic" }
)

$tempFiles = @()
foreach ($s in $samples) {
    $bytes = [Smiles2Img.Rendering.MoleculeRenderer]::Render($s.Smiles)
    $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ppt_smiles_$([Guid]::NewGuid().ToString('N')).png")
    [System.IO.File]::WriteAllBytes($tmp, $bytes)
    $s.TempPath = $tmp
    $tempFiles += $tmp
}

$msoTrue = -1
$msoFalse = 0

# 2. Automate PowerPoint
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue

try {
    $pres = $ppt.Presentations.Add($msoTrue)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

    # Background color (Clean dark slate or modern light)
    $bg = $slide.Background
    $bg.Fill.Solid()
    $bg.Fill.ForeColor.RGB = 0x1E1E2E # Sleek modern dark blue/slate (RGB: 46, 30, 30 in BGR)
    # Actually Office ForeColor.RGB is 0xBBGGRR -> for 0x0F172A (slate-900): B=2A, G=17, R=0F -> 0x2A170F

    # Title Banner
    $titleBox = $slide.Shapes.AddTextbox(1, 40, 25, 880, 50)
    $titleBox.TextFrame.TextRange.Text = "SMILES for PowerPoint — Chemical Structure Showcase"
    $titleBox.TextFrame.TextRange.Font.Name = "Segoe UI"
    $titleBox.TextFrame.TextRange.Font.Size = 24
    $titleBox.TextFrame.TextRange.Font.Bold = $msoTrue
    $titleBox.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

    $subBox = $slide.Shapes.AddTextbox(1, 40, 70, 880, 25)
    $subBox.TextFrame.TextRange.Text = "EPAM Indigo Vector Renderer & Ketcher 2D Integration • Instant In-Place Slide Editing"
    $subBox.TextFrame.TextRange.Font.Name = "Segoe UI"
    $subBox.TextFrame.TextRange.Font.Size = 13
    $subBox.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8

    # 4 Cards Layout
    $cardWidth = 205
    $cardHeight = 380
    $startY = 115
    $gap = 20
    $startX = 40

    for ($i = 0; $i -lt $samples.Count; $i++) {
        $item = $samples[$i]
        $cardX = $startX + ($i * ($cardWidth + $gap))

        # Card Background Shape
        $card = $slide.Shapes.AddShape(1, $cardX, $startY, $cardWidth, $cardHeight) # 1 = msoShapeRectangle
        $card.Fill.Solid()
        $card.Fill.ForeColor.RGB = 0x2D231E # Dark card background
        $card.Line.ForeColor.RGB = 0x473833 # Card border
        $card.Line.Weight = 1.2

        # Card Title
        $nameBox = $slide.Shapes.AddTextbox(1, $cardX + 10, $startY + 12, $cardWidth - 20, 28)
        $nameBox.TextFrame.TextRange.Text = $item.Name
        $nameBox.TextFrame.TextRange.Font.Name = "Segoe UI"
        $nameBox.TextFrame.TextRange.Font.Size = 13
        $nameBox.TextFrame.TextRange.Font.Bold = $msoTrue
        $nameBox.TextFrame.TextRange.Font.Color.RGB = 0x38BDF8 # Sky blue

        # Formula & Desc
        $infoBox = $slide.Shapes.AddTextbox(1, $cardX + 10, $startY + 40, $cardWidth - 20, 22)
        $infoBox.TextFrame.TextRange.Text = "$($item.Formula) • $($item.Desc)"
        $infoBox.TextFrame.TextRange.Font.Name = "Segoe UI"
        $infoBox.TextFrame.TextRange.Font.Size = 10
        $infoBox.TextFrame.TextRange.Font.Color.RGB = 0xCBD5E1

        # Molecule Image (White background box inside card)
        $imgBox = $slide.Shapes.AddShape(1, $cardX + 10, $startY + 70, $cardWidth - 20, 210)
        $imgBox.Fill.Solid()
        $imgBox.Fill.ForeColor.RGB = 0xFFFFFF # White canvas for molecule
        $imgBox.Line.Visible = $msoFalse

        # Add picture centered in imgBox
        $imgW = 175
        $imgH = 190
        $pic = $slide.Shapes.AddPicture(
            $item.TempPath,
            $msoFalse,
            $msoTrue,
            $cardX + 15,
            $startY + 80,
            $imgW,
            $imgH
        )
        $pic.AlternativeText = "SMILES:" + $item.Smiles
        $pic.Title = "SMILES:" + $item.Smiles
        try { $pic.Tags.Add("SMILES", $item.Smiles) } catch { }

        # SMILES text box at card bottom
        $smilesBox = $slide.Shapes.AddTextbox(1, $cardX + 10, $startY + 295, $cardWidth - 20, 70)
        $smilesBox.TextFrame.TextRange.Text = "SMILES:`n" + $item.Smiles
        $smilesBox.TextFrame.TextRange.Font.Name = "Consolas"
        $smilesBox.TextFrame.TextRange.Font.Size = 9
        $smilesBox.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8
    }

    # Save presentation and Export high-res PNG (1920x1080)
    $exportPath = [System.IO.Path]::GetFullPath($OutputPath)
    $slide.Export($exportPath, "PNG", 1920, 1080)
    Write-Host "[+] Saved PowerPoint Showcase: $exportPath" -ForegroundColor Green

    $docsPublicPath = (Join-Path $PSScriptRoot '../docs/public/powerpoint-showcase.png')
    if (Test-Path $exportPath) {
        Copy-Item -Path $exportPath -Destination $docsPublicPath -Force
        Write-Host "[+] Synced PowerPoint Showcase to docs/public: $docsPublicPath" -ForegroundColor Cyan
    }

    $pres.Close()
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    foreach ($tmp in $tempFiles) {
        try { if (Test-Path $tmp) { Remove-Item $tmp -Force } } catch { }
    }
}
