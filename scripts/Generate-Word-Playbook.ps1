<#
.SYNOPSIS
    Generates the official self-contained func_SMILES2IMG Microsoft Word Interactive Playbook (.docx)
    using live Microsoft Word + Smiles2Img.Word.AddIn so every 2D structure, KET reaction scheme,
    and Mol* 3D figure carries genuine embedded AlternativeText & CustomXMLParts metadata.
#>
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\examples\func_SMILES2IMG_Word_Playbook.docx'),
    [string]$PreviewPdfPath = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing, System.Windows.Forms, System.IO.Compression.FileSystem

$repoRoot = Split-Path -Parent $PSScriptRoot
$targetFile = [System.IO.Path]::GetFullPath($OutputPath)
$docsPublicFile = [System.IO.Path]::GetFullPath((Join-Path $repoRoot 'docs\public\examples\func_SMILES2IMG_Word_Playbook.docx'))
New-Item -ItemType Directory -Path (Split-Path -Parent $targetFile) -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $docsPublicFile) -Force | Out-Null

# Load Smiles2Img.Core + NativeIndigo
$coreDll = (Resolve-Path (Join-Path $repoRoot 'dist\x64\Word\Smiles2Img.Core.dll')).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))

$addinVersion = '1.8.0'
try {
    $propsXml = [xml](Get-Content (Join-Path $repoRoot 'Directory.Build.props') -Raw)
    if ($propsXml.Project.PropertyGroup.Version) { $addinVersion = [string]$propsXml.Project.PropertyGroup.Version }
} catch {}

Write-Host "=== Generating func_SMILES2IMG v$addinVersion — Self-Contained Microsoft Word Playbook ===" -ForegroundColor Cyan
Write-Host "Output : $targetFile" -ForegroundColor Gray

if (Test-Path -LiteralPath $targetFile) { Remove-Item -LiteralPath $targetFile -Force }

function Convert-DarkBgToTransparent([byte[]]$pngBytes) {
    $msIn = New-Object System.IO.MemoryStream(, $pngBytes)
    $src = [System.Drawing.Bitmap]::FromStream($msIn)
    $dst = New-Object System.Drawing.Bitmap($src.Width, $src.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $bg = $src.GetPixel(4, 4)
        for ($y = 0; $y -lt $src.Height; $y++) {
            for ($x = 0; $x -lt $src.Width; $x++) {
                $px = $src.GetPixel($x, $y)
                $d = [Math]::Abs([int]$px.R - [int]$bg.R) + [Math]::Abs([int]$px.G - [int]$bg.G) + [Math]::Abs([int]$px.B - [int]$bg.B)
                if ($d -le 18) { $dst.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, 0, 0, 0)) }
                else { $dst.SetPixel($x, $y, $px) }
            }
        }
        $msOut = New-Object System.IO.MemoryStream
        $dst.Save($msOut, [System.Drawing.Imaging.ImageFormat]::Png)
        $res = $msOut.ToArray(); $msOut.Dispose()
        return ,$res
    } finally {
        $dst.Dispose(); $src.Dispose(); $msIn.Dispose()
    }
}

function Get-1CrnPngBytes([IntPtr]$ownerHwnd) {
    $diskCache = Join-Path ([System.IO.Path]::GetTempPath()) 's2i_cached_1crn_trans.png'
    if (Test-Path -LiteralPath $diskCache) {
        $b = [System.IO.File]::ReadAllBytes($diskCache)
        if ($b.Length -gt 8000) { return ,$b }
    }
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $form = New-Object Smiles2Img.MolstarForm('Word', '1CRN', $null)
        $form.AutoApplyAfterLoadForTest = $true
        $owner = New-Object Smiles2Img.Win32Window($ownerHwnd)
        $t = New-Object System.Windows.Forms.Timer
        $t.Tag = $form; $t.Interval = 15000
        $t.Add_Tick({ param($s,$e); $s.Stop(); try { $s.Tag.Close() } catch {} })
        $t.Start()
        [void]$form.ShowDialog($owner)
        $t.Dispose()
        $img = $form.ResultImageBytes
        if (-not $img) { $img = $form.ResultUiViewportPngBytes }
        $form.Dispose()
        if ($img -and $img.Length -gt 8000) {
            $trans = Convert-DarkBgToTransparent $img
            [System.IO.File]::WriteAllBytes($diskCache, $trans)
            return ,$trans
        }
        Start-Sleep -Milliseconds 500
    }
    throw 'Could not obtain live 1CRN PNG from MolstarForm.'
}

function Add-Para {
    param(
        [object]$doc,
        [string]$text,
        [double]$size = 10,
        [bool]$bold = $false,
        [bool]$italic = $false,
        [int]$colorBgr = 0x262626,
        [double]$spaceBefore = 2,
        [double]$spaceAfter = 5,
        [string]$fontName = 'Segoe UI',
        [int]$align = 0 # 0=left, 1=center
    )
    $p = $doc.Paragraphs.Add()
    $p.Borders.Enable = 0
    $p.Range.Text = $text
    $p.Range.Font.Name = $fontName
    $p.Range.Font.Size = $size
    $p.Range.Font.Bold = if ($bold) { 1 } else { 0 }
    $p.Range.Font.Italic = if ($italic) { 1 } else { 0 }
    $p.Range.Font.Color = $colorBgr
    $p.Format.SpaceBefore = $spaceBefore
    $p.Format.SpaceAfter = $spaceAfter
    $p.Format.Alignment = $align
    $p.Range.InsertParagraphAfter()
}

function Add-ChapterHeading([object]$doc, [string]$numLabel, [string]$title, [bool]$pageBreakBefore = $false) {
    if ($pageBreakBefore) {
        $pBr = $doc.Paragraphs.Add()
        $pBr.Borders.Enable = 0
        $pBr.Format.Alignment = 0
        $pBr.Range.InsertBreak(7) # wdPageBreak
    }
    $p = $doc.Paragraphs.Add()
    $p.Range.Text = "$numLabel  $title"
    $p.Range.Font.Name = 'Segoe UI'
    $p.Range.Font.Size = 14
    $p.Range.Font.Bold = 1
    $p.Range.Font.Italic = 0
    $p.Range.Font.Color = 0x6E760F # Teal #0F766E in BGR
    $p.Format.Alignment = 0
    $p.Format.SpaceBefore = 6
    $p.Format.SpaceAfter = 6
    $p.Borders.Enable = 0
    $p.Borders.Item(-3).LineStyle = 1 # wdBorderBottom
    $p.Borders.Item(-3).LineWidth = 12 # 1.5pt
    $p.Borders.Item(-3).Color = 0xB8C499
    $p.Range.InsertParagraphAfter()
    # Explicitly clear inherited bottom border on the trailing paragraph
    $lastP = $doc.Paragraphs.Item($doc.Paragraphs.Count)
    $lastP.Borders.Enable = 0
    $lastP.Format.Alignment = 0
}

function Add-CalloutBox {
    param(
        [object]$doc,
        [string]$heading,
        [string[]]$lines,
        [int]$accentBgr = 0x6E760F,
        [int]$bgBgr = 0xF7FBF9
    )
    $pAnchor = $doc.Paragraphs.Add()
    $pAnchor.Borders.Enable = 0
    $pAnchor.Format.Alignment = 0
    $tbl = $doc.Tables.Add($pAnchor.Range, 1, 1)
    $tbl.AllowAutoFit = $false
    $tbl.PreferredWidthType = 3 # wdPreferredWidthPoints
    $tbl.PreferredWidth = 504
    $cell = $tbl.Cell(1, 1)
    $cell.Shading.BackgroundPatternColor = $bgBgr
    $cell.TopPadding = 5; $cell.BottomPadding = 5; $cell.LeftPadding = 10; $cell.RightPadding = 10
    $tbl.Borders.Enable = 0
    $cell.Borders.Item(-2).LineStyle = 1 # wdBorderLeft
    $cell.Borders.Item(-2).LineWidth = 24 # 3pt
    $cell.Borders.Item(-2).Color = $accentBgr
    $cell.Borders.Item(-1).LineStyle = 1; $cell.Borders.Item(-1).LineWidth = 4; $cell.Borders.Item(-1).Color = 0xE2E8F0
    $cell.Borders.Item(-3).LineStyle = 1; $cell.Borders.Item(-3).LineWidth = 4; $cell.Borders.Item(-3).Color = 0xE2E8F0
    $cell.Borders.Item(-4).LineStyle = 1; $cell.Borders.Item(-4).LineWidth = 4; $cell.Borders.Item(-4).Color = 0xE2E8F0

    $bodyText = $heading + "`r" + ($lines -join "`r")
    $cell.Range.Text = $bodyText
    $cell.Range.Font.Name = 'Segoe UI'
    $cell.Range.Font.Size = 9.2
    $cell.Range.Font.Bold = 0
    $cell.Range.Font.Italic = 0
    $cell.Range.Font.Color = 0x262626
    $cell.Range.ParagraphFormat.Alignment = 0
    $cell.Range.ParagraphFormat.SpaceBefore = 1
    $cell.Range.ParagraphFormat.SpaceAfter = 2.5
    $cell.Range.Paragraphs.Item(1).Range.Font.Bold = 1
    $cell.Range.Paragraphs.Item(1).Range.Font.Italic = 0
    $cell.Range.Paragraphs.Item(1).Range.Font.Size = 9.8
    $cell.Range.Paragraphs.Item(1).Range.Font.Color = $accentBgr

    $pAfter = $doc.Paragraphs.Add()
    $pAfter.Borders.Enable = 0
    $pAfter.Format.Alignment = 0
    $pAfter.Format.SpaceBefore = 0; $pAfter.Format.SpaceAfter = 3
    $pAfter.Range.InsertParagraphAfter()
}

function Add-StyledWordTable {
    param(
        [object]$doc,
        [string[]]$headers,
        [array]$rows,
        [double[]]$colWidthsPt,
        [int]$headerBgBgr = 0x6E760F,
        [double]$bodyFontSize = 9,
        [double]$cellPadV = 4
    )
    $pAnchor = $doc.Paragraphs.Add()
    $pAnchor.Borders.Enable = 0
    $pAnchor.Format.Alignment = 0
    $tbl = $doc.Tables.Add($pAnchor.Range, ($rows.Length + 1), $headers.Length)
    $tbl.AllowAutoFit = $false
    $tbl.Borders.Enable = 1
    $tbl.Borders.OutsideColor = 0xCBD5E1
    $tbl.Borders.InsideColor = 0xE2E8F0
    $tbl.Rows.Item(1).HeadingFormat = -1

    for ($c = 0; $c -lt $headers.Length; $c++) {
        $cell = $tbl.Cell(1, $c + 1)
        $cell.Width = $colWidthsPt[$c]
        $cell.Shading.BackgroundPatternColor = $headerBgBgr
        $cell.TopPadding = $cellPadV; $cell.BottomPadding = $cellPadV; $cell.LeftPadding = 6; $cell.RightPadding = 6
        $cell.Range.Text = $headers[$c]
        $cell.Range.Font.Name = 'Segoe UI'
        $cell.Range.Font.Size = 9.2
        $cell.Range.Font.Bold = 1
        $cell.Range.Font.Italic = 0
        $cell.Range.Font.Color = 0xFFFFFF
        $cell.Range.ParagraphFormat.Alignment = 0
        $cell.Range.ParagraphFormat.SpaceBefore = 0
        $cell.Range.ParagraphFormat.SpaceAfter = 0
    }

    for ($r = 0; $r -lt $rows.Length; $r++) {
        $bg = if ($r % 2 -eq 1) { 0xF8FAFC } else { 0xFFFFFF }
        for ($c = 0; $c -lt $headers.Length; $c++) {
            $cell = $tbl.Cell($r + 2, $c + 1)
            $cell.Width = $colWidthsPt[$c]
            $cell.Shading.BackgroundPatternColor = $bg
            $cell.TopPadding = $cellPadV; $cell.BottomPadding = $cellPadV; $cell.LeftPadding = 6; $cell.RightPadding = 6
            $cell.VerticalAlignment = 1 # wdCellAlignVerticalCenter
            $cell.Range.Text = [string]$rows[$r][$c]
            $cell.Range.Font.Name = 'Segoe UI'
            $cell.Range.Font.Size = $bodyFontSize
            $cell.Range.Font.Bold = if ($c -eq 0) { 1 } else { 0 }
            $cell.Range.Font.Italic = 0
            $cell.Range.Font.Color = 0x1E293B
            $cell.Range.ParagraphFormat.Alignment = 0
            $cell.Range.ParagraphFormat.SpaceBefore = 0
            $cell.Range.ParagraphFormat.SpaceAfter = 0
        }
    }

    $pAfter = $doc.Paragraphs.Add()
    $pAfter.Borders.Enable = 0
    $pAfter.Format.Alignment = 0
    $pAfter.Format.SpaceBefore = 0; $pAfter.Format.SpaceAfter = 3
    $pAfter.Range.InsertParagraphAfter()
    return $tbl
}

function Insert-LiveWordStructure {
    param(
        [object]$wordApp,
        [object]$wordAddInObj,
        [object]$targetRange,
        [string]$clipboardPayload,
        [double]$widthPt = 115
    )
    $targetRange.Select()
    $wordApp.Selection.Collapse(1) # wdCollapseStart
    [void][Smiles2Img.Editor.SafeClipboard]::SetText($clipboardPayload)
    Start-Sleep -Milliseconds 120
    [void]$wordAddInObj.ExecuteRibbonAction('OnQuickInsert')
    Start-Sleep -Milliseconds 350
    $doc = $wordApp.ActiveDocument
    if ($doc.InlineShapes.Count -gt 0) {
        $shp = $doc.InlineShapes.Item($doc.InlineShapes.Count)
        try { $shp.LockAspectRatio = -1 } catch {}
        $shp.Width = $widthPt
        return $shp
    }
    return $null
}

# Pre-build KET reactions & real descriptor summaries via NativeIndigo
$aspirinSmiles     = 'CC(=O)Oc1ccccc1C(=O)O'
$caffeineSmiles    = 'Cn1cnc2c1c(=O)n(C)c(=O)n2C'
$ibuprofenSmiles   = 'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O'
$paracetamolSmiles = 'CC(=O)Nc1ccc(O)cc1'
$ketAspirinRxn     = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('O=C(O)c1ccccc1O.CC(=O)OC(=O)C>>CC(=O)Oc1ccccc1C(=O)O')
$ketParacetamolRxn = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('Nc1ccc(O)cc1.CC(=O)OC(=O)C>>CC(=O)Nc1ccc(O)cc1')

function Get-CompoundProps([string]$name, [string]$smiles) {
    $info = [Smiles2Img.Rendering.MoleculeRenderer]::TryGetMoleculeInfo($smiles, $false)
    return @{
        Name     = $name
        Smiles   = $smiles
        Formula  = $info.Formula
        MW       = [Math]::Round($info.MolecularWeight, 2)
        LogP     = [Math]::Round($info.LogP, 2)
        TPSA     = [Math]::Round($info.Tpsa, 1)
        Lipinski = if ($info.PassesLipinski) { 'PASS (0 viol.)' } else { 'FAIL' }
    }
}

$compList = @(
    (Get-CompoundProps 'Aspirin'     $aspirinSmiles),
    (Get-CompoundProps 'Caffeine'    $caffeineSmiles),
    (Get-CompoundProps 'Ibuprofen'   $ibuprofenSmiles),
    (Get-CompoundProps 'Paracetamol' $paracetamolSmiles)
)

Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.DisplayAlerts = 0
$doc = $null

try {
    $wordAddIn = $null
    foreach ($ai in $word.COMAddIns) {
        if ($ai.ProgId -eq 'Smiles2Img.Word.AddIn') { $wordAddIn = $ai; break }
    }
    if (-not $wordAddIn) { throw 'Smiles2Img.Word.AddIn is not registered!' }
    if (-not $wordAddIn.Connect) { $wordAddIn.Connect = $true }
    $wordObj = $wordAddIn.Object

    $doc = $word.Documents.Add()
    $doc.PageSetup.TopMargin    = 36
    $doc.PageSetup.BottomMargin = 36
    $doc.PageSetup.LeftMargin   = 54
    $doc.PageSetup.RightMargin  = 54
    $wHwnd = [IntPtr]$word.ActiveWindow.Hwnd

    # =========================================================================
    # PAGE 1: Cover Banner, Overview & Complete Word [SMILES] Command Reference
    # =========================================================================
    Write-Host '[1/6] Building Page 1: Cover Banner & Ribbon Reference...' -ForegroundColor Yellow
    $pTop = $doc.Paragraphs.Item(1)
    $bannerTbl = $doc.Tables.Add($pTop.Range, 1, 1)
    $bannerTbl.AllowAutoFit = $false
    $bannerTbl.PreferredWidthType = 3
    $bannerTbl.PreferredWidth = 504
    $bannerTbl.Borders.Enable = 0
    $bc = $bannerTbl.Cell(1, 1)
    $bc.Shading.BackgroundPatternColor = 0x6E760F # Deep Teal #0F766E (BGR)
    $bc.TopPadding = 10; $bc.BottomPadding = 10; $bc.LeftPadding = 14; $bc.RightPadding = 14
    $bc.Range.Text = "func_SMILES2IMG v$addinVersion — Microsoft Word Interactive Playbook`rSelf-Contained Manuscript & Hands-On Workbook (Zero Web Docs Needed)"
    $bc.Range.Font.Name = 'Segoe UI'
    $bc.Range.Font.Color = 0xFFFFFF
    $bc.Range.Paragraphs.Item(1).Range.Font.Size = 16.5
    $bc.Range.Paragraphs.Item(1).Range.Font.Bold = 1
    $bc.Range.Paragraphs.Item(1).Format.SpaceAfter = 3
    $bc.Range.Paragraphs.Item(2).Range.Font.Size = 10
    $bc.Range.Paragraphs.Item(2).Format.SpaceAfter = 0

    Add-Para $doc '' 2 $false $false 0x262626 0 2

    Add-CalloutBox $doc 'How to Use This Playbook (100% Interactive Inside Word)' @(
        '• Every chemical structure, reaction scheme, and 3D protein figure in this document was inserted using the real func_SMILES2IMG Word Add-in with live embedded metadata (Alt Text + CustomXMLParts).',
        '• Double-click ANY 2D structure or reaction scheme to reopen EPAM Ketcher 2D in place — edit bonds/arrows and press Ctrl+Enter without breaking paragraph layout or table column widths.',
        '• Try the hands-on exercise boxes in Chapters 1-5 to master 0.1s Paste Structure, Molecule Info table export, Mol* 3D proteins, 1-click ACS B&W Batch Style, and lossless Vector SVG copy.',
        '• Local-First Privacy: 2D rendering, Ketcher 2D, reaction KET storage, SVG export, and all property calculations run 100% offline. Only optional PubChem name lookup and RCSB PDB fetch use the network.'
    ) 0x6E760F 0xF4FBF9

    Add-ChapterHeading $doc '0.' 'Microsoft Word [SMILES] Ribbon & 3-Level Right-Click Menu Reference' $false

    $ribbonHeaders = @('Ribbon Group / Menu', 'Command', 'Shortcut / Access', 'What It Does in Microsoft Word')
    $ribbonWidths  = @(108, 126, 95, 175)
    $ribbonRows = @(
        @('Insert Structure', 'Insert / Edit Structure', 'Ctrl+Shift+K / Double-Click', 'Opens EPAM Ketcher 2D to draw molecules or reaction schemes, or re-edit selected figure in place.'),
        @('Insert Structure', '3D Structure (Mol*)',     'Ribbon -> [SMILES]',          'Opens Mol* 3D viewer to load PDB IDs (1CRN, 4HHB, 6LU7) or local .pdb/.cif files as 2D PNG or 3D GLB.'),
        @('Insert Structure', 'Paste Structure',         'Ribbon or Right-Click',       'Converts selected text (SMILES, InChI, KET, or compound name) or clipboard into an inline structure in 0.1s.'),
        @('Structure Tools',  'Copy SMILES / InChI / KET','Ribbon or Right-Click',      'Extracts the exact SMILES, IUPAC InChI, or lossless KET JSON from any selected structure picture.'),
        @('Structure Tools',  'Copy SVG',                'Ribbon or Right-Click',       'Copies lossless vector SVG graphics directly to clipboard for Adobe Illustrator or Inkscape.'),
        @('Structure Tools',  'Molecule Info',           'Ribbon or Right-Click',       'Displays 17 descriptors + 2D preview; copies Summary, 17-col TSV, Multi-Line List, Markdown, or JSON.'),
        @('Structure Tools',  'Batch Style',             'Ribbon -> Batch Style',       '1-click document-wide restyle: Standard Color, Publication B&W (ACS), White Background, or Dark Mode.'),
        @('Right-Click Menu', 'Chemical Structure > ...','Right-Click Text / Picture',  '3-level cascading menu exposing Insert/Edit, Mol* 3D, Paste Structure, Copy SVG, and Copy Molecule Info.'),
        @('About',            'About func_SMILES2IMG',   'Ribbon -> [SMILES]',          'Shows installed version, built-in engine versions, open-source licenses, and Check for Updates.')
    )
    Add-StyledWordTable $doc $ribbonHeaders $ribbonRows $ribbonWidths 0x6E760F 8.5 2.5 | Out-Null

    # =========================================================================
    # PAGE 2: Chapter 1 — Inline Structures & 0.1-Second Text Conversion
    # =========================================================================
    Write-Host '[2/6] Building Page 2: Chapter 1 — Inline Structures & 0.1s Quick Convert...' -ForegroundColor Yellow
    Add-ChapterHeading $doc '1.' 'Inline Structures & 0.1-Second Text-to-Structure Conversion' $true
    Add-Para $doc 'Structures inserted in Word are embedded as aspect-ratio-locked InlineShapes so they flow cleanly with manuscript text and never drift when paragraphs shift. Below are two live structures embedded with genuine chemical metadata — double-click either one to reopen it in Ketcher 2D!' 10 $false $false 0x262626 2 8 | Out-Null

    $pEx1 = Add-Para $doc 'Example 1A — Acetylsalicylic Acid (Aspirin, SMILES embedded; double-click to edit):' 10 $true $false 0x1E293B 4 4
    $pImg1 = $doc.Paragraphs.Add()
    [void](Insert-LiveWordStructure $word $wordObj $pImg1.Range $aspirinSmiles 140)
    $pImg1.Range.InsertParagraphAfter()

    $pEx2 = Add-Para $doc 'Example 1B — Caffeine (Purine alkaloid, SMILES embedded; double-click to edit):' 10 $true $false 0x1E293B 8 4
    $pImg2 = $doc.Paragraphs.Add()
    [void](Insert-LiveWordStructure $word $wordObj $pImg2.Range $caffeineSmiles 140)
    $pImg2.Range.InsertParagraphAfter()

    Add-Para $doc '' 4 $false $false 0x262626 4 4 | Out-Null

    Add-CalloutBox $doc 'Hands-On Exercise 1 — Convert Plain Text into a Structure in 0.1 Seconds!' @(
        'Try it right now inside this box: highlight ANY of the three chemical identifiers below, then click [SMILES] -> [Paste Structure] (or Right-Click -> Chemical Structure -> Paste Structure):',
        '',
        '  1) SMILES string   :  CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O',
        '  2) Chemical Name   :  Paracetamol',
        '  3) IUPAC InChI     :  InChI=1S/C8H9NO2/c1-6(10)9-7-2-4-8(11)5-3-7/h2-5,11H,1H3,(H,9,10)',
        '',
        'Tip: If no text is selected, [Paste Structure] converts whatever SMILES, InChI, KET, or compound name is currently on your clipboard!'
    ) 0xBD5A18 0xF0F6FE

    # =========================================================================
    # PAGE 3: Chapter 2 — Real-World Use Case: Publication Compound Summary Table
    # =========================================================================
    Write-Host '[3/6] Building Page 3: Chapter 2 — Publication Compound Summary Table...' -ForegroundColor Yellow
    Add-ChapterHeading $doc '2.' 'Real-World Use Case: Journal Compound Summary & SAR Table' $true
    Add-Para $doc 'Medicinal chemistry manuscripts require crisp 2D structures aligned beside exact molecular formulas, weights, LogP/TPSA, and Lipinski Rule of 5 assessments. The table below uses Fixed Column Width (Right-Click Table -> AutoFit -> Fixed Column Width) so editing any structure in place never distorts adjacent text columns.' 10 $false $false 0x262626 2 8 | Out-Null

    $sarHeaders = @('Compound', '2D Structure (Double-Click)', 'Formula', 'MW (g/mol)', 'LogP / TPSA', 'Lipinski Ro5')
    $sarWidths  = @(78, 146, 62, 66, 82, 70)
    $sarRows = @()
    foreach ($c in $compList) {
        $sarRows += ,@($c.Name, '', $c.Formula, ("{0:F2}" -f $c.MW), ("{0:F2} / {1:F1}" -f $c.LogP, $c.TPSA), $c.Lipinski)
    }
    $sarTbl = Add-StyledWordTable $doc $sarHeaders $sarRows $sarWidths 0x1E3A8A
    for ($i = 0; $i -lt $compList.Count; $i++) {
        $imgCell = $sarTbl.Cell($i + 2, 2)
        $imgCell.Range.Text = ''
        $imgCell.Range.ParagraphFormat.Alignment = 1 # center
        [void](Insert-LiveWordStructure $word $wordObj $imgCell.Range $compList[$i].Smiles 115)
    }

    Add-CalloutBox $doc 'Hands-On Exercise 2 — Zero-Typo Property Extraction via [Molecule Info]' @(
        '1. Click ANY structure inside Table 2 above (e.g., Ibuprofen), then click [SMILES] -> [Molecule Info] (or Right-Click -> Chemical Structure -> Copy Molecule Info).',
        '2. Try these 3 time-saving copy formats:',
        '   • Copy All Summary (Text)          — pastes a clean 2-row Header + Value summary table.',
        '   • Copy as Excel Row (TSV Columns)  — pastes all 17 descriptors across tab-separated columns.',
        '   • Detailed Molecule Info Window... — inspect all 17 descriptors with a live 2D thumbnail, double-click any row to copy a single value, or export Markdown / JSON.'
    ) 0x6E760F 0xF4FBF9

    # =========================================================================
    # PAGE 4: Chapter 3 — Real-World Use Case: Multi-Step Reaction Schemes (KET)
    # =========================================================================
    Write-Host '[4/6] Building Page 4: Chapter 3 — Multi-Step Organic Reaction Schemes (KET)...' -ForegroundColor Yellow
    Add-ChapterHeading $doc '3.' 'Real-World Use Case: Multi-Step Organic Synthesis Schemes (KET)' $true
    Add-Para $doc 'Unlike plain SMILES, EPAM Ketcher KET JSON stores exact 2D coordinates, reaction arrows, + signs, catalysts, temperatures, and yields inside Word CustomXMLParts. When you share this .docx with a co-author who has func_SMILES2IMG installed, double-clicking the scheme restores every arrow and label losslessly!' 10 $false $false 0x262626 2 8 | Out-Null

    Add-Para $doc 'Scheme 1. Synthesis of Acetylsalicylic Acid (Aspirin) from Salicylic Acid & Acetic Anhydride (KET Embedded):' 10 $true $false 0x1E293B 4 4 | Out-Null
    $pRxn1 = $doc.Paragraphs.Add()
    $pRxn1.Format.Alignment = 1
    [void](Insert-LiveWordStructure $word $wordObj $pRxn1.Range ("KET:$ketAspirinRxn") 400)
    $pRxn1.Range.InsertParagraphAfter()

    Add-Para $doc 'Scheme 2. Synthesis of Acetaminophen (Paracetamol) from 4-Aminophenol & Acetic Anhydride (KET Embedded):' 10 $true $false 0x1E293B 10 4 | Out-Null
    $pRxn2 = $doc.Paragraphs.Add()
    $pRxn2.Format.Alignment = 1
    [void](Insert-LiveWordStructure $word $wordObj $pRxn2.Range ("KET:$ketParacetamolRxn") 400)
    $pRxn2.Range.InsertParagraphAfter()

    Add-CalloutBox $doc 'Hands-On Exercise 3 — Edit Reaction Conditions In-Place' @(
        '1. Double-click Scheme 1 above -> EPAM Ketcher 2D opens with Salicylic Acid, Acetic Anhydride, the + sign, reaction arrow, and Aspirin already positioned.',
        '2. Use the Ketcher Text tool (T) to add "H2SO4 (cat.), 85 C, 15 min (80%)" above the reaction arrow.',
        '3. Press Ctrl+Enter (Apply to Document). Word updates Scheme 1 cleanly in place with full KET metadata preserved!'
    ) 0x6E760F 0xF4FBF9

    # =========================================================================
    # PAGE 5: Chapter 4 — Real-World Use Case: 3D Protein Target Figure (Mol* 3D)
    # =========================================================================
    Write-Host '[5/6] Building Page 5: Chapter 4 — 3D Protein Target Figure (Mol* 3D)...' -ForegroundColor Yellow
    Add-ChapterHeading $doc '4.' 'Real-World Use Case: 3D Protein Target Figures (Mol* 3D)' $true
    Add-Para $doc 'Embed 3D macromolecular crystal structures straight from RCSB PDB / AlphaFold codes (1CRN, 4HHB, 6LU7, 1BNA) or offline local .pdb / .cif / .sdf files — without leaving Microsoft Word. Double-click Figure 1 below to reopen Mol* 3D with PDB 1CRN preloaded!' 10 $false $false 0x262626 2 8 | Out-Null

    $crnBytes = Get-1CrnPngBytes $wHwnd
    $tmp1Crn = Join-Path ([System.IO.Path]::GetTempPath()) ("word_playbook_1crn_" + [Guid]::NewGuid().ToString('N') + '.png')
    [System.IO.File]::WriteAllBytes($tmp1Crn, $crnBytes)
    try {
        $pMol = $doc.Paragraphs.Add()
        $pMol.Format.Alignment = 1
        $pic3d = $pMol.Range.InlineShapes.AddPicture($tmp1Crn, $false, $true)
        try { $pic3d.LockAspectRatio = -1 } catch {}
        $pic3d.Width = 260
        $meta = New-Object Smiles2Img.MolstarMetadata
        $meta.Id = [Guid]::NewGuid().ToString('N')
        $meta.Pdb = '1CRN'
        $meta.Mode = 'image'
        $meta.Embed = $true
        $meta.Background = 'transparent'
        $pic3d.AlternativeText = [Smiles2Img.MolstarSessionManager]::FormatAltText($meta)
        $pic3d.Title = [Smiles2Img.MolstarSessionManager]::FormatTitle('1CRN')
        $pMol.Range.InsertParagraphAfter()
    } finally {
        Remove-Item -LiteralPath $tmp1Crn -Force -ErrorAction SilentlyContinue
    }

    Add-Para $doc 'Figure 1. Crystal structure of Crambin (PDB: 1CRN) rendered with Mol* 3D (transparent background; double-click to reopen in Mol* 3D).' 9 $false $true 0x475569 2 10 'Segoe UI' 1 | Out-Null

    $molHeaders = @('Output Mode', 'Format', 'Best For', 'How to Use')
    $molWidths  = @(104, 84, 158, 158)
    $molRows = @(
        @('2D Picture (PNG)',  'High-DPI PNG',  'Print journals, PDF manuscripts, Word reports',       'Choose BG: Transparent (or White) in Mol* top bar -> click Apply to Document.'),
        @('3D Model (GLB)',    'Embedded .glb', 'Interactive 360-degree review inside Word 2019+/365', 'Select 3D Model (GLB) in Mol* -> drag center 3D handle in Word to rotate freely.'),
        @('Offline PDB / CIF', 'Local File',    'Firewalled labs blocking RCSB PDB / PDBe access',     'Click [Open File...] in Mol* toolbar to open local .pdb, .cif, .bcif, .sdf, or .mol2.')
    )
    Add-StyledWordTable $doc $molHeaders $molRows $molWidths 0x475569 | Out-Null

    # =========================================================================
    # PAGE 6: Chapter 5 — Pre-Submission ACS B&W Polish, Vector SVG & Troubleshooting
    # =========================================================================
    Write-Host '[6/6] Building Page 6: Chapter 5 — Pre-Submission Polish, SVG & Troubleshooting...' -ForegroundColor Yellow
    Add-ChapterHeading $doc '5.' 'Pre-Submission Polish: 1-Click ACS B&W, Vector SVG & Troubleshooting' $true

    Add-CalloutBox $doc 'Hands-On Exercise 4 — Convert This Entire Playbook to ACS Monochrome in 1 Click!' @(
        '1. Make sure NO single structure is selected (click empty paragraph text), then click [SMILES] -> [Batch Style] -> [Publication B&W].',
        '2. Scroll back through Chapters 1-3: every 2D molecule and reaction scheme across the entire manuscript is now crisp ACS / Nature black-and-white line art!',
        '3. Click [SMILES] -> [Batch Style] -> [Standard Color] anytime to restore IUPAC element colors.',
        '4. For Adobe Illustrator / Inkscape artwork: select ANY figure -> click [SMILES] -> [Copy SVG] (or Right-Click -> Chemical Structure -> Copy Vector SVG) -> press Ctrl+V in your vector editor.'
    ) 0x6E760F 0xF4FBF9

    Add-Para $doc 'Quick Troubleshooting Reference (Built-In Offline Guide)' 11 $true $false 0x1E293B 8 5 | Out-Null

    $tsHeaders = @('Issue / Question', 'Root Cause', 'Exact Solution')
    $tsWidths  = @(130, 168, 206)
    $tsRows = @(
        @('Sharing .docx with macOS / co-authors without add-in', 'Word structures are stored as standard embedded PNG/GLB media.', 'Co-authors view and print every figure at full resolution with zero add-in required. Chemical metadata stays safely preserved inside the file.'),
        @('Double-clicking a structure does not open Ketcher',    'Plain screenshot lacks Alt Text metadata, or Word COM add-in disabled.', 'Insert structures via [Insert / Edit Structure] or [Paste Structure]. Verify File -> Options -> Add-ins -> Manage: COM Add-ins -> func_SMILES2IMG is checked.'),
        @('Table columns shift after editing a structure',        'Word AutoFit to Contents dynamically resizes columns on replacement.', 'Right-click the table -> AutoFit -> Fixed Column Width so adjacent text columns remain locked.'),
        @('Mol* cannot fetch PDB IDs behind corporate firewall',  'HTTPS requests to RCSB PDB / PDBe are blocked by network proxy.', 'Download the .pdb or .cif file manually and load it 100% offline via [Open File...] in the Mol* toolbar.')
    )
    Add-StyledWordTable $doc $tsHeaders $tsRows $tsWidths 0x3B291E | Out-Null

    # Save DOCX & sync to docs/public/examples/
    Write-Host "`nSaving Word Playbook to: $targetFile" -ForegroundColor Cyan
    $doc.SaveAs2($targetFile, 16) # wdFormatXMLDocument (.docx)
    if ($PreviewPdfPath) {
        $pdfAbs = [System.IO.Path]::GetFullPath($PreviewPdfPath)
        New-Item -ItemType Directory -Path (Split-Path -Parent $pdfAbs) -Force | Out-Null
        $doc.ExportAsFixedFormat($pdfAbs, 17) # wdExportFormatPDF
        Write-Host "[OK] Exported PDF preview: $pdfAbs" -ForegroundColor Gray
    }
    $doc.Close($false)
    Copy-Item -LiteralPath $targetFile -Destination $docsPublicFile -Force
    Write-Host "[OK] Synced copy to: $docsPublicFile" -ForegroundColor Green

} finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    if ($word) {
        try { $word.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# OpenXML Integrity Verification
$sizeKB = [math]::Round((Get-Item -LiteralPath $targetFile).Length / 1KB, 1)
$zip = [System.IO.Compression.ZipFile]::OpenRead($targetFile)
try {
    $entries = @($zip.Entries | Select-Object -ExpandProperty FullName)
    $pngCount = @($entries | Where-Object { $_ -like 'word/media/*.png' }).Count
    $customXmlCount = @($entries | Where-Object { $_ -like 'customXml/item*.xml' }).Count
    $docXmlEntry = $zip.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($docXmlEntry.Open())
    $docXml = $sr.ReadToEnd(); $sr.Dispose()
    $smilesMetaCount = ([regex]::Matches($docXml, 'SMILES:')).Count
    $ketMetaCount    = ([regex]::Matches($docXml, 'KET:')).Count
    $molstarMetaCount = ([regex]::Matches($docXml, 'MOLSTAR:')).Count

    Write-Host "`n=== Word Playbook Verification ===" -ForegroundColor Cyan
    Write-Host "  File Size        : $sizeKB KB" -ForegroundColor Green
    Write-Host "  Embedded PNGs    : $pngCount unique PNG media (deduplicated across 9 embedded figures)" -ForegroundColor Green
    Write-Host "  CustomXML Parts  : $customXmlCount" -ForegroundColor Green
    Write-Host "  AltText Metadata : SMILES=$smilesMetaCount | KET=$ketMetaCount | MOLSTAR=$molstarMetaCount" -ForegroundColor Green
    if ($pngCount -lt 7) { throw "Expected >= 7 embedded PNGs in Word Playbook, found $pngCount" }
    if ($smilesMetaCount -lt 6 -or $molstarMetaCount -lt 1) { throw "Missing chemical AltText metadata inside Word Playbook!" }
} finally {
    $zip.Dispose()
}

Write-Host "`n[PASS] func_SMILES2IMG_Word_Playbook.docx generated and verified!" -ForegroundColor Green
