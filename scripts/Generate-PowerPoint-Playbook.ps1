<#
.SYNOPSIS
    Generates the official self-contained func_SMILES2IMG Microsoft PowerPoint Interactive Playbook (.pptx)
    using live Microsoft PowerPoint + Smiles2Img.PowerPoint.AddIn so every 2D structure, KET reaction scheme,
    and Mol* 3D protein carries genuine Shape.Tags & AlternativeText metadata and Speaker Notes.
#>
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\examples\func_SMILES2IMG_PowerPoint_Playbook.pptx'),
    [string]$ExportSlidesDir = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing, System.Windows.Forms, System.IO.Compression.FileSystem

$repoRoot = Split-Path -Parent $PSScriptRoot
$targetFile = [System.IO.Path]::GetFullPath($OutputPath)
$docsPublicFile = [System.IO.Path]::GetFullPath((Join-Path $repoRoot 'docs\public\examples\func_SMILES2IMG_PowerPoint_Playbook.pptx'))
New-Item -ItemType Directory -Path (Split-Path -Parent $targetFile) -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $docsPublicFile) -Force | Out-Null

# Load Smiles2Img.Core + NativeIndigo
$coreDll = (Resolve-Path (Join-Path $repoRoot 'dist\x64\PowerPoint\Smiles2Img.Core.dll')).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))

$addinVersion = '1.9.0'
try {
    $propsXml = [xml](Get-Content (Join-Path $repoRoot 'Directory.Build.props') -Raw)
    if ($propsXml.Project.PropertyGroup.Version) { $addinVersion = [string]$propsXml.Project.PropertyGroup.Version }
} catch {}

Write-Host "=== Generating func_SMILES2IMG v$addinVersion — 8-Slide Self-Contained PowerPoint Playbook ===" -ForegroundColor Cyan
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
        $form = New-Object Smiles2Img.MolstarForm('PowerPoint Slide', '1CRN', $null)
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

# BGR color palette (PowerPoint COM uses BGR 0xBBGGRR!)
$C_BG_SLATE   = 0x2A170F # #0F172A Deep Slate
$C_CARD_DARK  = 0x3B291E # #1E293B Card Slate
$C_CARD_LINE  = 0x554133 # #334155 Subtle Border
$C_TEAL       = 0xA6B814 # #14B8A6 Brand Teal
$C_TEAL_DARK  = 0x6E760F # #0F766E Dark Teal
$C_SKY        = 0xF8BD38 # #38BDF8 Sky Accent
$C_AMBER      = 0x24B4F5 # #F5B424 Warm Amber
$C_CORAL      = 0x476CED # #ED6C47 PowerPoint Coral
$C_WHITE      = 0xFFFFFF
$C_TEXT_MUTED = 0xD5C5B8 # Soft Slate Gray
$C_TEXT_DARK  = 0x2A170F

function New-DarkSlide {
    param([object]$pres, [string]$badge, [string]$title, [string]$subtitle, [string]$notes)
    $s = $pres.Slides.Add($pres.Slides.Count + 1, 12) # ppLayoutBlank
    $s.FollowMasterBackground = 0
    $s.Background.Fill.Solid()
    $s.Background.Fill.ForeColor.RGB = $C_BG_SLATE

    # Top accent bar
    $bar = $s.Shapes.AddShape(1, 0, 0, 960, 5)
    $bar.Fill.Solid(); $bar.Fill.ForeColor.RGB = $C_TEAL
    $bar.Line.Visible = 0

    if ($badge) {
        $b = $s.Shapes.AddTextbox(1, 36, 16, 400, 20)
        $b.TextFrame.TextRange.Text = $badge.ToUpperInvariant()
        $b.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $b.TextFrame.TextRange.Font.Size = 9.5
        $b.TextFrame.TextRange.Font.Bold = -1
        $b.TextFrame.TextRange.Font.Color.RGB = $C_TEAL
    }

    $t = $s.Shapes.AddTextbox(1, 36, 36, 888, 38)
    $t.TextFrame.TextRange.Text = $title
    $t.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $t.TextFrame.TextRange.Font.Size = 22
    $t.TextFrame.TextRange.Font.Bold = -1
    $t.TextFrame.TextRange.Font.Color.RGB = $C_WHITE

    if ($subtitle) {
        $sub = $s.Shapes.AddTextbox(1, 36, 73, 888, 24)
        $sub.TextFrame.TextRange.Text = $subtitle
        $sub.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $sub.TextFrame.TextRange.Font.Size = 11.5
        $sub.TextFrame.TextRange.Font.Color.RGB = $C_TEXT_MUTED
    }

    # Subtle footer
    $ft = $s.Shapes.AddTextbox(1, 36, 514, 888, 18)
    $ft.TextFrame.TextRange.Text = "func_SMILES2IMG v$addinVersion  |  Microsoft PowerPoint Interactive Playbook  |  Slide $($pres.Slides.Count) of 8"
    $ft.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $ft.TextFrame.TextRange.Font.Size = 8.5
    $ft.TextFrame.TextRange.Font.Color.RGB = 0x8B7464

    if ($notes) {
        $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes
    }
    return $s
}

function Add-Card {
    param([object]$slide, [double]$x, [double]$y, [double]$w, [double]$h, [int]$fillBgr = 0x3B291E, [int]$lineBgr = 0x554133)
    $card = $slide.Shapes.AddShape(5, $x, $y, $w, $h) # 5 = msoShapeRoundedRectangle
    $card.Adjustments.Item(1) = 0.04
    $card.Fill.Solid()
    $card.Fill.ForeColor.RGB = $fillBgr
    $card.Line.ForeColor.RGB = $lineBgr
    $card.Line.Weight = 1.0
    return $card
}

function Add-Text {
    param(
        [object]$slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$text,
        [double]$size = 11, [bool]$bold = $false, [int]$colorBgr = 0xFFFFFF,
        [string]$fontName = 'Segoe UI', [int]$align = 1 # 1=left, 2=center
    )
    $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.MarginLeft = 4; $tb.TextFrame.MarginRight = 4; $tb.TextFrame.MarginTop = 2; $tb.TextFrame.MarginBottom = 2
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $text
    $tr.Font.Name = $fontName
    $tr.Font.Size = $size
    $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.Font.Color.RGB = $colorBgr
    $tr.ParagraphFormat.Alignment = $align
    return $tb
}

function Insert-LivePptStructure {
    param(
        [object]$pptApp,
        [object]$pptAddInObj,
        [object]$slide,
        [string]$payload,
        [double]$boxX, [double]$boxY, [double]$boxW, [double]$boxH,
        [bool]$darkMode = $true
    )
    $slide.Select()
    try { $pptApp.ActiveWindow.Selection.Unselect() } catch {}
    [void][Smiles2Img.Editor.SafeClipboard]::SetText($payload)
    Start-Sleep -Milliseconds 120
    [void]$pptAddInObj.ExecuteRibbonAction('OnQuickInsert')
    Start-Sleep -Milliseconds 350
    $pic = $slide.Shapes.Item($slide.Shapes.Count)
    $pic.Select()
    if ($darkMode) {
        [void]$pptAddInObj.ExecuteRibbonAction('OnBatchStyleDark')
        Start-Sleep -Milliseconds 300
        $pic = $slide.Shapes.Item($slide.Shapes.Count)
    }
    $pic.LockAspectRatio = -1
    $pic.Width = $boxW
    if ($pic.Height -gt $boxH) { $pic.Height = $boxH }
    $pic.Left = $boxX + (($boxW - $pic.Width) / 2)
    $pic.Top  = $boxY + (($boxH - $pic.Height) / 2)
    try { $pptApp.ActiveWindow.Selection.Unselect() } catch {}
    return $pic
}

$aspirinSmiles     = 'CC(=O)Oc1ccccc1C(=O)O'
$caffeineSmiles    = 'Cn1cnc2c1c(=O)n(C)c(=O)n2C'
$ibuprofenSmiles   = 'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O'
$nicotineSmiles    = 'CN1CCC[C@H]1c2cccnc2'
$paracetamolSmiles = 'CC(=O)Nc1ccc(O)cc1'
$ketAspirinRxn     = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('O=C(O)c1ccccc1O.CC(=O)OC(=O)C>>CC(=O)Oc1ccccc1C(=O)O')

Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null

try {
    $pptAddIn = $null
    foreach ($ai in $ppt.COMAddIns) {
        if ($ai.ProgId -eq 'Smiles2Img.PowerPoint.AddIn') { $pptAddIn = $ai; break }
    }
    if (-not $pptAddIn) { throw 'Smiles2Img.PowerPoint.AddIn is not registered!' }
    if (-not $pptAddIn.Connect) { $pptAddIn.Connect = $true }
    $pptObj = $pptAddIn.Object

    $pres = $ppt.Presentations.Add(-1)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540
    $pHwnd = [IntPtr]$ppt.HWND

    # =========================================================================
    # SLIDE 1: Cover & Playbook Roadmap
    # =========================================================================
    Write-Host '[1/8] Building Slide 1: Cover & Playbook Roadmap...' -ForegroundColor Yellow
    $s1 = $pres.Slides.Add(1, 12)
    $s1.FollowMasterBackground = 0
    $s1.Background.Fill.Solid()
    $s1.Background.Fill.ForeColor.RGB = $C_BG_SLATE
    $bar1 = $s1.Shapes.AddShape(1, 0, 0, 960, 6)
    $bar1.Fill.Solid(); $bar1.Fill.ForeColor.RGB = $C_TEAL; $bar1.Line.Visible = 0

    Add-Text $s1 48 48 420 26 "func_SMILES2IMG v$addinVersion  ·  INTERACTIVE PLAYBOOK" 11 $true $C_TEAL | Out-Null
    Add-Text $s1 48 80 560 94 "Chemical Structures & 3D Proteins`rfor Microsoft PowerPoint" 32 $true $C_WHITE | Out-Null
    Add-Text $s1 48 184 540 60 "Self-contained hands-on slide deck — zero web documentation required. Every molecule, reaction, and 3D figure carries live metadata so you can double-click to edit, convert textboxes in 0.1s, and switch slide themes with 1-click Batch Style." 12.5 $false $C_TEXT_MUTED | Out-Null

    # Hero live Dark-Mode Caffeine preview on right side of Slide 1
    [void](Add-Card $s1 620 56 296 226 $C_CARD_DARK $C_TEAL_DARK)
    Add-Text $s1 636 68 264 22 'LIVE STRUCTURE (DOUBLE-CLICK TO EDIT!)' 9.5 $true $C_TEAL 'Segoe UI' 2 | Out-Null
    [void](Insert-LivePptStructure $ppt $pptObj $s1 $caffeineSmiles 648 96 240 162 $true)
    Add-Text $s1 636 256 264 20 'Caffeine (C8H10N4O2  ·  MW 194.19 g/mol)' 9.5 $false $C_TEXT_MUTED 'Segoe UI' 2 | Out-Null

    # 3 Roadmap Cards along bottom
    $roadCards = @(
        @{ Title = 'Slides 2–3 · Core Skills';      Body = 'Complete [SMILES] Ribbon & Right-Click Menu map + live 0.1s TextBox-to-Structure conversion lab.' },
        @{ Title = 'Slides 4–6 · Real-World Deck';  Body = '4-Card Dark Mode candidate grid, multi-step Ketcher KET synthesis slide, and Mol* 3D protein target (1CRN).' },
        @{ Title = 'Slides 7–8 · Q&A & Delivery';   Body = 'Live conference Q&A via Molecule Info, 1-click deck-wide Batch Style, A0 Poster Vector SVG & sharing.' }
    )
    for ($i = 0; $i -lt 3; $i++) {
        $cx = 48 + $i * 296
        [void](Add-Card $s1 $cx 316 276 156 $C_CARD_DARK $C_CARD_LINE)
        Add-Text $s1 ($cx + 16) 332 244 26 $roadCards[$i].Title 12.5 $true $C_SKY | Out-Null
        Add-Text $s1 ($cx + 16) 364 244 92 $roadCards[$i].Body 10.5 $false $C_TEXT_MUTED | Out-Null
    }
    Add-Text $s1 48 492 864 24 'Local-First Privacy: 2D rendering, Ketcher 2D, reaction KET, SVG & properties run 100% offline. Optional PubChem name & RCSB PDB fetch use the network.' 9.5 $false 0x9E8B7B | Out-Null
    $s1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Welcome to the func_SMILES2IMG PowerPoint Playbook! Start by double-clicking the Caffeine molecule on the right side of Slide 1 to open EPAM Ketcher 2D directly inside PowerPoint."

    # =========================================================================
    # SLIDE 2: Complete [SMILES] Ribbon & Right-Click Menu Reference
    # =========================================================================
    Write-Host '[2/8] Building Slide 2: Complete Ribbon & Menu Cheat Sheet...' -ForegroundColor Yellow
    $s2 = New-DarkSlide $pres 'Slide 2 · Quick Reference' 'PowerPoint [SMILES] Ribbon & 3-Level Context Menu' 'Everything you need on one slide — no web documentation required' `
        'Use Slide 2 as your built-in command cheat sheet. Every command on the top [SMILES] Ribbon tab is also accessible via Right-Click -> Chemical Structure >.'

    $tblShape = $s2.Shapes.AddTable(9, 4, 36, 110, 888, 384)
    $tbl = $tblShape.Table
    $tbl.Columns.Item(1).Width = 150
    $tbl.Columns.Item(2).Width = 205
    $tbl.Columns.Item(3).Width = 175
    $tbl.Columns.Item(4).Width = 358
    $s2Data = @(
        @('Group / Menu',      'Command',                  'How to Access',              'Action in Microsoft PowerPoint'),
        @('Insert Structure',  'Insert / Edit Structure',  'Ribbon / Double-Click',      'Open Ketcher 2D to sketch or re-edit selected molecule/reaction in place'),
        @('Insert Structure',  '3D Structure (Mol*)',      'Ribbon -> [SMILES]',         'Load PDB ID (1CRN, 4HHB, 6LU7) or local .pdb/.cif as 3D GLB or 2D PNG'),
        @('Insert Structure',  'Paste Structure',          'Ribbon / Right-Click',       'Convert selected TextBox (SMILES/InChI/Name) or clipboard in 0.1s'),
        @('Structure Tools',   'Copy SMILES / InChI / KET','Ribbon / Right-Click',       'Extract exact chemical identifier or lossless KET JSON from any slide shape'),
        @('Structure Tools',   'Copy SVG',                 'Ribbon / Right-Click',       'Copy lossless Vector SVG straight into Adobe Illustrator / A0 posters'),
                @('Structure Tools',   'Molecule Info',            'Ribbon / Right-Click',       'Inspect 12 descriptors live or copy Multi-Line List / TSV / Markdown / JSON'),
        @('Structure Tools',   'Batch Style',              'Ribbon -> Batch Style',      '1-click restyle: Dark Mode (White Ink), Color, ACS B&W, or White BG'),
        @('Right-Click Menu',  'Chemical Structure > ...', 'Right-Click Shape / Text',   '3-level cascading menu for editing, copying, and Molecule Info formats')
    )
    for ($r = 1; $r -le 9; $r++) {
        for ($c = 1; $c -le 4; $c++) {
            $cell = $tbl.Cell($r, $c)
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = if ($r -eq 1) { $C_TEAL_DARK } elseif ($r % 2 -eq 0) { $C_CARD_DARK } else { 0x332217 }
            $tr = $cell.Shape.TextFrame.TextRange
            $tr.Text = $s2Data[$r - 1][$c - 1]
            $tr.Font.Name = 'Segoe UI'
            $tr.Font.Size = if ($r -eq 1) { 10.5 } else { 9.8 }
            $tr.Font.Bold = if ($r -eq 1 -or $c -eq 2) { -1 } else { 0 }
            $tr.Font.Color.RGB = if ($r -eq 1) { $C_WHITE } elseif ($c -eq 2) { $C_SKY } else { 0xEFE7E0 }
        }
    }

    # =========================================================================
    # SLIDE 3: Hands-On Lab — 0.1-Second TextBox-to-Structure Conversion
    # =========================================================================
    Write-Host '[3/8] Building Slide 3: Hands-On 0.1s TextBox Conversion Lab...' -ForegroundColor Yellow
    $s3 = New-DarkSlide $pres 'Slide 3 · Hands-On Exercise' 'Convert Slide TextBoxes to Structures in 0.1 Seconds' 'Select any amber TextBox on the left -> click [SMILES] -> [Paste Structure] (then [Batch Style] -> [Dark Mode])' `
        'Exercise: Click one of the 3 plain-text boxes on the left so its border is selected, then click [Paste Structure] on the [SMILES] Ribbon tab. It instantly converts into a crisp, aspect-ratio-locked structure!'

    [void](Add-Card $s3 36 112 456 382 $C_CARD_DARK $C_CARD_LINE)
    Add-Text $s3 56 126 416 26 'TRY IT LIVE — SELECT ANY TEXTBOX BELOW & CLICK [PASTE STRUCTURE]' 10.5 $true $C_AMBER | Out-Null

    $labPrompts = @(
        @{ Label = '1. Plain SMILES String (Aspirin):';      Val = 'CC(=O)Oc1ccccc1C(=O)O'; Y = 164 },
        @{ Label = '2. Common Chemical Name (PubChem):';     Val = 'Paracetamol';           Y = 264 },
        @{ Label = '3. Chiral Alkaloid SMILES (Nicotine):';  Val = 'CN1CCC[C@H]1c2cccnc2';  Y = 364 }
    )
    foreach ($lp in $labPrompts) {
        Add-Text $s3 56 $lp.Y 416 22 $lp.Label 10.5 $true $C_TEXT_MUTED | Out-Null
        $tbTry = $s3.Shapes.AddTextbox(1, 56, ($lp.Y + 24), 416, 42)
        $tbTry.Fill.Solid(); $tbTry.Fill.ForeColor.RGB = 0x2A170F
        $tbTry.Line.ForeColor.RGB = $C_AMBER; $tbTry.Line.Weight = 1.25
        $tbTry.TextFrame.TextRange.Text = $lp.Val
        $tbTry.TextFrame.TextRange.Font.Name = 'Consolas'
        $tbTry.TextFrame.TextRange.Font.Size = 14
        $tbTry.TextFrame.TextRange.Font.Bold = -1
        $tbTry.TextFrame.TextRange.Font.Color.RGB = $C_AMBER
    }

    [void](Add-Card $s3 516 112 408 382 $C_CARD_DARK $C_TEAL_DARK)
    Add-Text $s3 536 126 368 26 'EXPECTED RESULT (LIVE DARK-MODE STRUCTURE)' 10.5 $true $C_TEAL | Out-Null
    [void](Insert-LivePptStructure $ppt $pptObj $s3 $aspirinSmiles 556 168 328 210 $true)
    Add-Text $s3 536 392 368 84 "• Transparent background — zero opaque white rectangle!`r• Aspect ratio locked automatically (never stretched).`r• Double-click the Aspirin structure above to open Ketcher 2D and modify any functional group in place." 10.5 $false $C_TEXT_MUTED | Out-Null

    # =========================================================================
    # SLIDE 4: Real-World Use Case — 4-Card Dark Mode Candidate Screening Grid
    # =========================================================================
    Write-Host '[4/8] Building Slide 4: 4-Card Lead Candidate Grid...' -ForegroundColor Yellow
    $s4 = New-DarkSlide $pres 'Slide 4 · Real-World Use Case' 'Lead Compound Candidates — Dark Slide Screening Grid' 'Transparent white-ink structures aligned cleanly across 4 comparison cards (zero white boxes)' `
        'All 4 structures on Slide 4 are live PowerPoint shapes inserted via func_SMILES2IMG and styled with [Batch Style] -> [Dark Mode]. Double-click any molecule to edit substituents in place without disturbing card alignment!'

    $cards = @(
        @{ Name = 'Aspirin';     Meta = 'C9H8O4  ·  MW 180.16';   Sub = 'LogP 1.31  ·  TPSA 63.6'; Smiles = $aspirinSmiles },
        @{ Name = 'Caffeine';    Meta = 'C8H10N4O2  ·  MW 194.19';Sub = 'LogP -1.03  ·  TPSA 61.8';Smiles = $caffeineSmiles },
        @{ Name = 'Nicotine';    Meta = 'C10H14N2  ·  MW 162.24'; Sub = 'LogP 1.85  ·  Fsp3 0.50';Smiles = $nicotineSmiles },
        @{ Name = 'Paracetamol'; Meta = 'C8H9NO2  ·  MW 151.17';  Sub = 'LogP 1.35  ·  TPSA 49.3'; Smiles = $paracetamolSmiles }
    )
    for ($i = 0; $i -lt 4; $i++) {
        $cx = 36 + $i * 226
        [void](Add-Card $s4 $cx 114 210 342 $C_CARD_DARK $C_CARD_LINE)
        Add-Text $s4 ($cx + 12) 126 186 28 $cards[$i].Name 15 $true $C_SKY | Out-Null
        Add-Text $s4 ($cx + 12) 154 186 22 $cards[$i].Meta 10 $false $C_WHITE | Out-Null
        [void](Insert-LivePptStructure $ppt $pptObj $s4 $cards[$i].Smiles ($cx + 16) 190 178 190 $true)
        Add-Text $s4 ($cx + 12) 416 186 26 $cards[$i].Sub 9.5 $false $C_TEAL 'Segoe UI' 2 | Out-Null
    }
    Add-Text $s4 36 468 888 26 'Pro Tip: Double-click any card molecule to edit in Ketcher 2D — replacement keeps exact card bounds and aspect ratio!' 10.5 $false $C_TEXT_MUTED | Out-Null

    # =========================================================================
    # SLIDE 5: Real-World Use Case — Multi-Step Reaction Pathway Slide (KET)
    # =========================================================================
    Write-Host '[5/8] Building Slide 5: Multi-Step Reaction Pathway Slide (KET)...' -ForegroundColor Yellow
    $s5 = New-DarkSlide $pres 'Slide 5 · Real-World Use Case' 'Synthetic Mechanism & Reaction Pathway Slide (KET)' 'Full reaction arrows, + signs, and coordinates preserved inside PowerPoint Shape.Tags' `
        'Double-click the Aspirin synthesis reaction scheme on Slide 5 to open EPAM Ketcher 2D with all reactants, arrows, and product coordinates restored. Press Ctrl+Enter to update cleanly in place.'

    [void](Add-Card $s5 56 116 848 312 $C_WHITE 0xCBD5E1)
    [void](Insert-LivePptStructure $ppt $pptObj $s5 ("KET:$ketAspirinRxn") 96 144 768 256 $false)

    [void](Add-Card $s5 56 444 848 52 $C_CARD_DARK $C_TEAL_DARK)
    Add-Text $s5 74 454 812 34 'Try It: Double-click the reaction scheme above to add catalyst / temperature conditions in Ketcher 2D -> press Ctrl+Enter! Or right-click -> Chemical Structure -> Copy Vector SVG for crisp A0 conference posters.' 10.5 $false $C_WHITE | Out-Null

    # =========================================================================
    # SLIDE 6: Real-World Use Case — 3D Macromolecular Target (Mol* 3D)
    # =========================================================================
    Write-Host '[6/8] Building Slide 6: 3D Protein Target (Mol* 3D)...' -ForegroundColor Yellow
    $s6 = New-DarkSlide $pres 'Slide 6 · Real-World Use Case' '3D Macromolecular Target (PDB: 1CRN Crambin)' 'Fetch PDB / AlphaFold structures or open offline .pdb / .cif files via [3D Structure (Mol*)]' `
        'Double-click the Crambin (1CRN) protein structure on the left to reopen Mol* 3D directly inside PowerPoint. Try switching between 3D Model (GLB) for 360-degree Morph transitions and 2D Picture (BG: Transparent) for ultra-light slide decks.'

    [void](Add-Card $s6 36 114 480 374 $C_CARD_DARK $C_TEAL_DARK)
    $crnBytes = Get-1CrnPngBytes $pHwnd
    $tmp1Crn = Join-Path ([System.IO.Path]::GetTempPath()) ("ppt_playbook_1crn_" + [Guid]::NewGuid().ToString('N') + '.png')
    [System.IO.File]::WriteAllBytes($tmp1Crn, $crnBytes)
    try {
        $shp3d = $s6.Shapes.AddPicture($tmp1Crn, 0, -1, 76, 134, 400, 300)
        $shp3d.LockAspectRatio = -1
        $meta = New-Object Smiles2Img.MolstarMetadata
        $meta.Id = [Guid]::NewGuid().ToString('N')
        $meta.Pdb = '1CRN'
        $meta.Mode = 'image'
        $meta.Embed = $true
        $meta.Background = 'transparent'
        $shp3d.AlternativeText = [Smiles2Img.MolstarSessionManager]::FormatAltText($meta)
        $shp3d.Title = [Smiles2Img.MolstarSessionManager]::FormatTitle('1CRN')
        $shp3d.Tags.Add('MOLSTAR', $meta.Id)
        $shp3d.Tags.Add('PDB', '1CRN')
    } finally {
        Remove-Item -LiteralPath $tmp1Crn -Force -ErrorAction SilentlyContinue
    }
    Add-Text $s6 52 454 448 24 'PDB: 1CRN (Crambin)  ·  Double-click to reopen in Mol* 3D!' 10 $true $C_TEAL 'Segoe UI' 2 | Out-Null

    [void](Add-Card $s6 540 114 384 374 $C_CARD_DARK $C_CARD_LINE)
    Add-Text $s6 560 130 344 26 '3-STEP 3D WORKFLOW IN POWERPOINT' 11 $true $C_SKY | Out-Null
    Add-Text $s6 560 162 344 134 "1. Click [SMILES] -> [3D Structure (Mol*)].`r`r2. Enter a 4-letter PDB code (1CRN, 4HHB, 6LU7, 1BNA) and click [Fetch] — or click [Open File...] for offline local .pdb / .cif / .sdf files.`r`r3. Pick an output mode & click [Apply to Slide]:" 10.5 $false $C_WHITE | Out-Null
    [void](Add-Card $s6 560 306 344 74 0x2A170F $C_TEAL_DARK)
    Add-Text $s6 572 314 320 60 "3D Model (GLB)  [PowerPoint 2019+/365]`rRotate 360° live during talks + combine with PowerPoint's Morph slide transition!" 9.8 $false $C_TEXT_MUTED | Out-Null
    [void](Add-Card $s6 560 392 344 74 0x2A170F $C_CARD_LINE)
    Add-Text $s6 572 400 320 60 "2D Picture (BG: Transparent)`rCrisp high-DPI transparent PNG for lightweight decks, PDFs & older PowerPoint versions." 9.8 $false $C_TEXT_MUTED | Out-Null

    # =========================================================================
    # SLIDE 7: Live Conference Q&A (Molecule Info) & 1-Click Deck Batch Style
    # =========================================================================
    Write-Host '[7/8] Building Slide 7: Live Q&A Molecule Info & Batch Style Lab...' -ForegroundColor Yellow
    $s7 = New-DarkSlide $pres 'Slide 7 · Hands-On Lab' 'Live Conference Q&A (Molecule Info) & 1-Click Batch Style' 'Inspect exact descriptors mid-talk or switch your entire presentation theme in a single click' `
        'Exercise 1: Select the Ibuprofen structure on the left and click [Molecule Info] to inspect MW, Exact Mass, LogP, TPSA, and Fsp3 live. Exercise 2: Unselect all shapes and click [Batch Style] -> [White Background] or [Publication B&W], then switch back to [Dark Mode]!'

    [void](Add-Card $s7 36 114 448 374 $C_CARD_DARK $C_CARD_LINE)
    Add-Text $s7 56 128 408 26 '1. LIVE Q&A — SELECT IBUPROFEN -> CLICK [MOLECULE INFO]' 10.5 $true $C_SKY | Out-Null
    [void](Insert-LivePptStructure $ppt $pptObj $s7 $ibuprofenSmiles 76 160 368 176 $true)
    [void](Add-Card $s7 56 348 408 122 0x2A170F $C_CARD_LINE)
    Add-Text $s7 68 356 384 108 "Sample Output of [Copy as Multi-Line List]:`r• Formula: C13H18O2   |   MW: 206.29 g/mol`r• Exact Mass: 206.1307 Da   |   LogP: 3.21   |   TPSA: 37.3 A²`r• Heavy Atoms: 15   |   Rings: 1 (1 Aromatic)`r• Fraction Csp3 (Fsp3): 0.46" 9.8 $false $C_TEXT_MUTED 'Consolas' | Out-Null

    [void](Add-Card $s7 508 114 416 374 $C_CARD_DARK $C_TEAL_DARK)
    Add-Text $s7 528 128 376 26 '2. DECK-WIDE [BATCH STYLE] — TRY IT RIGHT NOW!' 10.5 $true $C_TEAL | Out-Null
    Add-Text $s7 528 160 376 62 'Click empty background (no shape selected), then click [SMILES] -> [Batch Style] to instantly convert EVERY 2D structure across your slide deck:' 10.5 $false $C_WHITE | Out-Null

    $presets = @(
        @{ Name = 'Dark Mode (White Ink)'; Desc = 'Transparent canvas + crisp white carbon bonds — ideal for dark decks.' },
        @{ Name = 'Standard Color';        Desc = 'Transparent canvas + IUPAC CPK element colors for light templates.' },
        @{ Name = 'Publication B&W';       Desc = 'Pure black ACS/Nature monochrome line art for handouts & papers.' },
        @{ Name = 'White Background';      Desc = 'Opaque white box behind color structures for busy photo/gradient slides.' }
    )
    for ($p = 0; $p -lt 4; $p++) {
        $py = 232 + $p * 60
        [void](Add-Card $s7 528 $py 376 50 0x2A170F $C_CARD_LINE)
        Add-Text $s7 540 ($py + 5) 352 20 $presets[$p].Name 10.5 $true $C_AMBER | Out-Null
        Add-Text $s7 540 ($py + 25) 352 20 $presets[$p].Desc 9.2 $false $C_TEXT_MUTED | Out-Null
    }

    # =========================================================================
    # SLIDE 8: Sharing, A0 Poster Vector Export & Built-In Troubleshooting
    # =========================================================================
    Write-Host '[8/8] Building Slide 8: Sharing, Poster SVG & Troubleshooting...' -ForegroundColor Yellow
    $s8 = New-DarkSlide $pres 'Slide 8 · Delivery & Support' 'Sharing Decks, A0 Poster Vector SVG & Quick Troubleshooting' 'Built-in offline checklist before sending slides to conference AV laptops or co-authors' `
        'All embedded PNG structures and 3D GLB models are standard Office media — they display smoothly on macOS, iPad, and conference PCs even without func_SMILES2IMG installed!'

    $cols8 = @(
        @{
            Title = '1. Sharing to Any Conference PC / Mac'
            Color = $C_TEAL
            Body  = "• Every 2D figure and 3D .glb model is saved natively inside the .pptx file.`r`r• Co-authors on macOS, PowerPoint Web, or venue AV laptops view and present with ZERO add-in installed.`r`r• Chemical metadata (Shape.Tags + Alt Text) stays preserved so you can still double-click to edit later on Windows."
        },
        @{
            Title = '2. Large A0 Conference Poster Export'
            Color = $C_SKY
            Body  = "• Never scale low-res screenshots onto an A0/A1 poster.`r`r• Right-click any slide molecule -> Chemical Structure -> Copy Vector SVG (or Ribbon [Copy SVG]).`r`r• Press Ctrl+V straight into Adobe Illustrator, Inkscape, or CorelDRAW for infinite-DPI vector artwork."
        },
        @{
            Title = '3. Built-In Quick Troubleshooting'
            Color = $C_CORAL
            Body  = "• White box on dark slide?`r  Run [Batch Style] -> [Dark Mode], or pick [Background: Transparent] in Ketcher / Mol*.`r`r• 3D GLB not supported?`r  Needs PowerPoint 2019+/365; use Mol* [2D Picture] mode on older versions.`r`r• Missing [SMILES] tab?`r  File -> Options -> Add-ins -> Manage: COM Add-ins -> check func_SMILES2IMG."
        }
    )
    for ($i = 0; $i -lt 3; $i++) {
        $cx = 36 + $i * 304
        [void](Add-Card $s8 $cx 116 280 370 $C_CARD_DARK $C_CARD_LINE)
        Add-Text $s8 ($cx + 16) 134 248 40 $cols8[$i].Title 13 $true $cols8[$i].Color | Out-Null
        Add-Text $s8 ($cx + 16) 184 248 286 $cols8[$i].Body 10.5 $false 0xEFE7E0 | Out-Null
    }

    # Select Slide 1 and Save
    $s1.Select()
    Write-Host "`nSaving PowerPoint Playbook to: $targetFile" -ForegroundColor Cyan
    $pres.SaveAs($targetFile, 24) # 24 = ppSaveAsOpenXMLPresentation (.pptx)

    if ($ExportSlidesDir) {
        New-Item -ItemType Directory -Path $ExportSlidesDir -Force | Out-Null
        for ($si = 1; $si -le $pres.Slides.Count; $si++) {
            $outPng = Join-Path $ExportSlidesDir ("slide_{0:D2}.png" -f $si)
            $pres.Slides.Item($si).Export($outPng, 'PNG', 1600, 900)
        }
        Write-Host "[OK] Exported $($pres.Slides.Count) slide previews to: $ExportSlidesDir" -ForegroundColor Gray
    }

    $pres.Saved = -1
    $pres.Close()
    Copy-Item -LiteralPath $targetFile -Destination $docsPublicFile -Force
    Write-Host "[OK] Synced copy to: $docsPublicFile" -ForegroundColor Green

} finally {
    if ($pres) { try { $pres.Saved = -1; $pres.Close() } catch {} }
    if ($ppt) {
        try { $ppt.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# OpenXML Integrity Verification
$sizeKB = [math]::Round((Get-Item -LiteralPath $targetFile).Length / 1KB, 1)
$zip = [System.IO.Compression.ZipFile]::OpenRead($targetFile)
try {
    $entries = @($zip.Entries | Select-Object -ExpandProperty FullName)
    $slideCount = @($entries | Where-Object { $_ -like 'ppt/slides/slide*.xml' }).Count
    $pngCount   = @($entries | Where-Object { $_ -like 'ppt/media/*.png' }).Count
    $tagCount   = @($entries | Where-Object { $_ -like 'ppt/tags/tag*.xml' }).Count
    $notesCount = @($entries | Where-Object { $_ -like 'ppt/notesSlides/notesSlide*.xml' }).Count

    Write-Host "`n=== PowerPoint Playbook Verification ===" -ForegroundColor Cyan
    Write-Host "  File Size     : $sizeKB KB" -ForegroundColor Green
    Write-Host "  Slides        : $slideCount (expected 8)" -ForegroundColor Green
    Write-Host "  Embedded PNGs : $pngCount unique PNG media (deduplicated across 9 shapes)" -ForegroundColor Green
    Write-Host "  Shape.Tags    : $tagCount metadata tag files" -ForegroundColor Green
    Write-Host "  Speaker Notes : $notesCount slides" -ForegroundColor Green
    if ($slideCount -ne 8) { throw "Expected 8 slides, found $slideCount" }
    if ($pngCount -lt 7)   { throw "Expected >= 7 embedded PNGs, found $pngCount" }
    if ($tagCount -lt 6)   { throw "Expected >= 6 Shape.Tags metadata parts, found $tagCount" }
} finally {
    $zip.Dispose()
}

Write-Host "`n[PASS] func_SMILES2IMG_PowerPoint_Playbook.pptx generated and verified!" -ForegroundColor Green
