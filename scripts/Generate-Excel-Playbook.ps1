<#
.SYNOPSIS
    Generates the official self-contained func_SMILES2IMG Excel Interactive Playbook
    (10 tabs: Quick Start -> Core Rendering -> Properties/Lipinski -> SAR/SMARTS ->
    Conversions -> Styles -> Reactions/KET -> Real-World Lab Workflows -> Legacy .FLOAT ->
    Complete 40-Function, Ribbon & SDF Reference).
#>
param(
    [string]$AddInPath = (Join-Path $PSScriptRoot '..\dist\x64\Smiles2Img-AddIn64-packed.xll'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\examples\func_SMILES2IMG_Excel_Playbook.xlsx')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# C# Excel COM helper to set 2D object arrays and formulas cleanly across PowerShell dynamic binder
Add-Type -ReferencedAssemblies 'Microsoft.CSharp.dll' -TypeDefinition @'
using System;
public static class ExcelPlaybookHelper {
    public static void SetRangeValue2(object range, object[,] data) {
        dynamic r = range;
        r.Value2 = data;
    }
    public static void SetRangeFormula(object range, object[,] formulas) {
        dynamic r = range;
        try { r.Formula2 = formulas; }
        catch { r.Formula = formulas; }
    }
}
'@

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path -LiteralPath $AddInPath)) {
    $fallbackXll = Join-Path $env:APPDATA 'Microsoft\AddIns\Smiles2Img\Smiles2Img-AddIn64.xll'
    if (Test-Path -LiteralPath $fallbackXll) { $AddInPath = $fallbackXll }
}
$resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath).Path
$targetFile = [System.IO.Path]::GetFullPath($OutputPath)
$docsPublicFile = [System.IO.Path]::GetFullPath((Join-Path $repoRoot 'docs\public\examples\func_SMILES2IMG_Excel_Playbook.xlsx'))

New-Item -ItemType Directory -Path (Split-Path -Parent $targetFile) -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $docsPublicFile) -Force | Out-Null

# Load Smiles2Img.Core to generate authentic Indigo KET JSON for reaction & canvas demos
$coreDll = (Resolve-Path (Join-Path $repoRoot 'dist\x64\Word\Smiles2Img.Core.dll')).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))

$addinVersion = '1.8.0'
try {
    $propsXml = [xml](Get-Content (Join-Path $repoRoot 'Directory.Build.props') -Raw)
    if ($propsXml.Project.PropertyGroup.Version) { $addinVersion = [string]$propsXml.Project.PropertyGroup.Version }
} catch {}

Write-Host "=== Generating func_SMILES2IMG v$addinVersion — 10-Tab Self-Contained Excel Playbook ===" -ForegroundColor Cyan
Write-Host "Add-in : $resolvedAddIn" -ForegroundColor Gray
Write-Host "Output : $targetFile" -ForegroundColor Gray

if (Test-Path -LiteralPath $targetFile) { Remove-Item -LiteralPath $targetFile -Force }

function Invoke-Excel([scriptblock]$operation, [int]$timeoutSec = 45) {
    $deadline = [DateTime]::UtcNow.AddSeconds($timeoutSec)
    do {
        try { return (& $operation) }
        catch {
            $hex = '{0:X8}' -f $_.Exception.HResult
            if (($hex -notmatch '80010001|8001010A|800AC472' -and
                 $_.Exception.ToString() -notmatch '80010001|8001010A|800AC472|rejected') -or
                [DateTime]::UtcNow -ge $deadline) { throw }
            Start-Sleep -Milliseconds 350
        }
    } while ($true)
}

function Get-ColLetter([int]$colIdx1Based) {
    $n = $colIdx1Based
    $res = ''
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $res = [char](65 + $rem) + $res
        $n = [math]::Floor(($n - 1) / 26)
    }
    return $res
}

function Set-BannerAndCallout {
    param(
        [object]$sheet,
        [int]$numCols,
        [string]$title,
        [string]$subtitle,
        [int]$titleBgBgr,
        [int]$subBgBgr = 0xF5F3EF
    )
    $lastCol = Get-ColLetter $numCols
    Invoke-Excel {
        $r1 = $sheet.Range("A1:${lastCol}1")
        $r1.Merge()
        $sheet.Range('A1').Value2 = $title
        $r1.Font.Name = 'Segoe UI'
        $r1.Font.Size = 14
        $r1.Font.Bold = $true
        $r1.Font.Color = 0xFFFFFF
        $r1.Interior.Color = $titleBgBgr
        $r1.RowHeight = 34
        $r1.VerticalAlignment = -4108
        $r1.HorizontalAlignment = -4131
        $r1.IndentLevel = 1

        $r2 = $sheet.Range("A2:${lastCol}2")
        $r2.Merge()
        $sheet.Range('A2').Value2 = $subtitle
        $r2.Font.Name = 'Segoe UI'
        $r2.Font.Size = 9.5
        $r2.Font.Color = 0x333333
        $r2.Interior.Color = $subBgBgr
        $r2.RowHeight = 28
        $r2.WrapText = $true
        $r2.VerticalAlignment = -4108
        $r2.HorizontalAlignment = -4131
        $r2.IndentLevel = 1
    }
}

function Add-StyledTable {
    param(
        [object]$sheet,
        [int]$headerRow,
        [string[]]$headers,
        [array]$rows,
        [hashtable]$formulas = @{},
        [int]$headerBgBgr,
        [int[]]$colWidths = @(),
        [int]$rowHeight = 76,
        [int[]]$leftAlignCols = @(),
        [int[]]$monoCols = @()
    )
    $numCols = $headers.Length
    $numRows = $rows.Length
    $lastCol = Get-ColLetter $numCols
    $startRow = $headerRow + 1
    $endRow = $startRow + $numRows - 1

    Invoke-Excel {
        for ($c = 0; $c -lt $numCols; $c++) {
            $sheet.Cells.Item($headerRow, $c + 1).Value2 = $headers[$c]
        }
        $hr = $sheet.Range("A${headerRow}:${lastCol}${headerRow}")
        $hr.Font.Name = 'Segoe UI'
        $hr.Font.Bold = $true
        $hr.Font.Size = 10
        $hr.Font.Color = 0xFFFFFF
        $hr.Interior.Color = $headerBgBgr
        $hr.RowHeight = 26
        $hr.WrapText = $true
        $hr.VerticalAlignment = -4108
        $hr.HorizontalAlignment = -4108
    }

    if ($numRows -gt 0) {
        $dataArr = New-Object 'object[,]' $numRows, $numCols
        for ($r = 0; $r -lt $numRows; $r++) {
            for ($c = 0; $c -lt $numCols; $c++) {
                $val = $rows[$r][$c]
                if ($val -is [string] -and $val.StartsWith('=')) { $val = "'" + $val }
                $dataArr[$r, $c] = $val
            }
        }

        Invoke-Excel {
            $dr = $sheet.Range("A${startRow}:${lastCol}${endRow}")
            [ExcelPlaybookHelper]::SetRangeValue2($dr, $dataArr)
            $dr.Font.Name = 'Segoe UI'
            $dr.Font.Size = 9.5
            $dr.RowHeight = $rowHeight
            $dr.WrapText = $true
            $dr.Borders.LineStyle = 1
            $dr.Borders.Color = 0xDCDCDC
            $dr.VerticalAlignment = -4108
            $dr.HorizontalAlignment = -4108

            # Subtle zebra striping for readability
            for ($r = 0; $r -lt $numRows; $r++) {
                if ($r % 2 -eq 1) {
                    $sheet.Range("A$($startRow + $r):${lastCol}$($startRow + $r)").Interior.Color = 0xFAF9F7
                }
            }

            foreach ($lc in $leftAlignCols) {
                $cl = Get-ColLetter $lc
                $sheet.Range("${cl}${startRow}:${cl}${endRow}").HorizontalAlignment = -4131
            }
            foreach ($mc in $monoCols) {
                $cl = Get-ColLetter $mc
                $mr = $sheet.Range("${cl}${startRow}:${cl}${endRow}")
                $mr.Font.Name = 'Consolas'
                $mr.Font.Size = 9.5
            }
        }
    }

    if ($colWidths.Length -eq $numCols) {
        Invoke-Excel {
            for ($c = 0; $c -lt $numCols; $c++) {
                $sheet.Columns.Item($c + 1).ColumnWidth = $colWidths[$c]
            }
        }
    }

    foreach ($colKey in $formulas.Keys) {
        $colFormulas = $formulas[$colKey]
        $fArr = New-Object 'object[,]' $colFormulas.Length, 1
        for ($r = 0; $r -lt $colFormulas.Length; $r++) {
            $fArr[$r, 0] = $colFormulas[$r]
        }
        Invoke-Excel {
            $fr = $sheet.Range("${colKey}${startRow}:${colKey}$($startRow + $colFormulas.Length - 1)")
            [ExcelPlaybookHelper]::SetRangeFormula($fr, $fArr)
        }
    }
}

function Add-SectionHeader {
    param([object]$sheet, [int]$row, [int]$numCols, [string]$text, [int]$bgBgr = 0xEAECEF, [int]$fgBgr = 0x1E293B)
    $lastCol = Get-ColLetter $numCols
    Invoke-Excel {
        $rng = $sheet.Range("A${row}:${lastCol}${row}")
        $rng.Merge()
        $sheet.Range("A${row}").Value2 = $text
        $rng.Font.Bold = $true
        $rng.Font.Size = 11
        $rng.Font.Color = $fgBgr
        $rng.Interior.Color = $bgBgr
        $rng.RowHeight = 25
        $rng.VerticalAlignment = -4108
        $rng.HorizontalAlignment = -4131
        $rng.IndentLevel = 1
    }
}

# Pre-build real KET strings via EPAM Indigo NativeIndigo
$ketAspirinRxn     = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('O=C(O)c1ccccc1O.CC(=O)OC(=O)C>>CC(=O)Oc1ccccc1C(=O)O')
$ketParacetamolRxn = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('Nc1ccc(O)cc1.CC(=O)OC(=O)C>>CC(=O)Nc1ccc(O)cc1')
$ketAspirinMol     = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('CC(=O)Oc1ccccc1C(=O)O')
$ketCaffeineMol    = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('Cn1cnc2c1c(=O)n(C)c(=O)n2C')
$ketIbuprofenMol   = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('CC(C)Cc1ccc(cc1)C(C)C(=O)O')

$excel = $null
$workbook = $null

try {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 400

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Excel Version: $($excel.Version)" -ForegroundColor Gray
    if (-not $excel.RegisterXLL($resolvedAddIn)) {
        throw "Failed to register XLL: $resolvedAddIn"
    }
    Write-Host "[OK] Registered XLL add-in" -ForegroundColor Green

    $workbook = $excel.Workbooks.Add()
    $excel.Calculation = -4135 # xlCalculationManual while populating all 10 tabs cleanly
    while ($workbook.Worksheets.Count -lt 10) {
        $null = $workbook.Worksheets.Add($workbook.Worksheets.Item($workbook.Worksheets.Count))
    }

    $tabSpecs = @(
        @{ Name = '0. Start Here';             Color = 0x6E760F }, # Teal #0F766E (BGR)
        @{ Name = '1. Structure Rendering';    Color = 0xB4661B }, # Blue
        @{ Name = '2. Properties & Lipinski';  Color = 0x417C10 }, # Excel Green
        @{ Name = '3. SAR & Similarity';       Color = 0x993366 }, # Purple
        @{ Name = '4. Notation Conversions';   Color = 0x8F5E00 }, # Deep Cyan
        @{ Name = '5. Styles & Transforms';    Color = 0x0066CC }, # Amber/Orange
        @{ Name = '6. Reactions & KET';        Color = 0x52731E }, # Olive Teal
        @{ Name = '7. Real-World Workflows';   Color = 0x2647D2 }, # Coral Red
        @{ Name = '8. Legacy Excel (.FLOAT)';  Color = 0x336633 }, # Forest
        @{ Name = '9. Ribbon, SDF & Ref';      Color = 0x3B291E }  # Slate Navy
    )
    $sheets = @()
    for ($i = 0; $i -lt 10; $i++) {
        $ws = $workbook.Worksheets.Item($i + 1)
        $ws.Name = $tabSpecs[$i].Name
        $ws.Tab.Color = $tabSpecs[$i].Color
        $sheets += $ws
    }

    # =========================================================================
    # TAB 0: Start Here — Self-Contained Quickstart, Roadmap & 3-Min Sandbox
    # =========================================================================
    Write-Host '[1/10] Building Tab 0: Start Here...' -ForegroundColor Yellow
    $s0 = $sheets[0]
    Set-BannerAndCallout $s0 7 `
        "func_SMILES2IMG v$addinVersion — Self-Contained Microsoft Excel Playbook" `
        "No web documentation needed! Explore all 10 tabs from basics to real-world lab workflows. SMILES/InChI/KET rendering, all 32 property/SAR formulas, and Ketcher 2D run locally offline; optional PubChem Name/CAS lookup and Mol* PDB fetch use the network once and cache locally." `
        0x6E760F

    Add-SectionHeader $s0 4 7 'A. 3-Minute Interactive Sandbox — Try Editing Any Input (Col C) or Double-Click Any Structure Cell (Col E)!'
    $t0Headers = @('#', 'Capability', 'Editable Input', 'Formula Syntax (Copyable)', 'Live Result (In-Cell)', 'How to Try It Right Now', 'What You Learn')
    $t0Widths  = @(5, 22, 26, 36, 26, 30, 48)
    $t0Rows = @(
        @(1, 'Name -> 2D Structure',   'Aspirin',                              '=NAME2IMG(C6)',                  '', 'Change "Aspirin" to "Caffeine"',        'Resolves chemical names via PubChem (cached) into crisp in-cell 2D pictures'),
        @(2, 'CAS -> 2D Structure',    '58-08-2',                              '=NAME2IMG(C7)',                  '', 'Try "50-78-2" (Aspirin)',               'Direct CAS registry number lookup without knowing SMILES'),
        @(3, 'SMILES -> 2D Structure', 'CC(=O)Oc1ccccc1C(=O)O',                '=SMILES2IMG(C8)',                '', 'Double-click Col E to open Ketcher',    'Offline local Indigo rendering; scales cleanly when resizing rows/columns'),
        @(4, 'InChI -> 2D Structure',  'InChI=1S/C8H9NO2/c1-6(10)9-7-2-4-8(11)5-3-7/h2-5,11H,1H3,(H,9,10)', '=INCHI2IMG(C9)', '', 'Edit InChI or double-click cell', 'Direct rendering from IUPAC Standard InChI identifiers'),
        @(5, 'Molecular Weight (MW)',  'CC(=O)Oc1ccccc1C(=O)O',                '=SMILES2MW(C10)',                '', 'Use in math: =SMILES2MW(C10)*10',       'Returns a native Excel Double (180.16) for stoichiometry & mass formulas'),
        @(6, 'Chemical Formula',       'CC(=O)Oc1ccccc1C(=O)O',                '=SMILES2FORMULA(C11)',           '', 'Combine with text: & operator',         'Instant offline Hill formula (C9H8O4) with zero server calls'),
        @(7, 'Lipinski Rule of 5',     'CC(=O)Oc1ccccc1C(=O)O',                '=SMILES2LIPINSKI(C12)',          '', 'Filter with =COUNTIF(..., "PASS")',     'Screens oral drug-likeness (MW<=500, LogP<=5, HBD<=5, HBA<=10)'),
        @(8, 'Substructure Match',     'CC(=O)Oc1ccccc1C(=O)O',                '=SMILES.MATCH(C13, "c1ccccc1")', '', 'Swap "c1ccccc1" -> "C(=O)O"',           'Returns TRUE/FALSE for SMARTS/SMILES pharmacophore filtering'),
        @(9, 'ACS Publication B&W',    'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O',      '=SMILES2IMG(C14, , "bw stereo")','', 'Or click Ribbon -> Batch Style',        'Combine tokens: pure black ACS journal line art + chiral R/S labels'),
        @(10,'Dark Mode (White Ink)',  'Cn1cnc2c1c(=O)n(C)c(=O)n2C',           '=SMILES2IMG(C15, "#1E293B", "white")','', 'Right-click -> Copy Vector SVG',    'High-contrast white carbon bonds over dark cells or conference slides')
    )
    $t0Formulas = @{
        'E' = @(
            '=NAME2IMG(C6)',
            '=NAME2IMG(C7)',
            '=SMILES2IMG(C8)',
            '=INCHI2IMG(C9)',
            '=SMILES2MW(C10)',
            '=SMILES2FORMULA(C11)',
            '=SMILES2LIPINSKI(C12)',
            '=SMILES.MATCH(C13, "c1ccccc1")',
            '=SMILES2IMG(C14, , "bw stereo")',
            '=SMILES2IMG(C15, "#1E293B", "white")'
        )
    }
    Add-StyledTable $s0 5 $t0Headers $t0Rows $t0Formulas 0x88940D $t0Widths 72 @(2,3,4,6,7) @(3,4)

    Add-SectionHeader $s0 17 7 'B. Playbook Roadmap — What Each Tab Teaches (Click Any Tab Below to Jump In)'
    $mapHeaders = @('Tab', 'Sheet Name', 'Level', 'Primary Goal', 'Key Functions / Tools Demonstrated', 'Hands-On Exercise Included', 'Best For')
    $mapRows = @(
        @('0', '0. Start Here',            'Beginner',     'Instant orientation & live sandbox',           'Top 10 formulas + shortcuts',                     'Edit Col C & double-click Col E',           'First-time users'),
        @('1', '1. Structure Rendering',   'Beginner',     'Render from Name, CAS, SMILES & InChI',        '=NAME2IMG, =SMILES2IMG, =INCHI2IMG',              'Drag down across 12 compound rows',         'Catalogs & inventories'),
        @('2', '2. Properties & Lipinski', 'Intermediate', '18-col medicinal chemistry screening table',   'All 14 property & 3D complexity UDFs + KPIs',     'Live COUNTIF & AVERAGE dashboard',          'Drug discovery & ADMET'),
        @('3', '3. SAR & Similarity',      'Intermediate', 'Substructure highlighting & Tanimoto rank',    '=SMILES.MATCH, hl:SMARTS, =SMILES.SIMILARITY',    'Edit Query B3 + built-in 12-group SMARTS',  'SAR & hit expansion'),
        @('4', '4. Notation Conversions',  'Intermediate', 'Standardize & deduplicate chemical IDs',       '=NAME2SMILES, =INCHI2*, =SMILES2INCHIKEY',        'Automatic duplicate detector via InChIKey', 'Database cleaning'),
        @('5', '5. Styles & Transforms',   'Intermediate', 'ACS B&W, Dark Mode, NMR numbering, rotation',  'All 16 style presets + full token cheat sheet',   'Edit benchmark molecule B3',                'Papers & presentations'),
        @('6', '6. Reactions & KET',       'Advanced',     'Multi-step reaction schemes & KET JSON',       '=KET2IMG, =KET2FORMULA, =KET2MW, =KET2EXACTMASS', 'Double-click Aspirin synthesis in Ketcher', 'Synthetic chemistry'),
        @('7', '7. Real-World Workflows',  'Production',   'Copy-ready Stoichiometry & MPO Triage labs',   'Live mg/mmol scaling + Yield % + MPO IF/AND',     'Change Target mmol (H6) to rescale lab',    'Daily bench & med-chem'),
        @('8', '8. Legacy Excel (.FLOAT)', 'Compat',       'Excel 2016 / 2019 / 2021 floating shapes',     '=SMILES2IMG.FLOAT, =NAME2IMG.FLOAT',              '10 floating cell-anchored shapes',          'Excel 2016-2021 users'),
        @('9', '9. Ribbon, SDF & Ref',     'Reference',    'Complete 40-UDF catalog, Ribbon & .SDF guide', 'All 40 UDFs, 3-level context menu, Error codes',  'Export any tab to .sdf & re-import',        'Offline desktop reference')
    )
    Add-StyledTable $s0 18 $mapHeaders $mapRows @{} 0x475569 @() 28 @(2,4,5,6,7) @()

    # =========================================================================
    # TAB 1: Structure Rendering (Name, CAS, SMILES, InChI)
    # =========================================================================
    Write-Host '[2/10] Building Tab 1: Structure Rendering...' -ForegroundColor Yellow
    $s1 = $sheets[1]
    Set-BannerAndCallout $s1 8 `
        '1. Structure Rendering — Auto-Render from Name, CAS, SMILES & InChI' `
        'Tip: Set Row Height to 62-80 pt for comfortable viewing. On Excel 365/2024, images live natively inside cells and sort/filter with your table. On Excel 2016-2021, append .FLOAT (see Tab 8).' `
        0xB4661B

    $t1Headers = @('#', 'Input Type', 'Input (Name / CAS / SMILES / InChI)', 'Formula Used', '2D Structure (In-Cell)', 'Resolved SMILES', 'Resolved Name', 'Chemical Class & Notes')
    $t1Widths  = @(5, 15, 34, 24, 26, 36, 20, 42)
    $t1Rows = @(
        @(1,  'Common Name', 'Aspirin',                                                                   '=NAME2IMG(C4)',   '', '=NAME2SMILES(C4)',   '=SMILES2NAME(F4)',   'Acetylsalicylic acid — NSAID analgesic'),
        @(2,  'Common Name', 'Ibuprofen',                                                                 '=NAME2IMG(C5)',   '', '=NAME2SMILES(C5)',   '=SMILES2NAME(F5)',   'Propionic acid NSAID with 1 chiral center'),
        @(3,  'Common Name', 'Caffeine',                                                                  '=NAME2IMG(C6)',   '', '=NAME2SMILES(C6)',   '=SMILES2NAME(F6)',   'Methylxanthine purine alkaloid'),
        @(4,  'Common Name', 'Paracetamol',                                                               '=NAME2IMG(C7)',   '', '=NAME2SMILES(C7)',   '=SMILES2NAME(F7)',   'Acetaminophen — para-aminophenol analgesic'),
        @(5,  'CAS Number',  '50-78-2',                                                                   '=NAME2IMG(C8)',   '', '=NAME2SMILES(C8)',   '=SMILES2NAME(F8)',   'Aspirin CAS number — no SMILES required'),
        @(6,  'CAS Number',  '58-08-2',                                                                   '=NAME2IMG(C9)',   '', '=NAME2SMILES(C9)',   '=SMILES2NAME(F9)',   'Caffeine CAS registry lookup'),
        @(7,  'CAS Number',  '103-90-2',                                                                  '=NAME2IMG(C10)',  '', '=NAME2SMILES(C10)',  '=SMILES2NAME(F10)',  'Acetaminophen / Paracetamol CAS lookup'),
        @(8,  'SMILES',      'c1ccccc1',                                                                  '=SMILES2IMG(C11)','', 'c1ccccc1',           '=SMILES2NAME(C11)',  'Benzene — offline rendering (0 ms network)'),
        @(9,  'SMILES',      'CCO',                                                                       '=SMILES2IMG(C12)','', 'CCO',                '=SMILES2NAME(C12)',  'Ethanol — primary aliphatic alcohol'),
        @(10, 'SMILES',      'CN1CCC[C@H]1c2cccnc2',                                                      '=SMILES2IMG(C13)','', 'CN1CCC[C@H]1c2cccnc2','=SMILES2NAME(C13)', 'Nicotine — chiral pyrrolidine + pyridine'),
        @(11, 'SMILES',      'O=Cc1ccc(O)c(OC)c1',                                                        '=SMILES2IMG(C14)','', 'O=Cc1ccc(O)c(OC)c1', '=SMILES2NAME(C14)',  'Vanillin — phenolic aromatic aldehyde'),
        @(12, 'IUPAC InChI', 'InChI=1S/C8H10N4O2/c1-10-4-9-6-5(10)7(13)12(3)8(14)11(6)2/h4H,1-3H3',       '=INCHI2IMG(C15)', '', '=INCHI2SMILES(C15)', '=INCHI2NAME(C15)',   'Direct rendering & decoding from IUPAC InChI')
    )
    $t1Formulas = @{
        'E' = @('=NAME2IMG(C4)','=NAME2IMG(C5)','=NAME2IMG(C6)','=NAME2IMG(C7)','=NAME2IMG(C8)','=NAME2IMG(C9)','=NAME2IMG(C10)','=SMILES2IMG(C11)','=SMILES2IMG(C12)','=SMILES2IMG(C13)','=SMILES2IMG(C14)','=INCHI2IMG(C15)');
        'F' = @('=NAME2SMILES(C4)','=NAME2SMILES(C5)','=NAME2SMILES(C6)','=NAME2SMILES(C7)','=NAME2SMILES(C8)','=NAME2SMILES(C9)','=NAME2SMILES(C10)','c1ccccc1','CCO','CN1CCC[C@H]1c2cccnc2','O=Cc1ccc(O)c(OC)c1','=INCHI2SMILES(C15)');
        'G' = @('=SMILES2NAME(F4)','=SMILES2NAME(F5)','=SMILES2NAME(F6)','=SMILES2NAME(F7)','=SMILES2NAME(F8)','=SMILES2NAME(F9)','=SMILES2NAME(F10)','=SMILES2NAME(C11)','=SMILES2NAME(C12)','=SMILES2NAME(C13)','=SMILES2NAME(C14)','=INCHI2NAME(C15)')
    }
    Add-StyledTable $s1 3 $t1Headers $t1Rows $t1Formulas 0xC97A22 $t1Widths 74 @(2,3,4,6,8) @(3,4,6)

    # =========================================================================
    # TAB 2: Physicochemical Properties, 3D Complexity & Lipinski Ro5 Dashboard
    # =========================================================================
    Write-Host '[3/10] Building Tab 2: Properties & Lipinski Dashboard...' -ForegroundColor Yellow
    $s2 = $sheets[2]
    $t2Cols = @(
        @{ H = '#';              W = 5;  Fn = $null },
        @{ H = 'Compound';       W = 15; Fn = $null },
        @{ H = 'SMILES';         W = 24; Fn = $null },
        @{ H = '2D Structure';   W = 22; Fn = 'SMILES2IMG' },
        @{ H = 'Formula';        W = 13; Fn = 'SMILES2FORMULA' },
        @{ H = 'MW (g/mol)';     W = 11; Fn = 'SMILES2MW' },
        @{ H = 'Exact Mass';     W = 12; Fn = 'SMILES2EXACTMASS' },
        @{ H = 'LogP';           W = 8;  Fn = 'SMILES2LOGP' },
        @{ H = 'TPSA (A^2)';     W = 10; Fn = 'SMILES2TPSA' },
        @{ H = 'HBD';            W = 7;  Fn = 'SMILES2HBD' },
        @{ H = 'HBA';            W = 7;  Fn = 'SMILES2HBA' },
        @{ H = 'RotB';           W = 7;  Fn = 'SMILES2ROTBONDS' },
        @{ H = 'Heavy';          W = 7;  Fn = 'SMILES2HEAVY' },
        @{ H = 'Rings';          W = 7;  Fn = 'SMILES2RINGS' },
        @{ H = 'Arom';           W = 7;  Fn = 'SMILES2AROMATIC' },
        @{ H = 'Fsp3';           W = 8;  Fn = 'SMILES2FSP3' },
        @{ H = 'Chiral';         W = 7;  Fn = 'SMILES2CHIRAL' },
        @{ H = 'Lipinski Ro5';   W = 24; Fn = 'SMILES2LIPINSKI' },
        @{ H = 'Med-Chem Note';  W = 34; Fn = $null }
    )
    Set-BannerAndCallout $s2 $t2Cols.Length `
        '2. Physicochemical Descriptors, 3D Complexity & Lipinski Ro5 Screening (100% Offline)' `
        'Lipinski Rule of 5: MW <= 500 | LogP <= 5 | HBD <= 5 | HBA <= 10. Escape from Flatland: Fsp3 >= 0.25 and Aromatic Rings <= 3 improve solubility. All numeric outputs work directly inside AVERAGE / COUNTIF / IF!' `
        0x417C10

    $t2Drugs = @(
        @('Aspirin',      'CC(=O)Oc1ccccc1C(=O)O',                                                                                           'NSAID; 0 Ro5 violations; flat aromatic (Fsp3=0.11)'),
        @('Ibuprofen',    'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O',                                                                                 'NSAID; 0 violations; 1 chiral center; good Fsp3'),
        @('Caffeine',     'Cn1cnc2c1c(=O)n(C)c(=O)n2C',                                                                                      'Purine CNS stimulant; low TPSA; passes BBB'),
        @('Paracetamol',  'CC(=O)Nc1ccc(O)cc1',                                                                                              'Analgesic; compact MW=151; 0 violations'),
        @('Penicillin G', 'CC1(C(N2C(S1)C(C2=O)NC(=O)Cc3ccccc3)C(=O)O)C',                                                                    'Beta-lactam antibiotic; 3 chiral centers'),
        @('Morphine',     'CN1CC[C@]23[C@@H]4Oc5c3c(CC1[C@@H]2C=C[C@@H]4O)ccc5O',                                                            'Opioid alkaloid; rigid 5-ring 3D scaffold'),
        @('Dopamine',     'NCCc1ccc(O)c(O)c1',                                                                                               'Neurotransmitter; low MW catecholamine'),
        @('Nicotine',     'CN1CCC[C@H]1c2cccnc2',                                                                                            'Cholinergic ligand; 1 chiral center; LogP=1.2'),
        @('Artemisinin',  'CC1CCC2C(C(=O)OC3C24C1CCC(O3)(OO4)C)C',                                                                           'Nobel antimalarial; sp3-rich (Fsp3=0.93, 0 arom)'),
        @('Cholesterol',  'CC(C)CCCC(C)C1CCC2C1(CCC3C2CC=C4C3(CCC(C4)O)C)C',                                                                 'Sterol lipid; fails Ro5 due to LogP > 5'),
        @('Atorvastatin', 'CC(C)c1c(c(c(c1NC(=O)c2ccccc2)c3ccc(cc3)F)c4ccccc4)CC(O)CC(O)CC(=O)O',                                            'Statin; fails Ro5 (MW > 500, LogP > 5)'),
        @('Paclitaxel',   'CC1=C2C(C(=O)C3(C(CC4C(C3C(C(C2(C)C)(CC1OC(=O)C(C(c5ccccc5)NC(=O)c6ccccc6)O)O)OC(=O)c7ccccc7)(CO4)OC(=O)C)O)C)OC(=O)C','Taxol natural product; exceeds Ro5 (MW=854)')
    )
    $t2Start = 4
    $t2End = $t2Start + $t2Drugs.Length - 1
    $t2Rows = @()
    for ($i = 0; $i -lt $t2Drugs.Length; $i++) {
        $r = New-Object object[] $t2Cols.Length
        for ($c = 0; $c -lt $r.Length; $c++) { $r[$c] = '' }
        $r[0] = $i + 1; $r[1] = $t2Drugs[$i][0]; $r[2] = $t2Drugs[$i][1]; $r[$r.Length - 1] = $t2Drugs[$i][2]
        $t2Rows += ,$r
    }
    $t2Formulas = @{}
    for ($c = 0; $c -lt $t2Cols.Length; $c++) {
        if ($t2Cols[$c].Fn) {
            $colL = Get-ColLetter ($c + 1)
            $t2Formulas[$colL] = @($t2Start..$t2End | ForEach-Object { "=$($t2Cols[$c].Fn)(C$_)" })
        }
    }
    Add-StyledTable $s2 3 @($t2Cols | ForEach-Object { $_.H }) $t2Rows $t2Formulas 0x4E8E1E @($t2Cols | ForEach-Object { $_.W }) 74 @(2,3,19) @(3,5)

    # Live KPI Summary Row below the 12-drug screening table (Row 17)
    $kpiRow = $t2End + 2
    Add-SectionHeader $s2 ($kpiRow - 1) $t2Cols.Length 'Live Excel Summary Metrics (Calculated Directly from Numeric UDF Outputs Above)' 0xE6F4EA 0x134E4A
    Invoke-Excel {
        $s2.Range("B${kpiRow}").Value2 = 'Library Average / KPI'
        $s2.Range("B${kpiRow}").Font.Bold = $true
        $s2.Range("F${kpiRow}").Formula = "=ROUND(AVERAGE(F${t2Start}:F${t2End}),2)"
        $s2.Range("H${kpiRow}").Formula = "=ROUND(AVERAGE(H${t2Start}:H${t2End}),2)"
        $s2.Range("I${kpiRow}").Formula = "=ROUND(AVERAGE(I${t2Start}:I${t2End}),1)"
        $s2.Range("M${kpiRow}").Formula = "=ROUND(AVERAGE(M${t2Start}:M${t2End}),1)"
        $s2.Range("P${kpiRow}").Formula = "=ROUND(AVERAGE(P${t2Start}:P${t2End}),2)"
        $s2.Range("R${kpiRow}").Formula = "=COUNTIF(R${t2Start}:R${t2End},`"PASS`") & `" / `" & COUNTA(R${t2Start}:R${t2End}) & `" PASS`""
        $s2.Range("S${kpiRow}").Value2 = 'Shows UDFs return true Excel numbers & strings for native analytics!'
        $kr = $s2.Range("A${kpiRow}:S${kpiRow}")
        $kr.Interior.Color = 0xF0FDF4
        $kr.Font.Bold = $true
        $kr.RowHeight = 28
        $kr.Borders.LineStyle = 1
        $kr.Borders.Color = 0xB7E4C7
        $kr.VerticalAlignment = -4108
        $kr.HorizontalAlignment = -4108
        $s2.Range("S${kpiRow}").HorizontalAlignment = -4131
    }

    # =========================================================================
    # TAB 3: SAR Substructure Screening, Similarity Ranking & SMARTS Dictionary
    # =========================================================================
    Write-Host '[4/10] Building Tab 3: SAR, Similarity & SMARTS Dictionary...' -ForegroundColor Yellow
    $s3 = $sheets[3]
    Set-BannerAndCallout $s3 8 `
        '3. Substructure Search (=SMILES.MATCH), Highlighting (hl:) & Tanimoto Similarity (=SMILES.SIMILARITY)' `
        'Interactive Exercise: Edit Cell B3 (copy any SMARTS from the 12-group dictionary below!) or Cell F3 (Lead SMILES) to immediately recompute matches, red ring highlights, and similarity rankings.' `
        0x993366

    Invoke-Excel {
        $s3.Range('A3').Value2 = 'Query SMARTS ($B$3):'
        $s3.Range('A3').Font.Bold = $true
        $s3.Range('A3').HorizontalAlignment = -4152
        $s3.Range('B3').Value2 = 'c1ccccc1'
        $s3.Range('B3').Font.Bold = $true
        $s3.Range('B3').Font.Name = 'Consolas'
        $s3.Range('B3').Interior.Color = 0xE2E2FE
        $s3.Range('B3').HorizontalAlignment = -4108
        $s3.Range('C3').Value2 = 'Try: C(=O)O | c1ccncc1 | c1ccc(O)cc1'
        $s3.Range('C3').Font.Italic = $true
        $s3.Range('E3').Value2 = 'Lead SMILES ($F$3):'
        $s3.Range('E3').Font.Bold = $true
        $s3.Range('E3').HorizontalAlignment = -4152
        $s3.Range('F3').Value2 = 'CC(=O)Oc1ccccc1C(=O)O'
        $s3.Range('F3').Font.Bold = $true
        $s3.Range('F3').Font.Name = 'Consolas'
        $s3.Range('F3').Interior.Color = 0xFCE8D5
        $s3.Range('F3').HorizontalAlignment = -4108
        $s3.Range('G3').Value2 = '(Aspirin Lead)'
        $s3.Range('G3').Font.Italic = $true
        $s3.Rows.Item(3).RowHeight = 25
    }

    $t3Headers = @('#', 'Compound', 'Target SMILES', 'Match? (=SMILES.MATCH)', 'Highlighted Structure (hl:)', 'Tanimoto Similarity', 'Similarity Rank', 'SAR Pharmacophore Interpretation')
    $t3Widths  = @(5, 18, 30, 22, 26, 20, 14, 44)
    $t3Rows = @(
        @(1,  'Aspirin',     'CC(=O)Oc1ccccc1C(=O)O',                    '', '', '', '', 'Exact lead match (Similarity = 1.00); contains benzene core'),
        @(2,  'Salicylic Acid','O=C(O)c1ccccc1O',                        '', '', '', '', 'Aspirin precursor; high scaffold similarity + benzene ring'),
        @(3,  'Ibuprofen',   'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O',          '', '', '', '', 'Para-substituted benzene highlighted in bright red'),
        @(4,  'Paracetamol', 'CC(=O)Nc1ccc(O)cc1',                       '', '', '', '', 'Acetamide phenol sharing benzene ring with Aspirin'),
        @(5,  'Vanillin',    'O=Cc1ccc(O)c(OC)c1',                       '', '', '', '', 'Substituted phenolic aldehyde containing benzene'),
        @(6,  'Benzene',     'c1ccccc1',                                 '', '', '', '', 'Minimal exact match to benzene SMARTS query'),
        @(7,  'Pyridine',    'c1ccncc1',                                 '', '', '', '', 'Heteroaromatic N-ring -> FALSE for all-carbon benzene'),
        @(8,  'Nicotine',    'CN1CCC[C@H]1c2cccnc2',                     '', '', '', '', 'Contains pyridine (not benzene) -> FALSE'),
        @(9,  'Caffeine',    'Cn1cnc2c1c(=O)n(C)c(=O)n2C',               '', '', '', '', 'Fused purine xanthine heterocycle -> FALSE'),
        @(10, 'Artemisinin', 'CC1CCC2C(C(=O)OC3C24C1CCC(O3)(OO4)C)C',   '', '', '', '', 'All-aliphatic terpene endoperoxide -> FALSE')
    )
    $t3Formulas = @{
        'D' = @(5..14 | ForEach-Object { "=SMILES.MATCH(C$_, `$B`$3)" });
        'E' = @(5..14 | ForEach-Object { "=SMILES2IMG(C$_, , `"hl:`" & `$B`$3)" });
        'F' = @(5..14 | ForEach-Object { "=ROUND(SMILES.SIMILARITY(C$_, `$F`$3), 3)" });
        'G' = @(5..14 | ForEach-Object { "=RANK(F$_, `$F`$5:`$F`$14)" })
    }
    Add-StyledTable $s3 4 $t3Headers $t3Rows $t3Formulas 0xB34D80 $t3Widths 72 @(2,3,8) @(3)

    Add-SectionHeader $s3 16 8 'Built-In 12-Group Medicinal Chemistry SMARTS Quick-Copy Dictionary (Copy Any Pattern into Cell B3 Above!)'
    $smartsHeaders = @('#', 'Functional Group', 'SMARTS Pattern', 'Example Formula', 'Target Matches', 'Chemotype Category', 'Clinical / Synthetic Relevance', 'Copy Tip')
    $smartsRows = @(
        @(1,  'Benzene Ring',       'c1ccccc1',              '=SMILES.MATCH(C5,"c1ccccc1")',       'Aspirin, Ibuprofen', 'Aromatic',      'Core hydrophobic phenyl scaffold',             'Paste into B3'),
        @(2,  'Pyridine Ring',      'c1ccncc1',              '=SMILES.MATCH(C5,"c1ccncc1")',       'Pyridine, Nicotine', 'Heterocycle',   'H-bond acceptor aromatic N-ring',              'Paste into B3'),
        @(3,  'Carboxylic Acid',    'C(=O)[OH]',             '=SMILES.MATCH(C5,"C(=O)[OH]")',      'Aspirin, Ibuprofen', 'Acidic Group',  'Ionizable anionic pharmacophore (pKa ~4)',     'Paste into B3'),
        @(4,  'Ester',              '[#6]C(=O)O[#6]',        '=SMILES.MATCH(C5,"[#6]C(=O)O[#6]")', 'Aspirin, Paclitaxel','Prodrug Linker','Hydrolyzable ester prodrug motif',             'Paste into B3'),
        @(5,  'Amide Bond',         'C(=O)N',                '=SMILES.MATCH(C5,"C(=O)N")',         'Paracetamol, Pen G', 'Peptidomimetic','Metabolically stable H-bond linker',           'Paste into B3'),
        @(6,  'Phenol (-ArOH)',     'c[OH]',                 '=SMILES.MATCH(C5,"c[OH]")',          'Salicylic, Vanillin','H-Bond Donor',  'Antioxidant / phase-II glucuronidation site',  'Paste into B3'),
        @(7,  'Primary Amine',      '[NX3;H2]',              '=SMILES.MATCH(C5,"[NX3;H2]")',       'Dopamine',           'Basic Group',   'Cationic amine at physiological pH',           'Paste into B3'),
        @(8,  'Aryl Halide (F/Cl)', '[c][F,Cl,Br,I]',        '=SMILES.MATCH(C5,"[c][F,Cl,Br,I]")', 'Atorvastatin',       'Halogenated',   'Metabolic blocker / halogen bonding',          'Paste into B3'),
        @(9,  'Sulfonamide',        'S(=O)(=O)N',            '=SMILES.MATCH(C5,"S(=O)(=O)N")',     'Sulfa drugs',        'Bioisostere',   'Classic carboxylic acid bioisostere',          'Paste into B3'),
        @(10, 'Aldehyde',           '[CX3H1](=O)[#6]',       '=SMILES.MATCH(C5,"[CX3H1](=O)[#6]")','Vanillin',           'Reactive Alert','Electrophilic carbonyl / PAINS alert',         'Paste into B3'),
        @(11, 'Nitro Group',        '[$([NX3](=O)=O)]',      '=SMILES.MATCH(C5,"[$([NX3](=O)=O)]")','Nitroaromatics',    'Toxicity Alert','Mutagenicity / structural alert screen',       'Paste into B3'),
        @(12, 'Michael Acceptor',   '[C]=[C]-C(=O)',         '=SMILES.MATCH(C5,"[C]=[C]-C(=O)")',  'Covalent inhibitors','Covalent Warhead','Targeted covalent cysteine warhead',          'Paste into B3')
    )
    Add-StyledTable $s3 17 $smartsHeaders $smartsRows @{} 0x6B2D5C @() 26 @(2,3,4,5,7) @(3,4)

    # =========================================================================
    # TAB 4: Notation Conversions & Duplicate Structure Detection
    # =========================================================================
    Write-Host '[5/10] Building Tab 4: Notation Conversions & Deduplication...' -ForegroundColor Yellow
    $s4 = $sheets[4]
    Set-BannerAndCallout $s4 9 `
        '4. Cross-Notation Conversions & Chemical Library Deduplication (SMILES <-> InChI <-> InChIKey <-> Name)' `
        'Real-World Trick: Two chemists often write different valid SMILES for the exact same molecule (see Rows 8 & 9 below). Use =SMILES2INCHIKEY to automatically detect duplicates with 100% accuracy!' `
        0x8F5E00

    $t4Headers = @('#', 'Label / Source', 'Raw SMILES Input', '2D Structure', 'IUPAC InChI (=SMILES2INCHI)', 'InChIKey (=SMILES2INCHIKEY)', 'Round-Trip SMILES (=INCHI2SMILES)', 'MW via InChI (=INCHI2MW)', 'Duplicate Check (vs Row Above)')
    $t4Widths  = @(5, 22, 28, 24, 38, 28, 28, 15, 24)
    $t4Rows = @(
        @(1, 'Aspirin (Vendor A)',     'CC(=O)Oc1ccccc1C(=O)O',           '', '', '', '', '', ''),
        @(2, 'Caffeine (Vendor A)',    'Cn1cnc2c1c(=O)n(C)c(=O)n2C',      '', '', '', '', '', ''),
        @(3, 'Ibuprofen (Vendor A)',   'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O', '', '', '', '', '', ''),
        @(4, 'Paracetamol (Vendor A)', 'CC(=O)Nc1ccc(O)cc1',              '', '', '', '', '', ''),
        @(5, 'Aspirin (Alt SMILES 1)', 'CC(=O)Oc1ccccc1C(=O)O',           '', '', '', '', '', ''),
        @(6, 'Aspirin (Alt SMILES 2)', 'O=C(O)c1ccccc1OC(=O)C',           '', '', '', '', '', '')
    )
    $t4Formulas = @{
        'D' = @(4..9 | ForEach-Object { "=SMILES2IMG(C$_)" });
        'E' = @(4..9 | ForEach-Object { "=SMILES2INCHI(C$_)" });
        'F' = @(4..9 | ForEach-Object { "=SMILES2INCHIKEY(C$_)" });
        'G' = @(4..9 | ForEach-Object { "=INCHI2SMILES(E$_)" });
        'H' = @(4..9 | ForEach-Object { "=INCHI2MW(E$_)" });
        'I' = @(
            'Baseline',
            '=IF(COUNTIF($F$4:F5, F5)>1, "DUPLICATE!", "Unique")',
            '=IF(COUNTIF($F$4:F6, F6)>1, "DUPLICATE!", "Unique")',
            '=IF(COUNTIF($F$4:F7, F7)>1, "DUPLICATE!", "Unique")',
            '=IF(COUNTIF($F$4:F8, F8)>1, "DUPLICATE!", "Unique")',
            '=IF(COUNTIF($F$4:F9, F9)>1, "DUPLICATE!", "Unique")'
        )
    }
    Add-StyledTable $s4 3 $t4Headers $t4Rows $t4Formulas 0xA66E00 $t4Widths 76 @(2,3,5,7) @(3,5,6,7)
    Invoke-Excel {
        $s4.Range('I8:I9').Interior.Color = 0xD7D7FE
        $s4.Range('I8:I9').Font.Bold = $true
        $s4.Range('I8:I9').Font.Color = 0x1C1CB9
    }

    # =========================================================================
    # TAB 5: Drawing Styles, ACS B&W, Dark Mode & Transforms Matrix
    # =========================================================================
    Write-Host '[6/10] Building Tab 5: Styles & Transforms Matrix...' -ForegroundColor Yellow
    $s5 = $sheets[5]
    Set-BannerAndCallout $s5 7 `
        '5. Drawing Styles, Publication ACS B&W, Dark Mode & Geometric Transforms' `
        'Syntax: =SMILES2IMG(smiles, [bg], [style], [transform]). Mix multiple style tokens with spaces ("bw num stereo") or pipes ("bw|stereo"). Change Cell B3 to preview all 16 styles on your own molecule!' `
        0x0066CC

    Invoke-Excel {
        $s5.Range('A3').Value2 = 'Test SMILES ($B$3):'
        $s5.Range('A3').Font.Bold = $true
        $s5.Range('A3').HorizontalAlignment = -4152
        $s5.Range('B3').Value2 = 'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O'
        $s5.Range('B3').Font.Bold = $true
        $s5.Range('B3').Font.Name = 'Consolas'
        $s5.Range('B3').Interior.Color = 0xFCE8D5
        $s5.Range('B3').HorizontalAlignment = -4108
        $s5.Range('C3').Value2 = '(Ibuprofen — shows chiral stereo + benzene highlight cleanly)'
        $s5.Range('C3').Font.Italic = $true
        $s5.Rows.Item(3).RowHeight = 25
    }

    $t5Headers = @('#', 'Style Preset', 'Arg2 [bg]', 'Arg3 [style] / Arg4 [transform]', 'Formula Syntax (Copyable)', 'Live Rendered Output', 'Use Case & Publication Standard')
    $t5Widths  = @(5, 24, 15, 28, 52, 26, 44)
    $t5Rows = @(
        @(1,  'Standard CPK Color',      '(default)',   '(default)',                  '=SMILES2IMG($B$3)',                                          '', 'Transparent background + IUPAC element colors'),
        @(2,  'Solid White Canvas',      '"white"',     '(default)',                  '=SMILES2IMG($B$3, "white")',                                 '', 'Opaque white fill for clean PDF printing'),
        @(3,  'ACS Publication B&W',     '(default)',   '"bw"',                       '=SMILES2IMG($B$3, , "bw")',                                  '', 'Pure black bonds for ACS, RSC, Nature, patents'),
        @(4,  'Dark Mode (White Ink)',   '"#1E293B"',   '"white"',                    '=SMILES2IMG($B$3, "#1E293B", "white")',                      '', 'Crisp white bonds over dark cells / dark slides'),
        @(5,  'Custom Corporate Ink',    '"aliceblue"', '"ink=#1E40AF"',              '=SMILES2IMG($B$3, "aliceblue", "ink=#1E40AF")',              '', 'Any hex or CSS brand ink color'),
        @(6,  'NMR Atom Numbers',        '(default)',   '"num"',                      '=SMILES2IMG($B$3, , "num")',                                 '', '1-based atom indices for 1H/13C NMR assignment'),
        @(7,  'Chiral Stereolabels',     '(default)',   '"stereo"',                   '=SMILES2IMG($B$3, , "stereo")',                              '', 'Explicit Cahn-Ingold-Prelog (R/S, E/Z) labels'),
        @(8,  'Substructure Highlight',  '(default)',   '"hl:c1ccccc1"',              '=SMILES2IMG($B$3, , "hl:c1ccccc1")',                         '', 'Highlights phenyl scaffold in bright red'),
        @(9,  'Explicit Hydrogens',      '(default)',   '"h"',                        '=SMILES2IMG($B$3, , "h")',                                   '', 'Unfolds all implicit C-H bonds for steric review'),
        @(10, 'Show All Carbon "C"',     '(default)',   '"all-c"',                    '=SMILES2IMG($B$3, , "all-c")',                               '', 'Displays explicit C labels for teaching figures'),
        @(11, 'Chained: B&W + Stereo',   '(default)',   '"bw stereo"',                '=SMILES2IMG($B$3, , "bw stereo")',                           '', 'Space-chained ACS B&W + stereocenter labels'),
        @(12, 'Chained: Pipe Delimiter', '(default)',   '"bw|num"',                   '=SMILES2IMG($B$3, , "bw|num")',                              '', 'Pipe-delimited monochrome + NMR atom indices'),
        @(13, 'Rotate 90 deg CW',        '(default)',   'transform="rot90"',          '=SMILES2IMG($B$3, , , "rot90")',                             '', 'Compact vertical orientation for narrow columns'),
        @(14, 'Rotate 180 deg',          '(default)',   'transform="rot180"',         '=SMILES2IMG($B$3, , , "rot180")',                            '', '180-degree planar orientation alignment'),
        @(15, 'Horizontal Mirror Flip',  '(default)',   'transform="flip"',           '=SMILES2IMG($B$3, , , "flip")',                              '', 'Left-right mirror flip for SAR series alignment'),
        @(16, 'All-in-One Combined',     '"white"',     '"bw num stereo", "flip rot90"','=SMILES2IMG($B$3, "white", "bw num stereo", "flip rot90")','', 'White BG + B&W + Numbers + Stereo + Flip + Rot90')
    )
    $t5Formulas = @{
        'F' = @(
            '=SMILES2IMG($B$3)',
            '=SMILES2IMG($B$3, "white")',
            '=SMILES2IMG($B$3, , "bw")',
            '=SMILES2IMG($B$3, "#1E293B", "white")',
            '=SMILES2IMG($B$3, "aliceblue", "ink=#1E40AF")',
            '=SMILES2IMG($B$3, , "num")',
            '=SMILES2IMG($B$3, , "stereo")',
            '=SMILES2IMG($B$3, , "hl:c1ccccc1")',
            '=SMILES2IMG($B$3, , "h")',
            '=SMILES2IMG($B$3, , "all-c")',
            '=SMILES2IMG($B$3, , "bw stereo")',
            '=SMILES2IMG($B$3, , "bw|num")',
            '=SMILES2IMG($B$3, , , "rot90")',
            '=SMILES2IMG($B$3, , , "rot180")',
            '=SMILES2IMG($B$3, , , "flip")',
            '=SMILES2IMG($B$3, "white", "bw num stereo", "flip rot90")'
        )
    }
    Add-StyledTable $s5 4 $t5Headers $t5Rows $t5Formulas 0x0080E6 $t5Widths 74 @(2,4,5,7) @(3,4,5)

    # =========================================================================
    # TAB 6: Multi-Step Reactions & Lossless KET Canvas
    # =========================================================================
    Write-Host '[7/10] Building Tab 6: Reactions & KET Canvas...' -ForegroundColor Yellow
    $s6 = $sheets[6]
    Set-BannerAndCallout $s6 8 `
        '6. Multi-Step Reaction Schemes & Lossless EPAM Ketcher KET Canvas (=KET2IMG & =KET2*)' `
        'Interactive Exercise: Double-click any cell in Col E to launch Ketcher 2D with full arrows, + signs, and exact 2D coordinates restored. Add catalysts/conditions above the arrow and press Ctrl+Enter!' `
        0x52731E

    $t6Headers = @('#', 'Diagram / Reaction Title', 'KET JSON Data (Truncated View)', 'Image Formula', 'Live Rendered KET Diagram', 'Formula (=KET2FORMULA)', 'MW (=KET2MW)', 'Workflow & Round-Trip Notes')
    $t6Widths  = @(5, 24, 24, 22, 34, 18, 14, 42)
    $t6Rows = @(
        @(1, 'Aspirin Synthesis (Reaction)',     $ketAspirinRxn,     '=KET2IMG(C4)', '', 'Reaction Scheme', 'Multi-Comp', 'Salicylic Acid + Acetic Anhydride -> Aspirin. Double-click Col E to edit arrows!'),
        @(2, 'Paracetamol Synthesis (Reaction)', $ketParacetamolRxn, '=KET2IMG(C5)', '', 'Reaction Scheme', 'Multi-Comp', '4-Aminophenol + Acetic Anhydride -> Paracetamol. Preserves full reaction layout!'),
        @(3, 'Aspirin (Single-Mol KET)',         $ketAspirinMol,     '=KET2IMG(C6)', '', '', '', 'Exact 2D coordinates + live =KET2FORMULA & =KET2MW evaluation'),
        @(4, 'Caffeine (Single-Mol KET)',        $ketCaffeineMol,    '=KET2IMG(C7)', '', '', '', 'Exact ring orientation preserved across save/reopen/share'),
        @(5, 'Ibuprofen (Single-Mol KET)',       $ketIbuprofenMol,   '=KET2IMG(C8)', '', '', '', 'Exact 2D coordinates & functional groups preserved in Ketcher JSON')
    )
    $t6Formulas = @{
        'E' = @('=KET2IMG(C4)','=KET2IMG(C5)','=KET2IMG(C6)','=KET2IMG(C7)','=KET2IMG(C8)');
        'F' = @('Reaction Scheme','Reaction Scheme','=KET2FORMULA(C6)','=KET2FORMULA(C7)','=KET2FORMULA(C8)');
        'G' = @('N/A (Reaction)','N/A (Reaction)','=KET2MW(C6)','=KET2MW(C7)','=KET2MW(C8)')
    }
    Add-StyledTable $s6 3 $t6Headers $t6Rows $t6Formulas 0x628A24 $t6Widths 88 @(2,4,8) @(4,6)

    # =========================================================================
    # TAB 7: Real-World Lab Workflows (Stoichiometry Calculator & Hit-to-Lead MPO)
    # =========================================================================
    Write-Host '[8/10] Building Tab 7: Real-World Chemistry Workflows...' -ForegroundColor Yellow
    $s7 = $sheets[7]
    Set-BannerAndCallout $s7 10 `
        '7. Production Chemistry Templates — Organic Stoichiometry Scale-Up & Hit-to-Lead MPO Triage' `
        'Copy these two ready-to-use templates straight into your lab notebooks! In Table A, change Target Scale (Cell H6) from 10 to 50 mmol — every reagent mass (mg) and theoretical yield updates automatically!' `
        0x2647D2

    # --- Use Case A: Stoichiometry Reaction Scale-Up Sheet ---
    Add-SectionHeader $s7 4 10 'Use Case A: Organic Synthesis Stoichiometry Table (Aspirin Acetylation — Edit Cell H6 to Rescale Reaction!)' 0xE6EEFF 0x1E3A8A
    $s7AHeaders = @('Role', 'Reagent / Product', 'SMILES', '2D Structure', 'Formula', 'MW (g/mol)', 'Equiv.', 'Scale (mmol)', 'Required / Theor. Mass (mg)', 'Actual Yield & % Yield')
    $s7AWidths  = @(16, 20, 26, 24, 13, 12, 10, 13, 18, 22)
    $s7ARows = @(
        @('Limiting Reagent', 'Salicylic Acid',   'O=C(O)c1ccccc1O',       '', '', '', 1.0, 10.0, '', 'Weigh solid (1.0 equiv)'),
        @('Reagent (Excess)', 'Acetic Anhydride', 'CC(=O)OC(=O)C',         '', '', '', 1.5, '',   '', 'Liquid d=1.08 g/mL -> ~1.42 mL'),
        @('Target Product',   'Aspirin',          'CC(=O)Oc1ccccc1C(=O)O', '', '', '', 1.0, '',   '', '')
    )
    $s7AFormulas = @{
        'D' = @('=SMILES2IMG(C6)', '=SMILES2IMG(C7)', '=SMILES2IMG(C8)');
        'E' = @('=SMILES2FORMULA(C6)', '=SMILES2FORMULA(C7)', '=SMILES2FORMULA(C8)');
        'F' = @('=SMILES2MW(C6)', '=SMILES2MW(C7)', '=SMILES2MW(C8)');
        'H' = @('10', '=G7*$H$6', '=G8*$H$6');
        'I' = @('=ROUND(F6*H6, 1)', '=ROUND(F7*H7, 1)', '=ROUND(F8*H8, 1)');
        'J' = @('Weigh solid (1.0 equiv)', 'd=1.08 g/mL (~1.42 mL)', '="Isolated 1440 mg (" & TEXT(1440/I8,"0.0%") & ")"')
    }
    Add-StyledTable $s7 5 $s7AHeaders $s7ARows $s7AFormulas 0x3B5998 $s7AWidths 74 @(1,2,3,10) @(3,5)
    Invoke-Excel {
        $s7.Range('H6').Interior.Color = 0xCCF5FF
        $s7.Range('H6').Font.Bold = $true
        $s7.Range('J8').Interior.Color = 0xE6F4EA
        $s7.Range('J8').Font.Bold = $true
    }

    # --- Use Case B: Hit-to-Lead Candidate Triage (Multi-Parameter Optimization) ---
    Add-SectionHeader $s7 10 10 'Use Case B: Medicinal Chemistry Hit-to-Lead Triage (Auto-Decision: Lipinski PASS + Fsp3 >= 0.20 + LogP <= 4.0 + TPSA <= 90)' 0xF3E8FF 0x581C87
    $s7BHeaders = @('ID', 'Candidate', 'SMILES', '2D Structure', 'MW', 'LogP', 'TPSA', 'Fsp3', 'Lipinski Ro5', 'MPO Decision (Formula)')
    $s7BRows = @(
        @('CPD-01', 'Ibuprofen',   'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O', '', '', '', '', '', '', ''),
        @('CPD-02', 'Caffeine',    'Cn1cnc2c1c(=O)n(C)c(=O)n2C',      '', '', '', '', '', '', ''),
        @('CPD-03', 'Nicotine',    'CN1CCC[C@H]1c2cccnc2',            '', '', '', '', '', '', ''),
        @('CPD-04', 'Aspirin',     'CC(=O)Oc1ccccc1C(=O)O',           '', '', '', '', '', '', ''),
        @('CPD-05', 'Atorvastatin','CC(C)c1c(c(c(c1NC(=O)c2ccccc2)c3ccc(cc3)F)c4ccccc4)CC(O)CC(O)CC(=O)O', '', '', '', '', '', '', '')
    )
    $s7BFormulas = @{
        'D' = @(12..16 | ForEach-Object { "=SMILES2IMG(C$_)" });
        'E' = @(12..16 | ForEach-Object { "=SMILES2MW(C$_)" });
        'F' = @(12..16 | ForEach-Object { "=SMILES2LOGP(C$_)" });
        'G' = @(12..16 | ForEach-Object { "=SMILES2TPSA(C$_)" });
        'H' = @(12..16 | ForEach-Object { "=SMILES2FSP3(C$_)" });
        'I' = @(12..16 | ForEach-Object { "=SMILES2LIPINSKI(C$_)" });
        'J' = @(12..16 | ForEach-Object { "=IF(AND(I$_=`"PASS`",H$_>=0.2,F$_<=4,G$_<=90),`"ADVANCE`",`"OPTIMIZE`")" })
    }
    Add-StyledTable $s7 11 $s7BHeaders $s7BRows $s7BFormulas 0x7E22CE @() 72 @(2,3) @(3)
    Invoke-Excel {
        $s7.Range('I17').Value2 = 'Candidates Advancing:'
        $s7.Range('I17').Font.Bold = $true
        $s7.Range('I17').HorizontalAlignment = -4152
        $s7.Range('J17').Formula = '=COUNTIF(J12:J16,"ADVANCE") & " of " & COUNTA(J12:J16)'
        $s7.Range('J17').Font.Bold = $true
        $s7.Range('J17').Interior.Color = 0xE6F4EA
        $s7.Range('J17').HorizontalAlignment = -4108
        $s7.Rows.Item(17).RowHeight = 25
    }

    # =========================================================================
    # TAB 8: Legacy Excel 2016 / 2019 / 2021 Compatibility (.FLOAT Mode)
    # =========================================================================
    Write-Host '[9/10] Building Tab 8: Legacy Excel (.FLOAT)...' -ForegroundColor Yellow
    $s8 = $sheets[8]
    Set-BannerAndCallout $s8 8 `
        '8. Legacy Excel 2016 / 2019 / 2021 Compatibility (=SMILES2IMG.FLOAT & =NAME2IMG.FLOAT)' `
        'Excel 2016-2021 lacks native in-cell pictures — append .FLOAT to any image formula to insert cell-anchored shapes (xlMoveAndSize). All 32 property/SAR/conversion formulas work identically without .FLOAT!' `
        0x336633

    $t8Headers = @('#', 'Mode', 'Compound', 'Input (SMILES / Name)', 'Style Preset', 'Formula Syntax (Copyable)', 'Live Floating Shape (.FLOAT)', 'Compatibility Notes')
    $t8Widths  = @(5, 14, 16, 28, 22, 46, 28, 42)
    $t8Rows = @(
        @(1,  'SMILES',    'Ethanol',      'CCO',                              'Default Color',    '=SMILES2IMG.FLOAT(D4)',                                      '', 'Anchored floating shape that moves & resizes with cell'),
        @(2,  'Name',      'Aspirin',      'Aspirin',                          'Name Lookup',      '=NAME2IMG.FLOAT(D5)',                                        '', 'PubChem name lookup works identically in .FLOAT mode'),
        @(3,  'Background','Paracetamol',  'CC(=O)Nc1ccc(O)cc1',               'White Canvas',     '=SMILES2IMG.FLOAT(D6, "white")',                             '', 'Opaque white background card'),
        @(4,  'Monochrome','Aspirin',      'CC(=O)Oc1ccccc1C(=O)O',            'ACS B&W',          '=SMILES2IMG.FLOAT(D7, , "bw")',                              '', 'Monochrome patent/journal line art on Excel 2016-2021'),
        @(5,  'Stereo',    'L-Lactic Acid','C[C@H](O)C(=O)O',                  'Chiral R/S',       '=SMILES2IMG.FLOAT(D8, , "stereo")',                          '', 'Displays absolute stereochemistry labels'),
        @(6,  'Hydrogens', 'Ethanol',      'CCO',                              'Explicit H',       '=SMILES2IMG.FLOAT(D9, , "h")',                               '', 'Unfolded C-H hydrogens'),
        @(7,  'Combined',  'Aspirin',      'CC(=O)Oc1ccccc1C(=O)O',            'B&W + NMR #',      '=SMILES2IMG.FLOAT(D10, , "bw num")',                         '', 'Chained space options work identically in .FLOAT'),
        @(8,  'Dark Mode', 'Nicotine',     'CN1CCC[C@H]1c2cccnc2',             'Dark + Stereo',    '=SMILES2IMG.FLOAT(D11, "#1E293B", "white stereo")',          '', 'White ink over dark background'),
        @(9,  'Transform', 'Nicotine',     'CN1CCC[C@H]1c2cccnc2',             'Mirror Flip',      '=SMILES2IMG.FLOAT(D12, , , "flip")',                         '', 'Horizontal mirror flip'),
        @(10, 'All-in-One','Aspirin',      'CC(=O)Oc1ccccc1C(=O)O',            'White+Full+Rot90', '=SMILES2IMG.FLOAT(D13, "white", "bw num stereo", "rot90")',  '', 'Full option matrix supported on legacy Excel')
    )
    $t8Formulas = @{
        'G' = @(
            '=SMILES2IMG.FLOAT(D4)',
            '=NAME2IMG.FLOAT(D5)',
            '=SMILES2IMG.FLOAT(D6, "white")',
            '=SMILES2IMG.FLOAT(D7, , "bw")',
            '=SMILES2IMG.FLOAT(D8, , "stereo")',
            '=SMILES2IMG.FLOAT(D9, , "h")',
            '=SMILES2IMG.FLOAT(D10, , "bw num")',
            '=SMILES2IMG.FLOAT(D11, "#1E293B", "white stereo")',
            '=SMILES2IMG.FLOAT(D12, , , "flip")',
            '=SMILES2IMG.FLOAT(D13, "white", "bw num stereo", "rot90")'
        )
    }
    Add-StyledTable $s8 3 $t8Headers $t8Rows $t8Formulas 0x3D7A3D $t8Widths 76 @(2,3,4,6,8) @(4,6)

    # =========================================================================
    # TAB 9: Complete Offline Reference (Ribbon, 3-Level Menu, SDF & All 40 UDFs)
    # =========================================================================
    Write-Host '[10/10] Building Tab 9: Complete 40-Function, Ribbon & SDF Reference...' -ForegroundColor Yellow
    $s9 = $sheets[9]
    Set-BannerAndCallout $s9 6 `
        "9. Complete Offline Desktop Reference — [SMILES] Ribbon, .SDF Database I/O & All 40 Excel Functions" `
        'Everything you need in one tab: exact v1.8.0 Ribbon buttons, 3-level right-click menu, 12-column .sdf export/import guide, all 40 functions, and #N/A / #VALUE! / #NAME? fixes.' `
        0x3B291E

    Add-SectionHeader $s9 4 6 'A. Excel [SMILES] Ribbon Tab, 3-Level Right-Click Menu & .SDF Database Workflow'
    $r9Headers = @('Ribbon Group', 'Exact Button / Menu Label', 'Shortcut / Access', 'Primary Action', 'Output / Export Options', 'Tips & Cross-App Notes')
    $r9Widths  = @(18, 28, 22, 34, 32, 42)
    $r9Rows = @(
        @('Insert Structure', 'Insert / Edit (SMILES)',     'Ctrl+Shift+K / Double-Click', 'Opens EPAM Ketcher 2D editor',                    'Inserts =SMILES2IMG or =KET2IMG',          'Double-click any molecule cell to edit in 0.05s'),
        @('Insert Structure', 'Insert / Edit (InChI)',      'Ribbon -> [SMILES]',          'Opens Ketcher / inserts IUPAC InChI formula',     'Inserts =INCHI2IMG(...)',                  'Best for regulatory & PubChem InChI tables'),
        @('Insert Structure', 'Insert / Edit (KET)',        'Ribbon -> [SMILES]',          'Inserts lossless KET reaction / canvas formula',  'Inserts =KET2IMG(...)',                    'Preserves reaction arrows, +, and labels'),
        @('Insert Structure', '3D Structure (Mol*)',        'Ribbon -> [SMILES]',          'Opens Mol* 3D protein viewer (PDB / local file)', '3D Model (.glb) or Transparent 2D PNG',    'Try PDB IDs: 1CRN, 4HHB, 6LU7, 1BNA'),
        @('Insert Structure', 'Paste Structure',            'Ctrl+Shift+V',                'Converts clipboard SMILES/InChI/KET/Name in 0.1s','Instant structure without opening editor', 'Works in Excel, Word & PowerPoint'),
        @('Structure Tools',  'Copy SMILES / InChI / KET',  'Ribbon or Right-Click',       'Extracts exact identifier from selected cell',    'Plain text onto clipboard',                'Right-click -> Chemical Structure > ...'),
        @('Structure Tools',  'Copy SVG',                   'Ribbon or Right-Click',       'Copies lossless vector SVG graphics',             'Paste into Illustrator / Inkscape',        'Crisp at any poster/journal DPI'),
        @('Structure Tools',  'Molecule Info',              'Ribbon or Right-Click',       'Inspects 17 descriptors + 2D preview dialog',     'Summary, 17-col TSV, List, Markdown, JSON','Double-click any row to copy single value'),
        @('Structure Tools',  'Batch Style',                'Ribbon -> Batch Style',       '1-click restyle across selection or entire sheet','Color, ACS B&W, White BG, Dark Mode',      'Instant switch before journal submission'),
        @('Structure Tools',  'SDF File (Import / Export)', 'Ribbon -> SDF File',          'Imports .sdf into 12-col live formula table',     'Exports sheet rows back to V2000 .sdf',    'Export Tab 2 to .sdf, then re-import it!'),
        @('Diagnostics/About','Diagnostics & About',        'Ribbon -> [SMILES]',          'Inspects Indigo parser errors & checks updates',  'Logs, 3rd-party licenses, version check',  'Open Diagnostics if you see #VALUE!')
    )
    Add-StyledTable $s9 5 $r9Headers $r9Rows @{} 0x475569 $r9Widths 30 @(1,2,3,4,5,6) @()

    Add-SectionHeader $s9 17 6 'B. Complete 40-Function Excel UDF Catalog (Grouped by Category — Zero Web Docs Needed)'
    $f9Headers = @('Category', 'Function Name', 'Syntax', 'Return Type', 'Network?', 'Description')
    $f9Rows = @(
        @('In-Cell Image (4)',    'SMILES2IMG',        '=SMILES2IMG(smiles, [bg], [style], [transform])', 'In-Cell Picture', 'Offline',  'Render 2D structure from SMILES string'),
        @('In-Cell Image (4)',    'NAME2IMG',          '=NAME2IMG(nameOrCas, [bg], [style], [transform])','In-Cell Picture', 'PubChem*', 'Render 2D structure from common name or CAS (cached)'),
        @('In-Cell Image (4)',    'INCHI2IMG',         '=INCHI2IMG(inchi, [bg], [style], [transform])',   'In-Cell Picture', 'Offline',  'Render 2D structure from IUPAC InChI'),
        @('In-Cell Image (4)',    'KET2IMG',           '=KET2IMG(ketJson, [bg], [style], [transform])',   'In-Cell Picture', 'Offline',  'Render lossless Ketcher KET molecule or reaction scheme'),
        @('Legacy Float (4)',     'SMILES2IMG.FLOAT',  '=SMILES2IMG.FLOAT(smiles, [bg], [style], [tr])',  'Floating Shape',  'Offline',  'Excel 2016-2021 floating picture from SMILES'),
        @('Legacy Float (4)',     'NAME2IMG.FLOAT',    '=NAME2IMG.FLOAT(nameOrCas, [bg], [style], [tr])', 'Floating Shape',  'PubChem*', 'Excel 2016-2021 floating picture from Name/CAS'),
        @('Legacy Float (4)',     'INCHI2IMG.FLOAT',   '=INCHI2IMG.FLOAT(inchi, [bg], [style], [tr])',    'Floating Shape',  'Offline',  'Excel 2016-2021 floating picture from InChI'),
        @('Legacy Float (4)',     'KET2IMG.FLOAT',     '=KET2IMG.FLOAT(ketJson, [bg], [style], [tr])',    'Floating Shape',  'Offline',  'Excel 2016-2021 floating picture from KET'),
        @('Properties (9)',       'SMILES2FORMULA',    '=SMILES2FORMULA(smiles)',                         'Text (String)',   'Offline',  'Hill molecular formula (e.g. C9H8O4)'),
        @('Properties (9)',       'SMILES2MW',         '=SMILES2MW(smiles)',                              'Number (Double)', 'Offline',  'Average molecular weight (g/mol)'),
        @('Properties (9)',       'SMILES2EXACTMASS',  '=SMILES2EXACTMASS(smiles)',                       'Number (Double)', 'Offline',  'Monoisotopic exact mass for HRMS (Da)'),
        @('Properties (9)',       'SMILES2LOGP',       '=SMILES2LOGP(smiles)',                            'Number (Double)', 'Offline',  'Wildman-Crippen octanol-water LogP'),
        @('Properties (9)',       'SMILES2TPSA',       '=SMILES2TPSA(smiles)',                            'Number (Double)', 'Offline',  'Topological polar surface area (A^2)'),
        @('Properties (9)',       'SMILES2HBD',        '=SMILES2HBD(smiles)',                             'Number (Int)',    'Offline',  'Hydrogen bond donor count (NH + OH)'),
        @('Properties (9)',       'SMILES2HBA',        '=SMILES2HBA(smiles)',                             'Number (Int)',    'Offline',  'Hydrogen bond acceptor count (N + O)'),
        @('Properties (9)',       'SMILES2ROTBONDS',   '=SMILES2ROTBONDS(smiles)',                        'Number (Int)',    'Offline',  'Rotatable single bond count'),
        @('Properties (9)',       'SMILES2LIPINSKI',   '=SMILES2LIPINSKI(smiles)',                        'Text (String)',   'Offline',  'Lipinski Rule of 5 ("PASS" or violation details)'),
        @('3D Complexity (5)',    'SMILES2HEAVY',      '=SMILES2HEAVY(smiles)',                           'Number (Int)',    'Offline',  'Heavy (non-hydrogen) atom count'),
        @('3D Complexity (5)',    'SMILES2RINGS',      '=SMILES2RINGS(smiles)',                           'Number (Int)',    'Offline',  'Total ring count (SSSR)'),
        @('3D Complexity (5)',    'SMILES2AROMATIC',   '=SMILES2AROMATIC(smiles)',                        'Number (Int)',    'Offline',  'Aromatic ring count'),
        @('3D Complexity (5)',    'SMILES2FSP3',       '=SMILES2FSP3(smiles)',                            'Number (Double)', 'Offline',  'Fraction of sp3 carbons (0.00-1.00)'),
        @('3D Complexity (5)',    'SMILES2CHIRAL',     '=SMILES2CHIRAL(smiles)',                          'Number (Int)',    'Offline',  'Tetrahedral chiral center count'),
        @('SAR Search (2)',       'SMILES.MATCH',      '=SMILES.MATCH(targetSmiles, querySmarts)',        'Boolean',         'Offline',  'TRUE if target contains SMARTS/SMILES substructure'),
        @('SAR Search (2)',       'SMILES.SIMILARITY', '=SMILES.SIMILARITY(smilesA, smilesB)',            'Number (Double)', 'Offline',  'Morgan/ECFP Tanimoto similarity (0.00-1.00)'),
        @('Convert / Name (4)',   'NAME2SMILES',       '=NAME2SMILES(nameOrCas)',                         'Text (String)',   'PubChem*', 'Resolve chemical name or CAS to SMILES (cached)'),
        @('Convert / Name (4)',   'SMILES2NAME',       '=SMILES2NAME(smiles)',                            'Text (String)',   'PubChem*', 'Resolve SMILES to common substance name (cached)'),
        @('Convert / Name (4)',   'SMILES2INCHI',      '=SMILES2INCHI(smiles)',                           'Text (String)',   'Offline',  'Convert SMILES to IUPAC Standard InChI'),
        @('Convert / Name (4)',   'SMILES2INCHIKEY',   '=SMILES2INCHIKEY(smiles)',                        'Text (String)',   'Offline',  'Convert SMILES to 27-char InChIKey (deduplication)'),
        @('InChI Functions (6)',  'INCHI2SMILES',      '=INCHI2SMILES(inchi)',                            'Text (String)',   'Offline',  'Convert IUPAC InChI to SMILES'),
        @('InChI Functions (6)',  'INCHI2FORMULA',     '=INCHI2FORMULA(inchi)',                           'Text (String)',   'Offline',  'Hill formula directly from InChI'),
        @('InChI Functions (6)',  'INCHI2MW',          '=INCHI2MW(inchi)',                                'Number (Double)', 'Offline',  'Molecular weight directly from InChI'),
        @('InChI Functions (6)',  'INCHI2EXACTMASS',   '=INCHI2EXACTMASS(inchi)',                         'Number (Double)', 'Offline',  'Monoisotopic exact mass directly from InChI'),
        @('InChI Functions (6)',  'INCHI2INCHIKEY',    '=INCHI2INCHIKEY(inchi)',                          'Text (String)',   'Offline',  'InChIKey directly from InChI'),
        @('InChI Functions (6)',  'INCHI2NAME',        '=INCHI2NAME(inchi)',                              'Text (String)',   'PubChem*', 'Common substance name from InChI (cached)'),
        @('KET Functions (6)',    'KET2FORMULA',       '=KET2FORMULA(ketJson)',                           'Text (String)',   'Offline',  'Hill formula directly from KET JSON'),
        @('KET Functions (6)',    'KET2MW',            '=KET2MW(ketJson)',                                'Number (Double)', 'Offline',  'Molecular weight directly from KET JSON'),
        @('KET Functions (6)',    'KET2EXACTMASS',     '=KET2EXACTMASS(ketJson)',                         'Number (Double)', 'Offline',  'Monoisotopic exact mass directly from KET JSON'),
        @('KET Functions (6)',    'KET2INCHI',         '=KET2INCHI(ketJson)',                             'Text (String)',   'Offline',  'IUPAC InChI directly from KET JSON'),
        @('KET Functions (6)',    'KET2INCHIKEY',      '=KET2INCHIKEY(ketJson)',                          'Text (String)',   'Offline',  'InChIKey directly from KET JSON'),
        @('KET Functions (6)',    'KET2NAME',          '=KET2NAME(ketJson)',                              'Text (String)',   'PubChem*', 'Common substance name from KET JSON (cached)')
    )
    Add-StyledTable $s9 18 $f9Headers $f9Rows @{} 0x334155 @() 24 @(1,2,3,6) @(2,3)

    Add-SectionHeader $s9 60 6 'C. Built-In Troubleshooting Guide (#N/A, #VALUE!, #NAME? & Sharing Workbooks)'
    $trHeaders = @('Symptom / Code', 'Cause', 'Exact Fix', 'Notes', 'Applies To', 'Quick Check')
    $trRows = @(
        @('#N/A (~0.1s)',     'Async rendering or first-time PubChem lookup', 'Wait a moment — updates automatically',       'Cached locally for instant reuse',  'All Image/Name UDFs', 'Normal behavior'),
        @('#VALUE! (Error)',  'Invalid SMILES/InChI syntax or empty cell',    'Click Ribbon [SMILES] -> [Diagnostics]',      'Shows exact Indigo parser error',   'All UDFs',            'Check unclosed rings'),
        @('#NAME? (Missing)', 'Add-in not loaded or Excel 2016-2021',         'Re-run install.ps1 or append .FLOAT',         'See Tab 8 for .FLOAT formulas',     'Excel 2016-2021',     'File -> Options -> Add-ins'),
        @('Sharing .xlsx',  'Co-author on macOS or PC without add-in',      'Cached PNG images & values stay visible',     'Avoid Ctrl+Alt+F9 without add-in',  'All Excel versions',  'Export PDF for lock')
    )
    Add-StyledTable $s9 61 $trHeaders $trRows @{} 0x3B291E @() 28 @(1,2,3,4,5,6) @()

    Write-Host "`nRecalculating all 10 tabs and waiting for async image & .FLOAT rendering..." -ForegroundColor Cyan
    Invoke-Excel {
        $excel.Calculation = -4105 # xlCalculationAutomatic
        $s8.Activate()
        $excel.Calculate()
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $maxWaitSec = 75
    do {
        Start-Sleep -Seconds 2
        $floatShapes = 0
        $pendingErrors = 0
        Invoke-Excel {
            $floatShapes = $s8.Shapes.Count
            # Check key PubChem/async cells on Tabs 0, 1, 2, 4, 6, 7 for any lingering #N/A or errors
            foreach ($wsCheck in @($s0, $s1, $s2, $s3, $s4, $s5, $s6, $s7, $s8)) {
                $ur = $wsCheck.UsedRange
                $vals = $ur.Value2
                if ($vals -is [object[,]]) {
                    $rMax = $vals.GetLength(0); $cMax = $vals.GetLength(1)
                    for ($r = 1; $r -le $rMax; $r++) {
                        for ($c = 1; $c -le $cMax; $c++) {
                            $v = $vals[$r, $c]
                            if ($v -is [int] -and $v -lt -2146826000) { $pendingErrors++ }
                        }
                    }
                }
            }
        }
        Write-Host ("  Elapsed: {0,2}s | Tab 8 .FLOAT Shapes: {1}/10 | Pending Excel Errors: {2}" -f [int]$sw.Elapsed.TotalSeconds, $floatShapes, $pendingErrors) -ForegroundColor Gray
        if ($floatShapes -ge 10 -and $pendingErrors -eq 0 -and $sw.Elapsed.TotalSeconds -ge 12) {
            Start-Sleep -Seconds 3
            break
        }
    } while ($sw.Elapsed.TotalSeconds -lt $maxWaitSec)

    # Strict Zero-Error Audit across ALL 10 sheets before saving!
    Write-Host "`nRunning strict cell-by-cell zero-error audit across all 10 tabs..." -ForegroundColor Cyan
    $errorReport = New-Object 'System.Collections.Generic.List[string]'
    Invoke-Excel {
        foreach ($wsCheck in $sheets) {
            $ur = $wsCheck.UsedRange
            $r0 = $ur.Row; $c0 = $ur.Column
            $rCnt = $ur.Rows.Count; $cCnt = $ur.Columns.Count
            for ($r = 1; $r -le $rCnt; $r++) {
                for ($c = 1; $c -le $cCnt; $c++) {
                    $cell = $ur.Cells.Item($r, $c)
                    $frm = [string]$cell.Formula
                    # In-cell RichData image cells report #VALUE! via legacy COM; verified post-save via OpenXML vm= attribute
                    if ($frm -match '^=(SMILES2IMG|NAME2IMG|INCHI2IMG|KET2IMG)\(') { continue }
                    $txt = ([string]$cell.Text).Trim()
                    $val = $cell.Value2
                    $isErrVal = ($val -is [int] -and $val -lt -2146826000)
                    if ($isErrVal -or $txt -match '^#(REF!|NAME\?|VALUE!|N/A|DIV/0!|NUM!|NULL!)$') {
                        $addr = "$(Get-ColLetter ($c0 + $c - 1))$($r0 + $r - 1)"
                        $errorReport.Add("$($wsCheck.Name)!$addr = $txt (Formula: $frm)")
                    }
                }
            }
        }
    }
    if ($errorReport.Count -gt 0) {
        $errorReport | ForEach-Object { Write-Warning "  [ERR] $_" }
        throw "Workbook audit failed with $($errorReport.Count) Excel formula errors!"
    }
    Write-Host '[PASS] 0 formula errors across all 10 worksheets!' -ForegroundColor Green

    Write-Host "`nSaving Excel Playbook to: $targetFile" -ForegroundColor Cyan
    Invoke-Excel {
        $s0.Activate()
        $workbook.SaveAs($targetFile, 51)
    }
    Copy-Item -LiteralPath $targetFile -Destination $docsPublicFile -Force
    Write-Host "[OK] Synced copy to: $docsPublicFile" -ForegroundColor Green

} finally {
    if ($workbook) { try { $workbook.Close($false) } catch {} }
    if ($excel) {
        try {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# OpenXML Integrity Verification
$fileSizeKB = [math]::Round((Get-Item -LiteralPath $targetFile).Length / 1KB, 1)
$zip = [System.IO.Compression.ZipFile]::OpenRead($targetFile)
try {
    $entries = @($zip.Entries | Select-Object -ExpandProperty FullName)
    $sheetCount = @($entries | Where-Object { $_ -like 'xl/worksheets/sheet*.xml' }).Count
    $pngCount   = @($entries | Where-Object { $_ -like 'xl/media/*.png' }).Count
    $richCount  = @($entries | Where-Object { $_ -like 'xl/richData/*' }).Count
    $xmlErrors  = New-Object 'System.Collections.Generic.List[string]'
    foreach ($se in ($zip.Entries | Where-Object { $_.FullName -like 'xl/worksheets/sheet*.xml' })) {
        $sr = New-Object System.IO.StreamReader($se.Open())
        $sXml = $sr.ReadToEnd(); $sr.Dispose()
        foreach ($m in [regex]::Matches($sXml, '<c [^>]*t="e"[^>]*>.*?</c>')) {
            if ($m.Value -notmatch 'vm="') { $xmlErrors.Add("$($se.Name): $($m.Value)") }
        }
    }
    Write-Host "`n=== Excel Playbook Verification ===" -ForegroundColor Cyan
    Write-Host "  File Size : $fileSizeKB KB" -ForegroundColor Green
    Write-Host "  Sheets    : $sheetCount (10 visible + hidden richData helper)" -ForegroundColor Green
    Write-Host "  PNG Media : $pngCount embedded structure images" -ForegroundColor Green
    Write-Host "  RichData  : $richCount entries" -ForegroundColor Green
    Write-Host "  XML Errors: $($xmlErrors.Count) non-image error cells" -ForegroundColor $(if ($xmlErrors.Count -eq 0) { 'Green' } else { 'Red' })
    if ($xmlErrors.Count -gt 0) {
        $xmlErrors | ForEach-Object { Write-Warning "  [XML ERR] $_" }
        throw "OpenXML audit failed with $($xmlErrors.Count) non-image error cell(s)!"
    }
    if ($sheetCount -lt 10) { throw "Expected >= 10 sheets, found $sheetCount" }
    if ($pngCount -lt 30)   { throw "Expected >= 30 embedded PNG images, found $pngCount" }
} finally {
    $zip.Dispose()
}

Write-Host "`n[PASS] func_SMILES2IMG_Excel_Playbook.xlsx generated and verified!" -ForegroundColor Green
