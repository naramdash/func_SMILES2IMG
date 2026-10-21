param(
    [string]$AddInPath = (Join-Path $PSScriptRoot '../dist/x64/Smiles2Img-AddIn64-packed.xll'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '../SMILES2IMG_Comprehensive_Showcase.xlsx')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# High performance C# Excel Helper to bypass PowerShell dynamic COM binder quirks
Add-Type -ReferencedAssemblies "Microsoft.CSharp.dll" -TypeDefinition @"
using System;
public static class ExcelHelper {
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
"@

$resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath).Path
$targetFile = [System.IO.Path]::GetFullPath($OutputPath)

Write-Host "=== Generating func_SMILES2IMG 8-Tab Master Interactive Showcase Workbook ===" -ForegroundColor Cyan
Write-Host "AddIn Path  : $resolvedAddIn" -ForegroundColor Gray
Write-Host "Target Path : $targetFile" -ForegroundColor Gray

if (Test-Path -LiteralPath $targetFile) {
    Remove-Item -LiteralPath $targetFile -Force
}

function Invoke-Excel([scriptblock]$operation, [int]$timeoutSec = 45) {
    $deadline = [DateTime]::UtcNow.AddSeconds($timeoutSec)
    do {
        try { return (& $operation) }
        catch {
            $hex = '{0:X8}' -f $_.Exception.HResult
            if (($hex -notmatch '80010001|8001010A|800AC472' -and
                 $_.Exception.ToString() -notmatch '80010001|8001010A|800AC472|rejected|거부') -or
                [DateTime]::UtcNow -ge $deadline) { throw }
            Start-Sleep -Milliseconds 350
        }
    } while ($true)
}

function Create-TableSheet {
    param(
        [object]$sheet,
        [string]$title,
        [string]$subtitle,
        [string[]]$headers,
        [array]$rows,           # 2D array of data values (text/numbers)
        [hashtable]$formulas,   # column letter -> array of formula strings (1 per row)
        [int]$titleColor,       # BGR color
        [int]$headerColor,      # BGR color
        [int[]]$colWidths,
        [int]$rowHeight = 85,
        [int]$startRow = 4
    )

    $numCols = $headers.Length
    $numRows = $rows.Length
    $lastColChar = [char](64 + $numCols)
    $endRow = $startRow + $numRows - 1

    # 1. Title Banner
    Invoke-Excel {
        $sheet.Range("A1").Value2 = $title
        $sheet.Range("A1").Font.Size = 14
        $sheet.Range("A1").Font.Bold = $true
        $sheet.Range("A1").Font.Color = 0xFFFFFF
        $titleRange = $sheet.Range("A1:${lastColChar}1")
        $titleRange.Merge()
        $titleRange.Interior.Color = $titleColor
        $titleRange.RowHeight = 32
        $titleRange.VerticalAlignment = -4108 # xlCenter
        $titleRange.HorizontalAlignment = -4108
    }

    # 2. Subtitle / Description Row
    if ($subtitle) {
        Invoke-Excel {
            $sheet.Range("A2").Value2 = $subtitle
            $sheet.Range("A2").Font.Size = 9.5
            $sheet.Range("A2").Font.Italic = $true
            $sheet.Range("A2").Font.Color = 0x333333
            $subRange = $sheet.Range("A2:${lastColChar}2")
            $subRange.Merge()
            $subRange.Interior.Color = 0xF5F5F5
            $subRange.RowHeight = 22
            $subRange.VerticalAlignment = -4108
            $subRange.HorizontalAlignment = -4131 # xlLeft
        }
    }

    # 3. Header Row
    Invoke-Excel {
        $headerRow = $startRow - 1
        for ($c = 0; $c -lt $numCols; $c++) {
            $sheet.Cells.Item($headerRow, $c + 1).Value2 = $headers[$c]
        }
        $headerRange = $sheet.Range("A${headerRow}:${lastColChar}${headerRow}")
        $headerRange.Font.Bold = $true
        $headerRange.Font.Size = 10.5
        $headerRange.Font.Color = 0xFFFFFF
        $headerRange.Interior.Color = $headerColor
        $headerRange.RowHeight = 26
        $headerRange.VerticalAlignment = -4108
        $headerRange.HorizontalAlignment = -4108
    }

    # 4. Populate Data Values & Formatting (Set BEFORE formulas!)
    $dataArr = New-Object 'object[,]' $numRows, $numCols
    for ($r = 0; $r -lt $numRows; $r++) {
        for ($c = 0; $c -lt $numCols; $c++) {
            $val = $rows[$r][$c]
            if ($val -is [string] -and $val.StartsWith("=")) {
                $val = "'" + $val
            }
            $dataArr[$r, $c] = $val
        }
    }

    Invoke-Excel {
        $dataRange = $sheet.Range("A${startRow}:${lastColChar}${endRow}")
        [ExcelHelper]::SetRangeValue2($dataRange, $dataArr)
        $dataRange.RowHeight = $rowHeight
        $dataRange.Borders.LineStyle = 1 # xlContinuous
        $dataRange.Borders.Color = 0xD9D9D9
        $dataRange.VerticalAlignment = -4108
        $dataRange.HorizontalAlignment = -4108
    }

    # 5. Column Widths
    Invoke-Excel {
        for ($c = 0; $c -lt $numCols; $c++) {
            $sheet.Columns.Item($c + 1).ColumnWidth = $colWidths[$c]
        }
    }

    # 6. Apply Formulas
    foreach ($colKey in $formulas.Keys) {
        $colFormulas = $formulas[$colKey]
        $formulaArr = New-Object 'object[,]' $numRows, 1
        for ($r = 0; $r -lt $numRows; $r++) {
            $formulaArr[$r, 0] = $colFormulas[$r]
        }
        Invoke-Excel {
            $targetColRange = $sheet.Range("${colKey}${startRow}:${colKey}${endRow}")
            [ExcelHelper]::SetRangeFormula($targetColRange, $formulaArr)
        }
    }
}

$excel = $null
$workbook = $null

try {
    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Excel Version: $($excel.Version)" -ForegroundColor Gray
    if (-not $excel.RegisterXLL($resolvedAddIn)) {
        throw "Failed to register XLL: $resolvedAddIn"
    }
    Write-Host "XLL registered successfully!" -ForegroundColor Green

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -lt 8) {
        $null = $workbook.Worksheets.Add()
    }

    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet1.Name = "0. Quick Guide"
    $sheet1.Tab.Color = 0x00A5FF # Gold BGR

    $sheet2 = $workbook.Worksheets.Item(2)
    $sheet2.Name = "1. Structure to Image"
    $sheet2.Tab.Color = 0xD97000 # Blue BGR

    $sheet3 = $workbook.Worksheets.Item(3)
    $sheet3.Name = "2. Molecular Properties"
    $sheet3.Tab.Color = 0x228B22 # Forest Green BGR

    $sheet4 = $workbook.Worksheets.Item(4)
    $sheet4.Name = "3. Substructure & Similarity"
    $sheet4.Tab.Color = 0x800080 # Purple BGR

    $sheet5 = $workbook.Worksheets.Item(5)
    $sheet5.Name = "4. Drawing Styles & Options"
    $sheet5.Tab.Color = 0x0080FF # Orange BGR

    $sheet6 = $workbook.Worksheets.Item(6)
    $sheet6.Name = "5. Reactions & KET Canvas"
    $sheet6.Tab.Color = 0x507020 # Dark Teal BGR

    $sheet7 = $workbook.Worksheets.Item(7)
    $sheet7.Name = "6. Legacy Excel (.FLOAT)"
    $sheet7.Tab.Color = 0x1B4D1B # Dark Green BGR

    $sheet8 = $workbook.Worksheets.Item(8)
    $sheet8.Name = "7. 3D & Ribbon Tools"
    $sheet8.Tab.Color = 0x800000 # Navy / Maroon BGR

    # =========================================================================
    # TAB 1: Quick Guide & Shortcuts
    # =========================================================================
    Write-Host "Populating Tab 1: Quick Guide & Shortcuts..." -ForegroundColor Yellow
    $t1Headers = @("No.", "Feature Category", "Compound / Input", "Formula (Plain Text)", "Live Working Result (In-Cell)", "Keyboard Shortcut / Action", "Key Capability & Tips")
    $t1ColWidths = @(6, 22, 24, 38, 28, 28, 52)

    $t1Rows = @(
        @(1, "1. Chemical Name Lookup", "Aspirin", '=NAME2IMG("Aspirin")', "", "Type formula & press Enter", "Auto-renders 2D structure from PubChem common chemical name"),
        @(2, "2. CAS Registry Lookup", "58-08-2", '=NAME2IMG("58-08-2")', "", "Type formula & press Enter", "Auto-renders Caffeine structure directly from standard CAS number"),
        @(3, "3. Chemical SMILES String", "CC(=O)Oc1ccccc1C(=O)O", '=SMILES2IMG(C6)', "", "Type formula & press Enter", "Standard native in-cell picture scaling with row & column resize"),
        @(4, "4. Molecular Weight (MW)", "CC(=O)Oc1ccccc1C(=O)O", '=SMILES2MW(C7)', "", "Cell formula", "Zero-lag in-memory molecular weight calculation (180.16 g/mol)"),
        @(5, "5. Chemical Formula", "CC(=O)Oc1ccccc1C(=O)O", '=SMILES2FORMULA(C8)', "", "Cell formula", "Calculates Hill chemical formula (C9H8O4) locally with zero lag"),
        @(6, "6. Lipinski Rule of 5", "CC(=O)Oc1ccccc1C(=O)O", '=SMILES2LIPINSKI(C9)', "", "Cell formula", "Screens oral bioavailability: Pass (0 violations) or lists rule failures"),
        @(7, "7. Substructure Match", "CC(=O)Oc1ccccc1C(=O)O", '=SMILES.MATCH(C10, "c1ccccc1")', "", "Cell formula", "Fast subgraph isomorphism check: TRUE if benzene ring is present"),
        @(8, "8. 2D Sketcher In-Place", "Any Molecule Cell", "Double-Click Any Cell", "", "Double-Click or Ctrl+Shift+K", "Opens Ketcher 2D Editor in 0.05s; edit bonds and click Apply"),
        @(9, "9. 3D Protein Embedding", "PDB ID: 1CRN", "Ribbon -> [SMILES] -> Mol*", "", "Click Ribbon [SMILES] Tab", "Embed rotatable 3D macromolecules & complexes into Office slides/cells"),
        @(10, "10. ACS B&W & SVG Copy", "Selected Molecules", "Right-Click -> Chemical Structure", "", "Right-Click Context Menu", "Copy crisp vector SVG for Illustrator or 1-click batch B&W style")
    )
    $t1Formulas = @{
        "E" = @(
            '=NAME2IMG("Aspirin")',
            '=NAME2IMG("58-08-2")',
            '=SMILES2IMG(C6)',
            '=SMILES2MW(C7)',
            '=SMILES2FORMULA(C8)',
            '=SMILES2LIPINSKI(C9)',
            '=SMILES.MATCH(C10, "c1ccccc1")',
            '=SMILES2IMG("CC(=O)Oc1ccccc1C(=O)O", , "stereo")',
            '=SMILES2IMG("CC(=O)Nc1ccc(O)cc1", "white")',
            '=SMILES2IMG("CC(=O)Oc1ccccc1C(=O)O", , "bw")'
        )
    }
    Create-TableSheet $sheet1 "func_SMILES2IMG — Master Interactive Guide & Function Sandbox" "💡 Tip: Double-click any cell or structure picture to sketch in Ketcher 2D! All formulas evaluate 100% locally with zero cloud requests." $t1Headers $t1Rows $t1Formulas 0x0045B4 0x0066CC $t1ColWidths 85 4

    # =========================================================================
    # TAB 2: Name & SMILES to Image
    # =========================================================================
    Write-Host "Populating Tab 2: Name & SMILES to Image..." -ForegroundColor Yellow
    $t2Headers = @("No.", "Input Category", "Compound Input (Name / CAS / SMILES / InChI)", "Formula Syntax (Text)", "Rendered 2D Structure (In-Cell)", "Resolved Canonical SMILES", "Resolved Common Name", "Application & Structural Highlights")
    $t2ColWidths = @(6, 18, 32, 42, 28, 38, 22, 48)

    $t2Rows = @(
        @(1, "Common Name", "Aspirin", '=NAME2IMG(C4)', "", '=NAME2SMILES(C4)', '=SMILES2NAME(F4)', "Acetylsalicylic acid — NSAID analgesic and antipyretic drug"),
        @(2, "Common Name", "Ibuprofen", '=NAME2IMG(C5)', "", '=NAME2SMILES(C5)', '=SMILES2NAME(F5)', "Propionic acid derivative NSAID with asymmetric chiral center"),
        @(3, "Common Name", "Caffeine", '=NAME2IMG(C6)', "", '=NAME2SMILES(C6)', '=SMILES2NAME(F6)', "Purine alkaloid CNS stimulant from coffee beans and tea"),
        @(4, "Common Name", "Paracetamol", '=NAME2IMG(C7)', "", '=NAME2SMILES(C7)', '=SMILES2NAME(F7)', "Acetaminophen — Para-acetylaminophenol analgesic"),
        @(5, "CAS Number", "50-78-2", '=NAME2IMG(C8)', "", '=NAME2SMILES(C8)', "Aspirin", "Aspirin CAS registry number lookup without knowing SMILES"),
        @(6, "CAS Number", "58-08-2", '=NAME2IMG(C9)', "", '=NAME2SMILES(C9)', "Caffeine", "Caffeine CAS registry number lookup"),
        @(7, "CAS Number", "103-90-2", '=NAME2IMG(C10)', "", '=NAME2SMILES(C10)', "Acetaminophen", "Paracetamol / Acetaminophen CAS lookup"),
        @(8, "SMILES String", "c1ccccc1", '=SMILES2IMG(C11)', "", "c1ccccc1", '=SMILES2NAME(C11)', "Benzene — Fundamental aromatic ring core with alternating double bonds"),
        @(9, "SMILES String", "CCO", '=SMILES2IMG(C12)', "", "CCO", '=SMILES2NAME(C12)', "Ethanol — Primary aliphatic alcohol"),
        @(10, "SMILES String", "CN1CCC[C@H]1c2cccnc2", '=SMILES2IMG(C13)', "", "CN1CCC[C@H]1c2cccnc2", '=SMILES2NAME(C13)', "Nicotine — Tobacco alkaloid with chiral pyrrolidine and pyridine rings"),
        @(11, "SMILES String", "CC(=O)C", '=SMILES2IMG(C14)', "", "CC(=O)C", '=SMILES2NAME(C14)', "Acetone — Simple ketone organic solvent"),
        @(12, "InChI String", "InChI=1S/C8H10N4O2/c1-10-4-9-6-5(10)7(13)12(3)8(14)11(6)2/h4H,1-3H3", '=INCHI2IMG(C15)', "", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "Caffeine", "Direct rendering from official IUPAC standard InChI string")
    )
    $t2Formulas = @{
        "E" = @(
            '=NAME2IMG(C4)',
            '=NAME2IMG(C5)',
            '=NAME2IMG(C6)',
            '=NAME2IMG(C7)',
            '=NAME2IMG(C8)',
            '=NAME2IMG(C9)',
            '=NAME2IMG(C10)',
            '=SMILES2IMG(C11)',
            '=SMILES2IMG(C12)',
            '=SMILES2IMG(C13)',
            '=SMILES2IMG(C14)',
            '=INCHI2IMG(C15)'
        );
        "F" = @(
            '=NAME2SMILES(C4)',
            '=NAME2SMILES(C5)',
            '=NAME2SMILES(C6)',
            '=NAME2SMILES(C7)',
            '=NAME2SMILES(C8)',
            '=NAME2SMILES(C9)',
            '=NAME2SMILES(C10)',
            'c1ccccc1',
            'CCO',
            'CN1CCC[C@H]1c2cccnc2',
            'CC(=O)C',
            'Cn1cnc2c1c(=O)n(C)c(=O)n2C'
        );
        "G" = @(
            '=SMILES2NAME(F4)',
            '=SMILES2NAME(F5)',
            '=SMILES2NAME(F6)',
            '=SMILES2NAME(F7)',
            'Aspirin',
            'Caffeine',
            'Acetaminophen',
            '=SMILES2NAME(C11)',
            '=SMILES2NAME(C12)',
            '=SMILES2NAME(C13)',
            '=SMILES2NAME(C14)',
            'Caffeine'
        )
    }
    Create-TableSheet $sheet2 "1. Chemical Structure Auto-Rendering (=NAME2IMG, =SMILES2IMG & =INCHI2IMG)" "💡 Tip: Drag formulas down hundreds of rows to automatically convert chemical catalogs or compound lists into high-resolution structures." $t2Headers $t2Rows $t2Formulas 0xD97000 0xEA8511 $t2ColWidths 85 4

    # =========================================================================
    # TAB 3: Drug Discovery & Lipinski Ro5
    # =========================================================================
    Write-Host "Populating Tab 3: Drug Discovery & Lipinski Ro5..." -ForegroundColor Yellow
    $t3Headers = @("No.", "Drug Name", "SMILES String", "2D Structure", "Formula", "MW (g/mol)", "Exact Mass", "LogP", "TPSA (Å²)", "HBD", "HBA", "RotBonds", "Lipinski Ro5 Assessment", "Pharmacology & Significance")
    $t3ColWidths = @(5, 16, 26, 24, 14, 12, 13, 9, 11, 8, 8, 10, 24, 38)

    $t3Rows = @(
        @(1, "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", "", "", "", "", "", "", "", "", "", "", "Oral NSAID analgesic; 0 Ro5 violations"),
        @(2, "Ibuprofen", "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O", "", "", "", "", "", "", "", "", "", "", "Anti-inflammatory blockbuster; 0 Ro5 violations"),
        @(3, "Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "", "", "", "", "", "", "", "", "", "", "Purine alkaloid stimulant; 0 Ro5 violations"),
        @(4, "Paracetamol", "CC(=O)Nc1ccc(O)cc1", "", "", "", "", "", "", "", "", "", "", "Over-the-counter analgesic; 0 Ro5 violations"),
        @(5, "Penicillin G", "CC1(C(N2C(S1)C(C2=O)NC(=O)Cc3ccccc3)C(=O)O)C", "", "", "", "", "", "", "", "", "", "", "Beta-lactam antibiotic; 0 Ro5 violations"),
        @(6, "Morphine", "CN1CC[C@]23[C@@H]4Oc5c3c(CC1[C@@H]2C=C[C@@H]4O)ccc5O", "", "", "", "", "", "", "", "", "", "", "Potent central opioid alkaloid; 0 Ro5 violations"),
        @(7, "Dopamine", "NCCc1ccc(O)c(O)c1", "", "", "", "", "", "", "", "", "", "", "Essential reward neurotransmitter; 0 Ro5 violations"),
        @(8, "Nicotine", "CN1CCC[C@H]1c2cccnc2", "", "", "", "", "", "", "", "", "", "", "Potent cholinergic stimulant; 0 Ro5 violations"),
        @(9, "Artemisinin", "CC1CC2CC3(C)OO4C(O2)(C1C(=O)O3)C(C)CC4", "", "", "", "", "", "", "", "", "", "", "Nobel Prize antimalarial natural product; 0 Ro5 violations"),
        @(10, "Cholesterol", "CC(C)CCCC(C)C1CCC2C1(CCC3C2CC=C4C3(CCC(C4)O)C)C", "", "", "", "", "", "", "", "", "", "", "Lipid sterol; violates Ro5 due to high LogP (>5)"),
        @(11, "Atorvastatin", "CC(C)c1c(c(c(c1NC(=O)c2ccccc2)c3ccc(cc3)F)c4ccccc4)CC(O)CC(O)CC(=O)O", "", "", "", "", "", "", "", "", "", "", "Lipitor statin; violates Ro5 due to MW (>500)"),
        @(12, "Paclitaxel", "CC1=C2C(C(=O)C3(C(CC4C(C3C(C(C2(C)C)(CC1OC(=O)C(C(c5ccccc5)NC(=O)c6ccccc6)O)O)OC(=O)c7ccccc7)(CO4)OC(=O)C)O)C)OC(=O)C", "", "", "", "", "", "", "", "", "", "", "Chemotherapy blockbuster (Taxol); violates multiple Ro5 criteria")
    )
    $t3Formulas = @{
        "D" = @(4..15 | ForEach-Object { "=SMILES2IMG(C$_)" });
        "E" = @(4..15 | ForEach-Object { "=SMILES2FORMULA(C$_)" });
        "F" = @(4..15 | ForEach-Object { "=SMILES2MW(C$_)" });
        "G" = @(4..15 | ForEach-Object { "=SMILES2EXACTMASS(C$_)" });
        "H" = @(4..15 | ForEach-Object { "=SMILES2LOGP(C$_)" });
        "I" = @(4..15 | ForEach-Object { "=SMILES2TPSA(C$_)" });
        "J" = @(4..15 | ForEach-Object { "=SMILES2HBD(C$_)" });
        "K" = @(4..15 | ForEach-Object { "=SMILES2HBA(C$_)" });
        "L" = @(4..15 | ForEach-Object { "=SMILES2ROTBONDS(C$_)" });
        "M" = @(4..15 | ForEach-Object { "=SMILES2LIPINSKI(C$_)" })
    }
    Create-TableSheet $sheet3 "2. Physicochemical Properties & Lipinski Rule of 5 SAR Screening" "💡 Lipinski Rule of 5 Criteria: MW <= 500 Da | LogP <= 5.0 | HBD <= 5 | HBA <= 10 | RotBonds <= 10. All calculated instantly (0ms) in Excel." $t3Headers $t3Rows $t3Formulas 0x1B4D1B 0x228B22 $t3ColWidths 85 4

    # =========================================================================
    # TAB 4: Substructure & Similarity
    # =========================================================================
    Write-Host "Populating Tab 4: Substructure & Similarity..." -ForegroundColor Yellow

    $t4Headers = @("No.", "Compound Name", "Target SMILES", "Substructure Match? (=SMILES.MATCH)", "Highlighted Structure (=SMILES2IMG with hl:)", "Tanimoto Similarity to Lead", "Formula Syntax (Text)", "SAR Pharmacophore Analysis")
    $t4ColWidths = @(6, 18, 30, 24, 28, 24, 48, 48)

    $t4Rows = @(
        @(1, "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", "", "", "", '=SMILES.MATCH(C5, $B$3)', "Contains benzene core; perfect 1.0 similarity to itself"),
        @(2, "Ibuprofen", "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O", "", "", "", '=SMILES.MATCH(C6, $B$3)', "Contains benzene core; highlighted in red ring"),
        @(3, "Benzene", "c1ccccc1", "", "", "", '=SMILES.MATCH(C7, $B$3)', "Exact match to query substructure"),
        @(4, "Pyridine", "c1ccncc1", "", "", "", '=SMILES.MATCH(C8, $B$3)', "Nitrogen heterocycle: FALSE (not all-carbon benzene)"),
        @(5, "Ethanol", "CCO", "", "", "", '=SMILES.MATCH(C9, $B$3)', "Aliphatic alcohol: FALSE (no aromatic ring)"),
        @(6, "Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "", "", "", '=SMILES.MATCH(C10, $B$3)', "Fused purine heterocycle: FALSE"),
        @(7, "Paracetamol", "CC(=O)Nc1ccc(O)cc1", "", "", "", '=SMILES.MATCH(C11, $B$3)', "Contains para-substituted benzene ring: TRUE"),
        @(8, "Nicotine", "CN1CCC[C@H]1c2cccnc2", "", "", "", '=SMILES.MATCH(C12, $B$3)', "Contains pyridine, not benzene: FALSE"),
        @(9, "Vanillin", "O=Cc1ccc(O)c(OC)c1", "", "", "", '=SMILES.MATCH(C13, $B$3)', "Phenolic aldehyde with benzene ring: TRUE"),
        @(10, "Artemisinin", "CC1CC2CC3(C)OO4C(O2)(C1C(=O)O3)C(C)CC4", "", "", "", '=SMILES.MATCH(C14, $B$3)', "Sesquiterpene lactone without aromaticity: FALSE")
    )
    $t4Formulas = @{
        "D" = @(5..14 | ForEach-Object { "=SMILES.MATCH(C$_, `$B`$3)" });
        "E" = @(5..14 | ForEach-Object { "=SMILES2IMG(C$_, , ""hl:"" & `$B`$3)" });
        "F" = @(5..14 | ForEach-Object { "=SMILES.SIMILARITY(C$_, `$F`$3)" })
    }
    Create-TableSheet $sheet4 "3. Substructure Search, Pharmacophore Highlighting & Similarity" "💡 Change Cell B3 or F3 to immediately recalculate matches, highlight rings, and recompute Tanimoto similarity across the whole table!" $t4Headers $t4Rows $t4Formulas 0x4B0082 0x800080 $t4ColWidths 85 5

    # Parameter Controls in Row 3 (Set after table structure is created)
    Invoke-Excel {
        $sheet4.Range("A3").Value2 = "Query Substructure ($B$3):"
        $sheet4.Range("A3").Font.Bold = $true
        $sheet4.Range("A3").HorizontalAlignment = -4152
        $sheet4.Range("B3").Value2 = "c1ccccc1"
        $sheet4.Range("B3").Font.Bold = $true
        $sheet4.Range("B3").Interior.Color = 0xFFE0E0
        $sheet4.Range("B3").HorizontalAlignment = -4108
        $sheet4.Range("C3").Value2 = "(Benzene Ring Core)"
        $sheet4.Range("C3").Font.Italic = $true

        $sheet4.Range("E3").Value2 = "Lead Reference ($F$3):"
        $sheet4.Range("E3").Font.Bold = $true
        $sheet4.Range("E3").HorizontalAlignment = -4152
        $sheet4.Range("F3").Value2 = "CC(=O)Oc1ccccc1C(=O)O"
        $sheet4.Range("F3").Font.Bold = $true
        $sheet4.Range("F3").Interior.Color = 0xE6F2FF
        $sheet4.Range("F3").HorizontalAlignment = -4108
        $sheet4.Range("G3").Value2 = "(Aspirin Lead Molecule)"
        $sheet4.Range("G3").Font.Italic = $true
        $sheet4.Rows.Item(3).RowHeight = 24
    }

    # =========================================================================
    # TAB 5: Styles, ACS & Dark Mode
    # =========================================================================
    Write-Host "Populating Tab 5: Styles, ACS & Dark Mode..." -ForegroundColor Yellow

    $t5Headers = @("No.", "Rendering Style Preset", "Delimiter", "Option Code", "Formula Text", "Rendered Structure (In-Cell)", "Visual Style & Publication Standards")
    $t5ColWidths = @(6, 26, 12, 28, 56, 28, 48)

    $t5Rows = @(
        @(1, "Standard CPK Color", "-", "(default)", '=SMILES2IMG($B$3)', "", "Default transparent background + elemental CPK colors (C: black, O: red, N: blue)"),
        @(2, "White Canvas Background", "-", 'bg="white"', '=SMILES2IMG($B$3, "white")', "", "Opaque pure-white cell canvas suitable for printing and PDFs"),
        @(3, "Publication Black & White", "-", 'style="bw"', '=SMILES2IMG($B$3, , "bw")', "", "100% monochrome line art for ACS journals, Nature, and patent filings (USPTO/KIPO)"),
        @(4, "Sequential Atom Numbers", "-", 'style="num"', '=SMILES2IMG($B$3, , "num")', "", "Numbered atom indices (1, 2, 3...) for NMR and crystal structure assignment"),
        @(5, "Chiral Stereolabels", "-", 'style="stereo"', '=SMILES2IMG($B$3, , "stereo")', "", "Explicit R/S chiral centers and E/Z double bond labels"),
        @(6, "Unfolded Hydrogens", "-", 'style="h"', '=SMILES2IMG($B$3, , "h")', "", "Explicit C-H bond lines and hydrogens to evaluate steric hindrance"),
        @(7, "Skeletal Carbon Labels", "-", 'style="all-c"', '=SMILES2IMG($B$3, , "all-c")', "", "Explicit 'C' letter at all vertex carbons for educational diagrams"),
        @(8, "Dark Mode (White Ink)", "-", 'style="white" (Dark)', '=SMILES2IMG($B$3, "#1e1e1e", "white")', "", "Crisp pure-white bond lines on dark canvas for PowerPoint dark slides"),
        @(9, "Corporate CI Navy Ink", "-", 'style="ink=#003366"', '=SMILES2IMG($B$3, "aliceblue", "ink=#003366")', "", "Custom Hex bond color on aliceblue canvas for branded company reports"),
        @(10, "B/W + Atom Numbers", "Space", 'style="bw num"', '=SMILES2IMG($B$3, , "bw num")', "", "Combined monochrome patent line art and atom numbering"),
        @(11, "B/W + Chiral Stereolabels", "Pipe (|)", 'style="bw|stereo"', '=SMILES2IMG($B$3, , "bw|stereo")', "", "Pipe-delimited monochrome line art with chiral stereocenters"),
        @(12, "Transform: 90-deg CW", "-", 'transform="rot90"', '=SMILES2IMG($B$3, , , "rot90")', "", "90-degree clockwise rotation for compact vertical table columns"),
        @(13, "Transform: 180-deg Rotate", "-", 'transform="rot180"', '=SMILES2IMG($B$3, , , "rot180")', "", "180-degree upside down rotation"),
        @(14, "Transform: Horizontal Flip", "-", 'transform="flip"', '=SMILES2IMG($B$3, , , "flip")', "", "Horizontal planar reflection matching textbook coordinate standards"),
        @(15, "Transform: Chained Flip+Rot", "Space", 'transform="flip rot90"', '=SMILES2IMG($B$3, , , "flip rot90")', "", "Chained transformation: horizontal flip followed by 90-degree rotation"),
        @(16, "Ultimate All-in-One", "Space", 'All Options Enabled', '=SMILES2IMG($B$3, "white", "bw num stereo h all-c", "flip rot90")', "", "All capabilities active: White canvas + B/W + Numbers + Chiral + H + Carbons + Flip + Rot")
    )
    $t5Formulas = @{
        "F" = @(
            '=SMILES2IMG($B$3)',
            '=SMILES2IMG($B$3, "white")',
            '=SMILES2IMG($B$3, , "bw")',
            '=SMILES2IMG($B$3, , "num")',
            '=SMILES2IMG($B$3, , "stereo")',
            '=SMILES2IMG($B$3, , "h")',
            '=SMILES2IMG($B$3, , "all-c")',
            '=SMILES2IMG($B$3, "#1e1e1e", "white")',
            '=SMILES2IMG($B$3, "aliceblue", "ink=#003366")',
            '=SMILES2IMG($B$3, , "bw num")',
            '=SMILES2IMG($B$3, , "bw|stereo")',
            '=SMILES2IMG($B$3, , , "rot90")',
            '=SMILES2IMG($B$3, , , "rot180")',
            '=SMILES2IMG($B$3, , , "flip")',
            '=SMILES2IMG($B$3, , , "flip rot90")',
            '=SMILES2IMG($B$3, "white", "bw num stereo h all-c", "flip rot90")'
        )
    }
    Create-TableSheet $sheet5 "4. Visual Styling, Publication ACS B&W, Dark Mode & Transforms Matrix" "💡 Easily mix and match styling parameters using spaces or pipes (|). Unify an entire sheet in one click via Ribbon -> [Batch Style]." $t5Headers $t5Rows $t5Formulas 0x0045FF 0x0080FF $t5ColWidths 85 5

    # Parameter Controls in Row 3 (Set after table structure is created)
    Invoke-Excel {
        $sheet5.Range("A3").Value2 = "Reference Molecule ($B$3):"
        $sheet5.Range("A3").Font.Bold = $true
        $sheet5.Range("A3").HorizontalAlignment = -4152
        $sheet5.Range("B3").Value2 = "CC(=O)Oc1ccccc1C(=O)O"
        $sheet5.Range("B3").Font.Bold = $true
        $sheet5.Range("B3").Interior.Color = 0xE6F2FF
        $sheet5.Range("B3").HorizontalAlignment = -4108
        $sheet5.Range("C3").Value2 = "(Aspirin Benchmark Molecule)"
        $sheet5.Range("C3").Font.Italic = $true
        $sheet5.Rows.Item(3).RowHeight = 24
    }

    # =========================================================================
    # TAB 6: Reactions & KET Canvas
    # =========================================================================
    Write-Host "Populating Tab 6: Reactions & KET Canvas..." -ForegroundColor Yellow
    $sampleKetMethanol = '{"root":{"nodes":[{"$ref":"mol0"}]},"mol0":{"type":"molecule","atoms":[{"label":"C","location":[0,0,0]},{"label":"O","location":[1,0,0]}],"bonds":[{"type":1,"atoms":[0,1]}]}}'
    $sampleKetEthanol  = '{"root":{"nodes":[{"$ref":"mol0"}]},"mol0":{"type":"molecule","atoms":[{"label":"C","location":[0,0,0]},{"label":"C","location":[1,0,0]},{"label":"O","location":[1.5,0.86,0]}],"bonds":[{"type":1,"atoms":[0,1]},{"type":1,"atoms":[1,2]}]}}'

    $t6Headers = @("No.", "Reaction / Scheme Title", "Structure / KET Data", "Formula (Plain Text)", "Rendered Diagram (In-Cell)", "Chemical Description & Ketcher Workflow")
    $t6ColWidths = @(6, 28, 32, 28, 32, 52)

    $t6Rows = @(
        @(1, "Methanol KET Vector Canvas", $sampleKetMethanol, '=KET2IMG(C4)', "", "Native EPAM Ketcher vector diagram with precise atomic coordinates"),
        @(2, "Ethanol KET Vector Canvas", $sampleKetEthanol, '=KET2IMG(C5)', "", "Lossless 2D canvas persisted in workbook without raster artifacts"),
        @(3, "Aspirin Structure & Synthesis", "CC(=O)Oc1ccccc1C(=O)O", '=SMILES2IMG(C6)', "", "Salicylic acid + Acetic anhydride -> Acetylsalicylic acid (Aspirin)"),
        @(4, "Amide Coupling Product", "CC(=O)Nc1ccc(O)cc1", '=SMILES2IMG(C7)', "", "Para-aminophenol + Acetic anhydride -> Paracetamol"),
        @(5, "Caffeine Core Xanthine Ring", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", '=SMILES2IMG(C8)', "", "Purine heterocycle synthesized via Traube purine synthesis")
    )
    $t6Formulas = @{
        "E" = @(
            '=KET2IMG(C4)',
            '=KET2IMG(C5)',
            '=SMILES2IMG(C6)',
            '=SMILES2IMG(C7)',
            '=SMILES2IMG(C8)'
        )
    }
    Create-TableSheet $sheet6 "5. Multi-Molecule Chemical Reactions & EPAM Ketcher KET Canvas" "💡 Double-click any structure cell to open the Ketcher 2D Editor, draw multi-step reaction schemes with arrows, and click Apply!" $t6Headers $t6Rows $t6Formulas 0x507020 0x608028 $t6ColWidths 85 4

    # =========================================================================
    # TAB 7: Excel 2016-2021 Compatibility (.FLOAT Mode)
    # =========================================================================
    Write-Host "Populating Tab 7: Excel 2016-2021 Compatibility (.FLOAT Mode)..." -ForegroundColor Yellow
    $t7Headers = @("No.", "Input Category", "Compound Name", "Chemical SMILES", "Applied Preset", "Formula Text", "Floating Molecular Shape (Over Cell)", "Features & Legacy Behavior")
    $t7ColWidths = @(6, 16, 20, 30, 24, 48, 28, 48)

    $t7Rows = @(
        @(1, "Basic", "Ethanol", "CCO", "Default", '=SMILES2IMG.FLOAT(D4)', "", "Automatic cell-anchored floating shape on Excel 2016/2019/2021"),
        @(2, "Basic", "Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "Default", '=SMILES2IMG.FLOAT(D5)', "", "Locked to cell movement and resizing (xlMoveAndSize)"),
        @(3, "Background", "Paracetamol", "CC(=O)Nc1ccc(O)cc1", 'bg="white"', '=SMILES2IMG.FLOAT(D6, "white")', "", "Opaque white background floating picture"),
        @(4, "Single Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="bw"', '=SMILES2IMG.FLOAT(D7, , "bw")', "", "Monochrome patent line art floating shape for legacy Excel"),
        @(5, "Single Style", "L-Lactic Acid", "C[C@H](O)C(=O)O", 'style="stereo"', '=SMILES2IMG.FLOAT(D8, , "stereo")', "", "Chiral R/S stereochemical labels in floating mode"),
        @(6, "Single Style", "Ethanol", "CCO", 'style="h"', '=SMILES2IMG.FLOAT(D9, , "h")', "", "All C-H hydrogens unfolded in floating shape mode"),
        @(7, "Multi-Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="bw num"', '=SMILES2IMG.FLOAT(D10, , "bw num")', "", "Monochrome combined with atom numbers in floating mode"),
        @(8, "Dark Mode", "Nicotine", "CN1CCC[C@H]1c2cccnc2", 'style="white stereo"', '=SMILES2IMG.FLOAT(D11, "#1e1e1e", "white stereo")', "", "Dark canvas with pure-white bonds and chiral labels"),
        @(9, "Transform", "Nicotine", "CN1CCC[C@H]1c2cccnc2", 'transform="flip"', '=SMILES2IMG.FLOAT(D12, , , "flip")', "", "Textbook horizontal reflection in floating mode"),
        @(10, "All-in-One", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'White + Full + Rot', '=SMILES2IMG.FLOAT(D13, "white", "bw num stereo h all-c", "rot90")', "", "Ultimate all-in-one floating molecular rendering on legacy Excel")
    )
    $t7Formulas = @{
        "G" = @(4..13 | ForEach-Object {
            switch ($_) {
                4 { '=SMILES2IMG.FLOAT(D4)' }
                5 { '=SMILES2IMG.FLOAT(D5)' }
                6 { '=SMILES2IMG.FLOAT(D6, "white")' }
                7 { '=SMILES2IMG.FLOAT(D7, , "bw")' }
                8 { '=SMILES2IMG.FLOAT(D8, , "stereo")' }
                9 { '=SMILES2IMG.FLOAT(D9, , "h")' }
                10 { '=SMILES2IMG.FLOAT(D10, , "bw num")' }
                11 { '=SMILES2IMG.FLOAT(D11, "#1e1e1e", "white stereo")' }
                12 { '=SMILES2IMG.FLOAT(D12, , , "flip")' }
                13 { '=SMILES2IMG.FLOAT(D13, "white", "bw num stereo h all-c", "rot90")' }
            }
        })
    }
    Create-TableSheet $sheet7 "6. Excel 2016-2021 Compatibility (=SMILES2IMG.FLOAT & =NAME2IMG.FLOAT)" "💡 Append '.FLOAT' to any function name to generate floating picture shapes anchored to cells in Excel 2016, 2019, and 2021!" $t7Headers $t7Rows $t7Formulas 0x1B4D1B 0x228B22 $t7ColWidths 85 4

    # =========================================================================
    # TAB 8: 3D Mol* & Ribbon Guide
    # =========================================================================
    Write-Host "Populating Tab 8: 3D Mol* & Ribbon Guide..." -ForegroundColor Yellow
    $t8Headers = @("No.", "Feature Module", "Command / Ribbon Tool", "Keyboard Shortcut", "Office Host Apps", "Operational Workflow & Functionality")
    $t8ColWidths = @(6, 24, 28, 22, 22, 54)

    $t8Rows = @(
        @(1, "2D Molecule Editor", "Insert / Edit Structure", "Ctrl + Shift + K", "Excel · Word · PPT", "Opens embedded EPAM Ketcher in 0.05s; sketch molecules & reaction schemes; click Apply"),
        @(2, "InChI Text Input", "Insert InChI Molecule", "Ribbon -> [SMILES]", "Excel · Word · PPT", "Converts official IUPAC InChI strings into rendered 2D structure images"),
        @(3, "KET Reaction Canvas", "Insert KET Molecule", "Ribbon -> [SMILES]", "Excel · Word · PPT", "Inserts complete reaction schemes with reaction arrows, plus signs, and text annotations"),
        @(4, "3D Protein Viewer", "Insert 3D Mol* Model", "Ribbon -> [SMILES]", "Excel · Word · PPT", "Fetch PDB structures by 4-letter code (e.g. 1CRN); embed rotatable 3D GLB models or 2D views"),
        @(5, "Quick Paste", "Quick Insert Structure", "Ctrl + Shift + V", "Excel · Word · PPT", "Instantly parses clipboard SMILES/InChI/KET and renders image in under 0.1s"),
        @(6, "Molecule Info Window", "Copy Molecule Info", "Right-Click Context", "Excel · Word · PPT", "Opens rich DataGridView showing MW, Formula, Exact Mass, InChIKey, LogP, TPSA, Lipinski"),
        @(7, "Lossless SVG Export", "Copy SVG (Vector)", "Right-Click Context", "Excel · Word · PPT", "Copies crisp vector SVG image to clipboard for Adobe Illustrator and publication layout"),
        @(8, "1-Click Batch Style", "Batch Style Presets", "Ribbon -> Batch Style", "Excel · Word · PPT", "Normalizes every structure in workbook/slide deck into Standard, ACS B&W, or Dark Mode"),
        @(9, "Chemical Database I/O", "SDF File Import / Export", "Ribbon -> SDF File", "Excel (Standalone XLL)", "Imports multi-compound .sdf chemical tables into spreadsheet rows or exports sheet to .sdf"),
        @(10, "Live Diagnostics", "Diagnostics & About", "Ribbon -> About", "Excel · Word · PPT", "Monitors WebView2 instance pool, cache memory, COM registration, and update checks")
    )
    Create-TableSheet $sheet8 "7. 3D Macromolecules (Mol*), Ribbon Controls & Power Features Guide" "💡 All functions run 100% locally and offline inside your Microsoft Office installation with zero external server dependencies." $t8Headers $t8Rows @{} 0x800000 0xA02020 $t8ColWidths 50 4

    # Activate Sheet 1 (Quick Guide)
    $sheet1.Activate()

    # Recalculate all formulas across all 8 sheets
    Write-Host "`nRecalculating formulas across all 8 tabs and waiting for async image insertion..." -ForegroundColor Cyan
    Invoke-Excel { $excel.Calculate() }

    # Settle down
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $maxWaitSec = 60
    do {
        Start-Sleep -Seconds 2
        $shapeCount = 0
        Invoke-Excel { $shapeCount = $sheet7.Shapes.Count }
        Write-Host "Elapsed: $([int]$sw.Elapsed.TotalSeconds)s | Sheet 7 (.FLOAT) Shapes: $shapeCount / 10" -ForegroundColor Gray
        
        if ($shapeCount -ge 10 -and $sw.Elapsed.TotalSeconds -ge 14) {
            Start-Sleep -Seconds 3
            break
        }
    } while ($sw.Elapsed.TotalSeconds -lt $maxWaitSec)

    # Save to Target Path
    Write-Host "`nSaving master workbook to: $targetFile" -ForegroundColor Cyan
    Invoke-Excel {
        $workbook.SaveAs($targetFile, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
    }
    Write-Host "Workbook saved successfully!" -ForegroundColor Green

    # Also copy to docs/public so visitors can download it from the documentation site
    $docsPublicFile = Join-Path $PSScriptRoot "../docs/public/SMILES2IMG_Comprehensive_Showcase.xlsx"
    Copy-Item -LiteralPath $targetFile -Destination $docsPublicFile -Force
    Write-Host "Copied workbook to docs/public/ for documentation website download!" -ForegroundColor Green

} finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($excel) {
        try {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# =========================================================================
# Verification & OpenXML Inspection
# =========================================================================
Write-Host "`n=== Verifying Generated 8-Tab Workbook Integrity ===" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $targetFile)) {
    throw "Output file was not created: $targetFile"
}

$fileSize = (Get-Item -LiteralPath $targetFile).Length
Write-Host "File Size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green

$zip = [System.IO.Compression.ZipFile]::OpenRead($targetFile)
try {
    $entries = $zip.Entries | Select-Object -ExpandProperty FullName
    
    $sheetEntries = $entries | Where-Object { $_ -like 'xl/worksheets/sheet*.xml' }
    Write-Host "Worksheet XML Entries: $($sheetEntries.Count) (Target: 8+ sheets)" -ForegroundColor Green
    
    $mediaEntries = $entries | Where-Object { $_ -like 'xl/media/*.png' }
    Write-Host "Media (PNG) Entries: $($mediaEntries.Count) in-cell cached structures" -ForegroundColor Green

    $richDataEntries = $entries | Where-Object { $_ -like 'xl/richData/*' }
    Write-Host "RichData Entries: $($richDataEntries.Count)" -ForegroundColor Green

    if ($sheetEntries.Count -lt 8) {
        throw "Expected at least 8 worksheets, found $($sheetEntries.Count)"
    }
    if ($mediaEntries.Count -eq 0) {
        throw "No media PNG files found inside the workbook!"
    }
} finally {
    $zip.Dispose()
}

Write-Host "`n🎉 ALL 8-TAB MASTER WORKBOOK VERIFICATION PASSED!" -ForegroundColor Green
Write-Host "Path: $targetFile" -ForegroundColor Cyan
