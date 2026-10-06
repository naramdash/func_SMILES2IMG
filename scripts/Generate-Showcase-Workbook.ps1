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

Write-Host "=== Generating SMILES2IMG Comprehensive Showcase Workbook ===" -ForegroundColor Cyan
Write-Host "AddIn Path  : $resolvedAddIn" -ForegroundColor Gray
Write-Host "Target Path : $targetFile" -ForegroundColor Gray

if (Test-Path -LiteralPath $targetFile) {
    Remove-Item -LiteralPath $targetFile -Force
}

function Invoke-Excel([scriptblock]$operation) {
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    do {
        try { return (& $operation) }
        catch {
            if ($_.Exception.ToString() -notmatch '80010001|8001010A|800AC472|rejected by callee' -or
                [DateTime]::UtcNow -ge $deadline) { throw }
            Start-Sleep -Milliseconds 250
        }
    } while ($true)
}

function Populate-SheetData {
    param(
        [object]$sheet,
        [string]$title,
        [string[]]$headers,
        [array]$rows,
        [int]$bgrTitleColor,
        [int]$bgrHeaderColor,
        [int[]]$colWidths,
        [bool]$isFormulaCol7 = $true
    )

    $numCols = $headers.Length
    $numRows = $rows.Length
    $lastColChar = [char](64 + $numCols)
    $endRow = 3 + $numRows

    # 1. Title Row
    Invoke-Excel {
        $sheet.Range("A1").Value2 = $title
        $sheet.Range("A1").Font.Size = 15
        $sheet.Range("A1").Font.Bold = $true
        $sheet.Range("A1").Font.Color = 0xFFFFFF
        $titleRange = $sheet.Range("A1:${lastColChar}1")
        $titleRange.Merge()
        $titleRange.Interior.Color = $bgrTitleColor
        $titleRange.RowHeight = 36
        $titleRange.VerticalAlignment = -4108 # xlCenter
        $titleRange.HorizontalAlignment = -4108
    }

    # 2. Header Row
    Invoke-Excel {
        for ($c = 0; $c -lt $numCols; $c++) {
            $sheet.Cells.Item(3, $c + 1).Value2 = $headers[$c]
        }
        $headerRange = $sheet.Range("A3:${lastColChar}3")
        $headerRange.Font.Bold = $true
        $headerRange.Font.Size = 11
        $headerRange.Font.Color = 0xFFFFFF
        $headerRange.Interior.Color = $bgrHeaderColor
        $headerRange.RowHeight = 26
        $headerRange.VerticalAlignment = -4108
        $headerRange.HorizontalAlignment = -4108
    }

    # 3. Bulk Data Arrays
    $dataArr = New-Object 'object[,]' $numRows, $numCols
    $formulaArr = New-Object 'object[,]' $numRows, 1
    $formulaTextColIdx = if ($isFormulaCol7) { 5 } else { 4 }

    for ($r = 0; $r -lt $numRows; $r++) {
        $item = $rows[$r]
        for ($c = 0; $c -lt $numCols; $c++) {
            $val = $item[$c]
            if ($c -eq $formulaTextColIdx -and $val -is [string] -and $val.StartsWith("=")) {
                $val = "'" + $val
            }
            $dataArr[$r, $c] = $val
        }
        $formulaArr[$r, 0] = $item[$numCols]
    }

    # 4. Fast C# Bulk Assignment
    Invoke-Excel {
        $formulaTextColChar = if ($isFormulaCol7) { "F" } else { "E" }
        $sheet.Range("${formulaTextColChar}4:${formulaTextColChar}${endRow}").NumberFormat = "@"

        $dataRange = $sheet.Range("A4:${lastColChar}${endRow}")
        [ExcelHelper]::SetRangeValue2($dataRange, $dataArr)
        $dataRange.RowHeight = 85
        $dataRange.Borders.LineStyle = 1 # xlContinuous
        $dataRange.Borders.Color = 0xD9D9D9

        # Column widths
        for ($c = 0; $c -lt $numCols; $c++) {
            $sheet.Columns.Item($c + 1).ColumnWidth = $colWidths[$c]
        }

        # Apply Formulas
        $formulaColChar = if ($isFormulaCol7) { "G" } else { "F" }
        $formulaRange = $sheet.Range("${formulaColChar}4:${formulaColChar}${endRow}")
        [ExcelHelper]::SetRangeFormula($formulaRange, $formulaArr)
        $formulaRange.HorizontalAlignment = -4108
        $formulaRange.VerticalAlignment = -4108

        # Center Align specific columns
        $sheet.Range("A4:A${endRow}").HorizontalAlignment = -4108
        $sheet.Range("B4:B${endRow}").HorizontalAlignment = -4108
    }
}

$excel = $null
$workbook = $null

try {
    # Ensure clean Excel state
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
    while ($workbook.Worksheets.Count -lt 4) {
        $null = $workbook.Worksheets.Add()
    }

    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet1.Name = "SMILES2IMG (In-Cell 365)"

    $sheet2 = $workbook.Worksheets.Item(2)
    $sheet2.Name = "SMILES2IMG.FLOAT (Excel 2016+)"

    $sheet3 = $workbook.Worksheets.Item(3)
    $sheet3.Name = "Complex Natural Products"

    $sheet4 = $workbook.Worksheets.Item(4)
    $sheet4.Name = "Style & Transform Matrix"

    # =========================================================================
    # SHEET 1: SMILES2IMG (In-Cell 365)
    # =========================================================================
    Write-Host "Populating Sheet 1: SMILES2IMG (In-Cell 365)..." -ForegroundColor Yellow
    $sheet1.Tab.Color = 0xD97000 # Blue BGR

    $s1Headers = @("No.", "Category", "Compound Name", "Chemical SMILES", "Applied Option", "Formula (Text)", "Molecular Structure (In-Cell)", "Description & Notes")
    $s1ColWidths = @(6, 16, 22, 32, 26, 48, 28, 48)

    $s1Rows = @(
        @(1, "Basic", "Ethanol", "CCO", "Default (Transparent CPK)", '=SMILES2IMG(D4)', "", "Default transparent canvas + elemental CPK coloring", '=SMILES2IMG(D4)'),
        @(2, "Basic", "Benzene", "c1ccccc1", "Default", '=SMILES2IMG(D5)', "", "Automatic Kekulé alternating double bond representation", '=SMILES2IMG(D5)'),
        @(3, "Basic", "Acetone", "CC(=O)C", "Inline SMILES Literal", '=SMILES2IMG("CC(=O)C")', "", "Direct string literal input without external cell reference", '=SMILES2IMG("CC(=O)C")'),
        @(4, "Basic", "Pyridine", "c1ccncc1", "Default", '=SMILES2IMG(D7)', "", "Nitrogen heterocyclic aromatic ring structure", '=SMILES2IMG(D7)'),
        @(5, "Basic", "Ibuprofen", "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O", "Default", '=SMILES2IMG(D8)', "", "NSAID drug core with asymmetric chiral wedge bond", '=SMILES2IMG(D8)'),

        @(6, "Background", "Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", 'bg="white"', '=SMILES2IMG(D9, "white")', "", "Opaque pure-white canvas background", '=SMILES2IMG(D9, "white")'),
        @(7, "Background", "Paracetamol", "CC(=O)Nc1ccc(O)cc1", 'bg="cornsilk"', '=SMILES2IMG(D10, "cornsilk")', "", "Standard CSS color name (cornsilk light yellow)", '=SMILES2IMG(D10, "cornsilk")'),
        @(8, "Background", "Dopamine", "NCCc1ccc(O)c(O)c1", 'bg="aliceblue"', '=SMILES2IMG(D11, "aliceblue")', "", "Standard CSS color name (aliceblue soft light blue)", '=SMILES2IMG(D11, "aliceblue")'),
        @(9, "Background", "Serotonin", "NCCc1c[nH]c2ccc(O)cc12", 'bg="#F0F4F8"', '=SMILES2IMG(D12, "#F0F4F8")', "", "Custom 6-digit Hex background color code", '=SMILES2IMG(D12, "#F0F4F8")'),

        @(10, "Single Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="bw"', '=SMILES2IMG(D13, , "bw")', "", "100% monochrome line art for patent (USPTO/KIPO) & journals", '=SMILES2IMG(D13, , "bw")'),
        @(11, "Single Style", "L-Lactic Acid", "C[C@H](O)C(=O)O", 'style="stereo"', '=SMILES2IMG(D14, , "stereo")', "", "Explicit R/S chiral centers and E/Z double-bond stereochemistry", '=SMILES2IMG(D14, , "stereo")'),
        @(12, "Single Style", "Ethanol", "CCO", 'style="h"', '=SMILES2IMG(D15, , "h")', "", "Explicit C-H bond lines and hydrogens to inspect steric hindrance", '=SMILES2IMG(D15, , "h")'),
        @(13, "Single Style", "Benzene", "c1ccccc1", 'style="num"', '=SMILES2IMG(D16, , "num")', "", "Sequential atom indices (1, 2, 3...) for NMR peak assignment", '=SMILES2IMG(D16, , "num")'),
        @(14, "Single Style", "Acetone", "CC(=O)C", 'style="all-c"', '=SMILES2IMG(D17, , "all-c")', "", "Explicit 'C' label on all skeletal carbons (educational use)", '=SMILES2IMG(D17, , "all-c")'),
        @(15, "Single Style", "Ibuprofen", "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O", 'style="ink=#003366"', '=SMILES2IMG(D18, "aliceblue", "ink=#003366")', "", "Corporate CI navy bond lines on aliceblue canvas", '=SMILES2IMG(D18, "aliceblue", "ink=#003366")'),
        @(16, "Single Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="white" (Dark)', '=SMILES2IMG(D19, "#1e1e1e", "white")', "", "Pure-white bond lines for dark presentation slides", '=SMILES2IMG(D19, "#1e1e1e", "white")'),

        @(17, "Multi-Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="bw num" (Space)', '=SMILES2IMG(D20, , "bw num")', "", "Monochrome line art combined with atom index numbering", '=SMILES2IMG(D20, , "bw num")'),
        @(18, "Multi-Style", "L-Lactic Acid", "C[C@H](O)C(=O)O", 'style="bw|stereo" (Pipe)', '=SMILES2IMG(D21, , "bw|stereo")', "", "Pipe (|) delimited monochrome drawing with chiral annotations", '=SMILES2IMG(D21, , "bw|stereo")'),
        @(19, "Multi-Style", "Ethanol", "CCO", 'style="num h" (Space)', '=SMILES2IMG(D22, , "num h")', "", "Combined atom indexing and explicit hydrogens for reaction tracking", '=SMILES2IMG(D22, , "num h")'),
        @(20, "Multi-Style", "Nicotine", "CN1CCC[C@H]1c2cccnc2", 'style="white stereo num h"', '=SMILES2IMG(D23, "#1a1a1a", "white stereo num h")', "", "Dark canvas with pure-white bonds, chiral labels, numbers, and hydrogens", '=SMILES2IMG(D23, "#1a1a1a", "white stereo num h")'),

        @(21, "Transform", "Nicotine", "CN1CCC[C@H]1c2cccnc2", 'transform="flip"', '=SMILES2IMG(D24, , , "flip")', "", "Horizontal reflection matching standard textbook & database layout", '=SMILES2IMG(D24, , , "flip")'),
        @(22, "Transform", "Cholesterol", "CC(C)CCCC(C)C1CCC2C1(CCC3C2CC=C4C3(CCC(C4)O)C)C", 'transform="flipy"', '=SMILES2IMG(D25, , , "flipy")', "", "Vertical 2D planar reflection", '=SMILES2IMG(D25, , , "flipy")'),
        @(23, "Transform", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'transform="rot90"', '=SMILES2IMG(D26, , , "rot90")', "", "90-deg clockwise rotation for compact vertical column layouts", '=SMILES2IMG(D26, , , "rot90")'),
        @(24, "Transform", "Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", 'transform="rot180"', '=SMILES2IMG(D27, , , "rot180")', "", "180-deg upside-down rotation", '=SMILES2IMG(D27, , , "rot180")'),
        @(25, "Transform", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'transform="flip rot90"', '=SMILES2IMG(D28, , , "flip rot90")', "", "Chained transformation: horizontal flip followed by 90-deg rotation", '=SMILES2IMG(D28, , , "flip rot90")'),

        @(26, "All-in-One", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'White + Full (Space) + Rot', '=SMILES2IMG(D29, "white", "bw num stereo h all-c", "flip rot90")', "", "All features: White canvas + B/W + Atom numbers + Chiral + H + Carbons + Flip + 90-deg rot", '=SMILES2IMG(D29, "white", "bw num stereo h all-c", "flip rot90")'),
        @(27, "All-in-One", "Penicillin G", "CC1(C(N2C(S1)C(C2=O)NC(=O)Cc3ccccc3)C(=O)O)C", 'White + Full (Pipe) + Rot', '=SMILES2IMG(D30, "white", "bw|num|stereo|h|all-c", "flip rot90")', "", "Pipe (|) delimited ultimate all-in-one formula activating all capabilities", '=SMILES2IMG(D30, "white", "bw|num|stereo|h|all-c", "flip rot90")')
    )

    Populate-SheetData $sheet1 "SMILES2IMG Comprehensive Showcase (Microsoft 365 / Excel 2024 Native In-Cell)" $s1Headers $s1Rows 0x593E1F 0x794E1F $s1ColWidths $true

    # =========================================================================
    # SHEET 2: SMILES2IMG.FLOAT (Excel 2016+)
    # =========================================================================
    Write-Host "Populating Sheet 2: SMILES2IMG.FLOAT (Excel 2016+)..." -ForegroundColor Yellow
    $sheet2.Tab.Color = 0x228B22 # Forest Green BGR

    $s2Headers = @("No.", "Category", "Compound Name", "Chemical SMILES", "Applied Option", "Formula (Text)", "Floating Molecular Shape (Over Cell)", "Features & Behavior")
    $s2ColWidths = @(6, 16, 22, 32, 26, 48, 28, 48)

    $s2Rows = @(
        @(1, "Basic", "Ethanol", "CCO", "Default", '=SMILES2IMG.FLOAT(D4)', "", "Automatic cell-anchored floating picture on Excel 2016/2019/2021", '=SMILES2IMG.FLOAT(D4)'),
        @(2, "Basic", "Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "Default", '=SMILES2IMG.FLOAT(D5)', "", "Linked to cell movement and resizing (xlMoveAndSize)", '=SMILES2IMG.FLOAT(D5)'),
        @(3, "Background", "Paracetamol", "CC(=O)Nc1ccc(O)cc1", 'bg="white"', '=SMILES2IMG.FLOAT(D6, "white")', "", "Opaque white canvas floating shape", '=SMILES2IMG.FLOAT(D6, "white")'),
        @(4, "Single Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="bw"', '=SMILES2IMG.FLOAT(D7, , "bw")', "", "Monochrome patent line art floating shape for legacy Excel", '=SMILES2IMG.FLOAT(D7, , "bw")'),
        @(5, "Single Style", "L-Lactic Acid", "C[C@H](O)C(=O)O", 'style="stereo"', '=SMILES2IMG.FLOAT(D8, , "stereo")', "", "Chiral R/S stereochemical labels in floating mode", '=SMILES2IMG.FLOAT(D8, , "stereo")'),
        @(6, "Single Style", "Ethanol", "CCO", 'style="h"', '=SMILES2IMG.FLOAT(D9, , "h")', "", "All C-H hydrogens unfolded in floating mode", '=SMILES2IMG.FLOAT(D9, , "h")'),
        @(7, "Multi-Style", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'style="bw num"', '=SMILES2IMG.FLOAT(D10, , "bw num")', "", "Monochrome combined with atom indices in floating mode", '=SMILES2IMG.FLOAT(D10, , "bw num")'),
        @(8, "Multi-Style", "Nicotine", "CN1CCC[C@H]1c2cccnc2", 'style="white stereo"', '=SMILES2IMG.FLOAT(D11, "#1e1e1e", "white stereo")', "", "Dark mode canvas with pure-white bonds and chiral labels", '=SMILES2IMG.FLOAT(D11, "#1e1e1e", "white stereo")'),
        @(9, "Transform", "Nicotine", "CN1CCC[C@H]1c2cccnc2", 'transform="flip"', '=SMILES2IMG.FLOAT(D12, , , "flip")', "", "Textbook horizontal flip in floating mode", '=SMILES2IMG.FLOAT(D12, , , "flip")'),
        @(10, "All-in-One", "Aspirin", "CC(=O)Oc1ccccc1C(=O)O", 'White + Full (Space) + Rot', '=SMILES2IMG.FLOAT(D13, "white", "bw num stereo h all-c", "rot90")', "", "Ultimate all-in-one floating molecular rendering on legacy Excel", '=SMILES2IMG.FLOAT(D13, "white", "bw num stereo h all-c", "rot90")')
    )

    Populate-SheetData $sheet2 "SMILES2IMG.FLOAT Floating Mode Showcase (Excel 2016 / 2019 / 2021 Compatible)" $s2Headers $s2Rows 0x1B4D1B 0x228B22 $s2ColWidths $true

    # =========================================================================
    # SHEET 3: Complex Drugs & Natural Products
    # =========================================================================
    Write-Host "Populating Sheet 3: Complex Natural Products..." -ForegroundColor Yellow
    $sheet3.Tab.Color = 0x800080 # Purple BGR

    $s3Headers = @("No.", "Category", "Compound Name", "Chemical SMILES", "Structural Features & Chiral Centers", "Formula (Text)", "Molecular Structure (In-Cell)", "Pharmacology & Significance")
    $s3ColWidths = @(6, 18, 24, 34, 32, 48, 32, 48)

    $s3Rows = @(
        @(1, "Antimalarial", "Artemisinin", "CC1CC2CC3(C)OO4C(O2)(C1C(=O)O3)C(C)CC4", "7 chiral centers, endoperoxide bridge", '=SMILES2IMG(D4, , "stereo")', "", "Nobel Prize-winning antimalarial natural product from sweet wormwood", '=SMILES2IMG(D4, , "stereo")'),
        @(2, "Antineoplastic", "Paclitaxel", "CC1=C2C(C(=O)C3(C(CC4C(C3C(C(C2(C)C)(CC1OC(=O)C(C(c5ccccc5)NC(=O)c6ccccc6)O)O)OC(=O)c7ccccc7)(CO4)OC(=O)C)O)C)OC(=O)C", "11 chiral centers, complex polycyclic core", '=SMILES2IMG(D5)', "", "Taxus brevifolia bark-derived tubulin inhibitor anticancer blockbuster", '=SMILES2IMG(D5)'),
        @(3, "Biological Sterol", "Cholesterol", "CC(C)CCCC(C)C1CCC2C1(CCC3C2CC=C4C3(CCC(C4)O)C)C", "8 chiral centers, 4-ring fused steroid core", '=SMILES2IMG(D6, , "stereo")', "", "Cell membrane fluidity regulator and steroid hormone precursor", '=SMILES2IMG(D6, , "stereo")'),
        @(4, "Opioid Analgesic", "Morphine", "CN1CC[C@]23[C@@H]4Oc5c3c(CC1[C@@H]2C=C[C@@H]4O)ccc5O", "5 chiral centers, 5 fused rings", '=SMILES2IMG(D7, , "stereo")', "", "Opium poppy-derived potent central opioid analgesic alkaloid", '=SMILES2IMG(D7, , "stereo")'),
        @(5, "Beta-lactam", "Penicillin G", "CC1(C(N2C(S1)C(C2=O)NC(=O)Cc3ccccc3)C(=O)O)C", "3 chiral centers, thiazolidine-beta-lactam core", '=SMILES2IMG(D8, "white", "bw num stereo")', "", "Historic broad-spectrum antibiotic rendered in patent style", '=SMILES2IMG(D8, "white", "bw num stereo")'),
        @(6, "Lipid-Lowering", "Atorvastatin", "CC(C)c1c(c(c(c1NC(=O)c2ccccc2)c3ccc(cc3)F)c4ccccc4)CC(O)CC(O)CC(=O)O", "2 chiral centers, HMG-CoA reductase inhibitor", '=SMILES2IMG(D9)', "", "Top-selling statin blockbuster in pharmaceutical history (Lipitor)", '=SMILES2IMG(D9)'),
        @(7, "Biological Energy", "ATP", "Nc1ncnc2n(cnc12)[C@@H]3O[C@H](COP(=O)(O)OP(=O)(O)OP(=O)(O)O)[C@@H](O)[C@H]3O", "4 chiral centers, triphosphate anhydride", '=SMILES2IMG(D10)', "", "Universal energy currency powering cellular metabolism", '=SMILES2IMG(D10)'),
        @(8, "Antimicrobial", "Sulfamethoxazole", "Cc1cc(NS(=O)(=O)c2ccc(N)cc2)no1", "Isoxazole-containing sulfonamide", '=SMILES2IMG(D11, , "bw")', "", "Folate synthesis inhibitor antibacterial agent for UTI treatment", '=SMILES2IMG(D11, , "bw")'),
        @(9, "Organic Dye", "Indigo Dye", "O=C1C(=C2Nc3ccccc3C2=O)Nc4ccccc14", "Bis-indoline conjugated chromophore", '=SMILES2IMG(D12, "aliceblue", "ink=#003366")', "", "Classic denim dye rendered in corporate CI ink and canvas", '=SMILES2IMG(D12, "aliceblue", "ink=#003366")'),
        @(10, "Carbohydrate", "D-Glucose", "OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O", "5 chiral centers, pyranose hemiacetal ring", '=SMILES2IMG(D13, , "stereo num")', "", "Primary fuel of cellular respiration in biology & biochemistry", '=SMILES2IMG(D13, , "stereo num")'),
        @(11, "Neurotransmitter", "Dopamine", "NCCc1ccc(O)c(O)c1", "Catecholamine neurotransmitter", '=SMILES2IMG(D14, , "num h")', "", "Key neurotransmitter in reward pathways and motor control", '=SMILES2IMG(D14, , "num h")'),
        @(12, "Natural Flavoring", "Vanillin", "O=Cc1ccc(O)c(OC)c1", "Aromatic aldehyde-phenol", '=SMILES2IMG(D15, , "all-c")', "", "Classic phenolic flavor compound from vanilla pods", '=SMILES2IMG(D15, , "all-c")')
    )

    Populate-SheetData $sheet3 "Complex Natural Products & Blockbuster Pharmaceuticals Gallery" $s3Headers $s3Rows 0x4B0082 0x800080 $s3ColWidths $true

    # =========================================================================
    # SHEET 4: Style & Transform Matrix
    # =========================================================================
    Write-Host "Populating Sheet 4: Style & Transform Matrix..." -ForegroundColor Yellow
    $sheet4.Tab.Color = 0x0080FF # Orange BGR

    $s4Headers = @("No.", "Style / Transform Category", "Delimiter", "Option String", "Formula (Text)", "Rendered Structure (In-Cell)", "Visual Effect & Output Description")
    $s4ColWidths = @(6, 28, 12, 26, 68, 28, 48)

    Invoke-Excel {
        $sheet4.Range("A2").Value2 = "Ref:"
        $sheet4.Range("A2").Font.Bold = $true
        $sheet4.Range("A2").HorizontalAlignment = -4152 # xlRight
        $sheet4.Range("B2").Value2 = "CC(=O)Oc1ccccc1C(=O)O"
        $sheet4.Range("B2").Font.Bold = $true
        $sheet4.Range("B2").Interior.Color = 0xE6F2FF
        $sheet4.Range("B2").HorizontalAlignment = -4108 # xlCenter
        $sheet4.Range("C2").Value2 = "(Aspirin Reference Molecule)"
        $sheet4.Range("C2").Font.Italic = $true
        $sheet4.Range("C2").Font.Color = 0x595959
    }

    $s4Rows = @(
        @(1, "1. Default Rendering", "-", '(default)', '=SMILES2IMG($B$2)', "", "Transparent canvas + CPK elemental coloring", '=SMILES2IMG($B$2)'),
        @(2, "2. Black & White Line Art", "-", '"bw"', '=SMILES2IMG($B$2, , "bw")', "", "100% monochrome line art for patent offices (USPTO/KIPO)", '=SMILES2IMG($B$2, , "bw")'),
        @(3, "3. Sequential Atom Numbers", "-", '"num"', '=SMILES2IMG($B$2, , "num")', "", "Numbered atom indices (1, 2, 3...) for NMR assignment", '=SMILES2IMG($B$2, , "num")'),
        @(4, "4. Stereochemical Labels", "-", '"stereo"', '=SMILES2IMG($B$2, , "stereo")', "", "Chiral R/S centers and geometric E/Z stereolabels", '=SMILES2IMG($B$2, , "stereo")'),
        @(5, "5. Unfolded Hydrogens", "-", '"h"', '=SMILES2IMG($B$2, , "h")', "", "Explicit C-H bond lines displaying all hydrogens", '=SMILES2IMG($B$2, , "h")'),
        @(6, "6. Skeletal Carbon Labels", "-", '"all-c"', '=SMILES2IMG($B$2, , "all-c")', "", "Explicit 'C' text at every skeletal vertex", '=SMILES2IMG($B$2, , "all-c")'),
        @(7, "7. Dark Mode Pure-White Ink", "-", '"white"', '=SMILES2IMG($B$2, "#1e1e1e", "white")', "", "Crisp pure-white bond lines against dark cell canvas", '=SMILES2IMG($B$2, "#1e1e1e", "white")'),
        @(8, "8. Corporate Custom Ink", "-", '"ink=#003366"', '=SMILES2IMG($B$2, "aliceblue", "ink=#003366")', "", "Specified Hex color bond ink on aliceblue background", '=SMILES2IMG($B$2, "aliceblue", "ink=#003366")'),
        @(9, "9. B/W + Atom Numbers", "Space", '"bw num"', '=SMILES2IMG($B$2, , "bw num")', "", "Space-separated monochrome patent line art with numbers", '=SMILES2IMG($B$2, , "bw num")'),
        @(10, "10. B/W + Chiral Labels", "Pipe (|)", '"bw|stereo"', '=SMILES2IMG($B$2, , "bw|stereo")', "", "Pipe-separated monochrome line art with chiral labels", '=SMILES2IMG($B$2, , "bw|stereo")'),
        @(11, "11. Transform: 90-deg CW", "-", '"rot90"', '=SMILES2IMG($B$2, , , "rot90")', "", "90-deg clockwise rotation for vertical column layouts", '=SMILES2IMG($B$2, , , "rot90")'),
        @(12, "12. Transform: Horizontal Flip", "-", '"flip"', '=SMILES2IMG($B$2, , , "flip")', "", "Horizontally flipped matching chemical database conventions", '=SMILES2IMG($B$2, , , "flip")'),
        @(13, "13. Transform: Flip + 90-deg", "Space", '"flip rot90"', '=SMILES2IMG($B$2, , , "flip rot90")', "", "Chained flip followed by 90-deg rotation", '=SMILES2IMG($B$2, , , "flip rot90")'),
        @(14, "14. Full-Option (Space)", "Space", '"bw num stereo h all-c"', '=SMILES2IMG($B$2, "white", "bw num stereo h all-c", "flip rot90")', "", "All options: White + B/W + Numbers + Chiral + H + Carbons + Flip + 90-deg rot", '=SMILES2IMG($B$2, "white", "bw num stereo h all-c", "flip rot90")'),
        @(15, "15. Full-Option (Pipe)", "Pipe (|)", '"bw|num|stereo|h|all-c"', '=SMILES2IMG($B$2, "white", "bw|num|stereo|h|all-c", "flip rot90")', "", "Pipe-delimited ultimate all-in-one formula activating all features", '=SMILES2IMG($B$2, "white", "bw|num|stereo|h|all-c", "flip rot90")')
    )

    Populate-SheetData $sheet4 "Style & Transform Rendering Options Matrix (Aspirin CC(=O)Oc1ccccc1C(=O)O)" $s4Headers $s4Rows 0x0045FF 0x0080FF $s4ColWidths $false

    # Activate Sheet 1
    $sheet1.Activate()

    # Recalculate
    Write-Host "`nRecalculating formulas and waiting for async image insertion..." -ForegroundColor Cyan
    Invoke-Excel { $excel.Calculate() }

    # Settle down
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $maxWaitSec = 60
    do {
        Start-Sleep -Seconds 2
        $shapeCount = 0
        Invoke-Excel { $shapeCount = $sheet2.Shapes.Count }
        Write-Host "Elapsed: $([int]$sw.Elapsed.TotalSeconds)s | Sheet2 Floating Shapes: $shapeCount / $($s2Rows.Count)" -ForegroundColor Gray
        
        if ($shapeCount -ge $s2Rows.Count -and $sw.Elapsed.TotalSeconds -ge 12) {
            Start-Sleep -Seconds 3
            break
        }
    } while ($sw.Elapsed.TotalSeconds -lt $maxWaitSec)

    # Save to Target Path
    Write-Host "`nSaving workbook to: $targetFile" -ForegroundColor Cyan
    Invoke-Excel {
        $workbook.SaveAs($targetFile, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
    }
    Write-Host "Workbook saved successfully!" -ForegroundColor Green

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
# File Verification & OpenXML Inspection
# =========================================================================
Write-Host "`n=== Verifying Generated Workbook Integrity ===" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $targetFile)) {
    throw "Output file was not created: $targetFile"
}

$fileSize = (Get-Item -LiteralPath $targetFile).Length
Write-Host "File Size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Green

$zip = [System.IO.Compression.ZipFile]::OpenRead($targetFile)
try {
    $entries = $zip.Entries | Select-Object -ExpandProperty FullName
    
    $sheetEntries = $entries | Where-Object { $_ -like 'xl/worksheets/sheet*.xml' }
    Write-Host "Worksheet XML Entries: $($sheetEntries.Count)" -ForegroundColor Green
    
    $mediaEntries = $entries | Where-Object { $_ -like 'xl/media/*.png' }
    Write-Host "Media (PNG) Entries: $($mediaEntries.Count)" -ForegroundColor Green

    $richDataEntries = $entries | Where-Object { $_ -like 'xl/richData/*' }
    Write-Host "RichData Entries: $($richDataEntries.Count)" -ForegroundColor Green

    $metadataEntry = $entries | Where-Object { $_ -eq 'xl/metadata.xml' }
    if ($metadataEntry) {
        Write-Host "Metadata XML present: Yes (Native In-Cell XLRICHVALUE metadata)" -ForegroundColor Green
    } else {
        Write-Host "Metadata XML present: No" -ForegroundColor Yellow
    }

    if ($sheetEntries.Count -lt 4) {
        throw "Expected 4 worksheets, found $($sheetEntries.Count)"
    }
    if ($mediaEntries.Count -eq 0) {
        throw "No media PNG files found inside the workbook!"
    }
} finally {
    $zip.Dispose()
}

Write-Host "`n🎉 ALL VERIFICATION PASSED! File generated at:" -ForegroundColor Green
Write-Host $targetFile -ForegroundColor Cyan
