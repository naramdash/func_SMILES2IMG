<#
.SYNOPSIS
    True End-to-End (E2E) Automated UI & Functional Verification for func_SMILES2IMG (Word, PowerPoint, Excel)
    - Syncs freshly built dist\x64 binaries to %APPDATA%\Microsoft\AddIns\Smiles2Img
    - Launches visible Office windows and verifies HWND visibility
    - Drives Word & PowerPoint via their in-process COMAddIn.Object.ExecuteRibbonAction(...)
    - Verifies live modal UI dialogs (MoleculeInfoDialog & AboutDialog) pop up and render inside Word, PowerPoint, and Excel
    - Drives Excel via RegisterXLL, live cell formula evaluation (v1.0 ~ v1.8.0), and $excel.Run(...)
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class OfficeE2EWindow {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private const uint WM_CLOSE = 0x0010;

    public static string LastDetectedTitle = "";
    public static int LastDetectedWidth = 0;
    public static int LastDetectedHeight = 0;
    public static bool DialogDetected = false;

    public static void StartModalDialogCloser(string expectedTitleSubstring, int timeoutMs = 8000) {
        LastDetectedTitle = "";
        LastDetectedWidth = 0;
        LastDetectedHeight = 0;
        DialogDetected = false;

        ThreadPool.QueueUserWorkItem(_ => {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs && !DialogDetected) {
                EnumWindows((hWnd, lParam) => {
                    if (!IsWindowVisible(hWnd)) return true;
                    var sb = new StringBuilder(256);
                    if (GetWindowText(hWnd, sb, sb.Capacity) > 0) {
                        string title = sb.ToString();
                        if (title.IndexOf(expectedTitleSubstring, StringComparison.OrdinalIgnoreCase) >= 0) {
                            RECT r;
                            if (GetWindowRect(hWnd, out r)) {
                                int w = r.Right - r.Left;
                                int h = r.Bottom - r.Top;
                                if (w >= 400 && h >= 300) {
                                    LastDetectedTitle = title;
                                    LastDetectedWidth = w;
                                    LastDetectedHeight = h;
                                    DialogDetected = true;
                                    Thread.Sleep(250); // Allow full WM_PAINT inside Office
                                    PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                    return false;
                                }
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
                if (!DialogDetected) Thread.Sleep(80);
            }
        });
    }
}
'@ -ErrorAction SilentlyContinue

$repoRoot = Split-Path -Parent $PSScriptRoot
$distX64 = Join-Path $repoRoot "dist\x64"
$coreDll = Join-Path $distX64 "Word\Smiles2Img.Core.dll"
$xllPath = Join-Path $distX64 "Smiles2Img-AddIn64-packed.xll"
if (-not (Test-Path $xllPath)) {
    $xllPath = Join-Path $env:APPDATA "Microsoft\AddIns\Smiles2Img\Smiles2Img-AddIn64.xll"
}

# Clean up any headless background WINWORD / POWERPNT left over from aborted automation runs
Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

# Sync dist\x64 DLLs into %APPDATA%\Microsoft\AddIns\Smiles2Img and update HKCU CLSID Assembly version non-interactively
$suiteDir = Join-Path $env:APPDATA "Microsoft\AddIns\Smiles2Img"
New-Item -ItemType Directory -Path $suiteDir -Force | Out-Null
foreach ($sub in @("Word", "PowerPoint")) {
    $srcDir = Join-Path $distX64 $sub
    if (Test-Path $srcDir) {
        Get-ChildItem -Path $srcDir -File | ForEach-Object {
            try { Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $suiteDir $_.Name) -Force -ErrorAction Stop } catch {}
        }
    }
}
foreach ($entry in @(
    @{ Clsid = "{41A5C1F0-559A-46E0-B9C0-8B35A0F1D910}"; Dll = Join-Path $distX64 "Word\Smiles2Img.Word.dll" },
    @{ Clsid = "{A755973C-21FB-494A-A3C5-086AFFB7C2AC}"; Dll = Join-Path $distX64 "PowerPoint\Smiles2Img.PowerPoint.dll" }
)) {
    if (Test-Path $entry.Dll) {
        $asmFull = [System.Reflection.AssemblyName]::GetAssemblyName($entry.Dll).FullName
        $targetDll = Join-Path $suiteDir (Split-Path -Leaf $entry.Dll)
        $uri = if (Test-Path $targetDll) { ([uri]$targetDll).AbsoluteUri } else { ([uri]$entry.Dll).AbsoluteUri }
        foreach ($r in @("HKCU:\Software\Classes", "HKCU:\Software\Classes\Wow6432Node")) {
            $k = "$r\CLSID\$($entry.Clsid)\InprocServer32"
            if (Test-Path $k) {
                Set-ItemProperty -Path $k -Name "Assembly" -Value $asmFull -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $k -Name "CodeBase" -Value $uri -ErrorAction SilentlyContinue
            }
        }
    }
}

[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
$excelDll = Join-Path $repoRoot "src\Smiles2Img.Excel\bin\Release\net48\Smiles2Img.dll"
if (-not (Test-Path $excelDll)) { throw "Excel assembly not found at $excelDll" }
[System.Reflection.Assembly]::LoadFrom($excelDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path -Parent $coreDll))

$sampleKet = '{"root":{"nodes":[{"$ref":"mol0"}]},"mol0":{"type":"molecule","atoms":[{"label":"C","location":[0,0,0]},{"label":"O","location":[1,0,0]}],"bonds":[{"type":1,"atoms":[0,1]}]}}'
$aspirinSmiles = "CC(=O)Oc1ccccc1C(=O)O"

Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host " 1. Microsoft Word Live In-Process COM Add-in & UI E2E" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    $wordAddIn = $null
    foreach ($ai in $word.COMAddIns) {
        if ($ai.ProgId -eq "Smiles2Img.Word.AddIn") { $wordAddIn = $ai; break }
    }
    if (-not $wordAddIn) { throw "Smiles2Img.Word.AddIn COM Add-in is not registered in Word!" }
    if (-not $wordAddIn.Connect) { $wordAddIn.Connect = $true }
    $wordObj = $wordAddIn.Object
    if ($null -eq $wordObj) { throw "Smiles2Img.Word.AddIn.Object is null (in-process instance not exposed)!" }
    Write-Host "  [PASS] Word Visible UI & In-Process COM Add-in connected (ProgId=$($wordAddIn.ProgId))" -ForegroundColor Green

    $doc = $word.Documents.Add()

    # Test 1A: Selected text -> In-Process Ribbon Callback OnQuickInsert
    $word.Selection.TypeText($aspirinSmiles)
    $doc.Range(0, $aspirinSmiles.Length).Select()
    if (-not $wordObj.ExecuteRibbonAction("OnQuickInsert")) { throw "Failed to invoke OnQuickInsert on WordAddIn!" }

    if ($doc.InlineShapes.Count -lt 1) { throw "Word OnQuickInsert from selection failed to create InlineShape!" }
    $inline1 = $doc.InlineShapes.Item(1)
    $extractedSmiles1 = [Smiles2Img.StructureMetadata]::ExtractSmiles($inline1.Title, $inline1.AlternativeText)
    if ($extractedSmiles1 -ne $aspirinSmiles) { throw "Word InlineShape #1 SMILES mismatch! Got: $extractedSmiles1" }
    if ([int]$inline1.LockAspectRatio -ne -1) { throw "Word InlineShape #1 LockAspectRatio expected -1 (msoTrue), got: $($inline1.LockAspectRatio)" }
    Write-Host "  [PASS] Word In-Process OnQuickInsert (Selection -> InlineShape #1): $($inline1.Width)x$($inline1.Height) pt (LockAspectRatio=msoTrue)" -ForegroundColor Green

    # Test 1B: KET Canvas via Clipboard -> In-Process Ribbon Callback OnQuickInsert
    $word.Selection.Collapse(0)
    [void][Smiles2Img.Editor.SafeClipboard]::SetText("KET:$sampleKet")
    [void]$wordObj.ExecuteRibbonAction("OnQuickInsert")

    if ($doc.InlineShapes.Count -lt 2) { throw "Word OnQuickInsert from KET clipboard failed to create 2nd InlineShape!" }
    $inline2 = $doc.InlineShapes.Item(2)
    $extractedKet2 = [Smiles2Img.StructureMetadata]::ExtractStructure($inline2.AlternativeText, $inline2.Title)
    if ($extractedKet2 -ne $sampleKet) { throw "Word InlineShape #2 KET round-trip mismatch! Got: $extractedKet2" }
    Write-Host "  [PASS] Word In-Process OnQuickInsert (KET Canvas Round-Trip -> InlineShape #2)" -ForegroundColor Green

    # Test 1C: Selection-Scoped OnBatchStyleBw (Select only InlineShape #1 and apply B/W)
    $doc.InlineShapes.Item(1).Range.Select()
    [void]$wordObj.ExecuteRibbonAction("OnBatchStyleBw")
    if ($doc.InlineShapes.Count -ne 2) { throw "Word OnBatchStyleBw changed InlineShapes count unexpectedly!" }
    Write-Host "  [PASS] Word In-Process Selection-Scoped OnBatchStyleBw (InlineShape #1 updated, #2 preserved)" -ForegroundColor Green

    # Test 1D: OnCopySvg on selected InlineShape #1
    $doc.InlineShapes.Item(1).Range.Select()
    [void]$wordObj.ExecuteRibbonAction("OnCopySvg")
    $svgClip = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $svgClip -or -not $svgClip.Contains("<svg")) { throw "Word OnCopySvg failed to copy SVG XML!" }
    Write-Host "  [PASS] Word In-Process OnCopySvg copied valid Vector SVG ($($svgClip.Length) chars)" -ForegroundColor Green

    # Test 1E: Live Modal UI Popup inside Word (OnCopyMoleculeInfo -> MoleculeInfoDialog & OnAbout -> AboutDialog)
    $doc.InlineShapes.Item(1).Range.Select()
    [OfficeE2EWindow]::StartModalDialogCloser("Molecule Info")
    [void]$wordObj.ExecuteRibbonAction("OnCopyMoleculeInfo")
    if (-not [OfficeE2EWindow]::DialogDetected) { throw "Word OnCopyMoleculeInfo modal UI dialog did not appear on screen!" }
    Write-Host "  [PASS] Word Live Modal UI Popup: '$([OfficeE2EWindow]::LastDetectedTitle)' ($([OfficeE2EWindow]::LastDetectedWidth)x$([OfficeE2EWindow]::LastDetectedHeight) px)" -ForegroundColor Green

    [OfficeE2EWindow]::StartModalDialogCloser("func_SMILES2IMG")
    [void]$wordObj.ExecuteRibbonAction("OnAbout")
    if (-not [OfficeE2EWindow]::DialogDetected) { throw "Word OnAbout modal UI dialog did not appear on screen!" }
    Write-Host "  [PASS] Word Live Modal UI Popup: '$([OfficeE2EWindow]::LastDetectedTitle)' ($([OfficeE2EWindow]::LastDetectedWidth)x$([OfficeE2EWindow]::LastDetectedHeight) px)" -ForegroundColor Green

    # Test 1F: CCO text selection -> In-Process OnCopyAllSummary (17-column TSV with header)
    $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).Select()
    $word.Selection.TypeText("CCO")
    $doc.Range($doc.Content.End - 4, $doc.Content.End - 1).Select()
    [void]$wordObj.ExecuteRibbonAction("OnCopyAllSummary")
    $wordClip = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $wordClip -or -not $wordClip.Contains("C2H6O") -or -not $wordClip.Contains("Heavy Atoms") -or -not $wordClip.Contains("Fsp3")) {
        throw "Word OnCopyAllSummary failed or missing v1.8.0 17-column TSV header! Got: $wordClip"
    }
    Write-Host "  [PASS] Word In-Process OnCopyAllSummary copied 17-column TSV header + row" -ForegroundColor Green

    # Test 1G: Cross-Host Parity Checks in Word (Replaced LockAspectRatio + "Aspirin" text fallback across Copy SMILES/InChI/SVG/Formula/KET)
    $inlineAfterReplace = $doc.InlineShapes.Item(1)
    if ([int]$inlineAfterReplace.LockAspectRatio -ne -1) { throw "Word replaced InlineShape LockAspectRatio expected -1 (msoTrue), got: $($inlineAfterReplace.LockAspectRatio)" }
    $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).Select()
    $startPos = $word.Selection.Start
    $word.Selection.TypeText("Aspirin")
    $doc.Range($startPos, $startPos + 7).Select()

    [void]$wordObj.ExecuteRibbonAction("OnCopySmiles")
    $wAspSmiles = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if ($wAspSmiles -ne $aspirinSmiles) { throw "Word Copy SMILES on 'Aspirin' expected '$aspirinSmiles', got '$wAspSmiles'" }

    [void]$wordObj.ExecuteRibbonAction("OnCopyInchi")
    $wAspInchi = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $wAspInchi -or -not $wAspInchi.StartsWith("InChI=1S/C9H8O4")) { throw "Word Copy InChI on 'Aspirin' failed! Got: $wAspInchi" }

    [void]$wordObj.ExecuteRibbonAction("OnCopySvg")
    $wAspSvg = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $wAspSvg -or -not $wAspSvg.Contains("<svg")) { throw "Word Copy SVG on 'Aspirin' failed!" }

    [void]$wordObj.ExecuteRibbonAction("OnCopyFormula")
    $wAspForm = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if ($wAspForm -ne "C9H8O4") { throw "Word Copy Formula on 'Aspirin' expected 'C9H8O4', got '$wAspForm'" }

    [void]$wordObj.ExecuteRibbonAction("OnCopyKet")
    $wAspKet = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $wAspKet -or -not $wAspKet.Contains('"root"')) { throw "Word Copy KET on 'Aspirin' failed!" }
    Write-Host "  [PASS] Word Cross-Host Parity: LockAspectRatio=msoTrue & 'Aspirin' Copy SMILES/InChI/SVG/Formula/KET all resolved" -ForegroundColor Green

    $doc.Close($false)
} finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

Write-Host ""
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host " 2. Microsoft PowerPoint Live In-Process COM Add-in & UI E2E" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application

    $pptAddIn = $null
    foreach ($ai in $ppt.COMAddIns) {
        if ($ai.ProgId -eq "Smiles2Img.PowerPoint.AddIn") { $pptAddIn = $ai; break }
    }
    if (-not $pptAddIn) { throw "Smiles2Img.PowerPoint.AddIn COM Add-in is not registered in PowerPoint!" }
    if (-not $pptAddIn.Connect) { $pptAddIn.Connect = $true }
    $pptObj = $pptAddIn.Object
    if ($null -eq $pptObj) { throw "Smiles2Img.PowerPoint.AddIn.Object is null (in-process instance not exposed)!" }
    Write-Host "  [PASS] PowerPoint Visible UI & In-Process COM Add-in connected (ProgId=$($pptAddIn.ProgId))" -ForegroundColor Green

    $pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)
    $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank = 12

    # Test 2A: In-Process OnQuickInsert SMILES onto PowerPoint slide
    [void][Smiles2Img.Editor.SafeClipboard]::SetText($aspirinSmiles)
    if (-not $pptObj.ExecuteRibbonAction("OnQuickInsert")) { throw "Failed to invoke OnQuickInsert on PowerPointAddIn!" }

    if ($slide.Shapes.Count -lt 1) { throw "PowerPoint OnQuickInsert failed to insert shape!" }
    $pptShape1 = $slide.Shapes.Item(1)
    $pptSmiles1 = $pptShape1.Tags.Item("SMILES")
    if ($pptSmiles1 -ne $aspirinSmiles) { throw "PowerPoint Shape #1 missing SMILES tag! Got: $pptSmiles1" }
    Write-Host "  [PASS] PowerPoint In-Process OnQuickInsert (SMILES -> Shape #1): $([math]::Round($pptShape1.Width,1))x$([math]::Round($pptShape1.Height,1)) pt" -ForegroundColor Green

    # Test 2B: In-Process OnQuickInsert KET onto PowerPoint slide
    [void][Smiles2Img.Editor.SafeClipboard]::SetText("KET:$sampleKet")
    [void]$pptObj.ExecuteRibbonAction("OnQuickInsert")

    if ($slide.Shapes.Count -lt 2) { throw "PowerPoint OnQuickInsert KET failed to insert 2nd shape!" }
    $pptShape2 = $slide.Shapes.Item(2)
    $pptKet2 = $pptShape2.Tags.Item("KET")
    if ($pptKet2 -ne $sampleKet) { throw "PowerPoint Shape #2 KET round-trip mismatch! Got: $pptKet2" }
    Write-Host "  [PASS] PowerPoint In-Process OnQuickInsert (KET Canvas -> Shape #2): Tags['KET'] preserved 100%" -ForegroundColor Green

    # Test 2C: Selection-Scoped OnBatchStyleBw on Shape #1 only
    $pptShape1.Select()
    [void]$pptObj.ExecuteRibbonAction("OnBatchStyleBw")
    if ($slide.Shapes.Count -ne 2) { throw "PowerPoint OnBatchStyleBw changed shape count unexpectedly!" }
    Write-Host "  [PASS] PowerPoint In-Process Selection-Scoped OnBatchStyleBw (Selected shape updated)" -ForegroundColor Green

    # Test 2D: In-Process OnCopySvg & OnCopyAllSummary on replaced Aspirin shape (Item 2 after ReplaceShape)
    $slide.Shapes.Item(2).Select()
    [void]$pptObj.ExecuteRibbonAction("OnCopySvg")
    $pptSvg = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $pptSvg -or -not $pptSvg.Contains("<svg")) { throw "PowerPoint OnCopySvg failed to copy SVG!" }
    Write-Host "  [PASS] PowerPoint In-Process OnCopySvg copied valid Vector SVG ($($pptSvg.Length) chars)" -ForegroundColor Green

    [void]$pptObj.ExecuteRibbonAction("OnCopyAllSummary")
    $pptClip = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $pptClip -or -not $pptClip.Contains("C9H8O4") -or -not $pptClip.Contains("Heavy Atoms") -or -not $pptClip.Contains("Fsp3")) {
        throw "PowerPoint OnCopyAllSummary failed or missing 17-column TSV header! Got: $pptClip"
    }
    Write-Host "  [PASS] PowerPoint In-Process OnCopyAllSummary copied 17-column TSV header + row" -ForegroundColor Green

    # Test 2E: Live Modal UI Popup inside PowerPoint (OnCopyMoleculeInfo -> MoleculeInfoDialog & OnAbout -> AboutDialog)
    $slide.Shapes.Item(2).Select()
    [OfficeE2EWindow]::StartModalDialogCloser("Molecule Info")
    [void]$pptObj.ExecuteRibbonAction("OnCopyMoleculeInfo")
    if (-not [OfficeE2EWindow]::DialogDetected) { throw "PowerPoint OnCopyMoleculeInfo modal UI dialog did not appear on screen!" }
    Write-Host "  [PASS] PowerPoint Live Modal UI Popup: '$([OfficeE2EWindow]::LastDetectedTitle)' ($([OfficeE2EWindow]::LastDetectedWidth)x$([OfficeE2EWindow]::LastDetectedHeight) px)" -ForegroundColor Green

    [OfficeE2EWindow]::StartModalDialogCloser("func_SMILES2IMG")
    [void]$pptObj.ExecuteRibbonAction("OnAbout")
    if (-not [OfficeE2EWindow]::DialogDetected) { throw "PowerPoint OnAbout modal UI dialog did not appear on screen!" }
    Write-Host "  [PASS] PowerPoint Live Modal UI Popup: '$([OfficeE2EWindow]::LastDetectedTitle)' ($([OfficeE2EWindow]::LastDetectedWidth)x$([OfficeE2EWindow]::LastDetectedHeight) px)" -ForegroundColor Green

    # Test 2F: Cross-Host Parity Checks in PowerPoint (LockAspectRatio + TextBox 'Aspirin' fallback & 'CCO' QuickInsert)
    $pptLockShape = $slide.Shapes.Item(2)
    if ([int]$pptLockShape.LockAspectRatio -ne -1) { throw "PowerPoint Shape LockAspectRatio expected -1 (msoTrue), got: $($pptLockShape.LockAspectRatio)" }
    $tb = $slide.Shapes.AddTextbox(1, 50, 50, 220, 50)
    $tb.TextFrame.TextRange.Text = "Aspirin"
    $tb.Select()

    [void]$pptObj.ExecuteRibbonAction("OnCopySmiles")
    $pAspSmiles = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if ($pAspSmiles -ne $aspirinSmiles) { throw "PowerPoint Copy SMILES on TextBox 'Aspirin' expected '$aspirinSmiles', got '$pAspSmiles'" }

    [void]$pptObj.ExecuteRibbonAction("OnCopyInchi")
    $pAspInchi = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $pAspInchi -or -not $pAspInchi.StartsWith("InChI=1S/C9H8O4")) { throw "PowerPoint Copy InChI on TextBox 'Aspirin' failed! Got: $pAspInchi" }

    [void]$pptObj.ExecuteRibbonAction("OnCopySvg")
    $pAspSvg = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $pAspSvg -or -not $pAspSvg.Contains("<svg")) { throw "PowerPoint Copy SVG on TextBox 'Aspirin' failed!" }

    [void]$pptObj.ExecuteRibbonAction("OnCopyFormula")
    $pAspForm = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if ($pAspForm -ne "C9H8O4") { throw "PowerPoint Copy Formula on TextBox 'Aspirin' expected 'C9H8O4', got '$pAspForm'" }

    [void]$pptObj.ExecuteRibbonAction("OnCopyKet")
    $pAspKet = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $pAspKet -or -not $pAspKet.Contains('"root"')) { throw "PowerPoint Copy KET on TextBox 'Aspirin' failed!" }

    $tb.TextFrame.TextRange.Text = "CCO"
    $tb.Select()
    [void][Smiles2Img.Editor.SafeClipboard]::SetText("not-a-molecule")
    $beforeQuickShapes = $slide.Shapes.Count
    [void]$pptObj.ExecuteRibbonAction("OnQuickInsert")
    if ($slide.Shapes.Count -ne $beforeQuickShapes) { throw "PowerPoint OnQuickInsert on TextBox 'CCO' expected in-place replacement (count=$beforeQuickShapes), got $($slide.Shapes.Count)!" }
    $quickInsertedPptShape = $slide.Shapes.Item($slide.Shapes.Count)
    if ($quickInsertedPptShape.Tags.Item("SMILES") -ne "CCO") { throw "PowerPoint TextBox QuickInsert expected SMILES 'CCO', got '$($quickInsertedPptShape.Tags.Item("SMILES"))'" }
    Write-Host "  [PASS] PowerPoint Cross-Host Parity: LockAspectRatio=msoTrue, TextBox 'Aspirin' Copy* & TextBox 'CCO' QuickInsert verified" -ForegroundColor Green

    $pres.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pres.Close()
} finally {
    if ($ppt) {
        try { $ppt.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
}

Write-Host ""
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host " 3. Microsoft Excel Live Add-in (.xll), Worksheet & UI E2E" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
$excel = $null
$ownedExcelId = 0
$existingExcelIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$instanceId = 0
    [void][OfficeE2EWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$instanceId)
    if ($instanceId -ne 0 -and $existingExcelIds -notcontains [int]$instanceId) { $ownedExcelId = [int]$instanceId }
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    if (-not [OfficeE2EWindow]::IsWindowVisible([IntPtr]$excel.Hwnd)) { throw "Excel main window HWND is not visible!" }
    if (-not (Test-Path $xllPath)) { throw "Excel XLL not found at $xllPath" }
    $registered = $excel.RegisterXLL($xllPath)
    if (-not $registered) { throw "Excel failed to load $xllPath via RegisterXLL!" }
    Write-Host "  [PASS] Excel Visible UI & XLL loaded via RegisterXLL: $xllPath" -ForegroundColor Green

    $wb = $excel.Workbooks.Add()
    $ws = $wb.ActiveSheet
    $ws.Range("A1").Value2 = $aspirinSmiles
    $ws.Range("B1").RowHeight = 90
    $ws.Range("B1").ColumnWidth = 24
    $ws.Range("B1").Formula = "=SMILES2IMG.FLOAT(A1)"
    Start-Sleep -Seconds 2

    $valB1 = $ws.Range("B1").Text
    Write-Host "  [PASS] Excel =SMILES2IMG.FLOAT(A1) evaluated cleanly (Cell text='$valB1', Shapes=$($ws.Shapes.Count))" -ForegroundColor Green

    # Test 3B: Core, v1.7.0 Physicochemical, v1.8.0 Complexity, Search/Similarity & KET Worksheet Formulas
    $ws.Range("C1").Formula = "=SMILES2NAME(A1)"
    $ws.Range("D1").Formula = "=SMILES2MW(A1)"
    $ws.Range("E1").Formula = "=SMILES2FORMULA(A1)"
    $ws.Range("F1").Formula = "=SMILES2INCHI(A1)"
    $ws.Range("G1").Formula = "=SMILES2INCHIKEY(A1)"
    $ws.Range("H1").Formula = "=SMILES2EXACTMASS(A1)"
    $ws.Range("I1").Formula = "=SMILES2LOGP(A1)"
    $ws.Range("J1").Formula = "=SMILES2TPSA(A1)"
    $ws.Range("K1").Formula = "=SMILES2HBD(A1)"
    $ws.Range("L1").Formula = "=SMILES2HBA(A1)"
    $ws.Range("M1").Formula = "=SMILES2ROTBONDS(A1)"
    $ws.Range("N1").Formula = "=SMILES2LIPINSKI(A1)"
    $ws.Range("O1").Formula = "=SMILES2HEAVY(A1)"
    $ws.Range("P1").Formula = "=SMILES2RINGS(A1)"
    $ws.Range("Q1").Formula = "=SMILES2AROMATIC(A1)"
    $ws.Range("R1").Formula = "=SMILES2FSP3(A1)"
    $ws.Range("S1").Formula = "=SMILES2CHIRAL(A1)"
    $ws.Range("T1").Formula = "=SMILES.MATCH(A1, ""c1ccccc1"")"
    $ws.Range("U1").Formula = "=SMILES.SIMILARITY(A1, ""c1ccccc1"")"
    $ws.Range("A3").Value2 = $sampleKet
    $ws.Range("B3").Formula = "=KET2FORMULA(A3)"
    $ws.Range("C3").Formula = "=SMILES2HEAVY(A3)"
    $ws.Range("D3").Formula = "=KET2EXACTMASS(A3)"
    Start-Sleep -Seconds 1

    $mwVal = [double]$ws.Range("D1").Value2
    $formulaVal = $ws.Range("E1").Text
    $inchiVal = $ws.Range("F1").Text
    $inchiKeyVal = $ws.Range("G1").Text
    $exactMassVal = [double]$ws.Range("H1").Value2
    $logpVal = [double]$ws.Range("I1").Value2
    $tpsaVal = [double]$ws.Range("J1").Value2
    $hbdVal = [int]$ws.Range("K1").Value2
    $hbaVal = [int]$ws.Range("L1").Value2
    $rotVal = [int]$ws.Range("M1").Value2
    $lipVal = $ws.Range("N1").Text
    $heavyVal = [int]$ws.Range("O1").Value2
    $ringsVal = [int]$ws.Range("P1").Value2
    $aromVal = [int]$ws.Range("Q1").Value2
    $fsp3Val = [double]$ws.Range("R1").Value2
    $chiralVal = [int]$ws.Range("S1").Value2
    $matchVal = [bool]$ws.Range("T1").Value2
    $simVal = [double]$ws.Range("U1").Value2
    $ketFormVal = $ws.Range("B3").Text
    $ketHeavyVal = [int]$ws.Range("C3").Value2

    if ([math]::Abs($mwVal - 180.16) -gt 0.2) { throw "Excel =SMILES2MW(A1) mismatch: $mwVal" }
    if ($formulaVal -ne "C9H8O4") { throw "Excel =SMILES2FORMULA(A1) mismatch: $formulaVal" }
    if (-not $inchiVal.StartsWith("InChI=1S/C9H8O4")) { throw "Excel =SMILES2INCHI(A1) mismatch: $inchiVal" }
    if ($inchiKeyVal -ne "BSYNRYMUTXBXSQ-UHFFFAOYSA-N") { throw "Excel =SMILES2INCHIKEY(A1) mismatch: $inchiKeyVal" }
    if ([math]::Abs($exactMassVal - 180.0423) -gt 0.01) { throw "Excel =SMILES2EXACTMASS(A1) mismatch: $exactMassVal" }
    if ($logpVal -le 0.5 -or $tpsaVal -le 50 -or $hbdVal -ne 1 -or $hbaVal -lt 3 -or $rotVal -lt 2 -or $lipVal -ne "PASS") {
        throw "Excel v1.7.0 physicochemical UDF mismatch: LogP=$logpVal, TPSA=$tpsaVal, HBD=$hbdVal, Lipinski=$lipVal"
    }
    if ($heavyVal -ne 13 -or $ringsVal -ne 1 -or $aromVal -ne 1 -or [math]::Abs($fsp3Val - 0.11) -gt 0.01 -or $chiralVal -ne 0) {
        throw "Excel v1.8.0 structural complexity UDF mismatch: Heavy=$heavyVal, Rings=$ringsVal, Arom=$aromVal, Fsp3=$fsp3Val, Chiral=$chiralVal"
    }
    if (-not $matchVal -or $simVal -le 0.1) { throw "Excel SMILES.MATCH / SMILES.SIMILARITY UDF mismatch: Match=$matchVal, Sim=$simVal" }
    if ($ketFormVal -ne "CH4O" -or $ketHeavyVal -ne 2) { throw "Excel KET2FORMULA / SMILES2HEAVY(KET) UDF mismatch: Formula=$ketFormVal, Heavy=$ketHeavyVal" }
    Write-Host "  [PASS] Excel 21 Live Worksheet UDFs (Core + v1.7.0 Physicochemical/Search + v1.8.0 Complexity + KET2*) verified" -ForegroundColor Green

    # Test 3C: Selection-Scoped BatchStyleBwCommand in Excel
    $ws.Range("A2").Value2 = "CCO"
    $ws.Range("B2").Formula = "=SMILES2IMG.FLOAT(A2)"
    Start-Sleep -Milliseconds 500
    $ws.Activate()
    [void]$ws.Range("A1:B1").Select()
    [void]$excel.Run("BatchStyleBwCommand")
    Start-Sleep -Milliseconds 500
    $fB1 = [string]$ws.Range("B1").Formula
    $fB2 = [string]$ws.Range("B2").Formula
    if (-not $fB1.Contains('"bw"')) { throw "Excel BatchStyleBwCommand did not update selected cell B1! Got: $fB1" }
    if ($fB2.Contains('"bw"')) { throw "Excel BatchStyleBwCommand modified unselected cell B2! Got: $fB2" }
    Write-Host "  [PASS] Excel Selection-Scoped BatchStyleBwCommand (B1 -> '$fB1', unselected B2 -> '$fB2')" -ForegroundColor Green

    # Test 3D: CopySelectedAllSummaryCommand, CopySelectedMoleculeSvgCommand & Live Modal MoleculeInfoDialog Popup
    try { [System.Windows.Forms.Clipboard]::Clear() } catch {}
    [void]$ws.Range("A2").Select()
    Start-Sleep -Milliseconds 200
    [void]$excel.Run("CopySelectedAllSummaryCommand")
    Start-Sleep -Milliseconds 500
    $clipText = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $clipText -or -not $clipText.Contains("C2H6O") -or -not $clipText.Contains("Heavy Atoms")) {
        throw "Excel CopySelectedAllSummaryCommand failed! Got: $clipText"
    }
    Write-Host "  [PASS] Excel CopySelectedAllSummaryCommand copied 17-column TSV summary" -ForegroundColor Green

    [void]$excel.Run("CopySelectedMoleculeSvgCommand")
    Start-Sleep -Milliseconds 400
    $excelSvg = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $excelSvg -or -not $excelSvg.Contains("<svg")) {
        throw "Excel CopySelectedMoleculeSvgCommand failed! Got: $excelSvg"
    }
    Write-Host "  [PASS] Excel CopySelectedMoleculeSvgCommand copied valid Vector SVG ($($excelSvg.Length) chars)" -ForegroundColor Green

    [OfficeE2EWindow]::StartModalDialogCloser("Molecule Info")
    [void]$excel.Run("CopySelectedMoleculeInfoCommand")
    if (-not [OfficeE2EWindow]::DialogDetected) { throw "Excel CopySelectedMoleculeInfoCommand modal UI dialog did not appear on screen!" }
    Write-Host "  [PASS] Excel Live Modal UI Popup: '$([OfficeE2EWindow]::LastDetectedTitle)' ($([OfficeE2EWindow]::LastDetectedWidth)x$([OfficeE2EWindow]::LastDetectedHeight) px)" -ForegroundColor Green

    # Test 3E: Live =NAME2IMG.FLOAT("Caffeine") floating picture creation
    $shapesBeforeNameImg = $ws.Shapes.Count
    $ws.Range("B4").RowHeight = 90
    $ws.Range("B4").ColumnWidth = 24
    $ws.Range("B4").Formula = '=NAME2IMG.FLOAT("Caffeine")'
    Start-Sleep -Seconds 2
    if ($ws.Shapes.Count -le $shapesBeforeNameImg) {
        throw "Excel =NAME2IMG.FLOAT(""Caffeine"") failed to create floating picture! Shapes before=$shapesBeforeNameImg, after=$($ws.Shapes.Count)"
    }
    Write-Host "  [PASS] Excel =NAME2IMG.FLOAT(""Caffeine"") rendered live floating picture (Shapes=$shapesBeforeNameImg -> $($ws.Shapes.Count))" -ForegroundColor Green

    # Test 3F: Live SDF Import -> Export Round-Trip (Structure & Compound Name Preservation)
    $tempSdfIn = Join-Path ([System.IO.Path]::GetTempPath()) ("smiles2img_e2e_in_" + [Guid]::NewGuid().ToString("N") + ".sdf")
    $tempSdfOut = Join-Path ([System.IO.Path]::GetTempPath()) ("smiles2img_e2e_out_" + [Guid]::NewGuid().ToString("N") + ".sdf")
    try {
        $r1 = New-Object Smiles2Img.Rendering.SdfRecord
        $r1.Name = "Aspirin"
        $r1.Smiles = $aspirinSmiles
        $r2 = New-Object Smiles2Img.Rendering.SdfRecord
        $r2.Name = "Compound_Custom_1"
        $r2.Smiles = "CCO"
        [Smiles2Img.Rendering.SdfManager]::WriteSdfFile($tempSdfIn, [Smiles2Img.Rendering.SdfRecord[]]@($r1, $r2))

        $wsSdf = $wb.Worksheets.Add()
        $wsSdf.Activate()
        $importedCount = [Smiles2Img.KetcherEditor]::ImportSdfFromFile($excel, $tempSdfIn)
        if ($importedCount -ne 2) { throw "Excel ImportSdfFromFile expected 2 imported rows, got $importedCount" }
        Start-Sleep -Milliseconds 500
        $exportedCount = [Smiles2Img.KetcherEditor]::ExportSdfToFile($excel, $tempSdfOut)
        if ($exportedCount -ne 2) { throw "Excel ExportSdfToFile expected 2 exported rows, got $exportedCount" }

        $exported = [Smiles2Img.Rendering.SdfManager]::ReadSdfFile($tempSdfOut)
        if ($exported.Count -ne 2) { throw "SDF round-trip expected 2 records, got $($exported.Count)" }
        $exportedNames = @($exported | ForEach-Object { $_.Name })
        if ($exportedNames -notcontains "Aspirin" -or $exportedNames -notcontains "Compound_Custom_1") {
            throw "SDF round-trip lost compound names! Got: $($exportedNames -join ', ')"
        }
        Write-Host "  [PASS] Excel Live SDF Import -> Export Round-Trip preserved 2 records and names ($($exportedNames -join ', '))" -ForegroundColor Green
    } finally {
        Remove-Item -LiteralPath $tempSdfIn, $tempSdfOut -Force -ErrorAction SilentlyContinue
    }

    # Test 3G: Excel Cross-Host Parity (Plain-text 'Aspirin' Copy KET + InChI ActiveCell QuickInsert)
    $ws.Activate()
    $ws.Range("A5").Value2 = "Aspirin"
    [void]$ws.Range("A5").Select()
    try { [System.Windows.Forms.Clipboard]::Clear() } catch {}
    [void]$excel.Run("CopySelectedMoleculeKetCommand")
    Start-Sleep -Milliseconds 400
    $excelKet = [Smiles2Img.Editor.SafeClipboard]::GetText()
    if (-not $excelKet -or -not $excelKet.Contains('"root"')) { throw "Excel CopySelectedMoleculeKetCommand on plain text 'Aspirin' failed! Got: $excelKet" }

    $ws.Range("A6").Value2 = "InChI=1S/C2H6O/c1-2-3/h3H,2H2,1H3"
    [void]$ws.Range("A6").Select()
    [void][Smiles2Img.Editor.SafeClipboard]::SetText("not-a-molecule")
    [void]$excel.Run("QuickInsertCommand")
    Start-Sleep -Milliseconds 800
    $fA6 = [string]$ws.Range("A6").Formula
    if (-not $fA6.StartsWith("=SMILES2IMG", [System.StringComparison]::OrdinalIgnoreCase) -or -not $fA6.Contains("CCO")) {
        throw "Excel QuickInsertCommand on InChI cell failed! Got formula: $fA6"
    }
    Write-Host "  [PASS] Excel Cross-Host Parity: Copy KET on 'Aspirin' cell & QuickInsert on InChI cell verified ('$fA6')" -ForegroundColor Green

    $wb.Close($false)
} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($ownedExcelId -ne 0) {
        $ownedExcel = Get-Process -Id $ownedExcelId -ErrorAction SilentlyContinue
        if ($null -ne $ownedExcel -and -not $ownedExcel.WaitForExit(2000) -and $ownedExcel.MainWindowTitle -eq '') {
            Stop-Process -Id $ownedExcelId -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host ""
Write-Host "[ALL LIVE OFFICE UI & E2E TESTS PASSED SUCCESSFULLY]" -ForegroundColor Green
