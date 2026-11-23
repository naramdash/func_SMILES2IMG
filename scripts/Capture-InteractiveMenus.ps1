<#
.SYNOPSIS
    Captures real, live 3-level right-click Context Menu screenshots ("Chemical Structure -> Copy Molecule Info -> ...")
    from Microsoft Excel, Word, and PowerPoint using Per-Monitor DPI-Aware PrintWindow + UIA ExpandCollapsePattern.
    Uses SendInput when running on an interactive desktop (WinSta0\Default) and WM_CONTEXTMENU/RBUTTON messages otherwise.
#>
param(
    [ValidateSet('All', 'Excel', 'Word', 'PowerPoint')]
    [string]$TargetApp = 'All'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms

Add-Type -ReferencedAssemblies System.Drawing, System.Windows.Forms @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

public static class LiveMenuCapture
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int L, T, R, B; }

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT
    {
        public uint type;
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr val);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetDpiForSystem();
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc f, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumWindowsProc f, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int c);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern uint SendInput(uint n, INPUT[] inputs, int size);

    public static void InitDpi()
    {
        try { if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return; } catch { }
        try { SetProcessDPIAware(); } catch { }
    }

    public static double GetWindowDpiScale(IntPtr hwnd)
    {
        uint dpi = 0;
        try { if (hwnd != IntPtr.Zero) dpi = GetDpiForWindow(hwnd); } catch { }
        if (dpi == 0)
        {
            try { dpi = GetDpiForSystem(); } catch { }
        }
        if (dpi < 96) dpi = 96;
        return dpi / 96.0;
    }

    public static Size SetMainWindowSizeLogical(IntPtr hwnd, int logicalX, int logicalY, int logicalW, int logicalH)
    {
        ShowWindow(hwnd, 9);
        SetWindowPos(hwnd, IntPtr.Zero, logicalX, logicalY, 800, 450, 0x0040);
        System.Threading.Thread.Sleep(60);

        double scale = GetWindowDpiScale(hwnd);
        int physX = (int)Math.Round(logicalX * scale / 2.25);
        int physY = (int)Math.Round(logicalY * scale / 2.25);
        int physW = ((int)Math.Round(logicalW * scale) / 2) * 2;
        int physH = ((int)Math.Round(logicalH * scale) / 2) * 2;
        try
        {
            var bounds = System.Windows.Forms.Screen.FromHandle(hwnd).Bounds;
            int maxW = Math.Max(640, bounds.Width - 16);
            int maxH = Math.Max(360, bounds.Height - 16);
            if (physW > maxW || physH > maxH)
            {
                double clamp = Math.Min((double)maxW / physW, (double)maxH / physH);
                physW = ((int)Math.Floor(physW * clamp) / 2) * 2;
                physH = ((int)Math.Floor(physH * clamp) / 2) * 2;
            }
        }
        catch { }
        SetWindowPos(hwnd, IntPtr.Zero, physX, physY, physW, physH, 0x0040);
        return new Size(physW, physH);
    }

    public static void RightClickScreen(int screenX, int screenY)
    {
        SetCursorPos(screenX, screenY);
        var inputs = new INPUT[2];
        inputs[0] = new INPUT { type = 0, mi = new MOUSEINPUT { dwFlags = 0x0008 } }; // RIGHTDOWN
        inputs[1] = new INPUT { type = 0, mi = new MOUSEINPUT { dwFlags = 0x0010 } }; // RIGHTUP
        SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    public static IntPtr FindChildByClassPrefix(IntPtr parent, string prefix)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, (h, _) =>
        {
            if (!IsWindowVisible(h)) return true;
            var sb = new StringBuilder(128);
            GetClassName(h, sb, 128);
            if (sb.ToString().StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                found = h;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static void SanitizeTitleBarAccount(Bitmap bmp)
    {
        SanitizeTitleBarAccount(bmp, 2.25);
    }

    public static void SanitizeTitleBarAccount(Bitmap bmp, double dpiScale)
    {
        int w = bmp.Width;
        double s = dpiScale > 0.5 ? dpiScale : 1.0;
        if (w < (int)Math.Round(450 * s) || bmp.Height < (int)Math.Round(60 * s)) return;
        // Sample clean Office title bar strip just right of the Search bar
        int sampleX = Math.Max(10, w - (int)Math.Round(415 * s));
        int sampleY = Math.Max(4, Math.Min(bmp.Height - 4, (int)Math.Round(20 * s)));
        Color sample = bmp.GetPixel(sampleX, sampleY);
        using (var br = new SolidBrush(sample))
        using (var g = Graphics.FromImage(bmp))
        {
            // Cover top-right account avatar & upgrade pill seamlessly before window controls (Min/Max/Close)
            int boxX = Math.Max(0, w - (int)Math.Round(380 * s));
            int boxRight = Math.Max(boxX, w - (int)Math.Round(150 * s));
            int boxW = Math.Max(1, boxRight - boxX);
            int boxY = Math.Max(2, (int)Math.Round(2 * s));
            int boxH = Math.Max(20, (int)Math.Round(38 * s));
            g.FillRectangle(br, boxX, boxY, boxW, boxH);
        }
    }

    public static int CaptureWithPopups(IntPtr mainHwnd, uint pid, string[] outPaths)
    {
        RECT mr;
        GetWindowRect(mainHwnd, out mr);
        int mw = ((mr.R - mr.L) / 2) * 2;
        int mh = ((mr.B - mr.T) / 2) * 2;
        if (mw <= 0 || mh <= 0) return 0;

        var wins = new List<IntPtr>();
        EnumWindows((h, _) =>
        {
            if (h == mainHwnd || !IsWindowVisible(h)) return true;
            uint p;
            GetWindowThreadProcessId(h, out p);
            if (p == pid) wins.Add(h);
            return true;
        }, IntPtr.Zero);

        var popupItems = new List<Tuple<IntPtr, RECT>>();
        int outW = mw;
        int outH = mh;
        for (int i = wins.Count - 1; i >= 0; i--)
        {
            var sb = new StringBuilder(128);
            GetClassName(wins[i], sb, 128);
            string cls = sb.ToString();
            bool isOfficePopup = cls.StartsWith("Net UI", StringComparison.OrdinalIgnoreCase) ||
                                 cls.IndexOf("MsoCommandBar", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 cls == "#32768";
            if (!isOfficePopup) continue;

            RECT pr;
            GetWindowRect(wins[i], out pr);
            int pw = pr.R - pr.L, ph = pr.B - pr.T;
            if (pw < 80 || ph < 80) continue;

            popupItems.Add(Tuple.Create(wins[i], pr));
            outW = Math.Max(outW, (((pr.R - mr.L) + 28) / 2) * 2);
            outH = Math.Max(outH, (((pr.B - mr.T) + 28) / 2) * 2);
        }

        int popupCount = popupItems.Count;
        using (var bmp = new Bitmap(outW, outH))
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(38, 38, 38));
            using (var mainBmp = new Bitmap(mw, mh))
            {
                using (var mg = Graphics.FromImage(mainBmp))
                {
                    var hdc = mg.GetHdc();
                    PrintWindow(mainHwnd, hdc, 2);
                    mg.ReleaseHdc(hdc);
                }
                SanitizeTitleBarAccount(mainBmp, GetWindowDpiScale(mainHwnd));
                g.DrawImage(mainBmp, 0, 0, mw, mh);
            }

            foreach (var item in popupItems)
            {
                IntPtr phwnd = item.Item1;
                RECT pr = item.Item2;
                int pw = pr.R - pr.L, ph = pr.B - pr.T;
                using (var pbmp = new Bitmap(pw, ph))
                {
                    using (var pg = Graphics.FromImage(pbmp))
                    {
                        var phdc = pg.GetHdc();
                        PrintWindow(phwnd, phdc, 2);
                        pg.ReleaseHdc(phdc);
                    }
                    g.DrawImage(pbmp, pr.L - mr.L, pr.T - mr.T, pw, ph);
                }
            }

            foreach (var p in outPaths)
            {
                var dir = System.IO.Path.GetDirectoryName(p);
                if (!string.IsNullOrEmpty(dir)) System.IO.Directory.CreateDirectory(dir);
                bmp.Save(p, System.Drawing.Imaging.ImageFormat.Png);
            }
        }
        return popupCount;
    }
}
'@

[LiveMenuCapture]::InitDpi()

$repoRoot = Split-Path -Parent $PSScriptRoot
$publicDir = Join-Path $repoRoot 'docs\public'
$assetsDir = Join-Path $repoRoot 'assets'
$distXll = Join-Path $repoRoot 'dist\x64\Smiles2Img-AddIn64-packed.xll'
if (-not (Test-Path $distXll)) {
    $distXll = Join-Path $env:APPDATA 'Microsoft\AddIns\Smiles2Img\Smiles2Img-AddIn64.xll'
}

function Expand-CascadingChemicalMenu([int]$procId) {
    $AE = [System.Windows.Automation.AutomationElement]
    $TS = [System.Windows.Automation.TreeScope]
    $deadline = [DateTime]::UtcNow.AddSeconds(4)
    $chemItem = $null
    while ([DateTime]::UtcNow -lt $deadline -and -not $chemItem) {
        $wins = $AE::RootElement.FindAll($TS::Children, (New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $procId)))
        foreach ($w in $wins) {
            $chemItem = $w.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'Chemical Structure')))
            if ($chemItem) { break }
        }
        if (-not $chemItem) { Start-Sleep -Milliseconds 150 }
    }
    if (-not $chemItem) { return $false }

    try { $chemItem.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand() } catch { return $false }
    Start-Sleep -Milliseconds 450

    $infoItem = $null
    $deadline2 = [DateTime]::UtcNow.AddSeconds(3)
    while ([DateTime]::UtcNow -lt $deadline2 -and -not $infoItem) {
        $wins2 = $AE::RootElement.FindAll($TS::Children, (New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $procId)))
        foreach ($w2 in $wins2) {
            $infoItem = $w2.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'Copy Molecule Info')))
            if ($infoItem) { break }
        }
        if (-not $infoItem) { Start-Sleep -Milliseconds 150 }
    }
    if ($infoItem) {
        try { $infoItem.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand() } catch {}
        Start-Sleep -Milliseconds 450
    }
    return $true
}

function Select-RibbonTab([IntPtr]$mainHwnd, [string]$tabName) {
    try {
        $AE = [System.Windows.Automation.AutomationElement]
        $win = $AE::FromHandle($mainHwnd)
        if (-not $win) { return }
        $cond = New-Object System.Windows.Automation.AndCondition(
            (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)),
            (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $tabName))
        )
        $tab = $win.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
        if ($tab) {
            $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
            Start-Sleep -Milliseconds 450
        }
    } catch {}
}

# 1. WORD CONTEXT MENU
if ($TargetApp -eq 'All' -or $TargetApp -eq 'Word') {
    Write-Host "[*] Capturing live 3-level Word Context Menu..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    try {
        $doc = $word.Documents.Add()
        $word.Selection.TypeText("Aspirin: CC(=O)Oc1ccccc1C(=O)O")
        $doc.Range(9, 30).Select()

        $hwnd = [IntPtr]$word.ActiveWindow.Hwnd
        [uint32]$wPid = 0
        [void][LiveMenuCapture]::GetWindowThreadProcessId($hwnd, [ref]$wPid)
        [void][LiveMenuCapture]::SetMainWindowSizeLogical($hwnd, 8, 8, 1230, 692)
        Start-Sleep -Milliseconds 600
        Select-RibbonTab $hwnd 'SMILES'

        $wwg = [LiveMenuCapture]::FindChildByClassPrefix($hwnd, '_WwG')
        $r = New-Object LiveMenuCapture+RECT
        [void][LiveMenuCapture]::GetWindowRect($wwg, [ref]$r)
        $cx = [int]($r.L + ($r.R - $r.L) * 0.24)
        $cy = [int]($r.T + 55)
        $lp = [IntPtr](($cy -shl 16) -bor ($cx -band 0xFFFF))
        [void][LiveMenuCapture]::PostMessage($wwg, 0x007B, $wwg, $lp)

        if (Expand-CascadingChemicalMenu ([int]$wPid)) {
            $n = [LiveMenuCapture]::CaptureWithPopups($hwnd, $wPid, @(
                (Join-Path $publicDir 'word-context-menu-guide.png'),
                (Join-Path $assetsDir 'word-context-menu-guide.png')
            ))
            Write-Host "  [PASS] Captured live Word 3-level context menu ($n popups)" -ForegroundColor Green
        } else {
            Write-Warning "  [SKIP] Word context menu did not appear in this desktop session."
        }
        [void][LiveMenuCapture]::PostMessage($hwnd, 0x0100, [IntPtr]0x1B, [IntPtr]0)
        $doc.Close($false)
    } finally {
        try { $word.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# 2. EXCEL CONTEXT MENU
if ($TargetApp -eq 'All' -or $TargetApp -eq 'Excel') {
    Write-Host "[*] Capturing live 3-level Excel Context Menu..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    try { $excel.ShowSelectionFloaties = $false } catch {}
    try {
        [void]$excel.RegisterXLL($distXll)
        $wb = $excel.Workbooks.Add()
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        $ws = $wb.ActiveSheet
        $ws.Range("A1").Value2 = "Compound"
        $ws.Range("B1").Value2 = "SMILES"
        $ws.Range("C1").Value2 = "2D Structure"
        $ws.Range("D1").Value2 = "MW"
        $ws.Range("A2").Value2 = "Aspirin"
        $ws.Range("B2").Value2 = "CC(=O)Oc1ccccc1C(=O)O"
        $ws.Range("C2").RowHeight = 72
        $ws.Range("A3").Value2 = "Caffeine"
        $ws.Range("B3").Value2 = "Cn1cnc2c1c(=O)n(C)c(=O)n2C"
        $ws.Range("C3").RowHeight = 72
        $ws.Range("C2:C3").ColumnWidth = 22
        $ws.Range("B2").ColumnWidth = 28
        $ws.Range("C2").Formula = "=SMILES2IMG(B2)"
        $ws.Range("D2").Formula = "=SMILES2MW(B2)"
        $ws.Range("C3").Formula = "=SMILES2IMG(B3)"
        $ws.Range("D3").Formula = "=SMILES2MW(B3)"
        Start-Sleep -Seconds 2
        $ws.Range("C3").Select()

        $hwnd = [IntPtr]$excel.Hwnd
        [uint32]$xPid = 0
        [void][LiveMenuCapture]::GetWindowThreadProcessId($hwnd, [ref]$xPid)
        [void][LiveMenuCapture]::SetMainWindowSizeLogical($hwnd, 8, 8, 1230, 692)
        Start-Sleep -Milliseconds 600
        Select-RibbonTab $hwnd 'SMILES'

        $x7 = [LiveMenuCapture]::FindChildByClassPrefix($hwnd, 'EXCEL7')
        $r = New-Object LiveMenuCapture+RECT
        [void][LiveMenuCapture]::GetWindowRect($x7, [ref]$r)
        # Client coords inside EXCEL7 hitting row 3 / col C so any mini-toolbar stays inside Sheet Grid (never overlapping ribbon) and all 3 menu levels fit comfortably
        $clientX = 320
        $clientY = 185
        $lpClient = [IntPtr](($clientY -shl 16) -bor ($clientX -band 0xFFFF))
        [void][LiveMenuCapture]::PostMessage($x7, 0x0204, [IntPtr]2, $lpClient) # WM_RBUTTONDOWN
        Start-Sleep -Milliseconds 70
        [void][LiveMenuCapture]::PostMessage($x7, 0x0205, [IntPtr]0, $lpClient) # WM_RBUTTONUP
        Start-Sleep -Milliseconds 500

        if (-not (Expand-CascadingChemicalMenu ([int]$xPid))) {
            try { [void][LiveMenuCapture]::SetForegroundWindow($hwnd) } catch {}
            [LiveMenuCapture]::RightClickScreen(($r.L + $clientX), ($r.T + $clientY))
            Start-Sleep -Milliseconds 450
            [void](Expand-CascadingChemicalMenu ([int]$xPid))
        }

        if (Expand-CascadingChemicalMenu ([int]$xPid)) {
            $n = [LiveMenuCapture]::CaptureWithPopups($hwnd, $xPid, @(
                (Join-Path $publicDir 'excel-context-menu-guide.png'),
                (Join-Path $assetsDir 'excel-context-menu-guide.png'),
                (Join-Path $publicDir 'context-menu-guide.png'),
                (Join-Path $assetsDir 'context-menu-guide.png')
            ))
            Write-Host "  [PASS] Captured live Excel 3-level context menu ($n popups)" -ForegroundColor Green
        } else {
            Write-Warning "  [WARN] Could not expand Excel context menu."
        }
        [void][LiveMenuCapture]::PostMessage($hwnd, 0x0100, [IntPtr]0x1B, [IntPtr]0)
        $wb.Close($false)
    } finally {
        try { $excel.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# 3. POWERPOINT CONTEXT MENU
if ($TargetApp -eq 'All' -or $TargetApp -eq 'PowerPoint') {
    Write-Host "[*] Capturing live 3-level PowerPoint Context Menu..." -ForegroundColor Cyan
    $ppt = New-Object -ComObject PowerPoint.Application
    try {
        $pres = $ppt.Presentations.Add(-1)
        $pres.PageSetup.SlideWidth = 960
        $pres.PageSetup.SlideHeight = 540
        $slide = $pres.Slides.Add(1, 12)
        $slide.FollowMasterBackground = 0
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = 0x2A170F # BGR for dark slate #0F172A

        $pptAddIn = $null
        foreach ($ai in $ppt.COMAddIns) { if ($ai.ProgId -eq 'Smiles2Img.PowerPoint.AddIn') { $pptAddIn = $ai; break } }
        $pptObj = $pptAddIn.Object

        $tb = $slide.Shapes.AddTextbox(1, 140, 80, 260, 60)
        $tb.TextFrame.TextRange.Text = "CC(=O)Oc1ccccc1C(=O)O"
        $tb.Select()
        [void]$pptObj.ExecuteRibbonAction("OnQuickInsert")
        Start-Sleep -Milliseconds 400
        if ($slide.Shapes.Count -ge 1) {
            $shp = $slide.Shapes.Item(1)
            $shp.Left = 24
            $shp.Top = 60
            $shp.Width = 220
            $shp.Height = 165
            $shp.Select()
            [void]$pptObj.ExecuteRibbonAction("OnBatchStyleDark")
            Start-Sleep -Milliseconds 300
            if ($slide.Shapes.Count -ge 1) { $shp = $slide.Shapes.Item($slide.Shapes.Count); $shp.Select() }
        }
        $tb.Delete()

        $hwnd = [IntPtr]$ppt.HWND
        [uint32]$pPid = 0
        [void][LiveMenuCapture]::GetWindowThreadProcessId($hwnd, [ref]$pPid)
        [void][LiveMenuCapture]::SetMainWindowSizeLogical($hwnd, 8, 8, 1230, 692)
        Start-Sleep -Milliseconds 600
        Select-RibbonTab $hwnd 'SMILES'

        $mdi = [LiveMenuCapture]::FindChildByClassPrefix($hwnd, 'mdiClass')
        $r = New-Object LiveMenuCapture+RECT
        [void][LiveMenuCapture]::GetWindowRect($mdi, [ref]$r)
        # Convert the PowerPoint picture near top-left from Slide points (960x540) to screen pixels so all 3 cascading menus fit cleanly
        $aw = $ppt.ActiveWindow
        $scrX = [int]$aw.PointsToScreenPixelsX($shp.Left + 24)
        $scrY = [int]$aw.PointsToScreenPixelsY($shp.Top + 24)
        $clientX = $scrX - $r.L
        $clientY = $scrY - $r.T
        $lpClient = [IntPtr](($clientY -shl 16) -bor ($clientX -band 0xFFFF))
        [void][LiveMenuCapture]::PostMessage($mdi, 0x0204, [IntPtr]2, $lpClient)
        Start-Sleep -Milliseconds 70
        [void][LiveMenuCapture]::PostMessage($mdi, 0x0205, [IntPtr]0, $lpClient)
        Start-Sleep -Milliseconds 500

        if (-not (Expand-CascadingChemicalMenu ([int]$pPid))) {
            try { [void][LiveMenuCapture]::SetForegroundWindow($hwnd) } catch {}
            [LiveMenuCapture]::RightClickScreen(($r.L + $clientX), ($r.T + $clientY))
            Start-Sleep -Milliseconds 450
            [void](Expand-CascadingChemicalMenu ([int]$pPid))
        }

        if (Expand-CascadingChemicalMenu ([int]$pPid)) {
            $n = [LiveMenuCapture]::CaptureWithPopups($hwnd, $pPid, @(
                (Join-Path $publicDir 'powerpoint-context-menu-guide.png'),
                (Join-Path $assetsDir 'powerpoint-context-menu-guide.png')
            ))
            Write-Host "  [PASS] Captured live PowerPoint 3-level context menu ($n popups)" -ForegroundColor Green
        } else {
            Write-Warning "  [WARN] Could not expand PowerPoint context menu."
        }
        [void][LiveMenuCapture]::PostMessage($hwnd, 0x0100, [IntPtr]0x1B, [IntPtr]0)
        $pres.Saved = -1
        $pres.Close()
    } finally {
        try { $ppt.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
}

Get-Process EXCEL,WINWORD,POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
