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
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
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
        int w = bmp.Width;
        if (w < 900 || bmp.Height < 120) return;
        // Sample the Office title bar background left of the search bar (x=240, y=22)
        Color sample = bmp.GetPixel(240, 22);
        using (var br = new SolidBrush(sample))
        using (var g = Graphics.FromImage(bmp))
        {
            // Cover top-right account avatar & upgrade pill before window controls (Min/Max/Close)
            int boxW = 495;
            int boxX = Math.Max(0, w - 730);
            g.FillRectangle(br, boxX, 8, boxW, 68);
        }
    }

    public static int CaptureWithPopups(IntPtr mainHwnd, uint pid, string[] outPaths)
    {
        RECT mr;
        GetWindowRect(mainHwnd, out mr);
        int mw = ((mr.R - mr.L) / 2) * 2;
        int mh = ((mr.B - mr.T) / 2) * 2;
        if (mw <= 0 || mh <= 0) return 0;

        int popupCount = 0;
        using (var bmp = new Bitmap(mw, mh))
        using (var g = Graphics.FromImage(bmp))
        {
            var hdc = g.GetHdc();
            PrintWindow(mainHwnd, hdc, 2);
            g.ReleaseHdc(hdc);
            SanitizeTitleBarAccount(bmp);

            var wins = new List<IntPtr>();
            EnumWindows((h, _) =>
            {
                if (h == mainHwnd || !IsWindowVisible(h)) return true;
                uint p;
                GetWindowThreadProcessId(h, out p);
                if (p == pid) wins.Add(h);
                return true;
            }, IntPtr.Zero);

            for (int i = wins.Count - 1; i >= 0; i--)
            {
                var sb = new StringBuilder(128);
                GetClassName(wins[i], sb, 128);
                string cls = sb.ToString();
                if (cls.Contains("InputIndicator") || cls.Contains("IME") || cls.Contains("Tooltip")) continue;

                RECT pr;
                GetWindowRect(wins[i], out pr);
                int pw = pr.R - pr.L, ph = pr.B - pr.T;
                if (pw < 80 || ph < 80) continue;

                popupCount++;
                using (var pbmp = new Bitmap(pw, ph))
                {
                    using (var pg = Graphics.FromImage(pbmp))
                    {
                        var phdc = pg.GetHdc();
                        PrintWindow(wins[i], phdc, 2);
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
        [void][LiveMenuCapture]::SetWindowPos($hwnd, [IntPtr]::Zero, 40, 10, 2400, 1720, 0x0040)
        Start-Sleep -Milliseconds 600

        $wwg = [LiveMenuCapture]::FindChildByClassPrefix($hwnd, '_WwG')
        $r = New-Object LiveMenuCapture+RECT
        [void][LiveMenuCapture]::GetWindowRect($wwg, [ref]$r)
        $cx = [int]($r.L + ($r.R - $r.L) * 0.24)
        $cy = [int]($r.T + 85)
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
    try {
        [void]$excel.RegisterXLL($distXll)
        $wb = $excel.Workbooks.Add()
        $ws = $wb.ActiveSheet
        $ws.Range("A1").Value2 = "Compound"
        $ws.Range("B1").Value2 = "SMILES"
        $ws.Range("C1").Value2 = "2D Structure"
        $ws.Range("D1").Value2 = "MW"
        $ws.Range("A2").Value2 = "Aspirin"
        $ws.Range("B2").Value2 = "CC(=O)Oc1ccccc1C(=O)O"
        $ws.Range("C2").RowHeight = 85
        $ws.Range("C2").ColumnWidth = 22
        $ws.Range("B2").ColumnWidth = 26
        $ws.Range("C2").Formula = "=SMILES2IMG(B2)"
        $ws.Range("D2").Formula = "=SMILES2MW(B2)"
        Start-Sleep -Seconds 2
        $ws.Range("C2").Select()

        $hwnd = [IntPtr]$excel.Hwnd
        [uint32]$xPid = 0
        [void][LiveMenuCapture]::GetWindowThreadProcessId($hwnd, [ref]$xPid)
        [void][LiveMenuCapture]::SetWindowPos($hwnd, [IntPtr]::Zero, 40, 10, 2400, 1720, 0x0040)
        Start-Sleep -Milliseconds 600

        $x7 = [LiveMenuCapture]::FindChildByClassPrefix($hwnd, 'EXCEL7')
        $r = New-Object LiveMenuCapture+RECT
        [void][LiveMenuCapture]::GetWindowRect($x7, [ref]$r)
        # Client coords inside EXCEL7 hitting row 2 / col C near top of sheet
        $clientX = 290
        $clientY = 95
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
        [void][LiveMenuCapture]::SetWindowPos($hwnd, [IntPtr]::Zero, 40, 10, 2400, 1720, 0x0040)
        Start-Sleep -Milliseconds 600

        $mdi = [LiveMenuCapture]::FindChildByClassPrefix($hwnd, 'mdiClass')
        $r = New-Object LiveMenuCapture+RECT
        [void][LiveMenuCapture]::GetWindowRect($mdi, [ref]$r)
        # Convert the PowerPoint picture left-center from Slide points (960x540) to screen pixels
        $aw = $ppt.ActiveWindow
        $scrX = [int]$aw.PointsToScreenPixelsX($shp.Left + $shp.Width * 0.20)
        $scrY = [int]$aw.PointsToScreenPixelsY($shp.Top + $shp.Height * 0.25)
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
