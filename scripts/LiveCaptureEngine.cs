using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

public static class LiveCaptureEngine
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int L, T, R, B; }

    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr val);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc f, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int c);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int c);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);

    public static void InitDpi()
    {
        try { if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return; } catch { }
        try { SetProcessDPIAware(); } catch { }
    }

    public static uint GetPidForHwnd(IntPtr hwnd)
    {
        uint p;
        GetWindowThreadProcessId(hwnd, out p);
        return p;
    }

    public static void SetMainWindowSize(IntPtr hwnd, int x, int y, int w, int h)
    {
        ShowWindow(hwnd, 9); // SW_RESTORE
        SetWindowPos(hwnd, IntPtr.Zero, x, y, w, h, 0x0040);
    }

    public static bool SelectRibbonTab(IntPtr mainHwnd, string tabName)
    {
        try
        {
            var win = AutomationElement.FromHandle(mainHwnd);
            if (win == null) return false;
            var cond = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                new PropertyCondition(AutomationElement.NameProperty, tabName));
            var tab = win.FindFirst(TreeScope.Descendants, cond);
            if (tab == null) return false;
            var pat = (SelectionItemPattern)tab.GetCurrentPattern(SelectionItemPattern.Pattern);
            pat.Select();
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static Point FindControlRelativeCenter(IntPtr mainHwnd, params string[] names)
    {
        RECT mr;
        GetWindowRect(mainHwnd, out mr);
        try
        {
            var win = AutomationElement.FromHandle(mainHwnd);
            if (win == null) return new Point((mr.R - mr.L) / 2, 160);
            foreach (var name in names)
            {
                if (string.IsNullOrEmpty(name)) continue;
                var el = win.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
                if (el != null)
                {
                    var br = el.Current.BoundingRectangle;
                    if (!br.IsEmpty && br.Width > 0 && br.Height > 0)
                    {
                        int rx = (int)(br.Left + br.Width * 0.5) - mr.L;
                        int ry = (int)(br.Top + br.Height * 0.5) - mr.T;
                        return new Point(Math.Max(20, rx), Math.Max(20, ry));
                    }
                }
            }
        }
        catch { }
        return new Point(160, 160);
    }

    public static void SanitizeTitleBar(Bitmap bmp)
    {
        int w = bmp.Width;
        if (w < 900 || bmp.Height < 120) return;
        Color sample = bmp.GetPixel(Math.Min(240, w / 5), 22);
        using (var br = new SolidBrush(sample))
        using (var g = Graphics.FromImage(bmp))
        {
            int boxW = 495;
            int boxX = Math.Max(0, w - 730);
            g.FillRectangle(br, boxX, 8, boxW, 68);
        }
    }

    private static void DrawCursor(Graphics g, int x, int y, float ripple)
    {
        if (x < 0 || y < 0) return;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        if (ripple > 0f && ripple <= 1f)
        {
            int r = (int)(12 + ripple * 34);
            int alpha = (int)(220 * (1f - ripple));
            using (var pen = new Pen(Color.FromArgb(alpha, 16, 185, 129), 4f))
            {
                g.DrawEllipse(pen, x - r, y - r, r * 2, r * 2);
            }
        }

        Point[] pts = new[]
        {
            new Point(x, y),
            new Point(x, y + 34),
            new Point(x + 9, y + 26),
            new Point(x + 17, y + 40),
            new Point(x + 23, y + 37),
            new Point(x + 15, y + 23),
            new Point(x + 26, y + 23)
        };
        using (var shadow = new SolidBrush(Color.FromArgb(90, 0, 0, 0)))
        {
            g.TranslateTransform(2, 2);
            g.FillPolygon(shadow, pts);
            g.ResetTransform();
        }
        using (var fill = new SolidBrush(Color.FromArgb(255, 15, 23, 42)))
        using (var border = new Pen(Color.White, 2.5f))
        {
            g.FillPolygon(fill, pts);
            g.DrawPolygon(border, pts);
        }
    }

    private static bool IsCapturableChildWindow(IntPtr h, IntPtr mainHwnd, uint officePid, uint runnerPid)
    {
        if (h == mainHwnd || !IsWindowVisible(h)) return false;
        uint p;
        GetWindowThreadProcessId(h, out p);
        if (p != officePid && p != runnerPid) return false;

        var clsSb = new StringBuilder(128);
        GetClassName(h, clsSb, 128);
        string cls = clsSb.ToString();
        if (cls.Contains("InputIndicator") || cls.Contains("IME") || cls.Contains("Tooltip") ||
            cls.Contains("MSCTFIME") || cls.Contains("ConsoleWindowClass") || cls.Contains("PseudoConsole"))
            return false;

        var titleSb = new StringBuilder(256);
        GetWindowText(h, titleSb, 256);
        string title = titleSb.ToString();

        // Only include runner-owned windows if they are our actual add-in UI dialogs/viewers
        if (p == runnerPid && p != officePid)
        {
            if (title.IndexOf("Ketcher", StringComparison.OrdinalIgnoreCase) < 0 &&
                title.IndexOf("Mol*", StringComparison.OrdinalIgnoreCase) < 0 &&
                title.IndexOf("Molecule Info", StringComparison.OrdinalIgnoreCase) < 0 &&
                title.IndexOf("func_SMILES2IMG", StringComparison.OrdinalIgnoreCase) < 0 &&
                title.IndexOf("Diagnostics", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }
        }

        RECT r;
        if (!GetWindowRect(h, out r)) return false;
        int w = r.R - r.L, ht = r.B - r.T;
        return w >= 120 && ht >= 80;
    }

    public static void CaptureCompositeFrame(
        IntPtr mainHwnd,
        uint officePid,
        uint runnerPid,
        int cursorX,
        int cursorY,
        float ripple,
        string outPngPath)
    {
        RECT mr;
        if (!GetWindowRect(mainHwnd, out mr)) return;
        int mw = ((mr.R - mr.L) / 2) * 2;
        int mh = ((mr.B - mr.T) / 2) * 2;
        if (mw <= 0 || mh <= 0) return;

        using (var bmp = new Bitmap(mw, mh))
        using (var g = Graphics.FromImage(bmp))
        {
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            var hdc = g.GetHdc();
            PrintWindow(mainHwnd, hdc, 2);
            g.ReleaseHdc(hdc);
            SanitizeTitleBar(bmp);

            var overlays = new List<IntPtr>();
            EnumWindows((h, _) =>
            {
                if (IsCapturableChildWindow(h, mainHwnd, officePid, runnerPid))
                {
                    overlays.Add(h);
                }
                return true;
            }, IntPtr.Zero);

            for (int i = overlays.Count - 1; i >= 0; i--)
            {
                IntPtr oh = overlays[i];
                RECT or;
                GetWindowRect(oh, out or);
                int ow = or.R - or.L, ohh = or.B - or.T;
                if (ow <= 0 || ohh <= 0) continue;

                using (var obmp = new Bitmap(ow, ohh))
                {
                    using (var og = Graphics.FromImage(obmp))
                    {
                        var ohdc = og.GetHdc();
                        PrintWindow(oh, ohdc, 2);
                        og.ReleaseHdc(ohdc);
                    }

                    int drawW = ow, drawH = ohh;
                    int drawX = or.L - mr.L;
                    int drawY = or.T - mr.T;

                    // If an editor window (like Ketcher/Mol* 2666x1702) exceeds mainHwnd, fit & center cleanly
                    if (drawW > mw - 60 || drawH > mh - 60)
                    {
                        double scale = Math.Min((mw - 80.0) / drawW, (mh - 80.0) / drawH);
                        drawW = (int)(drawW * scale);
                        drawH = (int)(drawH * scale);
                        drawX = (mw - drawW) / 2;
                        drawY = (mh - drawH) / 2;
                    }
                    else if (drawX < 0 || drawY < 0 || drawX + drawW > mw || drawY + drawH > mh)
                    {
                        drawX = Math.Max(20, Math.Min(drawX, mw - drawW - 20));
                        drawY = Math.Max(20, Math.Min(drawY, mh - drawH - 20));
                    }

                    // Subtle drop shadow under modal/popup
                    using (var sh = new SolidBrush(Color.FromArgb(65, 15, 23, 42)))
                    {
                        g.FillRectangle(sh, drawX + 6, drawY + 8, drawW, drawH);
                    }
                    g.DrawImage(obmp, drawX, drawY, drawW, drawH);
                }
            }

            DrawCursor(g, cursorX, cursorY, ripple);

            var dir = Path.GetDirectoryName(outPngPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            bmp.Save(outPngPath, ImageFormat.Png);
        }
    }

    public static void VerifyImageDiverse(string path, int minW, int minH)
    {
        using (var bmp = new Bitmap(path))
        {
            if (bmp.Width < minW || bmp.Height < minH)
                throw new InvalidOperationException("Captured image too small: " + path + " (" + bmp.Width + "x" + bmp.Height + ")");
            var colors = new HashSet<int>();
            int sx = Math.Max(1, bmp.Width / 40);
            int sy = Math.Max(1, bmp.Height / 40);
            for (int y = 0; y < bmp.Height; y += sy)
                for (int x = 0; x < bmp.Width; x += sx)
                    colors.Add(bmp.GetPixel(x, y).ToArgb());
            if (colors.Count < 8)
                throw new InvalidOperationException("Captured image appears blank/uniform (" + colors.Count + " colors): " + path);
        }
    }

    private static volatile bool _recordingModal = false;
    private static ManualResetEvent _modalDone = new ManualResetEvent(true);

    public static void StartModalSnapLoop(
        IntPtr mainHwnd,
        uint officePid,
        uint runnerPid,
        string framesDir,
        int startIndex,
        string autoCloseTitleSubstring,
        int holdMsBeforeClose)
    {
        _recordingModal = true;
        _modalDone.Reset();
        LastModalFrameIndex = startIndex;
        ThreadPool.QueueUserWorkItem(_ =>
        {
            int idx = startIndex;
            bool closed = false;
            var sw = Stopwatch.StartNew();
            IntPtr detectedHwnd = IntPtr.Zero;
            var detectedSw = new Stopwatch();
            try
            {
                while (_recordingModal && sw.ElapsedMilliseconds < 20000)
                {
                    if (detectedHwnd == IntPtr.Zero && !string.IsNullOrEmpty(autoCloseTitleSubstring))
                    {
                        EnumWindows((h, __) =>
                        {
                            if (!IsCapturableChildWindow(h, mainHwnd, officePid, runnerPid)) return true;
                            var sb = new StringBuilder(256);
                            GetWindowText(h, sb, 256);
                            if (sb.ToString().IndexOf(autoCloseTitleSubstring, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                detectedHwnd = h;
                                detectedSw.Start();
                                return false;
                            }
                            return true;
                        }, IntPtr.Zero);
                    }

                    idx++;
                    CaptureCompositeFrame(mainHwnd, officePid, runnerPid, -1, -1, 0f,
                        Path.Combine(framesDir, string.Format("f_{0:D4}.png", idx)));
                    LastModalFrameIndex = idx;

                    if (detectedHwnd != IntPtr.Zero && !closed && detectedSw.ElapsedMilliseconds >= holdMsBeforeClose)
                    {
                        closed = true;
                        PostMessage(detectedHwnd, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
                    }
                    Thread.Sleep(85);
                }
            }
            finally
            {
                LastModalFrameIndex = idx;
                _modalDone.Set();
            }
        });
    }

    public static int LastModalFrameIndex = 0;

    public static int StopModalSnapLoop()
    {
        _recordingModal = false;
        _modalDone.WaitOne(2500);
        return LastModalFrameIndex;
    }
}
