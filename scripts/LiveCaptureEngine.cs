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
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetDpiForSystem();
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

    public static Size GetWindowSize(IntPtr hwnd)
    {
        RECT mr;
        if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out mr)) return Size.Empty;
        return new Size(Math.Max(0, mr.R - mr.L), Math.Max(0, mr.B - mr.T));
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
        try
        {
            using (var tmp = new Bitmap(16, 16))
            using (var g = Graphics.FromImage(tmp))
            {
                var hdc = g.GetHdc();
                PrintWindow(hwnd, hdc, 0);
                g.ReleaseHdc(hdc);
            }
        }
        catch { }
    }

    public static Size SetMainWindowSizeLogical(IntPtr hwnd, int logicalX, int logicalY, int logicalW, int logicalH)
    {
        ShowWindow(hwnd, 9); // SW_RESTORE
        SetWindowPos(hwnd, IntPtr.Zero, logicalX, logicalY, 800, 450, 0x0040);
        Thread.Sleep(60);

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

        SetMainWindowSize(hwnd, physX, physY, physW, physH);
        return new Size(physW, physH);
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
        SanitizeTitleBar(bmp, 2.25);
    }

    public static void SanitizeTitleBar(Bitmap bmp, double dpiScale)
    {
        int w = bmp.Width;
        double s = dpiScale > 0.5 ? dpiScale : 1.0;
        if (w < (int)Math.Round(450 * s) || bmp.Height < (int)Math.Round(60 * s)) return;
        int sampleX = Math.Max(10, w - (int)Math.Round(415 * s));
        int sampleY = Math.Max(4, Math.Min(bmp.Height - 4, (int)Math.Round(20 * s)));
        Color sample = bmp.GetPixel(sampleX, sampleY);
        using (var br = new SolidBrush(sample))
        using (var g = Graphics.FromImage(bmp))
        {
            int boxX = Math.Max(0, w - (int)Math.Round(380 * s));
            int boxRight = Math.Max(boxX, w - (int)Math.Round(150 * s));
            int boxW = Math.Max(1, boxRight - boxX);
            int boxY = Math.Max(2, (int)Math.Round(2 * s));
            int boxH = Math.Max(20, (int)Math.Round(38 * s));
            g.FillRectangle(br, boxX, boxY, boxW, boxH);
        }
    }

    private static void DrawCursor(Graphics g, int x, int y, float ripple, float scale)
    {
        if (x < 0 || y < 0) return;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        float sc = scale > 0.25f ? scale : 1.0f;

        if (ripple > 0f && ripple <= 1f)
        {
            float r = (12f + ripple * 34f) * sc;
            int alpha = (int)(220 * (1f - ripple));
            using (var pen = new Pen(Color.FromArgb(alpha, 16, 185, 129), Math.Max(1.5f, 4f * sc)))
            {
                g.DrawEllipse(pen, x - r, y - r, r * 2f, r * 2f);
            }
        }

        PointF[] pts = new[]
        {
            new PointF(x, y),
            new PointF(x, y + 34f * sc),
            new PointF(x + 9f * sc, y + 26f * sc),
            new PointF(x + 17f * sc, y + 40f * sc),
            new PointF(x + 23f * sc, y + 37f * sc),
            new PointF(x + 15f * sc, y + 23f * sc),
            new PointF(x + 26f * sc, y + 23f * sc)
        };
        using (var shadow = new SolidBrush(Color.FromArgb(90, 0, 0, 0)))
        {
            float sh = Math.Max(1f, 2f * sc);
            g.TranslateTransform(sh, sh);
            g.FillPolygon(shadow, pts);
            g.ResetTransform();
        }
        using (var fill = new SolidBrush(Color.FromArgb(255, 15, 23, 42)))
        using (var border = new Pen(Color.White, Math.Max(1.2f, 2.5f * sc)))
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

        // Skip borderless WinForms CopiedToast banners from polluting Office screenshots/videos
        if (cls.IndexOf("WindowsForms10", StringComparison.OrdinalIgnoreCase) >= 0 && string.IsNullOrWhiteSpace(title))
            return false;

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
            SanitizeTitleBar(bmp, GetWindowDpiScale(mainHwnd));

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

            DrawCursor(g, cursorX, cursorY, ripple, Math.Max(0.45f, mw / 2768.0f));

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
    public static int LastActiveModalShotFrameIndex = 0;

    private static bool IsOverlayRendered(IntPtr oh, string titleSub)
    {
        bool isKetcher = !string.IsNullOrEmpty(titleSub) && titleSub.IndexOf("Ketcher", StringComparison.OrdinalIgnoreCase) >= 0;
        bool isMolstar = !string.IsNullOrEmpty(titleSub) && titleSub.IndexOf("Mol*", StringComparison.OrdinalIgnoreCase) >= 0;
        if (!isKetcher && !isMolstar) return true;

        RECT r;
        if (!GetWindowRect(oh, out r)) return false;
        int w = r.R - r.L, ht = r.B - r.T;
        if (w < 200 || ht < 200) return false;

        try
        {
            using (var bmp = new Bitmap(w, ht))
            {
                bool pwOk = false;
                using (var g = Graphics.FromImage(bmp))
                {
                    var hdc = g.GetHdc();
                    pwOk = PrintWindow(oh, hdc, 2);
                    g.ReleaseHdc(hdc);
                }
                if (!pwOk) return false;
                if (isKetcher)
                {
                    // Require BOTH white Ketcher canvas (brightWhite >= 200, rejecting all-black unrendered frames)
                    // AND drawn dark/colored chemical ink pixels (darkInk >= 3 inside central canvas).
                    int brightWhite = 0;
                    int darkInk = 0;
                    for (int iy = 0; iy < 36; iy++)
                    {
                        int py = (int)(ht * (0.20 + 0.60 * iy / 35.0));
                        for (int ix = 0; ix < 36; ix++)
                        {
                            int px = (int)(w * (0.16 + 0.68 * ix / 35.0));
                            var c = bmp.GetPixel(px, py);
                            if (c.R >= 245 && c.G >= 245 && c.B >= 245) brightWhite++;
                            int mn = Math.Min(c.R, Math.Min(c.G, c.B));
                            if (mn < 195) darkInk++;
                        }
                    }
                    return brightWhite >= 200 && darkInk >= 3;
                }
                else
                {
                    // Mol*: sample central 3D viewport (x: 15%..70%, y: 20%..85%) for non-uniform / colored geometry
                    var quantized = new HashSet<int>();
                    int vivid = 0;
                    for (int iy = 0; iy < 14; iy++)
                    {
                        int py = (int)(ht * (0.20 + 0.65 * iy / 13.0));
                        for (int ix = 0; ix < 14; ix++)
                        {
                            int px = (int)(w * (0.15 + 0.55 * ix / 13.0));
                            var c = bmp.GetPixel(px, py);
                            quantized.Add(((c.R >> 4) << 8) | ((c.G >> 4) << 4) | (c.B >> 4));
                            if (Math.Max(c.R, Math.Max(c.G, c.B)) - Math.Min(c.R, Math.Min(c.G, c.B)) > 20)
                                vivid++;
                        }
                    }
                    return quantized.Count >= 8 && vivid >= 2;
                }
            }
        }
        catch
        {
            return true;
        }
    }

    public static void StartModalSnapLoop(
        IntPtr mainHwnd,
        uint officePid,
        uint runnerPid,
        string framesDir,
        int startIndex,
        string autoCloseTitleSubstring,
        int holdMsBeforeClose)
    {
        _modalDone.WaitOne(2500);
        _recordingModal = true;
        _modalDone.Reset();
        LastModalFrameIndex = startIndex;
        LastActiveModalShotFrameIndex = 0;
        ThreadPool.QueueUserWorkItem(_ =>
        {
            int idx = startIndex;
            int firstActiveIdx = 0;
            bool closed = false;
            bool readyToCount = false;
            var sw = Stopwatch.StartNew();
            IntPtr detectedHwnd = IntPtr.Zero;
            var detectedSw = new Stopwatch();
            try
            {
                while (_recordingModal && sw.ElapsedMilliseconds < 32000)
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

                    if (detectedHwnd != IntPtr.Zero)
                    {
                        if (!readyToCount)
                        {
                            if (IsOverlayRendered(detectedHwnd, autoCloseTitleSubstring) || detectedSw.ElapsedMilliseconds >= 13000)
                            {
                                readyToCount = true;
                                detectedSw.Restart();
                            }
                        }

                        if (readyToCount)
                        {
                            idx++;
                            string firstPath = Path.Combine(framesDir, string.Format("f_{0:D4}.png", idx));
                            CaptureCompositeFrame(mainHwnd, officePid, runnerPid, -1, -1, 0f, firstPath);
                            int capturedIdx = idx;
                            for (int dup = 0; dup < 2; dup++)
                            {
                                idx++;
                                File.Copy(firstPath, Path.Combine(framesDir, string.Format("f_{0:D4}.png", idx)), true);
                            }
                            LastModalFrameIndex = idx;
                            if (firstActiveIdx == 0) firstActiveIdx = capturedIdx;
                            if (!closed)
                            {
                                LastActiveModalShotFrameIndex = firstActiveIdx + Math.Max(0, ((capturedIdx - firstActiveIdx) * 2) / 3);
                            }

                            if (!closed && detectedSw.ElapsedMilliseconds >= holdMsBeforeClose)
                            {
                                closed = true;
                                PostMessage(detectedHwnd, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
                            }
                        }
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
