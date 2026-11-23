using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace Smiles2Img.Editor;

/// <summary>
/// Provides high-resolution, vector-rendered WinForms Icons for chemical dialogs,
/// 3D molecular viewers, and Office add-in dialog windows.
/// </summary>
public static class UiIcons
{

    public static Icon? ChemicalIcon => _chemicalIcon.Value;
    public static Icon? Molstar3DIcon => _molstar3DIcon.Value;
    public static Icon? MoleculeInfoIcon => _moleculeInfoIcon.Value;
    public static Icon? AboutIcon => _aboutIcon.Value;

    private static readonly Lazy<Icon?> _chemicalIcon = new(CreateChemicalIcon);
    private static readonly Lazy<Icon?> _molstar3DIcon = new(CreateMolstar3DIcon);
    private static readonly Lazy<Icon?> _moleculeInfoIcon = new(CreateMoleculeInfoIcon);
    private static readonly Lazy<Icon?> _aboutIcon = new(CreateAboutIcon);

    private static GraphicsPath CreateRoundedRectPath(RectangleF bounds, float radius)
    {
        float d = radius * 2f;
        var path = new GraphicsPath();
        path.AddArc(bounds.X, bounds.Y, d, d, 180f, 90f);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270f, 90f);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0f, 90f);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90f, 90f);
        path.CloseFigure();
        return path;
    }

    private static void DrawHexFBrandMark(Graphics g, float size, Color inkColor, float insetPx = 0f)
    {
        float scale = Math.Max(0.1f, (size - insetPx * 2f) / 64f);
        var state = g.Save();
        try
        {
            g.TranslateTransform(insetPx, insetPx);
            g.ScaleTransform(scale, scale);

            PointF[] hex =
            {
                new(32f, 6f),
                new(54.5f, 19f),
                new(54.5f, 45f),
                new(32f, 58f),
                new(9.5f, 45f),
                new(9.5f, 19f)
            };
            using (var hexPen = new Pen(inkColor, 4.2f) { LineJoin = LineJoin.Round })
            {
                g.DrawPolygon(hexPen, hex);
            }

            g.TranslateTransform(32f, 32f);
            g.ScaleTransform(0.72f, 0.72f);

            using var stemPath = new GraphicsPath();
            stemPath.AddBezier(new PointF(9.6f, -15.4f), new PointF(8.2f, -19.2f), new PointF(3.3f, -19.0f), new PointF(2.0f, -11.5f));
            stemPath.AddLine(new PointF(2.0f, -11.5f), new PointF(-2.0f, 11.5f));
            stemPath.AddBezier(new PointF(-2.0f, 11.5f), new PointF(-3.3f, 19.0f), new PointF(-8.2f, 19.2f), new PointF(-9.6f, 15.4f));

            using (var stemPen = new Pen(inkColor, 4.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            {
                g.DrawPath(stemPen, stemPath);
            }

            using (var crossPen = new Pen(inkColor, 4.0f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawLine(crossPen, -6.0f, -2.6f, 6.8f, -2.6f);
            }

            using var dotBrush = new SolidBrush(inkColor);
            g.FillEllipse(dotBrush, 9.6f - 2.3f, -15.4f - 2.3f, 4.6f, 4.6f);
            g.FillEllipse(dotBrush, -9.6f - 2.3f, 15.4f - 2.3f, 4.6f, 4.6f);
        }
        finally
        {
            g.Restore(state);
        }
    }

    /// <summary>
    /// Renders the 1:1 monochrome Hex-f brand mark onto a transparent bitmap at the requested pixel size.
    /// </summary>
    public static Bitmap CreateAppLogoBitmap(int size = 64)
    {
        int px = Math.Max(16, size);
        var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.Clear(Color.Transparent);
        DrawHexFBrandMark(g, px, Color.FromArgb(17, 17, 17), insetPx: 0f);
        return bmp;
    }

    private static Icon? CreateBrandWindowIcon()
    {
        try
        {
            using var bmp = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);

            using (var bgPath = CreateRoundedRectPath(new RectangleF(1f, 1f, 30f, 30f), 7f))
            using (var bgBrush = new SolidBrush(Color.White))
            using (var rimPen = new Pen(Color.FromArgb(212, 212, 216), 1.0f))
            {
                g.FillPath(bgBrush, bgPath);
                g.DrawPath(rimPen, bgPath);
            }

            DrawHexFBrandMark(g, 32f, Color.FromArgb(17, 17, 17), insetPx: 3f);
            return ConvertBitmapToIcon(bmp);
        }
        catch
        {
            return null;
        }
    }

    private static Icon? CreateChemicalIcon() => CreateBrandWindowIcon();

    private static Icon? CreateMolstar3DIcon()
    {
        try
        {
            using var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Outer deep violet-blue background
            using var bgBrush = new SolidBrush(Color.FromArgb(15, 23, 42));
            g.FillEllipse(bgBrush, 1, 1, 30, 30);

            // 3D Isometric Bond Lines
            using var bondPen = new Pen(Color.FromArgb(148, 163, 184), 1.8f);
            PointF center = new(16, 16);
            PointF top = new(16, 5.5f);
            PointF right = new(26, 21.5f);
            PointF left = new(6, 21.5f);

            g.DrawLine(bondPen, center, top);
            g.DrawLine(bondPen, center, right);
            g.DrawLine(bondPen, center, left);

            // Outer 3D bounding triangle bonds
            using var outerPen = new Pen(Color.FromArgb(71, 85, 105), 1.2f) { DashStyle = DashStyle.Dot };
            g.DrawPolygon(outerPen, new[] { top, right, left });

            // 3D Atomic Spheres
            // Top: Violet
            using var violetBrush = new SolidBrush(Color.FromArgb(168, 85, 247));
            g.FillEllipse(violetBrush, top.X - 3.5f, top.Y - 3.5f, 7, 7);

            // Right: Emerald
            using var emeraldBrush = new SolidBrush(Color.FromArgb(16, 185, 129));
            g.FillEllipse(emeraldBrush, right.X - 3.5f, right.Y - 3.5f, 7, 7);

            // Left: Coral/Rose
            using var roseBrush = new SolidBrush(Color.FromArgb(244, 63, 94));
            g.FillEllipse(roseBrush, left.X - 3.5f, left.Y - 3.5f, 7, 7);

            // Center: Cyan / Sky Blue with highlighted core
            using var cyanBrush = new SolidBrush(Color.FromArgb(14, 165, 233));
            g.FillEllipse(cyanBrush, center.X - 4.5f, center.Y - 4.5f, 9, 9);
            using var coreBrush = new SolidBrush(Color.White);
            g.FillEllipse(coreBrush, center.X - 1.5f, center.Y - 2.5f, 3, 3);

            return ConvertBitmapToIcon(bmp);
        }
        catch
        {
            return null;
        }
    }

    private static Icon? CreateMoleculeInfoIcon()
    {
        try
        {
            using var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Outer background
            using var bgBrush = new SolidBrush(Color.FromArgb(15, 23, 42));
            g.FillEllipse(bgBrush, 1, 1, 30, 30);

            // Small chemical structure on the left
            using var chemPen = new Pen(Color.FromArgb(56, 189, 248), 1.6f);
            PointF[] hex =
            {
                new(11, 8),
                new(16, 11),
                new(16, 17),
                new(11, 20),
                new(6, 17),
                new(6, 11)
            };
            g.DrawPolygon(chemPen, hex);

            // Atom node
            using var redBrush = new SolidBrush(Color.FromArgb(239, 68, 68));
            g.FillEllipse(redBrush, 9.5f, 6.5f, 3.5f, 3.5f);

            // Properties / Info Lines on the right
            using var linePen1 = new Pen(Color.FromArgb(245, 158, 11), 2.0f);
            using var linePen2 = new Pen(Color.FromArgb(203, 213, 225), 1.8f);
            using var linePen3 = new Pen(Color.FromArgb(148, 163, 184), 1.8f);

            g.DrawLine(linePen1, 19, 10, 26, 10);
            g.DrawLine(linePen2, 19, 15, 26, 15);
            g.DrawLine(linePen3, 19, 20, 24, 20);

            return ConvertBitmapToIcon(bmp);
        }
        catch
        {
            return null;
        }
    }

    private static Icon? CreateAboutIcon() => CreateBrandWindowIcon();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private static Icon? ConvertBitmapToIcon(Bitmap bmp)
    {
        try
        {
            var hIcon = bmp.GetHicon();
            try
            {
                using var temp = Icon.FromHandle(hIcon);
                return (Icon)temp.Clone();
            }
            finally
            {
                DestroyIcon(hIcon);
            }
        }
        catch
        {
            return null;
        }
    }
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, Bitmap> _propIconCache = new();

    /// <summary>
    /// Returns a clean, modern, high-DPI aware vector-rendered bitmap icon for each molecule info property row.
    /// </summary>
    public static Bitmap GetPropertyIcon(string propertyName, float dpiScale = 1.0f)
    {
        string key = $"{propertyName}_{dpiScale:F2}";
        return _propIconCache.GetOrAdd(key, _ => CreatePropertyBitmap(propertyName, dpiScale));
    }

    private static Bitmap CreatePropertyBitmap(string propertyName, float dpiScale)
    {
        int size = Math.Max(18, (int)Math.Round(20 * dpiScale));
        var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;

        float s = size / 20f; // normalize to 20x20 base canvas

        switch (propertyName)
        {
            case "Substance Name":
            {
                // Blue Label Tag with rounded corner and punch hole
                using var brush = new SolidBrush(Color.FromArgb(37, 99, 235));
                using var path = new GraphicsPath();
                path.AddPolygon(new PointF[]
                {
                    new(3 * s, 4 * s),
                    new(11 * s, 4 * s),
                    new(17 * s, 10 * s),
                    new(11 * s, 16 * s),
                    new(3 * s, 16 * s)
                });
                g.FillPath(brush, path);
                using var holeBrush = new SolidBrush(Color.White);
                g.FillEllipse(holeBrush, 5.5f * s, 8.5f * s, 3 * s, 3 * s);
                break;
            }
            case "IUPAC Name":
            {
                // Indigo Official IUPAC Document with text lines
                using var docBrush = new SolidBrush(Color.FromArgb(79, 70, 229));
                g.FillRectangle(docBrush, 4 * s, 3 * s, 12 * s, 14 * s);
                using var whitePen = new Pen(Color.White, 1.2f * s);
                g.DrawLine(whitePen, 6.5f * s, 6.5f * s, 13.5f * s, 6.5f * s);
                g.DrawLine(whitePen, 6.5f * s, 9.5f * s, 13.5f * s, 9.5f * s);
                g.DrawLine(whitePen, 6.5f * s, 12.5f * s, 11f * s, 12.5f * s);
                break;
            }
            case "Chemical Formula":
            {
                // Emerald Atomic Orbital
                using var orbPen = new Pen(Color.FromArgb(16, 185, 129), 1.3f * s);
                g.DrawEllipse(orbPen, 3 * s, 6 * s, 14 * s, 8 * s);
                var state = g.Save();
                g.TranslateTransform(10 * s, 10 * s);
                g.RotateTransform(60);
                g.TranslateTransform(-10 * s, -10 * s);
                g.DrawEllipse(orbPen, 3 * s, 6 * s, 14 * s, 8 * s);
                g.Restore(state);
                using var nucleus = new SolidBrush(Color.FromArgb(5, 150, 105));
                g.FillEllipse(nucleus, 8.5f * s, 8.5f * s, 3 * s, 3 * s);
                break;
            }
            case "Molecular Weight":
            {
                // Amber Precision Balance / Scale
                using var scalePen = new Pen(Color.FromArgb(217, 119, 6), 1.5f * s);
                g.DrawLine(scalePen, 10 * s, 3 * s, 10 * s, 16 * s); // pillar
                g.DrawLine(scalePen, 4 * s, 6 * s, 16 * s, 6 * s);   // beam
                g.DrawLine(scalePen, 6 * s, 16 * s, 14 * s, 16 * s); // base

                // Pans
                using var panBrush = new SolidBrush(Color.FromArgb(245, 158, 11));
                g.DrawLine(scalePen, 4 * s, 6 * s, 2.5f * s, 11 * s);
                g.DrawLine(scalePen, 4 * s, 6 * s, 5.5f * s, 11 * s);
                g.FillPie(panBrush, 1.5f * s, 9.5f * s, 5 * s, 3.5f * s, 0, 180);

                g.DrawLine(scalePen, 16 * s, 6 * s, 14.5f * s, 11 * s);
                g.DrawLine(scalePen, 16 * s, 6 * s, 17.5f * s, 11 * s);
                g.FillPie(panBrush, 13.5f * s, 9.5f * s, 5 * s, 3.5f * s, 0, 180);
                break;
            }
            case "Monoisotopic Mass":
            {
                // Violet Microscope / Mass Spec Spectrometry
                using var brush = new SolidBrush(Color.FromArgb(124, 58, 237));
                using var pen = new Pen(Color.FromArgb(124, 58, 237), 1.8f * s);
                g.DrawLine(pen, 13 * s, 4 * s, 7 * s, 12 * s); // tube
                g.FillEllipse(brush, 11.5f * s, 2.5f * s, 4 * s, 3 * s); // eyepiece
                g.FillRectangle(brush, 5 * s, 12 * s, 7 * s, 2 * s);     // stage
                g.DrawLine(pen, 9 * s, 12 * s, 9 * s, 16 * s);           // arm
                g.FillRectangle(brush, 4 * s, 15.5f * s, 10 * s, 2 * s); // base
                break;
            }
            case "LogP / TPSA":
            {
                // Cyan / Teal Partition Droplet (Octanol/Water)
                using var path = new GraphicsPath();
                path.AddBezier(10 * s, 2 * s, 16 * s, 10 * s, 16 * s, 16 * s, 10 * s, 17.5f * s);
                path.AddBezier(10 * s, 17.5f * s, 4 * s, 16 * s, 4 * s, 10 * s, 10 * s, 2 * s);
                using var brushTeal = new SolidBrush(Color.FromArgb(14, 165, 233));
                g.FillPath(brushTeal, path);
                using var innerPen = new Pen(Color.White, 1.2f * s);
                g.DrawLine(innerPen, 10 * s, 4 * s, 10 * s, 16 * s);
                break;
            }
            case "Rings & Aromaticity":
            {
                // Royal Blue Bicyclic Aromatic Hexagon + Inner Ring
                using var ringPen = new Pen(Color.FromArgb(37, 99, 235), 1.6f * s);
                PointF[] hex =
                {
                    new(10 * s, 3 * s),
                    new(16 * s, 6.5f * s),
                    new(16 * s, 13.5f * s),
                    new(10 * s, 17 * s),
                    new(4 * s, 13.5f * s),
                    new(4 * s, 6.5f * s)
                };
                g.DrawPolygon(ringPen, hex);
                using var aromPen = new Pen(Color.FromArgb(59, 130, 246), 1.2f * s);
                g.DrawEllipse(aromPen, 6.8f * s, 6.8f * s, 6.4f * s, 6.4f * s);
                break;
            }
            case "3D Complexity (Fsp3)":
            {
                // Violet/Emerald Tetrahedral sp3 Stereocenter
                using var bondPen = new Pen(Color.FromArgb(109, 40, 217), 1.5f * s);
                g.DrawLine(bondPen, 10 * s, 10.5f * s, 10 * s, 3.5f * s);
                g.DrawLine(bondPen, 10 * s, 10.5f * s, 4 * s, 15.5f * s);
                g.DrawLine(bondPen, 10 * s, 10.5f * s, 16 * s, 15.5f * s);
                using var centerBrush = new SolidBrush(Color.FromArgb(16, 185, 129));
                using var nodeBrush = new SolidBrush(Color.FromArgb(139, 92, 246));
                g.FillEllipse(nodeBrush, 8 * s, 2 * s, 4 * s, 4 * s);
                g.FillEllipse(nodeBrush, 2.5f * s, 13.5f * s, 4 * s, 4 * s);
                g.FillEllipse(nodeBrush, 13.5f * s, 13.5f * s, 4 * s, 4 * s);
                g.FillEllipse(centerBrush, 7.5f * s, 8 * s, 5 * s, 5 * s);
                break;
            }
            case "Canonical SMILES":
            {
                // Slate Chemical Chain (Benzene Ring + Double Bond)
                using var bondPen = new Pen(Color.FromArgb(51, 65, 85), 1.6f * s);
                PointF[] pts =
                {
                    new(10 * s, 3.5f * s),
                    new(16 * s, 7 * s),
                    new(16 * s, 13 * s),
                    new(10 * s, 16.5f * s),
                    new(4 * s, 13 * s),
                    new(4 * s, 7 * s)
                };
                g.DrawPolygon(bondPen, pts);
                using var innerPen = new Pen(Color.FromArgb(100, 116, 139), 1.0f * s);
                g.DrawLine(innerPen, 6.5f * s, 8.5f * s, 6.5f * s, 11.5f * s);
                break;
            }
            case "InChI":
            {
                // Teal IUPAC Landmark Column / Badge
                using var badgeBrush = new SolidBrush(Color.FromArgb(13, 148, 136));
                g.FillPolygon(badgeBrush, new PointF[] { new(10 * s, 3 * s), new(16 * s, 6 * s), new(4 * s, 6 * s) });
                g.FillRectangle(badgeBrush, 5 * s, 7 * s, 2 * s, 7 * s);
                g.FillRectangle(badgeBrush, 9 * s, 7 * s, 2 * s, 7 * s);
                g.FillRectangle(badgeBrush, 13 * s, 7 * s, 2 * s, 7 * s);
                g.FillRectangle(badgeBrush, 4 * s, 14.5f * s, 12 * s, 2.5f * s);
                break;
            }
            case "InChIKey":
            {
                // Golden Hash Security Key
                using var keyPen = new Pen(Color.FromArgb(202, 138, 4), 1.8f * s);
                g.DrawEllipse(keyPen, 4 * s, 4 * s, 6 * s, 6 * s);
                g.DrawLine(keyPen, 9 * s, 9 * s, 16 * s, 16 * s); // stem
                g.DrawLine(keyPen, 13 * s, 13 * s, 15 * s, 11 * s); // tooth 1
                g.DrawLine(keyPen, 15 * s, 15 * s, 17 * s, 13 * s); // tooth 2
                break;
            }
            default:
            {
                // Generic property bullet
                using var brush = new SolidBrush(Color.FromArgb(100, 116, 139));
                g.FillEllipse(brush, 6 * s, 6 * s, 8 * s, 8 * s);
                break;
            }
        }

        return bmp;
    }
}
