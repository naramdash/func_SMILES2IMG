using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

// SMILES2IMG.FLOAT places a picture over the formula cell for Excel builds without
// in-cell images (2016/2019/2021). The pictures themselves are the only state:
// Title holds "SMILES2IMG.FLOAT|<cache key>|<aspect ratio>" (it survives copy/paste,
// unlike the shape name), and a picture belongs to the cell under its top-left corner.
// Placement is xlMoveAndSize, so Excel moves/stretches the picture with its cell;
// the aspect ratio is restored on the next recalculation or via the Refit command.
internal static class FloatingPictures
{
    private const string TitleMarker = "SMILES2IMG.FLOAT|";
    private const string NamePrefix = "SMILES2IMG.FLOAT_";
    private const double Padding = 2.0;
    private const int RenderCacheLimit = 256;
    private static readonly Regex FormulaPattern = new(@"SMILES2IMG\.FLOAT\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Dictionary<string, byte[]> Rendered = new();
    private static Excel.Application? events;

    // Rendering also validates the SMILES, so the worksheet function calls this on every
    // calculation; the cache keeps repeated recalculations cheap.
    internal static byte[] Render(string key, Func<byte[]> render)
    {
        lock (Rendered)
            if (Rendered.TryGetValue(key, out var cached)) return cached;

        var png = render();
        lock (Rendered)
        {
            if (Rendered.Count >= RenderCacheLimit) Rendered.Clear();
            Rendered[key] = png;
        }
        return png;
    }

    // Largest rectangle with the given aspect ratio centered inside the cell, minus padding.
    internal static (double Left, double Top, double Width, double Height) Fit(double left, double top, double width, double height, double aspect)
    {
        if (!(aspect > 0) || double.IsInfinity(aspect)) aspect = 1;
        var w = Math.Max(width - 2 * Padding, 1);
        var h = Math.Max(height - 2 * Padding, 1);
        if (w / h > aspect) w = h * aspect;
        else h = w / aspect;
        return (left + (width - w) / 2, top + (height - h) / 2, w, h);
    }

    internal static double Aspect(byte[] png)
    {
        if (png.Length < 24) return 1;
        uint width = ((uint)png[16] << 24) | ((uint)png[17] << 16) | ((uint)png[18] << 8) | png[19];
        uint height = ((uint)png[20] << 24) | ((uint)png[21] << 16) | ((uint)png[22] << 8) | png[23];
        return width > 0 && height > 0 ? (double)width / height : 1;
    }

    internal static string Title(string key, double aspect) =>
        $"{TitleMarker}{key}|{aspect.ToString("R", CultureInfo.InvariantCulture)}";

    internal static bool TryParseTitle(string? title, out string key, out double aspect)
    {
        key = "";
        aspect = 1;
        if (title == null || !title.StartsWith(TitleMarker, StringComparison.Ordinal)) return false;
        var body = title.Substring(TitleMarker.Length);
        var separator = body.LastIndexOf('|');
        if (separator <= 0) return false;
        key = body.Substring(0, separator);
        if (!double.TryParse(body.Substring(separator + 1), NumberStyles.Float, CultureInfo.InvariantCulture, out aspect) || !(aspect > 0))
            aspect = 1;
        return true;
    }

    internal static bool HasFormula(Excel.Range cell) =>
        cell.Formula is string formula && formula.TrimStart().StartsWith("=") && FormulaPattern.IsMatch(formula);

    // Pictures owned by the add-in on a sheet, grouped by the (row, column) of their anchor cell.
    internal static Dictionary<(int Row, int Column), List<Excel.Shape>> Index(Excel.Worksheet sheet)
    {
        var index = new Dictionary<(int, int), List<Excel.Shape>>();
        foreach (var shape in Owned(sheet))
        {
            var anchor = Anchor(shape);
            var position = (anchor.Row, anchor.Column);
            if (!index.TryGetValue(position, out var list)) index[position] = list = new List<Excel.Shape>();
            list.Add(shape);
        }
        return index;
    }

    // Called only from a queued macro. key == null removes the cell's picture.
    internal static void Sync(Excel.Worksheet sheet, Excel.Range cell, string? key, string? smiles, byte[]? png,
        Dictionary<(int Row, int Column), List<Excel.Shape>> index)
    {
        var area = Area(cell);
        var anchor = (Excel.Range)area.Cells[1, 1];
        var position = (anchor.Row, anchor.Column);
        if (!index.TryGetValue(position, out var shapes)) index[position] = shapes = new List<Excel.Shape>();
        if (!HasFormula(anchor)) key = null; // the formula was deleted or replaced meanwhile

        Excel.Shape? keep = null;
        var keepAspect = 1.0;
        foreach (var shape in shapes.ToArray())
        {
            if (keep == null && key != null && TryParseTitle(shape.Title, out var shapeKey, out var aspect) && shapeKey == key)
            {
                keep = shape;
                keepAspect = aspect;
                continue;
            }
            shape.Delete();
            shapes.Remove(shape);
        }
        if (key == null || png == null) return;

        if (keep != null)
        {
            FitTo(keep, area, keepAspect);
            return;
        }
        var added = Add(sheet, area, key, smiles ?? "", png);
        if (added != null) shapes.Add(added);
    }

    internal static void RefitActiveSheet()
    {
        var application = (Excel.Application)ExcelDnaUtil.Application;
        if (application.ActiveSheet is not Excel.Worksheet sheet) return;
        var screenUpdating = application.ScreenUpdating;
        try
        {
            application.ScreenUpdating = false;
            foreach (var shape in Owned(sheet))
            {
                try
                {
                    TryParseTitle(shape.Title, out _, out var aspect);
                    FitTo(shape, Area(Anchor(shape)), aspect);
                }
                catch { }
            }
        }
        finally { application.ScreenUpdating = screenUpdating; }
    }

    // Removes pictures whose formula was cleared or overwritten. Deleting rows or columns
    // already deletes xlMoveAndSize pictures inside them.
    internal static void Attach()
    {
        events = (Excel.Application)ExcelDnaUtil.Application;
        events.SheetChange += OnSheetChange;
    }

    internal static void Detach()
    {
        if (events == null) return;
        events.SheetChange -= OnSheetChange;
        events = null;
    }

    private static void OnSheetChange(object sh, Excel.Range target)
    {
        try
        {
            if (sh is not Excel.Worksheet sheet || sheet.Shapes.Count == 0 || events == null) return;
            foreach (var shape in Owned(sheet))
            {
                try
                {
                    var area = Area(shape.TopLeftCell);
                    var anchor = (Excel.Range)area.Cells[1, 1];
                    if (events.Intersect(area, target) != null && !HasFormula(anchor)) shape.Delete();
                }
                catch { }
            }
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("SMILES2IMG.FLOAT cleanup: " + error.Message);
        }
    }

    private static Excel.Shape? Add(Excel.Worksheet sheet, Excel.Range area, string key, string smiles, byte[] png)
    {
        double width = Convert.ToDouble(area.Width), height = Convert.ToDouble(area.Height);
        if (width < 1 || height < 1) return null; // hidden row/column; placed on a later recalculation
        var aspect = Aspect(png);
        double areaLeft = Convert.ToDouble(area.Left), areaTop = Convert.ToDouble(area.Top);
        var (left, top, w, h) = Fit(areaLeft, areaTop, width, height, aspect);
        var path = Path.Combine(Path.GetTempPath(), "smiles2img-float-" + Guid.NewGuid().ToString("N") + ".png");
        try
        {
            File.WriteAllBytes(path, png);
            var shape = sheet.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, (float)left, (float)top, (float)w, (float)h);
            shape.Name = NamePrefix + Guid.NewGuid().ToString("N").Substring(0, 12);
            shape.Title = Title(key, aspect);
            shape.AlternativeText = smiles;
            shape.LockAspectRatio = MsoTriState.msoTrue;
            shape.Placement = Excel.XlPlacement.xlMoveAndSize;
            return shape;
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    private static void FitTo(Excel.Shape shape, Excel.Range area, double aspect)
    {
        double width = Convert.ToDouble(area.Width), height = Convert.ToDouble(area.Height);
        if (width < 1 || height < 1) return; // hidden: Excel restores the picture when shown again
        double areaLeft = Convert.ToDouble(area.Left), areaTop = Convert.ToDouble(area.Top);
        var (left, top, w, h) = Fit(areaLeft, areaTop, width, height, aspect);
        if (Math.Abs(shape.Left - left) < 0.5 && Math.Abs(shape.Top - top) < 0.5 &&
            Math.Abs(shape.Width - w) < 0.5 && Math.Abs(shape.Height - h) < 0.5) return;
        shape.LockAspectRatio = MsoTriState.msoFalse;
        shape.Width = (float)w;
        shape.Height = (float)h;
        shape.Left = (float)left;
        shape.Top = (float)top;
        shape.LockAspectRatio = MsoTriState.msoTrue;
    }

    private static List<Excel.Shape> Owned(Excel.Worksheet sheet)
    {
        var owned = new List<Excel.Shape>();
        foreach (Excel.Shape shape in sheet.Shapes)
        {
            string? title;
            try { title = shape.Title; }
            catch { continue; } // some control types do not expose Title
            if (title != null && title.StartsWith(TitleMarker, StringComparison.Ordinal)) owned.Add(shape);
        }
        return owned;
    }

    private static Excel.Range Anchor(Excel.Shape shape) => (Excel.Range)Area(shape.TopLeftCell).Cells[1, 1];

    private static Excel.Range Area(Excel.Range cell) => cell.MergeCells is true ? cell.MergeArea : cell;
}
