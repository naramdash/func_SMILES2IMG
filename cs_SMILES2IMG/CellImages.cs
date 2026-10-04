using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelDna.Integration;
using Smiles2Img.Rendering;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

// XLLs can return references to rich-value cells. Store the local image on a
// private worksheet and return its reference, leaving the user's formula intact.
// The sheet is the only state: column A holds the image, B a hash of the SMILES
// and color mode (MATCH ignores case and treats * as a wildcard), and C the readable SMILES.
internal static class CellImages
{
    private const string SheetName = "__SMILES2IMG";

    internal static ExcelReference? Find(ExcelReference caller, string smiles, bool color = true, string? transform = null, string? background = null)
    {
        var sheet = Prefix(caller) + SheetName;
        try
        {
            var keys = new ExcelReference(0, 1048575, 1, 1, sheet);
            return XlCall.Excel(XlCall.xlfMatch, Key(smiles, color, transform, background), keys, 0) is double position
                ? new ExcelReference((int)position - 1, (int)position - 1, 0, 0, sheet)
                : null;
        }
        catch (XlCallException) { return null; } // the cache sheet does not exist yet
    }

    // Called only from the queued macro, never during worksheet calculation.
    internal static void Put(ExcelReference caller, Excel.Range cell, string smiles, bool color, string? transform, string? background, byte[] png)
    {
        if (Find(caller, smiles, color, transform, background) != null) return;
        var application = (Excel.Application)ExcelDnaUtil.Application;
        var workbook = (Excel.Workbook)cell.Worksheet.Parent;
        var previousSheet = application.ActiveSheet;
        var previousSelection = application.Selection;
        var screenUpdating = application.ScreenUpdating;
        var enableEvents = application.EnableEvents;
        Excel.Worksheet? store = null;
        try
        {
            application.ScreenUpdating = false;
            application.EnableEvents = false;
            store = GetStore(workbook);
            store.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            workbook.Activate();
            store.Activate();
            var row = ((Excel.Range)store.Cells[store.Rows.Count, 2]).End[Excel.XlDirection.xlUp].Row + 1;
            var target = (Excel.Range)store.Cells[row, 1];
            NativeImageWorkbook.CopyTo(target, png);
            if (target.Value2 == null || target.Value2 is string value && value.Length == 0)
                throw new InvalidOperationException("Excel did not create an in-cell image.");
            store.Cells[row, 2] = Key(smiles, color, transform, background);
            store.Cells[row, 3] = smiles;
        }
        finally
        {
            try
            {
                if (previousSheet != null) ((dynamic)previousSheet).Activate();
                else cell.Worksheet.Activate();
                if (previousSelection != null) ((dynamic)previousSelection).Select();
                if (store != null) store.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            }
            finally
            {
                application.EnableEvents = enableEvents;
                application.ScreenUpdating = screenUpdating;
            }
        }
    }

    private static Excel.Worksheet GetStore(Excel.Workbook workbook)
    {
        var store = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == SheetName);
        if (store != null) return store;
        store = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
        store.Name = SheetName;
        store.Cells[1, 1] = "Image";
        store.Cells[1, 2] = "Key";
        store.Cells[1, 3] = "SMILES";
        ((Excel.Range)store.Columns[3]).NumberFormat = "@";
        return store;
    }

    // "[Book.xlsx]" of the calling cell's workbook.
    private static string Prefix(ExcelReference caller)
    {
        var sheet = (string)XlCall.Excel(XlCall.xlSheetNm, caller);
        return sheet.Substring(0, sheet.IndexOf(']') + 1);
    }

    private static string Key(string smiles, bool color = true, string? transform = null, string? background = null)
    {
        using var sha = SHA1.Create();
        var norm = string.IsNullOrEmpty(transform) ? "" : "_" + transform!.Trim().ToLowerInvariant();
        var bg = string.IsNullOrEmpty(background) ? "" : "_BG_" + background!.Trim().ToLowerInvariant();
        var prefix = (color ? "C" : "BW") + bg + norm + "_";
        return prefix + BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(smiles))).Replace("-", "");
    }
}
