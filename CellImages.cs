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
    internal const string UnsupportedMessage = "Needs Excel 2024/365: use SMILES2IMG.FLOAT";

    // Set for the session once Excel shows it cannot hold in-cell images (2016/2019/2021).
    internal static volatile bool NativeImagesUnsupported;

    internal static ExcelReference? Find(ExcelReference caller, string smiles, RenderOptions options)
    {
        var sheet = Prefix(caller) + SheetName;
        try
        {
            var keys = new ExcelReference(0, 1048575, 1, 1, sheet);
            return XlCall.Excel(XlCall.xlfMatch, Key(smiles, options), keys, 0) is double position
                ? new ExcelReference((int)position - 1, (int)position - 1, 0, 0, sheet)
                : null;
        }
        catch (XlCallException) { return null; } // the cache sheet does not exist yet
    }

    // Called only from the queued macro, never during worksheet calculation.
    internal static void Put(ExcelReference caller, Excel.Range cell, string smiles, RenderOptions options, byte[] png)
    {
        if (Find(caller, smiles, options) != null) return;
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

            // Excel 2016/2019/2021 cannot hold the rich image and leaves the copy empty or as
            // the file's #VALUE! fallback (TYPE 16 and no rich data). In that case, rollback the
            // cache entry and flag native images as unsupported for the session.
            var isInvalid = target.Value2 == null ||
                            target.Value2 is string value && value.Length == 0 ||
                            (store.Evaluate("TYPE(A" + row + ")") is double type && type == 16);

            if (isInvalid && !HoldsRichValue(target))
            {
                NativeImagesUnsupported = true;
                target.Clear();
                if (row == 2) // the cache sheet was created for this attempt; do not leave it behind
                {
                    var displayAlerts = application.DisplayAlerts;
                    application.DisplayAlerts = false;
                    try { store.Delete(); store = null; }
                    finally { application.DisplayAlerts = displayAlerts; }
                }
                throw new InvalidOperationException(UnsupportedMessage);
            }

            if (isInvalid)
                throw new InvalidOperationException("Excel did not create an in-cell image.");

            store.Cells[row, 2] = Key(smiles, options);
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

    // Late-bound: Range.HasRichDataType is missing from older Excel object models.
    private static bool HoldsRichValue(Excel.Range cell)
    {
        try { return ((dynamic)cell).HasRichDataType is true; }
        catch { return false; }
    }

    // "[Book.xlsx]" of the calling cell's workbook.
    private static string Prefix(ExcelReference caller)
    {
        var sheet = (string)XlCall.Excel(XlCall.xlSheetNm, caller);
        var idx = sheet.IndexOf(']');
        return idx >= 0 ? sheet.Substring(0, idx + 1) : "";
    }

    internal static string Key(string smiles, RenderOptions options)
    {
        using var sha = SHA1.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(smiles));
        var prefix = options.ToKeyString();
        var sb = new StringBuilder(prefix.Length + 1 + hash.Length * 2);
        sb.Append(prefix).Append('_');
        foreach (var b in hash)
            sb.Append(b.ToString("X2"));
        return sb.ToString();
    }
}
