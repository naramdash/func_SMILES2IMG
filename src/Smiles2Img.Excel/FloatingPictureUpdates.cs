using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

// Pictures can only be changed after calculation. Requests are coalesced per cell
// (the latest wins) and applied in one macro, indexing each sheet's pictures once.
internal static class FloatingPictureUpdates
{
    private sealed class Pending(ExcelReference caller, string? key, string? smiles, byte[]? png)
    {
        public ExcelReference Caller { get; } = caller;
        public string? Key { get; } = key;
        public string? Smiles { get; } = smiles;
        public byte[]? Png { get; } = png;
    }

    private static readonly object Gate = new();
    private static Dictionary<string, Pending> pending = new();
    private static bool scheduled;

    internal static void Place(ExcelReference caller, string key, string smiles, byte[] png) => Enqueue(new Pending(caller, key, smiles, png));

    internal static void Remove(ExcelReference caller) => Enqueue(new Pending(caller, null, null, null));

    private static void Enqueue(Pending item)
    {
        lock (Gate)
        {
            var cellKey = $"{item.Caller.SheetId}:{item.Caller.RowFirst}:{item.Caller.ColumnFirst}";
            pending[cellKey] = item;
            if (scheduled) return;
            scheduled = true;
        }
        ExcelAsyncUtil.QueueAsMacro(Flush);
    }

    private static void Flush()
    {
        Dictionary<string, Pending> batch;
        lock (Gate)
        {
            batch = pending;
            pending = new Dictionary<string, Pending>();
            scheduled = false;
        }
        if (!AddIn.IsOpen || batch.Count == 0) return;

        var application = (Excel.Application)ExcelDnaUtil.Application;
        var screenUpdating = application.ScreenUpdating;
        try
        {
            application.ScreenUpdating = false;
            foreach (var group in batch.Values.GroupBy(item => item.Caller.SheetId))
            {
                try
                {
                    var sheet = ResolveWorksheet(application, group.First().Caller);
                    if (sheet == null) continue;
                    var index = FloatingPictures.Index(sheet);
                    foreach (var item in group)
                    {
                        try
                        {
                            var cell = (Excel.Range)sheet.Cells[item.Caller.RowFirst + 1, item.Caller.ColumnFirst + 1];
                            FloatingPictures.Sync(sheet, cell, item.Key, item.Smiles, item.Png, index);
                        }
                        catch (Exception error)
                        {
                            LogDisplay.WriteLine("SMILES2IMG.FLOAT picture update: " + error.Message);
                        }
                    }
                }
                catch (Exception error)
                {
                    LogDisplay.WriteLine("SMILES2IMG.FLOAT sheet update: " + error.Message);
                }
            }
        }
        finally
        {
            application.ScreenUpdating = screenUpdating;
        }
    }

    private static Excel.Worksheet? ResolveWorksheet(Excel.Application application, ExcelReference caller)
    {
        try
        {
            var sheetRef = (string)XlCall.Excel(XlCall.xlSheetNm, caller);
            var closeBracket = sheetRef.IndexOf(']');
            if (closeBracket > 1)
            {
                var bookName = sheetRef.Substring(1, closeBracket - 1);
                var sheetName = sheetRef.Substring(closeBracket + 1);
                try { return (Excel.Worksheet)application.Workbooks[bookName].Worksheets[sheetName]; }
                catch
                {
                    foreach (Excel.Workbook wb in application.Workbooks)
                    {
                        if (string.Equals(wb.Name, bookName, StringComparison.OrdinalIgnoreCase))
                            return (Excel.Worksheet)wb.Worksheets[sheetName];
                    }
                }
            }
        }
        catch { }

        try
        {
            var address = (string)XlCall.Excel(XlCall.xlfReftext, caller, true);
            return application.Range[address].Worksheet;
        }
        catch { return null; }
    }
}
