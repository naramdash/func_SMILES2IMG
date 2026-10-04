using System;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

// Excel mutations must run after calculation. A queued update is skipped when the
// formula was deleted or replaced; a stale update only adds an unused cache image.
internal static class CellImageUpdates
{
    private static readonly Regex FormulaPattern = new(@"^=\s*(?:_xll\.)?SMILES2IMG\s*\(", RegexOptions.IgnoreCase);

    internal static void Queue(ExcelReference caller, string smiles, bool color, string? transform, string? background, byte[] png)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (!AddIn.IsOpen) return;
            try
            {
                var application = (Excel.Application)ExcelDnaUtil.Application;
                var address = (string)XlCall.Excel(XlCall.xlfReftext, caller, true);
                var cell = application.Range[address];
                if (cell.Formula is not string formula || !FormulaPattern.IsMatch(formula)) return;

                CellImages.Put(caller, cell, smiles, color, transform, background, png);
                cell.Dirty();
                cell.Calculate();
            }
            catch (Exception error)
            {
                LogDisplay.WriteLine("SMILES2IMG image update: " + error.Message);
            }
        });
    }
}
