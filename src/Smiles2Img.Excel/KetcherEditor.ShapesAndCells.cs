using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using com.epam.indigo;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Microsoft.Office.Core;
using Smiles2Img.Rendering;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

/// <summary>
/// Handles Excel worksheet formula cell resolution, KET canvas shape creation/replacement,
/// and structure/SMILES extraction.
/// </summary>
public static partial class KetcherEditor
{
    private static readonly Regex FormulaRefPattern = new(
        @"(?:SMILES2IMG|KET2IMG)(?:\.FLOAT)?\s*\(\s*(?:""(?<literal>(?:""""|[^""])*)""|(?<ref>(?:'[^']+'!)?[^,)\s]+))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static void EditSelectedMolecule(Excel.Application app)
    {
        var shape = GetSelectedShape(app);
        if (shape != null)
        {
            var structStr = ExtractStructure(shape);
            if (!string.IsNullOrEmpty(structStr))
            {
                EditShape(app, shape);
                return;
            }

            MessageBox.Show(
                "The selected object is not a SMILES or KET molecular structure image.",
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        if (app.ActiveCell is Excel.Range cell)
        {
            var sheet = (Excel.Worksheet)app.ActiveSheet;
            OpenEditor(sheet, cell.Address, preferKet: null);
        }
    }

    public static void EditShape(Excel.Application app, Excel.Shape targetShape)
    {
        try
        {
            var currentStructure = ExtractStructure(targetShape);
            if (string.IsNullOrEmpty(currentStructure)) return;

            using var form = new KetcherForm("Canvas", currentStructure ?? "");
            var result = form.ShowDialog(new Win32Window(ExcelDnaUtil.WindowHandle));

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultKet) || !string.IsNullOrWhiteSpace(form.ResultSmiles)))
            {
                var newKet = form.ResultKet;
                var newSmiles = form.ResultSmiles?.Trim() ?? "";
                var newPngBytes = StructureMetadata.ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;

                ReplaceShape(targetShape, newKet, newSmiles, newPngBytes);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while editing structure:\n" + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void InsertCanvasMolecule(Excel.Application app)
    {
        try
        {
            if (app.ActiveSheet is not Excel.Worksheet sheet) return;

            using var form = new KetcherForm("Canvas", "");
            var result = form.ShowDialog(new Win32Window(ExcelDnaUtil.WindowHandle));

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultKet) || !string.IsNullOrWhiteSpace(form.ResultSmiles)))
            {
                var newKet = form.ResultKet;
                var newSmiles = form.ResultSmiles?.Trim() ?? "";
                var newPngBytes = StructureMetadata.ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;

                InsertCanvasShapeOnSheet(app, sheet, newKet, newSmiles, newPngBytes);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while inserting canvas structure:\n" + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static void InsertCanvasShapeOnSheet(
        Excel.Application app, Excel.Worksheet sheet, string? ket, string? smiles, byte[] pngBytes, Excel.Range? anchorCell = null)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);

            float left = 100f;
            float top = 100f;
            var targetCell = anchorCell ?? (app.ActiveCell as Excel.Range);
            if (targetCell != null)
            {
                left = (float)Convert.ToDouble(targetCell.Left);
                top = (float)Convert.ToDouble(targetCell.Top);
            }

            var (width, height) = StructureMetadata.GetScaledDimensions(pngBytes, 300f, 200f, minWidth: 200f, maxWidth: 400f);

            var shape = sheet.Shapes.AddPicture(
                Filename: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: width,
                Height: height);

            shape.AlternativeText = StructureMetadata.FormatAlternativeText(ket, smiles);
            shape.Title = StructureMetadata.FormatTitle(smiles, "Molecule Canvas");
            shape.Placement = Excel.XlPlacement.xlMoveAndSize;
            shape.LockAspectRatio = MsoTriState.msoTrue;
            shape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    public static void ReplaceShape(Excel.Shape oldShape, string? newKet, string? newSmiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);

            var sheet = (Excel.Worksheet)oldShape.Parent;
            var oldPlacement = oldShape.Placement;
            var (left, top, newWidth, newHeight) = StructureMetadata.FitReplacementBounds(
                (float)oldShape.Left, (float)oldShape.Top, (float)oldShape.Width, (float)oldShape.Height, pngBytes);

            oldShape.Delete();

            var newShape = sheet.Shapes.AddPicture(
                Filename: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: newWidth,
                Height: newHeight);

            newShape.AlternativeText = StructureMetadata.FormatAlternativeText(newKet, newSmiles);
            newShape.Title = StructureMetadata.FormatTitle(newSmiles);
            newShape.Placement = oldPlacement;
            newShape.LockAspectRatio = MsoTriState.msoTrue;
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    public static string? ExtractStructure(Excel.Shape shape)
    {
        string? alt = null;
        string? title = null;
        try { alt = shape.AlternativeText; } catch { }
        try { title = shape.Title; } catch { }

        var extracted = StructureMetadata.ExtractStructure(alt, title);
        if (!string.IsNullOrWhiteSpace(extracted))
            return extracted;

        if (!string.IsNullOrWhiteSpace(title) && FloatingPictures.TryParseTitle(title, out var key, out _) && IsLikelySmiles(key))
            return key;

        if (!string.IsNullOrWhiteSpace(alt) && IsLikelySmiles(alt!.Trim()))
            return alt.Trim();

        return null;
    }

    public static string? ExtractSmiles(Excel.Shape shape)
    {
        string? title = null;
        string? alt = null;
        try { title = shape.Title; } catch { }
        try { alt = shape.AlternativeText; } catch { }

        var extracted = StructureMetadata.ExtractSmiles(title, alt);
        if (!string.IsNullOrWhiteSpace(extracted))
            return extracted;

        if (!string.IsNullOrWhiteSpace(title) && FloatingPictures.TryParseTitle(title, out var key, out _) && IsLikelySmiles(key))
            return key;

        if (!string.IsNullOrWhiteSpace(alt) && IsLikelySmiles(alt!.Trim()))
            return alt.Trim();

        return null;
    }

    public static string? ExtractSmilesFromCell(Excel.Range cell)
    {
        try
        {
            string formula = Convert.ToString((object?)cell.Formula) ?? "";
            string val = Convert.ToString((object?)cell.Value2) ?? "";

            if (!string.IsNullOrEmpty(formula) &&
                (formula.IndexOf("SMILES2IMG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 formula.IndexOf("KET2IMG", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                var match = FormulaRefPattern.Match(formula);
                if (match.Success)
                {
                    if (match.Groups["literal"].Success)
                        return match.Groups["literal"].Value.Replace("\"\"", "\"").Trim();

                    if (match.Groups["ref"].Success)
                    {
                        var refAddr = match.Groups["ref"].Value;
                        try
                        {
                            var sheet = (Excel.Worksheet)cell.Parent;
                            var app = (Excel.Application)sheet.Application;
                            Excel.Range refCell = refAddr.Contains("!") ? (Excel.Range)app.Range[refAddr] : (Excel.Range)sheet.Range[refAddr];
                            return (Convert.ToString(refCell.Value2) ?? "").Trim();
                        }
                        catch { }
                    }
                }
            }

            if (IsLikelySmiles(val))
                return val.Trim();
        }
        catch { }

        return null;
    }

    internal static Excel.Shape? GetSelectedShape(Excel.Application app)
    {
        try
        {
            var sel = app.Selection;
            if (sel == null || sel is Excel.Range) return null;

            if (sel is Excel.ShapeRange sr && sr.Count > 0) return sr.Item(1);
            if (sel is Excel.Shape s) return s;

            try
            {
                var shapeRangeProp = sel!.GetType().InvokeMember("ShapeRange", System.Reflection.BindingFlags.GetProperty, null, sel, null) as Excel.ShapeRange;
                if (shapeRangeProp != null && shapeRangeProp.Count > 0)
                    return shapeRangeProp.Item(1);
            }
            catch { }

            try
            {
                var nameVal = sel!.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, sel, null);
                string? name = Convert.ToString(nameVal);
                if (!string.IsNullOrEmpty(name) && app.ActiveSheet is Excel.Worksheet ws)
                    return ws.Shapes.Item(name!);
            }
            catch { }
        }
        catch { }
        return null;
    }

    private static void OnSheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
    {
        try
        {
            var cell = (Excel.Range)target.Cells[1, 1];
            string formula = Convert.ToString((object?)cell.Formula) ?? "";

            bool isStructureCell = !string.IsNullOrEmpty(formula) &&
                                   (formula.IndexOf("SMILES2IMG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    formula.IndexOf("KET2IMG", StringComparison.OrdinalIgnoreCase) >= 0);

            if (isStructureCell)
            {
                cancel = true;
                var address = cell.Address;
                var sheet = (Excel.Worksheet)sh;
                ExcelAsyncUtil.QueueAsMacro(() => OpenEditor(sheet, address, preferKet: null));
            }
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher double-click error: " + ex.Message);
        }
    }

    internal static void OpenEditor(Excel.Worksheet sheet, string address, bool? preferKet = false)
    {
        try
        {
            var cell = (Excel.Range)sheet.Range[address];
            string formula = Convert.ToString((object?)cell.Formula) ?? "";
            string val = Convert.ToString((object?)cell.Value2) ?? "";

            var targetCellToWrite = cell;
            var initialSmiles = "";
            var isFormulaLiteral = false;
            bool isKetFunction = preferKet.HasValue
                ? preferKet.Value
                : formula.IndexOf("KET2IMG", StringComparison.OrdinalIgnoreCase) >= 0;

            if (!string.IsNullOrEmpty(formula) &&
                (formula.IndexOf("SMILES2IMG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 formula.IndexOf("KET2IMG", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                var match = FormulaRefPattern.Match(formula);
                if (match.Success)
                {
                    if (match.Groups["ref"].Success)
                    {
                        var refAddr = match.Groups["ref"].Value;
                        try
                        {
                            var app = (Excel.Application)sheet.Application;
                            Excel.Range refCell = refAddr.Contains("!")
                                ? (Excel.Range)app.Range[refAddr]
                                : (Excel.Range)sheet.Range[refAddr];
                            targetCellToWrite = refCell;
                            initialSmiles = Convert.ToString(refCell.Value2) ?? "";
                        }
                        catch
                        {
                            targetCellToWrite = cell;
                        }
                    }
                    else if (match.Groups["literal"].Success)
                    {
                        isFormulaLiteral = true;
                        initialSmiles = match.Groups["literal"].Value.Replace("\"\"", "\"");
                    }
                }
            }
            else
            {
                initialSmiles = val;
            }

            using var form = new KetcherForm(targetCellToWrite.Address, initialSmiles);
            var result = form.ShowDialog(new Win32Window(ExcelDnaUtil.WindowHandle));

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultSmiles) || !string.IsNullOrWhiteSpace(form.ResultKet)))
            {
                var newSmiles = form.ResultSmiles?.Trim();
                var newKet = form.ResultKet?.Trim();
                if (!string.IsNullOrEmpty(newKet))
                {
                    newKet = StructureMetadata.CompactKet(newKet);
                }

                // If user selected SMILES mode but Ketcher couldn't generate SMILES (e.g. multi-component reaction canvas), fallback to KET
                if (!isKetFunction && string.IsNullOrEmpty(newSmiles) && !string.IsNullOrEmpty(newKet))
                {
                    isKetFunction = true;
                }

                var dataToWrite = isKetFunction
                    ? (!string.IsNullOrEmpty(newKet) ? newKet : newSmiles)
                    : newSmiles;

                if (string.IsNullOrEmpty(dataToWrite)) return;

                var funcName = isKetFunction
                    ? (CellImages.NativeImagesUnsupported ? "KET2IMG.FLOAT" : "KET2IMG")
                    : (CellImages.NativeImagesUnsupported ? "SMILES2IMG.FLOAT" : "SMILES2IMG");

                if (targetCellToWrite != cell)
                {
                    targetCellToWrite.Value2 = dataToWrite;
                    cell.Formula = $"={funcName}({targetCellToWrite.Address})";
                }
                else if (isFormulaLiteral)
                {
                    var escaped = dataToWrite!.Replace("\"", "\"\"");
                    if (escaped.Length > 8000)
                    {
                        MessageBox.Show(
                            "The molecular structure exceeds Excel's 8,192-character formula limit.\nConsider simplifying the structure or referencing another cell.",
                            "SMILES for Excel",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }
                    var updatedFormula = FormulaRefPattern.Replace(formula, m =>
                    {
                        return $"{funcName}(\"{escaped}\"";
                    }, 1);
                    cell.Formula = updatedFormula;
                }
                else
                {
                    var escaped = dataToWrite!.Replace("\"", "\"\"");
                    if (escaped.Length > 8000)
                    {
                        MessageBox.Show(
                            "The molecular structure exceeds Excel's 8,192-character formula limit.\nConsider simplifying the structure or referencing another cell.",
                            "SMILES for Excel",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    try
                    {
                        if (string.IsNullOrEmpty(formula) && targetCellToWrite == cell)
                        {
                            if (Convert.ToDouble(cell.RowHeight) < 85) cell.RowHeight = 95;
                            if (Convert.ToDouble(cell.ColumnWidth) < 16) cell.ColumnWidth = 20;
                        }
                    }
                    catch { }

                    cell.Formula = $"={funcName}(\"{escaped}\")";
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while launching editor: " + ex.Message,
                "SMILES2IMG Ketcher",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    internal static bool IsLikelySmiles(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        text = text.Trim();
        if (text.Length < 1 || text.Length > 2000) return false;
        if (double.TryParse(text, out _)) return false;
        if (!Regex.IsMatch(text, @"[CNOFPSKBrIcnops]")) return false;

        try
        {
            MoleculeRenderer.Normalize(text);
            using var indigo = new Indigo();
            using var mol = indigo.loadMolecule(text);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
