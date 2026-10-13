using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using com.epam.indigo;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Microsoft.Office.Core;
using Smiles2Img.Rendering;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

public static class KetcherEditor
{
    private const string ContextMenuTag = "SMILES2IMG_CONTEXT_MENU";
    private const string KetPrefix = "KET:";
    private const string SmilesPrefix = "SMILES:";

    private static readonly Regex FormulaRefPattern = new(
        @"SMILES2IMG(?:\.FLOAT)?\s*\(\s*(?:""(?<literal>[^""]*)""|(?<ref>(?:'[^']+'!)?[^,)\s]+))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex SmilesTagRegex = new(
        @"(?:^|\b|\s)SMILES:\s*(?<smiles>[^\r\n]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    private const int WH_MOUSE = 7;
    private const int WM_LBUTTONDBLCLK = 0x0203;

    private static IntPtr hMouseHook = IntPtr.Zero;
    private static HookProc? mouseHookProc;
    private static bool isEditing = false;

    private static Excel.Application? events;

    internal static void Attach()
    {
        try
        {
            events = (Excel.Application)ExcelDnaUtil.Application;
            events.SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;
            // Clean up any legacy CommandBar buttons; modern Excel uses SmilesRibbon <contextMenus>
            RemoveCellContextMenu();

            try
            {
                mouseHookProc = MouseHookCallback;
                hMouseHook = SetWindowsHookEx(WH_MOUSE, mouseHookProc, IntPtr.Zero, GetCurrentThreadId());
            }
            catch { }
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher attach failed: " + ex.Message);
        }
    }

    internal static void Detach()
    {
        try
        {
            if (hMouseHook != IntPtr.Zero)
            {
                try { UnhookWindowsHookEx(hMouseHook); } catch { }
                hMouseHook = IntPtr.Zero;
            }
            mouseHookProc = null;

            RemoveCellContextMenu();
            if (events != null)
            {
                events.SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
                events = null;
            }
        }
        catch { }
    }

    private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDBLCLK && !isEditing && events != null)
        {
            try
            {
                var targetShape = GetSelectedShape(events);
                if (targetShape != null)
                {
                    var structStr = ExtractStructure(targetShape);
                    if (!string.IsNullOrEmpty(structStr))
                    {
                        isEditing = true;
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            try
                            {
                                EditShape(events, targetShape);
                            }
                            finally
                            {
                                isEditing = false;
                            }
                        });
                        return (IntPtr)1; // Suppress Excel's default action (e.g. Picture Format ribbon tab)
                    }
                }
            }
            catch { }
        }
        return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
    }

    private static void RemoveCellContextMenu()
    {
        try
        {
            if (events == null) return;
            foreach (CommandBar bar in events.CommandBars)
            {
                if (bar.Name == "Cell" || bar.Name == "List Range Popup")
                {
                    try
                    {
                        for (int i = bar.Controls.Count; i >= 1; i--)
                        {
                            var ctrl = bar.Controls[i];
                            if (ctrl.Tag == ContextMenuTag)
                            {
                                ctrl.Delete();
                            }
                        }
                    }
                    catch { }
                }
            }
        }
        catch { }
    }

    [ExcelCommand(ShortCut = "^+K", MenuName = "SMILES2IMG", MenuText = "분자 구조식 삽입/편집 (Ctrl+Shift+K)")]
    public static void OpenKetcherCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            if (targetShape != null)
            {
                var structStr = ExtractStructure(targetShape);
                if (!string.IsNullOrEmpty(structStr))
                {
                    EditShape(app, targetShape);
                    return;
                }
            }

            if (app.ActiveCell is not Excel.Range activeCell) return;
            var sheet = (Excel.Worksheet)app.ActiveSheet;
            var address = activeCell.Address;
            OpenEditor(sheet, address);
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher command error: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "캔버스 구조식 삽입 (KET)")]
    public static void InsertCanvasMoleculeCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            InsertCanvasMolecule(app);
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher canvas insert error: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "구조식 편집")]
    public static void EditSelectedMoleculeCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            EditSelectedMolecule(app);
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher edit error: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "SMILES 복사 (추출)")]
    public static void CopySelectedMoleculeSmilesCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? smiles = null;

            if (targetShape != null)
            {
                smiles = ExtractSmiles(targetShape);
            }
            else if (app.ActiveCell is Excel.Range cell)
            {
                smiles = ExtractSmilesFromCell(cell);
            }

            if (string.IsNullOrEmpty(smiles))
            {
                MessageBox.Show(
                    "선택된 개체 또는 셀에서 SMILES 구조식을 찾을 수 없습니다.\n분자 구조식 이미지 또는 수식 셀을 선택해 주세요.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Clipboard.SetText(smiles!);
            MessageBox.Show(
                $"클립보드에 SMILES가 복사되었습니다:\n\n{smiles}",
                "SMILES 추출 완료",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "SMILES 추출 중 오류가 발생했습니다: " + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "캔버스 데이터(KET) 복사")]
    public static void CopySelectedMoleculeKetCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? ket = null;

            if (targetShape != null)
            {
                ket = ExtractStructure(targetShape);
            }

            if (string.IsNullOrEmpty(ket))
            {
                MessageBox.Show(
                    "선택된 이미지에서 캔버스 구조식(KET)을 찾을 수 없습니다.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Clipboard.SetText(ket!);
            MessageBox.Show(
                "클립보드에 캔버스 데이터(KET)가 복사되었습니다.",
                "KET 복사 완료",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "캔버스 데이터 복사 중 오류가 발생했습니다: " + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "분자 정보(화학식·분자량) 복사")]
    public static void CopySelectedMoleculeInfoCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? targetData = null;

            if (targetShape != null)
            {
                targetData = ExtractSmiles(targetShape) ?? ExtractStructure(targetShape);
            }
            else if (app.ActiveCell is Excel.Range cell)
            {
                targetData = ExtractSmilesFromCell(cell);
            }

            var info = MoleculeRenderer.TryGetMoleculeInfo(targetData);
            if (info == null)
            {
                MessageBox.Show(
                    "선택된 셀 또는 이미지에서 화학식/분자량 정보를 계산할 수 없습니다.\n먼저 분자 구조식 셀이나 이미지를 선택해 주세요.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Clipboard.SetText(info.FormatClipboardSummary());
            MessageBox.Show(
                $"클립보드에 분자 물성 정보가 복사되었습니다:\n\n{info.FormatDisplayMessage()}",
                "분자 정보 (화학식 · 분자량)",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "분자 정보 계산 중 오류가 발생했습니다: " + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "클립보드 SMILES/KET 즉시 삽입")]
    public static void QuickInsertCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet is not Excel.Worksheet sheet) return;

            string? rawInput = null;
            if (Clipboard.ContainsText())
            {
                rawInput = Clipboard.GetText()?.Trim();
            }

            // Fallback to current cell plain text if clipboard doesn't have valid structure
            if (app.ActiveCell is Excel.Range activeCell)
            {
                string cellFormula = Convert.ToString((object?)activeCell.Formula) ?? "";
                string cellVal = Convert.ToString((object?)activeCell.Value2)?.Trim() ?? "";
                if (!cellFormula.StartsWith("=", StringComparison.Ordinal) && !string.IsNullOrEmpty(cellVal) && IsLikelySmiles(cellVal))
                {
                    if (string.IsNullOrEmpty(rawInput) || !MoleculeRenderer.TryRenderStructure(rawInput, RenderOptions.Default, out _, out _, out _))
                    {
                        rawInput = cellVal;
                    }
                }
            }

            if (!MoleculeRenderer.TryRenderStructure(rawInput, RenderOptions.Default, out var pngBytes, out var ket, out var smiles) || pngBytes == null)
            {
                MessageBox.Show(
                    "클립보드(또는 활성 셀)에서 유효한 SMILES 또는 KET 구조식 데이터를 찾을 수 없습니다.\n먼저 SMILES 문자열(예: CC(=O)Oc1ccccc1C(=O)O)을 복사한 후 실행해 주세요.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // If it's a single molecule SMILES (and not a KET canvas / reaction), insert as cell formula!
            if (string.IsNullOrEmpty(ket) && !string.IsNullOrEmpty(smiles) && !smiles!.Contains(">") && app.ActiveCell is Excel.Range targetCell)
            {
                try
                {
                    if (Convert.ToDouble(targetCell.RowHeight) < 85) targetCell.RowHeight = 95;
                    if (Convert.ToDouble(targetCell.ColumnWidth) < 16) targetCell.ColumnWidth = 20;
                }
                catch { }

                var funcName = CellImages.NativeImagesUnsupported ? "SMILES2IMG.FLOAT" : "SMILES2IMG";
                var escaped = smiles.Replace("\"", "\"\"");
                targetCell.Formula = $"={funcName}(\"{escaped}\")";
                return;
            }

            // Otherwise (KET canvas or reaction), insert as floating canvas shape
            var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
            try
            {
                File.WriteAllBytes(tempFile, pngBytes);
                float left = 100f;
                float top = 100f;
                float width = 300f;
                float height = 200f;

                if (app.ActiveCell is Excel.Range cellPos)
                {
                    left = (float)Convert.ToDouble(cellPos.Left);
                    top = (float)Convert.ToDouble(cellPos.Top);
                }

                try
                {
                    using var ms = new MemoryStream(pngBytes);
                    using var img = System.Drawing.Image.FromStream(ms);
                    float aspect = img.Width / (float)img.Height;
                    width = Math.Max(200f, Math.Min(400f, img.Width));
                    height = width / aspect;
                }
                catch { }

                var shape = sheet.Shapes.AddPicture(
                    Filename: tempFile,
                    LinkToFile: MsoTriState.msoFalse,
                    SaveWithDocument: MsoTriState.msoTrue,
                    Left: left,
                    Top: top,
                    Width: width,
                    Height: height);

                shape.AlternativeText = !string.IsNullOrEmpty(ket) ? KetPrefix + ket : SmilesPrefix + (smiles ?? "");
                shape.Title = !string.IsNullOrEmpty(smiles) ? SmilesPrefix + smiles : "Molecule Canvas";
                shape.Placement = Excel.XlPlacement.xlMoveAndSize;
                shape.LockAspectRatio = MsoTriState.msoTrue;
                shape.Select();
            }
            finally
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "클립보드 구조식 즉시 삽입 중 오류가 발생했습니다: " + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

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
            else
            {
                MessageBox.Show(
                    "선택된 개체는 SMILES/KET 분자 구조식 이미지가 아닙니다.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
        }

        if (app.ActiveCell is Excel.Range cell)
        {
            var sheet = (Excel.Worksheet)app.ActiveSheet;
            OpenEditor(sheet, cell.Address);
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
                var newPngBytes = ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;

                ReplaceShape(targetShape, newKet, newSmiles, newPngBytes);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "분자 구조식 편집 중 오류가 발생했습니다:\n" + ex.Message,
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
                var newPngBytes = ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;

                var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
                try
                {
                    File.WriteAllBytes(tempFile, newPngBytes);

                    float left = 100f;
                    float top = 100f;
                    float width = 300f;
                    float height = 200f;

                    if (app.ActiveCell is Excel.Range cell)
                    {
                        left = (float)Convert.ToDouble(cell.Left);
                        top = (float)Convert.ToDouble(cell.Top);
                    }

                    try
                    {
                        using var ms = new MemoryStream(newPngBytes);
                        using var img = System.Drawing.Image.FromStream(ms);
                        float aspect = img.Width / (float)img.Height;
                        width = Math.Max(200f, Math.Min(400f, img.Width));
                        height = width / aspect;
                    }
                    catch { }

                    var newShape = sheet.Shapes.AddPicture(
                        Filename: tempFile,
                        LinkToFile: MsoTriState.msoFalse,
                        SaveWithDocument: MsoTriState.msoTrue,
                        Left: left,
                        Top: top,
                        Width: width,
                        Height: height);

                    newShape.AlternativeText = !string.IsNullOrEmpty(newKet) ? KetPrefix + newKet : SmilesPrefix + newSmiles;
                    newShape.Title = !string.IsNullOrEmpty(newSmiles) ? SmilesPrefix + newSmiles : "Molecule";
                    newShape.Placement = Excel.XlPlacement.xlMoveAndSize;
                    newShape.LockAspectRatio = MsoTriState.msoTrue;
                    newShape.Select();
                }
                finally
                {
                    try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "캔버스 구조식 삽입 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void ReplaceShape(Excel.Shape oldShape, string? newKet, string? newSmiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);

            var sheet = (Excel.Worksheet)oldShape.Parent;
            float left = (float)oldShape.Left;
            float top = (float)oldShape.Top;
            float oldWidth = (float)oldShape.Width;
            float oldHeight = (float)oldShape.Height;
            var oldPlacement = oldShape.Placement;

            float newWidth = oldWidth;
            float newHeight = oldHeight;
            try
            {
                using var ms = new MemoryStream(pngBytes);
                using var img = System.Drawing.Image.FromStream(ms);
                float aspect = img.Width / (float)img.Height;
                if (oldWidth / oldHeight > aspect)
                {
                    newHeight = oldHeight;
                    newWidth = oldHeight * aspect;
                    left += (oldWidth - newWidth) / 2f;
                }
                else
                {
                    newWidth = oldWidth;
                    newHeight = oldWidth / aspect;
                    top += (oldHeight - newHeight) / 2f;
                }
            }
            catch { }

            oldShape.Delete();

            var newShape = sheet.Shapes.AddPicture(
                Filename: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: newWidth,
                Height: newHeight);

            newShape.AlternativeText = !string.IsNullOrEmpty(newKet) ? KetPrefix + newKet : SmilesPrefix + newSmiles;
            newShape.Title = !string.IsNullOrEmpty(newSmiles) ? SmilesPrefix + newSmiles : "Molecule";
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
        try
        {
            var alt = shape.AlternativeText;
            if (!string.IsNullOrWhiteSpace(alt))
            {
                var trimmed = alt.Trim();
                if (trimmed.StartsWith(KetPrefix, StringComparison.OrdinalIgnoreCase))
                    return trimmed.Substring(KetPrefix.Length).Trim();
                var idx = trimmed.IndexOf(KetPrefix, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                    return trimmed.Substring(idx + KetPrefix.Length).Trim();
            }
        }
        catch { }

        try
        {
            var alt = shape.AlternativeText;
            if (!string.IsNullOrWhiteSpace(alt))
            {
                var trimmed = alt.Trim();
                if (trimmed.StartsWith(SmilesPrefix, StringComparison.OrdinalIgnoreCase))
                    return trimmed.Substring(SmilesPrefix.Length).Trim();
                if (IsLikelySmiles(trimmed))
                    return trimmed;
            }
        }
        catch { }

        try
        {
            var title = shape.Title;
            if (!string.IsNullOrWhiteSpace(title))
            {
                var trimmed = title.Trim();
                if (trimmed.StartsWith(SmilesPrefix, StringComparison.OrdinalIgnoreCase))
                    return trimmed.Substring(SmilesPrefix.Length).Trim();

                if (FloatingPictures.TryParseTitle(title, out var key, out _))
                {
                    if (IsLikelySmiles(key)) return key;
                }

                var match = SmilesTagRegex.Match(trimmed);
                if (match.Success) return match.Groups["smiles"].Value.Trim();
            }
        }
        catch { }

        return null;
    }

    public static string? ExtractSmiles(Excel.Shape shape)
    {
        try
        {
            var title = shape.Title;
            if (!string.IsNullOrWhiteSpace(title))
            {
                var trimmed = title.Trim();
                if (trimmed.StartsWith(SmilesPrefix, StringComparison.OrdinalIgnoreCase))
                    return trimmed.Substring(SmilesPrefix.Length).Trim();

                if (FloatingPictures.TryParseTitle(title, out var key, out _))
                {
                    if (IsLikelySmiles(key)) return key;
                }

                var match = SmilesTagRegex.Match(trimmed);
                if (match.Success) return match.Groups["smiles"].Value.Trim();
            }
        }
        catch { }

        try
        {
            var alt = shape.AlternativeText;
            if (!string.IsNullOrWhiteSpace(alt))
            {
                var trimmed = alt.Trim();
                if (trimmed.StartsWith(SmilesPrefix, StringComparison.OrdinalIgnoreCase))
                    return trimmed.Substring(SmilesPrefix.Length).Trim();
                var match = SmilesTagRegex.Match(trimmed);
                if (match.Success) return match.Groups["smiles"].Value.Trim();
                if (IsLikelySmiles(trimmed))
                    return trimmed;
            }
        }
        catch { }

        return null;
    }

    public static string? ExtractSmilesFromCell(Excel.Range cell)
    {
        try
        {
            string formula = Convert.ToString((object?)cell.Formula) ?? "";
            string val = Convert.ToString((object?)cell.Value2) ?? "";

            if (!string.IsNullOrEmpty(formula) && formula.IndexOf("SMILES2IMG", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var match = FormulaRefPattern.Match(formula);
                if (match.Success)
                {
                    if (match.Groups["literal"].Success)
                    {
                        return match.Groups["literal"].Value.Trim();
                    }
                    if (match.Groups["ref"].Success)
                    {
                        var refAddr = match.Groups["ref"].Value;
                        try
                        {
                            var sheet = (Excel.Worksheet)cell.Parent;
                            var app = (Excel.Application)sheet.Application;
                            Excel.Range refCell = refAddr.Contains("!") ? (Excel.Range)app.Range[refAddr] : (Excel.Range)sheet.Range[refAddr];
                            var refVal = Convert.ToString(refCell.Value2) ?? "";
                            return refVal.Trim();
                        }
                        catch { }
                    }
                }
            }

            if (IsLikelySmiles(val))
            {
                return val.Trim();
            }
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
                {
                    return shapeRangeProp.Item(1);
                }
            }
            catch { }

            try
            {
                var nameVal = sel!.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, sel, null);
                string? name = Convert.ToString(nameVal);
                if (!string.IsNullOrEmpty(name) && app.ActiveSheet is Excel.Worksheet ws)
                {
                    return ws.Shapes.Item(name!);
                }
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

            // 오직 SMILES2IMG 수식이 걸려있는 분자 이미지 셀(Structure Cell)만 더블클릭 감지!
            // 일반 텍스트 셀은 더블클릭해도 엑셀 정상 편집으로 작동 (텍스트 오작동 100% 방지)
            var isStructureCell = !string.IsNullOrEmpty(formula) &&
                                 formula.IndexOf("SMILES2IMG", StringComparison.OrdinalIgnoreCase) >= 0;

            if (isStructureCell)
            {
                cancel = true; // 엑셀 기본 커서 편집 취소
                var address = cell.Address;
                var sheet = (Excel.Worksheet)sh;

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    OpenEditor(sheet, address);
                });
            }
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher double-click error: " + ex.Message);
        }
    }

    internal static void OpenEditor(Excel.Worksheet sheet, string address)
    {
        try
        {
            var cell = (Excel.Range)sheet.Range[address];
            string formula = Convert.ToString((object?)cell.Formula) ?? "";
            string val = Convert.ToString((object?)cell.Value2) ?? "";

            var targetCellToWrite = cell;
            var initialSmiles = "";
            var isFormulaLiteral = false;

            if (!string.IsNullOrEmpty(formula) && formula.IndexOf("SMILES2IMG", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var match = FormulaRefPattern.Match(formula);
                if (match.Success)
                {
                    if (match.Groups["ref"].Success)
                    {
                        var refAddr = match.Groups["ref"].Value;
                        try
                        {
                            Excel.Range refCell;
                            if (refAddr.Contains("!"))
                            {
                                var app = (Excel.Application)sheet.Application;
                                refCell = (Excel.Range)app.Range[refAddr];
                            }
                            else
                            {
                                refCell = (Excel.Range)sheet.Range[refAddr];
                            }
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
                        initialSmiles = match.Groups["literal"].Value;
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
                var newKet = form.ResultKet;

                if (!string.IsNullOrEmpty(newSmiles))
                {
                    if (isFormulaLiteral)
                    {
                        // 기존 리터럴 수식 업데이트
                        var updatedFormula = FormulaRefPattern.Replace(formula, m =>
                        {
                            var prefix = m.Value.Substring(0, m.Value.IndexOf('"'));
                            return prefix + "\"" + newSmiles + "\"";
                        }, 1);
                        cell.Formula = updatedFormula;
                    }
                    else if (targetCellToWrite != cell)
                    {
                        // 참조 셀(=SMILES2IMG(A2))의 원본 갱신
                        targetCellToWrite.Value2 = newSmiles;
                    }
                    else
                    {
                        // 빈 셀에서 시작했거나 단축키/우클릭으로 삽입한 경우:
                        // 텍스트 대신 셀 자체가 분자 구조식 이미지(Structure Cell)가 되도록 수식으로 삽입!
                        var funcName = CellImages.NativeImagesUnsupported ? "SMILES2IMG.FLOAT" : "SMILES2IMG";
                        cell.Formula = $"={funcName}(\"{newSmiles}\")";
                    }
                }
                else if (!string.IsNullOrEmpty(newKet))
                {
                    // User drew reaction / canvas with KET not convertible to single SMILES formula
                    // Insert as floating canvas shape over the cell
                    var newPngBytes = ResolvePngBytes(form, newKet, null);
                    if (newPngBytes != null)
                    {
                        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
                        try
                        {
                            File.WriteAllBytes(tempFile, newPngBytes);
                            float left = (float)Convert.ToDouble(cell.Left);
                            float top = (float)Convert.ToDouble(cell.Top);
                            float width = 300f;
                            float height = 200f;
                            try
                            {
                                using var ms = new MemoryStream(newPngBytes);
                                using var img = System.Drawing.Image.FromStream(ms);
                                float aspect = img.Width / (float)img.Height;
                                width = Math.Max(200f, Math.Min(400f, img.Width));
                                height = width / aspect;
                            }
                            catch { }

                            var shape = sheet.Shapes.AddPicture(
                                Filename: tempFile,
                                LinkToFile: MsoTriState.msoFalse,
                                SaveWithDocument: MsoTriState.msoTrue,
                                Left: left,
                                Top: top,
                                Width: width,
                                Height: height);

                            shape.AlternativeText = KetPrefix + newKet;
                            shape.Title = "Molecule Canvas";
                            shape.Placement = Excel.XlPlacement.xlMoveAndSize;
                            shape.LockAspectRatio = MsoTriState.msoTrue;
                            shape.Select();
                        }
                        finally
                        {
                            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "에디터 실행 중 오류가 발생했습니다: " + ex.Message,
                "SMILES2IMG Ketcher",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static byte[]? ResolvePngBytes(KetcherForm form, string? ket, string? smiles)
    {
        var pngBytes = form.ResultPngBytes;
        if (pngBytes != null && pngBytes.Length > 0) return pngBytes;

        var opts = new RenderOptions { Background = form.ResultBackground };

        if (!string.IsNullOrWhiteSpace(ket))
        {
            try
            {
                pngBytes = MoleculeRenderer.RenderKet(ket!, opts);
                if (pngBytes != null && pngBytes.Length > 0) return pngBytes;
            }
            catch { }
        }

        if (!string.IsNullOrWhiteSpace(smiles))
        {
            try
            {
                pngBytes = MoleculeRenderer.Render(smiles!, opts);
                if (pngBytes != null && pngBytes.Length > 0) return pngBytes;
            }
            catch { }
        }

        return null;
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
