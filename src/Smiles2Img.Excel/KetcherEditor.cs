using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Microsoft.Office.Core;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using Smiles2Img.Notifications;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

/// <summary>
/// Manages Excel event hooks, keyboard shortcuts, and ribbon/context-menu commands
/// for the embedded EPAM Ketcher molecular editor.
/// </summary>
public static partial class KetcherEditor
{
    private const string ContextMenuTag = "SMILES2IMG_CONTEXT_MENU";

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
    private static bool isEditing;
    private static Excel.Application? events;

    internal static void Attach()
    {
        try
        {
            events = (Excel.Application)ExcelDnaUtil.Application;
            events.SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;

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
                if (targetShape != null && !string.IsNullOrEmpty(ExtractStructure(targetShape)))
                {
                    isEditing = true;
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { EditShape(events, targetShape); }
                        finally { isEditing = false; }
                    });
                    return (IntPtr)1;
                }
            }
            catch { }
        }
        return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
    }

public enum StructureFormatPreference
{
    Auto = 0,
    Smiles = 1,
    InChI = 2,
    Ket = 3
}

    [ExcelCommand(ShortCut = "^+K", MenuName = "SMILES2IMG", MenuText = "Insert/Edit Structure (SMILES2IMG) (Ctrl+Shift+K)")]
    public static void OpenKetcherCommand()
    {
        OpenKetcherWithPreference(StructureFormatPreference.Smiles);
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Insert/Edit Structure (INCHI2IMG)")]
    public static void OpenKetcherInchiCommand()
    {
        OpenKetcherWithPreference(StructureFormatPreference.InChI);
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Insert/Edit Structure (KET2IMG)")]
    public static void OpenKetcherKetCommand()
    {
        OpenKetcherWithPreference(StructureFormatPreference.Ket);
    }

    private static void OpenKetcherWithPreference(StructureFormatPreference preference)
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            if (targetShape != null && !string.IsNullOrEmpty(ExtractStructure(targetShape)))
            {
                EditShape(app, targetShape);
                return;
            }

            if (app.ActiveCell is not Excel.Range activeCell) return;
            var sheet = (Excel.Worksheet)app.ActiveSheet;
            OpenEditor(sheet, activeCell.Address, preference);
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Ketcher command error: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Insert Canvas Structure (KET)")]
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

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Edit Structure")]
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

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy SMILES")]
    public static void CopySelectedMoleculeSmilesCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? smiles = null;

            if (targetShape != null)
                smiles = ExtractSmiles(targetShape);
            else if (app.ActiveCell is Excel.Range cell)
                smiles = ExtractSmilesFromCell(cell);

            if (string.IsNullOrEmpty(smiles))
            {
                MoleculeNotification.NotifyWarning("SMILES", "No SMILES structure found in selected cell or object.");
                return;
            }

            SafeClipboard.SetText(smiles!);
            MoleculeNotification.NotifyCopied("SMILES", smiles!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("SMILES", "Error copying SMILES: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy InChI")]
    public static void CopySelectedMoleculeInchiCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? targetData = null;

            if (targetShape != null)
                targetData = ExtractSmiles(targetShape) ?? ExtractStructure(targetShape);
            else if (app.ActiveCell is Excel.Range cell)
                targetData = ExtractSmilesFromCell(cell) ?? ExtractKetFromCell(cell);

            var inchi = MoleculeRenderer.TryConvertToInchi(targetData);
            if (string.IsNullOrEmpty(inchi))
            {
                MoleculeNotification.NotifyWarning("InChI", "Could not generate InChI from selected cell or object.");
                return;
            }

            SafeClipboard.SetText(inchi!);
            MoleculeNotification.NotifyCopied("InChI", inchi!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChI", "Error copying InChI: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy Canvas Data (KET)")]
    public static void CopySelectedMoleculeKetCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? ket = null;

            if (targetShape != null)
            {
                var structStr = ExtractStructure(targetShape);
                ket = MoleculeRenderer.TryConvertToKet(structStr);
            }
            else if (app.ActiveCell is Excel.Range cell)
            {
                ket = ExtractKetFromCell(cell);
            }

            if (string.IsNullOrEmpty(ket))
            {
                MoleculeNotification.NotifyWarning("Canvas Data (KET)", "No canvas structure found in selected cell or object.");
                return;
            }

            SafeClipboard.SetText(ket!);
            MoleculeNotification.NotifyCopied("Canvas Data (KET)", ket!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Canvas Data (KET)", "Error copying canvas data: " + ex.Message);
        }
    }

    public static MoleculeInfo? GetSelectedMoleculeInfo(Excel.Application app, bool resolvePubChem = false)
    {
        var targetShape = GetSelectedShape(app);
        string? targetData = null;

        if (targetShape != null)
        {
            if (StructureMetadata.IsMolstar(targetShape.AlternativeText, targetShape.Title))
            {
                var meta = MolstarSessionManager.ParseAltText(targetShape.AlternativeText, targetShape.Title);
                if (meta != null && !string.IsNullOrWhiteSpace(meta.Pdb))
                {
                    (string Title, double? WeightKDa) pdbInfo = resolvePubChem ? PubChemResolver.ResolvePdb(meta.Pdb) : ("", (double?)null);
                    return new MoleculeInfo
                    {
                        CompoundName = !string.IsNullOrEmpty(pdbInfo.Title) ? pdbInfo.Title : $"PDB {meta.Pdb}",
                        IupacName = $"PDB Entry: {meta.Pdb}",
                        Formula = "Macromolecular Complex",
                        MolecularWeight = pdbInfo.WeightKDa.HasValue ? pdbInfo.WeightKDa.Value * 1000.0 : 0.0,
                        ExactMass = 0,
                        Smiles = $"PDB:{meta.Pdb}",
                        IsReaction = false
                    };
                }
            }
            targetData = ExtractSmiles(targetShape) ?? ExtractStructure(targetShape);
        }
        else
        {
            Excel.Range? cell = null;
            try { cell = app.ActiveCell as Excel.Range; } catch { }
            if (cell == null)
            {
                try { cell = app.Selection as Excel.Range; } catch { }
            }
            if (cell != null)
            {
                if (cell.Count > 1)
                {
                    try { cell = (Excel.Range)cell.Cells[1, 1]; } catch { }
                }
                targetData = ExtractSmilesFromCell(cell) ?? ExtractKetFromCell(cell);
            }
        }

        if (string.IsNullOrEmpty(targetData)) return null;
        return MoleculeRenderer.TryGetMoleculeInfo(targetData, resolvePubChem);
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy Substance Name")]
    public static void CopySelectedSubstanceNameCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null || string.IsNullOrWhiteSpace(info.CompoundName))
            {
                MoleculeNotification.NotifyWarning("Substance Name", "Could not resolve substance name from selected cell or object.");
                return;
            }
            SafeClipboard.SetText(info.CompoundName);
            MoleculeNotification.NotifyCopied("Substance Name", info.CompoundName);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Substance Name", "Error copying substance name: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy Chemical Formula")]
    public static void CopySelectedFormulaCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: false);
            if (info == null || string.IsNullOrWhiteSpace(info.Formula))
            {
                MoleculeNotification.NotifyWarning("Formula", "Could not compute chemical formula from selected cell or object.");
                return;
            }
            SafeClipboard.SetText(info.Formula);
            MoleculeNotification.NotifyCopied("Formula", info.Formula);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Formula", "Error copying chemical formula: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy Molecular Weight")]
    public static void CopySelectedMolecularWeightCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: false);
            if (info == null || info.MolecularWeight <= 0)
            {
                MoleculeNotification.NotifyWarning("Molecular Weight", "Could not compute molecular weight from selected cell or object.");
                return;
            }
            string mw = info.MolecularWeight.ToString("F2");
            SafeClipboard.SetText(mw);
            MoleculeNotification.NotifyCopied("Molecular Weight", $"{mw} g/mol");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecular Weight", "Error copying molecular weight: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy InChIKey")]
    public static void CopySelectedInchiKeyCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: false);
            if (info == null || string.IsNullOrWhiteSpace(info.InChIKey))
            {
                MoleculeNotification.NotifyWarning("InChIKey", "Could not generate InChIKey from selected cell or object.");
                return;
            }
            SafeClipboard.SetText(info.InChIKey);
            MoleculeNotification.NotifyCopied("InChIKey", info.InChIKey);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChIKey", "Error copying InChIKey: " + ex.Message);
        }
    }

    [ExcelCommand(Name = "CopySelectedAllSummaryCommand", MenuName = "SMILES2IMG", MenuText = "Copy All Summary")]
    public static void CopySelectedAllSummaryCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from selected cell or object.");
                return;
            }
            string summary = info.FormatClipboardSummary(MoleculeSummaryFormat.ExcelTsvWithHeader);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.ExcelTsvWithHeader);
            SafeClipboard.SetTextAndHtml(summary, html);
            MoleculeNotification.NotifyCopied("Molecule Info", info.SubstanceName ?? info.Formula ?? "Copied summary");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying molecule summary: " + ex.Message);
        }
    }

    [ExcelCommand(Name = "CopySelectedTsvRowCommand", MenuName = "SMILES2IMG", MenuText = "Copy as Excel Row (TSV Columns)")]
    public static void CopySelectedTsvRowCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from selected cell or object.");
                return;
            }
            string tsv = info.FormatClipboardSummary(MoleculeSummaryFormat.ExcelTsvRow);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.ExcelTsvRow);
            SafeClipboard.SetTextAndHtml(tsv, html);
            MoleculeNotification.NotifyCopied("Excel Row (TSV)", info.SubstanceName ?? info.Formula ?? "Copied row");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying TSV row: " + ex.Message);
        }
    }

    [ExcelCommand(Name = "CopySelectedMultiLineCommand", MenuName = "SMILES2IMG", MenuText = "Copy as Multi-Line List")]
    public static void CopySelectedMultiLineCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from selected cell or object.");
                return;
            }
            string multi = info.FormatClipboardSummary(MoleculeSummaryFormat.MultiLine);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.MultiLine);
            SafeClipboard.SetTextAndHtml(multi, html);
            MoleculeNotification.NotifyCopied("Multi-Line List", info.SubstanceName ?? info.Formula ?? "Copied list");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying multi-line list: " + ex.Message);
        }
    }

    [ExcelCommand(Name = "CopySelectedMarkdownCommand", MenuName = "SMILES2IMG", MenuText = "Copy as Markdown Table")]
    public static void CopySelectedMarkdownCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from selected cell or object.");
                return;
            }
            string md = info.FormatClipboardSummary(MoleculeSummaryFormat.MarkdownTable);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.MarkdownTable);
            SafeClipboard.SetTextAndHtml(md, html);
            MoleculeNotification.NotifyCopied("Markdown Table", info.SubstanceName ?? info.Formula ?? "Copied table");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying markdown table: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Molecule Info (Name, MW & Formula)...")]
    public static void CopySelectedMoleculeInfoCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MessageBox.Show(
                    "Cannot retrieve formula/molecular weight for the selected cell or image.\nPlease select a molecular structure cell or image.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = ExcelDnaUtil.WindowHandle;
            IWin32Window? owner = hwnd != IntPtr.Zero ? new Win32Window(hwnd) : null;
            if (owner != null)
                MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties", owner);
            else
                MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while calculating molecule info: " + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Paste Structure")]
    public static void QuickInsertCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet is not Excel.Worksheet sheet) return;

            string? rawInput = Clipboard.ContainsText() ? Clipboard.GetText()?.Trim() : null;

            // Fallback to current cell plain text if clipboard doesn't have valid structure
            if (app.ActiveCell is Excel.Range activeCell)
            {
                string cellFormula = Convert.ToString((object?)activeCell.Formula) ?? "";
                string cellVal = Convert.ToString((object?)activeCell.Value2)?.Trim() ?? "";
                if (!cellFormula.StartsWith("=", StringComparison.Ordinal) && !string.IsNullOrEmpty(cellVal) && IsLikelySmiles(cellVal))
                {
                    if (string.IsNullOrEmpty(rawInput) || !MoleculeRenderer.TryRenderStructure(rawInput, RenderOptions.Default, out _, out _, out _))
                        rawInput = cellVal;
                }
            }

            if (!MoleculeRenderer.TryRenderStructure(rawInput, RenderOptions.Default, out var pngBytes, out var ket, out var smiles) || pngBytes == null)
            {
                MessageBox.Show(
                    "No valid SMILES or KET structure data found on clipboard (or active cell).\nPlease copy a SMILES string (e.g. CC(=O)Oc1ccccc1C(=O)O) and try again.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (app.ActiveCell is Excel.Range targetCell)
            {
                var isKet = !string.IsNullOrEmpty(ket);
                var dataToWrite = isKet ? StructureMetadata.CompactKet(ket) : smiles;

                if (!string.IsNullOrEmpty(dataToWrite))
                {
                    var escaped = dataToWrite!.Replace("\"", "\"\"");
                    if (escaped.Length <= 8000)
                    {
                        try
                        {
                            if (Convert.ToDouble(targetCell.RowHeight) < 85) targetCell.RowHeight = 95;
                            if (Convert.ToDouble(targetCell.ColumnWidth) < 16) targetCell.ColumnWidth = 20;
                        }
                        catch { }

                        var funcName = isKet
                            ? (CellImages.NativeImagesUnsupported ? "KET2IMG.FLOAT" : "KET2IMG")
                            : (CellImages.NativeImagesUnsupported ? "SMILES2IMG.FLOAT" : "SMILES2IMG");

                        targetCell.Formula = $"={funcName}(\"{escaped}\")";
                        return;
                    }
                }
            }

            InsertCanvasShapeOnSheet(app, sheet, ket, smiles, pngBytes);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred during quick paste: " + ex.Message,
                "SMILES for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Insert 3D Structure (Mol*)")]
    public static void InsertMolstarCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            ExcelMolstarManager.InsertMolstarStructure(app);
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Mol* insert error: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Edit 3D Structure (Mol*)")]
    public static void EditMolstarCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            ExcelMolstarManager.EditSelectedMolstar(app);
        }
        catch (Exception ex)
        {
            LogDisplay.WriteLine("Mol* edit error: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy as Vector SVG")]
    public static void CopySelectedMoleculeSvgCommand()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var targetShape = GetSelectedShape(app);
            string? targetData = null;

            if (targetShape != null)
                targetData = ExtractStructure(targetShape) ?? ExtractSmiles(targetShape);
            else if (app.ActiveCell is Excel.Range cell)
                targetData = ExtractKetFromCell(cell) ?? ExtractSmilesFromCell(cell);

            var svg = MoleculeRenderer.RenderSvg(targetData);
            if (string.IsNullOrEmpty(svg) || !SvgClipboardHelper.CopySvgToClipboard(svg))
            {
                MoleculeNotification.NotifyWarning("Copy SVG", "Could not render SVG from the selected cell or object.");
                return;
            }

            MoleculeNotification.NotifyCopied("Vector SVG", "Scalable SVG copied to clipboard");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Copy SVG", "Error copying SVG: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Batch Style: Standard Color (Transparent)")]
    public static void BatchStyleColorTransCommand()
        => ApplyBatchStyle("trans", "color", "Standard Color (Transparent)");

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Batch Style: Publication B&W (ACS Style)")]
    public static void BatchStyleBwCommand()
        => ApplyBatchStyle("trans", "bw", "Publication B&W");

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Batch Style: White Background (Color)")]
    public static void BatchStyleWhiteBgCommand()
        => ApplyBatchStyle("white", "color", "White Background");

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Batch Style: Dark Mode (White Ink)")]
    public static void BatchStyleDarkCommand()
        => ApplyBatchStyle("trans", "white", "Dark Mode (White Ink)");

    public static int ApplyBatchStyle(string background, string style, string presetLabel = "Preset")
    {
        int updated = 0;
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet is not Excel.Worksheet sheet) return 0;

            var options = RenderOptions.Parse(background, style, null);
            var formulaRegex = new System.Text.RegularExpressions.Regex(
                @"^=\s*(?<fn>(?:SMILES2IMG|KET2IMG|INCHI2IMG|NAME2IMG)(?:\.FLOAT)?)\s*\(\s*(?<arg1>""(?:""""|[^""])*""|[^,)]+)(?:\s*,\s*[^,)]*)*\s*\)\s*$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            var used = sheet.UsedRange;
            if (used != null)
            {
                int rows = Math.Min(used.Rows.Count, 500);
                int cols = Math.Min(used.Columns.Count, 50);
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        try
                        {
                            var cell = (Excel.Range)used.Cells[r, c];
                            string formula = Convert.ToString((object?)cell.Formula) ?? "";
                            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=", StringComparison.Ordinal))
                                continue;

                            var m = formulaRegex.Match(formula);
                            if (m.Success)
                            {
                                var fn = m.Groups["fn"].Value;
                                var arg1 = m.Groups["arg1"].Value.Trim();
                                cell.Formula = $"={fn}({arg1}, \"{background}\", \"{style}\")";
                                updated++;
                            }
                        }
                        catch { }
                    }
                }
            }

            // Also update standalone canvas shapes (not managed by FloatingPictures)
            try
            {
                var shapesToUpdate = new System.Collections.Generic.List<Excel.Shape>();
                foreach (Excel.Shape s in sheet.Shapes)
                {
                    string? title = null;
                    try { title = s.Title; } catch { }
                    if (!string.IsNullOrEmpty(title) && FloatingPictures.TryParseTitle(title, out _, out _))
                        continue; // Managed by formula

                    if (StructureMetadata.IsMolstar(s.AlternativeText, title))
                        continue;

                    if (!string.IsNullOrEmpty(ExtractStructure(s)))
                        shapesToUpdate.Add(s);
                }

                foreach (var s in shapesToUpdate)
                {
                    var raw = ExtractStructure(s);
                    if (MoleculeRenderer.TryRenderStructure(raw, options, out var pngBytes, out var ket, out var smiles) && pngBytes != null)
                    {
                        ReplaceShape(s, ket, smiles, pngBytes);
                        updated++;
                    }
                }
            }
            catch { }

            if (updated > 0)
                MoleculeNotification.NotifyCopied("Batch Style", $"Updated {updated} structure(s) to {presetLabel}");
            else
                MoleculeNotification.NotifyWarning("Batch Style", "No SMILES structures found on the active sheet.");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Batch Style", "Error applying batch style: " + ex.Message);
        }
        return updated;
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Import SDF File...")]
    public static void ImportSdfCommand()
    {
        try
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Import Chemical Structure Data File (.sdf)",
                Filter = "MDL Structure Data Files (*.sdf;*.mol)|*.sdf;*.mol|All Files (*.*)|*.*",
                Multiselect = false
            };

            var hwnd = ExcelDnaUtil.WindowHandle;
            var dr = hwnd != IntPtr.Zero ? ofd.ShowDialog(new Win32Window(hwnd)) : ofd.ShowDialog();
            if (dr != DialogResult.OK || string.IsNullOrWhiteSpace(ofd.FileName))
                return;

            var app = (Excel.Application)ExcelDnaUtil.Application;
            int count = ImportSdfFromFile(app, ofd.FileName);
            if (count > 0)
                MoleculeNotification.NotifyCopied("SDF Import", $"Imported {count} molecule(s) from {System.IO.Path.GetFileName(ofd.FileName)}");
            else
                MoleculeNotification.NotifyWarning("SDF Import", "No valid molecules found in the selected SDF file.");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error importing SDF file:\n" + ex.Message, "SMILES for Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public static int ImportSdfFromFile(Excel.Application app, string filePath)
    {
        if (app.ActiveSheet is not Excel.Worksheet sheet)
            return 0;

        var records = SdfManager.ReadSdfFile(filePath, maxRecords: 500);
        if (records.Count == 0)
            return 0;

        bool prevScreenUpdating = app.ScreenUpdating;
        try
        {
            app.ScreenUpdating = false;

            int startRow = 1;
            try
            {
                var used = sheet.UsedRange;
                if (used != null && (used.Rows.Count > 1 || used.Columns.Count > 1 || Convert.ToString((object?)((Excel.Range)sheet.Cells[1, 1]).Value2) != null))
                {
                    startRow = used.Row + used.Rows.Count + 1;
                }
            }
            catch { }

            // Header row
            string[] headers = ["Name", "Structure", "SMILES", "Formula", "MW (g/mol)", "Exact Mass", "LogP", "Lipinski"];
            for (int c = 0; c < headers.Length; c++)
            {
                var hCell = (Excel.Range)sheet.Cells[startRow, c + 1];
                hCell.Value2 = headers[c];
                hCell.Font.Bold = true;
            }

            try
            {
                ((Excel.Range)sheet.Columns[1]).ColumnWidth = 18;
                ((Excel.Range)sheet.Columns[2]).ColumnWidth = 18;
                ((Excel.Range)sheet.Columns[3]).ColumnWidth = 28;
                ((Excel.Range)sheet.Columns[4]).ColumnWidth = 14;
                ((Excel.Range)sheet.Columns[5]).ColumnWidth = 12;
                ((Excel.Range)sheet.Columns[6]).ColumnWidth = 14;
                ((Excel.Range)sheet.Columns[7]).ColumnWidth = 10;
                ((Excel.Range)sheet.Columns[8]).ColumnWidth = 10;
            }
            catch { }

            string imgFunc = CellImages.NativeImagesUnsupported ? "SMILES2IMG.FLOAT" : "SMILES2IMG";
            int imported = 0;
            for (int i = 0; i < records.Count; i++)
            {
                var rec = records[i];
                int row = startRow + 1 + i;
                var rowRange = (Excel.Range)sheet.Rows[row];
                try { if (Convert.ToDouble(rowRange.RowHeight) < 75) rowRange.RowHeight = 80; } catch { }

                ((Excel.Range)sheet.Cells[row, 1]).Value2 = !string.IsNullOrWhiteSpace(rec.Name) ? rec.Name : $"Compound #{i + 1}";
                ((Excel.Range)sheet.Cells[row, 3]).Value2 = rec.Smiles;
                ((Excel.Range)sheet.Cells[row, 2]).Formula = $"={imgFunc}(C{row})";
                ((Excel.Range)sheet.Cells[row, 4]).Formula = $"=SMILES2FORMULA(C{row})";
                ((Excel.Range)sheet.Cells[row, 5]).Formula = $"=SMILES2MW(C{row})";
                ((Excel.Range)sheet.Cells[row, 6]).Formula = $"=SMILES2EXACTMASS(C{row})";
                ((Excel.Range)sheet.Cells[row, 7]).Formula = $"=SMILES2LOGP(C{row})";
                ((Excel.Range)sheet.Cells[row, 8]).Formula = $"=SMILES2LIPINSKI(C{row})";
                imported++;
            }

            return imported;
        }
        finally
        {
            try { app.ScreenUpdating = prevScreenUpdating; } catch { }
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Export Sheet to SDF...")]
    public static void ExportSdfCommand()
    {
        try
        {
            using var sfd = new SaveFileDialog
            {
                Title = "Export Worksheet Molecules to SDF (.sdf)",
                Filter = "MDL Structure Data File (*.sdf)|*.sdf|All Files (*.*)|*.*",
                FileName = "molecules_export.sdf",
                DefaultExt = "sdf"
            };

            var hwnd = ExcelDnaUtil.WindowHandle;
            var dr = hwnd != IntPtr.Zero ? sfd.ShowDialog(new Win32Window(hwnd)) : sfd.ShowDialog();
            if (dr != DialogResult.OK || string.IsNullOrWhiteSpace(sfd.FileName))
                return;

            var app = (Excel.Application)ExcelDnaUtil.Application;
            int count = ExportSdfToFile(app, sfd.FileName);
            if (count > 0)
                MoleculeNotification.NotifyCopied("SDF Export", $"Exported {count} molecule(s) to {System.IO.Path.GetFileName(sfd.FileName)}");
            else
                MoleculeNotification.NotifyWarning("SDF Export", "No valid SMILES structures found on the active worksheet to export.");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error exporting SDF file:\n" + ex.Message, "SMILES for Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public static int ExportSdfToFile(Excel.Application app, string filePath)
    {
        if (app.ActiveSheet is not Excel.Worksheet sheet)
            return 0;

        var records = new System.Collections.Generic.List<SdfRecord>();
        var seenSmilesRows = new System.Collections.Generic.HashSet<int>();

        var used = sheet.UsedRange;
        if (used != null)
        {
            int rows = Math.Min(used.Rows.Count, 1000);
            int cols = Math.Min(used.Columns.Count, 30);

            for (int r = 1; r <= rows; r++)
            {
                string? rowSmiles = null;
                string? rowName = null;

                for (int c = 1; c <= cols; c++)
                {
                    try
                    {
                        var cell = (Excel.Range)used.Cells[r, c];
                        string? extracted = ExtractSmilesFromCell(cell);
                        if (!string.IsNullOrWhiteSpace(extracted) && string.IsNullOrEmpty(rowSmiles))
                        {
                            rowSmiles = extracted;
                        }
                        else
                        {
                            string val = Convert.ToString((object?)cell.Value2)?.Trim() ?? "";
                            string formula = Convert.ToString((object?)cell.Formula) ?? "";
                            if (string.IsNullOrEmpty(rowName) &&
                                !string.IsNullOrEmpty(val) &&
                                !formula.StartsWith("=", StringComparison.Ordinal) &&
                                !string.Equals(val, "Name", StringComparison.OrdinalIgnoreCase) &&
                                !string.Equals(val, "SMILES", StringComparison.OrdinalIgnoreCase) &&
                                !IsLikelySmiles(val) &&
                                !double.TryParse(val, out _))
                            {
                                rowName = val;
                            }
                        }
                    }
                    catch { }
                }

                if (!string.IsNullOrWhiteSpace(rowSmiles) && seenSmilesRows.Add(r))
                {
                    records.Add(new SdfRecord
                    {
                        Name = rowName ?? $"Row_{r}",
                        Smiles = rowSmiles!
                    });
                }
            }
        }

        if (records.Count == 0)
            return 0;

        SdfManager.WriteSdfFile(filePath, records);
        return records.Count;
    }
}
