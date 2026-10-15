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
                                ctrl.Delete();
                        }
                    }
                    catch { }
                }
            }
        }
        catch { }
    }

    [ExcelCommand(ShortCut = "^+K", MenuName = "SMILES2IMG", MenuText = "Insert/Edit Structure (SMILES2IMG) (Ctrl+Shift+K)")]
    public static void OpenKetcherCommand()
    {
        OpenKetcherWithPreference(preferKet: false);
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Insert/Edit Structure (KET2IMG)")]
    public static void OpenKetcherKetCommand()
    {
        OpenKetcherWithPreference(preferKet: true);
    }

    private static void OpenKetcherWithPreference(bool? preferKet)
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
            OpenEditor(sheet, activeCell.Address, preferKet);
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

            Clipboard.SetText(smiles!);
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

            Clipboard.SetText(inchi!);
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

            Clipboard.SetText(ket!);
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
        else if (app.ActiveCell is Excel.Range cell)
        {
            targetData = ExtractSmilesFromCell(cell) ?? ExtractKetFromCell(cell);
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
            Clipboard.SetText(info.CompoundName);
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
            Clipboard.SetText(info.Formula);
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
            Clipboard.SetText(mw);
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
            Clipboard.SetText(info.InChIKey);
            MoleculeNotification.NotifyCopied("InChIKey", info.InChIKey);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChIKey", "Error copying InChIKey: " + ex.Message);
        }
    }

    [ExcelCommand(MenuName = "SMILES2IMG", MenuText = "Copy All Summary")]
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
            string summary = info.FormatClipboardSummary();
            Clipboard.SetText(summary);
            MoleculeNotification.NotifyCopied("Molecule Info", info.SubstanceName ?? info.Formula ?? "Copied summary");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying molecule summary: " + ex.Message);
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

            IWin32Window? owner = app.Hwnd != 0 ? new Win32Window((IntPtr)app.Hwnd) : null;
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
}
