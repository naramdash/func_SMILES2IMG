using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Microsoft.Office.Core;
using Smiles2Img.Rendering;
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

    [ExcelCommand(ShortCut = "^+K", MenuName = "SMILES2IMG", MenuText = "분자 구조식 삽입/편집 (Ctrl+Shift+K)")]
    public static void OpenKetcherCommand()
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
            OpenEditor(sheet, activeCell.Address);
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
                smiles = ExtractSmiles(targetShape);
            else if (app.ActiveCell is Excel.Range cell)
                smiles = ExtractSmilesFromCell(cell);

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
            var ket = targetShape != null ? ExtractStructure(targetShape) : null;

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
                targetData = ExtractSmiles(targetShape) ?? ExtractStructure(targetShape);
            else if (app.ActiveCell is Excel.Range cell)
                targetData = ExtractSmilesFromCell(cell);

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
                    "클립보드(또는 활성 셀)에서 유효한 SMILES 또는 KET 구조식 데이터를 찾을 수 없습니다.\n먼저 SMILES 문자열(예: CC(=O)Oc1ccccc1C(=O)O)을 복사한 후 실행해 주세요.",
                    "SMILES for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

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

            InsertCanvasShapeOnSheet(app, sheet, ket, smiles, pngBytes);
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
}
