using System;
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

    private static readonly Regex FormulaRefPattern = new(
        @"SMILES2IMG(?:\.FLOAT)?\s*\(\s*(?:""(?<literal>[^""]*)""|(?<ref>(?:'[^']+'!)?[^,)\s]+))",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static Excel.Application? events;

    internal static void Attach()
    {
        try
        {
            events = (Excel.Application)ExcelDnaUtil.Application;
            events.SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;
            AddCellContextMenu();
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
            RemoveCellContextMenu();
            if (events != null)
            {
                events.SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
                events = null;
            }
        }
        catch { }
    }

    private static void AddCellContextMenu()
    {
        try
        {
            if (events == null) return;
            RemoveCellContextMenu();

            foreach (CommandBar bar in events.CommandBars)
            {
                if (bar.Name == "Cell" || bar.Name == "List Range Popup")
                {
                    try
                    {
                        var button = (CommandBarButton)bar.Controls.Add(
                            MsoControlType.msoControlButton,
                            Temporary: true,
                            Before: 1);
                        button.Caption = "🧬 분자 구조식 삽입 (Ketcher)...";
                        button.Tag = ContextMenuTag;
                        button.OnAction = "OpenKetcherCommand";
                        button.BeginGroup = true;
                    }
                    catch { }
                }
            }
        }
        catch { }
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

            if (result == DialogResult.OK && form.ResultSmiles != null)
            {
                var newSmiles = form.ResultSmiles.Trim();
                if (string.IsNullOrEmpty(newSmiles)) return;

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
