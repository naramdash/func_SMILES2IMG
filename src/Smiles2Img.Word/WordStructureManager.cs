using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Rendering;
using MSWord = Microsoft.Office.Interop.Word;

namespace Smiles2Img.Word;

public static class WordStructureManager
{
    public static void InsertMolecule(MSWord.Application app)
    {
        try
        {
            if (app.Documents.Count == 0)
            {
                app.Documents.Add();
            }

            var hwnd = GetHwnd(app);
            using var form = new KetcherForm("Word", "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultKet) || !string.IsNullOrWhiteSpace(form.ResultSmiles)))
            {
                var ket = form.ResultKet;
                var smiles = form.ResultSmiles?.Trim() ?? "";
                var pngBytes = StructureMetadata.ResolvePngBytes(form, ket, smiles);
                if (pngBytes == null) return;
                InsertMoleculeAsPicture(app, ket, smiles, pngBytes);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "분자 구조식 삽입 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void EditSelectedMolecule(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            MSWord.InlineShape? targetInline = null;
            MSWord.Shape? targetShape = null;
            string? currentStructure = null;

            if (sel.InlineShapes.Count > 0)
            {
                targetInline = sel.InlineShapes[1];
                currentStructure = ExtractStructure(targetInline.AlternativeText, targetInline.Title);
            }
            else if (sel.ShapeRange.Count > 0)
            {
                targetShape = sel.ShapeRange[1];
                currentStructure = ExtractStructure(targetShape.AlternativeText, targetShape.Title);
            }

            if (string.IsNullOrEmpty(currentStructure))
            {
                MessageBox.Show(
                    "선택된 분자 구조식 이미지를 찾을 수 없습니다.\n먼저 편집하려는 구조식 그림을 클릭하여 선택해 주세요.",
                    "SMILES for Word",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = new KetcherForm("Word", currentStructure ?? "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultKet) || !string.IsNullOrWhiteSpace(form.ResultSmiles)))
            {
                var newKet = form.ResultKet;
                var newSmiles = form.ResultSmiles?.Trim() ?? "";
                var newPngBytes = StructureMetadata.ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;

                if (targetInline != null)
                {
                    ReplaceInlineShape(targetInline, newKet, newSmiles, newPngBytes);
                }
                else if (targetShape != null)
                {
                    ReplaceShape(targetShape, newKet, newSmiles, newPngBytes);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "분자 구조식 편집 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void OnWindowBeforeDoubleClick(MSWord.Selection sel, ref bool cancel)
    {
        try
        {
            if (sel.InlineShapes.Count > 0)
            {
                var shape = sel.InlineShapes[1];
                if (!string.IsNullOrEmpty(ExtractStructure(shape.AlternativeText, shape.Title)))
                {
                    cancel = true;
                    EditSelectedMolecule(sel.Application);
                }
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                if (!string.IsNullOrEmpty(ExtractStructure(shape.AlternativeText, shape.Title)))
                {
                    cancel = true;
                    EditSelectedMolecule(sel.Application);
                }
            }
        }
        catch { }
    }

    private static void InsertMoleculeAsPicture(MSWord.Application app, string? ket, string smiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var shape = app.Selection.InlineShapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true);

            shape.AlternativeText = StructureMetadata.FormatAlternativeText(ket, smiles);
            shape.Title = StructureMetadata.FormatTitle(smiles);
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceInlineShape(MSWord.InlineShape oldShape, string? newKet, string newSmiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var oldWidth = oldShape.Width;
            var range = oldShape.Range;

            var newShape = range.InlineShapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true);

            oldShape.Delete();

            var (_, height) = StructureMetadata.GetScaledDimensions(pngBytes, oldWidth, oldWidth, minWidth: oldWidth, maxWidth: oldWidth);
            newShape.Width = oldWidth;
            newShape.Height = height;
            newShape.AlternativeText = StructureMetadata.FormatAlternativeText(newKet, newSmiles);
            newShape.Title = StructureMetadata.FormatTitle(newSmiles);
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceShape(MSWord.Shape oldShape, string? newKet, string newSmiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var doc = (MSWord.Document)oldShape.Parent;
            var (left, top, newWidth, newHeight) = StructureMetadata.FitReplacementBounds(
                oldShape.Left, oldShape.Top, oldShape.Width, oldShape.Height, pngBytes);

            oldShape.Delete();

            var newShape = doc.Shapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true,
                Left: left,
                Top: top,
                Width: newWidth,
                Height: newHeight);

            newShape.AlternativeText = StructureMetadata.FormatAlternativeText(newKet, newSmiles);
            newShape.Title = StructureMetadata.FormatTitle(newSmiles);
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    public static string? ExtractStructure(params string?[] texts) =>
        StructureMetadata.ExtractStructure(texts);

    public static string? ExtractSmiles(params string?[] texts) =>
        StructureMetadata.ExtractSmiles(texts);

    public static void CopySelectedMoleculeSmiles(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            string? smiles = null;

            if (sel.InlineShapes.Count > 0)
            {
                var targetInline = sel.InlineShapes[1];
                smiles = ExtractSmiles(targetInline.Title, targetInline.AlternativeText);
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var targetShape = sel.ShapeRange[1];
                smiles = ExtractSmiles(targetShape.Title, targetShape.AlternativeText);
            }

            if (string.IsNullOrEmpty(smiles))
            {
                MessageBox.Show(
                    "선택된 이미지에서 SMILES 구조식을 찾을 수 없습니다.\n먼저 분자 구조식 이미지를 클릭하여 선택해 주세요.",
                    "SMILES for Word",
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
                "SMILES 추출 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void CopySelectedMoleculeKet(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            string? ket = null;

            if (sel.InlineShapes.Count > 0)
            {
                var targetInline = sel.InlineShapes[1];
                ket = ExtractStructure(targetInline.AlternativeText, targetInline.Title);
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var targetShape = sel.ShapeRange[1];
                ket = ExtractStructure(targetShape.AlternativeText, targetShape.Title);
            }

            if (string.IsNullOrEmpty(ket))
            {
                MessageBox.Show(
                    "선택된 이미지에서 캔버스 구조식(KET)을 찾을 수 없습니다.",
                    "SMILES for Word",
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
                "캔버스 데이터 복사 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void QuickInsertFromSelectionOrClipboard(MSWord.Application app)
    {
        try
        {
            if (app.Documents.Count == 0)
            {
                app.Documents.Add();
            }

            var sel = app.Selection;
            if (sel != null && sel.Start != sel.End && !string.IsNullOrWhiteSpace(sel.Text))
            {
                var selectedText = sel.Text.Trim();
                if (MoleculeRenderer.TryRenderStructure(selectedText, RenderOptions.Default, out var selPng, out var selKet, out var selSmiles) && selPng != null)
                {
                    sel.Text = "";
                    InsertMoleculeAsPicture(app, selKet, selSmiles ?? "", selPng);
                    return;
                }
            }

            if (Clipboard.ContainsText())
            {
                var clipText = Clipboard.GetText()?.Trim();
                if (MoleculeRenderer.TryRenderStructure(clipText, RenderOptions.Default, out var clipPng, out var clipKet, out var clipSmiles) && clipPng != null)
                {
                    InsertMoleculeAsPicture(app, clipKet, clipSmiles ?? "", clipPng);
                    return;
                }
            }

            MessageBox.Show(
                "선택된 텍스트나 클립보드에서 유효한 SMILES 또는 KET 구조식 데이터를 찾을 수 없습니다.\n먼저 SMILES 문자열(예: CC(=O)Oc1ccccc1C(=O)O)을 복사하거나 문서에서 텍스트를 드래그해 선택해 주세요.",
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "클립보드 구조식 즉시 삽입 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void CopySelectedMoleculeInfo(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            string? targetData = null;

            if (sel.InlineShapes.Count > 0)
            {
                var targetInline = sel.InlineShapes[1];
                targetData = ExtractSmiles(targetInline.Title, targetInline.AlternativeText)
                             ?? ExtractStructure(targetInline.AlternativeText, targetInline.Title);
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var targetShape = sel.ShapeRange[1];
                targetData = ExtractSmiles(targetShape.Title, targetShape.AlternativeText)
                             ?? ExtractStructure(targetShape.AlternativeText, targetShape.Title);
            }
            else if (sel != null && sel.Start != sel.End && !string.IsNullOrWhiteSpace(sel.Text))
            {
                targetData = sel.Text.Trim();
            }

            var info = MoleculeRenderer.TryGetMoleculeInfo(targetData);
            if (info == null)
            {
                MessageBox.Show(
                    "선택된 분자 구조식에서 화학식/분자량 정보를 계산할 수 없습니다.\n먼저 분자 구조식 이미지를 클릭하여 선택해 주세요.",
                    "SMILES for Word",
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
                "분자 정보 계산 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static IntPtr GetHwnd(MSWord.Application app)
    {
        try
        {
            if (app.ActiveWindow != null)
                return new IntPtr(app.ActiveWindow.Hwnd);
        }
        catch { }
        return IntPtr.Zero;
    }
}
