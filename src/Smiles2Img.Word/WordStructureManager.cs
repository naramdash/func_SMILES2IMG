using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Smiles2Img.Rendering;
using MSWord = Microsoft.Office.Interop.Word;

namespace Smiles2Img.Word;

public static class WordStructureManager
{
    private const string KetPrefix = "KET:";
    private const string SmilesPrefix = "SMILES:";
    private static readonly Regex SmilesTagRegex = new(@"\bSMILES:(?<smiles>[^\r\n]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

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
                var pngBytes = ResolvePngBytes(form, ket, smiles);
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
                var newPngBytes = ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;

                if (targetInline != null)
                {
                    ReplaceInlineShape(app, targetInline, newKet, newSmiles, newPngBytes);
                }
                else if (targetShape != null)
                {
                    ReplaceShape(app, targetShape, newKet, newSmiles, newPngBytes);
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
                var structStr = ExtractStructure(shape.AlternativeText, shape.Title);
                if (!string.IsNullOrEmpty(structStr))
                {
                    cancel = true;
                    var app = sel.Application;
                    EditSelectedMolecule(app);
                }
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                var structStr = ExtractStructure(shape.AlternativeText, shape.Title);
                if (!string.IsNullOrEmpty(structStr))
                {
                    cancel = true;
                    var app = sel.Application;
                    EditSelectedMolecule(app);
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
            var sel = app.Selection;
            var shape = sel.InlineShapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true);

            shape.AlternativeText = !string.IsNullOrEmpty(ket) ? KetPrefix + ket : SmilesPrefix + smiles;
            shape.Title = !string.IsNullOrEmpty(smiles) ? SmilesPrefix + smiles : "Molecule";
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceInlineShape(MSWord.Application app, MSWord.InlineShape oldShape, string? newKet, string newSmiles, byte[] pngBytes)
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

            try
            {
                using var ms = new MemoryStream(pngBytes);
                using var img = System.Drawing.Image.FromStream(ms);
                float aspect = img.Width / (float)img.Height;
                newShape.Width = oldWidth;
                newShape.Height = oldWidth / aspect;
            }
            catch { }

            newShape.AlternativeText = !string.IsNullOrEmpty(newKet) ? KetPrefix + newKet : SmilesPrefix + newSmiles;
            newShape.Title = !string.IsNullOrEmpty(newSmiles) ? SmilesPrefix + newSmiles : "Molecule";
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceShape(MSWord.Application app, MSWord.Shape oldShape, string? newKet, string newSmiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");

        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var doc = (MSWord.Document)oldShape.Parent;

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

            var newShape = doc.Shapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true,
                Left: left,
                Top: top,
                Width: newWidth,
                Height: newHeight);

            newShape.AlternativeText = !string.IsNullOrEmpty(newKet) ? KetPrefix + newKet : SmilesPrefix + newSmiles;
            newShape.Title = !string.IsNullOrEmpty(newSmiles) ? SmilesPrefix + newSmiles : "Molecule";
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    public static string? ExtractStructure(params string?[] texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var trimmed = text!.Trim();
            if (trimmed.StartsWith(KetPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return trimmed.Substring(KetPrefix.Length).Trim();
            }
            var ketIdx = trimmed.IndexOf(KetPrefix, StringComparison.OrdinalIgnoreCase);
            if (ketIdx >= 0)
            {
                return trimmed.Substring(ketIdx + KetPrefix.Length).Trim();
            }
        }
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var match = SmilesTagRegex.Match(text!);
            if (match.Success)
            {
                return match.Groups["smiles"].Value.Trim();
            }
        }
        return null;
    }

    public static string? ExtractSmiles(params string?[] texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var match = SmilesTagRegex.Match(text!);
            if (match.Success)
            {
                return match.Groups["smiles"].Value.Trim();
            }
        }
        return null;
    }

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

    private static byte[]? ResolvePngBytes(KetcherForm form, string? ket, string? smiles)
    {
        var pngBytes = form.ResultPngBytes;
        if (pngBytes != null && pngBytes.Length > 0) return pngBytes;

        if (!string.IsNullOrWhiteSpace(ket))
        {
            try
            {
                pngBytes = MoleculeRenderer.RenderKet(ket!);
                if (pngBytes != null && pngBytes.Length > 0) return pngBytes;
            }
            catch { }
        }

        if (!string.IsNullOrWhiteSpace(smiles))
        {
            try
            {
                pngBytes = MoleculeRenderer.Render(smiles!, RenderOptions.Default);
                if (pngBytes != null && pngBytes.Length > 0) return pngBytes;
            }
            catch { }
        }

        return null;
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
