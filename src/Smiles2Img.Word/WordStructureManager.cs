using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Smiles2Img.Rendering;
using MSWord = Microsoft.Office.Interop.Word;

namespace Smiles2Img.Word;

public static class WordStructureManager
{
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

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(form.ResultSmiles))
            {
                var smiles = form.ResultSmiles!.Trim();
                InsertSmilesAsPicture(app, smiles);
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
            string? currentSmiles = null;

            if (sel.InlineShapes.Count > 0)
            {
                targetInline = sel.InlineShapes[1];
                currentSmiles = ExtractSmiles(targetInline.AlternativeText, targetInline.Title);
            }
            else if (sel.ShapeRange.Count > 0)
            {
                targetShape = sel.ShapeRange[1];
                currentSmiles = ExtractSmiles(targetShape.AlternativeText, targetShape.Title);
            }

            if (string.IsNullOrEmpty(currentSmiles))
            {
                MessageBox.Show(
                    "선택된 분자 구조식 이미지를 찾을 수 없습니다.\n먼저 편집하려는 구조식 그림을 클릭하여 선택해 주세요.",
                    "SMILES for Word",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = new KetcherForm("Word", currentSmiles ?? "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(form.ResultSmiles))
            {
                var newSmiles = form.ResultSmiles!.Trim();
                if (targetInline != null)
                {
                    ReplaceInlineShape(app, targetInline, newSmiles);
                }
                else if (targetShape != null)
                {
                    ReplaceShape(app, targetShape, newSmiles);
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
                var smiles = ExtractSmiles(shape.AlternativeText, shape.Title);
                if (!string.IsNullOrEmpty(smiles))
                {
                    cancel = true;
                    var app = sel.Application;
                    EditSelectedMolecule(app);
                }
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                var smiles = ExtractSmiles(shape.AlternativeText, shape.Title);
                if (!string.IsNullOrEmpty(smiles))
                {
                    cancel = true;
                    var app = sel.Application;
                    EditSelectedMolecule(app);
                }
            }
        }
        catch { }
    }

    private static void InsertSmilesAsPicture(MSWord.Application app, string smiles)
    {
        var pngBytes = MoleculeRenderer.Render(smiles, RenderOptions.Default);
        var tempFile = Path.Combine(Path.GetTempPath(), $"smiles_{Guid.NewGuid():N}.png");

        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var sel = app.Selection;
            var shape = sel.InlineShapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true);

            shape.AlternativeText = SmilesPrefix + smiles;
            shape.Title = SmilesPrefix + smiles;
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceInlineShape(MSWord.Application app, MSWord.InlineShape oldShape, string newSmiles)
    {
        var pngBytes = MoleculeRenderer.Render(newSmiles, RenderOptions.Default);
        var tempFile = Path.Combine(Path.GetTempPath(), $"smiles_{Guid.NewGuid():N}.png");

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

            newShape.AlternativeText = SmilesPrefix + newSmiles;
            newShape.Title = SmilesPrefix + newSmiles;
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceShape(MSWord.Application app, MSWord.Shape oldShape, string newSmiles)
    {
        var pngBytes = MoleculeRenderer.Render(newSmiles, RenderOptions.Default);
        var tempFile = Path.Combine(Path.GetTempPath(), $"smiles_{Guid.NewGuid():N}.png");

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

            newShape.AlternativeText = SmilesPrefix + newSmiles;
            newShape.Title = SmilesPrefix + newSmiles;
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static string? ExtractSmiles(params string?[] texts)
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
