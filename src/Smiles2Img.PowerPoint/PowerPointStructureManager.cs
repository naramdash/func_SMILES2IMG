using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Smiles2Img;
using Smiles2Img.Rendering;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Smiles2Img.PowerPoint;

public static class PowerPointStructureManager
{
    private const string SmilesPrefix = "SMILES:";
    private static readonly Regex SmilesTagRegex = new(@"\bSMILES:(?<smiles>[^\r\n]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static void InsertMolecule(MSPowerPoint.Application app)
    {
        try
        {
            if (app.Presentations.Count == 0)
            {
                var pres = app.Presentations.Add();
                pres.Slides.Add(1, MSPowerPoint.PpSlideLayout.ppLayoutBlank);
            }
            else if (app.ActivePresentation.Slides.Count == 0)
            {
                app.ActivePresentation.Slides.Add(1, MSPowerPoint.PpSlideLayout.ppLayoutBlank);
            }

            var slide = GetActiveSlide(app);
            if (slide == null)
            {
                MessageBox.Show(
                    "활성화된 슬라이드가 없습니다.\n분자 구조식을 삽입할 슬라이드를 먼저 선택해 주세요.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = new KetcherForm("PowerPoint", "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(form.ResultSmiles))
            {
                var smiles = form.ResultSmiles!.Trim();
                InsertSmilesAsPicture(app, slide, smiles);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "분자 구조식 삽입 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void EditSelectedMolecule(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel == null || sel.Type != MSPowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MessageBox.Show(
                    "선택된 분자 구조식 이미지를 찾을 수 없습니다.\n먼저 편집하려는 구조식 그림을 클릭하여 선택해 주세요.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var targetShape = sel.ShapeRange[1];
            var currentSmiles = ExtractSmiles(targetShape);

            if (string.IsNullOrEmpty(currentSmiles))
            {
                MessageBox.Show(
                    "선택된 개체는 SMILES 분자 구조식 이미지가 아닙니다.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = new KetcherForm("PowerPoint", currentSmiles ?? "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(form.ResultSmiles))
            {
                var newSmiles = form.ResultSmiles!.Trim();
                ReplaceShape(targetShape, newSmiles);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "분자 구조식 편집 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void OnWindowBeforeDoubleClick(MSPowerPoint.Selection sel, ref bool cancel)
    {
        try
        {
            if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                var smiles = ExtractSmiles(shape);
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

    private static void InsertSmilesAsPicture(MSPowerPoint.Application app, MSPowerPoint.Slide slide, string smiles)
    {
        var pngBytes = MoleculeRenderer.Render(smiles, RenderOptions.Default);
        var tempFile = Path.Combine(Path.GetTempPath(), $"smiles_{Guid.NewGuid():N}.png");

        try
        {
            File.WriteAllBytes(tempFile, pngBytes);

            float slideWidth = 960f;
            float slideHeight = 540f;
            try
            {
                slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
            }
            catch { }

            float width = 240f;
            float height = 180f;
            using (var ms = new MemoryStream(pngBytes))
            using (var img = System.Drawing.Image.FromStream(ms))
            {
                width = img.Width * 0.75f;
                height = img.Height * 0.75f;
            }

            float left = (slideWidth - width) / 2f;
            float top = (slideHeight - height) / 2f;
            if (left < 50f) left = 50f;
            if (top < 50f) top = 50f;

            var shape = slide.Shapes.AddPicture(
                FileName: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: width,
                Height: height);

            shape.AlternativeText = SmilesPrefix + smiles;
            shape.Title = SmilesPrefix + smiles;
            try { shape.Tags.Add("SMILES", smiles); } catch { }
            shape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceShape(MSPowerPoint.Shape oldShape, string newSmiles)
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
            var slide = (MSPowerPoint.Slide)oldShape.Parent;

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

            var newShape = slide.Shapes.AddPicture(
                FileName: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: newWidth,
                Height: newHeight);

            newShape.AlternativeText = SmilesPrefix + newSmiles;
            newShape.Title = SmilesPrefix + newSmiles;
            try { newShape.Tags.Add("SMILES", newSmiles); } catch { }
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static string? ExtractSmiles(MSPowerPoint.Shape shape)
    {
        try
        {
            var tag = shape.Tags["SMILES"];
            if (!string.IsNullOrWhiteSpace(tag)) return tag.Trim();
        }
        catch { }

        return ExtractSmilesFromText(shape.AlternativeText, shape.Title);
    }

    private static string? ExtractSmilesFromText(params string?[] texts)
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

    private static MSPowerPoint.Slide? GetActiveSlide(MSPowerPoint.Application app)
    {
        try
        {
            if (app.ActiveWindow?.View?.Slide is MSPowerPoint.Slide slide)
                return slide;
        }
        catch { }

        try
        {
            if (app.ActiveWindow?.Selection?.SlideRange?.Count > 0)
                return app.ActiveWindow.Selection.SlideRange[1];
        }
        catch { }

        return null;
    }

    private static IntPtr GetHwnd(MSPowerPoint.Application app)
    {
        try
        {
            if (app.ActiveWindow != null)
                return new IntPtr(app.ActiveWindow.HWND);
        }
        catch { }
        return IntPtr.Zero;
    }
}
