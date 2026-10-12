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
    private const string KetPrefix = "KET:";
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

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultKet) || !string.IsNullOrWhiteSpace(form.ResultSmiles)))
            {
                var ket = form.ResultKet;
                var smiles = form.ResultSmiles?.Trim() ?? "";
                var pngBytes = form.ResultPngBytes ?? MoleculeRenderer.Render(smiles, RenderOptions.Default);
                InsertMoleculeAsPicture(app, slide, ket, smiles, pngBytes);
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

    public static void EditSelectedMolecule(MSPowerPoint.Application app, MSPowerPoint.Shape? targetShape = null)
    {
        try
        {
            if (targetShape == null)
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

                targetShape = sel.ShapeRange[1];
            }

            var currentStructure = ExtractStructure(targetShape);

            if (string.IsNullOrEmpty(currentStructure))
            {
                MessageBox.Show(
                    "선택된 개체는 SMILES/KET 분자 구조식 이미지가 아닙니다.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = new KetcherForm("PowerPoint", currentStructure ?? "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (!string.IsNullOrWhiteSpace(form.ResultKet) || !string.IsNullOrWhiteSpace(form.ResultSmiles)))
            {
                var newKet = form.ResultKet;
                var newSmiles = form.ResultSmiles?.Trim() ?? "";
                var newPngBytes = form.ResultPngBytes ?? MoleculeRenderer.Render(newSmiles, RenderOptions.Default);
                ReplaceShape(targetShape, newKet, newSmiles, newPngBytes);
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
                var structStr = ExtractStructure(shape);
                if (!string.IsNullOrEmpty(structStr))
                {
                    cancel = true;
                    var app = sel.Application;
                    EditSelectedMolecule(app, shape);
                }
            }
        }
        catch { }
    }

    private static void InsertMoleculeAsPicture(MSPowerPoint.Application app, MSPowerPoint.Slide slide, string? ket, string smiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");

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

            shape.AlternativeText = !string.IsNullOrEmpty(ket) ? KetPrefix + ket : SmilesPrefix + smiles;
            shape.Title = !string.IsNullOrEmpty(smiles) ? SmilesPrefix + smiles : "Molecule";
            if (!string.IsNullOrEmpty(ket)) try { shape.Tags.Add("KET", ket); } catch { }
            if (!string.IsNullOrEmpty(smiles)) try { shape.Tags.Add("SMILES", smiles); } catch { }
            shape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ReplaceShape(MSPowerPoint.Shape oldShape, string? newKet, string newSmiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");

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

            newShape.AlternativeText = !string.IsNullOrEmpty(newKet) ? KetPrefix + newKet : SmilesPrefix + newSmiles;
            newShape.Title = !string.IsNullOrEmpty(newSmiles) ? SmilesPrefix + newSmiles : "Molecule";
            if (!string.IsNullOrEmpty(newKet)) try { newShape.Tags.Add("KET", newKet); } catch { }
            if (!string.IsNullOrEmpty(newSmiles)) try { newShape.Tags.Add("SMILES", newSmiles); } catch { }
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    public static string? ExtractStructure(MSPowerPoint.Shape shape)
    {
        try
        {
            var ketTag = shape.Tags["KET"];
            if (!string.IsNullOrWhiteSpace(ketTag)) return ketTag.Trim();
        }
        catch { }

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

        // Fallback to legacy SMILES
        try
        {
            var smilesTag = shape.Tags["SMILES"];
            if (!string.IsNullOrWhiteSpace(smilesTag)) return smilesTag.Trim();
        }
        catch { }

        return ExtractSmilesFromText(shape.AlternativeText, shape.Title);
    }

    public static string? ExtractSmiles(MSPowerPoint.Shape shape)
    {
        try
        {
            var smilesTag = shape.Tags["SMILES"];
            if (!string.IsNullOrWhiteSpace(smilesTag)) return smilesTag.Trim();
        }
        catch { }

        return ExtractSmilesFromText(shape.Title, shape.AlternativeText);
    }

    public static void CopySelectedMoleculeSmiles(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel == null || sel.Type != MSPowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MessageBox.Show(
                    "선택된 이미지에서 SMILES 구조식을 찾을 수 없습니다.\n먼저 분자 구조식 이미지를 클릭하여 선택해 주세요.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var targetShape = sel.ShapeRange[1];
            var smiles = ExtractSmiles(targetShape);

            if (string.IsNullOrEmpty(smiles))
            {
                MessageBox.Show(
                    "선택된 이미지에서 SMILES 구조식을 찾을 수 없습니다.\n먼저 분자 구조식 이미지를 클릭하여 선택해 주세요.",
                    "SMILES for PowerPoint",
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
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void CopySelectedMoleculeKet(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel == null || sel.Type != MSPowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MessageBox.Show(
                    "선택된 분자 구조식 이미지를 찾을 수 없습니다.\n먼저 분자 구조식 이미지를 클릭하여 선택해 주세요.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var targetShape = sel.ShapeRange[1];
            var ket = ExtractStructure(targetShape);

            if (string.IsNullOrEmpty(ket))
            {
                MessageBox.Show(
                    "선택된 이미지에서 캔버스 구조식(KET)을 찾을 수 없습니다.",
                    "SMILES for PowerPoint",
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
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
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
