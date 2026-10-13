using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Rendering;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Smiles2Img.PowerPoint;

public static class PowerPointStructureManager
{
    public static void InsertMolecule(MSPowerPoint.Application app)
    {
        try
        {
            var slide = EnsureActiveSlide(app);
            if (slide == null)
            {
                MessageBox.Show(
                    "No active slide found.\nPlease select a slide to insert the molecule structure first.",
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
                var pngBytes = StructureMetadata.ResolvePngBytes(form, ket, smiles);
                if (pngBytes == null) return;
                InsertMoleculeAsPicture(app, slide, ket, smiles, pngBytes);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while inserting molecule:\n" + ex.Message,
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
                        "No molecular structure image selected.\nPlease select a molecular structure picture first.",
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
                    "The selected object is not a SMILES/KET molecule structure image.",
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
                var newPngBytes = StructureMetadata.ResolvePngBytes(form, newKet, newSmiles);
                if (newPngBytes == null) return;
                ReplaceShape(targetShape, newKet, newSmiles, newPngBytes);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while editing molecule:\n" + ex.Message,
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
                if (!string.IsNullOrEmpty(ExtractStructure(shape)))
                {
                    cancel = true;
                    EditSelectedMolecule(sel.Application, shape);
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

            var (width, height) = StructureMetadata.GetScaledDimensions(pngBytes, 240f, 180f, scale: 0.75f);
            float left = Math.Max(50f, (slideWidth - width) / 2f);
            float top = Math.Max(50f, (slideHeight - height) / 2f);

            var shape = slide.Shapes.AddPicture(
                FileName: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: width,
                Height: height);

            ApplyShapeMetadata(shape, ket, smiles);
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
            var slide = (MSPowerPoint.Slide)oldShape.Parent;
            var (left, top, newWidth, newHeight) = StructureMetadata.FitReplacementBounds(
                oldShape.Left, oldShape.Top, oldShape.Width, oldShape.Height, pngBytes);

            oldShape.Delete();

            var newShape = slide.Shapes.AddPicture(
                FileName: tempFile,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: newWidth,
                Height: newHeight);

            ApplyShapeMetadata(newShape, newKet, newSmiles);
            newShape.Select();
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void ApplyShapeMetadata(MSPowerPoint.Shape shape, string? ket, string? smiles)
    {
        shape.AlternativeText = StructureMetadata.FormatAlternativeText(ket, smiles);
        shape.Title = StructureMetadata.FormatTitle(smiles);
        if (!string.IsNullOrEmpty(ket)) try { shape.Tags.Add("KET", ket); } catch { }
        if (!string.IsNullOrEmpty(smiles)) try { shape.Tags.Add("SMILES", smiles); } catch { }
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
            var extracted = StructureMetadata.ExtractStructure(shape.AlternativeText, shape.Title);
            if (!string.IsNullOrWhiteSpace(extracted)) return extracted;
        }
        catch { }

        try
        {
            var smilesTag = shape.Tags["SMILES"];
            if (!string.IsNullOrWhiteSpace(smilesTag)) return smilesTag.Trim();
        }
        catch { }

        return null;
    }

    public static string? ExtractSmiles(MSPowerPoint.Shape shape)
    {
        try
        {
            var smilesTag = shape.Tags["SMILES"];
            if (!string.IsNullOrWhiteSpace(smilesTag)) return smilesTag.Trim();
        }
        catch { }

        return StructureMetadata.ExtractSmiles(shape.Title, shape.AlternativeText);
    }

    public static void CopySelectedMoleculeSmiles(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel == null || sel.Type != MSPowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MessageBox.Show(
                    "Could not find SMILES structure from the selected image.\nPlease select a molecular structure picture first.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var smiles = ExtractSmiles(sel.ShapeRange[1]);
            if (string.IsNullOrEmpty(smiles))
            {
                MessageBox.Show(
                    "Could not find SMILES structure from the selected image.\nPlease select a molecular structure picture first.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Clipboard.SetText(smiles!);
            MessageBox.Show(
                $"SMILES copied to clipboard:\n\n{smiles}",
                "SMILES Extracted",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while extracting SMILES:\n" + ex.Message,
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
                    "No molecular structure image selected.\nPlease select a molecular structure picture first.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var ket = ExtractStructure(sel.ShapeRange[1]);
            if (string.IsNullOrEmpty(ket))
            {
                MessageBox.Show(
                    "Could not find canvas structure (KET) from the selected image.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Clipboard.SetText(ket!);
            MessageBox.Show(
                "Canvas data (KET) copied to clipboard.",
                "KET Copied",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while copying canvas data:\n" + ex.Message,
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void QuickInsertFromClipboard(MSPowerPoint.Application app)
    {
        try
        {
            var slide = EnsureActiveSlide(app);
            if (slide == null)
            {
                MessageBox.Show(
                    "No active slide found.\nPlease select a slide to insert the molecule structure first.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // 1. Check if text is selected inside a slide text frame
            try
            {
                var sel = app.ActiveWindow?.Selection;
                if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
                {
                    var selectedText = sel.TextRange.Text?.Trim();
                    if (!string.IsNullOrWhiteSpace(selectedText) &&
                        MoleculeRenderer.TryRenderStructure(selectedText, RenderOptions.Default, out var selPng, out var selKet, out var selSmiles) &&
                        selPng != null)
                    {
                        InsertMoleculeAsPicture(app, slide, selKet, selSmiles ?? "", selPng);
                        return;
                    }
                }
            }
            catch { }

            // 2. Check Windows Clipboard
            if (Clipboard.ContainsText())
            {
                var clipText = Clipboard.GetText()?.Trim();
                if (MoleculeRenderer.TryRenderStructure(clipText, RenderOptions.Default, out var clipPng, out var clipKet, out var clipSmiles) && clipPng != null)
                {
                    InsertMoleculeAsPicture(app, slide, clipKet, clipSmiles ?? "", clipPng);
                    return;
                }
            }

            MessageBox.Show(
                "No valid SMILES or KET structure data found in clipboard or selected text.\nPlease copy a SMILES string (e.g. CC(=O)Oc1ccccc1C(=O)O) first.",
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while inserting structure from clipboard:\n" + ex.Message,
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void CopySelectedMoleculeInfo(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            string? targetData = null;

            if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                targetData = ExtractSmiles(shape) ?? ExtractStructure(shape);
            }
            else if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
            {
                targetData = sel.TextRange.Text?.Trim();
            }

            var info = MoleculeRenderer.TryGetMoleculeInfo(targetData);
            if (info == null)
            {
                MessageBox.Show(
                    "Could not calculate formula/molecular weight from the selected structure.\nPlease select a molecular structure picture on the slide first.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Clipboard.SetText(info.FormatClipboardSummary());
            MessageBox.Show(
                $"Molecular properties copied to clipboard:\n\n{info.FormatDisplayMessage()}",
                "Molecule Info (Formula / MW)",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while calculating molecule info:\n" + ex.Message,
                "SMILES for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static MSPowerPoint.Slide? EnsureActiveSlide(MSPowerPoint.Application app)
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
