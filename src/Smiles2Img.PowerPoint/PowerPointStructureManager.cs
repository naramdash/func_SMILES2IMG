using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Editor;
using Smiles2Img.Notifications;
using Smiles2Img.Rendering;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Smiles2Img.PowerPoint;

public static partial class PowerPointStructureManager
{
    public static void InsertMolecule(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                var targetShape = sel.ShapeRange[1];
                if (!string.IsNullOrEmpty(ExtractStructure(targetShape)))
                {
                    EditSelectedMolecule(app, targetShape);
                    return;
                }
            }

            var slide = EnsureActiveSlide(app);
            if (slide == null)
            {
                MessageBox.Show(
                    "No active slide found.\nPlease select a slide to insert the molecule structure first.",
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = KetcherForm.Rent("PowerPoint", "");
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
                "func_SMILES2IMG",
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
                        "func_SMILES2IMG",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                targetShape = sel.ShapeRange[1];
            }

            if (StructureMetadata.IsMolstar(targetShape.AlternativeText, targetShape.Title))
            {
                EditSelectedMolstar(app, targetShape);
                return;
            }

            var currentStructure = ExtractStructure(targetShape);
            if (string.IsNullOrEmpty(currentStructure))
            {
                MessageBox.Show(
                    "The selected object is not a SMILES/KET molecule structure image.",
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = KetcherForm.Rent("PowerPoint", currentStructure ?? "");
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
                "func_SMILES2IMG",
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
                    if (StructureMetadata.IsMolstar(shape.AlternativeText, shape.Title))
                    {
                        EditSelectedMolstar(sel.Application, shape);
                    }
                    else
                    {
                        EditSelectedMolecule(sel.Application, shape);
                    }
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

    private static void ReplaceShape(MSPowerPoint.Shape oldShape, string? newKet, string newSmiles, byte[] pngBytes, bool selectNew = true)
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
            if (selectNew)
            {
                try { newShape.Select(); } catch { }
            }
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
            if (StructureMetadata.IsMolstar(shape.AlternativeText, shape.Title))
                return shape.AlternativeText;
        }
        catch { }

        try
        {
            var molTag = shape.Tags["MOLSTAR"];
            if (!string.IsNullOrWhiteSpace(molTag)) return shape.AlternativeText;
        }
        catch { }

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

    private static void GetSelectedShapeOrText(MSPowerPoint.Application app, out MSPowerPoint.Shape? shape, out string? selectedText)
    {
        shape = null;
        selectedText = null;
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel == null) return;

            if (sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                shape = sel.ShapeRange[1];
            }
            else if (sel.Type == MSPowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
            {
                selectedText = sel.TextRange.Text?.Trim();
            }
        }
        catch { }
    }

    public static void CopySelectedMoleculeSmiles(MSPowerPoint.Application app)
    {
        GetSelectedShapeOrText(app, out var shape, out var selectedText);
        if (shape == null && string.IsNullOrWhiteSpace(selectedText))
        {
            MoleculeNotification.NotifyWarning("SMILES", "Please select a molecular structure picture or text first.");
            return;
        }
        MoleculeActionHelper.CopySmiles(
            MoleculeActionHelper.ResolveSmilesWithTextFallback(shape != null ? ExtractSmiles(shape) : null, selectedText),
            "Could not find SMILES structure from the selected image or text.");
    }

    public static void CopySelectedMoleculeInchi(MSPowerPoint.Application app)
    {
        GetSelectedShapeOrText(app, out var shape, out var selectedText);
        var targetData = shape != null ? (ExtractSmiles(shape) ?? ExtractStructure(shape)) : selectedText;
        MoleculeActionHelper.CopyInchi(
            targetData,
            "Could not generate InChI from the selected structure or text.");
    }

    public static void CopySelectedMoleculeKet(MSPowerPoint.Application app)
    {
        GetSelectedShapeOrText(app, out var shape, out var selectedText);
        if (shape == null && string.IsNullOrWhiteSpace(selectedText))
        {
            MoleculeNotification.NotifyWarning("Canvas Data (KET)", "Please select a molecular structure picture or text first.");
            return;
        }
        MoleculeActionHelper.CopyKet(
            MoleculeActionHelper.ResolveKetWithTextFallback(shape != null ? ExtractStructure(shape) : null, selectedText),
            "Could not find canvas structure (KET) from the selected image or text.");
    }

    public static void CopySelectedMoleculeSvg(MSPowerPoint.Application app)
    {
        GetSelectedShapeOrText(app, out var shape, out var selectedText);
        var targetData = shape != null ? (ExtractStructure(shape) ?? ExtractSmiles(shape)) : selectedText;
        if (string.IsNullOrEmpty(targetData))
        {
            MoleculeNotification.NotifyWarning("SVG Vector", "Please select a molecular structure picture or text first.");
            return;
        }
        MoleculeActionHelper.CopySvg(
            targetData,
            "Could not generate SVG vector image from the selected structure.");
    }

    public static void ApplyBatchStyle(MSPowerPoint.Application app, string background, string style, string presetLabel)
    {
        try
        {
            if (app.Presentations.Count == 0)
            {
                MoleculeNotification.NotifyWarning("Batch Style", "No active PowerPoint presentation found.");
                return;
            }

            var opts = RenderOptions.Parse(background, style, null);
            var shapesToUpdate = new System.Collections.Generic.List<MSPowerPoint.Shape>();

            try
            {
                var sel = app.ActiveWindow?.Selection;
                if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                    foreach (MSPowerPoint.Shape s in sel.ShapeRange)
                    {
                        if (!StructureMetadata.IsMolstar(s.AlternativeText, s.Title) && !string.IsNullOrEmpty(ExtractStructure(s)))
                            shapesToUpdate.Add(s);
                    }
                }
            }
            catch { }
            if (shapesToUpdate.Count == 0)
            foreach (MSPowerPoint.Slide slide in app.ActivePresentation.Slides)
            {
                foreach (MSPowerPoint.Shape s in slide.Shapes)
                {
                    if (!StructureMetadata.IsMolstar(s.AlternativeText, s.Title) && !string.IsNullOrEmpty(ExtractStructure(s)))
                        shapesToUpdate.Add(s);
                }
            }

            int updated = 0;
            foreach (var s in shapesToUpdate)
            {
                var raw = ExtractStructure(s);
                var smiles = ExtractSmiles(s) ?? "";
                if (!string.IsNullOrEmpty(raw) && MoleculeRenderer.TryRenderStructure(raw, opts, out var pngBytes, out var ket, out var resolvedSmiles) && pngBytes != null)
                {
                    ReplaceShape(s, ket ?? (raw!.TrimStart().StartsWith("{") ? raw : null), !string.IsNullOrEmpty(smiles) ? smiles : (resolvedSmiles ?? ""), pngBytes, selectNew: false);
                    updated++;
                }
            }

            if (updated > 0)
                MoleculeNotification.NotifyCopied("Batch Style Applied", $"Updated {updated} structure(s) to {presetLabel}");
            else
                MoleculeNotification.NotifyWarning("Batch Style", "No SMILES/KET structures found in the active presentation.");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Batch Style", "Error applying batch style: " + ex.Message);
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
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            GetSelectedShapeOrText(app, out _, out var selectedText);
            if (!string.IsNullOrWhiteSpace(selectedText) &&
                MoleculeRenderer.TryRenderStructure(selectedText, RenderOptions.Default, out var selPng, out var selKet, out var selSmiles) &&
                selPng != null)
            {
                InsertMoleculeAsPicture(app, slide, selKet, selSmiles ?? "", selPng);
                return;
            }

            if (SafeClipboard.ContainsText())
            {
                var clipText = SafeClipboard.GetText()?.Trim();
                if (MoleculeRenderer.TryRenderStructure(clipText, RenderOptions.Default, out var clipPng, out var clipKet, out var clipSmiles) && clipPng != null)
                {
                    InsertMoleculeAsPicture(app, slide, clipKet, clipSmiles ?? "", clipPng);
                    return;
                }
            }

            MessageBox.Show(
                "No valid SMILES or KET structure data found in clipboard or selected text.\nPlease copy a SMILES string (e.g. CC(=O)Oc1ccccc1C(=O)O) first.",
                "func_SMILES2IMG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while inserting structure from clipboard:\n" + ex.Message,
                "func_SMILES2IMG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void CopySelectedMoleculeInfo(MSPowerPoint.Application app) => ShowSelectedMoleculeInfo(app);

    public static MoleculeInfo? GetSelectedMoleculeInfo(MSPowerPoint.Application app, bool resolvePubChem = false)
    {
        GetSelectedShapeOrText(app, out var shape, out var selectedText);
        if (shape != null)
        {
            var molstarInfo = MoleculeActionHelper.TryResolveMolstarInfo(shape.AlternativeText, shape.Title, resolvePubChem);
            if (molstarInfo != null) return molstarInfo;
        }

        var targetData = shape != null ? (ExtractSmiles(shape) ?? ExtractStructure(shape)) : selectedText;
        return string.IsNullOrEmpty(targetData) ? null : MoleculeRenderer.TryGetMoleculeInfo(targetData, resolvePubChem);
    }

    public static void CopySelectedSubstanceName(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopySubstanceName(r => GetSelectedMoleculeInfo(app, r), "Could not resolve substance name from the selected structure.");

    public static void CopySelectedFormula(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopyFormula(r => GetSelectedMoleculeInfo(app, r), "Could not compute formula from the selected structure.");

    public static void CopySelectedMolecularWeight(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopyMolecularWeight(r => GetSelectedMoleculeInfo(app, r), "Could not compute molecular weight from the selected structure.");

    public static void CopySelectedInchiKey(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopyInchiKey(r => GetSelectedMoleculeInfo(app, r), "Could not generate InChIKey from the selected structure.");

    public static void CopySelectedAllSummary(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.ExcelTsvWithHeader, "Molecule Info", "Copied summary", "Could not retrieve molecule info from the selected structure.", "molecule summary");

    public static void CopySelectedTsvRow(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.ExcelTsvRow, "Excel Row (TSV)", "Copied row", "Could not retrieve molecule info from the selected structure.", "TSV row");

    public static void CopySelectedMultiLine(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.MultiLine, "Multi-Line List", "Copied list", "Could not retrieve molecule info from the selected structure.", "multi-line list");

    public static void CopySelectedMarkdown(MSPowerPoint.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.MarkdownTable, "Markdown Table", "Copied table", "Could not retrieve molecule info from the selected structure.", "markdown table");

    public static void ShowSelectedMoleculeInfo(MSPowerPoint.Application app) =>
        MoleculeActionHelper.ShowMoleculeInfoDialog(
            r => GetSelectedMoleculeInfo(app, r),
            GetHwnd(app),
            "Could not retrieve molecule properties from the selected structure.\nPlease select a molecular structure picture or text on the slide first.");

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
