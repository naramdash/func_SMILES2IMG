using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Editor;
using Smiles2Img.Notifications;
using Smiles2Img.Rendering;
using MSWord = Microsoft.Office.Interop.Word;

namespace Smiles2Img.Word;

public static partial class WordStructureManager
{
    public static void InsertMolecule(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            if (sel != null)
            {
                if (sel.InlineShapes.Count > 0 && !string.IsNullOrEmpty(ExtractStructure(sel.InlineShapes[1].AlternativeText, sel.InlineShapes[1].Title)))
                {
                    EditSelectedMolecule(app);
                    return;
                }
                if (sel.ShapeRange.Count > 0 && !string.IsNullOrEmpty(ExtractStructure(sel.ShapeRange[1].AlternativeText, sel.ShapeRange[1].Title)))
                {
                    EditSelectedMolecule(app);
                    return;
                }
            }

            if (app.Documents.Count == 0)
            {
                app.Documents.Add();
            }

            var hwnd = GetHwnd(app);
            using var form = KetcherForm.Rent("Word", "");
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
                "An error occurred while inserting molecule:\n" + ex.Message,
                "func_SMILES2IMG",
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
                    "No molecular structure image selected.\nPlease select a molecular structure picture first.",
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = KetcherForm.Rent("Word", currentStructure ?? "");
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
                "An error occurred while editing molecule:\n" + ex.Message,
                "func_SMILES2IMG",
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
                if (StructureMetadata.IsMolstar(shape.AlternativeText, shape.Title))
                {
                    cancel = true;
                    EditSelectedMolstar(sel.Application);
                    return;
                }
                if (!string.IsNullOrEmpty(ExtractStructure(shape.AlternativeText, shape.Title)))
                {
                    cancel = true;
                    EditSelectedMolecule(sel.Application);
                    return;
                }
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                if (StructureMetadata.IsMolstar(shape.AlternativeText, shape.Title))
                {
                    cancel = true;
                    EditSelectedMolstar(sel.Application);
                    return;
                }
                if (!string.IsNullOrEmpty(ExtractStructure(shape.AlternativeText, shape.Title)))
                {
                    cancel = true;
                    EditSelectedMolecule(sel.Application);
                    return;
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

    private static void GetSelectionTexts(MSWord.Application app, out string? altText, out string? title, out string? selectedText)
    {
        altText = null;
        title = null;
        selectedText = null;

        try
        {
            var sel = app.Selection;
            if (sel == null) return;

            if (sel.InlineShapes.Count > 0)
            {
                var inlineShape = sel.InlineShapes[1];
                altText = inlineShape.AlternativeText;
                title = inlineShape.Title;
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                altText = shape.AlternativeText;
                title = shape.Title;
            }
            else if (sel.Start != sel.End && !string.IsNullOrWhiteSpace(sel.Text))
            {
                selectedText = sel.Text.Trim();
            }
        }
        catch { }
    }

    public static void CopySelectedMoleculeSmiles(MSWord.Application app)
    {
        GetSelectionTexts(app, out var altText, out var title, out var selectedText);
        MoleculeActionHelper.CopySmiles(
            MoleculeActionHelper.ResolveSmilesWithTextFallback(ExtractSmiles(title, altText), selectedText),
            "Could not find SMILES structure from the selected image or text.");
    }

    public static void CopySelectedMoleculeInchi(MSWord.Application app)
    {
        GetSelectionTexts(app, out var altText, out var title, out var selectedText);
        var targetData = ExtractSmiles(title, altText) ?? ExtractStructure(altText, title) ?? selectedText;
        MoleculeActionHelper.CopyInchi(
            targetData,
            "Could not generate InChI from the selected structure or text.");
    }

    public static void CopySelectedMoleculeKet(MSWord.Application app)
    {
        GetSelectionTexts(app, out var altText, out var title, out var selectedText);
        MoleculeActionHelper.CopyKet(
            MoleculeActionHelper.ResolveKetWithTextFallback(ExtractStructure(altText, title), selectedText),
            "Could not find canvas structure (KET) from the selected image or text.");
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

            if (SafeClipboard.ContainsText())
            {
                var clipText = SafeClipboard.GetText()?.Trim();
                if (MoleculeRenderer.TryRenderStructure(clipText, RenderOptions.Default, out var clipPng, out var clipKet, out var clipSmiles) && clipPng != null)
                {
                    InsertMoleculeAsPicture(app, clipKet, clipSmiles ?? "", clipPng);
                    return;
                }
            }

            MessageBox.Show(
                "No valid SMILES or KET structure data found in selection or clipboard.\nPlease copy a SMILES string (e.g. CC(=O)Oc1ccccc1C(=O)O) or select text in the document first.",
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

    public static void CopySelectedMoleculeInfo(MSWord.Application app) => ShowSelectedMoleculeInfo(app);

    public static MoleculeInfo? GetSelectedMoleculeInfo(MSWord.Application app, bool resolvePubChem = false)
    {
        GetSelectionTexts(app, out var altText, out var title, out var selectedText);
        if (altText != null || title != null)
        {
            var molstarInfo = MoleculeActionHelper.TryResolveMolstarInfo(altText, title, resolvePubChem);
            if (molstarInfo != null) return molstarInfo;
        }

        var targetData = ExtractSmiles(title, altText) ?? ExtractStructure(altText, title) ?? selectedText;
        return string.IsNullOrEmpty(targetData) ? null : MoleculeRenderer.TryGetMoleculeInfo(targetData, resolvePubChem);
    }

    public static void CopySelectedSubstanceName(MSWord.Application app) =>
        MoleculeActionHelper.CopySubstanceName(r => GetSelectedMoleculeInfo(app, r), "Could not resolve substance name from the selected structure.");

    public static void CopySelectedFormula(MSWord.Application app) =>
        MoleculeActionHelper.CopyFormula(r => GetSelectedMoleculeInfo(app, r), "Could not compute formula from the selected structure.");

    public static void CopySelectedMolecularWeight(MSWord.Application app) =>
        MoleculeActionHelper.CopyMolecularWeight(r => GetSelectedMoleculeInfo(app, r), "Could not compute molecular weight from the selected structure.");

    public static void CopySelectedInchiKey(MSWord.Application app) =>
        MoleculeActionHelper.CopyInchiKey(r => GetSelectedMoleculeInfo(app, r), "Could not generate InChIKey from the selected structure.");

    public static void CopySelectedAllSummary(MSWord.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.ExcelTsvWithHeader, "Molecule Info", "Copied summary", "Could not retrieve molecule info from the selected structure.", "molecule summary");

    public static void CopySelectedTsvRow(MSWord.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.ExcelTsvRow, "Excel Row (TSV)", "Copied row", "Could not retrieve molecule info from the selected structure.", "TSV row");

    public static void CopySelectedMultiLine(MSWord.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.MultiLine, "Multi-Line List", "Copied list", "Could not retrieve molecule info from the selected structure.", "multi-line list");

    public static void CopySelectedMarkdown(MSWord.Application app) =>
        MoleculeActionHelper.CopySummary(r => GetSelectedMoleculeInfo(app, r), MoleculeSummaryFormat.MarkdownTable, "Markdown Table", "Copied table", "Could not retrieve molecule info from the selected structure.", "markdown table");

    public static void ShowSelectedMoleculeInfo(MSWord.Application app) =>
        MoleculeActionHelper.ShowMoleculeInfoDialog(
            r => GetSelectedMoleculeInfo(app, r),
            GetHwnd(app),
            "Could not retrieve molecule properties from the selected structure.\nPlease select a molecular structure picture or text first.");

    public static void CopySelectedMoleculeSvg(MSWord.Application app)
    {
        GetSelectionTexts(app, out var altText, out var title, out var selectedText);
        var targetData = ExtractStructure(altText, title) ?? ExtractSmiles(title, altText) ?? selectedText;
        MoleculeActionHelper.CopySvg(
            targetData,
            "Could not render SVG from the selected structure or text.",
            title: "Vector SVG",
            successMessage: "Scalable SVG copied to clipboard");
    }

    public static int ApplyBatchStyle(MSWord.Application app, string background, string style, string presetLabel = "Preset")
    {
        int updated = 0;
        try
        {
            if (app.Documents.Count == 0 || app.ActiveDocument == null)
                return 0;

            var doc = app.ActiveDocument;
            var options = RenderOptions.Parse(background, style, null);

            var inlineList = new System.Collections.Generic.List<MSWord.InlineShape>();
            var sel = app.Selection;
            bool useSelection = sel != null && (sel.InlineShapes.Count > 0 || sel.ShapeRange.Count > 0);
            foreach (MSWord.InlineShape s in (useSelection ? sel!.InlineShapes : doc.InlineShapes))
            {
                try
                {
                    if (StructureMetadata.IsMolstar(s.AlternativeText, s.Title))
                        continue;
                    if (!string.IsNullOrEmpty(ExtractStructure(s.AlternativeText, s.Title)))
                        inlineList.Add(s);
                }
                catch { }
            }

            foreach (var s in inlineList)
            {
                try
                {
                    var raw = ExtractStructure(s.AlternativeText, s.Title);
                    if (MoleculeRenderer.TryRenderStructure(raw, options, out var pngBytes, out var ket, out var smiles) && pngBytes != null)
                    {
                        ReplaceInlineShape(s, ket, smiles ?? "", pngBytes);
                        updated++;
                    }
                }
                catch { }
            }

            var floatList = new System.Collections.Generic.List<MSWord.Shape>();
            foreach (MSWord.Shape s in (useSelection ? (System.Collections.IEnumerable)sel!.ShapeRange : doc.Shapes))
            {
                try
                {
                    if (StructureMetadata.IsMolstar(s.AlternativeText, s.Title))
                        continue;
                    if (!string.IsNullOrEmpty(ExtractStructure(s.AlternativeText, s.Title)))
                        floatList.Add(s);
                }
                catch { }
            }

            foreach (var s in floatList)
            {
                try
                {
                    var raw = ExtractStructure(s.AlternativeText, s.Title);
                    if (MoleculeRenderer.TryRenderStructure(raw, options, out var pngBytes, out var ket, out var smiles) && pngBytes != null)
                    {
                        ReplaceShape(s, ket, smiles ?? "", pngBytes);
                        updated++;
                    }
                }
                catch { }
            }

            if (updated > 0)
                MoleculeNotification.NotifyCopied("Batch Style", $"Updated {updated} structure(s) to {presetLabel}");
            else
                MoleculeNotification.NotifyWarning("Batch Style", "No SMILES structures found in the active document.");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Batch Style", "Error applying batch style: " + ex.Message);
        }
        return updated;
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
