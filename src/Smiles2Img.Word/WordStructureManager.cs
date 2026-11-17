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
    private static bool _isDoubleClickHandling;

    private static string? CleanWordText(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var cleaned = raw!.Trim(' ', '\t', '\r', '\n', '\v', '\f', '\a');
        return string.IsNullOrWhiteSpace(cleaned) ? null : cleaned;
    }

    private static void SetLockAspectRatio<T>(T target) where T : class
    {
        try
        {
            var prop = typeof(T).GetProperty("LockAspectRatio");
            if (prop == null) return;
            object value = prop.PropertyType.IsEnum ? Enum.ToObject(prop.PropertyType, -1) : -1;
            prop.SetValue(target, value, null);
        }
        catch { }
    }

    private static void LockShapeAspectRatio(MSWord.InlineShape shape) => SetLockAspectRatio(shape);

    private static void LockShapeAspectRatio(MSWord.Shape shape) => SetLockAspectRatio(shape);

    private static void LockShapeAspectRatio(object shape)
    {
        if (shape is MSWord.InlineShape inlineShape)
        {
            SetLockAspectRatio(inlineShape);
            return;
        }
        if (shape is MSWord.Shape floatShape)
        {
            SetLockAspectRatio(floatShape);
            return;
        }
    }

    private static bool TryGetSelectedGraphic(
        MSWord.Selection? sel,
        out MSWord.InlineShape? inlineShape,
        out MSWord.Shape? floatShape,
        out string? altText,
        out string? title)
    {
        inlineShape = null;
        floatShape = null;
        altText = null;
        title = null;
        if (sel == null) return false;

        if (sel.InlineShapes.Count > 0)
        {
            inlineShape = sel.InlineShapes[1];
            altText = inlineShape.AlternativeText;
            title = inlineShape.Title;
            return true;
        }

        if (sel.ShapeRange.Count > 0)
        {
            floatShape = sel.ShapeRange[1];
            altText = floatShape.AlternativeText;
            title = floatShape.Title;
            return true;
        }

        return false;
    }

    private static bool TryOpenGraphicForDoubleClick(MSWord.Application app, string? altText, string? title)
    {
        if (StructureMetadata.IsMolstar(altText, title))
        {
            EditSelectedMolstar(app);
            return true;
        }
        if (!string.IsNullOrEmpty(ExtractStructure(altText, title)))
        {
            EditSelectedMolecule(app);
            return true;
        }
        return false;
    }

    public static void InsertMolecule(MSWord.Application app)
    {
        try
        {
            if (TryGetSelectedGraphic(app.Selection, out _, out _, out var altText, out var title) &&
                TryOpenGraphicForDoubleClick(app, altText, title))
            {
                return;
            }

            if (app.Documents.Count == 0)
            {
                app.Documents.Add();
            }

            GetSelectionTexts(app, out _, out _, out var selText);
            var initialStructure = StructureMetadata.ExtractPreloadForKetcher(selText);
            var hwnd = GetHwnd(app);
            using var form = KetcherForm.Rent("Word", initialStructure);
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
            TryGetSelectedGraphic(app.Selection, out var targetInline, out var targetShape, out var altText, out var title);
            if (StructureMetadata.IsMolstar(altText, title))
            {
                EditSelectedMolstar(app);
                return;
            }

            var currentStructure = ExtractStructure(altText, title);
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
        if (_isDoubleClickHandling) return;
        try
        {
            if (TryGetSelectedGraphic(sel, out _, out _, out var altText, out var title) &&
                TryOpenGraphicForDoubleClick(sel.Application, altText, title))
            {
                cancel = true;
            }
        }
        catch { }
    }

    /// <summary>
    /// Intercepts WM_LBUTTONDBLCLK on floating Word shapes and 3D models whose double-click event
    /// is otherwise swallowed by Word's Picture/3D Format pane.
    /// </summary>
    public static bool TryHandleFloatingDoubleClick(MSWord.Application app)
    {
        if (_isDoubleClickHandling) return false;
        try
        {
            var sel = app.Selection;
            if (TryGetSelectedGraphic(sel, out var inlineShape, out _, out var altText, out var title) &&
                inlineShape == null)
            {
                _isDoubleClickHandling = true;
                try
                {
                    return TryOpenGraphicForDoubleClick(app, altText, title);
                }
                finally
                {
                    _isDoubleClickHandling = false;
                }
            }
        }
        catch { }
        return false;
    }

    /// <summary>
    /// Resolves the InlineShapes target and insertion range, anchoring to the container paragraph
    /// when a floating Shape or Text Box is currently selected.
    /// </summary>

    internal static (MSWord.InlineShapes InlineShapes, MSWord.Range Range) ResolveInsertTarget(MSWord.Application app)
    {
        var sel = app.Selection;
        if (sel.Type == MSWord.WdSelectionType.wdSelectionShape && sel.ShapeRange.Count > 0)
        {
            var anchor = sel.ShapeRange[1].Anchor;
            return (anchor.InlineShapes, anchor);
        }
        return (sel.InlineShapes, sel.Range);
    }

    private static void InsertMoleculeAsPicture(MSWord.Application app, string? ket, string smiles, byte[] pngBytes)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"structure_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(tempFile, pngBytes);
            var (inlineShapes, _) = ResolveInsertTarget(app);
            var shape = inlineShapes.AddPicture(
                FileName: tempFile,
                LinkToFile: false,
                SaveWithDocument: true);

            shape.AlternativeText = StructureMetadata.FormatAlternativeText(ket, smiles);
            shape.Title = StructureMetadata.FormatTitle(smiles);
            LockShapeAspectRatio(shape);
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
            LockShapeAspectRatio(newShape);
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
            LockShapeAspectRatio(newShape);
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

            if (TryGetSelectedGraphic(sel, out var inlineShape, out var shape, out altText, out title))
            {
                if (inlineShape != null || shape == null)
                    return;

                if (sel.Start != sel.End && !string.IsNullOrWhiteSpace(sel.Text))
                {
                    selectedText = CleanWordText(sel.Text);
                }
                else if (string.IsNullOrEmpty(ExtractStructure(altText, title)) &&
                         string.IsNullOrEmpty(ExtractSmiles(title, altText)))
                {
                    try
                    {
                        if (shape.TextFrame.HasText != 0)
                            selectedText = CleanWordText(shape.TextFrame.TextRange.Text);
                    }
                    catch { }
                }
                return;
            }

            if (sel.Start != sel.End && !string.IsNullOrWhiteSpace(sel.Text))
            {
                selectedText = CleanWordText(sel.Text);
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
        var targetData = MoleculeActionHelper.ResolveStructureOrText(ExtractSmiles(title, altText), ExtractStructure(altText, title), selectedText);
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
            bool TryInsertCandidate(string? raw, bool clearSelection = false)
            {
                var candidate = CleanWordText(raw);
                if (candidate == null ||
                    !MoleculeRenderer.TryRenderStructure(candidate, RenderOptions.Default, out var png, out var ket, out var smiles) ||
                    png == null)
                {
                    return false;
                }

                if (clearSelection && sel != null)
                    sel.Text = "";

                InsertMoleculeAsPicture(app, ket, smiles ?? "", png);
                return true;
            }

            if (sel != null && sel.Start != sel.End && TryInsertCandidate(sel.Text, clearSelection: true))
                return;

            GetSelectionTexts(app, out _, out _, out var shapeText);
            if (TryInsertCandidate(shapeText))
                return;

            if (SafeClipboard.ContainsText() && TryInsertCandidate(SafeClipboard.GetText()))
                return;

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

        var targetData = MoleculeActionHelper.ResolveStructureOrText(ExtractSmiles(title, altText), ExtractStructure(altText, title), selectedText);
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
        var targetData = MoleculeActionHelper.ResolveStructureOrText(ExtractStructure(altText, title), ExtractSmiles(title, altText), selectedText);
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
