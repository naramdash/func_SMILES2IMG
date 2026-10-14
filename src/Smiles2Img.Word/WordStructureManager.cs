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
                "An error occurred while inserting molecule:\n" + ex.Message,
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
                    "No molecular structure image selected.\nPlease select a molecular structure picture first.",
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
                "An error occurred while editing molecule:\n" + ex.Message,
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void InsertMolstarStructure(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            if (sel != null)
            {
                if (sel.InlineShapes.Count > 0 && StructureMetadata.IsMolstar(sel.InlineShapes[1].AlternativeText, sel.InlineShapes[1].Title))
                {
                    EditSelectedMolstar(app);
                    return;
                }
                if (sel.ShapeRange.Count > 0 && StructureMetadata.IsMolstar(sel.ShapeRange[1].AlternativeText, sel.ShapeRange[1].Title))
                {
                    EditSelectedMolstar(app);
                    return;
                }
            }

            if (app.Documents.Count == 0)
            {
                app.Documents.Add();
            }

            var hwnd = GetHwnd(app);
            using var form = new MolstarForm("Word", "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (form.ResultImageBytes != null || form.ResultGlbBytes != null))
            {
                ApplyMolstarInsert(app, form);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while opening Mol* 3D Viewer:\n" + ex.Message,
                "Mol* 3D for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void EditSelectedMolstar(MSWord.Application app)
    {
        try
        {
            var sel = app.Selection;
            MSWord.InlineShape? targetInline = null;
            MSWord.Shape? targetShape = null;
            string? altText = null;
            string? title = null;

            if (sel.InlineShapes.Count > 0)
            {
                targetInline = sel.InlineShapes[1];
                altText = targetInline.AlternativeText;
                title = targetInline.Title;
            }
            else if (sel.ShapeRange.Count > 0)
            {
                targetShape = sel.ShapeRange[1];
                altText = targetShape.AlternativeText;
                title = targetShape.Title;
            }

            var meta = MolstarSessionManager.ParseAltText(altText, title);
            if (meta == null)
            {
                MessageBox.Show(
                    "The selected object is not a Mol* 3D structure object.\nPlease select a Mol* structure image or 3D model first.",
                    "Mol* 3D for Word",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string? snapshotJson = MolstarSessionManager.LoadLocalCache(meta.Id);
            if (string.IsNullOrEmpty(snapshotJson) && app.ActiveDocument != null)
            {
                try
                {
                    snapshotJson = MolstarSessionManager.LoadFromCustomXmlParts(app.ActiveDocument, meta.Id);
                }
                catch { }
            }

            var hwnd = GetHwnd(app);
            using var form = new MolstarForm("Word", meta.Pdb, snapshotJson);
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (form.ResultImageBytes != null || form.ResultGlbBytes != null))
            {
                if (targetInline != null)
                {
                    ApplyMolstarReplaceInline(app, targetInline, form, meta);
                }
                else if (targetShape != null)
                {
                    ApplyMolstarReplaceShape(app, targetShape, form, meta);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while editing Mol* 3D structure:\n" + ex.Message,
                "Mol* 3D for Word",
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

    private static void ApplyMolstarInsert(MSWord.Application app, MolstarForm form)
    {
        var sessionId = Guid.NewGuid().ToString("N");
        var meta = new MolstarMetadata
        {
            Id = sessionId,
            Pdb = form.ResultPdbId,
            Mode = form.ResultMode,
            Embed = form.ResultEmbedInDoc,
            Background = form.ResultBackground
        };

        if (!string.IsNullOrEmpty(form.ResultSnapshotJson))
        {
            MolstarSessionManager.SaveLocalCache(sessionId, form.ResultSnapshotJson!);
            if (form.ResultEmbedInDoc && app.ActiveDocument != null)
            {
                MolstarSessionManager.SaveToCustomXmlParts(app.ActiveDocument, sessionId, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        dynamic? newShape = null;
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                dynamic inlineShapes = app.Selection.InlineShapes;
                newShape = inlineShapes.Add3DModel(tempGlb, false, true, app.Selection.Range);
            }
            catch
            {
                meta.Mode = "image";
            }
            finally
            {
                try { if (File.Exists(tempGlb)) File.Delete(tempGlb); } catch { }
            }
        }

        if (newShape == null && form.ResultImageBytes != null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes);
                var shape = app.Selection.InlineShapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: false,
                    SaveWithDocument: true);

                var (w, h) = StructureMetadata.GetScaledDimensions(form.ResultImageBytes!, 320f, 240f, scale: 0.85f);
                shape.Width = w;
                shape.Height = h;
                newShape = shape;
            }
            finally
            {
                try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
            }
        }

        if (newShape != null)
        {
            newShape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
            newShape.Title = $"Mol* 3D: {(string.IsNullOrEmpty(form.ResultPdbId) ? "Structure" : form.ResultPdbId)}";
            newShape.Select();
        }
    }

    private static void ApplyMolstarReplaceInline(MSWord.Application app, MSWord.InlineShape oldShape, MolstarForm form, MolstarMetadata oldMeta)
    {
        var sessionId = oldMeta.Id;
        if (string.IsNullOrEmpty(sessionId)) sessionId = Guid.NewGuid().ToString("N");

        var meta = new MolstarMetadata
        {
            Id = sessionId,
            Pdb = form.ResultPdbId,
            Mode = form.ResultMode,
            Embed = form.ResultEmbedInDoc,
            Background = form.ResultBackground
        };

        if (!string.IsNullOrEmpty(form.ResultSnapshotJson))
        {
            MolstarSessionManager.SaveLocalCache(sessionId, form.ResultSnapshotJson!);
            if (form.ResultEmbedInDoc && app.ActiveDocument != null)
            {
                MolstarSessionManager.SaveToCustomXmlParts(app.ActiveDocument, sessionId, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        var oldWidth = oldShape.Width;
        var range = oldShape.Range;
        oldShape.Delete();

        dynamic? newShape = null;
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                dynamic inlineShapes = range.InlineShapes;
                newShape = inlineShapes.Add3DModel(tempGlb, false, true, range);
            }
            catch
            {
                meta.Mode = "image";
            }
            finally
            {
                try { if (File.Exists(tempGlb)) File.Delete(tempGlb); } catch { }
            }
        }

        if (newShape == null && form.ResultImageBytes != null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes);
                var shape = range.InlineShapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: false,
                    SaveWithDocument: true);

                var (_, h) = StructureMetadata.GetScaledDimensions(form.ResultImageBytes!, oldWidth, oldWidth, minWidth: oldWidth, maxWidth: oldWidth);
                shape.Width = oldWidth;
                shape.Height = h;
                newShape = shape;
            }
            finally
            {
                try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
            }
        }

        if (newShape != null)
        {
            newShape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
            newShape.Title = $"Mol* 3D: {(string.IsNullOrEmpty(form.ResultPdbId) ? "Structure" : form.ResultPdbId)}";
            newShape.Select();
        }
    }

    private static void ApplyMolstarReplaceShape(MSWord.Application app, MSWord.Shape oldShape, MolstarForm form, MolstarMetadata oldMeta)
    {
        var sessionId = oldMeta.Id;
        if (string.IsNullOrEmpty(sessionId)) sessionId = Guid.NewGuid().ToString("N");

        var meta = new MolstarMetadata
        {
            Id = sessionId,
            Pdb = form.ResultPdbId,
            Mode = form.ResultMode,
            Embed = form.ResultEmbedInDoc,
            Background = form.ResultBackground
        };

        if (!string.IsNullOrEmpty(form.ResultSnapshotJson))
        {
            MolstarSessionManager.SaveLocalCache(sessionId, form.ResultSnapshotJson!);
            if (form.ResultEmbedInDoc && app.ActiveDocument != null)
            {
                MolstarSessionManager.SaveToCustomXmlParts(app.ActiveDocument, sessionId, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        var doc = (MSWord.Document)oldShape.Parent;
        float oldLeft = oldShape.Left;
        float oldTop = oldShape.Top;
        float oldWidth = oldShape.Width;
        float oldHeight = oldShape.Height;
        oldShape.Delete();

        dynamic? newShape = null;
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                dynamic shapes = doc.Shapes;
                newShape = shapes.Add3DModel(tempGlb, false, true, oldLeft, oldTop, oldWidth, oldHeight);
            }
            catch
            {
                meta.Mode = "image";
            }
            finally
            {
                try { if (File.Exists(tempGlb)) File.Delete(tempGlb); } catch { }
            }
        }

        if (newShape == null && form.ResultImageBytes != null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes);
                var (left, top, w, h) = StructureMetadata.FitReplacementBounds(oldLeft, oldTop, oldWidth, oldHeight, form.ResultImageBytes!);
                newShape = doc.Shapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: false,
                    SaveWithDocument: true,
                    Left: left,
                    Top: top,
                    Width: w,
                    Height: h);
            }
            finally
            {
                try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
            }
        }

        if (newShape != null)
        {
            newShape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
            newShape.Title = $"Mol* 3D: {(string.IsNullOrEmpty(form.ResultPdbId) ? "Structure" : form.ResultPdbId)}";
            newShape.Select();
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
                    "Could not find SMILES structure from the selected image.\nPlease select a molecular structure picture first.",
                    "SMILES for Word",
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
                var raw = ExtractStructure(targetInline.AlternativeText, targetInline.Title);
                ket = MoleculeRenderer.TryConvertToKet(raw);
            }
            else if (sel.ShapeRange.Count > 0)
            {
                var targetShape = sel.ShapeRange[1];
                var raw = ExtractStructure(targetShape.AlternativeText, targetShape.Title);
                ket = MoleculeRenderer.TryConvertToKet(raw);
            }

            if (string.IsNullOrEmpty(ket))
            {
                MessageBox.Show(
                    "Could not find canvas structure (KET) from the selected image.",
                    "SMILES for Word",
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
                "No valid SMILES or KET structure data found in selection or clipboard.\nPlease copy a SMILES string (e.g. CC(=O)Oc1ccccc1C(=O)O) or select text in the document first.",
                "SMILES for Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while inserting structure from clipboard:\n" + ex.Message,
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
                    "Could not calculate formula/molecular weight from the selected structure.\nPlease select a molecular structure picture first.",
                    "SMILES for Word",
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
