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

    public static void InsertMolstarStructure(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                var targetShape = sel.ShapeRange[1];
                if (StructureMetadata.IsMolstar(targetShape.AlternativeText, targetShape.Title))
                {
                    EditSelectedMolstar(app, targetShape);
                    return;
                }
            }

            var slide = EnsureActiveSlide(app);
            if (slide == null)
            {
                MessageBox.Show(
                    "No active slide found.\nPlease select a slide to insert the 3D structure first.",
                    "Mol* 3D for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            using var form = new MolstarForm("PowerPoint", "");
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (form.ResultImageBytes != null || form.ResultGlbBytes != null))
            {
                ApplyMolstarInsert(app, slide, form);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while opening Mol* 3D Viewer:\n" + ex.Message,
                "Mol* 3D for PowerPoint",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public static void EditSelectedMolstar(MSPowerPoint.Application app, MSPowerPoint.Shape? targetShape = null)
    {
        try
        {
            if (targetShape == null)
            {
                var sel = app.ActiveWindow?.Selection;
                if (sel == null || sel.Type != MSPowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
                {
                    MessageBox.Show(
                        "No 3D molecular structure selected.\nPlease select a Mol* structure shape or 3D model first.",
                        "Mol* 3D for PowerPoint",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                targetShape = sel.ShapeRange[1];
            }

            var meta = MolstarSessionManager.ParseAltText(targetShape.AlternativeText, targetShape.Title);
            if (meta == null)
            {
                MessageBox.Show(
                    "The selected object is not a Mol* 3D structure object.",
                    "Mol* 3D for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Restore snapshot:
            // 1. Try local cache
            string? snapshotJson = MolstarSessionManager.LoadLocalCache(meta.Id);
            // 2. Try CustomXMLParts
            if (string.IsNullOrEmpty(snapshotJson))
            {
                try
                {
                    snapshotJson = MolstarSessionManager.LoadFromCustomXmlParts(app.ActivePresentation, meta.Id);
                }
                catch { }
            }

            var hwnd = GetHwnd(app);
            using var form = new MolstarForm("PowerPoint", meta.Pdb, snapshotJson);
            var result = hwnd != IntPtr.Zero ? form.ShowDialog(new Win32Window(hwnd)) : form.ShowDialog();

            if (result == DialogResult.OK && (form.ResultImageBytes != null || form.ResultGlbBytes != null))
            {
                ApplyMolstarReplace(app, targetShape, form, meta);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while editing Mol* 3D structure:\n" + ex.Message,
                "Mol* 3D for PowerPoint",
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

    private static void ApplyMolstarInsert(MSPowerPoint.Application app, MSPowerPoint.Slide slide, MolstarForm form)
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
            if (form.ResultEmbedInDoc)
            {
                MolstarSessionManager.SaveToCustomXmlParts(app.ActivePresentation, sessionId, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        float slideWidth = 960f, slideHeight = 540f;
        try
        {
            slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
        }
        catch { }

        MSPowerPoint.Shape? newShape = null;
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                float w = 320f, h = 320f;
                float left = Math.Max(50f, (slideWidth - w) / 2f);
                float top = Math.Max(50f, (slideHeight - h) / 2f);
                dynamic shapes = slide.Shapes;
                newShape = shapes.Add3DModel(tempGlb, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, w, h);
            }
            catch
            {
                // If 3D Model insertion fails or is unsupported, fallback to 2D image
                meta.Mode = "image";
            }
            finally
            {
                try { if (File.Exists(tempGlb)) File.Delete(tempGlb); } catch { }
            }
        }

        if (newShape == null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes!);
                var (width, height) = StructureMetadata.GetScaledDimensions(form.ResultImageBytes!, 320f, 240f, scale: 0.85f);
                float left = Math.Max(50f, (slideWidth - width) / 2f);
                float top = Math.Max(50f, (slideHeight - height) / 2f);

                newShape = slide.Shapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: MsoTriState.msoFalse,
                    SaveWithDocument: MsoTriState.msoTrue,
                    Left: left,
                    Top: top,
                    Width: width,
                    Height: height);
            }
            finally
            {
                try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
            }
        }

        ApplyMolstarShapeMetadata(newShape, meta);
        newShape.Select();
    }

    private static void ApplyMolstarReplace(MSPowerPoint.Application app, MSPowerPoint.Shape oldShape, MolstarForm form, MolstarMetadata oldMeta)
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
            if (form.ResultEmbedInDoc)
            {
                MolstarSessionManager.SaveToCustomXmlParts(app.ActivePresentation, sessionId, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        var slide = (MSPowerPoint.Slide)oldShape.Parent;
        float oldLeft = oldShape.Left;
        float oldTop = oldShape.Top;
        float oldWidth = oldShape.Width;
        float oldHeight = oldShape.Height;
        oldShape.Delete();

        MSPowerPoint.Shape? newShape = null;
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                dynamic shapes = slide.Shapes;
                newShape = shapes.Add3DModel(tempGlb, MsoTriState.msoFalse, MsoTriState.msoTrue, oldLeft, oldTop, oldWidth, oldHeight);
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

        if (newShape == null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes!);
                var (left, top, newWidth, newHeight) = StructureMetadata.FitReplacementBounds(
                    oldLeft, oldTop, oldWidth, oldHeight, form.ResultImageBytes!);

                newShape = slide.Shapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: MsoTriState.msoFalse,
                    SaveWithDocument: MsoTriState.msoTrue,
                    Left: left,
                    Top: top,
                    Width: newWidth,
                    Height: newHeight);
            }
            finally
            {
                try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
            }
        }

        ApplyMolstarShapeMetadata(newShape, meta);
        newShape.Select();
    }

    private static void ApplyMolstarShapeMetadata(MSPowerPoint.Shape shape, MolstarMetadata meta)
    {
        shape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
        shape.Title = $"Mol* 3D [{meta.Pdb}]";
        try { shape.Tags.Add("MOLSTAR", meta.Id); } catch { }
        try { shape.Tags.Add("PDB", meta.Pdb); } catch { }
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

            var raw = ExtractStructure(sel.ShapeRange[1]);
            var ket = MoleculeRenderer.TryConvertToKet(raw);
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
