using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Editor;
using Smiles2Img.Notifications;
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
            using var form = MolstarForm.Rent("PowerPoint", "");
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
            using var form = MolstarForm.Rent("PowerPoint", meta.Pdb, snapshotJson);
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
                MoleculeNotification.NotifyWarning("SMILES", "Please select a molecular structure picture first.");
                return;
            }

            var smiles = ExtractSmiles(sel.ShapeRange[1]);
            if (string.IsNullOrEmpty(smiles))
            {
                MoleculeNotification.NotifyWarning("SMILES", "Could not find SMILES structure from the selected image.");
                return;
            }

            SafeClipboard.SetText(smiles!);
            MoleculeNotification.NotifyCopied("SMILES", smiles!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("SMILES", "Error extracting SMILES: " + ex.Message);
        }
    }

    public static void CopySelectedMoleculeInchi(MSPowerPoint.Application app)
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

            var inchi = MoleculeRenderer.TryConvertToInchi(targetData);
            if (string.IsNullOrEmpty(inchi))
            {
                MoleculeNotification.NotifyWarning("InChI", "Could not generate InChI from the selected structure or text.");
                return;
            }

            SafeClipboard.SetText(inchi!);
            MoleculeNotification.NotifyCopied("InChI", inchi!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChI", "Error generating InChI: " + ex.Message);
        }
    }

    public static void CopySelectedMoleculeKet(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            if (sel == null || sel.Type != MSPowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MoleculeNotification.NotifyWarning("Canvas Data (KET)", "Please select a molecular structure picture first.");
                return;
            }

            var raw = ExtractStructure(sel.ShapeRange[1]);
            var ket = MoleculeRenderer.TryConvertToKet(raw);
            if (string.IsNullOrEmpty(ket))
            {
                MoleculeNotification.NotifyWarning("Canvas Data (KET)", "Could not find canvas structure (KET) from the selected image.");
                return;
            }

            SafeClipboard.SetText(ket!);
            MoleculeNotification.NotifyCopied("Canvas Data (KET)", ket!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Canvas Data (KET)", "Error copying canvas data: " + ex.Message);
        }
    }

    public static void CopySelectedMoleculeSvg(MSPowerPoint.Application app)
    {
        try
        {
            var sel = app.ActiveWindow?.Selection;
            string? targetData = null;

            if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
            {
                var shape = sel.ShapeRange[1];
                targetData = ExtractStructure(shape) ?? ExtractSmiles(shape);
            }
            else if (sel != null && sel.Type == MSPowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
            {
                targetData = sel.TextRange.Text?.Trim();
            }

            if (string.IsNullOrEmpty(targetData))
            {
                MoleculeNotification.NotifyWarning("SVG Vector", "Please select a molecular structure picture or text first.");
                return;
            }

            var svg = MoleculeRenderer.RenderSvg(targetData, RenderOptions.Default);
            if (string.IsNullOrEmpty(svg))
            {
                MoleculeNotification.NotifyWarning("SVG Vector", "Could not generate SVG vector image from the selected structure.");
                return;
            }

            SvgClipboardHelper.CopySvgToClipboard(svg);
            MoleculeNotification.NotifyCopied("SVG Vector", "Vector graphic copied to clipboard");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("SVG Vector", "Error copying SVG: " + ex.Message);
        }
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

    public static void CopySelectedMoleculeInfo(MSPowerPoint.Application app) => ShowSelectedMoleculeInfo(app);

    public static MoleculeInfo? GetSelectedMoleculeInfo(MSPowerPoint.Application app, bool resolvePubChem = false)
    {
        var sel = app.ActiveWindow?.Selection;
        if (sel == null) return null;

        string? targetData = null;

        if (sel.Type == MSPowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
        {
            var shape = sel.ShapeRange[1];
            if (StructureMetadata.IsMolstar(shape.AlternativeText, shape.Title))
            {
                var meta = MolstarSessionManager.ParseAltText(shape.AlternativeText, shape.Title);
                if (meta != null && !string.IsNullOrWhiteSpace(meta.Pdb))
                {
                    (string Title, double? WeightKDa) pdbInfo = resolvePubChem ? PubChemResolver.ResolvePdb(meta.Pdb) : ("", (double?)null);
                    return new MoleculeInfo
                    {
                        CompoundName = !string.IsNullOrEmpty(pdbInfo.Title) ? pdbInfo.Title : $"PDB {meta.Pdb}",
                        IupacName = $"PDB Entry: {meta.Pdb}",
                        Formula = "Macromolecular Complex",
                        MolecularWeight = pdbInfo.WeightKDa.HasValue ? pdbInfo.WeightKDa.Value * 1000.0 : 0.0,
                        ExactMass = 0,
                        Smiles = $"PDB:{meta.Pdb}",
                        IsReaction = false
                    };
                }
            }
            targetData = ExtractSmiles(shape) ?? ExtractStructure(shape);
        }
        else if (sel.Type == MSPowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
        {
            targetData = sel.TextRange.Text?.Trim();
        }

        if (string.IsNullOrEmpty(targetData)) return null;
        return MoleculeRenderer.TryGetMoleculeInfo(targetData, resolvePubChem);
    }

    public static void CopySelectedSubstanceName(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null || string.IsNullOrWhiteSpace(info.CompoundName))
            {
                MoleculeNotification.NotifyWarning("Substance Name", "Could not resolve substance name from the selected structure.");
                return;
            }
            SafeClipboard.SetText(info.CompoundName);
            MoleculeNotification.NotifyCopied("Substance Name", info.CompoundName);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Substance Name", "Error copying substance name: " + ex.Message);
        }
    }

    public static void CopySelectedFormula(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: false);
            if (info == null || string.IsNullOrWhiteSpace(info.Formula))
            {
                MoleculeNotification.NotifyWarning("Formula", "Could not compute formula from the selected structure.");
                return;
            }
            SafeClipboard.SetText(info.Formula);
            MoleculeNotification.NotifyCopied("Chemical Formula", info.Formula);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Formula", "Error copying formula: " + ex.Message);
        }
    }

    public static void CopySelectedMolecularWeight(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: false);
            if (info == null || info.MolecularWeight <= 0)
            {
                MoleculeNotification.NotifyWarning("Molecular Weight", "Could not compute molecular weight from the selected structure.");
                return;
            }
            string mw = info.MolecularWeight.ToString("F2");
            SafeClipboard.SetText(mw);
            MoleculeNotification.NotifyCopied("Molecular Weight", $"{mw} g/mol");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecular Weight", "Error copying molecular weight: " + ex.Message);
        }
    }

    public static void CopySelectedInchiKey(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: false);
            if (info == null || string.IsNullOrWhiteSpace(info.InChIKey))
            {
                MoleculeNotification.NotifyWarning("InChIKey", "Could not generate InChIKey from the selected structure.");
                return;
            }
            SafeClipboard.SetText(info.InChIKey);
            MoleculeNotification.NotifyCopied("InChIKey", info.InChIKey);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChIKey", "Error copying InChIKey: " + ex.Message);
        }
    }

    public static void CopySelectedAllSummary(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from the selected structure.");
                return;
            }
            string summary = info.FormatClipboardSummary(MoleculeSummaryFormat.MultiLine);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.MultiLine);
            SafeClipboard.SetTextAndHtml(summary, html);
            MoleculeNotification.NotifyCopied("Molecule Info", info.SubstanceName ?? info.Formula ?? "Copied summary");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying molecule summary: " + ex.Message);
        }
    }

    public static void CopySelectedTsvRow(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from the selected structure.");
                return;
            }
            string tsv = info.FormatClipboardSummary(MoleculeSummaryFormat.ExcelTsvRow);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.ExcelTsvRow);
            SafeClipboard.SetTextAndHtml(tsv, html);
            MoleculeNotification.NotifyCopied("Excel Row (TSV)", info.SubstanceName ?? info.Formula ?? "Copied row");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying TSV row: " + ex.Message);
        }
    }

    public static void CopySelectedMultiLine(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from the selected structure.");
                return;
            }
            string multi = info.FormatClipboardSummary(MoleculeSummaryFormat.MultiLine);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.MultiLine);
            SafeClipboard.SetTextAndHtml(multi, html);
            MoleculeNotification.NotifyCopied("Multi-Line List", info.SubstanceName ?? info.Formula ?? "Copied list");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying multi-line list: " + ex.Message);
        }
    }

    public static void CopySelectedMarkdown(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", "Could not retrieve molecule info from the selected structure.");
                return;
            }
            string md = info.FormatClipboardSummary(MoleculeSummaryFormat.MarkdownTable);
            string? html = info.FormatHtmlForSummary(MoleculeSummaryFormat.MarkdownTable);
            SafeClipboard.SetTextAndHtml(md, html);
            MoleculeNotification.NotifyCopied("Markdown Table", info.SubstanceName ?? info.Formula ?? "Copied table");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", "Error copying markdown table: " + ex.Message);
        }
    }

    public static void ShowSelectedMoleculeInfo(MSPowerPoint.Application app)
    {
        try
        {
            var info = GetSelectedMoleculeInfo(app, resolvePubChem: true);
            if (info == null)
            {
                MessageBox.Show(
                    "Could not retrieve molecule properties from the selected structure.\nPlease select a molecular structure picture or text on the slide first.",
                    "SMILES for PowerPoint",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var hwnd = GetHwnd(app);
            if (hwnd != IntPtr.Zero)
                MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties", new Win32Window(hwnd));
            else
                MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties");
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
