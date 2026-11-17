using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Editor;
using MSWord = Microsoft.Office.Interop.Word;

namespace Smiles2Img.Word;

public static partial class WordStructureManager
{
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

            string initialPdb = "";
            try
            {
                GetSelectionTexts(app, out _, out _, out var selText);
                initialPdb = MolstarSessionManager.ExtractInitialPdbTarget(selText);
            }
            catch { }

            var hwnd = GetHwnd(app);
            using var form = MolstarForm.Rent("Word", initialPdb);
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
                "func_SMILES2IMG",
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
                    "func_SMILES2IMG",
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
            using var form = MolstarForm.Rent("Word", meta.Pdb, snapshotJson);
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
                "func_SMILES2IMG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static void ApplyMolstarInsert(MSWord.Application app, MolstarForm form)
    {
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.Documents.Count > 0 ? app.ActiveDocument : null, form);
        var (inlineShapes, range) = ResolveInsertTarget(app);
        dynamic? newShape = TryAddMolstarInline3DModel(
            () => inlineShapes,
            range,
            meta,
            form);

        if (newShape == null && form.ResultImageBytes != null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes);
                var shape = inlineShapes.AddPicture(
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

        FinalizeMolstarShape(newShape, meta, form);
    }

    private static void ApplyMolstarReplaceInline(MSWord.Application app, MSWord.InlineShape oldShape, MolstarForm form, MolstarMetadata oldMeta)
    {
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.Documents.Count > 0 ? app.ActiveDocument : null, form, oldMeta.Id);
        var oldWidth = oldShape.Width;
        var range = oldShape.Range;

        dynamic? newShape = TryAddMolstarInline3DModel(
            () => range.InlineShapes,
            range,
            meta,
            form);

        if (newShape == null && form.ResultImageBytes != null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.png");
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
            try { oldShape.Delete(); } catch { }
        }

        FinalizeMolstarShape(newShape, meta, form);
    }

    private static void ApplyMolstarReplaceShape(MSWord.Application app, MSWord.Shape oldShape, MolstarForm form, MolstarMetadata oldMeta)
    {
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.Documents.Count > 0 ? app.ActiveDocument : null, form, oldMeta.Id);
        var doc = (MSWord.Document)oldShape.Parent;
        float oldLeft = oldShape.Left;
        float oldTop = oldShape.Top;
        float oldWidth = oldShape.Width;
        float oldHeight = oldShape.Height;
        oldShape.Delete();

        dynamic? newShape = null;
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.glb");
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
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.png");
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

        FinalizeMolstarShape(newShape, meta, form);
    }

    private static dynamic? TryAddMolstarInline3DModel(Func<MSWord.InlineShapes> getInlineShapes, MSWord.Range anchorRange, MolstarMetadata meta, MolstarForm form)
    {
        if (form.ResultMode != "3dmodel" || form.ResultGlbBytes == null || form.ResultGlbBytes.Length == 0)
        {
            return null;
        }

        var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.glb");
        try
        {
            File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
            dynamic inlineShapes = getInlineShapes();
            return inlineShapes.Add3DModel(tempGlb, false, true, anchorRange);
        }
        catch
        {
            meta.Mode = "image";
            return null;
        }
        finally
        {
            try { if (File.Exists(tempGlb)) File.Delete(tempGlb); } catch { }
        }
    }

    private static void FinalizeMolstarShape(dynamic? newShape, MolstarMetadata meta, MolstarForm form)
    {
        if (newShape != null)
        {
            newShape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
            newShape.Title = $"Mol* 3D: {(string.IsNullOrEmpty(form.ResultPdbId) ? "Structure" : form.ResultPdbId)}";
            LockShapeAspectRatio(newShape);
            newShape.Select();
        }
    }
}
