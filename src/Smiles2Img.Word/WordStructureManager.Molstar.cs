using System;
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

        if (newShape is null && form.ResultImageBytes is { Length: > 0 } imageBytes)
        {
            newShape = StructureMetadata.WithTempFile(imageBytes, $"molstar_{meta.Id}", "png", tempPng =>
            {
                var shape = inlineShapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: false,
                    SaveWithDocument: true);

                var (w, h) = StructureMetadata.GetScaledDimensions(imageBytes, 320f, 240f, scale: 0.85f);
                shape.Width = w;
                shape.Height = h;
                return (dynamic)shape;
            });
        }

        FinalizeMolstarShape(newShape, meta);
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

        if (newShape is null && form.ResultImageBytes is { Length: > 0 } imageBytes)
        {
            newShape = StructureMetadata.WithTempFile(imageBytes, $"molstar_{meta.Id}", "png", tempPng =>
            {
                var shape = range.InlineShapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: false,
                    SaveWithDocument: true);

                var (_, h) = StructureMetadata.GetScaledDimensions(imageBytes, oldWidth, oldWidth, minWidth: oldWidth, maxWidth: oldWidth);
                shape.Width = oldWidth;
                shape.Height = h;
                return (dynamic)shape;
            });
        }

        if (newShape is not null)
        {
            try { oldShape.Delete(); } catch { }
        }

        FinalizeMolstarShape(newShape, meta);
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
        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes is { Length: > 0 } glbBytes)
        {
            try
            {
                newShape = StructureMetadata.WithTempFile<dynamic>(glbBytes, $"molstar_{meta.Id}", "glb", tempGlb =>
                {
                    dynamic shapes = doc.Shapes;
                    return shapes.Add3DModel(tempGlb, false, true, oldLeft, oldTop, oldWidth, oldHeight);
                });
            }
            catch
            {
                meta.Mode = "image";
            }
        }

        if (newShape is null && form.ResultImageBytes is { Length: > 0 } imageBytes)
        {
            newShape = StructureMetadata.WithTempFile(imageBytes, $"molstar_{meta.Id}", "png", tempPng =>
            {
                var (left, top, w, h) = StructureMetadata.FitReplacementBounds(oldLeft, oldTop, oldWidth, oldHeight, imageBytes);
                return (dynamic)doc.Shapes.AddPicture(
                    FileName: tempPng,
                    LinkToFile: false,
                    SaveWithDocument: true,
                    Left: left,
                    Top: top,
                    Width: w,
                    Height: h);
            });
        }

        FinalizeMolstarShape(newShape, meta);
    }

    private static dynamic? TryAddMolstarInline3DModel(Func<MSWord.InlineShapes> getInlineShapes, MSWord.Range anchorRange, MolstarMetadata meta, MolstarForm form)
    {
        if (form.ResultMode != "3dmodel" || form.ResultGlbBytes == null || form.ResultGlbBytes.Length == 0)
        {
            return null;
        }

        try
        {
            return StructureMetadata.WithTempFile<dynamic>(form.ResultGlbBytes, $"molstar_{meta.Id}", "glb", tempGlb =>
            {
                dynamic inlineShapes = getInlineShapes();
                return inlineShapes.Add3DModel(tempGlb, false, true, anchorRange);
            });
        }
        catch
        {
            meta.Mode = "image";
            return null;
        }
    }

    private static void FinalizeMolstarShape(dynamic? newShape, MolstarMetadata meta)
    {
        if (newShape != null)
        {
            newShape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
            newShape.Title = MolstarSessionManager.FormatTitle(meta.Pdb);
            LockShapeAspectRatio(newShape);
            newShape.Select();
        }
    }
}
