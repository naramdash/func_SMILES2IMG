using System;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Editor;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using MSPowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Smiles2Img.PowerPoint;

public static partial class PowerPointStructureManager
{
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
                    "func_SMILES2IMG",
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
                "func_SMILES2IMG",
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
                        "func_SMILES2IMG",
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
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string? snapshotJson = MolstarSessionManager.LoadLocalCache(meta.Id);
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
                "func_SMILES2IMG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static void ApplyMolstarInsert(MSPowerPoint.Application app, MSPowerPoint.Slide slide, MolstarForm form)
    {
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.ActivePresentation, form);

        float slideWidth = 960f, slideHeight = 540f;
        try
        {
            slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
        }
        catch { }

        MSPowerPoint.Shape? newShape = TryAddMolstar3DModel(
            slide,
            meta,
            form,
            Math.Max(50f, (slideWidth - 320f) / 2f),
            Math.Max(50f, (slideHeight - 320f) / 2f),
            320f,
            320f);

        if (newShape == null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.png");
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
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.ActivePresentation, form, oldMeta.Id);

        var slide = (MSPowerPoint.Slide)oldShape.Parent;
        float oldLeft = oldShape.Left;
        float oldTop = oldShape.Top;
        float oldWidth = oldShape.Width;
        float oldHeight = oldShape.Height;
        oldShape.Delete();

        MSPowerPoint.Shape? newShape = TryAddMolstar3DModel(
            slide,
            meta,
            form,
            oldLeft,
            oldTop,
            oldWidth,
            oldHeight);

        if (newShape == null)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.png");
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

    private static MSPowerPoint.Shape? TryAddMolstar3DModel(
        MSPowerPoint.Slide slide,
        MolstarMetadata meta,
        MolstarForm form,
        float left,
        float top,
        float width,
        float height)
    {
        if (form.ResultMode != "3dmodel" || form.ResultGlbBytes == null || form.ResultGlbBytes.Length == 0)
        {
            return null;
        }

        var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.glb");
        try
        {
            File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
            dynamic shapes = slide.Shapes;
            return shapes.Add3DModel(tempGlb, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
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

    private static void ApplyMolstarShapeMetadata(MSPowerPoint.Shape shape, MolstarMetadata meta)
    {
        shape.AlternativeText = MolstarSessionManager.FormatAltText(meta);
        shape.Title = $"Mol* 3D [{meta.Pdb}]";
        try { shape.Tags.Add("MOLSTAR", meta.Id); } catch { }
        try { shape.Tags.Add("PDB", meta.Pdb); } catch { }
    }
}
