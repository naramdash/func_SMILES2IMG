using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img;

/// <summary>
/// Manages embedded Mol* 3D macromolecular viewer interactions, floating 3D model (.glb)
/// and 2D picture (.png) insertion, in-place re-editing, and session persistence in Microsoft Excel.
/// </summary>
public static class ExcelMolstarManager
{
    public static void InsertMolstarStructure(Excel.Application app)
    {
        try
        {
            var shape = KetcherEditor.GetSelectedShape(app);
            if (shape != null && StructureMetadata.IsMolstar(shape.AlternativeText, shape.Title))
            {
                EditSelectedMolstar(app, shape);
                return;
            }

            if (app.ActiveSheet is not Excel.Worksheet sheet)
            {
                MessageBox.Show(
                    "No active worksheet found.\nPlease open a worksheet to insert the 3D structure first.",
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Check if active cell contains a PDB ID or structure path
            string initialPdb = "";
            try
            {
                if (app.ActiveCell is Excel.Range cell)
                {
                    initialPdb = MolstarSessionManager.ExtractInitialPdbTarget(Convert.ToString((object?)cell.Value2));
                }
            }
            catch { }

            using var form = MolstarForm.Rent("Excel Worksheet", initialPdb);
            var result = form.ShowDialog(new Win32Window(ExcelDnaUtil.WindowHandle));

            if (result == DialogResult.OK && (form.ResultImageBytes != null || form.ResultGlbBytes != null))
            {
                ApplyMolstarInsert(app, sheet, form);
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

    public static void EditSelectedMolstar(Excel.Application app, Excel.Shape? targetShape = null)
    {
        try
        {
            if (targetShape == null)
            {
                targetShape = KetcherEditor.GetSelectedShape(app);
                if (targetShape == null)
                {
                    MessageBox.Show(
                        "No 3D molecular structure selected.\nPlease select a Mol* structure shape or 3D model first.",
                        "func_SMILES2IMG",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }
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

            // Restore snapshot:
            // 1. Try local cache
            string? snapshotJson = MolstarSessionManager.LoadLocalCache(meta.Id);
            // 2. Try CustomXMLParts
            if (string.IsNullOrEmpty(snapshotJson))
            {
                try
                {
                    snapshotJson = MolstarSessionManager.LoadFromCustomXmlParts(app.ActiveWorkbook, meta.Id);
                }
                catch { }
            }

            using var form = MolstarForm.Rent("Excel Worksheet", meta.Pdb, snapshotJson);
            var result = form.ShowDialog(new Win32Window(ExcelDnaUtil.WindowHandle));

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

    private static void ApplyMolstarInsert(Excel.Application app, Excel.Worksheet sheet, MolstarForm form)
    {
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.ActiveWorkbook, form);

        float left = 100f, top = 100f;
        try
        {
            if (app.ActiveCell is Excel.Range cell)
            {
                left = Convert.ToSingle(cell.Left);
                top = Convert.ToSingle(cell.Top);
            }
        }
        catch { }

        Excel.Shape? newShape = TryAddMolstar3DModel(sheet, meta, form, left, top, 320f, 320f);
        if (newShape == null && form.ResultImageBytes != null && form.ResultImageBytes.Length > 0)
        {
            var (imgW, imgH) = StructureMetadata.GetScaledDimensions(form.ResultImageBytes, 320f, 240f, scale: 0.85f);
            newShape = AddMolstarPicture(sheet, meta, form.ResultImageBytes, left, top, imgW, imgH);
        }

        if (newShape != null)
        {
            ApplyShapeMetadata(newShape, meta);
            newShape.Select();
        }
    }

    private static void ApplyMolstarReplace(Excel.Application app, Excel.Shape oldShape, MolstarForm form, MolstarMetadata oldMeta)
    {
        var meta = MolstarSessionManager.BuildAndPersistMetadata(app.ActiveWorkbook, form, oldMeta.Id);

        var sheet = (Excel.Worksheet)oldShape.Parent;
        float left = Convert.ToSingle(oldShape.Left);
        float top = Convert.ToSingle(oldShape.Top);
        float width = Convert.ToSingle(oldShape.Width);
        float height = Convert.ToSingle(oldShape.Height);
        var oldPlacement = oldShape.Placement;

        oldShape.Delete();

        Excel.Shape? newShape = TryAddMolstar3DModel(sheet, meta, form, left, top, width, height);
        if (newShape == null && form.ResultImageBytes != null && form.ResultImageBytes.Length > 0)
        {
            (left, top, width, height) = StructureMetadata.FitReplacementBounds(left, top, width, height, form.ResultImageBytes);
            newShape = AddMolstarPicture(sheet, meta, form.ResultImageBytes, left, top, width, height);
        }

        if (newShape != null)
        {
            ApplyShapeMetadata(newShape, meta);
            newShape.Placement = oldPlacement;
            newShape.Select();
        }
    }

    private static Excel.Shape? TryAddMolstar3DModel(
        Excel.Worksheet sheet, MolstarMetadata meta, MolstarForm form, float left, float top, float width, float height)
    {
        if (form.ResultMode != "3dmodel" || form.ResultGlbBytes == null || form.ResultGlbBytes.Length == 0)
            return null;

        try
        {
            return StructureMetadata.WithTempFile(form.ResultGlbBytes, $"molstar_{meta.Id}", "glb", tempGlb =>
            {
                dynamic shapes = sheet.Shapes;
                return (Excel.Shape)shapes.Add3DModel(tempGlb, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
            });
        }
        catch
        {
            // Fallback to 2D image if 3D model is unsupported
            meta.Mode = "image";
            return null;
        }
    }

    private static Excel.Shape AddMolstarPicture(
        Excel.Worksheet sheet, MolstarMetadata meta, byte[] pngBytes, float left, float top, float width, float height) =>
        StructureMetadata.WithTempFile(pngBytes, $"molstar_{meta.Id}", "png", tempPng =>
            sheet.Shapes.AddPicture(
                Filename: tempPng,
                LinkToFile: MsoTriState.msoFalse,
                SaveWithDocument: MsoTriState.msoTrue,
                Left: left,
                Top: top,
                Width: width,
                Height: height));

    private static void ApplyShapeMetadata(Excel.Shape shape, MolstarMetadata meta)
    {
        try { shape.AlternativeText = MolstarSessionManager.FormatAltText(meta); } catch { }
        try { shape.Title = MolstarSessionManager.FormatTitle(meta.Pdb); } catch { }
        try { shape.Placement = Excel.XlPlacement.xlMoveAndSize; } catch { }
        try { shape.LockAspectRatio = MsoTriState.msoTrue; } catch { }
    }
}
