using System;
using System.IO;
using System.Text.RegularExpressions;
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
    private static readonly Regex PdbPattern = new(@"^[1-9][0-9A-Za-z]{3}$", RegexOptions.Compiled);

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
                    "Mol* 3D for Excel",
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
                    string text = Convert.ToString((object?)cell.Value2)?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(text))
                    {
                        if (PdbPattern.IsMatch(text))
                        {
                            initialPdb = text;
                        }
                        else if (text.StartsWith("PDB:", StringComparison.OrdinalIgnoreCase))
                        {
                            initialPdb = text.Substring(4).Trim();
                        }
                        else if (text.EndsWith(".pdb", StringComparison.OrdinalIgnoreCase) ||
                                 text.EndsWith(".cif", StringComparison.OrdinalIgnoreCase))
                        {
                            initialPdb = text;
                        }
                    }
                }
            }
            catch { }

            using var form = new MolstarForm("Excel Worksheet", initialPdb);
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
                "Mol* 3D for Excel",
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
                        "Mol* 3D for Excel",
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
                    "Mol* 3D for Excel",
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

            using var form = new MolstarForm("Excel Worksheet", meta.Pdb, snapshotJson);
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
                "Mol* 3D for Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private static void ApplyMolstarInsert(Excel.Application app, Excel.Worksheet sheet, MolstarForm form)
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
                MolstarSessionManager.SaveToCustomXmlParts(app.ActiveWorkbook, sessionId, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        // Target placement: near ActiveCell or (100, 100)
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

        float w = 320f, h = 320f;
        Excel.Shape? newShape = null;

        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                dynamic shapes = sheet.Shapes;
                newShape = shapes.Add3DModel(tempGlb, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, w, h);
            }
            catch
            {
                // Fallback to 2D image if 3D model is unsupported
                meta.Mode = "image";
            }
            finally
            {
                try { if (File.Exists(tempGlb)) File.Delete(tempGlb); } catch { }
            }
        }

        if (newShape == null && form.ResultImageBytes != null && form.ResultImageBytes.Length > 0)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{sessionId}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes);
                newShape = sheet.Shapes.AddPicture(
                    Filename: tempPng,
                    LinkToFile: MsoTriState.msoFalse,
                    SaveWithDocument: MsoTriState.msoTrue,
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
            ApplyShapeMetadata(newShape, meta);
            newShape.Select();
        }
    }

    private static void ApplyMolstarReplace(Excel.Application app, Excel.Shape oldShape, MolstarForm form, MolstarMetadata meta)
    {
        meta.Pdb = form.ResultPdbId;
        meta.Mode = form.ResultMode;
        meta.Embed = form.ResultEmbedInDoc;
        meta.Background = form.ResultBackground;

        if (!string.IsNullOrEmpty(form.ResultSnapshotJson))
        {
            MolstarSessionManager.SaveLocalCache(meta.Id, form.ResultSnapshotJson!);
            if (form.ResultEmbedInDoc)
            {
                MolstarSessionManager.SaveToCustomXmlParts(app.ActiveWorkbook, meta.Id, form.ResultPdbId, form.ResultSnapshotJson!);
            }
        }

        var sheet = (Excel.Worksheet)oldShape.Parent;
        float left = Convert.ToSingle(oldShape.Left);
        float top = Convert.ToSingle(oldShape.Top);
        float width = Convert.ToSingle(oldShape.Width);
        float height = Convert.ToSingle(oldShape.Height);
        var oldPlacement = oldShape.Placement;

        oldShape.Delete();

        Excel.Shape? newShape = null;

        if (form.ResultMode == "3dmodel" && form.ResultGlbBytes != null && form.ResultGlbBytes.Length > 0)
        {
            var tempGlb = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.glb");
            try
            {
                File.WriteAllBytes(tempGlb, form.ResultGlbBytes);
                dynamic shapes = sheet.Shapes;
                newShape = shapes.Add3DModel(tempGlb, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
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

        if (newShape == null && form.ResultImageBytes != null && form.ResultImageBytes.Length > 0)
        {
            var tempPng = Path.Combine(Path.GetTempPath(), $"molstar_{meta.Id}.png");
            try
            {
                File.WriteAllBytes(tempPng, form.ResultImageBytes);
                newShape = sheet.Shapes.AddPicture(
                    Filename: tempPng,
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

        if (newShape != null)
        {
            ApplyShapeMetadata(newShape, meta);
            newShape.Placement = oldPlacement;
            newShape.Select();
        }
    }

    private static void ApplyShapeMetadata(Excel.Shape shape, MolstarMetadata meta)
    {
        try { shape.AlternativeText = MolstarSessionManager.FormatAltText(meta); } catch { }
        try { shape.Title = StructureMetadata.FormatTitle(meta.Pdb, "Mol* 3D Structure"); } catch { }
        try { shape.Placement = Excel.XlPlacement.xlMoveAndSize; } catch { }
        try { shape.LockAspectRatio = MsoTriState.msoTrue; } catch { }
    }
}
