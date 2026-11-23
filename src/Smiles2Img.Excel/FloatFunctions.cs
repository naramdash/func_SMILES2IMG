using System;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

namespace Smiles2Img;

// Excel 2016/2019/2021 variant of SMILES2IMG: same arguments, but the molecule is a
// picture floating over the cell instead of a native in-cell image.
public static class FloatFunctions
{
    [ExcelFunction(Name = "SMILES2IMG.FLOAT", Category = "SMILES",
        Description = "[Excel 2016-2021] Cell-anchored floating 2D image (offline Indigo). Accepts SMILES, >>, InChI, or KET; #VALUE! on bad syntax.", IsMacroType = true)]
    public static object ImgFloat(
        [ExcelArgument(Name = "smiles", Description = Functions.ArgSmilesImgDesc)] string smiles,
        [ExcelArgument(Name = "background", Description = Functions.ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = Functions.ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = Functions.ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2IMG.FLOAT";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        try
        {
            var options = Functions.ParseOptions(background, style, transform);

            if (MoleculeRenderer.IsKet(smiles))
            {
                var cleanKet = MoleculeRenderer.CleanKet(smiles);
                var key = CellImages.Key(cleanKet, options);
                var png = FloatingPictures.Render(key, () => MoleculeRenderer.RenderKet(cleanKet, options) ?? throw new InvalidOperationException("Invalid KET data"));
                FloatingPictureUpdates.Place(caller, key, cleanKet, png);
                return "";
            }

            var input = MoleculeRenderer.Normalize(smiles);
            var keySmiles = CellImages.Key(input, options);
            var pngSmiles = FloatingPictures.Render(keySmiles, () => MoleculeRenderer.Render(input, options));
            FloatingPictureUpdates.Place(caller, keySmiles, input, pngSmiles);
            return "";
        }
        catch (InvalidSmilesException ex)
        {
            LogDisplay.WriteLine("SMILES2IMG.FLOAT: " + ex.Message);
            FloatingPictureUpdates.Remove(caller);
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("SMILES2IMG.FLOAT: " + error.Message);
            FloatingPictureUpdates.Remove(caller);
            return ExcelError.ExcelErrorNA;
        }
    }

    [ExcelFunction(Name = "KET2IMG.FLOAT", Category = "SMILES",
        Description = "[Excel 2016-2021] Cell-anchored floating 2D image from strict Ketcher KET JSON (molecule or reaction, offline); #VALUE! if not KET.", IsMacroType = true)]
    public static object KetImgFloat(
        [ExcelArgument(Name = "ket", Description = Functions.ArgKetImgDesc)] string ket,
        [ExcelArgument(Name = "background", Description = Functions.ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = Functions.ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = Functions.ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "KET2IMG.FLOAT";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        try
        {
            if (string.IsNullOrWhiteSpace(ket) || !MoleculeRenderer.IsKet(ket)) { FloatingPictureUpdates.Remove(caller); return ExcelError.ExcelErrorValue; }
            var options = Functions.ParseOptions(background, style, transform);
            var cleanKet = MoleculeRenderer.CleanKet(ket);
            var key = CellImages.Key(cleanKet, options);
            var png = FloatingPictures.Render(key, () => MoleculeRenderer.RenderKet(cleanKet, options) ?? throw new InvalidOperationException("Invalid KET data"));
            FloatingPictureUpdates.Place(caller, key, cleanKet, png);
            return "";
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("KET2IMG.FLOAT: " + error.Message);
            FloatingPictureUpdates.Remove(caller);
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "INCHI2IMG.FLOAT", Category = "SMILES",
        Description = "[Excel 2016-2021] Cell-anchored floating 2D image from strict IUPAC InChI (offline Indigo); #VALUE! if input is not InChI.", IsMacroType = true)]
    public static object InchiImgFloat(
        [ExcelArgument(Name = "inchi", Description = Functions.ArgInchiImgDesc)] string inchi,
        [ExcelArgument(Name = "background", Description = Functions.ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = Functions.ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = Functions.ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "INCHI2IMG.FLOAT";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;
        if (string.IsNullOrWhiteSpace(inchi) || !MoleculeRenderer.IsInchi(inchi))
        {
            FloatingPictureUpdates.Remove(caller);
            return ExcelError.ExcelErrorValue;
        }
        return ImgFloat(inchi, background, style, transform);
    }

    [ExcelFunction(Name = "NAME2IMG.FLOAT", Category = "SMILES",
        Description = "[Excel 2016-2021] Cell-anchored floating 2D image from name/CAS. 104 built-in offline + CAS check; else PubChem (30d cache, #N/A if offline).", IsMacroType = true)]
    public static object NameImgFloat(
        [ExcelArgument(Name = "name", Description = Functions.ArgNameDesc)] string name,
        [ExcelArgument(Name = "background", Description = Functions.ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = Functions.ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = Functions.ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "NAME2IMG.FLOAT";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;
        if (string.IsNullOrWhiteSpace(name))
        {
            FloatingPictureUpdates.Remove(caller);
            return ExcelError.ExcelErrorValue;
        }
        var smiles = PubChemResolver.ResolveNameToSmiles(name);
        if (string.IsNullOrWhiteSpace(smiles))
        {
            FloatingPictureUpdates.Remove(caller);
            return ExcelError.ExcelErrorNA;
        }
        return ImgFloat(smiles!, background, style, transform);
    }

    [ExcelCommand(MenuName = "SMILES", MenuText = "Refit floating images")]
    public static void RefitFloatingImages()
    {
        try { FloatingPictures.RefitActiveSheet(); }
        catch (Exception error) { LogDisplay.WriteLine("SMILES2IMG.FLOAT refit: " + error.Message); }
    }
}
