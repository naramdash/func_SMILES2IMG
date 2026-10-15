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
        Description = "[Excel 2016+] Places a molecular structure picture over the cell (moves and sizes with the cell). Use on Excel 2016/2019/2021.", IsMacroType = true)]
    public static object ImgFloat(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, Ketcher (KET) string, or cell reference (e.g., A2, \"CCO\").")] string smiles,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"num\", \"stereo\", \"h\", \"all-c\", \"white\", combined (e.g. \"bw num\" or \"bw|stereo\").")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
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
        Description = "[Excel 2016+] Places a Ketcher (KET) structure/reaction picture over the cell. Use on Excel 2016/2019/2021.", IsMacroType = true)]
    public static object KetImgFloat(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference containing one (e.g., A2, \"{...}\").")] string ket,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"white\", combined.")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "KET2IMG.FLOAT";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        try
        {
            if (string.IsNullOrWhiteSpace(ket)) return ExcelError.ExcelErrorValue;
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
        Description = "[Excel 2016+] Places an InChI molecular structure picture over the cell. Use on Excel 2016/2019/2021.", IsMacroType = true)]
    public static object InchiImgFloat(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference (e.g., A2, \"InChI=1S/...\").")] string inchi,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"num\", \"stereo\", \"h\", \"all-c\", \"white\", combined.")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
        => ImgFloat(inchi, background, style, transform);

    [ExcelCommand(MenuName = "SMILES", MenuText = "Refit floating images")]
    public static void RefitFloatingImages()
    {
        try { FloatingPictures.RefitActiveSheet(); }
        catch (Exception error) { LogDisplay.WriteLine("SMILES2IMG.FLOAT refit: " + error.Message); }
    }
}
