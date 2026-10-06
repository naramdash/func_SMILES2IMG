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
        [ExcelArgument(Name = "smiles", Description = "A SMILES string or a cell reference containing one (e.g., A2, \"CCO\").")] string smiles,
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
            var input = MoleculeRenderer.Normalize(smiles);
            var key = CellImages.Key(input, options);
            var png = FloatingPictures.Render(key, () => MoleculeRenderer.Render(input, options));
            FloatingPictureUpdates.Place(caller, key, input, png);
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

    [ExcelCommand(MenuName = "SMILES", MenuText = "Refit floating images")]
    public static void RefitFloatingImages()
    {
        try { FloatingPictures.RefitActiveSheet(); }
        catch (Exception error) { LogDisplay.WriteLine("SMILES2IMG.FLOAT refit: " + error.Message); }
    }
}
