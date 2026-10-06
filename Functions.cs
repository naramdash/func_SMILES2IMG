using System;
using System.Runtime.CompilerServices;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

[assembly: InternalsVisibleTo("Smiles2Img.Tests")]

namespace Smiles2Img;

public static class Functions
{
    [ExcelFunction(Name = "SMILES2IMG", Category = "SMILES",
        Description = "[Excel 2024 / Microsoft 365] Renders a high-resolution molecular structure image into the cell. On Excel 2016/2019/2021 use SMILES2IMG.FLOAT.", IsMacroType = true)]
    public static object Img(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string or a cell reference containing one (e.g., A2, \"CCO\").")] string smiles,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"num\", \"stereo\", \"h\", \"all-c\", \"white\", combined (e.g. \"bw num\" or \"bw|stereo\").")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2IMG";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        if (CellImages.NativeImagesUnsupported) return CellImages.UnsupportedMessage;

        try
        {
            var options = ParseOptions(background, style, transform);
            var input = MoleculeRenderer.Normalize(smiles);
            if (CellImages.Find(caller, input, options) is { } image) return image;

            var png = MoleculeRenderer.Render(input, options);
            CellImageUpdates.Queue(caller, input, options, png);
            return "";
        }
        catch (InvalidSmilesException ex)
        {
            LogDisplay.WriteLine("SMILES2IMG: " + ex.Message);
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("SMILES2IMG: " + error.Message);
            return ExcelError.ExcelErrorNA;
        }
    }

    /// <summary>
    /// Delegates parsing of Excel formula arguments to RenderOptions.
    /// </summary>
    internal static RenderOptions ParseOptions(object? bgArg, object? styleArg, object? transformArg)
        => RenderOptions.Parse(bgArg, styleArg, transformArg);

    [ExcelCommand(MenuName = "SMILES", MenuText = "Show diagnostics")]
    public static void ShowDiagnostics() => LogDisplay.Show();
}
