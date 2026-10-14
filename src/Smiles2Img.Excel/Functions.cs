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
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, Ketcher (KET) string, or cell reference (e.g., A2, \"CCO\").")] string smiles,
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

            if (MoleculeRenderer.IsKet(smiles))
            {
                var cleanKet = MoleculeRenderer.CleanKet(smiles);
                if (CellImages.Find(caller, cleanKet, options) is { } cachedKet) return cachedKet;

                var pngKet = MoleculeRenderer.RenderKet(cleanKet, options);
                if (pngKet == null) return ExcelError.ExcelErrorValue;

                CellImageUpdates.Queue(caller, cleanKet, options, pngKet);
                return "";
            }

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

    [ExcelFunction(Name = "KET2IMG", Category = "SMILES",
        Description = "[Excel 2024 / Microsoft 365] Renders a Ketcher (KET) structure/reaction image into the cell. On Excel 2016/2019/2021 use KET2IMG.FLOAT.", IsMacroType = true)]
    public static object KetImg(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference containing one (e.g., A2, \"{...}\").")] string ket,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"white\", combined (e.g. \"bw\").")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "KET2IMG";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        if (CellImages.NativeImagesUnsupported) return CellImages.UnsupportedMessage;

        try
        {
            if (string.IsNullOrWhiteSpace(ket)) return ExcelError.ExcelErrorValue;
            var options = ParseOptions(background, style, transform);
            var cleanKet = MoleculeRenderer.CleanKet(ket);
            if (CellImages.Find(caller, cleanKet, options) is { } image) return image;

            var png = MoleculeRenderer.RenderKet(cleanKet, options);
            if (png == null) return ExcelError.ExcelErrorValue;

            CellImageUpdates.Queue(caller, cleanKet, options, png);
            return "";
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("KET2IMG: " + error.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2NAME", Category = "SMILES",
        Description = "Resolves chemical substance name from PubChem for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesName(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2NAME";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles);
            if (info != null && !string.IsNullOrWhiteSpace(info.CompoundName))
                return info.CompoundName;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2MW", Category = "SMILES",
        Description = "Calculates molecular weight (MW) in g/mol for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesMW(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2MW";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles);
            if (info != null && !info.IsReaction && info.MolecularWeight > 0)
                return Math.Round(info.MolecularWeight, 2);
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2FORMULA", Category = "SMILES",
        Description = "Calculates Hill chemical formula for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesFormula(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2FORMULA";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles);
            if (info != null && !info.IsReaction && !string.IsNullOrWhiteSpace(info.Formula))
                return info.Formula;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "KET2NAME", Category = "SMILES",
        Description = "Resolves chemical substance name from PubChem for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetName(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesName(ket);

    [ExcelFunction(Name = "KET2MW", Category = "SMILES",
        Description = "Calculates molecular weight (MW) in g/mol for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetMW(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesMW(ket);

    [ExcelFunction(Name = "KET2FORMULA", Category = "SMILES",
        Description = "Calculates Hill chemical formula for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetFormula(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesFormula(ket);

    /// <summary>
    /// Delegates parsing of Excel formula arguments to RenderOptions.
    /// </summary>
    internal static RenderOptions ParseOptions(object? bgArg, object? styleArg, object? transformArg)
        => RenderOptions.Parse(bgArg, styleArg, transformArg);

    [ExcelCommand(MenuName = "SMILES", MenuText = "Show diagnostics")]
    public static void ShowDiagnostics() => LogDisplay.Show();
}
