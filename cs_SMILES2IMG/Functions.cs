using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

[assembly: InternalsVisibleTo("Smiles2Img.Tests")]

namespace Smiles2Img;

public static class Functions
{
    [ExcelFunction(Name = "SMILES2IMG", Category = "SMILES",
        Description = "Renders a high-resolution molecular structure image from a SMILES string into the cell.", IsMacroType = true)]
    public static object Img(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string or a cell reference containing one (e.g., A2, \"CCO\").")] string smiles,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" or \"transparent\", \"white\" (default), or hex color (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "color", Description = "Optional: TRUE for color (default), FALSE for black & white mode.")] object? color = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2IMG";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;
        try
        {
            var (isColor, trans, bg) = ParseOptions(background, color, transform);
            var input = MoleculeRenderer.Normalize(smiles);
            if (CellImages.Find(caller, input, isColor, trans, bg) is { } image) return image;
            var png = MoleculeRenderer.Render(input, isColor, trans, bg);
            CellImageUpdates.Queue(caller, input, isColor, trans, bg, png);
            return "";
        }
        catch (InvalidSmilesException) { return ExcelError.ExcelErrorValue; }
        catch (Exception error)
        {
            LogDisplay.WriteLine("SMILES2IMG: " + error.Message);
            return ExcelError.ExcelErrorNA;
        }
    }

    internal static (bool Color, string? Transform, string? Background) ParseOptions(object? bgArg, object? colorArg, object? transformArg)
    {
        var isColor = true;
        string? background = null;

        if (bgArg is bool bgBool)
        {
            isColor = bgBool;
        }
        else if (bgArg is string bgStr && !string.IsNullOrWhiteSpace(bgStr))
        {
            var t = bgStr.Trim().ToLowerInvariant();
            if (t == "false" || t == "bw")
                isColor = false;
            else if (t is "transparent" or "trans" or "clear" or "nobg" or "none")
                background = "transparent";
            else if (t is "white" or "bg=white")
                background = "white";
            else if (t.StartsWith("bg="))
                background = bgStr.Substring(3).Trim();
            else
                background = bgStr.Trim();
        }

        if (colorArg is bool cBool)
        {
            isColor = cBool;
        }
        else if (colorArg is double d)
        {
            isColor = d != 0;
        }
        else if (colorArg is string cStr && !string.IsNullOrWhiteSpace(cStr))
        {
            if (cStr.Equals("false", StringComparison.OrdinalIgnoreCase) || cStr.Equals("bw", StringComparison.OrdinalIgnoreCase))
                isColor = false;
        }

        string? transform = null;
        if (transformArg is string transStr && !string.IsNullOrWhiteSpace(transStr))
        {
            var transforms = new List<string>();
            foreach (var part in transStr.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var t = part.Trim().ToLowerInvariant();
                if (t is "flip" or "flipx" or "flipy" or "rot90" or "rot180" or "rot270")
                    transforms.Add(t);
            }
            if (transforms.Count > 0)
                transform = string.Join(",", transforms);
        }

        return (isColor, transform, background);
    }

    [ExcelCommand(MenuName = "SMILES", MenuText = "Show diagnostics")]
    public static void ShowDiagnostics() => LogDisplay.Show();
}
