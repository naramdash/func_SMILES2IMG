using System;
using System.Collections.Generic;

namespace Smiles2Img.Rendering;

public sealed class RenderOptions
{
    public static RenderOptions Default => new();
    public bool Color { get; set; } = true;
    public string? BaseColor { get; set; }
    public bool AtomNumbers { get; set; }
    public bool StereoLabels { get; set; }
    public bool UnfoldHydrogens { get; set; }
    public bool AllCarbons { get; set; }
    public string? Background { get; set; } = "transparent";
    public string? Transform { get; set; }

    public string ToKeyString()
    {
        var parts = new List<string>(8)
        {
            Color ? "C" : "BW"
        };

        if (!string.IsNullOrWhiteSpace(BaseColor))
            parts.Add("INK-" + BaseColor!.Trim().ToLowerInvariant());
        if (AtomNumbers)
            parts.Add("NUM");
        if (StereoLabels)
            parts.Add("STR");
        if (UnfoldHydrogens)
            parts.Add("H");
        if (AllCarbons)
            parts.Add("CARB");
        if (!string.IsNullOrWhiteSpace(Background) && Background != "transparent")
            parts.Add("BG-" + Background!.Trim().ToLowerInvariant());
        if (!string.IsNullOrWhiteSpace(Transform))
            parts.Add("TR-" + Transform!.Trim().ToLowerInvariant());

        return string.Join("_", parts);
    }

    private static readonly char[] StyleSeparators = [' ', ',', '|', ';', '+', '/'];
    private static readonly char[] TransformSeparators = [' ', ',', '|', ';'];

    /// <summary>
    /// Parses Excel formula arguments into validated RenderOptions.
    /// Supports legacy booleans/numbers and modern string tokens.
    /// </summary>
    public static RenderOptions Parse(object? bgArg, object? styleArg, object? transformArg)
    {
        var options = new RenderOptions();

        // 1. Background (2nd argument)
        if (bgArg is bool bgBool)
        {
            options.Color = bgBool;
            options.Background = "transparent";
        }
        else if (bgArg is string bgStr && !string.IsNullOrWhiteSpace(bgStr))
        {
            var raw = bgStr.Trim().ToLowerInvariant();
            if (raw is "false" or "bw" or "black" or "mono")
            {
                options.Color = false;
                options.Background = "transparent";
            }
            else
            {
                options.Background = ColorHelper.Normalize(raw, isBackground: true);
            }
        }

        // 2. Style / Domain options (3rd argument)
        var extraTransforms = new List<string>();
        if (styleArg is bool sBool)
        {
            options.Color = sBool;
        }
        else if (styleArg is double d)
        {
            options.Color = d != 0;
        }
        else if (styleArg is string sStr && !string.IsNullOrWhiteSpace(sStr))
        {
            var tokens = sStr.Split(StyleSeparators, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawToken in tokens)
            {
                var token = rawToken.Trim().ToLowerInvariant();
                switch (token)
                {
                    case "bw" or "mono" or "monochrome" or "black" or "gray" or "grayscale" or "false":
                        options.Color = false;
                        break;
                    case "color" or "colour" or "cpk" or "true":
                        options.Color = true;
                        break;
                    case "white" or "light":
                        options.BaseColor = "white";
                        break;
                    case "num" or "number" or "numbers" or "idx" or "index" or "atom-num" or "atom-nums" or "atom-ids" or "atom-number":
                        options.AtomNumbers = true;
                        break;
                    case "stereo" or "chiral" or "abs" or "ext" or "stereochemistry":
                        options.StereoLabels = true;
                        break;
                    case "h" or "hydrogen" or "hydrogens" or "unfoldh" or "all-h":
                        options.UnfoldHydrogens = true;
                        break;
                    case "carbon" or "carbons" or "all-c" or "c":
                        options.AllCarbons = true;
                        break;
                    case "flip" or "flipx" or "flipy" or "rot90" or "rot180" or "rot270":
                        extraTransforms.Add(token);
                        break;
                    default:
                        if (token.StartsWith("ink=", StringComparison.Ordinal) ||
                            token.StartsWith("line=", StringComparison.Ordinal) ||
                            token.StartsWith("base=", StringComparison.Ordinal))
                        {
                            options.BaseColor = ColorHelper.Normalize(token);
                        }
                        break;
                }
            }
        }

        // 3. Transform (4th argument)
        var transforms = new List<string>(extraTransforms);
        if (transformArg is string transStr && !string.IsNullOrWhiteSpace(transStr))
        {
            foreach (var part in transStr.Split(TransformSeparators, StringSplitOptions.RemoveEmptyEntries))
            {
                var t = part.Trim().ToLowerInvariant();
                if (t is "flip" or "flipx" or "flipy" or "rot90" or "rot180" or "rot270")
                    transforms.Add(t);
            }
        }

        if (transforms.Count > 0)
            options.Transform = string.Join(",", transforms);

        return options;
    }
}
