using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

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
    public string? HighlightQuery { get; set; }

    public string ToKeyString()
    {
        var parts = new List<string>(9)
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
        if (!string.IsNullOrWhiteSpace(HighlightQuery))
            parts.Add("HL-" + ComputeShortSha1Hex(HighlightQuery!.Trim()));

        return string.Join("_", parts);
    }

    private static string ComputeShortSha1Hex(string text)
    {
        using var sha = SHA1.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
        var sb = new StringBuilder(16);
        for (var i = 0; i < 8 && i < hash.Length; i++)
            sb.Append(hash[i].ToString("X2"));
        return sb.ToString();
    }

    private static readonly char[] PrimaryStyleSeparators = [' ', ',', '|', ';'];
    private static readonly char[] SecondaryStyleSeparators = ['+', '/'];
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
            if (raw is "false" or "bw" or "mono")
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
            foreach (var rawToken in ExtractStyleTokens(sStr))
            {
                if (TryExtractHighlightQuery(rawToken, out var hlQuery))
                {
                    options.HighlightQuery = hlQuery;
                    continue;
                }

                var token = rawToken.ToLowerInvariant();
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
                    case "rot90" or "rot180" or "rot270":
                        extraTransforms.Add(token);
                        break;
                    case "flip" or "flipx" or "flipy":
                        // Mirror reflection is disabled for stereochemical reliability; silently ignored for legacy workbook compatibility.
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
                if (t is "rot90" or "rot180" or "rot270")
                    transforms.Add(t);
            }
        }

        if (transforms.Count > 0)
            options.Transform = string.Join(",", transforms);

        return options;
    }

    private static readonly string[] HighlightPrefixes = ["highlight:", "highlight=", "hl:", "hl="];

    internal static List<string> ExtractStyleTokens(string? styleText)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(styleText)) return result;

        var primaryTokens = SplitRespectingBrackets(styleText!, PrimaryStyleSeparators);
        foreach (var primaryRaw in primaryTokens)
        {
            var trimmedPrimary = primaryRaw.Trim();
            if (TryExtractHighlightQuery(trimmedPrimary, out _))
            {
                result.Add(trimmedPrimary);
                continue;
            }

            var subTokens = SplitRespectingBrackets(trimmedPrimary, SecondaryStyleSeparators, stopAtHighlightPrefix: true);
            foreach (var rawToken in subTokens)
            {
                var trimmed = rawToken.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                    result.Add(trimmed);
            }
        }
        return result;
    }

    internal static bool IsColorOrInkToken(string token)
    {
        var lower = token.Trim().ToLowerInvariant();
        return lower is "bw" or "mono" or "monochrome" or "black" or "gray" or "grayscale" or "false"
            or "color" or "colour" or "cpk" or "true"
            or "white" or "light"
            || lower.StartsWith("ink=", StringComparison.Ordinal)
            || lower.StartsWith("line=", StringComparison.Ordinal)
            || lower.StartsWith("base=", StringComparison.Ordinal);
    }

    private static bool StartsWithHighlightPrefix(string token)
    {
        foreach (var prefix in HighlightPrefixes)
        {
            if (token.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static List<string> SplitRespectingBrackets(string input, char[] separators, bool stopAtHighlightPrefix = false)
    {
        var tokens = new List<string>();
        var start = 0;
        var depth = 0;
        for (var i = 0; i < input.Length; i++)
        {
            var ch = input[i];
            if (ch == '[') depth++;
            else if (ch == ']' && depth > 0) depth--;
            else if (depth == 0 && Array.IndexOf(separators, ch) >= 0)
            {
                var seg = input.Substring(start, i - start).Trim();
                if (!string.IsNullOrEmpty(seg))
                    tokens.Add(seg);
                start = i + 1;
                if (stopAtHighlightPrefix && StartsWithHighlightPrefix(input.Substring(start).TrimStart()))
                    break;
            }
        }
        var tail = input.Substring(start).Trim();
        if (!string.IsNullOrEmpty(tail))
            tokens.Add(tail);
        return tokens;
    }

    private static bool TryExtractHighlightQuery(string token, out string query)
    {
        query = "";
        foreach (var prefix in HighlightPrefixes)
        {
            if (token.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                var extracted = token.Substring(prefix.Length).Trim();
                if (!string.IsNullOrEmpty(extracted))
                {
                    query = extracted;
                    return true;
                }
            }
        }
        return false;
    }
}
