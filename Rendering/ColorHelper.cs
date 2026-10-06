using System;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Smiles2Img.Rendering;

internal static class ColorHelper
{
    private static readonly Regex HexRegex = new(@"\A[0-9a-fA-F]{3}(?:[0-9a-fA-F]{3})?\z", RegexOptions.Compiled);

    /// <summary>
    /// Normalizes a color or background token (e.g. "bg=#FFFF00", "FFFF00", "trans", "white").
    /// Returns "transparent" for transparent canvas tokens.
    /// Returns standardized hex (#RRGGBB / #RGB) or standard color name.
    /// </summary>
    internal static string? Normalize(string? raw, bool isBackground = false)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;

        var token = raw!.Trim().ToLowerInvariant();

        if (token.StartsWith("bg=", StringComparison.Ordinal)) token = token.Substring(3).Trim();
        else if (token.StartsWith("ink=", StringComparison.Ordinal)) token = token.Substring(4).Trim();
        else if (token.StartsWith("line=", StringComparison.Ordinal)) token = token.Substring(5).Trim();
        else if (token.StartsWith("base=", StringComparison.Ordinal)) token = token.Substring(5).Trim();

        if (isBackground && token is "transparent" or "trans" or "none" or "clear" or "nobg")
            return "transparent";

        if (HexRegex.IsMatch(token))
            token = "#" + token;

        return token;
    }

    /// <summary>
    /// Safely attempts to parse a normalized color token into RGB fractions (0.0 .. 1.0).
    /// Returns false if transparent or invalid.
    /// </summary>
    internal static bool TryGetRgb(string? colorToken, out double r, out double g, out double b)
    {
        r = g = b = -1.0;
        if (string.IsNullOrWhiteSpace(colorToken)) return false;

        var token = Normalize(colorToken);
        if (token is null or "transparent") return false;

        if (token is "white" or "light")
        {
            r = g = b = 1.0;
            return true;
        }

        try
        {
            var color = ColorTranslator.FromHtml(token);
            if (color.A == 0) return false;

            r = color.R / 255.0;
            g = color.G / 255.0;
            b = color.B / 255.0;
            return true;
        }
        catch
        {
            return false;
        }
    }
}
