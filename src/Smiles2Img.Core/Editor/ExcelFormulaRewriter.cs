using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

public enum FormulaRewriteStatus
{
    NotApplicable = 0,
    Updated = 1,
    SkippedComposite = 2
}

/// <summary>
/// Lossless Excel formula argument parser and style rewriter for SMILES/KET/INCHI/NAME image UDFs.
/// Preserves argument 1 (including nested calls/references), non-color style tokens (num, stereo, h, all-c, hl:*),
/// dynamic style expressions, and argument 4+ (rot90/rot180/rot270).
/// </summary>
public static class ExcelFormulaRewriter
{
    private static readonly HashSet<string> ImageFunctionNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "SMILES2IMG",
        "KET2IMG",
        "INCHI2IMG",
        "NAME2IMG",
        "SMILES2IMG.FLOAT",
        "KET2IMG.FLOAT",
        "INCHI2IMG.FLOAT",
        "NAME2IMG.FLOAT"
    };

    private static readonly Regex CompositeImageCallRegex = new(
        @"(?<![A-Za-z0-9_.])(?:_xll\.)?(?:SMILES2IMG|KET2IMG|INCHI2IMG|NAME2IMG)(?:\.FLOAT)?\s*\(",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex BatchPrefixExprRegex = new(
        @"^""(?:color|bw|white)\s+""\s*&\s*\((?<inner>[\s\S]*)\)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static FormulaRewriteStatus TryRewriteImageStyle(
        string? formula,
        string background,
        string style,
        out string rewrittenFormula)
    {
        rewrittenFormula = formula ?? "";
        if (string.IsNullOrWhiteSpace(formula))
            return FormulaRewriteStatus.NotApplicable;

        var trimmed = formula!.Trim();
        if (!trimmed.StartsWith("=", StringComparison.Ordinal))
            return FormulaRewriteStatus.NotApplicable;

        if (!TryParseTopLevelCall(trimmed, out var prefix, out var funcName, out var openParenIdx))
            return ContainsImageCallOutsideStrings(trimmed)
                ? FormulaRewriteStatus.SkippedComposite
                : FormulaRewriteStatus.NotApplicable;

        if (!ImageFunctionNames.Contains(funcName))
            return ContainsImageCallOutsideStrings(trimmed)
                ? FormulaRewriteStatus.SkippedComposite
                : FormulaRewriteStatus.NotApplicable;

        if (!TrySplitTopLevelArguments(trimmed, openParenIdx, out var args, out var closeParenIdx))
            return FormulaRewriteStatus.SkippedComposite;

        // Require the closing parenthesis to end the formula expression cleanly (ignoring trailing whitespace).
        for (int i = closeParenIdx + 1; i < trimmed.Length; i++)
        {
            if (!char.IsWhiteSpace(trimmed[i]))
                return FormulaRewriteStatus.SkippedComposite;
        }

        if (args.Count < 1 || string.IsNullOrWhiteSpace(args[0]))
            return FormulaRewriteStatus.SkippedComposite;

        var arg1 = args[0].Trim();
        var arg2 = QuoteExcelString(background ?? "trans");
        var arg3Raw = args.Count >= 3 ? args[2].Trim() : "";
        var arg3 = RewriteStyleArgument(arg3Raw, style ?? "color");

        var rebuilt = new StringBuilder();
        rebuilt.Append('=').Append(prefix).Append(funcName).Append('(');
        rebuilt.Append(arg1);
        rebuilt.Append(", ").Append(arg2);
        rebuilt.Append(", ").Append(arg3);

        // Find last non-whitespace trailing argument at index >= 3 so empty placeholders aren't left dangling,
        // while preserving any transform (or extra positional argument) verbatim.
        int lastNonEmptyExtra = -1;
        for (int i = 3; i < args.Count; i++)
        {
            if (!string.IsNullOrWhiteSpace(args[i]))
                lastNonEmptyExtra = i;
        }

        for (int i = 3; i <= lastNonEmptyExtra; i++)
        {
            rebuilt.Append(", ");
            var extra = args[i].Trim();
            if (!string.IsNullOrEmpty(extra))
                rebuilt.Append(extra);
        }

        rebuilt.Append(')');
        rewrittenFormula = rebuilt.ToString();
        return FormulaRewriteStatus.Updated;
    }

    private static bool TryParseTopLevelCall(
        string trimmedFormula,
        out string prefix,
        out string funcName,
        out int openParenIdx)
    {
        prefix = "";
        funcName = "";
        openParenIdx = -1;

        int i = 1; // skip '='
        while (i < trimmedFormula.Length && char.IsWhiteSpace(trimmedFormula[i])) i++;

        var prefixSb = new StringBuilder();
        if (i < trimmedFormula.Length && trimmedFormula[i] == '@')
        {
            prefixSb.Append('@');
            i++;
            while (i < trimmedFormula.Length && char.IsWhiteSpace(trimmedFormula[i])) i++;
        }

        if (i + 5 <= trimmedFormula.Length &&
            trimmedFormula.Substring(i, 5).Equals("_xll.", StringComparison.OrdinalIgnoreCase))
        {
            prefixSb.Append(trimmedFormula, i, 5);
            i += 5;
        }

        int nameStart = i;
        while (i < trimmedFormula.Length && IsFunctionNameChar(trimmedFormula[i])) i++;
        if (i == nameStart)
            return false;

        var name = trimmedFormula.Substring(nameStart, i - nameStart);
        while (i < trimmedFormula.Length && char.IsWhiteSpace(trimmedFormula[i])) i++;
        if (i >= trimmedFormula.Length || trimmedFormula[i] != '(')
            return false;

        prefix = prefixSb.ToString();
        funcName = name;
        openParenIdx = i;
        return true;
    }

    private static bool IsFunctionNameChar(char ch) =>
        char.IsLetterOrDigit(ch) || ch == '_' || ch == '.';

    private static bool TrySplitTopLevelArguments(
        string text,
        int openParenIdx,
        out List<string> args,
        out int closeParenIdx)
    {
        args = new List<string>();
        closeParenIdx = -1;

        int argStart = openParenIdx + 1;
        int parenDepth = 0;
        int braceDepth = 0;
        bool inString = false;
        bool inSingleQuote = false;

        for (int i = openParenIdx + 1; i < text.Length; i++)
        {
            char ch = text[i];

            if (inString)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++; // skip escaped quote
                        continue;
                    }
                    inString = false;
                }
                continue;
            }

            if (inSingleQuote)
            {
                if (ch == '\'')
                {
                    if (i + 1 < text.Length && text[i + 1] == '\'')
                    {
                        i++; // skip escaped apostrophe in sheet name
                        continue;
                    }
                    inSingleQuote = false;
                }
                continue;
            }

            if (ch == '"')
            {
                inString = true;
            }
            else if (ch == '\'')
            {
                inSingleQuote = true;
            }
            else if (ch == '{')
            {
                braceDepth++;
            }
            else if (ch == '}' && braceDepth > 0)
            {
                braceDepth--;
            }
            else if (ch == '(')
            {
                parenDepth++;
            }
            else if (ch == ')')
            {
                if (parenDepth == 0 && braceDepth == 0)
                {
                    args.Add(text.Substring(argStart, i - argStart));
                    closeParenIdx = i;
                    return true;
                }
                if (parenDepth > 0)
                    parenDepth--;
            }
            else if (ch == ',' && parenDepth == 0 && braceDepth == 0)
            {
                args.Add(text.Substring(argStart, i - argStart));
                argStart = i + 1;
            }
        }

        return false;
    }

    private static bool ContainsImageCallOutsideStrings(string formula)
    {
        // Strip string literals first so literal strings mentioning "SMILES2IMG(" aren't misclassified.
        var stripped = new StringBuilder(formula.Length);
        bool inString = false;
        for (int i = 0; i < formula.Length; i++)
        {
            char ch = formula[i];
            if (inString)
            {
                if (ch == '"')
                {
                    if (i + 1 < formula.Length && formula[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    inString = false;
                }
                stripped.Append(' ');
            }
            else if (ch == '"')
            {
                inString = true;
                stripped.Append(' ');
            }
            else
            {
                stripped.Append(ch);
            }
        }

        return CompositeImageCallRegex.IsMatch(stripped.ToString());
    }

    private static string RewriteStyleArgument(string rawArg3, string targetStyle)
    {
        var styleNorm = string.IsNullOrWhiteSpace(targetStyle) ? "color" : targetStyle.Trim();

        // 1. Missing, boolean, numeric, or pure Excel string literal -> preserve structural tokens in string.
        if (string.IsNullOrWhiteSpace(rawArg3) ||
            rawArg3.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
            rawArg3.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
            double.TryParse(rawArg3, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            return QuoteExcelString(styleNorm);
        }

        if (TryUnquoteExcelString(rawArg3, out var unquoted))
        {
            var kept = new List<string> { styleNorm };
            foreach (var token in RenderOptions.ExtractStyleTokens(unquoted))
            {
                if (!RenderOptions.IsColorOrInkToken(token))
                    kept.Add(token);
            }
            return QuoteExcelString(string.Join(" ", kept));
        }

        // 2. Dynamic expression (e.g., "hl:" & $B$3 or cell reference C2) -> unwrap any prior Batch Style prefix.
        var expr = UnwrapPriorBatchPrefix(rawArg3);
        return $"{QuoteExcelString(styleNorm + " ")} & ({expr})";
    }

    private static string UnwrapPriorBatchPrefix(string expr)
    {
        var current = expr.Trim();
        for (int guard = 0; guard < 4; guard++)
        {
            var m = BatchPrefixExprRegex.Match(current);
            if (!m.Success) break;
            var inner = m.Groups["inner"].Value.Trim();
            if (!AreOuterParensMatched(current, current.IndexOf('(')))
                break;
            current = inner;
        }
        return current;
    }

    private static bool AreOuterParensMatched(string text, int openIdx)
    {
        if (openIdx < 0 || openIdx >= text.Length || text[openIdx] != '(')
            return false;

        int depth = 0;
        bool inString = false;
        for (int i = openIdx; i < text.Length; i++)
        {
            char ch = text[i];
            if (inString)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { i++; continue; }
                    inString = false;
                }
                continue;
            }

            if (ch == '"') inString = true;
            else if (ch == '(') depth++;
            else if (ch == ')')
            {
                depth--;
                if (depth == 0) return i == text.Length - 1;
                if (depth < 0) return false;
            }
        }
        return false;
    }

    internal static bool TryUnquoteExcelString(string raw, out string value)
    {
        value = "";
        if (string.IsNullOrEmpty(raw) || raw.Length < 2 || raw[0] != '"' || raw[raw.Length - 1] != '"')
            return false;

        var sb = new StringBuilder(raw.Length - 2);
        for (int i = 1; i < raw.Length; i++)
        {
            char ch = raw[i];
            if (ch == '"')
            {
                if (i + 1 < raw.Length && raw[i + 1] == '"')
                {
                    sb.Append('"');
                    i++;
                    continue;
                }
                // Closing quote must be the final character of raw.
                if (i == raw.Length - 1)
                {
                    value = sb.ToString();
                    return true;
                }
                return false;
            }
            sb.Append(ch);
        }
        return false;
    }

    internal static string QuoteExcelString(string text) =>
        "\"" + (text ?? "").Replace("\"", "\"\"") + "\"";
}
