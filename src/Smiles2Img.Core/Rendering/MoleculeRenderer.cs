using System;
using System.Collections.Generic;
using System.Linq;
using com.epam.indigo;

namespace Smiles2Img.Rendering;

[Serializable]
public sealed class InvalidSmilesException(string message) : Exception(message);

public enum MoleculeSummaryFormat
{
    CompactOneLine,
    MultiLine,
    ExcelTsvRow,
    ExcelTsvWithHeader,
    MarkdownTable,
    Json
}

public sealed class MoleculeInfo
{
    public string CompoundName { get; set; } = "";
    public string IupacName { get; set; } = "";
    public string SubstanceName => !string.IsNullOrWhiteSpace(CompoundName) ? CompoundName : IupacName;
    public string Formula { get; set; } = "";
    public double MolecularWeight { get; set; }
    public double ExactMass { get; set; }
    public double LogP { get; set; }
    public double Tpsa { get; set; }
    public int HBondDonors { get; set; }
    public int HBondAcceptors { get; set; }
    public int RotatableBonds { get; set; }
    public bool PassesLipinski => MolecularWeight <= 500.0 && LogP <= 5.0 && HBondDonors <= 5 && HBondAcceptors <= 10;
    public string Smiles { get; set; } = "";
    public bool IsReaction { get; set; }

    public string InChI { get; set; } = "";
    public string InChIKey { get; set; } = "";

    public string FormatClipboardSummary(MoleculeSummaryFormat format = MoleculeSummaryFormat.CompactOneLine)
    {
        return format switch
        {
            MoleculeSummaryFormat.MultiLine => FormatMultiLine(),
            MoleculeSummaryFormat.ExcelTsvRow => FormatExcelTsv(includeHeader: false),
            MoleculeSummaryFormat.ExcelTsvWithHeader => FormatExcelTsv(includeHeader: true),
            MoleculeSummaryFormat.MarkdownTable => FormatMarkdownTable(),
            MoleculeSummaryFormat.Json => FormatJson(),
            _ => FormatCompactOneLine()
        };
    }

    public string FormatCompactOneLine()
    {
        if (IsReaction)
            return $"Reaction SMILES: {Smiles}";

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(CompoundName))
            parts.Add($"Name: {CompoundName}");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            parts.Add($"IUPAC: {IupacName}");
        parts.Add($"Formula: {Formula}");
        parts.Add($"MW: {MolecularWeight:F2} g/mol");
        if (ExactMass > 0)
            parts.Add($"Exact Mass: {ExactMass:F4}");
        parts.Add($"LogP: {LogP:F2}");
        parts.Add($"TPSA: {Tpsa:F1} Å²");
        parts.Add($"HBD/HBA/RotB: {HBondDonors}/{HBondAcceptors}/{RotatableBonds}");
        parts.Add($"SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            parts.Add($"InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            parts.Add($"InChIKey: {InChIKey}");

        return string.Join(" | ", parts);
    }

    public string FormatMultiLine()
    {
        if (IsReaction)
            return $"Reaction SMILES:\r\n{Smiles}";

        var lines = new List<string>();
        if (!string.IsNullOrWhiteSpace(CompoundName))
            lines.Add($"Substance Name: {CompoundName}");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            lines.Add($"IUPAC Name: {IupacName}");
        lines.Add($"Formula: {Formula}");
        lines.Add($"Molecular Weight: {MolecularWeight:F2} g/mol");
        if (ExactMass > 0)
            lines.Add($"Exact Mass: {ExactMass:F4}");
        lines.Add($"LogP: {LogP:F2}");
        lines.Add($"TPSA: {Tpsa:F1} Å²");
        lines.Add($"H-Bond Donors: {HBondDonors}");
        lines.Add($"H-Bond Acceptors: {HBondAcceptors}");
        lines.Add($"Rotatable Bonds: {RotatableBonds}");
        lines.Add($"Lipinski Rule of 5: {(PassesLipinski ? "Passed" : "Failed")}");
        lines.Add($"SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            lines.Add($"InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            lines.Add($"InChIKey: {InChIKey}");

        return string.Join("\r\n", lines);
    }

    public string FormatExcelTsv(bool includeHeader = false)
    {
        if (IsReaction)
            return includeHeader ? "Type\tSMILES\r\nReaction\t" + Smiles : "Reaction\t" + Smiles;

        var rowValues = new List<string>
        {
            SubstanceName,
            Formula,
            MolecularWeight.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
            ExactMass > 0 ? ExactMass.ToString("F4", System.Globalization.CultureInfo.InvariantCulture) : "",
            LogP.ToString("F2", System.Globalization.CultureInfo.InvariantCulture),
            Tpsa.ToString("F1", System.Globalization.CultureInfo.InvariantCulture),
            HBondDonors.ToString(),
            HBondAcceptors.ToString(),
            RotatableBonds.ToString(),
            PassesLipinski ? "Pass" : "Fail",
            Smiles,
            InChIKey
        };
        string row = string.Join("\t", rowValues);

        if (!includeHeader)
            return row;

        var headerValues = new[]
        {
            "Name", "Formula", "MW (g/mol)", "Exact Mass", "LogP", "TPSA (Å²)",
            "HBD", "HBA", "RotB", "Lipinski", "SMILES", "InChIKey"
        };
        string header = string.Join("\t", headerValues);
        return header + "\r\n" + row;
    }

    public string FormatMarkdownTable()
    {
        if (IsReaction)
            return "| Property | Value |\r\n| :--- | :--- |\r\n| **Type** | Chemical Reaction |\r\n| **SMILES** | `" + Smiles + "` |\r\n";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("| Property | Value |");
        sb.AppendLine("| :--- | :--- |");
        if (!string.IsNullOrWhiteSpace(SubstanceName))
            sb.AppendLine($"| **Substance Name** | {SubstanceName} |");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            sb.AppendLine($"| **IUPAC Name** | {IupacName} |");
        sb.AppendLine($"| **Formula** | {Formula} |");
        sb.AppendLine($"| **Molecular Weight** | {MolecularWeight:F2} g/mol |");
        if (ExactMass > 0)
            sb.AppendLine($"| **Exact Mass** | {ExactMass:F4} |");
        sb.AppendLine($"| **LogP** | {LogP:F2} |");
        sb.AppendLine($"| **TPSA** | {Tpsa:F1} Å² |");
        sb.AppendLine($"| **HBD / HBA / RotB** | {HBondDonors} / {HBondAcceptors} / {RotatableBonds} |");
        sb.AppendLine($"| **Lipinski Rule of 5** | {(PassesLipinski ? "Passed" : "Failed")} |");
        sb.AppendLine($"| **SMILES** | `{Smiles}` |");
        if (!string.IsNullOrWhiteSpace(InChI))
            sb.AppendLine($"| **InChI** | `{InChI}` |");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            sb.AppendLine($"| **InChIKey** | `{InChIKey}` |");

        return sb.ToString().TrimEnd();
    }

    public string FormatJson()
    {
        if (IsReaction)
            return "{\r\n  \"type\": \"reaction\",\r\n  \"smiles\": \"" + EscapeJson(Smiles) + "\"\r\n}";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("{");
        sb.AppendLine($"  \"name\": \"{EscapeJson(SubstanceName)}\",");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            sb.AppendLine($"  \"iupacName\": \"{EscapeJson(IupacName)}\",");
        sb.AppendLine($"  \"formula\": \"{EscapeJson(Formula)}\",");
        sb.AppendLine($"  \"molecularWeight\": {MolecularWeight.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)},");
        if (ExactMass > 0)
            sb.AppendLine($"  \"exactMass\": {ExactMass.ToString("F4", System.Globalization.CultureInfo.InvariantCulture)},");
        sb.AppendLine($"  \"logP\": {LogP.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)},");
        sb.AppendLine($"  \"tpsa\": {Tpsa.ToString("F1", System.Globalization.CultureInfo.InvariantCulture)},");
        sb.AppendLine($"  \"hBondDonors\": {HBondDonors},");
        sb.AppendLine($"  \"hBondAcceptors\": {HBondAcceptors},");
        sb.AppendLine($"  \"rotatableBonds\": {RotatableBonds},");
        sb.AppendLine($"  \"passesLipinski\": {(PassesLipinski ? "true" : "false")},");
        sb.AppendLine($"  \"smiles\": \"{EscapeJson(Smiles)}\",");
        if (!string.IsNullOrWhiteSpace(InChI))
            sb.AppendLine($"  \"inchi\": \"{EscapeJson(InChI)}\",");
        sb.AppendLine($"  \"inchiKey\": \"{EscapeJson(InChIKey)}\"");
        sb.Append("}");
        return sb.ToString();
    }

    private static string EscapeJson(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s!.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    public string FormatHtmlTable()
    {
        var rows = new List<(string Key, string Val)>();
        if (!string.IsNullOrWhiteSpace(SubstanceName)) rows.Add(("Substance Name", SubstanceName));
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            rows.Add(("IUPAC Name", IupacName));
        rows.Add(("Formula", Formula));
        rows.Add(("Molecular Weight", $"{MolecularWeight:F2} g/mol"));
        if (ExactMass > 0) rows.Add(("Exact Mass", $"{ExactMass:F4}"));
        rows.Add(("LogP", $"{LogP:F2}"));
        rows.Add(("TPSA", $"{Tpsa:F1} &#197;&#178;"));
        rows.Add(("HBD / HBA / RotB", $"{HBondDonors} / {HBondAcceptors} / {RotatableBonds}"));
        rows.Add(("Lipinski Rule of 5", PassesLipinski ? "Passed" : "Failed"));
        rows.Add(("SMILES", Smiles));
        if (!string.IsNullOrWhiteSpace(InChIKey)) rows.Add(("InChIKey", InChIKey));

        var sb = new System.Text.StringBuilder();
        sb.Append("<table style=\"border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;font-size:10pt;border:1px solid #d1d5db;\">");
        sb.Append("<thead><tr style=\"background-color:#184c78;color:#ffffff;\">");
        sb.Append("<th style=\"padding:6px 14px;border:1px solid #d1d5db;text-align:left;\">Property</th>");
        sb.Append("<th style=\"padding:6px 14px;border:1px solid #d1d5db;text-align:left;\">Value</th>");
        sb.Append("</tr></thead><tbody>");
        bool alt = false;
        foreach (var (k, v) in rows)
        {
            string bg = alt ? "#f9fafb" : "#ffffff";
            sb.Append($"<tr style=\"background-color:{bg};\">");
            sb.Append($"<td style=\"padding:5px 12px;border:1px solid #e5e7eb;font-weight:600;color:#374151;\">{System.Net.WebUtility.HtmlEncode(k)}</td>");
            sb.Append($"<td style=\"padding:5px 12px;border:1px solid #e5e7eb;color:#111827;\">{System.Net.WebUtility.HtmlEncode(v)}</td>");
            sb.Append("</tr>");
            alt = !alt;
        }
        sb.Append("</tbody></table>");
        return sb.ToString();
    }

    public string FormatDisplayMessage()
    {
        if (IsReaction)
            return $"[Reaction]\n\n• Reaction SMILES: {Smiles}";

        var sb = new System.Text.StringBuilder();
        if (!string.IsNullOrWhiteSpace(CompoundName))
            sb.AppendLine($"• Substance Name: {CompoundName}");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            sb.AppendLine($"• IUPAC Name: {IupacName}");
        sb.AppendLine($"• Formula: {Formula}");
        sb.AppendLine($"• Molecular Weight (MW): {MolecularWeight:F2} g/mol");
        if (ExactMass > 0)
            sb.AppendLine($"• Exact Mass: {ExactMass:F4}");
        sb.AppendLine($"• LogP (Crippen): {LogP:F2} | TPSA: {Tpsa:F1} Å²");
        sb.AppendLine($"• Lipinski Ro5: {(PassesLipinski ? "Pass" : "Violation")} (HBD: {HBondDonors}, HBA: {HBondAcceptors}, RotBonds: {RotatableBonds})");
        sb.AppendLine($"• Canonical SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            sb.AppendLine($"• InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            sb.Append($"• InChIKey: {InChIKey}");
        return sb.ToString().TrimEnd();
    }
}

public static class MoleculeRenderer
{
    private const int RingTopology = 10;
    private static readonly char[] TransformSeparators = [' ', ',', '|', ';'];

    public static bool IsInchi(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text!.Trim();
        return trimmed.StartsWith("InChI=", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("1S/", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("1/", StringComparison.OrdinalIgnoreCase);
    }

    public static string NormalizeInchi(string input)
    {
        var trimmed = input.Trim();
        if (trimmed.StartsWith("InChI=", StringComparison.OrdinalIgnoreCase))
            return trimmed;
        if (trimmed.StartsWith("1S/", StringComparison.OrdinalIgnoreCase) || trimmed.StartsWith("1/", StringComparison.OrdinalIgnoreCase))
            return "InChI=" + trimmed;
        return trimmed;
    }

    public static string Normalize(string smiles)
    {
        var input = smiles?.Trim();
        if (string.IsNullOrEmpty(input) || input!.Length > 2000)
            throw new InvalidSmilesException("Enter a SMILES or InChI string containing 1 to 2000 characters.");

        if (IsInchi(input))
            return NormalizeInchi(input!);

        // Fast syntax validation (bracket and parenthesis balancing)
        var branches = 0;
        var inAtom = false;
        foreach (var character in input)
        {
            if (character == '[' && !inAtom) inAtom = true;
            else if (character == ']' && inAtom) inAtom = false;
            else if (character is '[' or ']') throw Invalid();
            else if (!inAtom && character == '(') branches++;
            else if (!inAtom && character == ')' && --branches < 0) throw Invalid();
        }

        if (inAtom || branches != 0) throw Invalid();
        return input;
    }

    public static byte[] Render(string smiles, bool color = true, string? transform = null, string? background = null)
        => Render(smiles, new RenderOptions { Color = color, Transform = transform, Background = background });

    public static byte[] Render(string smiles, RenderOptions options)
        => RenderCore(smiles, options, "png");

    private static byte[] RenderCore(string smiles, RenderOptions options, string outputFormat)
    {
        NativeIndigo.EnsureInitialized();
        options ??= new RenderOptions();
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);

        indigo.setOption("render-output-format", outputFormat);
        indigo.setOption("render-bond-length", 80.0);
        indigo.setOption("render-margins", 30, 30);
        indigo.setOption("render-relative-thickness", 1.5);
        indigo.setOption("aromaticity-model", "generic");
        indigo.setOption("smart-layout", "true");
        indigo.setOption("layout-orientation", "horizontal");
        indigo.setOption("render-coloring", options.Color);

        ApplyBackground(indigo, options.Background);

        if (!string.IsNullOrWhiteSpace(options.BaseColor))
            ApplyBaseColor(indigo, options.BaseColor);

        if (options.AtomNumbers)
        {
            indigo.setOption("render-atom-ids-visible", "true");
            indigo.setOption("render-atom-bond-ids-from-one", "true");
        }

        indigo.setOption("render-stereo-style", options.StereoLabels ? "ext" : "none");

        using var molecule = Parse(indigo, Normalize(smiles));

        if (!string.IsNullOrWhiteSpace(options.HighlightQuery))
            ApplySubstructureHighlight(indigo, molecule, options.HighlightQuery);

        if (options.UnfoldHydrogens)
            molecule.unfoldHydrogens();
        else
            ExposeRingJunctionChiralHydrogens(molecule);

        if (options.AllCarbons)
            indigo.setOption("render-label-mode", "all");
        else if (molecule.countAtoms() <= 4)
            indigo.setOption("render-label-mode", "terminal-hetero");
        else
            indigo.setOption("render-label-mode", "hetero");

        molecule.dearomatize();
        molecule.layout();
        ApplyTransform(molecule, options.Transform);

        if (!options.AllCarbons)
            LabelHeteroMethyls(molecule);

        return renderer.renderToBuffer(molecule);
    }

    public static bool IsKet(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text!.Trim();
        return trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase) ||
               (trimmed.StartsWith("{") && (trimmed.Contains("\"root\"") || trimmed.Contains("\"nodes\"")));
    }

    public static string CleanKet(string? ket)
    {
        if (string.IsNullOrWhiteSpace(ket)) return "";
        var trimmed = ket!.Trim();
        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            return trimmed.Substring(4).Trim();
        return trimmed;
    }

    public static string? TryConvertToSmiles(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var trimmed = input!.Trim();

        if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            return trimmed.Substring(7).Trim();

        if (IsKet(trimmed))
        {
            var info = TryGetMoleculeInfo(trimmed, resolvePubChem: false);
            if (!string.IsNullOrEmpty(info?.Smiles))
                return info!.Smiles;
            return null;
        }

        if (IsInchi(trimmed))
        {
            NativeIndigo.EnsureInitialized();
            try
            {
                using var indigo = new Indigo();
                using var inchi = new IndigoInchi(indigo);
                using var mol = inchi.loadMolecule(NormalizeInchi(trimmed));
                return mol.canonicalSmiles() ?? mol.smiles();
            }
            catch { }
            return null;
        }

        return trimmed;
    }

    public static string? TryConvertToKet(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var trimmed = input!.Trim();

        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(4).Trim();

        if (IsKet(trimmed))
            return StructureMetadata.CompactKet(trimmed);

        if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(7).Trim();

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            IndigoObject? obj = null;
            try
            {
                if (IsInchi(trimmed))
                {
                    using var inchi = new IndigoInchi(indigo);
                    obj = inchi.loadMolecule(NormalizeInchi(trimmed));
                }
                else
                {
                    obj = indigo.loadMolecule(trimmed);
                }
            }
            catch
            {
                try { obj = indigo.loadReaction(trimmed); } catch { }
            }

            if (obj != null)
            {
                using (obj)
                {
                    var json = obj.json();
                    if (!string.IsNullOrWhiteSpace(json))
                        return StructureMetadata.CompactKet(json);
                }
            }
        }
        catch { }

        return null;
    }

    public static string? TryConvertToInchi(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var trimmed = input!.Trim();

        if (IsInchi(trimmed))
            return NormalizeInchi(trimmed);

        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(4).Trim();
        else if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(7).Trim();

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            using var inchi = new IndigoInchi(indigo);
            using var mol = indigo.loadMolecule(trimmed);
            return inchi.getInchi(mol);
        }
        catch
        {
            var smiles = TryConvertToSmiles(trimmed);
            if (!string.IsNullOrEmpty(smiles) && !string.Equals(smiles, trimmed, StringComparison.Ordinal))
            {
                try
                {
                    using var indigo = new Indigo();
                    using var inchi = new IndigoInchi(indigo);
                    using var mol = indigo.loadMolecule(smiles);
                    return inchi.getInchi(mol);
                }
                catch { }
            }
        }

        return null;
    }

    public static string? TryConvertToInchiKey(string? input)
    {
        var inchiStr = TryConvertToInchi(input);
        if (string.IsNullOrEmpty(inchiStr)) return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            using var inchi = new IndigoInchi(indigo);
            return inchi.getInchiKey(inchiStr);
        }
        catch { }

        return null;
    }

    public static byte[]? RenderKet(string ket, RenderOptions? options = null)
        => RenderKetCore(ket, options, "png");

    private static byte[]? RenderKetCore(string ket, RenderOptions? options, string outputFormat)
    {
        if (string.IsNullOrWhiteSpace(ket)) return null;
        var cleanKet = CleanKet(ket);
        NativeIndigo.EnsureInitialized();
        options ??= new RenderOptions();
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);

        indigo.setOption("render-output-format", outputFormat);
        indigo.setOption("render-bond-length", 80.0);
        indigo.setOption("render-margins", 30, 30);
        indigo.setOption("render-relative-thickness", 1.5);
        indigo.setOption("aromaticity-model", "generic");
        indigo.setOption("render-coloring", options.Color);

        ApplyBackground(indigo, options.Background);

        if (!string.IsNullOrWhiteSpace(options.BaseColor))
            ApplyBaseColor(indigo, options.BaseColor);

        IndigoObject? obj = null;
        var isReaction = false;
        try
        {
            obj = indigo.loadMolecule(cleanKet);
        }
        catch
        {
            try
            {
                obj = indigo.loadReaction(cleanKet);
                isReaction = true;
            }
            catch { }
        }

        if (obj == null) return null;

        using (obj)
        {
            if (!isReaction && !string.IsNullOrWhiteSpace(options.HighlightQuery))
                ApplySubstructureHighlight(indigo, obj, options.HighlightQuery);

            try
            {
                if (!string.IsNullOrWhiteSpace(options.Transform))
                    ApplyTransform(obj, options.Transform);
            }
            catch { }

            return renderer.renderToBuffer(obj);
        }
    }

    public static string? RenderSvg(string? rawInput, RenderOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(rawInput)) return null;
        var input = rawInput!.Trim();
        var smilesIdx = input.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
            input = input.Substring(0, smilesIdx).Trim();

        if (input.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(4).Trim();
        else if (input.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrWhiteSpace(input)) return null;
        options ??= RenderOptions.Default;

        if (input.StartsWith("{", StringComparison.Ordinal) && input.Contains("\"root\""))
        {
            try
            {
                var bytes = RenderKetCore(input, options, "svg");
                if (bytes != null && bytes.Length > 0)
                    return System.Text.Encoding.UTF8.GetString(bytes);
            }
            catch { }
        }

        try
        {
            var bytes = RenderCore(input, options, "svg");
            if (bytes != null && bytes.Length > 0)
                return System.Text.Encoding.UTF8.GetString(bytes);
        }
        catch { }

        try
        {
            var bytes = RenderKetCore(input, options, "svg");
            if (bytes != null && bytes.Length > 0)
                return System.Text.Encoding.UTF8.GetString(bytes);
        }
        catch { }

        return null;
    }

    public static bool? IsSubstructureMatch(string? targetInput, string? queryInput)
    {
        if (string.IsNullOrWhiteSpace(targetInput) || string.IsNullOrWhiteSpace(queryInput))
            return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            using var targetMol = TryLoadMoleculeForAnalysis(indigo, targetInput!);
            if (targetMol == null) return null;

            targetMol.aromatize();

            using var queryObj = TryLoadQueryForMatching(indigo, queryInput!.Trim());
            if (queryObj == null) return null;

            using var matcher = indigo.substructureMatcher(targetMol);
            using var match = matcher.match(queryObj);
            return match != null;
        }
        catch
        {
            return null;
        }
    }

    public static double? CalculateSimilarity(string? input1, string? input2, string? metric = "tanimoto")
    {
        if (string.IsNullOrWhiteSpace(input1) || string.IsNullOrWhiteSpace(input2))
            return null;

        var normMetric = (metric ?? "tanimoto").Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(normMetric) || normMetric is "default" or "tan")
            normMetric = "tanimoto";
        else if (normMetric is "tversky")
            normMetric = "tversky";
        else if (normMetric is "euclid" or "euclid-sub")
            normMetric = "euclid-sub";

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            using var mol1 = TryLoadMoleculeForAnalysis(indigo, input1!);
            using var mol2 = TryLoadMoleculeForAnalysis(indigo, input2!);
            if (mol1 == null || mol2 == null) return null;

            mol1.aromatize();
            mol2.aromatize();

            using var fp1 = mol1.fingerprint("sim");
            using var fp2 = mol2.fingerprint("sim");
            var sim = indigo.similarity(fp1, fp2, normMetric);
            return Math.Round((double)sim, 4);
        }
        catch
        {
            return null;
        }
    }

    private static IndigoObject? TryLoadMoleculeForAnalysis(Indigo indigo, string rawInput)
    {
        var input = rawInput.Trim();
        var smilesIdx = input.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
            input = input.Substring(0, smilesIdx).Trim();

        if (input.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(4).Trim();
        else if (input.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrWhiteSpace(input)) return null;

        try
        {
            if (IsInchi(input))
            {
                using var inchi = new IndigoInchi(indigo);
                var m = inchi.loadMolecule(NormalizeInchi(input));
                if (m.countAtoms() > 0) return m;
                m.Dispose();
                return null;
            }
            else
            {
                var m = indigo.loadMolecule(input);
                if (m.countAtoms() > 0) return m;
                m.Dispose();
                return null;
            }
        }
        catch
        {
            return null;
        }
    }

    private static IndigoObject? TryLoadQueryForMatching(Indigo indigo, string query)
    {
        try
        {
            var q = indigo.loadQueryMolecule(query);
            try { q.aromatize(); } catch { }
            return q;
        }
        catch
        {
            try
            {
                return indigo.loadSmarts(query);
            }
            catch
            {
                return null;
            }
        }
    }

    private static void ApplySubstructureHighlight(Indigo indigo, IndigoObject molecule, string? query)
    {
        if (string.IsNullOrWhiteSpace(query)) return;
        try
        {
            indigo.setOption("render-highlight-color-enabled", true);
            indigo.setOption("render-highlight-color", 0.92f, 0.18f, 0.18f);
            indigo.setOption("render-highlight-thickness-enabled", true);

            molecule.aromatize();
            using var queryObj = TryLoadQueryForMatching(indigo, query!.Trim());
            if (queryObj == null) return;

            using var matcher = indigo.substructureMatcher(molecule);
            foreach (IndigoObject m in matcher.iterateMatches(queryObj))
            {
                using (m)
                {
                    foreach (IndigoObject qa in queryObj.iterateAtoms())
                    {
                        try { m.mapAtom(qa)?.highlight(); } catch { }
                    }
                    foreach (IndigoObject qb in queryObj.iterateBonds())
                    {
                        try { m.mapBond(qb)?.highlight(); } catch { }
                    }
                }
            }
        }
        catch { }
    }

    public static MoleculeInfo? TryGetMoleculeInfo(string? structureOrSmiles, bool resolvePubChem = false)
    {
        if (string.IsNullOrWhiteSpace(structureOrSmiles)) return null;
        var input = structureOrSmiles!.Trim();
        string? fallbackSmiles = null;

        var smilesIdx = input.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
        {
            fallbackSmiles = input.Substring(smilesIdx + 8).Trim();
            input = input.Substring(0, smilesIdx).Trim();
        }

        if (input.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(4).Trim();
        else if (input.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrWhiteSpace(input))
            input = fallbackSmiles ?? "";

        if (string.IsNullOrWhiteSpace(input)) return null;

        var result = TryGetMoleculeInfoCore(input, resolvePubChem);
        if (result == null && !string.IsNullOrWhiteSpace(fallbackSmiles))
            result = TryGetMoleculeInfoCore(fallbackSmiles!, resolvePubChem);
        return result;
    }

    private static MoleculeInfo? TryGetMoleculeInfoCore(string input, bool resolvePubChem)
    {
        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            try
            {
                IndigoObject? mol = null;
                IndigoInchi? inchi = null;
                try
                {
                    if (IsInchi(input))
                    {
                        inchi = new IndigoInchi(indigo);
                        mol = inchi.loadMolecule(NormalizeInchi(input));
                    }
                    else
                    {
                        mol = indigo.loadMolecule(input);
                    }
                }
                catch { }

                if (mol != null && mol.countAtoms() > 0)
                {
                    using (mol)
                    {
                        var rawFormula = "";
                        try { rawFormula = mol.grossFormula()?.Replace(" ", "") ?? ""; } catch { }
                        double mw = 0;
                        try { mw = mol.molecularWeight(); } catch { }
                        double exact = 0;
                        try { exact = mol.monoisotopicMass(); } catch { }
                        double logP = 0;
                        try { logP = Math.Round((double)mol.logP(), 2); } catch { }
                        double tpsa = 0;
                        try { tpsa = Math.Round((double)mol.tpsa(false), 2); } catch { }
                        int hbd = 0;
                        try { hbd = mol.numHydrogenBondDonors(); } catch { }
                        int hba = 0;
                        try { hba = mol.numHydrogenBondAcceptors(); } catch { }
                        int rotBonds = 0;
                        try { rotBonds = mol.numRotatableBonds(); } catch { }
                        var canSmiles = "";
                        try { canSmiles = mol.canonicalSmiles() ?? ""; } catch { try { canSmiles = mol.smiles() ?? ""; } catch { } }

                        string title = "";
                        string iupac = "";
                        if (resolvePubChem && !string.IsNullOrWhiteSpace(canSmiles))
                        {
                            var resolved = PubChemResolver.ResolveCompound(canSmiles);
                            title = resolved.CompoundName;
                            iupac = resolved.IupacName;
                        }

                        string canInchi = "";
                        string canInchiKey = "";
                        try
                        {
                            if (IsInchi(input))
                            {
                                canInchi = NormalizeInchi(input);
                                inchi ??= new IndigoInchi(indigo);
                                canInchiKey = inchi.getInchiKey(canInchi) ?? "";
                            }
                            else
                            {
                                inchi ??= new IndigoInchi(indigo);
                                canInchi = inchi.getInchi(mol) ?? "";
                                if (!string.IsNullOrEmpty(canInchi))
                                    canInchiKey = inchi.getInchiKey(canInchi) ?? "";
                            }
                        }
                        catch { }
                        finally
                        {
                            inchi?.Dispose();
                        }

                        return new MoleculeInfo
                        {
                            CompoundName = title,
                            IupacName = iupac,
                            Formula = rawFormula,
                            MolecularWeight = mw,
                            ExactMass = exact,
                            LogP = logP,
                            Tpsa = tpsa,
                            HBondDonors = hbd,
                            HBondAcceptors = hba,
                            RotatableBonds = rotBonds,
                            Smiles = canSmiles,
                            InChI = canInchi,
                            InChIKey = canInchiKey,
                            IsReaction = false
                        };
                    }
                }
            }
            catch { }

            try
            {
                using var rxn = indigo.loadReaction(input);
                var rxnSmiles = "";
                try { rxnSmiles = rxn.smiles() ?? ""; } catch { }
                if (!string.IsNullOrEmpty(rxnSmiles))
                {
                    return new MoleculeInfo
                    {
                        Smiles = rxnSmiles,
                        IsReaction = true
                    };
                }
            }
            catch { }
        }
        catch { }

        return null;
    }

    public static bool TryRenderStructure(string? rawInput, RenderOptions? options, out byte[]? pngBytes, out string? ket, out string? smiles)
    {
        pngBytes = null;
        ket = null;
        smiles = null;

        if (string.IsNullOrWhiteSpace(rawInput)) return false;
        var input = rawInput!.Trim();
        string? embeddedSmiles = null;

        var smilesIdx = input.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
        {
            embeddedSmiles = input.Substring(smilesIdx + 8).Trim();
            input = input.Substring(0, smilesIdx).Trim();
        }

        if (input.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(4).Trim();
        else if (input.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrWhiteSpace(input)) return false;

        options ??= RenderOptions.Default;

        // 1. Check if input is KET JSON
        if (input.StartsWith("{", StringComparison.Ordinal) && input.Contains("\"root\""))
        {
            try
            {
                pngBytes = RenderKet(input, options);
                if (pngBytes != null && pngBytes.Length > 0)
                {
                    ket = input;
                    smiles = !string.IsNullOrWhiteSpace(embeddedSmiles)
                        ? embeddedSmiles
                        : (TryGetMoleculeInfo(input)?.Smiles ?? "");
                    return true;
                }
            }
            catch { }
        }

        // 2. Try standard SMILES rendering
        try
        {
            pngBytes = Render(input, options);
            if (pngBytes != null && pngBytes.Length > 0)
            {
                smiles = IsInchi(input) ? (TryConvertToSmiles(input) ?? NormalizeInchi(input)) : Normalize(input);
                return true;
            }
        }
        catch { }

        // 3. Try Reaction SMILES or Molfile via RenderKet
        try
        {
            pngBytes = RenderKet(input, options);
            if (pngBytes != null && pngBytes.Length > 0)
            {
                var info = TryGetMoleculeInfo(input);
                smiles = info?.Smiles ?? input;
                return true;
            }
        }
        catch { }

        return false;
    }

    // Expose explicit H only on stereocenters whose non-H bonds are all part of rings
    // (e.g. ring junctions, bridgeheads). This prevents Indigo from putting wedge bonds
    // directly onto ring bonds while keeping standard skeletal formulas everywhere else.
    private static void ExposeRingJunctionChiralHydrogens(IndigoObject molecule)
    {
        if (molecule.countStereocenters() == 0) return;

        var chainBonds = new Dictionary<int, int>();
        foreach (IndigoObject atom in molecule.iterateAtoms()) chainBonds[atom.index()] = 0;
        foreach (IndigoObject bond in molecule.iterateBonds())
        {
            if (bond.topology() != RingTopology)
            {
                chainBonds[bond.source().index()]++;
                chainBonds[bond.destination().index()]++;
            }
        }

        var junctions = new HashSet<int>();
        foreach (IndigoObject sc in molecule.iterateStereocenters())
        {
            var idx = sc.index();
            if (chainBonds[idx] == 0 && sc.countHydrogens() > 0)
                junctions.Add(idx);
        }

        if (junctions.Count == 0) return;

        molecule.unfoldHydrogens();
        var toRemove = new List<int>();
        foreach (IndigoObject atom in molecule.iterateAtoms())
        {
            if (atom.symbol() == "H")
            {
                var keep = atom.iterateNeighbors().Cast<IndigoObject>().Any(n => junctions.Contains(n.index()));
                if (!keep) toRemove.Add(atom.index());
            }
        }

        toRemove.Sort();
        for (var i = toRemove.Count - 1; i >= 0; i--)
            molecule.getAtom(toRemove[i]).remove();
    }

    private static void ApplyTransform(IndigoObject molecule, string? transform)
    {
        if (string.IsNullOrWhiteSpace(transform)) return;
        var tokens = transform!.ToLowerInvariant().Split(TransformSeparators, StringSplitOptions.RemoveEmptyEntries);
        foreach (var token in tokens)
        {
            switch (token)
            {
                case "flip" or "flipx":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(-p[0], p[1], p[2]);
                    }
                    break;
                case "flipy":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(p[0], -p[1], p[2]);
                    }
                    break;
                case "rot90":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(p[1], -p[0], p[2]);
                    }
                    break;
                case "rot180":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(-p[0], -p[1], p[2]);
                    }
                    break;
                case "rot270":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(-p[1], p[0], p[2]);
                    }
                    break;
            }
        }
    }

    private static void LabelHeteroMethyls(IndigoObject molecule)
    {
        if (molecule.countAtoms() <= 4) return;
        foreach (IndigoObject atom in molecule.iterateAtoms())
        {
            if (atom.symbol() == "C" && atom.degree() == 1)
            {
                var neighbor = atom.iterateNeighbors().Cast<IndigoObject>().FirstOrDefault();
                if (neighbor != null)
                {
                    var sym = neighbor.symbol();
                    if (sym is "N" or "O" or "S")
                    {
                        var cPos = atom.xyz();
                        var nPos = neighbor.xyz();
                        atom.resetAtom(cPos[0] < nPos[0] - 0.2f ? "H3C" : "CH3");
                    }
                }
            }
        }
    }

    private static void ApplyBackground(Indigo indigo, string? background)
    {
        if (ColorHelper.TryGetRgb(background, out var r, out var g, out var b))
            indigo.setOption("render-background-color", r, g, b);
        else
            indigo.setOption("render-background-color", -1.0, -1.0, -1.0); // transparent
    }

    private static void ApplyBaseColor(Indigo indigo, string? baseColor)
    {
        if (ColorHelper.TryGetRgb(baseColor, out var r, out var g, out var b))
            indigo.setOption("render-base-color", r, g, b);
    }

    private static IndigoObject Parse(Indigo indigo, string input)
    {
        IndigoObject molecule;
        try
        {
            if (IsInchi(input))
            {
                using var inchi = new IndigoInchi(indigo);
                molecule = inchi.loadMolecule(NormalizeInchi(input));
            }
            else
            {
                molecule = indigo.loadMolecule(input);
            }
        }
        catch (IndigoException ex) { throw new InvalidSmilesException("Indigo: " + ex.Message); }
        if (molecule.countAtoms() == 0)
        {
            molecule.Dispose();
            throw Invalid();
        }
        return molecule;
    }

    private static InvalidSmilesException Invalid() => new("Invalid SMILES or InChI string.");
}
