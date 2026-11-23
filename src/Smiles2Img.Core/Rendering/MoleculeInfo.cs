using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;

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
    public bool HasQueryOrDummyAtoms { get; set; }
    public double LogP { get; set; }
    public bool HasLogP { get; set; }
    public double Tpsa { get; set; }
    public bool HasTpsa { get; set; }
    public double TpsaWithSP { get; set; }
    public bool HasTpsaWithSP { get; set; }
    public int HeavyAtomCount { get; set; }
    public int RingCount { get; set; }
    public int AromaticRingCount { get; set; }
    public double Fsp3 { get; set; }
    public string Smiles { get; set; } = "";
    public bool IsReaction { get; set; }
    public bool IsMacromolecule { get; set; }

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
        if (!IsMacromolecule) parts.Add($"LogP: {LogP:F2}");
        if (!IsMacromolecule) parts.Add($"TPSA: {Tpsa:F1} Å²");
        if (HeavyAtomCount > 0)
            parts.Add($"Heavy/Rings/Arom/Fsp3: {HeavyAtomCount}/{RingCount}/{AromaticRingCount}/{Fsp3:F2}");
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
        if (!IsMacromolecule)
        {
            lines.Add($"LogP: {LogP:F2}");
            lines.Add($"TPSA: {Tpsa:F1} Å²");
            lines.Add($"Heavy Atoms: {HeavyAtomCount}");
            lines.Add($"Rings (Total / Aromatic): {RingCount} / {AromaticRingCount}");
            lines.Add($"Fraction Csp3 (Fsp3): {Fsp3:F2}");
        }
        lines.Add($"SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            lines.Add($"InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            lines.Add($"InChIKey: {InChIKey}");

        return string.Join("\r\n", lines);
    }

    private static string CleanTsvField(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s!.Replace("\t", " ").Replace("\r", "").Replace("\n", " ").Trim();
    }

    public string FormatExcelTsv(bool includeHeader = false)
    {
        if (IsReaction)
            return includeHeader ? "Type\tSMILES\r\nReaction\t" + Smiles : "Reaction\t" + Smiles;

        var rowValues = new List<string>
        {
            CleanTsvField(SubstanceName),
            CleanTsvField(Formula),
            MolecularWeight.ToString("F2", CultureInfo.InvariantCulture),
            ExactMass > 0 ? ExactMass.ToString("F4", CultureInfo.InvariantCulture) : "",
            LogP.ToString("F2", CultureInfo.InvariantCulture),
            Tpsa.ToString("F1", CultureInfo.InvariantCulture),
            HeavyAtomCount.ToString(),
            RingCount.ToString(),
            AromaticRingCount.ToString(),
            Fsp3.ToString("F2", CultureInfo.InvariantCulture),
            CleanTsvField(Smiles),
            CleanTsvField(InChIKey)
        };
        string row = string.Join("\t", rowValues);

        if (!includeHeader)
            return row;

        var headerValues = new[]
        {
            "Name", "Formula", "MW (g/mol)", "Exact Mass", "LogP", "TPSA (Å²)",
            "Heavy Atoms", "Rings", "Aromatic Rings", "Fsp3", "SMILES", "InChIKey"
        };
        string header = string.Join("\t", headerValues);
        return header + "\r\n" + row;
    }

    public string FormatMarkdownTable()
    {
        if (IsReaction)
            return "| Property | Value |\r\n| :--- | :--- |\r\n| **Type** | Chemical Reaction |\r\n| **SMILES** | `" + Smiles + "` |\r\n";

        var sb = new StringBuilder();
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
        if (!IsMacromolecule) sb.AppendLine($"| **LogP** | {LogP:F2} |");
        if (!IsMacromolecule) sb.AppendLine($"| **TPSA** | {Tpsa:F1} Å² |");
        if (!IsMacromolecule) sb.AppendLine($"| **Heavy Atoms** | {HeavyAtomCount} |");
        if (!IsMacromolecule) sb.AppendLine($"| **Rings (Total / Aromatic)** | {RingCount} / {AromaticRingCount} |");
        if (!IsMacromolecule) sb.AppendLine($"| **Fraction Csp3 (Fsp3)** | {Fsp3:F2} |");
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

        var sb = new StringBuilder();
        sb.AppendLine("{");
        sb.AppendLine($"  \"name\": \"{EscapeJson(SubstanceName)}\",");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            sb.AppendLine($"  \"iupacName\": \"{EscapeJson(IupacName)}\",");
        sb.AppendLine($"  \"formula\": \"{EscapeJson(Formula)}\",");
        sb.AppendLine($"  \"molecularWeight\": {MolecularWeight.ToString("F2", CultureInfo.InvariantCulture)},");
        if (ExactMass > 0)
            sb.AppendLine($"  \"exactMass\": {ExactMass.ToString("F4", CultureInfo.InvariantCulture)},");
        if (!IsMacromolecule)
        {
            sb.AppendLine($"  \"logP\": {LogP.ToString("F2", CultureInfo.InvariantCulture)},");
            sb.AppendLine($"  \"tpsa\": {Tpsa.ToString("F1", CultureInfo.InvariantCulture)},");
            sb.AppendLine($"  \"heavyAtomCount\": {HeavyAtomCount},");
            sb.AppendLine($"  \"ringCount\": {RingCount},");
            sb.AppendLine($"  \"aromaticRingCount\": {AromaticRingCount},");
            sb.AppendLine($"  \"fsp3\": {Fsp3.ToString("F2", CultureInfo.InvariantCulture)},");
        }
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
        if (!IsMacromolecule) rows.Add(("LogP", $"{LogP:F2}"));
        if (!IsMacromolecule) rows.Add(("TPSA", $"{Tpsa:F1} Å²"));
        if (!IsMacromolecule) rows.Add(("Heavy Atoms", $"{HeavyAtomCount}"));
        if (!IsMacromolecule) rows.Add(("Rings (Total / Aromatic)", $"{RingCount} / {AromaticRingCount}"));
        if (!IsMacromolecule) rows.Add(("Fraction Csp3 (Fsp3)", $"{Fsp3:F2}"));
        rows.Add(("SMILES", Smiles));
        if (!string.IsNullOrWhiteSpace(InChIKey)) rows.Add(("InChIKey", InChIKey));

        var sb = new StringBuilder();
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
            sb.Append($"<td style=\"padding:5px 12px;border:1px solid #e5e7eb;font-weight:600;color:#374151;\">{WebUtility.HtmlEncode(k)}</td>");
            sb.Append($"<td style=\"padding:5px 12px;border:1px solid #e5e7eb;color:#111827;\">{WebUtility.HtmlEncode(v)}</td>");
            sb.Append("</tr>");
            alt = !alt;
        }
        sb.Append("</tbody></table>");
        return sb.ToString();
    }

    public string FormatHtmlHorizontalTable(bool includeHeader = true)
    {
        if (IsReaction)
        {
            return "<table style=\"border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;font-size:10pt;border:1px solid #d1d5db;\">" +
                   "<thead><tr style=\"background-color:#184c78;color:#ffffff;\"><th style=\"padding:6px 14px;border:1px solid #d1d5db;text-align:left;\">Type</th><th style=\"padding:6px 14px;border:1px solid #d1d5db;text-align:left;\">SMILES</th></tr></thead>" +
                   $"<tbody><tr><td style=\"padding:5px 12px;border:1px solid #e5e7eb;\">Reaction</td><td style=\"padding:5px 12px;border:1px solid #e5e7eb;\">{WebUtility.HtmlEncode(Smiles)}</td></tr></tbody></table>";
        }

        var headerValues = new[]
        {
            "Name", "Formula", "MW (g/mol)", "Exact Mass", "LogP", "TPSA (Å²)",
            "Heavy Atoms", "Rings", "Aromatic Rings", "Fsp3", "SMILES", "InChIKey"
        };
        var rowValues = new[]
        {
            CleanTsvField(SubstanceName),
            CleanTsvField(Formula),
            MolecularWeight.ToString("F2", CultureInfo.InvariantCulture),
            ExactMass > 0 ? ExactMass.ToString("F4", CultureInfo.InvariantCulture) : "",
            LogP.ToString("F2", CultureInfo.InvariantCulture),
            Tpsa.ToString("F1", CultureInfo.InvariantCulture),
            HeavyAtomCount.ToString(),
            RingCount.ToString(),
            AromaticRingCount.ToString(),
            Fsp3.ToString("F2", CultureInfo.InvariantCulture),
            CleanTsvField(Smiles),
            CleanTsvField(InChIKey)
        };

        var sb = new StringBuilder();
        sb.Append("<table style=\"border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;font-size:10pt;border:1px solid #d1d5db;\">");
        if (includeHeader)
        {
            sb.Append("<thead><tr style=\"background-color:#184c78;color:#ffffff;\">");
            foreach (var h in headerValues)
            {
                sb.Append($"<th style=\"padding:6px 12px;border:1px solid #d1d5db;text-align:left;\">{WebUtility.HtmlEncode(h)}</th>");
            }
            sb.Append("</tr></thead>");
        }
        sb.Append("<tbody><tr>");
        foreach (var v in rowValues)
        {
            sb.Append($"<td style=\"padding:5px 12px;border:1px solid #e5e7eb;color:#111827;\">{WebUtility.HtmlEncode(v)}</td>");
        }
        sb.Append("</tr></tbody></table>");
        return sb.ToString();
    }

    public string? FormatHtmlForSummary(MoleculeSummaryFormat format)
    {
        return format switch
        {
            MoleculeSummaryFormat.MultiLine => FormatHtmlTable(),
            MoleculeSummaryFormat.ExcelTsvRow => FormatHtmlHorizontalTable(includeHeader: false),
            MoleculeSummaryFormat.ExcelTsvWithHeader => FormatHtmlHorizontalTable(includeHeader: true),
            MoleculeSummaryFormat.MarkdownTable => FormatHtmlTable(),
            _ => null
        };
    }

    public string FormatDisplayMessage()
    {
        if (IsReaction)
            return $"[Reaction]\n\n• Reaction SMILES: {Smiles}";

        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(CompoundName))
            sb.AppendLine($"• Substance Name: {CompoundName}");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            sb.AppendLine($"• IUPAC Name: {IupacName}");
        sb.AppendLine($"• Formula: {Formula}");
        sb.AppendLine($"• Molecular Weight (MW): {MolecularWeight:F2} g/mol");
        if (ExactMass > 0)
            sb.AppendLine($"• Exact Mass: {ExactMass:F4}");
        if (!IsMacromolecule)
        {
            sb.AppendLine($"• LogP (Crippen): {LogP:F2} | TPSA: {Tpsa:F1} Å²");
            sb.AppendLine($"• Complexity: Heavy={HeavyAtomCount}, Rings={RingCount} (Arom={AromaticRingCount}), Fsp3={Fsp3:F2}");
        }
        sb.AppendLine($"• Canonical SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            sb.AppendLine($"• InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            sb.Append($"• InChIKey: {InChIKey}");
        return sb.ToString().TrimEnd();
    }
}
