using System;
using System.Globalization;
using System.Windows.Forms;
using Smiles2Img.Notifications;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

public enum QuickInsertSourceKind
{
    ExplicitStructure = 0,
    ValidatedCas = 1,
    AmbiguousSmiles = 2,
    ChemicalName = 3
}

public sealed class QuickInsertResolution
{
    public string OriginalInput { get; set; } = "";
    public QuickInsertSourceKind SourceKind { get; set; }
    public bool RequiresConfirmation =>
        SourceKind is QuickInsertSourceKind.AmbiguousSmiles or QuickInsertSourceKind.ChemicalName;
    public byte[] PngBytes { get; set; } = Array.Empty<byte>();
    public string? Ket { get; set; }
    public string? Smiles { get; set; }
    public string Formula { get; set; } = "";
    public double MolecularWeight { get; set; }
    public string ResolvedName { get; set; } = "";
}

/// <summary>
/// Shared cross-host command execution helper for Excel, Word, and PowerPoint.
/// Centralizes clipboard copying, format conversion, notifications, and MoleculeInfo dialog display.
/// </summary>
public static class MoleculeActionHelper
{
    public static string? ResolveSmilesWithTextFallback(string? extractedSmiles, string? selectedText, bool resolvePubChem = true)
    {
        if (!string.IsNullOrWhiteSpace(extractedSmiles))
            return extractedSmiles;
        if (string.IsNullOrWhiteSpace(selectedText))
            return null;

        var trimmed = selectedText!.Trim();
        var info = MoleculeRenderer.TryGetMoleculeInfo(trimmed, resolvePubChem: false);
        if (info != null && !string.IsNullOrWhiteSpace(info.Smiles))
            return info.Smiles;

        if (MoleculeRenderer.TryRenderStructure(trimmed, RenderOptions.Default, out _, out _, out var renderedSmiles) &&
            !string.IsNullOrWhiteSpace(renderedSmiles))
        {
            return renderedSmiles;
        }

        return resolvePubChem
            ? PubChemResolver.ResolveNameToSmiles(trimmed)
            : PubChemResolver.ResolveNameToSmilesOffline(trimmed);
    }

    /// <summary>
    /// Resolves usable structure/SMILES data from extracted metadata first, falling back to
    /// offline structure parsing or chemical name lookup on selected document text.
    /// </summary>
    public static string? ResolveStructureOrText(
        string? extractedPrimary,
        string? extractedSecondary,
        string? selectedText,
        bool resolvePubChem = true)
    {
        if (!string.IsNullOrWhiteSpace(extractedPrimary))
            return extractedPrimary;
        if (!string.IsNullOrWhiteSpace(extractedSecondary))
            return extractedSecondary;
        return ResolveSmilesWithTextFallback(null, selectedText, resolvePubChem);
    }

    public static string? ResolveKetWithTextFallback(string? extractedStructure, string? selectedText)
    {
        if (!string.IsNullOrWhiteSpace(extractedStructure) || string.IsNullOrWhiteSpace(selectedText))
            return extractedStructure;
        return MoleculeRenderer.TryConvertToKet(selectedText)
            ?? MoleculeRenderer.TryConvertToKet(PubChemResolver.ResolveNameToSmiles(selectedText));
    }

    /// <summary>
    /// Validates a CAS Registry Number including its mathematical check digit.
    /// </summary>
    public static bool IsValidCasNumber(string? input) => PubChemResolver.IsValidCasNumber(input);

    /// <summary>
    /// Identifies plain-letter strings that are valid SMILES syntactically yet easily collide
    /// with ordinary English words, initials, or element symbols (e.g. "I", "C", "CO", "NO", "COP", "POP", "COCO").
    /// Explicit prefixes ("SMILES:", "InChI=", "KET:"), punctuation/ring digits, and "CC..." aliphatic chains (e.g. "CCO") are allowed.
    /// </summary>
    public static bool IsAmbiguousPlainAlphaSmiles(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return false;
        var trimmed = input!.Trim();

        if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("InChI=", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("1S/", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("1/", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("{", StringComparison.Ordinal))
        {
            return false;
        }

        for (int i = 0; i < trimmed.Length; i++)
        {
            char ch = trimmed[i];
            bool isAsciiLetter = (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
            if (!isAsciiLetter)
                return false;
        }

        if (trimmed.Length <= 2)
            return true;

        return !trimmed.StartsWith("CC", StringComparison.Ordinal);
    }

    private static bool LooksLikeChemicalNameOrCas(string input)
    {
        if (string.IsNullOrWhiteSpace(input) || input.Length < 2 || input.Length > 160)
            return false;
        if (input.IndexOf('\r') >= 0 || input.IndexOf('\n') >= 0)
            return false;
        if (input.StartsWith("{", StringComparison.Ordinal) ||
            input.StartsWith("=", StringComparison.Ordinal) ||
            input.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            input.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            return false;

        // Reject malformed digit-hyphen-digit triplets that fail the CAS check digit.
        if (PubChemResolver.IsNumericTriplet(input) && !IsValidCasNumber(input))
            return false;

        var tokens = input.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return tokens.Length is >= 1 and <= 6;
    }

    /// <summary>
    /// Resolves and renders quick-insert input from SMILES, InChI, KET, checksum-validated CAS,
    /// or chemical name, classifying whether user confirmation is required before insertion.
    /// </summary>
    public static bool TryResolveQuickInsertInput(
        string? rawInput,
        RenderOptions? options,
        out QuickInsertResolution? resolution,
        bool resolvePubChem = true)
    {
        resolution = null;
        if (string.IsNullOrWhiteSpace(rawInput))
            return false;

        var trimmed = rawInput!.Trim();
        if (MoleculeRenderer.TryRenderStructure(trimmed, options, out var pngBytes, out var ket, out var smiles) &&
            pngBytes != null)
        {
            var kind = IsAmbiguousPlainAlphaSmiles(trimmed)
                ? QuickInsertSourceKind.AmbiguousSmiles
                : QuickInsertSourceKind.ExplicitStructure;

            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles ?? trimmed, resolvePubChem: false);
            resolution = new QuickInsertResolution
            {
                OriginalInput = trimmed,
                SourceKind = kind,
                PngBytes = pngBytes,
                Ket = ket,
                Smiles = smiles,
                Formula = info?.Formula ?? "",
                MolecularWeight = info?.MolecularWeight ?? 0
            };
            return true;
        }

        var isCas = IsValidCasNumber(trimmed);
        if (!isCas && !LooksLikeChemicalNameOrCas(trimmed))
            return false;

        var resolvedSmiles = resolvePubChem
            ? PubChemResolver.ResolveNameToSmiles(trimmed)
            : PubChemResolver.ResolveNameToSmilesOffline(trimmed);
        if (string.IsNullOrWhiteSpace(resolvedSmiles))
            return false;

        if (!MoleculeRenderer.TryRenderStructure(resolvedSmiles, options, out pngBytes, out ket, out smiles) ||
            pngBytes == null)
        {
            return false;
        }

        var molInfo = MoleculeRenderer.TryGetMoleculeInfo(smiles ?? resolvedSmiles, resolvePubChem: false);
        var resolvedName = "";
        if (!isCas)
        {
            try
            {
                resolvedName = PubChemResolver.ResolveCompound(smiles ?? resolvedSmiles, timeoutMs: 1).CompoundName;
            }
            catch { }
        }

        resolution = new QuickInsertResolution
        {
            OriginalInput = trimmed,
            SourceKind = isCas ? QuickInsertSourceKind.ValidatedCas : QuickInsertSourceKind.ChemicalName,
            PngBytes = pngBytes,
            Ket = ket,
            Smiles = smiles,
            Formula = molInfo?.Formula ?? "",
            MolecularWeight = molInfo?.MolecularWeight ?? 0,
            ResolvedName = !string.IsNullOrWhiteSpace(resolvedName) ? resolvedName : trimmed
        };
        return true;
    }

    public static bool ConfirmQuickInsertIfNeeded(QuickInsertResolution? resolution, IntPtr ownerHwnd)
    {
        if (resolution == null) return false;
        if (!resolution.RequiresConfirmation) return true;

        var mwText = resolution.MolecularWeight > 0
            ? $" (MW {resolution.MolecularWeight.ToString("F2", CultureInfo.InvariantCulture)} g/mol)"
            : "";
        var formulaLine = !string.IsNullOrWhiteSpace(resolution.Formula)
            ? $"• Formula: {resolution.Formula}{mwText}\n"
            : "";
        var smilesLine = !string.IsNullOrWhiteSpace(resolution.Smiles)
            ? $"• SMILES: {resolution.Smiles}\n"
            : "";

        string message;
        if (resolution.SourceKind == QuickInsertSourceKind.AmbiguousSmiles)
        {
            message =
                $"Text \"{resolution.OriginalInput}\" could be an ordinary word or abbreviation, or a SMILES string:\n\n" +
                formulaLine +
                smilesLine +
                "\nInsert this molecular structure?\n(Tip: prefix with \"SMILES:\" to skip confirmation.)";
        }
        else
        {
            var nameLine = !string.IsNullOrWhiteSpace(resolution.ResolvedName)
                ? $"• Matched Name: {resolution.ResolvedName}\n"
                : "";
            message =
                $"Interpreted \"{resolution.OriginalInput}\" as a chemical compound name:\n\n" +
                nameLine +
                formulaLine +
                smilesLine +
                "\nInsert this molecular structure?";
        }

        return DialogHelper.ConfirmYesNo(message, DialogHelper.ToOwner(ownerHwnd)) == DialogResult.Yes;
    }

    /// <summary>
    /// Renders quick-insert input from SMILES, InChI, KET, or chemical name/CAS (with offline dictionary + optional PubChem lookup).
    /// </summary>
    public static bool TryRenderQuickInsertInput(
        string? rawInput,
        RenderOptions? options,
        out byte[]? pngBytes,
        out string? ket,
        out string? smiles,
        bool resolvePubChem = true)
    {
        if (TryResolveQuickInsertInput(rawInput, options, out var res, resolvePubChem) && res != null)
        {
            pngBytes = res.PngBytes;
            ket = res.Ket;
            smiles = res.Smiles;
            return true;
        }
        pngBytes = null;
        ket = null;
        smiles = null;
        return false;
    }

    public static MoleculeInfo? TryResolveMolstarInfo(string? altText, string? title, bool resolvePubChem)
    {
        if (!StructureMetadata.IsMolstar(altText, title))
            return null;

        var meta = MolstarSessionManager.ParseAltText(altText, title);
        if (meta == null || string.IsNullOrWhiteSpace(meta.Pdb))
            return null;

        (string Title, double? WeightKDa) pdbInfo = resolvePubChem ? PubChemResolver.ResolvePdb(meta.Pdb) : ("", null);
        return new MoleculeInfo
        {
            CompoundName = !string.IsNullOrEmpty(pdbInfo.Title) ? pdbInfo.Title : $"PDB {meta.Pdb}",
            IupacName = $"PDB Entry: {meta.Pdb}",
            Formula = "Macromolecular Complex",
            MolecularWeight = pdbInfo.WeightKDa.HasValue ? pdbInfo.WeightKDa.Value * 1000.0 : 0.0,
            ExactMass = 0,
            Smiles = $"PDB:{meta.Pdb}",
            IsReaction = false,
            IsMacromolecule = true
        };
    }

    public static void CopySmiles(string? smiles, string emptyWarning)
    {
        try
        {
            if (string.IsNullOrEmpty(smiles))
            {
                MoleculeNotification.NotifyWarning("SMILES", emptyWarning);
                return;
            }
            SafeClipboard.SetText(smiles!);
            MoleculeNotification.NotifyCopied("SMILES", smiles!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("SMILES", "Error extracting SMILES: " + ex.Message);
        }
    }

    public static void CopyInchi(string? targetData, string emptyWarning)
    {
        try
        {
            var inchi = MoleculeRenderer.TryConvertToInchi(targetData);
            if (string.IsNullOrEmpty(inchi))
            {
                MoleculeNotification.NotifyWarning("InChI", emptyWarning);
                return;
            }
            SafeClipboard.SetText(inchi!);
            MoleculeNotification.NotifyCopied("InChI", inchi!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChI", "Error generating InChI: " + ex.Message);
        }
    }

    public static void CopyKet(string? rawStructureOrKet, string emptyWarning, bool alreadyKet = false)
    {
        try
        {
            var ket = alreadyKet ? rawStructureOrKet : MoleculeRenderer.TryConvertToKet(rawStructureOrKet);
            if (string.IsNullOrEmpty(ket))
            {
                MoleculeNotification.NotifyWarning("Canvas Data (KET)", emptyWarning);
                return;
            }
            SafeClipboard.SetText(ket!);
            MoleculeNotification.NotifyCopied("Canvas Data (KET)", ket!);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Canvas Data (KET)", "Error copying canvas data: " + ex.Message);
        }
    }

    public static void CopySvg(string? targetData, string emptyWarning, string title = "SVG Vector", string successMessage = "Vector graphic copied to clipboard")
    {
        try
        {
            var svg = MoleculeRenderer.RenderSvg(targetData, RenderOptions.Default);
            if (string.IsNullOrEmpty(svg) || !SvgClipboardHelper.CopySvgToClipboard(svg))
            {
                MoleculeNotification.NotifyWarning(title, emptyWarning);
                return;
            }
            MoleculeNotification.NotifyCopied(title, successMessage);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning(title, "Error copying SVG: " + ex.Message);
        }
    }

    public static void CopySubstanceName(Func<bool, MoleculeInfo?> getInfo, string emptyWarning)
    {
        try
        {
            var info = getInfo(true);
            if (info == null || string.IsNullOrWhiteSpace(info.CompoundName))
            {
                MoleculeNotification.NotifyWarning("Substance Name", emptyWarning);
                return;
            }
            SafeClipboard.SetText(info.CompoundName);
            MoleculeNotification.NotifyCopied("Substance Name", info.CompoundName);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Substance Name", "Error copying substance name: " + ex.Message);
        }
    }

    public static void CopyFormula(Func<bool, MoleculeInfo?> getInfo, string emptyWarning, string copiedTitle = "Chemical Formula")
    {
        try
        {
            var info = getInfo(false);
            if (info == null || string.IsNullOrWhiteSpace(info.Formula))
            {
                MoleculeNotification.NotifyWarning("Formula", emptyWarning);
                return;
            }
            SafeClipboard.SetText(info.Formula);
            MoleculeNotification.NotifyCopied(copiedTitle, info.Formula);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Formula", "Error copying formula: " + ex.Message);
        }
    }

    public static void CopyMolecularWeight(Func<bool, MoleculeInfo?> getInfo, string emptyWarning)
    {
        try
        {
            var info = getInfo(false);
            if (info == null || info.MolecularWeight <= 0)
            {
                MoleculeNotification.NotifyWarning("Molecular Weight", emptyWarning);
                return;
            }
            string mw = info.MolecularWeight.ToString("F2");
            SafeClipboard.SetText(mw);
            MoleculeNotification.NotifyCopied("Molecular Weight", $"{mw} g/mol");
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecular Weight", "Error copying molecular weight: " + ex.Message);
        }
    }

    public static void CopyInchiKey(Func<bool, MoleculeInfo?> getInfo, string emptyWarning)
    {
        try
        {
            var info = getInfo(false);
            if (info == null || string.IsNullOrWhiteSpace(info.InChIKey))
            {
                MoleculeNotification.NotifyWarning("InChIKey", emptyWarning);
                return;
            }
            SafeClipboard.SetText(info.InChIKey);
            MoleculeNotification.NotifyCopied("InChIKey", info.InChIKey);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("InChIKey", "Error copying InChIKey: " + ex.Message);
        }
    }

    public static void CopySummary(
        Func<bool, MoleculeInfo?> getInfo,
        MoleculeSummaryFormat format,
        string copiedTitle,
        string fallbackDetail,
        string emptyWarning,
        string errorActionDescription)
    {
        try
        {
            var info = getInfo(true);
            if (info == null)
            {
                MoleculeNotification.NotifyWarning("Molecule Info", emptyWarning);
                return;
            }
            string text = info.FormatClipboardSummary(format);
            string? html = info.FormatHtmlForSummary(format);
            SafeClipboard.SetTextAndHtml(text, html);
            MoleculeNotification.NotifyCopied(copiedTitle, info.SubstanceName ?? info.Formula ?? fallbackDetail);
        }
        catch (Exception ex)
        {
            MoleculeNotification.NotifyWarning("Molecule Info", $"Error copying {errorActionDescription}: " + ex.Message);
        }
    }

    public static void ShowMoleculeInfoDialog(Func<bool, MoleculeInfo?> getInfo, IntPtr ownerHwnd, string emptyMessage)
    {
        try
        {
            var info = getInfo(true);
            if (info == null)
            {
                DialogHelper.ShowInfo(emptyMessage, ownerHwnd);
                return;
            }

            MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties", DialogHelper.ToOwner(ownerHwnd));
        }
        catch (Exception ex)
        {
            DialogHelper.ShowWarning("An error occurred while calculating molecule info:\n" + ex.Message, ownerHwnd);
        }
    }
}
