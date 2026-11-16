using System;
using System.Windows.Forms;
using Smiles2Img.Notifications;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

/// <summary>
/// Shared cross-host command execution helper for Excel, Word, and PowerPoint.
/// Centralizes clipboard copying, format conversion, notifications, and MoleculeInfo dialog display.
/// </summary>
public static class MoleculeActionHelper
{
    public static string? ResolveSmilesWithTextFallback(string? extractedSmiles, string? selectedText)
    {
        if (!string.IsNullOrWhiteSpace(extractedSmiles) || string.IsNullOrWhiteSpace(selectedText))
            return extractedSmiles;
        return MoleculeRenderer.TryConvertToSmiles(selectedText) ?? PubChemResolver.ResolveNameToSmiles(selectedText);
    }

    public static string? ResolveKetWithTextFallback(string? extractedStructure, string? selectedText)
    {
        if (!string.IsNullOrWhiteSpace(extractedStructure) || string.IsNullOrWhiteSpace(selectedText))
            return extractedStructure;
        return MoleculeRenderer.TryConvertToKet(selectedText)
            ?? MoleculeRenderer.TryConvertToKet(PubChemResolver.ResolveNameToSmiles(selectedText));
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
                MessageBox.Show(
                    emptyMessage,
                    "func_SMILES2IMG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (ownerHwnd != IntPtr.Zero)
                MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties", new Win32Window(ownerHwnd));
            else
                MoleculeInfoDialog.ShowModal(info, "Molecule Info & Chemical Properties");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "An error occurred while calculating molecule info:\n" + ex.Message,
                "func_SMILES2IMG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
