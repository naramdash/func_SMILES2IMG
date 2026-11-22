using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Smiles2Img.Rendering;

namespace Smiles2Img;

/// <summary>
/// Shared helper for embedding, extracting, and resolving KET/SMILES structure metadata
/// and aspect-ratio-preserving image geometry across Word, PowerPoint, and Excel.
/// </summary>
public static class StructureMetadata
{
    public const string KetPrefix = "KET:";
    public const string SmilesPrefix = "SMILES:";
    public const string MolstarPrefix = "MOLSTAR:";

    public static bool IsMolstar(params string?[] texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            if (text!.IndexOf(MolstarPrefix, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }
        return false;
    }

    private static readonly Regex SmilesTagRegex = new(
        @"(?:^|\b|\s)SMILES:\s*(?<smiles>[^\r\n]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static string FormatAlternativeText(string? ket, string? smiles)
    {
        if (!string.IsNullOrEmpty(ket))
            return KetPrefix + ket;
        return SmilesPrefix + (smiles ?? "");
    }

    public static string FormatTitle(string? smiles, string defaultTitle = "Molecule")
    {
        if (!string.IsNullOrEmpty(smiles))
            return SmilesPrefix + smiles;
        return defaultTitle;
    }

    public static string? ExtractStructure(params string?[] texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var trimmed = text!.Trim();

            if (trimmed.StartsWith(KetPrefix, StringComparison.OrdinalIgnoreCase))
                return StripTrailingSmilesTag(trimmed.Substring(KetPrefix.Length).Trim());

            var ketIdx = trimmed.IndexOf(KetPrefix, StringComparison.OrdinalIgnoreCase);
            if (ketIdx >= 0)
                return StripTrailingSmilesTag(trimmed.Substring(ketIdx + KetPrefix.Length).Trim());
        }

        return ExtractSmiles(texts);
    }

    public static string? ExtractSmiles(params string?[] texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var trimmed = text!.Trim();

            if (trimmed.StartsWith(SmilesPrefix, StringComparison.OrdinalIgnoreCase))
                return trimmed.Substring(SmilesPrefix.Length).Trim();

            var match = SmilesTagRegex.Match(trimmed);
            if (match.Success)
                return match.Groups["smiles"].Value.Trim();
        }

        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var trimmed = text!.Trim();
            if (MoleculeRenderer.IsKet(trimmed))
            {
                var converted = MoleculeRenderer.TryConvertToSmiles(trimmed);
                if (!string.IsNullOrEmpty(converted))
                    return converted;
            }
            if (MoleculeRenderer.IsInchi(trimmed))
            {
                var converted = MoleculeRenderer.TryConvertToSmiles(trimmed);
                if (!string.IsNullOrEmpty(converted))
                    return converted;
                return MoleculeRenderer.NormalizeInchi(trimmed);
            }
        }

        return null;
    }

    public static byte[]? ResolvePngBytes(KetcherForm form, string? ket, string? smiles)
    {
        var pngBytes = form.ResultPngBytes;
        if (pngBytes != null && pngBytes.Length > 0) return pngBytes;

        var opts = new RenderOptions { Background = form.ResultBackground };

        if (!string.IsNullOrWhiteSpace(ket))
        {
            try
            {
                pngBytes = MoleculeRenderer.RenderKet(ket!, opts);
                if (pngBytes != null && pngBytes.Length > 0) return pngBytes;
            }
            catch { }
        }

        if (!string.IsNullOrWhiteSpace(smiles))
        {
            try
            {
                pngBytes = MoleculeRenderer.Render(smiles!, opts);
                if (pngBytes != null && pngBytes.Length > 0) return pngBytes;
            }
            catch { }
        }

        return null;
    }

    public static (float Left, float Top, float Width, float Height) FitReplacementBounds(
        float oldLeft, float oldTop, float oldWidth, float oldHeight, byte[] pngBytes)
    {
        float left = oldLeft;
        float top = oldTop;
        float newWidth = oldWidth;
        float newHeight = oldHeight;

        try
        {
            using var ms = new MemoryStream(pngBytes);
            using var img = Image.FromStream(ms);
            float aspect = img.Width / (float)img.Height;
            if (oldWidth / oldHeight > aspect)
            {
                newHeight = oldHeight;
                newWidth = oldHeight * aspect;
                left += (oldWidth - newWidth) / 2f;
            }
            else
            {
                newWidth = oldWidth;
                newHeight = oldWidth / aspect;
                top += (oldHeight - newHeight) / 2f;
            }
        }
        catch { }

        return (left, top, newWidth, newHeight);
    }

    /// <summary>
    /// Computes centered, aspect-ratio-preserving bounds inside a slide when converting a PowerPoint
    /// TextBox in-place, expanding thin single-line TextBoxes into a readable target box clamped to the slide.
    /// </summary>
    public static (float Left, float Top, float Width, float Height) FitTextBoxReplacementBounds(
        float boxLeft, float boxTop, float boxWidth, float boxHeight, float slideWidth, float slideHeight, byte[] pngBytes)
    {
        if (boxWidth >= 90f && boxHeight >= 75f)
            return FitReplacementBounds(boxLeft, boxTop, boxWidth, boxHeight, pngBytes);

        float cx = boxLeft + boxWidth / 2f;
        float cy = boxTop + boxHeight / 2f;
        float targetWidth = Math.Min(Math.Max(boxWidth, 160f), Math.Min(240f, Math.Max(120f, slideWidth - 24f)));
        float targetHeight = Math.Min(Math.Max(boxHeight, 88f), Math.Min(130f, Math.Max(80f, slideHeight - 24f)));

        var (left, top, width, height) = FitReplacementBounds(
            cx - targetWidth / 2f,
            cy - targetHeight / 2f,
            targetWidth,
            targetHeight,
            pngBytes);

        left = Math.Max(12f, Math.Min(left, Math.Max(12f, slideWidth - width - 12f)));
        top = Math.Max(12f, Math.Min(top, Math.Max(12f, slideHeight - height - 12f)));
        return (left, top, width, height);
    }

    public static (float Width, float Height) GetScaledDimensions(
        byte[] pngBytes, float defaultWidth = 300f, float defaultHeight = 200f, float scale = 1.0f, float minWidth = 0f, float maxWidth = float.MaxValue)
    {
        try
        {
            using var ms = new MemoryStream(pngBytes);
            using var img = Image.FromStream(ms);
            float aspect = img.Width / (float)img.Height;
            float width = img.Width * scale;
            if (minWidth > 0f || maxWidth < float.MaxValue)
                width = Math.Max(minWidth, Math.Min(maxWidth, width));
            float height = width / aspect;
            return (width, height);
        }
        catch
        {
            return (defaultWidth, defaultHeight);
        }
    }

    public static string CompactKet(string? ket)
    {
        if (string.IsNullOrWhiteSpace(ket)) return "";
        try
        {
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var obj = serializer.DeserializeObject(ket!);
            return serializer.Serialize(obj);
        }
        catch
        {
            return ket!.Trim();
        }
    }

    private static string StripTrailingSmilesTag(string ketText)
    {
        var smilesIdx = ketText.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
            return ketText.Substring(0, smilesIdx).Trim();
        return ketText;
    }

    /// <summary>
    /// Writes binary payload to a unique temporary file, runs <paramref name="action"/>,
    /// and guarantees best-effort deletion afterward.
    /// </summary>
    public static T WithTempFile<T>(byte[] bytes, string filePrefix, string extension, Func<string, T> action)
    {
        var ext = extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;
        var tempFile = Path.Combine(Path.GetTempPath(), $"{filePrefix}_{Guid.NewGuid():N}{ext}");
        try
        {
            File.WriteAllBytes(tempFile, bytes);
            return action(tempFile);
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    /// <summary>
    /// Writes binary payload to a unique temporary file, runs <paramref name="action"/>,
    /// and guarantees best-effort deletion afterward.
    /// </summary>
    public static void WithTempFile(byte[] bytes, string filePrefix, string extension, Action<string> action)
    {
        WithTempFile<object?>(bytes, filePrefix, extension, path =>
        {
            action(path);
            return null;
        });
    }

    /// <summary>
    /// Extracts an offline-validated SMILES or KET structure (or built-in/cached chemical name)
    /// from selected document text to preload into the Ketcher 2D editor without network latency.
    /// </summary>
    public static string ExtractPreloadForKetcher(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        var trimmed = text!.Trim();

        if (MoleculeRenderer.TryRenderStructure(trimmed, RenderOptions.Default, out _, out var ket, out var smiles))
        {
            if (!string.IsNullOrWhiteSpace(ket)) return ket!;
            if (!string.IsNullOrWhiteSpace(smiles)) return smiles!;
        }

        return PubChemResolver.ResolveNameToSmilesOffline(trimmed) ?? "";
    }
}
