using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;

namespace Smiles2Img.Editor;

/// <summary>
/// Shared offline web bundle extractor and SHA-256 version fingerprint verifier
/// for embedded WebView2 assets (EPAM Ketcher 2D and Mol* 3D).
/// </summary>
public static class EmbeddedAssetExtractor
{
    public static string EnsureExtracted(ref string? cachedFolder, string devSubfolderName, string cacheSubfolderName, string resourceName)
    {
        if (cachedFolder != null && File.Exists(Path.Combine(cachedFolder, "index.html")))
            return cachedFolder;

        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var devFolder = Path.Combine(baseDir, "assets", devSubfolderName);
        if (File.Exists(Path.Combine(devFolder, "index.html")))
        {
            cachedFolder = devFolder;
            return devFolder;
        }

        var targetDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Smiles2Img", cacheSubfolderName);

        var indexPath = Path.Combine(targetDir, "index.html");
        var manifestPath = Path.Combine(targetDir, ".version");

        var assembly = typeof(EmbeddedAssetExtractor).Assembly;
        var currentVersion = GetBundleFingerprint(resourceName, assembly);

        var needsExtract = !File.Exists(indexPath) ||
                           !File.Exists(manifestPath) ||
                           !string.Equals(File.ReadAllText(manifestPath).Trim(), currentVersion, StringComparison.OrdinalIgnoreCase);

        if (needsExtract)
        {
            Directory.CreateDirectory(targetDir);
            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream != null)
            {
                using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
                foreach (var entry in archive.Entries)
                {
                    var destPath = Path.Combine(targetDir, entry.FullName);
                    var dir = Path.GetDirectoryName(destPath);
                    if (!string.IsNullOrEmpty(dir))
                        Directory.CreateDirectory(dir);

                    if (string.IsNullOrEmpty(entry.Name)) continue;
                    entry.ExtractToFile(destPath, overwrite: true);
                }
                try { File.WriteAllText(manifestPath, currentVersion); } catch { }
            }
        }

        cachedFolder = targetDir;
        return targetDir;
    }

    public static string GetBundleFingerprint(string resourceName, Assembly? assembly = null)
    {
        assembly ??= typeof(EmbeddedAssetExtractor).Assembly;
        var asmVersion = assembly.GetName().Version?.ToString() ?? "1.0.0.0";
        try
        {
            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) return asmVersion;

            var zipLength = stream.Length;
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var indexEntry = archive.GetEntry("index.html");
            if (indexEntry != null)
            {
                using var entryStream = indexEntry.Open();
                using var sha = SHA256.Create();
                var hashBytes = sha.ComputeHash(entryStream);
                var hashHex = BitConverter.ToString(hashBytes).Replace("-", "").Substring(0, 16);
                return $"{asmVersion}-{zipLength}-{hashHex}";
            }
            return $"{asmVersion}-{zipLength}";
        }
        catch
        {
            return asmVersion;
        }
    }
}
