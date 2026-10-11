using System;
using System.IO;
using System.IO.Compression;

namespace Smiles2Img;

internal static class KetcherAssetManager
{
    private static string? cachedFolder;

    internal static string EnsureKetcherFolder()
    {
        if (cachedFolder != null && File.Exists(Path.Combine(cachedFolder, "index.html")))
            return cachedFolder;

        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var devFolder = Path.Combine(baseDir, "assets", "ketcher_bundle");
        if (File.Exists(Path.Combine(devFolder, "index.html")))
        {
            cachedFolder = devFolder;
            return devFolder;
        }

        var targetDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Smiles2Img", "ketcher");

        var indexPath = Path.Combine(targetDir, "index.html");
        var manifestPath = Path.Combine(targetDir, ".version");
        var currentVersion = typeof(KetcherAssetManager).Assembly.GetName().Version?.ToString() ?? "1.0.0.0";

        var needsExtract = !File.Exists(indexPath) || !File.Exists(manifestPath) || File.ReadAllText(manifestPath).Trim() != currentVersion;

        if (needsExtract)
        {
            Directory.CreateDirectory(targetDir);
            var assembly = typeof(KetcherAssetManager).Assembly;
            using var stream = assembly.GetManifestResourceStream("Smiles2Img.Ketcher.zip");
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
}

