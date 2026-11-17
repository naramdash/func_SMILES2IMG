using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Smiles2Img.Rendering;

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

/// <summary>
/// Shared WebView2 host initialization, DPI window sizing, viewport preview capture,
/// and JS string escaping helpers used by KetcherForm and MolstarForm.
/// </summary>
internal static class WebView2HostHelper
{
    public static void ApplyDpiScaledWindowBounds(Form form, int targetBaseW, int targetBaseH, int minClientW, int minClientH)
    {
        float dpiScale = DialogHelper.GetDpiScale(form);
        var screen = Screen.FromControl(form) ?? Screen.PrimaryScreen;
        int targetW = Math.Min((int)(targetBaseW * dpiScale), (int)(screen.WorkingArea.Width * 0.94));
        int targetH = Math.Min((int)(targetBaseH * dpiScale), (int)(screen.WorkingArea.Height * 0.94));
        form.ClientSize = new Size(Math.Max(targetW, minClientW), Math.Max(targetH, minClientH));
        form.MinimumSize = new Size(Math.Min(880, screen.WorkingArea.Width), Math.Min(600, screen.WorkingArea.Height));
        form.StartPosition = FormStartPosition.CenterParent;
    }

    public static async Task<CoreWebView2Environment> InitializeWebViewCoreAsync(
        WebView2 webView,
        CoreWebView2Environment? sharedEnv,
        string procFolderSuffix,
        string tempPrefix,
        string virtualHostName,
        string assetFolder)
    {
        NativeIndigo.EnsureWebView2Loader();

        var baseUserDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Smiles2Img", "WebView2UserData");
        Directory.CreateDirectory(baseUserDataFolder);

        // Isolate per host process (Excel / Word / PowerPoint) to prevent lock contention
        string procName = Process.GetCurrentProcess().ProcessName + procFolderSuffix;
        var procUserDataFolder = Path.Combine(baseUserDataFolder, procName);
        Directory.CreateDirectory(procUserDataFolder);

        try
        {
            sharedEnv ??= await CoreWebView2Environment.CreateAsync(null, procUserDataFolder, null);
            await webView.EnsureCoreWebView2Async(sharedEnv);
        }
        catch (Exception ex) when (ex.HResult == unchecked((int)0x800700AA) || ex.Message.Contains("0x800700AA") || ex.Message.IndexOf("in use", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            // Fallback to a dedicated unique session directory if locked by an orphaned browser process
            string fallbackFolder = Path.Combine(Path.GetTempPath(), tempPrefix + procName + "_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(fallbackFolder);
            sharedEnv = await CoreWebView2Environment.CreateAsync(null, fallbackFolder, null);
            await webView.EnsureCoreWebView2Async(sharedEnv);
        }

        webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
        webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;

        webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
            virtualHostName,
            assetFolder,
            CoreWebView2HostResourceAccessKind.Allow);

        return sharedEnv;
    }

    public static async Task<byte[]?> TryCaptureViewportPngAsync(WebView2 webView)
    {
        if (webView.CoreWebView2 == null) return null;
        try
        {
            using var uiMs = new MemoryStream();
            await webView.CoreWebView2.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, uiMs);
            return uiMs.ToArray();
        }
        catch
        {
            return null;
        }
    }

    public static string EscapeJs(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s!.Replace("\\", "\\\\")
                .Replace("'", "\\'")
                .Replace("\"", "\\\"")
                .Replace("\r", "")
                .Replace("\n", "\\n");
    }
}
