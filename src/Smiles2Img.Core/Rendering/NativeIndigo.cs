using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Smiles2Img.Rendering;

public static class NativeIndigo
{
    private const uint SearchDllDirectory = 0x00000100;
    private const uint SearchDefaultDirectories = 0x00001000;

    public static string? NativeDirectory { get; private set; }

    private static readonly object initLock = new();

    public static void EnsureInitialized(string? fallbackDirectory = null)
    {
        if (!string.IsNullOrEmpty(NativeDirectory)) return;
        lock (initLock)
        {
            if (!string.IsNullOrEmpty(NativeDirectory)) return;
            var dir = fallbackDirectory;
            if (string.IsNullOrEmpty(dir))
            {
                try { dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); } catch { }
            }
            Initialize(dir ?? "");
        }
    }

    public static void Initialize(string fallbackDirectory)
    {
        if (!string.IsNullOrEmpty(NativeDirectory)) return;
        lock (initLock)
        {
            if (!string.IsNullOrEmpty(NativeDirectory)) return;
            var targetDir = GetOrExtractNativeDirectory(fallbackDirectory);
            NativeDirectory = targetDir;
            try { SetDllDirectory(targetDir); } catch { }
            EnsureWebView2Loader();

            var dllsToLoad = Environment.Is64BitProcess
                ? new[] { "vcruntime140.dll", "vcruntime140_1.dll", "msvcp140.dll", "indigo.dll", "indigo-renderer.dll", "indigo-inchi.dll" }
                : new[] { "vcruntime140.dll", "msvcp140.dll", "indigo.dll", "indigo-renderer.dll", "indigo-inchi.dll" };

            foreach (var name in dllsToLoad)
            {
                var path = Path.Combine(targetDir, name);
                if (File.Exists(path))
                {
                    if (LoadLibraryEx(path, IntPtr.Zero, SearchDllDirectory | SearchDefaultDirectories) == IntPtr.Zero)
                    {
                        if (LoadLibrary(path) == IntPtr.Zero)
                        {
                            var error = Marshal.GetLastWin32Error();
                            if (name.StartsWith("indigo"))
                                throw new Win32Exception(error, "Cannot load Indigo native library: " + path);
                        }
                    }
                }
            }
        }
    }

    public static void EnsureWebView2Loader()
    {
        if (!string.IsNullOrEmpty(NativeDirectory))
        {
            try
            {
                Microsoft.Web.WebView2.Core.CoreWebView2Environment.SetLoaderDllFolderPath(NativeDirectory);
            }
            catch { }
        }
    }

    private static string GetOrExtractNativeDirectory(string fallbackDirectory)
    {
        if (!string.IsNullOrEmpty(fallbackDirectory) &&
            File.Exists(Path.Combine(fallbackDirectory, "indigo.dll")) &&
            File.Exists(Path.Combine(fallbackDirectory, "indigo-renderer.dll")))
        {
            return fallbackDirectory;
        }

        if (!string.IsNullOrEmpty(fallbackDirectory))
        {
            var suiteDir = Path.Combine(fallbackDirectory, "Smiles2Img");
            if (File.Exists(Path.Combine(suiteDir, "indigo.dll")) &&
                File.Exists(Path.Combine(suiteDir, "indigo-renderer.dll")))
            {
                return suiteDir;
            }
        }

        var bitness = Environment.Is64BitProcess ? "x64" : "x86";
        var cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Smiles2Img", "native", bitness);

        if (Directory.Exists(cacheDir) &&
            File.Exists(Path.Combine(cacheDir, "indigo.dll")) &&
            File.Exists(Path.Combine(cacheDir, "indigo-renderer.dll")) &&
            File.Exists(Path.Combine(cacheDir, "indigo-inchi.dll")))
        {
            return cacheDir;
        }

        Directory.CreateDirectory(cacheDir);

        var assembly = typeof(NativeIndigo).Assembly;
        var prefix = "Smiles2Img.Native." + bitness + ".";

        foreach (var resourceName in assembly.GetManifestResourceNames())
        {
            if (!resourceName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;

            var fileName = resourceName.Substring(prefix.Length);
            var destPath = Path.Combine(cacheDir, fileName);

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) continue;

            if (!File.Exists(destPath) || new FileInfo(destPath).Length != stream.Length)
            {
                try
                {
                    using var fileStream = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None);
                    stream.CopyTo(fileStream);
                }
                catch (IOException) { }
            }
        }

        return cacheDir;
    }

    [DllImport("kernel32.dll", EntryPoint = "LoadLibraryExW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryEx(string fileName, IntPtr file, uint flags);

    [DllImport("kernel32.dll", EntryPoint = "LoadLibraryW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibrary(string fileName);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool SetDllDirectory(string lpPathName);
}
