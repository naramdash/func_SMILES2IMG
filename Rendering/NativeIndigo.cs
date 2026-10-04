using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Smiles2Img.Rendering;

internal static class NativeIndigo
{
    private const uint SearchDllDirectory = 0x00000100;
    private const uint SearchDefaultDirectories = 0x00001000;

    internal static void Initialize(string fallbackDirectory)
    {
        var targetDir = GetOrExtractNativeDirectory(fallbackDirectory);

        var dllsToLoad = Environment.Is64BitProcess
            ? new[] { "vcruntime140.dll", "vcruntime140_1.dll", "msvcp140.dll", "indigo.dll", "indigo-renderer.dll" }
            : new[] { "vcruntime140.dll", "msvcp140.dll", "indigo.dll", "indigo-renderer.dll" };

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

    private static string GetOrExtractNativeDirectory(string fallbackDirectory)
    {
        if (!string.IsNullOrEmpty(fallbackDirectory) &&
            File.Exists(Path.Combine(fallbackDirectory, "indigo.dll")) &&
            File.Exists(Path.Combine(fallbackDirectory, "indigo-renderer.dll")))
        {
            return fallbackDirectory;
        }

        var bitness = Environment.Is64BitProcess ? "x64" : "x86";
        var cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Smiles2Img", "native", bitness);

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
}
