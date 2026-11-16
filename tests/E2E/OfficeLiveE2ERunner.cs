using System;
using System.Diagnostics;
using System.IO;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests.E2E;

internal static class OfficeLiveE2ERunner
{
    public static void RunLiveOfficeSuite()
    {
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var repoRoot = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
        var scriptPath = Path.Combine(repoRoot, "scripts", "Test-OfficeE2E.ps1");
        if (!File.Exists(scriptPath))
        {
            scriptPath = Path.GetFullPath(Path.Combine("scripts", "Test-OfficeE2E.ps1"));
        }
        Check(File.Exists(scriptPath), $"Live Office E2E script not found at {scriptPath}");

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var proc = Process.Start(psi)!;
        proc.OutputDataReceived += (_, e) => { if (e.Data != null) Console.WriteLine("    " + e.Data); };
        proc.ErrorDataReceived += (_, e) => { if (e.Data != null) Console.Error.WriteLine("    " + e.Data); };
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();
        proc.WaitForExit();

        Check(proc.ExitCode == 0, $"Live Office COM & XLL E2E suite failed with exit code {proc.ExitCode}");
        PassSuite("E2E Live Office Suite (Word In-Process COM Add-in + PowerPoint In-Process COM Add-in + Excel Live .xll)");
    }
}
