using System;
using System.IO;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

namespace Smiles2Img;

public sealed class AddIn : IExcelAddIn
{
    internal static bool IsOpen { get; private set; }

    public void AutoOpen()
    {
        try
        {
            NativeIndigo.Initialize(Path.GetDirectoryName(ExcelDnaUtil.XllPath)!);
            IntelliSenseServer.Install();
            IsOpen = true;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { FloatingPictures.Attach(); }
                catch (Exception error) { LogDisplay.WriteLine("SMILES2IMG.FLOAT events: " + error.Message); }
            });
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("SMILES2IMG startup: " + error.Message);
        }
    }

    public void AutoClose()
    {
        try
        {
            FloatingPictures.Detach();
        }
        catch { }
        try
        {
            IntelliSenseServer.Uninstall();
        }
        catch { }
        IsOpen = false;
    }
}
