using System;
using System.IO;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

namespace Smiles2Img;

public sealed class AddIn : IExcelAddIn
{
    internal static volatile bool IsOpen;

    public void AutoOpen()
    {
        try
        {
            IsOpen = true;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { IntelliSenseServer.Install(); }
                catch (Exception error) { LogDisplay.WriteLine("IntelliSense install: " + error.Message); }
                try { FloatingPictures.Attach(); }
                catch (Exception error) { LogDisplay.WriteLine("SMILES2IMG.FLOAT events: " + error.Message); }
                try { KetcherEditor.Attach(); }
                catch (Exception error) { LogDisplay.WriteLine("Ketcher events: " + error.Message); }
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
            KetcherEditor.Detach();
        }
        catch { }
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
