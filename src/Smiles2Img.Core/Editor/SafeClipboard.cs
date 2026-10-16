using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Robust clipboard helper that handles momentary clipboard locks (CLIPBRD_E_CANT_OPEN)
/// with exponential backoff retries, preventing Office application host crashes.
/// </summary>
public static class SafeClipboard
{
    public static bool SetText(string? text, int maxRetries = 5, int delayMs = 40)
    {
        if (string.IsNullOrEmpty(text)) return false;

        for (int i = 0; i < maxRetries; i++)
        {
            try
            {
                Clipboard.SetDataObject(text!, true, 5, 40);
                return true;
            }
            catch (ExternalException)
            {
                if (i < maxRetries - 1)
                    Thread.Sleep(delayMs * (i + 1));
            }
            catch
            {
                break;
            }
        }

        try
        {
            Clipboard.SetText(text!);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static string? GetText(int maxRetries = 3, int delayMs = 30)
    {
        for (int i = 0; i < maxRetries; i++)
        {
            try
            {
                if (Clipboard.ContainsText())
                {
                    return Clipboard.GetText();
                }
                return null;
            }
            catch (ExternalException)
            {
                if (i < maxRetries - 1)
                    Thread.Sleep(delayMs * (i + 1));
            }
            catch
            {
                break;
            }
        }
        return null;
    }

    public static bool ContainsText()
    {
        try
        {
            return Clipboard.ContainsText();
        }
        catch
        {
            return false;
        }
    }
}
