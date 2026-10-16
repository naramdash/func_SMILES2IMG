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

    /// <summary>
    /// Wrap standard HTML in Windows CF_HTML (Clipboard Format HTML) with byte-accurate UTF-8 offsets.
    /// </summary>
    public static string WrapInCfHtml(string htmlSnippet)
    {
        const string headerTemplate =
            "Version:0.9\r\n" +
            "StartHTML:{0:D8}\r\n" +
            "EndHTML:{1:D8}\r\n" +
            "StartFragment:{2:D8}\r\n" +
            "EndFragment:{3:D8}\r\n";

        const string preFragment = "<!DOCTYPE html><html><body><!--StartFragment-->";
        const string postFragment = "<!--EndFragment--></body></html>";

        var utf8 = System.Text.Encoding.UTF8;
        int dummyHeaderLen = utf8.GetByteCount(string.Format(System.Globalization.CultureInfo.InvariantCulture, headerTemplate, 0, 0, 0, 0));
        int startHtml = dummyHeaderLen;
        int startFragment = startHtml + utf8.GetByteCount(preFragment);
        int endFragment = startFragment + utf8.GetByteCount(htmlSnippet);
        int endHtml = endFragment + utf8.GetByteCount(postFragment);

        string header = string.Format(System.Globalization.CultureInfo.InvariantCulture, headerTemplate, startHtml, endHtml, startFragment, endFragment);
        return header + preFragment + htmlSnippet + postFragment;
    }

    /// <summary>
    /// Sets both Plain Text and rich HTML table on the clipboard with retry resilience.
    /// Pasting in Word/PowerPoint renders a rich formatted table, while Notepad/Excel gets clean text.
    /// </summary>
    public static bool SetTextAndHtml(string? plainText, string? htmlSnippet = null, int maxRetries = 5, int delayMs = 40)
    {
        if (string.IsNullOrEmpty(plainText)) return false;

        if (string.IsNullOrWhiteSpace(htmlSnippet))
            return SetText(plainText, maxRetries, delayMs);

        var dataObject = new DataObject();
        dataObject.SetData(DataFormats.UnicodeText, plainText);
        dataObject.SetData(DataFormats.Text, plainText);
        try
        {
            dataObject.SetData(DataFormats.Html, WrapInCfHtml(htmlSnippet!));
        }
        catch { }

        for (int i = 0; i < maxRetries; i++)
        {
            try
            {
                Clipboard.SetDataObject(dataObject, true, 5, 40);
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

        return SetText(plainText, maxRetries, delayMs);
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
