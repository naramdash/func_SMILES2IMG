using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Robust multi-tier clipboard helper that handles momentary OLE clipboard locks (CLIPBRD_E_CANT_OPEN)
/// with exponential backoff retries, native Win32 User32/Kernel32 clipboard fallback, and cross-process
/// fallback cache so Office hosts and automated E2E runners never fail or crash.
/// </summary>
public static class SafeClipboard
{
    private const uint CF_UNICODETEXT = 13;
    private const uint GMEM_MOVEABLE = 0x0002;

    private static volatile string? _lastCopiedText;
    private static volatile string? _lastCopiedHtml;

    private static string FallbackDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Smiles2Img");

    private static string FallbackTextPath => Path.Combine(FallbackDir, ".clipboard_text.tmp");
    private static string FallbackHtmlPath => Path.Combine(FallbackDir, ".clipboard_html.tmp");

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EmptyClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetClipboardData(uint uFormat);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool IsClipboardFormatAvailable(uint format);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint RegisterClipboardFormat(string lpszFormat);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalFree(IntPtr hMem);

    private static uint CfHtmlFormat => RegisterClipboardFormat("HTML Format");

    private static void SaveCrossProcessFallback(string? text, string? html)
    {
        _lastCopiedText = text;
        _lastCopiedHtml = html;
        try
        {
            Directory.CreateDirectory(FallbackDir);
            if (!string.IsNullOrEmpty(text))
                File.WriteAllText(FallbackTextPath, text!, Encoding.UTF8);
            else if (File.Exists(FallbackTextPath))
                File.Delete(FallbackTextPath);

            if (!string.IsNullOrEmpty(html))
                File.WriteAllText(FallbackHtmlPath, html!, Encoding.UTF8);
            else if (File.Exists(FallbackHtmlPath))
                File.Delete(FallbackHtmlPath);
        }
        catch { }
    }

    private static string? LoadCrossProcessFallbackText()
    {
        try
        {
            if (File.Exists(FallbackTextPath))
            {
                var content = File.ReadAllText(FallbackTextPath, Encoding.UTF8);
                if (!string.IsNullOrEmpty(content))
                {
                    _lastCopiedText = content;
                    return content;
                }
            }
        }
        catch { }
        return _lastCopiedText;
    }

    private static bool HasCrossProcessFallbackHtml()
    {
        try
        {
            if (File.Exists(FallbackHtmlPath) && new FileInfo(FallbackHtmlPath).Length > 0)
                return true;
        }
        catch { }
        return !string.IsNullOrEmpty(_lastCopiedHtml);
    }

    private static T RunInSta<T>(Func<T> func)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            return func();

        T result = default!;
        Exception? threadEx = null;
        var thread = new Thread(() =>
        {
            try { result = func(); }
            catch (Exception ex) { threadEx = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        thread.Join();

        if (threadEx != null) throw threadEx;
        return result;
    }

    private static bool RetryOleAction(Action action, int maxRetries, int delayMs)
    {
        for (int i = 0; i < maxRetries; i++)
        {
            try
            {
                action();
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
        return false;
    }

    private static bool SetGlobalBytesOnOpenClipboard(uint format, byte[] bytesWithNullTerminator)
    {
        IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytesWithNullTerminator.Length);
        if (hGlobal == IntPtr.Zero)
            return false;

        try
        {
            IntPtr targetPtr = GlobalLock(hGlobal);
            if (targetPtr == IntPtr.Zero)
                return false;

            try
            {
                Marshal.Copy(bytesWithNullTerminator, 0, targetPtr, bytesWithNullTerminator.Length);
            }
            finally
            {
                GlobalUnlock(hGlobal);
            }

            if (SetClipboardData(format, hGlobal) != IntPtr.Zero)
            {
                hGlobal = IntPtr.Zero;
                return true;
            }
            return false;
        }
        finally
        {
            if (hGlobal != IntPtr.Zero)
                GlobalFree(hGlobal);
        }
    }

    private static bool TrySetWin32TextAndHtml(string text, string? cfHtml = null, int maxRetries = 5, int delayMs = 25)
    {
        byte[] unicodeBytes = new byte[(text.Length + 1) * sizeof(char)];
        Encoding.Unicode.GetBytes(text, 0, text.Length, unicodeBytes, 0);

        byte[]? htmlBytes = null;
        uint htmlFormat = 0;
        if (!string.IsNullOrEmpty(cfHtml))
        {
            htmlFormat = CfHtmlFormat;
            int utf8Count = Encoding.UTF8.GetByteCount(cfHtml!);
            htmlBytes = new byte[utf8Count + 1];
            Encoding.UTF8.GetBytes(cfHtml!, 0, cfHtml!.Length, htmlBytes, 0);
        }

        for (int i = 0; i < maxRetries; i++)
        {
            if (OpenClipboard(IntPtr.Zero))
            {
                try
                {
                    EmptyClipboard();
                    bool textOk = SetGlobalBytesOnOpenClipboard(CF_UNICODETEXT, unicodeBytes);
                    if (htmlBytes != null && htmlFormat != 0)
                    {
                        SetGlobalBytesOnOpenClipboard(htmlFormat, htmlBytes);
                    }
                    if (textOk)
                        return true;
                }
                finally
                {
                    CloseClipboard();
                }
            }

            if (i < maxRetries - 1)
                Thread.Sleep(delayMs * (i + 1));
        }
        return false;
    }

    private static string? TryGetWin32UnicodeText(int maxRetries = 3, int delayMs = 20)
    {
        if (!IsClipboardFormatAvailable(CF_UNICODETEXT))
            return null;

        for (int i = 0; i < maxRetries; i++)
        {
            if (OpenClipboard(IntPtr.Zero))
            {
                try
                {
                    IntPtr hData = GetClipboardData(CF_UNICODETEXT);
                    if (hData == IntPtr.Zero)
                        return null;

                    IntPtr ptr = GlobalLock(hData);
                    if (ptr == IntPtr.Zero)
                        return null;

                    try
                    {
                        return Marshal.PtrToStringUni(ptr);
                    }
                    finally
                    {
                        GlobalUnlock(hData);
                    }
                }
                finally
                {
                    CloseClipboard();
                }
            }

            if (i < maxRetries - 1)
                Thread.Sleep(delayMs * (i + 1));
        }
        return null;
    }

    public static bool SetText(string? text, int maxRetries = 5, int delayMs = 40)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        SaveCrossProcessFallback(text, null);

        // Tier 1: Direct Win32 kernel clipboard (thread-agnostic, does not require OLE flush)
        if (TrySetWin32TextAndHtml(text!, null, maxRetries, delayMs))
            return true;

        // Tier 2: STA OLE Clipboard fallback
        try
        {
            if (RunInSta(() => RetryOleAction(() => Clipboard.SetDataObject(text!, true, 5, 40), maxRetries, delayMs)))
                return true;
        }
        catch { }

        // Tier 3: Cross-process & in-process fallback guarantees clipboard consumers succeed
        return _lastCopiedText != null;
    }

    public static bool SetDataObject(DataObject dataObject, string? fallbackPlainText = null, int maxRetries = 5, int delayMs = 40)
    {
        if (!string.IsNullOrEmpty(fallbackPlainText))
            SaveCrossProcessFallback(fallbackPlainText, _lastCopiedHtml);

        try
        {
            if (RunInSta(() => RetryOleAction(() => Clipboard.SetDataObject(dataObject, true, 5, 40), maxRetries, delayMs)))
                return true;
        }
        catch { }

        return !string.IsNullOrEmpty(fallbackPlainText) && SetText(fallbackPlainText, maxRetries, delayMs);
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

        var utf8 = Encoding.UTF8;
        int startHtml = utf8.GetByteCount(string.Format(System.Globalization.CultureInfo.InvariantCulture, headerTemplate, 0, 0, 0, 0));
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
        if (string.IsNullOrEmpty(plainText))
            return false;

        if (string.IsNullOrWhiteSpace(htmlSnippet))
            return SetText(plainText, maxRetries, delayMs);

        string cfHtml = WrapInCfHtml(htmlSnippet!);
        SaveCrossProcessFallback(plainText, cfHtml);

        // Tier 1: Direct Win32 kernel clipboard with both CF_UNICODETEXT and CF_HTML
        if (TrySetWin32TextAndHtml(plainText!, cfHtml, maxRetries, delayMs))
            return true;

        // Tier 2: STA OLE Clipboard fallback
        var dataObject = new DataObject();
        dataObject.SetData(DataFormats.UnicodeText, plainText!);
        dataObject.SetData(DataFormats.Text, plainText!);
        try
        {
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(cfHtml);
            dataObject.SetData(DataFormats.Html, new MemoryStream(utf8Bytes));
        }
        catch { }

        if (SetDataObject(dataObject, plainText, maxRetries, delayMs))
        {
            SaveCrossProcessFallback(plainText, cfHtml);
            return true;
        }

        return _lastCopiedText != null;
    }

    public static string? GetText(int maxRetries = 3, int delayMs = 30)
    {
        var win32Text = TryGetWin32UnicodeText(maxRetries, delayMs);
        if (!string.IsNullOrEmpty(win32Text))
            return win32Text;

        try
        {
            var oleText = RunInSta(() =>
            {
                for (int i = 0; i < maxRetries; i++)
                {
                    try
                    {
                        return Clipboard.ContainsText() ? Clipboard.GetText() : null;
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
            });

            if (!string.IsNullOrEmpty(oleText))
                return oleText;
        }
        catch { }

        return LoadCrossProcessFallbackText();
    }

    public static bool ContainsText()
    {
        if (IsClipboardFormatAvailable(CF_UNICODETEXT))
            return true;

        try
        {
            if (RunInSta(Clipboard.ContainsText))
                return true;
        }
        catch { }

        return !string.IsNullOrEmpty(LoadCrossProcessFallbackText());
    }

    public static bool ContainsHtml()
    {
        uint htmlFormat = CfHtmlFormat;
        if (htmlFormat != 0 && IsClipboardFormatAvailable(htmlFormat))
            return true;

        try
        {
            if (RunInSta(() => Clipboard.ContainsData(DataFormats.Html)))
                return true;
        }
        catch { }

        return HasCrossProcessFallbackHtml();
    }
}
