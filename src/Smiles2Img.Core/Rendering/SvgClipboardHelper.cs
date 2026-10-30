using System.IO;
using System.Text;
using System.Windows.Forms;
using Smiles2Img.Editor;

namespace Smiles2Img.Rendering;

public static class SvgClipboardHelper
{
    /// <summary>
    /// Copies an SVG XML string to the Windows Clipboard in both vector graphic stream format
    /// ("image/svg+xml") for Office 365 / Illustrator / Inkscape / Figma and Unicode text format
    /// for code/XML editors, with full STA and Win32 fallback protection.
    /// </summary>
    public static bool CopySvgToClipboard(string? svgXml)
    {
        if (string.IsNullOrWhiteSpace(svgXml))
            return false;

        try
        {
            var utf8Bytes = Encoding.UTF8.GetBytes(svgXml!);
            var dataObj = new DataObject();

            // Standard MIME clipboard format recognized by Office 365, Inkscape, Illustrator, Chromium, and Figma
            dataObj.SetData("image/svg+xml", false, new MemoryStream(utf8Bytes));
            dataObj.SetData("Scalable Vector Graphics", false, new MemoryStream(utf8Bytes));
            dataObj.SetData(DataFormats.UnicodeText, true, svgXml!);
            dataObj.SetData(DataFormats.Text, true, svgXml!);

            return SafeClipboard.SetDataObject(dataObj, svgXml!);
        }
        catch
        {
            return SafeClipboard.SetText(svgXml);
        }
    }
}
