using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Smiles2Img.Rendering;

public static class SvgClipboardHelper
{
    /// <summary>
    /// Copies an SVG XML string to the Windows Clipboard in both vector graphic stream format
    /// ("image/svg+xml") for Office 365 / Illustrator / Inkscape / Figma and Unicode text format
    /// for code/XML editors.
    /// </summary>
    public static bool CopySvgToClipboard(string? svgXml)
    {
        if (string.IsNullOrWhiteSpace(svgXml))
            return false;

        try
        {
            var utf8Bytes = Encoding.UTF8.GetBytes(svgXml);
            var dataObj = new DataObject();

            // Standard MIME clipboard format recognized by Office 365, Inkscape, Illustrator, Chromium, and Figma
            dataObj.SetData("image/svg+xml", false, new MemoryStream(utf8Bytes));
            dataObj.SetData("Scalable Vector Graphics", false, new MemoryStream(utf8Bytes));
            dataObj.SetData(DataFormats.UnicodeText, true, svgXml);
            dataObj.SetData(DataFormats.Text, true, svgXml);

            Clipboard.SetDataObject(dataObj, true);
            return true;
        }
        catch
        {
            try
            {
                Clipboard.SetText(svgXml);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
