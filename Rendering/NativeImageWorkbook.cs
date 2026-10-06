using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smiles2Img.Rendering;

internal static class NativeImageWorkbook
{
    // An Excel-created single-cell rich-image package. Importing an existing rich
    // value avoids the asynchronous UI conversion performed by PlacePictureInCell.
    internal static void CopyTo(Excel.Range destination, byte[] png)
    {
        var path = Path.Combine(Path.GetTempPath(), "smiles2img-native-" + Guid.NewGuid().ToString("N") + ".xlsx");
        Excel.Workbook? source = null;
        try
        {
            using (var template = Assembly.GetExecutingAssembly().GetManifestResourceStream("Smiles2Img.NativeImageTemplate.xlsx")!)
            using (var file = File.Create(path)) template.CopyTo(file);
            using (var file = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Update))
            {
                archive.GetEntry("xl/media/image1.png")!.Delete();
                using var image = archive.CreateEntry("xl/media/image1.png").Open();
                image.Write(png, 0, png.Length);
            }
            var application = (Excel.Application)ExcelDnaUtil.Application;
            source = application.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true, AddToMru: false);
            var sourceCell = ((Excel.Worksheet)source.Worksheets.Item[1]).Range["A1"];
            sourceCell.Copy(Destination: destination);
            application.CutCopyMode = 0;
        }
        finally
        {
            try { source?.Close(SaveChanges: false); }
            catch { }

            if (source != null)
            {
                try { Marshal.ReleaseComObject(source); }
                catch { }
            }

            if (File.Exists(path))
            {
                try { File.Delete(path); }
                catch { }
            }
        }
    }
}
