using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using Smiles2Img.Tests.E2E;
using Smiles2Img.Tests.Unit;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--capture-dialog")
        {
            NativeIndigo.Initialize(AppDomain.CurrentDomain.BaseDirectory);
            var info = MoleculeRenderer.TryGetMoleculeInfo("C1C=CC=C1", resolvePubChem: true);
            using var dlg = new MoleculeInfoDialog(info!);
            dlg.Show();
            dlg.Refresh();
            Application.DoEvents();
            using var bmp = new Bitmap(dlg.Width, dlg.Height);
            dlg.DrawToBitmap(bmp, new Rectangle(0, 0, dlg.Width, dlg.Height));
            string outPath = args.Length > 1 ? args[1] : "dialog_capture.png";
            bmp.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);
            Console.WriteLine("SAVED:" + outPath);
            return 0;
        }

        bool runUnit = true;
        bool runWebViewE2E = true;
        bool runOfficeLiveE2E = false;

        if (args.Any(a => string.Equals(a, "--unit", StringComparison.OrdinalIgnoreCase)))
        {
            runUnit = true;
            runWebViewE2E = false;
            runOfficeLiveE2E = false;
        }
        else if (args.Any(a => string.Equals(a, "--e2e", StringComparison.OrdinalIgnoreCase)))
        {
            runUnit = false;
            runWebViewE2E = true;
            runOfficeLiveE2E = false;
        }
        else if (args.Any(a => string.Equals(a, "--e2e-office", StringComparison.OrdinalIgnoreCase)))
        {
            runUnit = false;
            runWebViewE2E = false;
            runOfficeLiveE2E = true;
        }
        else if (args.Any(a => string.Equals(a, "--full", StringComparison.OrdinalIgnoreCase)))
        {
            runUnit = true;
            runWebViewE2E = true;
            runOfficeLiveE2E = true;
        }

        try
        {
            ResetCounters();
            NativeIndigo.Initialize(AppDomain.CurrentDomain.BaseDirectory);
            Console.WriteLine($"func_SMILES2IMG Test Runner ({(Environment.Is64BitProcess ? "x64" : "x86")})");

            if (runUnit)
            {
                PrintCategoryHeader("CATEGORY 1: UNIT TESTS (Rendering, Cheminformatics, Real Substances, UI & Ribbon XML)");
                RenderingUnitTests.RunAll();
                CheminformaticsUnitTests.RunAll();
                RealSubstanceEdgeCaseTests.RunAll();
                UiAndRibbonUnitTests.RunAll();
            }

            if (runWebViewE2E || runOfficeLiveE2E)
            {
                PrintCategoryHeader("CATEGORY 2: END-TO-END (E2E) TESTS (WebView2 Editors & Live Office COM/XLL)");
                if (runWebViewE2E)
                {
                    WebView2EditorE2ETests.RunAll();
                }
                if (runOfficeLiveE2E)
                {
                    OfficeLiveE2ERunner.RunLiveOfficeSuite();
                }
            }

            Console.WriteLine();
            Console.WriteLine($"[ALL {PassedSuites} TEST SUITES PASSED SUCCESSFULLY]");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    internal static readonly HashSet<string> KnownInvalidImageMsos = new(StringComparer.OrdinalIgnoreCase)
    {
        "FormatBullet", "Key", "FontName", "ShapesGallery", "HighContrast",
        "TableCustom", "DocumentSummary", "SourceCodeInsert", "DrawingCanvasInsert", "ShapeFillColorPicker",
        "Tag", "InvertColors", "Dark", "NightMode"
    };

    internal static readonly HashSet<string> VerifiedValidOfficeImageMsos = new(StringComparer.OrdinalIgnoreCase)
    {
        "PensGallery",
        "TextBoxInsert",
        "SmartArtInsert",
        "Insert3DModelFromFile",
        "Paste",
        "Copy",
        "SignatureLineInsert",
        "CopyPicture",
        "FileProperties",
        "FontSize",
        "Subscript",
        "Calculator",
        "FilePermission",
        "FileDocumentInspect",
        "TableInsert",
        "Bullets",
        "TableOfContentsGallery",
        "PictureRecolorGallery",
        "ColorMenu",
        "BlackAndWhiteGrayscale",
        "BlackAndWhiteLightGrayscale",
        "BlackAndWhiteBlack",
        "DatabaseQueryNew",
        "ImportTextFile",
        "FileSaveAs",
        "WatchWindow",
        "Info",
        "DesignMode",
        "EquationInsertNew",
        "FunctionWizard",
        "GroupTaskProperties",
        "PropertySheet"
    };

    internal static void ValidateRibbonXml(string xml, Type addInType, string appName)
    {
        Check(!string.IsNullOrWhiteSpace(xml), $"{appName} Ribbon XML is non-empty");
        var doc = XDocument.Parse(xml);
        Check(doc.Root != null, $"{appName} XML has root element");

        var onActions = doc.Descendants()
            .Select(e => (string?)e.Attribute("onAction"))
            .Where(a => !string.IsNullOrEmpty(a))
            .Distinct()
            .ToList();

        foreach (var action in onActions)
        {
            var method = addInType.GetMethod(action!, BindingFlags.Public | BindingFlags.Instance);
            Check(method != null, $"{appName} onAction callback '{action}' must exist on {addInType.Name}");
        }

        var allIds = doc.Descendants()
            .Select(e => (string?)e.Attribute("id"))
            .Where(id => !string.IsNullOrEmpty(id))
            .ToList();
        var duplicates = allIds.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        Check(duplicates.Count == 0, $"{appName} duplicate control IDs: {string.Join(", ", duplicates)}");

        var imageMsos = doc.Descendants()
            .Select(e => (string?)e.Attribute("imageMso"))
            .Where(img => !string.IsNullOrEmpty(img))
            .ToList();
        foreach (var img in imageMsos)
        {
            Check(!KnownInvalidImageMsos.Contains(img!), $"{appName} uses forbidden/invalid imageMso '{img}'");
            Check(VerifiedValidOfficeImageMsos.Contains(img!), $"{appName} uses unverified imageMso '{img}'. Please verify it is a valid Office button icon and add to VerifiedValidOfficeImageMsos in tests/Program.cs and AGENTS.md.");
        }
    }
}
