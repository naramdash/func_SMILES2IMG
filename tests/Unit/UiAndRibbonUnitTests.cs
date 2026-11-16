using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Smiles2Img;
using Smiles2Img.Editor;
using Smiles2Img.PowerPoint;
using Smiles2Img.Rendering;
using Smiles2Img.Word;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests.Unit;

internal static class UiAndRibbonUnitTests
{
    public static void RunAll()
    {
        RunOfflineAssetExtractionTests();
        RunRibbonAndContextMenuXmlTests();
        RunWinFormsDialogsIconsAndClipboardTests();
        RunSemVerComparisonTests();
    }

    private static void RunOfflineAssetExtractionTests()
    {
        Check(KetcherEditor.IsLikelySmiles("c1ccccc1"), "Ketcher detects benzene SMILES");
        Check(KetcherEditor.IsLikelySmiles("CC(=O)Oc1ccccc1C(=O)O"), "Ketcher detects aspirin SMILES");
        Check(KetcherEditor.IsLikelySmiles("CCO"), "Ketcher detects CCO SMILES");
        Check(!KetcherEditor.IsLikelySmiles("12345"), "Ketcher rejects plain numbers");
        Check(!KetcherEditor.IsLikelySmiles("Hello world"), "Ketcher rejects plain text");
        Check(!KetcherEditor.IsLikelySmiles(""), "Ketcher rejects empty string");

        var fingerprint = KetcherAssetManager.GetBundleFingerprint();
        Check(!string.IsNullOrWhiteSpace(fingerprint) && fingerprint.Contains("-"), "Ketcher bundle fingerprint is generated");
        var ketcherFolder = KetcherAssetManager.EnsureKetcherFolder();
        Check(Directory.Exists(ketcherFolder), "Ketcher folder exists");
        Check(File.Exists(Path.Combine(ketcherFolder, "index.html")), "Ketcher index.html exists");
        var indexContent = File.ReadAllText(Path.Combine(ketcherFolder, "index.html"));
        Check(indexContent.Contains("getCanvasSvg"), "Ketcher index.html includes getCanvasSvg");
        Check(indexContent.Contains("calculateCanvasBBox"), "Ketcher index.html includes calculateCanvasBBox");
        Check(indexContent.Contains("toggleBackground"), "Ketcher index.html includes toggleBackground");

        var molstarFp = MolstarAssetManager.GetBundleFingerprint();
        Check(!string.IsNullOrWhiteSpace(molstarFp) && molstarFp.Contains("-"), "Molstar bundle fingerprint is generated");
        var molstarFolder = MolstarAssetManager.EnsureMolstarFolder();
        Check(Directory.Exists(molstarFolder) && File.Exists(Path.Combine(molstarFolder, "index.html")), "Molstar offline bundle extracted cleanly");

        PassSuite("Unit: Ketcher & Mol* Offline Bundle Extraction & SHA-256 Fingerprints");
    }

    private static void RunRibbonAndContextMenuXmlTests()
    {
        var excelRibbon = new SmilesRibbon();
        var excelXml = excelRibbon.GetCustomUI("");
        Program.ValidateRibbonXml(excelXml, typeof(SmilesRibbon), "Excel");
        var excelDoc = XDocument.Parse(excelXml);
        Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextExcelCell" &&
                                              (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "Excel ContextMenuCell contains menuContextExcelCell anchored after PasteGalleryMini");
        Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextExcelList" &&
                                              (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "Excel ContextMenuListRange contains menuContextExcelList anchored after PasteGalleryMini");
        Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("idMso") == "ContextMenuPicture" &&
                                              e.Descendants().Any(b => (string?)b.Attribute("insertBeforeMso") == "Cut")),
            "Excel ContextMenuPicture has items anchored before Cut");

        var wordAddIn = new WordAddIn();
        var wordXml = wordAddIn.GetCustomUI("");
        Program.ValidateRibbonXml(wordXml, typeof(WordAddIn), "Word");
        var wordDoc = XDocument.Parse(wordXml);
        Check(wordDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextWordText" &&
                                             (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "Word ContextMenuText contains menuContextWordText anchored after PasteGalleryMini");
        Check(wordDoc.Descendants().Any(e => (string?)e.Attribute("idMso") == "ContextMenuPicture" &&
                                             e.Descendants().Any(b => (string?)b.Attribute("insertBeforeMso") == "Cut")),
            "Word ContextMenuPicture has items anchored before Cut");

        var pptAddIn = new PowerPointAddIn();
        var pptXml = pptAddIn.GetCustomUI("");
        Program.ValidateRibbonXml(pptXml, typeof(PowerPointAddIn), "PowerPoint");
        var pptDoc = XDocument.Parse(pptXml);
        Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextPptFrame" &&
                                            (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "PowerPoint ContextMenuFrame contains menuContextPptFrame anchored after PasteGalleryMini");
        Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextPptText" &&
                                            (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "PowerPoint ContextMenuText contains menuContextPptText anchored after PasteGalleryMini");
        Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextPptTable" &&
                                            (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "PowerPoint ContextMenuTable contains menuContextPptTable anchored after PasteGalleryMini");
        Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("idMso") == "ContextMenuPicture" &&
                                            e.Descendants().Any(b => (string?)b.Attribute("insertBeforeMso") == "Cut")),
            "PowerPoint ContextMenuPicture has items anchored before Cut");

        PassSuite("Unit: Excel, Word & PowerPoint Ribbon/ContextMenu XML & VerifiedValidOfficeImageMsos");
    }

    private static void RunWinFormsDialogsIconsAndClipboardTests()
    {
        Check(UiIcons.MoleculeInfoIcon != null && UiIcons.MoleculeInfoIcon.Handle != IntPtr.Zero, "UiIcons.MoleculeInfoIcon has valid handle");
        Check(UiIcons.AboutIcon != null && UiIcons.AboutIcon.Handle != IntPtr.Zero, "UiIcons.AboutIcon has valid handle");
        Check(UiIcons.ChemicalIcon != null && UiIcons.ChemicalIcon.Handle != IntPtr.Zero, "UiIcons.ChemicalIcon has valid handle");
        Check(UiIcons.Molstar3DIcon != null && UiIcons.Molstar3DIcon.Handle != IntPtr.Zero, "UiIcons.Molstar3DIcon has valid handle");
        Check(UiIcons.GetPropertyIcon("Rings & Aromaticity") != null, "UiIcons.GetPropertyIcon('Rings & Aromaticity') valid");
        Check(UiIcons.GetPropertyIcon("3D Complexity (Fsp3)") != null, "UiIcons.GetPropertyIcon('3D Complexity (Fsp3)') valid");

        using (var aboutDlg = new AboutDialog())
        {
            Check(aboutDlg is not null && !string.IsNullOrWhiteSpace(aboutDlg.Text), "AboutDialog instantiated with window title");
        }

        var ccoInfo = MoleculeRenderer.TryGetMoleculeInfo("CCO", resolvePubChem: false)!;
        using (var ccoDlg = new MoleculeInfoDialog(ccoInfo))
        {
            Check(ccoDlg is not null && ccoDlg.Icon != null && ccoDlg.Icon.Handle != IntPtr.Zero, "MoleculeInfoDialog instantiated with valid Win32 icon handle");
        }

        var summary = ccoInfo.FormatClipboardSummary();
        Check(SafeClipboard.SetText(summary), "SafeClipboard.SetText succeeds");
        Check(SafeClipboard.ContainsText() && SafeClipboard.GetText() == summary, "SafeClipboard.GetText round-trip");
        Check(SafeClipboard.SetTextAndHtml(ccoInfo.FormatMultiLine(), ccoInfo.FormatHtmlTable()), "SafeClipboard.SetTextAndHtml succeeds");
        Check(SafeClipboard.ContainsHtml(), "Clipboard contains DataFormats.Html");

        CopiedToast.Show("Test Molecule", "CCO", isWarning: false);
        CopiedToast.Show("Test Warning", "Notice", isWarning: true);

        PassSuite("Unit: WinForms Dialogs (AboutDialog, MoleculeInfoDialog), GDI+ UiIcons & SafeClipboard");
    }

    private static void RunSemVerComparisonTests()
    {
        Check(!AboutDialog.IsNewerVersion("v1.6.0", "v1.7.0-preview.5"), "1.6.0 is not newer than 1.7.0-preview.5");
        Check(!AboutDialog.IsNewerVersion("1.6.0", "1.7.0-preview.5"), "1.6.0 is not newer than 1.7.0-preview.5 without 'v'");
        Check(!AboutDialog.IsNewerVersion("1.7.0-preview.6", "1.7.0-preview.5"), "preview updates are ignored");
        Check(!AboutDialog.IsNewerVersion("1.8.0-preview.1", "1.7.0-preview.5"), "higher preview updates are ignored");
        Check(AboutDialog.IsNewerVersion("v1.7.0", "1.7.0-preview.5"), "1.7.0 stable is newer than 1.7.0-preview.5");
        Check(AboutDialog.IsNewerVersion("2.0.0", "1.7.0-preview.5"), "2.0.0 stable is newer than 1.7.0-preview.5");
        Check(!AboutDialog.IsNewerVersion("1.6.9", "1.7.0"), "1.6.9 is not newer than 1.7.0");
        Check(!AboutDialog.IsNewerVersion(null, "1.7.0"), "null is not newer");

        PassSuite("Unit: AboutDialog.IsNewerVersion SemVer Comparison Logic");
    }
}
