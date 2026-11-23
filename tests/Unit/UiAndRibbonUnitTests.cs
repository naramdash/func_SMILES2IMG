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
        RunBatchStyleFormulaRewriterTests();
        RunMouseDoubleClickHookFilterTests();
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
        Check(excelRibbon.FriendlyName == "func_SMILES2IMG (Ribbon UI)",
            $"Excel ribbon FriendlyName unified: {excelRibbon.FriendlyName}");
        Check(!string.IsNullOrWhiteSpace(excelRibbon.Description) && excelRibbon.Description.Contains("Excel"),
            $"Excel ribbon Description populated: {excelRibbon.Description}");
        var excelXml = excelRibbon.GetCustomUI("");
        Program.ValidateRibbonXml(excelXml, typeof(SmilesRibbon), "Excel");
        var excelDoc = XDocument.Parse(excelXml);
        Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextExcelCell" &&
                                              (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "Excel ContextMenuCell contains menuContextExcelCell anchored after PasteGalleryMini");
        Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextExcelList" &&
                                              (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "Excel ContextMenuListRange contains menuContextExcelList anchored after PasteGalleryMini");
        static void AssertContextMenuSubmenuBeforeCut(XDocument doc, string host, string idMso, string menuId) =>
            Check(doc.Descendants().Any(e => (string?)e.Attribute("idMso") == idMso &&
                                             e.Elements().Any(m => m.Name.LocalName == "menu" &&
                                                                   (string?)m.Attribute("id") == menuId &&
                                                                   (string?)m.Attribute("insertBeforeMso") == "Cut")),
                $"{host} {idMso} wraps items inside {menuId} anchored before Cut");

        AssertContextMenuSubmenuBeforeCut(excelDoc, "Excel", "ContextMenuPicture", "menuContextExcelPicture");
        AssertContextMenuSubmenuBeforeCut(excelDoc, "Excel", "ContextMenuShape", "menuContextExcelShape");
        Check(excelDoc.Descendants().Any(e => e.Name.LocalName == "button" &&
                                              (string?)e.Attribute("id") == "btnImportSdf" &&
                                              (string?)e.Attribute("onAction") == "OnImportSdf"),
            "Excel ribbon exposes standalone Import SDF button");
        Check(!excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "btnExportSdf" ||
                                               (string?)e.Attribute("onAction") == "OnExportSdf"),
            "Excel ribbon no longer exposes heuristic Export SDF action");

        var wordAddIn = new WordAddIn();
        var wordXml = wordAddIn.GetCustomUI("");
        Program.ValidateRibbonXml(wordXml, typeof(WordAddIn), "Word");
        var wordDoc = XDocument.Parse(wordXml);
        Check(wordDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextWordText" &&
                                             (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
            "Word ContextMenuText contains menuContextWordText anchored after PasteGalleryMini");
        AssertContextMenuSubmenuBeforeCut(wordDoc, "Word", "ContextMenuInlinePicture", "menuContextWordInline");
        AssertContextMenuSubmenuBeforeCut(wordDoc, "Word", "ContextMenuPicture", "menuContextWordPicture");
        AssertContextMenuSubmenuBeforeCut(wordDoc, "Word", "ContextMenuShape", "menuContextWordShape");

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
        AssertContextMenuSubmenuBeforeCut(pptDoc, "PowerPoint", "ContextMenuPicture", "menuContextPptPicture");
        AssertContextMenuSubmenuBeforeCut(pptDoc, "PowerPoint", "ContextMenuShape", "menuContextPptShape");

        foreach (var (hostName, doc) in new[] { ("Excel", excelDoc), ("Word", wordDoc), ("PowerPoint", pptDoc) })
        {
            Check(doc.Descendants().Any(e => e.Name.LocalName == "button" &&
                                             (string?)e.Attribute("onAction") == "OnOpenDocs" &&
                                             (string?)e.Attribute("label") == "Online Docs" &&
                                             (string?)e.Attribute("imageMso") == "TableOfContentsGallery"),
                $"{hostName}: Online Docs button is present in About group");
            Check(doc.Descendants().Any(e => e.Name.LocalName == "button" &&
                                             (string?)e.Attribute("onAction") == "OnAbout" &&
                                             (string?)e.Attribute("label") == "About func_SMILES2IMG"),
                $"{hostName}: About button label is 'About func_SMILES2IMG'");
            Check(!doc.Descendants().Any(e => (string?)e.Attribute("label") == "About SMILES"),
                $"{hostName}: no ribbon element uses prohibited 'About SMILES' label");

            var contextMenus = doc.Descendants().Where(e => e.Name.LocalName == "contextMenu").ToList();
            Check(contextMenus.Count > 0, $"{hostName}: has <contextMenu> entries");
            foreach (var cm in contextMenus)
            {
                var idMso = (string?)cm.Attribute("idMso") ?? "(unknown)";
                var directButtons = cm.Elements().Where(e => e.Name.LocalName == "button").ToList();
                Check(directButtons.Count == 0,
                    $"{hostName} {idMso}: no flattened top-level <button> items directly inside <contextMenu>");
                Check(cm.Elements().Any(e => e.Name.LocalName == "menu" &&
                                             (string?)e.Attribute("label") == "Chemical Structure"),
                    $"{hostName} {idMso}: exposes a top-level 'Chemical Structure' flyout <menu>");
            }
        }

        PassSuite("Unit: Excel, Word & PowerPoint Ribbon/ContextMenu XML & VerifiedValidOfficeImageMsos");
    }

    private static void RunBatchStyleFormulaRewriterTests()
    {
        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SMILES2IMG(A2)", "trans", "bw", out var r1) == FormulaRewriteStatus.Updated &&
              r1 == "=SMILES2IMG(A2, \"trans\", \"bw\")",
            "Rewrites basic SMILES2IMG formula cleanly: " + r1);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SMILES2IMG(A2, \"white\", \"color\", \"rot90\")", "trans", "bw", out var r2) == FormulaRewriteStatus.Updated &&
              r2 == "=SMILES2IMG(A2, \"trans\", \"bw\", \"rot90\")",
            "Preserves 4th transform argument: " + r2);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SMILES2IMG(A2, \"white\", \"bw num stereo hl:[N,O]\", \"rot180\")", "trans", "white", out var r3) == FormulaRewriteStatus.Updated &&
              r3 == "=SMILES2IMG(A2, \"trans\", \"white num stereo hl:[N,O]\", \"rot180\")",
            "Preserves structural style tokens (num, stereo, hl:[N,O]) and replaces only color/ink: " + r3);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SMILES2IMG(C5, , \"hl:\" & $B$3)", "trans", "bw", out var r4) == FormulaRewriteStatus.Updated &&
              r4 == "=SMILES2IMG(C5, \"trans\", \"bw \" & (\"hl:\" & $B$3))",
            "Preserves dynamic highlight expression in arg 3: " + r4);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle(r4, "trans", "white", out var r4Second) == FormulaRewriteStatus.Updated &&
              r4Second == "=SMILES2IMG(C5, \"trans\", \"white \" & (\"hl:\" & $B$3))",
            "Repeated Batch Style unwraps previous batch prefix instead of nesting: " + r4Second);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SMILES2IMG(TRIM(A2), \"white\", \"h\")", "trans", "color", out var r5) == FormulaRewriteStatus.Updated &&
              r5 == "=SMILES2IMG(TRIM(A2), \"trans\", \"color h\")",
            "Preserves nested call in arg 1 and h flag: " + r5);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=KET2IMG(\"{\"\"root\"\":{\"\"nodes\"\":[]}}\")", "white", "color", out var r6) == FormulaRewriteStatus.Updated &&
              r6 == "=KET2IMG(\"{\"\"root\"\":{\"\"nodes\"\":[]}}\", \"white\", \"color\")",
            "Preserves escaped KET JSON literal in arg 1: " + r6);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=@_xll.INCHI2IMG.FLOAT('Sheet 1'!A2, , \"num\")", "trans", "bw", out var r7) == FormulaRewriteStatus.Updated &&
              r7 == "=@_xll.INCHI2IMG.FLOAT('Sheet 1'!A2, \"trans\", \"bw num\")",
            "Preserves @ and _xll. prefixes, quoted sheet ref, and .FLOAT suffix: " + r7);

        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=IF(A2=\"\", \"\", SMILES2IMG(A2))", "trans", "bw", out _) == FormulaRewriteStatus.SkippedComposite,
            "Skips composite IF(..., SMILES2IMG(...)) formulas without corrupting them");
        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SMILES2IMG(A2) & \"\"", "trans", "bw", out _) == FormulaRewriteStatus.SkippedComposite,
            "Skips image formula with trailing expression");
        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=SUM(A1:A10)", "trans", "bw", out _) == FormulaRewriteStatus.NotApplicable,
            "Ignores unrelated formulas");
        Check(ExcelFormulaRewriter.TryRewriteImageStyle("=IF(A1=\"SMILES2IMG(\", 1, 2)", "trans", "bw", out _) == FormulaRewriteStatus.NotApplicable,
            "Does not confuse string literals mentioning SMILES2IMG( for composite formulas");

        PassSuite("Unit: Lossless ExcelFormulaRewriter for Batch Style");
    }

    private static void RunMouseDoubleClickHookFilterTests()
    {
        var dblClick = (IntPtr)0x0203;
        Check(ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(0, dblClick, 1, "EXCEL7", ["EXCEL7"]),
            "Accepts HC_ACTION double-click in EXCEL7 client area");
        Check(ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(0, dblClick, 1, "_WwG", ["_WwG"]),
            "Accepts HC_ACTION double-click in Word _WwG client area");
        Check(ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(0, dblClick, 1, "mdiClass", ["mdiClass", "paneClassDC"]),
            "Accepts HC_ACTION double-click in PowerPoint mdiClass client area");
        Check(!ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(3, dblClick, 1, "EXCEL7", ["EXCEL7"]),
            "Rejects HC_NOREMOVE hook calls");
        Check(!ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(-1, dblClick, 1, "EXCEL7", ["EXCEL7"]),
            "Rejects negative nCode hook calls");
        Check(!ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(0, dblClick, 2, "EXCEL7", ["EXCEL7"]),
            "Rejects non-client hits (e.g. HTCAPTION / borders)");
        Check(!ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(0, dblClick, 1, "NetUIHWND", ["EXCEL7"]),
            "Rejects ribbon / task pane NetUIHWND double-clicks");
        Check(!ThreadMouseDoubleClickHook.ShouldHandleDoubleClick(0, dblClick, 1, "PPTFrameClass", ["mdiClass", "paneClassDC"]),
            "Rejects frame/chrome window double-clicks in PowerPoint");

        PassSuite("Unit: ThreadMouseDoubleClickHook HC_ACTION, Hit-Test & Canvas Class Filtering");
    }

    private static void RunWinFormsDialogsIconsAndClipboardTests()
    {
        Check(UiIcons.MoleculeInfoIcon != null && UiIcons.MoleculeInfoIcon.Handle != IntPtr.Zero, "UiIcons.MoleculeInfoIcon has valid handle");
        Check(UiIcons.AboutIcon != null && UiIcons.AboutIcon.Handle != IntPtr.Zero, "UiIcons.AboutIcon has valid handle");
        Check(UiIcons.ChemicalIcon != null && UiIcons.ChemicalIcon.Handle != IntPtr.Zero, "UiIcons.ChemicalIcon has valid handle");
        Check(UiIcons.Molstar3DIcon != null && UiIcons.Molstar3DIcon.Handle != IntPtr.Zero, "UiIcons.Molstar3DIcon has valid handle");
        Check(UiIcons.GetPropertyIcon("Rings & Aromaticity") != null, "UiIcons.GetPropertyIcon('Rings & Aromaticity') valid");
        Check(UiIcons.GetPropertyIcon("3D Complexity (Fsp3)") != null, "UiIcons.GetPropertyIcon('3D Complexity (Fsp3)') valid");

        Check(AboutDialog.GetDocsUrl("Excel") == "https://naramdash.github.io/func_SMILES2IMG/excel/", "GetDocsUrl routes Excel to /excel/");
        Check(AboutDialog.GetDocsUrl("Word") == "https://naramdash.github.io/func_SMILES2IMG/word/", "GetDocsUrl routes Word to /word/");
        Check(AboutDialog.GetDocsUrl("PowerPoint") == "https://naramdash.github.io/func_SMILES2IMG/powerpoint/", "GetDocsUrl routes PowerPoint to /powerpoint/");
        Check(AboutDialog.GetDocsUrl(null) == "https://naramdash.github.io/func_SMILES2IMG/", "GetDocsUrl defaults to root docs URL");

        using (var aboutDlg = new AboutDialog("Excel"))
        {
            Check(aboutDlg is not null && !string.IsNullOrWhiteSpace(aboutDlg.Text), "AboutDialog instantiated with window title");
            Check(aboutDlg!.TargetDocsUrl == "https://naramdash.github.io/func_SMILES2IMG/excel/", "AboutDialog stores host-specific TargetDocsUrl");
            Check(FindButton(aboutDlg, "Open Docs") != null, "AboutDialog contains 'Open Docs' action button");
        }

        static System.Windows.Forms.Button? FindButton(System.Windows.Forms.Control root, string text)
        {
            foreach (System.Windows.Forms.Control child in root.Controls)
            {
                if (child is System.Windows.Forms.Button btn && btn.Text == text) return btn;
                var nested = FindButton(child, text);
                if (nested != null) return nested;
            }
            return null;
        }

        var ccoInfo = MoleculeRenderer.TryGetMoleculeInfo("CCO", resolvePubChem: false)!;
        using (var ccoDlg = new MoleculeInfoDialog(ccoInfo))
        {
            Check(ccoDlg is not null && ccoDlg.Icon != null && ccoDlg.Icon.Handle != IntPtr.Zero, "MoleculeInfoDialog instantiated with valid Win32 icon handle");
        }

        // Verify ultra-long compound names keep the 2D header preview inside client bounds
        var longInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O", resolvePubChem: false)!;
        longInfo.CompoundName = "(2S)-2-{[(R)-{[(2R,3S,4R,5R)-5-(4-aminopyrrolo[2,1-f][1,2,4]triazin-7-yl)-5-cyano-3,4-dihydroxytetrahydrofuran-2-yl]methoxy}(phenoxy)phosphoryl]amino}propanoate";
        using (var longDlg = new MoleculeInfoDialog(longInfo))
        {
            longDlg.CreateControl();
            longDlg.PerformLayout();
            var pic = FindPictureBox(longDlg);
            Check(pic != null && pic!.Right <= longDlg.ClientSize.Width,
                $"Long compound name keeps header 2D preview within client bounds ({pic?.Right} <= {longDlg.ClientSize.Width})");
        }

        static System.Windows.Forms.PictureBox? FindPictureBox(System.Windows.Forms.Control root)
        {
            foreach (System.Windows.Forms.Control child in root.Controls)
            {
                if (child is System.Windows.Forms.PictureBox pb) return pb;
                var nested = FindPictureBox(child);
                if (nested != null) return nested;
            }
            return null;
        }

        var summary = ccoInfo.FormatClipboardSummary();
        Check(SafeClipboard.SetText(summary), "SafeClipboard.SetText succeeds");
        Check(SafeClipboard.ContainsText() && SafeClipboard.GetText() == summary, "SafeClipboard.GetText round-trip");
        Check(SafeClipboard.SetTextAndHtml(ccoInfo.FormatMultiLine(), ccoInfo.FormatHtmlTable()), "SafeClipboard.SetTextAndHtml succeeds");
        Check(SafeClipboard.ContainsHtml(), "Clipboard contains DataFormats.Html");

        CopiedToast.Show("Test Molecule", "CCO", isWarning: false);
        CopiedToast.Show("Test Warning", "Notice", isWarning: true);

        string? capturedTempPath = null;
        var tempResult = StructureMetadata.WithTempFile(new byte[] { 1, 2, 3, 4 }, "unit_probe", "bin", path =>
        {
            capturedTempPath = path;
            Check(File.Exists(path) && File.ReadAllBytes(path).Length == 4, "WithTempFile writes payload to existing temp file");
            return 42;
        });
        Check(tempResult == 42 && capturedTempPath != null && !File.Exists(capturedTempPath), "WithTempFile cleans up temp file after completion");
        Check(MolstarSessionManager.FormatTitle("4HHB") == "Mol* 3D: 4HHB" &&
              MolstarSessionManager.FormatTitle("") == "Mol* 3D: Structure",
            "MolstarSessionManager.FormatTitle formats valid PDB and default title consistently");

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
        Check(AboutDialog.IsNewerVersion("v1.8.0", "v1.8.0-preview.4"), "1.8.0 stable is newer than 1.8.0-preview.4");
        Check(!AboutDialog.IsNewerVersion("v1.8.0", "v1.8.0"), "1.8.0 is not newer than 1.8.0");

        PassSuite("Unit: AboutDialog.IsNewerVersion SemVer Comparison Logic");
    }
}
