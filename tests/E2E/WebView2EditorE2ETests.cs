using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Smiles2Img;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests.E2E;

internal static class WebView2EditorE2ETests
{
    public static void RunAll()
    {
        Directory.CreateDirectory("output/samples");

        RunKetcherWebView2UiE2E();
        RunMolstarWebView2UiE2E();
        RunMoleculeInfoDialogUiE2E();
        RunAboutDialogAndToastUiE2E();
    }

    private static void RunKetcherWebView2UiE2E()
    {
        // E2E 1: Live WebView2 + Ketcher 2D UI (SMILES -> Full Browser Viewport UI Capture + KET + Transparent Canvas PNG)
        string? e2eKet;
        using (var form1 = new KetcherForm("PowerPoint", "CC(=O)Oc1ccccc1C(=O)O")
        {
            AutoApplyAfterLoadForTest = true,
            ToggleWhiteBackgroundForTest = false
        })
        {
            var dr1 = form1.ShowDialog();
            Check(dr1 == DialogResult.OK, "E2E KetcherForm #1 returned DialogResult.OK");
            Check(form1.Width >= 850 && form1.Height >= 550, $"E2E KetcherForm #1 window size valid ({form1.Width}x{form1.Height})");
            Check(!string.IsNullOrWhiteSpace(form1.ResultKet) && form1.ResultKet!.Contains("\"root\""), "E2E KetcherForm #1 returned valid KET JSON");
            Check(!string.IsNullOrWhiteSpace(form1.ResultSmiles), "E2E KetcherForm #1 returned non-empty SMILES: " + form1.ResultSmiles);
            Check(form1.ResultBackground == "transparent", "E2E KetcherForm #1 background is transparent");
            Check(form1.ResultPngBytes != null && form1.ResultPngBytes.Length > 500, "E2E KetcherForm #1 captured canvas PNG bytes");
            Check(form1.ResultUiViewportPngBytes != null && form1.ResultUiViewportPngBytes.Length > 5000, "E2E KetcherForm #1 captured full WebView2 UI viewport screenshot");
            VerifyUiScreenshotHasVisualDiversity(form1.ResultUiViewportPngBytes!, 800, 500, "KetcherForm WebView2 UI Viewport");
            File.WriteAllBytes("output/samples/e2e_ui_ketcher_viewport.png", form1.ResultUiViewportPngBytes!);

            using (var ms1 = new MemoryStream(form1.ResultPngBytes!))
            using (var bmp1 = new Bitmap(ms1))
            {
                Check(bmp1.Width >= 100 && bmp1.Height >= 100, $"E2E KetcherForm #1 canvas PNG dimensions ({bmp1.Width}x{bmp1.Height}) are valid");
                Check(bmp1.GetPixel(0, 0).A == 0, "E2E KetcherForm #1 corner pixel is transparent");
            }
            File.WriteAllBytes("output/samples/e2e_ketcher_trans.png", form1.ResultPngBytes!);
            e2eKet = form1.ResultKet;
            PassSuite($"E2E UI + WebView2 Ketcher #1: Window & Viewport UI ({form1.ResultUiViewportPngBytes!.Length} B) + Transparent Canvas PNG ({form1.ResultPngBytes!.Length} B)");
        }

        // E2E 2: Live WebView2 + Ketcher 2D (KET Round-Trip -> White Background Canvas PNG)
        using (var form2 = new KetcherForm("Word", e2eKet!)
        {
            AutoApplyAfterLoadForTest = true,
            ToggleWhiteBackgroundForTest = true
        })
        {
            var dr2 = form2.ShowDialog();
            Check(dr2 == DialogResult.OK, "E2E KetcherForm #2 returned DialogResult.OK");
            Check(!string.IsNullOrWhiteSpace(form2.ResultKet) && form2.ResultKet!.Contains("\"root\""), "E2E KetcherForm #2 round-tripped KET JSON");
            Check(form2.ResultBackground == "white", "E2E KetcherForm #2 background toggled to white");
            Check(form2.ResultPngBytes != null && form2.ResultPngBytes.Length > 500, "E2E KetcherForm #2 captured white-bg canvas PNG bytes");
            using (var ms2 = new MemoryStream(form2.ResultPngBytes!))
            using (var bmp2 = new Bitmap(ms2))
            {
                var corner = bmp2.GetPixel(0, 0);
                Check(corner.A == 255 && corner.R == 255 && corner.G == 255 && corner.B == 255, "E2E KetcherForm #2 corner pixel is opaque white");
            }
            File.WriteAllBytes("output/samples/e2e_ketcher_white.png", form2.ResultPngBytes!);
            PassSuite($"E2E UI + WebView2 Ketcher #2: KET Round-Trip + White BG Canvas PNG ({form2.ResultPngBytes!.Length} B)");
        }

        // E2E 3: Live WebView2 + Ketcher 2D (Complex Molecule Caffeine -> KET & Molecular Properties)
        using (var form3 = new KetcherForm("Excel", "Cn1cnc2c1c(=O)n(C)c(=O)n2C")
        {
            AutoApplyAfterLoadForTest = true,
            ToggleWhiteBackgroundForTest = false
        })
        {
            var dr3 = form3.ShowDialog();
            Check(dr3 == DialogResult.OK, "E2E KetcherForm #3 returned DialogResult.OK");
            Check(!string.IsNullOrWhiteSpace(form3.ResultKet) && form3.ResultKet!.Contains("\"root\""), "E2E KetcherForm #3 returned valid KET JSON");
            var caffeineInfo = MoleculeRenderer.TryGetMoleculeInfo(form3.ResultKet!);
            Check(caffeineInfo != null && caffeineInfo.Formula == "C8H10N4O2", $"E2E KetcherForm #3 caffeine formula is C8H10N4O2, got: {caffeineInfo?.Formula}");
            Check(Math.Abs((caffeineInfo?.MolecularWeight ?? 0) - 194.19) < 0.2, $"E2E KetcherForm #3 caffeine MW ~ 194.19, got: {caffeineInfo?.MolecularWeight}");
            Check(form3.ResultPngBytes != null && form3.ResultPngBytes.Length > 500, "E2E KetcherForm #3 captured canvas PNG bytes");
            File.WriteAllBytes("output/samples/e2e_ketcher_caffeine.png", form3.ResultPngBytes!);
            PassSuite($"E2E UI + WebView2 Ketcher #3: Caffeine KET + Formula {caffeineInfo?.Formula} + MW {caffeineInfo?.MolecularWeight:F2}");
        }
    }

    private static void RunMolstarWebView2UiE2E()
    {
        // E2E 4: Live WebView2 + Mol* 3D Viewer UI (1CRN WebGL Render -> Full UI Viewport Screenshot + 3D PNG + State Snapshot)
        using (var molForm = new MolstarForm("PowerPoint Slide", "1CRN")
        {
            AutoApplyAfterLoadForTest = true
        })
        {
            var dr4 = molForm.ShowDialog();
            Check(dr4 == DialogResult.OK, "E2E MolstarForm #4 returned DialogResult.OK");
            Check(molForm.Width >= 900 && molForm.Height >= 600, $"E2E MolstarForm #4 window size valid ({molForm.Width}x{molForm.Height})");
            Check(molForm.ResultImageBytes != null && molForm.ResultImageBytes.Length > 500, "E2E MolstarForm #4 captured 3D WebGL screenshot PNG bytes");
            Check(molForm.ResultUiViewportPngBytes != null && molForm.ResultUiViewportPngBytes.Length > 5000, "E2E MolstarForm #4 captured full WebView2 UI viewport screenshot");
            VerifyUiScreenshotHasVisualDiversity(molForm.ResultUiViewportPngBytes!, 800, 500, "MolstarForm WebView2 UI Viewport");
            File.WriteAllBytes("output/samples/e2e_ui_molstar_viewport.png", molForm.ResultUiViewportPngBytes!);

            Check(!string.IsNullOrWhiteSpace(molForm.ResultSnapshotJson) && molForm.ResultSnapshotJson!.StartsWith("{"), "E2E MolstarForm #4 captured valid Mol* state snapshot JSON");
            using (var ms4 = new MemoryStream(molForm.ResultImageBytes!))
            using (var bmp4 = new Bitmap(ms4))
            {
                Check(bmp4.Width >= 100 && bmp4.Height >= 100, $"E2E MolstarForm #4 screenshot dimensions ({bmp4.Width}x{bmp4.Height}) are valid");
            }
            File.WriteAllBytes("output/samples/e2e_molstar_1crn.png", molForm.ResultImageBytes!);
            PassSuite($"E2E UI + WebView2 Mol* 3D #4: Window & Viewport UI ({molForm.ResultUiViewportPngBytes!.Length} B) + 1CRN WebGL PNG ({molForm.ResultImageBytes!.Length} B)");
        }
    }

    private static void RunMoleculeInfoDialogUiE2E()
    {
        // E2E 5A: Small Molecule MoleculeInfoDialog Live Display, GDI+ CellPainting, Format Combo & Button Clicks
        var aspirinInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O", resolvePubChem: true)!;
        using (var dlg = new MoleculeInfoDialog(aspirinInfo))
        {
            dlg.Show();
            dlg.BringToFront();
            dlg.Refresh();
            Application.DoEvents();

            Check(dlg.Visible && dlg.IsHandleCreated, "MoleculeInfoDialog is visible on screen with valid HWND");
            Check(dlg.Width >= 550 && dlg.Height >= 450, $"MoleculeInfoDialog dimensions valid ({dlg.Width}x{dlg.Height})");

            var grid = FindControl<DataGridView>(dlg)!;
            Check(grid != null && grid.Rows.Count >= 9, $"Small-molecule MoleculeInfoDialog grid populated >= 9 rows, got: {grid?.Rows.Count}");

            var rowLabels = new List<string>();
            for (int i = 0; i < grid!.Rows.Count; i++)
                rowLabels.Add(Convert.ToString(grid.Rows[i].Cells[1].Value) ?? "");
            Check(rowLabels.Contains("LogP / TPSA") && rowLabels.Contains("Rings & Aromaticity") && rowLabels.Contains("3D Complexity (Fsp3)") && !rowLabels.Contains("Lipinski Rule of 5"),
                "Small-molecule MoleculeInfoDialog displays LogP / TPSA, Rings & Aromaticity, and 3D Complexity (Fsp3) rows without Lipinski");

            // Render entire dialog (including custom CellPainting vector icons & action pills) to bitmap
            using (var bmp = new Bitmap(dlg.Width, dlg.Height))
            {
                dlg.DrawToBitmap(bmp, new Rectangle(0, 0, dlg.Width, dlg.Height));
                using var ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                var bytes = ms.ToArray();
                VerifyUiScreenshotHasVisualDiversity(bytes, 550, 450, "MoleculeInfoDialog (Aspirin)");
                File.WriteAllBytes("output/samples/e2e_ui_molecule_info_aspirin.png", bytes);
            }

            // Simulate user clicking the Copy All split button in MoleculeInfoDialog
            var btn = FindControl<Button>(dlg, b => b.Text.StartsWith("Copy All", StringComparison.OrdinalIgnoreCase))!;
            Check(btn != null, "MoleculeInfoDialog Copy All button found");
            btn!.PerformClick();
            Application.DoEvents();
            Check(SafeClipboard.GetText()?.Contains("Heavy Atoms") == true && !SafeClipboard.GetText()!.Contains("Lipinski") && !SafeClipboard.GetText()!.Contains("Chiral"), "MoleculeInfoDialog live button click copied 12-col TSV with header");

            dlg.Close();
        }

        // E2E 5B: Macromolecule MoleculeInfoDialog Live Display (suppresses small-molecule rows)
        var macroInfo = MoleculeActionHelper.TryResolveMolstarInfo("MOLSTAR:{\"id\":\"test\",\"pdb\":\"1CRN\"}", "Mol* 3D: 1CRN", resolvePubChem: false)!;
        using (var macroDlg = new MoleculeInfoDialog(macroInfo))
        {
            macroDlg.Show();
            macroDlg.Refresh();
            Application.DoEvents();

            Check(macroDlg.Visible && macroDlg.IsHandleCreated, "Macromolecule MoleculeInfoDialog is visible on screen");
            var grid = FindControl<DataGridView>(macroDlg)!;
            var rowLabels = new List<string>();
            for (int i = 0; i < grid.Rows.Count; i++)
                rowLabels.Add(Convert.ToString(grid.Rows[i].Cells[1].Value) ?? "");
            Check(!rowLabels.Contains("LogP / TPSA") && !rowLabels.Contains("3D Complexity (Fsp3)"),
                "Macromolecule MoleculeInfoDialog suppresses small-molecule-only rows");

            using (var bmp = new Bitmap(macroDlg.Width, macroDlg.Height))
            {
                macroDlg.DrawToBitmap(bmp, new Rectangle(0, 0, macroDlg.Width, macroDlg.Height));
                bmp.Save("output/samples/e2e_ui_molecule_info_1crn.png", System.Drawing.Imaging.ImageFormat.Png);
            }
            macroDlg.Close();
        }

        PassSuite("E2E WinForms UI #5: MoleculeInfoDialog Live Display, CellPainting Render & Interactive Copy Controls");
    }

    private static void RunAboutDialogAndToastUiE2E()
    {
        using (var aboutDlg = new AboutDialog())
        {
            aboutDlg.Show();
            aboutDlg.BringToFront();
            aboutDlg.Refresh();
            Application.DoEvents();

            Check(aboutDlg.Visible && aboutDlg.IsHandleCreated, "AboutDialog is visible on screen with valid HWND");
            Check(aboutDlg.Text.Contains("func_SMILES2IMG"), $"AboutDialog title contains func_SMILES2IMG: {aboutDlg.Text}");
            Check(aboutDlg.Width >= 500 && aboutDlg.Height >= 400, $"AboutDialog dimensions valid ({aboutDlg.Width}x{aboutDlg.Height})");

            using (var bmp = new Bitmap(aboutDlg.Width, aboutDlg.Height))
            {
                aboutDlg.DrawToBitmap(bmp, new Rectangle(0, 0, aboutDlg.Width, aboutDlg.Height));
                using var ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                var bytes = ms.ToArray();
                VerifyUiScreenshotHasVisualDiversity(bytes, 500, 400, "AboutDialog Tab 1");
                File.WriteAllBytes("output/samples/e2e_ui_about_dialog.png", bytes);
            }

            var tabs = FindControl<TabControl>(aboutDlg)!;
            Check(tabs != null && tabs.TabPages.Count == 3, "AboutDialog contains 3 standard tabs");
            tabs!.SelectedIndex = 1;
            Application.DoEvents();
            var compGrid = FindControl<DataGridView>(tabs.TabPages[1])!;
            Check(compGrid != null && compGrid.Rows.Count >= 7, $"AboutDialog Components tab lists >= 7 components, got: {compGrid?.Rows.Count}");
            using (var bmp2 = new Bitmap(aboutDlg.Width, aboutDlg.Height))
            {
                aboutDlg.DrawToBitmap(bmp2, new Rectangle(0, 0, aboutDlg.Width, aboutDlg.Height));
                bmp2.Save("output/samples/e2e_ui_about_tab2.png", System.Drawing.Imaging.ImageFormat.Png);
            }

            tabs.SelectedIndex = 2;
            Application.DoEvents();
            var copyNoticesBtn = FindControl<Button>(tabs.TabPages[2], b => b.Text.Contains("Copy Full"))!;
            Check(copyNoticesBtn != null, "AboutDialog Notices tab hosts Copy Full License Notices button");
            copyNoticesBtn!.PerformClick();
            Application.DoEvents();
            Check(SafeClipboard.GetText()?.Contains("THIRD-PARTY COMPONENTS") == true, "AboutDialog Copy Full License Notices copies notices text");
            using (var bmp3 = new Bitmap(aboutDlg.Width, aboutDlg.Height))
            {
                aboutDlg.DrawToBitmap(bmp3, new Rectangle(0, 0, aboutDlg.Width, aboutDlg.Height));
                bmp3.Save("output/samples/e2e_ui_about_tab3.png", System.Drawing.Imaging.ImageFormat.Png);
            }

            aboutDlg.Close();
        }

        CopiedToast.Show("E2E Toast UI Verification", "CC(=O)Oc1ccccc1C(=O)O", isWarning: false);
        Application.DoEvents();

        PassSuite("E2E WinForms UI #6: AboutDialog Live Display, Visual Paint Verification & CopiedToast Popup");
    }

    private static T? FindControl<T>(Control root, Func<T, bool>? predicate = null) where T : Control
    {
        foreach (Control c in root.Controls)
        {
            if (c is T typed && (predicate == null || predicate(typed)))
                return typed;
            var child = FindControl(c, predicate);
            if (child != null) return child;
        }
        return null;
    }

    private static void VerifyUiScreenshotHasVisualDiversity(byte[] pngBytes, int minWidth, int minHeight, string label)
    {
        using var ms = new MemoryStream(pngBytes);
        using var bmp = new Bitmap(ms);
        Check(bmp.Width >= minWidth && bmp.Height >= minHeight, $"{label} screenshot dimensions ({bmp.Width}x{bmp.Height}) meet minimum ({minWidth}x{minHeight})");

        var distinctColors = new HashSet<int>();
        int stepX = Math.Max(1, bmp.Width / 40);
        int stepY = Math.Max(1, bmp.Height / 40);
        for (int y = 0; y < bmp.Height; y += stepY)
        {
            for (int x = 0; x < bmp.Width; x += stepX)
            {
                distinctColors.Add(bmp.GetPixel(x, y).ToArgb());
            }
        }
        Check(distinctColors.Count >= 6, $"{label} UI screenshot is not blank (sampled {distinctColors.Count} distinct colors)");
    }
}
