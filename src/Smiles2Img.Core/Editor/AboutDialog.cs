using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Modern, responsive, and DPI-aware About & License Information Dialog
/// displaying project metadata, author info, GitHub repository link,
/// and complete third-party open-source components and licenses.
/// </summary>
public sealed class AboutDialog : Form
{
    private class ComponentInfo
    {
        public string Name { get; set; } = "";
        public string Version { get; set; } = "";
        public string License { get; set; } = "";
        public string Website { get; set; } = "";
        public string Copyright { get; set; } = "";
        public string Description { get; set; } = "";
    }

    private static readonly List<ComponentInfo> Components = new()
    {
        new ComponentInfo
        {
            Name = "EPAM Ketcher",
            Version = "3.18.0",
            License = "Apache License 2.0",
            Website = "https://github.com/epam/ketcher",
            Copyright = "Copyright (c) 2011-2024 EPAM Systems, Inc.",
            Description = "High-performance standalone molecular structure editor, chemical reaction canvas, and KET format serialization."
        },
        new ComponentInfo
        {
            Name = "EPAM Indigo Toolkit",
            Version = "1.49.0-rc.2",
            License = "Apache License 2.0",
            Website = "https://github.com/epam/Indigo",
            Copyright = "Copyright (c) 2009-2024 EPAM Systems, Inc.",
            Description = "Cheminformatics engine providing molecular layout computation, aromatization, 0ms local property calculation, and high-resolution PNG rendering."
        },
        new ComponentInfo
        {
            Name = "Mol* (Molstar) 3D Viewer",
            Version = "5.13.1",
            License = "Apache License 2.0 / MIT",
            Website = "https://github.com/molstar/molstar",
            Copyright = "Copyright (c) 2018-2024 Mol* contributors",
            Description = "Comprehensive macromolecular 3D visualization engine for PDB and AlphaFold structures, supporting interactive 3D model (.glb) generation."
        },
        new ComponentInfo
        {
            Name = "IUPAC InChI Toolkit",
            Version = "1.07.1",
            License = "IUPAC / InChI-Trust Licence #1.0",
            Website = "https://www.inchi-trust.org",
            Copyright = "Copyright (c) IUPAC and InChI Trust Limited",
            Description = "Standard algorithm for International Chemical Identifier (InChI) conversion, structure normalization, and 27-character InChIKey hashing."
        },
        new ComponentInfo
        {
            Name = "Excel-DNA & IntelliSense",
            Version = "1.9.0",
            License = "zlib / MIT License",
            Website = "https://github.com/Excel-DNA/ExcelDna",
            Copyright = "Copyright (c) 2005-2024 Govert van Drimmelen",
            Description = "High-speed C-API add-in framework connecting .NET to Microsoft Excel, providing in-cell function auto-completion and argument tooltips."
        },
        new ComponentInfo
        {
            Name = "Microsoft.Web.WebView2",
            Version = "1.0.4258.31",
            License = "Microsoft Software License",
            Website = "https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            Copyright = "Copyright (c) Microsoft Corporation",
            Description = "Chromium-based native embedding runtime enabling EPAM Ketcher and Mol* canvases inside Office without external browser processes."
        },
        new ComponentInfo
        {
            Name = "System.Drawing.Common",
            Version = "4.7.2",
            License = "MIT License",
            Website = "https://github.com/dotnet/runtime",
            Copyright = "Copyright (c) .NET Foundation and Contributors",
            Description = "Cross-platform GDI+ graphics primitives for rasterization, font measurement, and in-memory image streaming."
        }
    };

    public AboutDialog(string? appName = null)
    {
        Text = "About SMILES for Microsoft Office";

        float dpiScale = 1.0f;
        try
        {
            using var g = CreateGraphics();
            dpiScale = Math.Max(1.0f, g.DpiX / 96.0f);
        }
        catch { }

        var screen = Screen.FromControl(this) ?? Screen.PrimaryScreen;
        int targetW = Math.Min((int)(820 * dpiScale), (int)(screen.WorkingArea.Width * 0.92));
        int targetH = Math.Min((int)(680 * dpiScale), (int)(screen.WorkingArea.Height * 0.92));
        ClientSize = new Size(Math.Max(targetW, 720), Math.Max(targetH, 560));
        MinimumSize = new Size(Math.Min(650, screen.WorkingArea.Width), Math.Min(480, screen.WorkingArea.Height));

        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = new Font("Segoe UI", 9.25f, FontStyle.Regular);
        BackColor = Color.FromArgb(249, 250, 251);
        KeyPreview = true;
        KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Escape) Close();
        };

        var toolTip = new ToolTip
        {
            AutoPopDelay = 8000,
            InitialDelay = 400,
            ReshowDelay = 200,
            ShowAlways = true
        };

        // Header Panel
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.FromArgb(20, 50, 85),
            Padding = new Padding(24, 16, 24, 16)
        };

        var lblTitle = new Label
        {
            Text = "SMILES for Microsoft Office Suite",
            UseMnemonic = false,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 13.5f, FontStyle.Bold),
            AutoSize = true,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 0, 0, 4)
        };

        var lblSubtitle = new Label
        {
            Text = "Chemical Structure & Macromolecule Add-in for Excel · Word · PowerPoint",
            UseMnemonic = false,
            ForeColor = Color.FromArgb(215, 230, 248),
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            AutoSize = true,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 4, 0, 0)
        };

        pnlHeader.Controls.Add(lblSubtitle);
        pnlHeader.Controls.Add(lblTitle);

        // Bottom Action Panel
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = (int)(56 * dpiScale),
            BackColor = Color.FromArgb(243, 244, 246),
            Padding = new Padding(20, 10, 20, 10)
        };

        var btnClose = new Button
        {
            Text = "Close",
            DialogResult = DialogResult.OK,
            Size = new Size((int)(90 * dpiScale), (int)(34 * dpiScale)),
            Dock = DockStyle.Right,
            BackColor = Color.FromArgb(20, 50, 85),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Bold)
        };
        btnClose.FlatAppearance.BorderSize = 0;
        btnClose.Click += (s, e) => Close();

        var btnOpenGitHub = new Button
        {
            Text = "Open GitHub Repo",
            Size = new Size((int)(140 * dpiScale), (int)(34 * dpiScale)),
            Dock = DockStyle.Left,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(30, 41, 59),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular),
            Margin = new Padding(0, 0, 8, 0)
        };
        btnOpenGitHub.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
        btnOpenGitHub.Click += (s, e) => OpenBrowser("https://github.com/naramdash/func_SMILES2IMG");

        var btnCopyAll = new Button
        {
            Text = "Copy Notices",
            Size = new Size((int)(110 * dpiScale), (int)(34 * dpiScale)),
            Dock = DockStyle.Left,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(30, 41, 59),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular),
            Margin = new Padding(8, 0, 0, 0)
        };
        btnCopyAll.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
        btnCopyAll.Click += (s, e) =>
        {
            try
            {
                Clipboard.SetText(BuildFullNoticesText());
                btnCopyAll.Text = "✓ Copied!";
                btnCopyAll.BackColor = Color.FromArgb(236, 253, 245);
                btnCopyAll.ForeColor = Color.FromArgb(16, 149, 102);

                var timer = new Timer { Interval = 1800 };
                timer.Tick += (ts, te) =>
                {
                    btnCopyAll.Text = "Copy Notices";
                    btnCopyAll.BackColor = Color.White;
                    btnCopyAll.ForeColor = Color.FromArgb(30, 41, 59);
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
            }
            catch { }
        };

        pnlBottom.Controls.Add(btnCopyAll);
        pnlBottom.Controls.Add(btnOpenGitHub);
        pnlBottom.Controls.Add(btnClose);

        // Center Content Tab Control
        var tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
            Padding = new Point(14, 8),
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
        };

        // Tab 1: Project & Author Overview
        var tabOverview = new TabPage("Overview & Author")
        {
            BackColor = Color.White,
            Padding = new Padding(20)
        };
        BuildOverviewTab(tabOverview, dpiScale);

        // Tab 2: Third-Party Dependencies
        var tabComponents = new TabPage("Third-Party Components")
        {
            BackColor = Color.White,
            Padding = new Padding(16)
        };
        BuildComponentsTab(tabComponents, dpiScale, toolTip);

        // Tab 3: Complete Notices Text
        var tabFullNotices = new TabPage("License Notices (Full Text)")
        {
            BackColor = Color.White,
            Padding = new Padding(16)
        };
        BuildFullNoticesTab(tabFullNotices, dpiScale);

        tabControl.TabPages.Add(tabOverview);
        tabControl.TabPages.Add(tabComponents);
        tabControl.TabPages.Add(tabFullNotices);

        Controls.Add(tabControl);
        Controls.Add(pnlBottom);
        Controls.Add(pnlHeader);
    }

    private static void BuildOverviewTab(TabPage tab, float dpiScale)
    {
        var pnl = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 8,
            AutoScroll = true,
            Padding = new Padding(12)
        };
        pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(160 * dpiScale)));
        pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        void AddRow(string label, Control valueControl)
        {
            var lbl = new Label
            {
                Text = label,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(55, 65, 81),
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 6, 12, 6)
            };
            valueControl.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            valueControl.Margin = new Padding(0, 4, 0, 10);
            pnl.Controls.Add(lbl);
            pnl.Controls.Add(valueControl);
        }

        AddRow("Project:", new Label
        {
            Text = "func_SMILES2IMG (SMILES for Microsoft Office Suite)",
            Font = new Font("Segoe UI", 9.75f, FontStyle.Bold),
            ForeColor = Color.FromArgb(17, 24, 39),
            AutoSize = true
        });

        AddRow("Version:", new Label
        {
            Text = $"{GetProductVersion()} (Unified Office Edition)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(37, 99, 235),
            AutoSize = true
        });

        AddRow("Author / Creator:", new Label
        {
            Text = "KimJuho (naramdash)",
            Font = new Font("Segoe UI", 9.75f, FontStyle.Bold),
            ForeColor = Color.FromArgb(17, 24, 39),
            AutoSize = true
        });

        var linkGithub = new LinkLabel
        {
            Text = "https://github.com/naramdash/func_SMILES2IMG",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            LinkColor = Color.FromArgb(37, 99, 235),
            AutoSize = true,
            Cursor = Cursors.Hand
        };
        linkGithub.LinkClicked += (s, e) => OpenBrowser("https://github.com/naramdash/func_SMILES2IMG");
        AddRow("GitHub Repository:", linkGithub);

        var linkDocs = new LinkLabel
        {
            Text = "https://naramdash.github.io/func_SMILES2IMG/",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            LinkColor = Color.FromArgb(37, 99, 235),
            AutoSize = true,
            Cursor = Cursors.Hand
        };
        linkDocs.LinkClicked += (s, e) => OpenBrowser("https://naramdash.github.io/func_SMILES2IMG/");
        AddRow("Documentation Site:", linkDocs);

        AddRow("Project License:", new Label
        {
            Text = "MIT License (Free & Open Source)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(16, 149, 102),
            AutoSize = true
        });

        var txtDesc = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(75, 85, 99),
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular),
            Text = "SMILES for Microsoft Office Suite seamlessly integrates chemical structure sketching, formula calculation, and 3D macromolecular visualization into Excel, Word, and PowerPoint.\r\n\r\n" +
                   "• Microsoft Excel: In-cell (=SMILES2IMG, =INCHI2IMG, =KET2IMG), floating (.FLOAT), and property formulas (=SMILES2NAME, =SMILES2MW, =SMILES2FORMULA, =SMILES2INCHI, =SMILES2INCHIKEY).\r\n" +
                   "• Microsoft Word: Inline and floating high-resolution molecular images with 100% 2D KET canvas preservation and live round-trip editing.\r\n" +
                   "• Microsoft PowerPoint: Slide shapes with KET metadata, 3D macromolecular Mol* integration, and instant formula conversion.\r\n" +
                   "• Built-in Engines: EPAM Ketcher, EPAM Indigo C++ Toolkit, Mol* 3D Viewer, and IUPAC InChI Library running purely locally and offline.",
            Height = (int)(140 * dpiScale)
        };
        AddRow("Description:", txtDesc);

        tab.Controls.Add(pnl);
    }

    private static void BuildComponentsTab(TabPage tab, float dpiScale, ToolTip toolTip)
    {
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = (int)(220 * dpiScale),
            SplitterWidth = 6
        };

        var lv = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular)
        };
        lv.Columns.Add("Component", (int)(190 * dpiScale));
        lv.Columns.Add("Version", (int)(90 * dpiScale));
        lv.Columns.Add("License", (int)(160 * dpiScale));
        lv.Columns.Add("Website / Repository", (int)(260 * dpiScale));

        foreach (var c in Components)
        {
            var item = new ListViewItem(c.Name);
            item.SubItems.Add(c.Version);
            item.SubItems.Add(c.License);
            item.SubItems.Add(c.Website);
            item.Tag = c;
            lv.Items.Add(item);
        }

        var pnlDetail = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            BackColor = Color.FromArgb(250, 250, 250)
        };

        var lblDetailTitle = new Label
        {
            Dock = DockStyle.Top,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(17, 24, 39),
            Height = (int)(26 * dpiScale)
        };

        var linkDetailWeb = new LinkLabel
        {
            Dock = DockStyle.Top,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular),
            LinkColor = Color.FromArgb(37, 99, 235),
            Height = (int)(22 * dpiScale),
            Cursor = Cursors.Hand
        };
        linkDetailWeb.LinkClicked += (s, e) =>
        {
            if (!string.IsNullOrEmpty(linkDetailWeb.Text)) OpenBrowser(linkDetailWeb.Text);
        };

        var txtDetailDesc = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular),
            Margin = new Padding(0, 6, 0, 0)
        };

        pnlDetail.Controls.Add(txtDetailDesc);
        pnlDetail.Controls.Add(linkDetailWeb);
        pnlDetail.Controls.Add(lblDetailTitle);

        void UpdateDetail(ComponentInfo c)
        {
            lblDetailTitle.Text = $"{c.Name} (v{c.Version}) — {c.License}";
            linkDetailWeb.Text = c.Website;
            txtDetailDesc.Text = $"{c.Copyright}\r\n\r\n{c.Description}\r\n\r\nLicense Terms: {c.License}\r\nSource: {c.Website}";
        }

        lv.SelectedIndexChanged += (s, e) =>
        {
            if (lv.SelectedItems.Count > 0 && lv.SelectedItems[0].Tag is ComponentInfo ci)
            {
                UpdateDetail(ci);
            }
        };

        if (lv.Items.Count > 0)
        {
            lv.Items[0].Selected = true;
            UpdateDetail(Components[0]);
        }

        split.Panel1.Controls.Add(lv);
        split.Panel2.Controls.Add(pnlDetail);
        tab.Controls.Add(split);
    }

    private static void BuildFullNoticesTab(TabPage tab, float dpiScale)
    {
        var txt = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(30, 41, 59),
            Font = new Font("Consolas", 9.0f, FontStyle.Regular),
            Text = BuildFullNoticesText()
        };
        tab.Controls.Add(txt);
    }

    public static string BuildFullNoticesText()
    {
        var sb = new StringBuilder();
        sb.AppendLine("================================================================================");
        sb.AppendLine(" SMILES for Microsoft Office Suite (func_SMILES2IMG)");
        sb.AppendLine($" Version: {GetProductVersion()}");
        sb.AppendLine(" Author: KimJuho (naramdash)");
        sb.AppendLine(" Repository: https://github.com/naramdash/func_SMILES2IMG");
        sb.AppendLine(" Documentation: https://naramdash.github.io/func_SMILES2IMG/");
        sb.AppendLine(" License: MIT License");
        sb.AppendLine("================================================================================");
        sb.AppendLine();
        sb.AppendLine("Copyright (c) 2026 KimJuho");
        sb.AppendLine();
        sb.AppendLine("Permission is hereby granted, free of charge, to any person obtaining a copy");
        sb.AppendLine("of this software and associated documentation files (the \"Software\"), to deal");
        sb.AppendLine("in the Software without restriction, including without limitation the rights");
        sb.AppendLine("to use, copy, modify, merge, publish, distribute, sublicense, and/or sell");
        sb.AppendLine("copies of the Software, and to permit persons to whom the Software is");
        sb.AppendLine("furnished to do so, subject to the following conditions:");
        sb.AppendLine();
        sb.AppendLine("The above copyright notice and this permission notice shall be included in all");
        sb.AppendLine("copies or substantial portions of the Software.");
        sb.AppendLine();
        sb.AppendLine("THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR");
        sb.AppendLine("IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,");
        sb.AppendLine("FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE");
        sb.AppendLine("AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER");
        sb.AppendLine("LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,");
        sb.AppendLine("OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE");
        sb.AppendLine("SOFTWARE.");
        sb.AppendLine();
        sb.AppendLine("================================================================================");
        sb.AppendLine(" THIRD-PARTY COMPONENTS & OPEN-SOURCE NOTICES");
        sb.AppendLine("================================================================================");
        sb.AppendLine();

        foreach (var c in Components)
        {
            sb.AppendLine($"--- {c.Name} (v{c.Version}) ---");
            sb.AppendLine($"• License:     {c.License}");
            sb.AppendLine($"• Copyright:   {c.Copyright}");
            sb.AppendLine($"• Website:     {c.Website}");
            sb.AppendLine($"• Description: {c.Description}");
            sb.AppendLine();
        }

        sb.AppendLine("Notice regarding Microsoft Visual C++ Runtime:");
        sb.AppendLine("Indigo.Net includes Microsoft Visual C++ runtime DLLs (vcruntime140.dll, msvcp140.dll).");
        sb.AppendLine("Their respective Microsoft license terms apply; they are not relicensed under Apache-2.0.");
        sb.AppendLine();
        sb.AppendLine("Notice regarding IUPAC InChI library:");
        sb.AppendLine("The IUPAC International Chemical Identifier (InChI) library is provided under");
        sb.AppendLine("the IUPAC/InChI-Trust Licence #1.0.");

        return sb.ToString();
    }

    public static string GetProductVersion()
    {
        try
        {
            var asm = typeof(AboutDialog).Assembly;
            var infoVer = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrWhiteSpace(infoVer))
            {
                var v = infoVer!;
                var plusIdx = v.IndexOf('+');
                if (plusIdx > 0) v = v.Substring(0, plusIdx);
                return "v" + v;
            }
            var ver = asm.GetName().Version;
            return ver != null ? $"v{ver.Major}.{ver.Minor}.{ver.Build}" : "v1.6.0";
        }
        catch
        {
            return "v1.6.0";
        }
    }

    private static void OpenBrowser(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch
        {
            try { Process.Start(url); } catch { }
        }
    }

    public static void ShowModal(IWin32Window? owner = null)
    {
        try
        {
            using var dlg = new AboutDialog();
            if (owner != null)
            {
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.ShowDialog(owner);
            }
            else
            {
                dlg.StartPosition = FormStartPosition.CenterScreen;
                dlg.ShowDialog();
            }
        }
        catch
        {
            OpenBrowser("https://github.com/naramdash/func_SMILES2IMG");
        }
    }
}
