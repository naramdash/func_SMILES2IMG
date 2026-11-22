using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Standard Windows Forms About &amp; License dialog displaying version info,
/// update check status, bundled open-source components, and full license notices.
/// </summary>
public sealed class AboutDialog : Form
{
    private sealed class ComponentInfo
    {
        public string Name { get; set; } = "";
        public string Version { get; set; } = "";
        public string License { get; set; } = "";
        public string Website { get; set; } = "";
        public string Copyright { get; set; } = "";
        public string Description { get; set; } = "";
    }

    private static readonly List<ComponentInfo> Components =
    [
        new()
        {
            Name = "EPAM Ketcher",
            Version = "3.18.0",
            License = "Apache License 2.0",
            Website = "https://github.com/epam/ketcher",
            Copyright = "Copyright (c) 2011-2024 EPAM Systems, Inc.",
            Description = "High-performance standalone molecular structure editor, chemical reaction canvas, and KET format serialization."
        },
        new()
        {
            Name = "EPAM Indigo Toolkit",
            Version = "1.49.0",
            License = "Apache License 2.0",
            Website = "https://github.com/epam/Indigo",
            Copyright = "Copyright (c) 2009-2024 EPAM Systems, Inc.",
            Description = "Cheminformatics engine providing molecular layout computation, aromatization, 0ms local property calculation, and high-resolution PNG rendering."
        },
        new()
        {
            Name = "Mol* (Molstar) 3D Viewer",
            Version = "5.13.1",
            License = "Apache License 2.0 / MIT",
            Website = "https://github.com/molstar/molstar",
            Copyright = "Copyright (c) 2018-2024 Mol* contributors",
            Description = "Comprehensive macromolecular 3D visualization engine for PDB and AlphaFold structures, supporting interactive 3D model (.glb) generation."
        },
        new()
        {
            Name = "IUPAC InChI Toolkit",
            Version = "1.07.1",
            License = "IUPAC / InChI-Trust Licence #1.0",
            Website = "https://www.inchi-trust.org",
            Copyright = "Copyright (c) IUPAC and InChI Trust Limited",
            Description = "Standard algorithm for International Chemical Identifier (InChI) conversion, structure normalization, and 27-character InChIKey hashing."
        },
        new()
        {
            Name = "Excel-DNA & IntelliSense",
            Version = "1.9.0",
            License = "zlib / MIT License",
            Website = "https://github.com/Excel-DNA/ExcelDna",
            Copyright = "Copyright (c) 2005-2024 Govert van Drimmelen",
            Description = "High-speed C-API add-in framework connecting .NET to Microsoft Excel, providing in-cell function auto-completion and argument tooltips."
        },
        new()
        {
            Name = "Microsoft.Web.WebView2",
            Version = "1.0.4314.35",
            License = "Microsoft Software License",
            Website = "https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            Copyright = "Copyright (c) Microsoft Corporation",
            Description = "Chromium-based native embedding runtime enabling EPAM Ketcher and Mol* canvases inside Office without external browser processes."
        },
        new()
        {
            Name = "System.Drawing.Common",
            Version = "4.7.2",
            License = "MIT License",
            Website = "https://github.com/dotnet/runtime",
            Copyright = "Copyright (c) .NET Foundation and Contributors",
            Description = "Cross-platform GDI+ graphics primitives for rasterization, font measurement, and in-memory image streaming."
        }
    ];

    private const string RepoUrl = "https://github.com/naramdash/func_SMILES2IMG";
    private const string DocsUrl = "https://naramdash.github.io/func_SMILES2IMG/";

    public string TargetDocsUrl { get; }

    private readonly Image? _logoBitmap;

    public static string GetDocsUrl(string? appName = null)
    {
        if (string.Equals(appName, "Excel", StringComparison.OrdinalIgnoreCase))
            return DocsUrl + "excel/";
        if (string.Equals(appName, "Word", StringComparison.OrdinalIgnoreCase))
            return DocsUrl + "word/";
        if (string.Equals(appName, "PowerPoint", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(appName, "Ppt", StringComparison.OrdinalIgnoreCase))
            return DocsUrl + "powerpoint/";
        return DocsUrl;
    }

    public AboutDialog(string? appName = null)
    {
        TargetDocsUrl = GetDocsUrl(appName);

        float dpiScale = DialogHelper.ConfigureStandardForm(
            this,
            "About func_SMILES2IMG",
            UiIcons.AboutIcon,
            baseWidth: 780,
            baseHeight: 650,
            minWidth: 620,
            minHeight: 480);

        int pad = (int)(14 * dpiScale);
        int btnH = Math.Max(28, (int)(28 * dpiScale));

        int iconSize = (int)(36 * dpiScale);
        _logoBitmap = (UiIcons.ChemicalIcon ?? UiIcons.AboutIcon)?.ToBitmap();
        var picLogo = new PictureBox
        {
            Size = new Size(iconSize, iconSize),
            SizeMode = PictureBoxSizeMode.Zoom,
            Image = _logoBitmap,
            Margin = new Padding(0, 2, (int)(12 * dpiScale), 0)
        };

        var headerContainer = DialogHelper.CreateHeaderStrip(
            dpiScale,
            "func_SMILES2IMG",
            $"{GetProductVersion()}  ·  MIT License  ·  Chemical Structure & 3D Add-in for Excel, Word & PowerPoint",
            leftVisual: picLogo);

        var leftActions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Margin = Padding.Empty
        };

        var btnCheckUpdate = DialogHelper.CreateStandardButton("Check for Updates", (int)(120 * dpiScale), btnH, new Padding(0, 0, (int)(10 * dpiScale), 0));
        var lblUpdateStatus = new Label
        {
            Text = "",
            UseMnemonic = false,
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Margin = new Padding(0, (int)(6 * dpiScale), 0, 0)
        };
        btnCheckUpdate.Click += async (_, _) => await CheckForUpdatesAsync(btnCheckUpdate, lblUpdateStatus);
        leftActions.Controls.Add(btnCheckUpdate);
        leftActions.Controls.Add(lblUpdateStatus);

        var rightActions = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Anchor = AnchorStyles.Right,
            Margin = Padding.Empty
        };

        var btnClose = DialogHelper.CreateStandardButton("Close", (int)(88 * dpiScale), btnH);
        btnClose.DialogResult = DialogResult.OK;
        btnClose.Click += (_, _) => Close();

        var btnOpenGitHub = DialogHelper.CreateStandardButton("Open GitHub", (int)(105 * dpiScale), btnH, new Padding(0, 0, (int)(8 * dpiScale), 0));
        btnOpenGitHub.Click += (_, _) => DialogHelper.OpenUrl(RepoUrl);

        var btnOpenDocs = DialogHelper.CreateStandardButton("Open Docs", (int)(100 * dpiScale), btnH, new Padding(0, 0, (int)(8 * dpiScale), 0));
        btnOpenDocs.Click += (_, _) => DialogHelper.OpenUrl(TargetDocsUrl);

        rightActions.Controls.Add(btnClose);
        rightActions.Controls.Add(btnOpenGitHub);
        rightActions.Controls.Add(btnOpenDocs);

        var bottomContainer = DialogHelper.CreateBottomActionBar(dpiScale, leftActions, rightActions);

        // Center native TabControl
        var contentHost = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(pad, (int)(10 * dpiScale), pad, (int)(8 * dpiScale))
        };

        var tabControl = new TabControl { Dock = DockStyle.Fill };

        var tabOverview = new TabPage("Overview & Author") { UseVisualStyleBackColor = true, Padding = new Padding(pad) };
        BuildOverviewTab(tabOverview, dpiScale, TargetDocsUrl);

        var tabComponents = new TabPage("Third-Party Components") { UseVisualStyleBackColor = true, Padding = new Padding(pad) };
        BuildComponentsTab(tabComponents, dpiScale);

        var tabNotices = new TabPage("License Notices (Full Text)") { UseVisualStyleBackColor = true, Padding = new Padding(pad) };
        BuildFullNoticesTab(tabNotices, dpiScale, btnH);

        tabControl.TabPages.AddRange(new[] { tabOverview, tabComponents, tabNotices });
        contentHost.Controls.Add(tabControl);

        Controls.Add(contentHost);
        Controls.Add(bottomContainer);
        Controls.Add(headerContainer);

        AcceptButton = btnClose;
        CancelButton = btnClose;
    }

    private async Task CheckForUpdatesAsync(Button btnCheckUpdate, Label lblStatus)
    {
        try
        {
            btnCheckUpdate.Enabled = false;
            lblStatus.ForeColor = SystemColors.GrayText;
            lblStatus.Text = "Checking GitHub Releases...";
            var currentVerStr = GetProductVersion();

            var latestTag = await Task.Run(FetchLatestGitHubReleaseTag);
            if (IsDisposed) return;

            btnCheckUpdate.Enabled = true;
            if (string.IsNullOrWhiteSpace(latestTag))
            {
                lblStatus.Text = "Could not reach GitHub Releases.";
                DialogHelper.ShowInfo("Could not reach GitHub Releases API. Opening releases page in browser.", Handle);
                DialogHelper.OpenUrl(RepoUrl + "/releases");
                return;
            }

            var cleanLatest = latestTag!.Trim().TrimStart('v', 'V');
            var cleanCurrent = currentVerStr.Trim().TrimStart('v', 'V');

            if (IsNewerVersion(cleanLatest, cleanCurrent))
            {
                lblStatus.ForeColor = Color.FromArgb(180, 83, 9);
                lblStatus.Text = $"Update available: v{cleanLatest}";

                const string updateCmd = "irm https://github.com/naramdash/func_SMILES2IMG/releases/latest/download/install.ps1 | iex";
                var dr = DialogHelper.ConfirmYesNo(
                    $"A new version (v{cleanLatest}) is available! (Current: v{cleanCurrent})\n\n" +
                    "Click [Yes] to copy the 1-line PowerShell update command to your clipboard and open the release page.",
                    this);

                if (dr == DialogResult.Yes)
                {
                    try { SafeClipboard.SetText(updateCmd); } catch { }
                    DialogHelper.OpenUrl(RepoUrl + "/releases/latest");
                }
            }
            else
            {
                lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
                lblStatus.Text = $"Up to date (v{cleanCurrent})";
            }
        }
        catch
        {
            if (!IsDisposed)
            {
                btnCheckUpdate.Enabled = true;
                lblStatus.Text = "";
            }
        }
    }

    private static void BuildOverviewTab(TabPage tab, float dpiScale, string docsUrl)
    {
        var metaGrid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            Margin = Padding.Empty,
            Padding = new Padding(2, 0, 2, (int)(8 * dpiScale))
        };
        metaGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(130 * dpiScale)));
        metaGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        void AddRow(string labelText, Control valueControl)
        {
            metaGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var lbl = new Label
            {
                Text = labelText,
                Font = new Font(SystemFonts.MessageBoxFont ?? Control.DefaultFont, FontStyle.Bold),
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 4, 8, 4)
            };
            valueControl.Anchor = AnchorStyles.Left;
            valueControl.Margin = new Padding(0, 4, 0, 4);
            metaGrid.Controls.Add(lbl);
            metaGrid.Controls.Add(valueControl);
        }

        static LinkLabel CreateLink(string url)
        {
            var link = new LinkLabel { Text = url, AutoSize = true, Cursor = Cursors.Hand };
            link.LinkClicked += (_, _) => DialogHelper.OpenUrl(url);
            return link;
        }

        AddRow("Project:", new Label { Text = "func_SMILES2IMG", AutoSize = true, UseMnemonic = false });
        AddRow("Version:", new Label { Text = $"{GetProductVersion()} (Excel · Word · PowerPoint)", AutoSize = true, UseMnemonic = false });
        AddRow("Author:", new Label { Text = "KimJuho (naramdash)", AutoSize = true, UseMnemonic = false });
        AddRow("License:", new Label { Text = "MIT License (Free & Open Source)", AutoSize = true, UseMnemonic = false });
        AddRow("Repository:", CreateLink(RepoUrl));
        AddRow("Documentation:", CreateLink(docsUrl));

        var grpCapabilities = new GroupBox
        {
            Text = "Capabilities && Offline Local Engines",
            Dock = DockStyle.Fill,
            Padding = new Padding((int)(10 * dpiScale), (int)(8 * dpiScale), (int)(10 * dpiScale), (int)(10 * dpiScale))
        };

        var (scrollHost, _) = DialogHelper.CreateScrollableWrappedLabel(
            "func_SMILES2IMG integrates chemical structure sketching, 40 worksheet formulas, and 3D macromolecular visualization directly inside Excel, Word, and PowerPoint.\r\n\r\n" +
            "• Microsoft Excel: In-cell (=SMILES2IMG, =INCHI2IMG, =KET2IMG, =NAME2IMG), floating (.FLOAT), 32 property/complexity/conversion/search UDFs (=SMILES2MW, =SMILES2LOGP, =SMILES2LIPINSKI, =SMILES2FSP3, =SMILES.MATCH), and 12-column SDF import/export.\r\n" +
            "• Microsoft Word: Inline and floating high-resolution structure images, lossless Vector SVG copy, 1-click Batch Style normalization, and live KET metadata round-trip editing.\r\n" +
            "• Microsoft PowerPoint: Slide shapes with KET metadata, 3D Mol* (.glb / 2D) models, Batch Style presets, and instant property export.\r\n" +
            "• Built-in Offline Engines: EPAM Ketcher, EPAM Indigo C++ Toolkit, Mol* 3D Viewer, and IUPAC InChI Library running locally without external servers.");

        grpCapabilities.Controls.Add(scrollHost);

        tab.Controls.Add(grpCapabilities);
        tab.Controls.Add(metaGrid);
    }

    private static void BuildComponentsTab(TabPage tab, float dpiScale)
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 58f));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 42f));

        // Standard DataGridView with AutoSizeColumnsMode.Fill avoids trailing blank space
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AllowUserToOrderColumns = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            BackgroundColor = SystemColors.Window,
            ScrollBars = ScrollBars.Vertical,
            ShowCellToolTips = true
        };
        grid.DefaultCellStyle.Padding = new Padding(4, 2, 4, 2);

        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Component", HeaderText = "Component", AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Version", HeaderText = "Version", AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "License", HeaderText = "License", AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Website", HeaderText = "Website (Double-click row to open)", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable });

        foreach (var c in Components)
        {
            int idx = grid.Rows.Add(c.Name, c.Version, c.License, c.Website);
            grid.Rows[idx].Tag = c;
        }

        // Standard GroupBox details pane (no fake TextBox caret or idle vertical scrollbar)
        var grpDetail = new GroupBox
        {
            Text = "Selected Component Details",
            Dock = DockStyle.Fill,
            Padding = new Padding((int)(10 * dpiScale), (int)(8 * dpiScale), (int)(10 * dpiScale), (int)(8 * dpiScale)),
            Margin = new Padding(0, (int)(8 * dpiScale), 0, 0)
        };

        var detailLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3
        };
        detailLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        detailLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var lblTitle = new Label
        {
            Font = new Font(SystemFonts.MessageBoxFont ?? Control.DefaultFont, FontStyle.Bold),
            AutoSize = true,
            UseMnemonic = false,
            Margin = new Padding(0, 0, 0, 4)
        };

        var linkWeb = new LinkLabel
        {
            AutoSize = true,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 0, 6)
        };
        linkWeb.LinkClicked += (_, _) =>
        {
            if (!string.IsNullOrWhiteSpace(linkWeb.Text))
                DialogHelper.OpenUrl(linkWeb.Text);
        };

        var (detailScroll, lblDetailDesc) = DialogHelper.CreateScrollableWrappedLabel();

        detailLayout.Controls.Add(lblTitle, 0, 0);
        detailLayout.Controls.Add(linkWeb, 0, 1);
        detailLayout.Controls.Add(detailScroll, 0, 2);
        grpDetail.Controls.Add(detailLayout);

        void UpdateDetail(ComponentInfo c)
        {
            lblTitle.Text = $"{c.Name} (v{c.Version}) — {c.License}";
            linkWeb.Text = c.Website;
            lblDetailDesc.Text = $"{c.Copyright}\r\n\r\n{c.Description}";
        }

        grid.SelectionChanged += (_, _) =>
        {
            if (grid.CurrentRow?.Tag is ComponentInfo ci)
                UpdateDetail(ci);
        };
        grid.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex >= 0 && grid.Rows[e.RowIndex].Tag is ComponentInfo ci && !string.IsNullOrEmpty(ci.Website))
                DialogHelper.OpenUrl(ci.Website);
        };

        if (grid.Rows.Count > 0)
        {
            grid.Rows[0].Selected = true;
            UpdateDetail(Components[0]);
        }

        layout.Controls.Add(grid, 0, 0);
        layout.Controls.Add(grpDetail, 0, 1);
        tab.Controls.Add(layout);
    }

    private static void BuildFullNoticesTab(TabPage tab, float dpiScale, int btnH)
    {
        var topBar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Padding = new Padding(0, 0, 0, (int)(8 * dpiScale))
        };

        var btnCopyNotices = DialogHelper.CreateStandardButton("Copy Full License Notices", (int)(160 * dpiScale), btnH, new Padding(0, 0, (int)(10 * dpiScale), 0));
        btnCopyNotices.TabIndex = 0;
        var lblCopyStatus = new Label
        {
            Text = "",
            UseMnemonic = false,
            AutoSize = true,
            ForeColor = Color.FromArgb(22, 101, 52),
            Margin = new Padding(0, (int)(6 * dpiScale), 0, 0)
        };

        btnCopyNotices.Click += (_, _) =>
        {
            try
            {
                SafeClipboard.SetText(BuildFullNoticesText());
                lblCopyStatus.Text = "Copied full notices to clipboard.";
            }
            catch { }
        };

        topBar.Controls.Add(btnCopyNotices);
        topBar.Controls.Add(lblCopyStatus);

        var txt = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            WordWrap = true,
            TabIndex = 1,
            ScrollBars = ScrollBars.Vertical,
            BackColor = SystemColors.Window,
            ForeColor = SystemColors.WindowText,
            Font = new Font("Consolas", 9.0f, FontStyle.Regular),
            Text = BuildFullNoticesText()
        };
        tab.Enter += (_, _) => txt.Select(0, 0);

        tab.Controls.Add(txt);
        tab.Controls.Add(topBar);
    }

    public static string BuildFullNoticesText()
    {
        var sb = new StringBuilder();
        sb.AppendLine("================================================================");
        sb.AppendLine(" func_SMILES2IMG (Excel · Word · PowerPoint)");
        sb.AppendLine($" Version: {GetProductVersion()}");
        sb.AppendLine(" Author: KimJuho (naramdash)");
        sb.AppendLine($" Repository: {RepoUrl}");
        sb.AppendLine($" Documentation: {DocsUrl}");
        sb.AppendLine(" License: MIT License");
        sb.AppendLine("================================================================");
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
        sb.AppendLine("================================================================");
        sb.AppendLine(" THIRD-PARTY COMPONENTS & OPEN-SOURCE NOTICES");
        sb.AppendLine("================================================================");
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
            return ver != null ? $"v{ver.Major}.{ver.Minor}.{ver.Build}" : "v1.8.1";
        }
        catch
        {
            return "v1.8.1";
        }
    }

    private static string? FetchLatestGitHubReleaseTag()
    {
        try
        {
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            var req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(
                "https://api.github.com/repos/naramdash/func_SMILES2IMG/releases/latest");
            req.Method = "GET";
            req.UserAgent = "func_SMILES2IMG";
            req.Accept = "application/vnd.github+json";
            req.Timeout = 4000;
            req.ReadWriteTimeout = 4000;

            using var resp = (System.Net.HttpWebResponse)req.GetResponse();
            using var stream = resp.GetResponseStream();
            if (stream == null) return null;
            using var reader = new System.IO.StreamReader(stream, Encoding.UTF8);
            var json = reader.ReadToEnd();
            var match = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value.Trim();
        }
        catch { }
        return null;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _logoBitmap?.Dispose();
        }
        base.Dispose(disposing);
    }

    public static void ShowModal(IWin32Window? owner = null, string? appName = null)
    {
        try
        {
            using var dlg = new AboutDialog(appName);
            DialogHelper.ShowModalDialog(dlg, owner);
        }
        catch
        {
            DialogHelper.OpenUrl(RepoUrl);
        }
    }

    public static bool IsNewerVersion(string? latestVersion, string? currentVersion)
    {
        if (string.IsNullOrWhiteSpace(latestVersion) || string.IsNullOrWhiteSpace(currentVersion))
            return false;

        SplitVersion(latestVersion!, out var latestBase, out var latestPre);
        SplitVersion(currentVersion!, out var currentBase, out var currentPre);

        // Only notify when a non-prerelease version supersedes current version
        if (!string.IsNullOrEmpty(latestPre))
            return false;

        int cmpBase = CompareNumericParts(latestBase, currentBase);
        if (cmpBase > 0) return true;
        if (cmpBase < 0) return false;

        return !string.IsNullOrEmpty(currentPre);
    }

    internal static void SplitVersion(string verStr, out string baseVer, out string preRelease)
    {
        var clean = verStr.Trim().TrimStart('v', 'V');
        var plusIdx = clean.IndexOf('+');
        if (plusIdx >= 0) clean = clean.Substring(0, plusIdx);

        var hyphenIdx = clean.IndexOf('-');
        if (hyphenIdx >= 0)
        {
            baseVer = clean.Substring(0, hyphenIdx).Trim();
            preRelease = clean.Substring(hyphenIdx + 1).Trim();
        }
        else
        {
            baseVer = clean.Trim();
            preRelease = string.Empty;
        }
    }

    internal static int CompareNumericParts(string a, string b)
    {
        var partsA = a.Split('.');
        var partsB = b.Split('.');
        int max = Math.Max(partsA.Length, partsB.Length);
        for (int i = 0; i < max; i++)
        {
            int numA = i < partsA.Length && int.TryParse(partsA[i], out var na) ? na : 0;
            int numB = i < partsB.Length && int.TryParse(partsB[i], out var nb) ? nb : 0;
            if (numA != numB) return numA.CompareTo(numB);
        }
        return 0;
    }
}
