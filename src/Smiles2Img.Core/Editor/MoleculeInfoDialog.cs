using System;
using System.Drawing;
using System.Windows.Forms;
using Smiles2Img.Notifications;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

/// <summary>
/// Modern, responsive, and resizable Windows Forms dialog displaying chemical substance info,
/// IUPAC name, molecular weight, formula, exact mass, SMILES, and InChI with individual
/// and all-in-one copy actions.
/// </summary>
public sealed class MoleculeInfoDialog : Form
{
    private static MoleculeSummaryFormat _lastUsedFormat = MoleculeSummaryFormat.ExcelTsvWithHeader;

    private readonly MoleculeInfo _info;
    private readonly Label _lblStatus;

    public MoleculeInfoDialog(MoleculeInfo info, string? title = null)
    {
        _info = info ?? throw new ArgumentNullException(nameof(info));

        Text = title ?? "Molecule Info & Chemical Properties";

        float dpiScale = 1.0f;
        try
        {
            using var g = CreateGraphics();
            dpiScale = Math.Max(1.0f, g.DpiX / 96.0f);
        }
        catch { }

        var screen = Screen.FromControl(this) ?? Screen.PrimaryScreen;
        int targetW = Math.Min((int)(940 * dpiScale), (int)(screen.WorkingArea.Width * 0.92));
        int targetH = Math.Min((int)(660 * dpiScale), (int)(screen.WorkingArea.Height * 0.92));
        ClientSize = new Size(Math.Max(targetW, 800), Math.Max(targetH, 560));
        MinimumSize = new Size(Math.Min(700, screen.WorkingArea.Width), Math.Min(500, screen.WorkingArea.Height));

        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = new Font("Segoe UI", 9.25f, FontStyle.Regular);
        BackColor = Color.FromArgb(249, 250, 251);

        if (UiIcons.MoleculeInfoIcon != null)
        {
            try { Icon = UiIcons.MoleculeInfoIcon; } catch { }
        }

        KeyPreview = true;
        KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        };

        var toolTip = new ToolTip
        {
            AutoPopDelay = 8000,
            InitialDelay = 400,
            ReshowDelay = 200,
            ShowAlways = true
        };

        // Header Panel: Auto-sized so both lines are 100% visible on any DPI scale without clipping
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.FromArgb(20, 50, 85),
            Padding = new Padding(24, 16, 20, 16)
        };

        var lblHeaderTitle = new Label
        {
            Text = "Molecular Properties & Substance Info",
            UseMnemonic = false,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 13f, FontStyle.Bold),
            AutoSize = true,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 0, 0, 4)
        };

        var lblHeaderSub = new Label
        {
            Text = "Calculated via Indigo Chemical Engine & Resolved via PubChem / RCSB PDB",
            UseMnemonic = false,
            ForeColor = Color.FromArgb(203, 225, 245),
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            AutoSize = true,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 4, 0, 0)
        };

        pnlHeader.Controls.Add(lblHeaderSub);
        pnlHeader.Controls.Add(lblHeaderTitle);

        // Bottom Action Panel
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = (int)Math.Max(64, 56 * dpiScale),
            BackColor = Color.FromArgb(243, 244, 246),
            Padding = new Padding(16, 12, 16, 12)
        };

        var sepBottom = new Panel
        {
            Dock = DockStyle.Top,
            Height = 1,
            BackColor = Color.FromArgb(229, 231, 235)
        };
        pnlBottom.Controls.Add(sepBottom);

        var pnlButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            WrapContents = false
        };

        var btnClose = new Button
        {
            Text = "Close",
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular),
            ForeColor = Color.FromArgb(55, 65, 81),
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Size = new Size((int)(95 * dpiScale), 36),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 2, 0, 0)
        };
        btnClose.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
        btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(243, 244, 246);
        btnClose.Click += (s, e) => Close();
        toolTip.SetToolTip(btnClose, "Close this dialog (Esc)");

        _lblStatus = new Label
        {
            Text = "💡 Double-click any row, click 'Copy', or press Ctrl+C to copy. Use 'Copy All' for spreadsheets.",
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            ForeColor = Color.FromArgb(75, 85, 99),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            AutoEllipsis = true,
            Padding = new Padding(6, 0, 10, 0)
        };

        var pnlSplitCopy = new TableLayoutPanel
        {
            RowCount = 1,
            ColumnCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 2, 12, 0),
            Padding = Padding.Empty,
            BackColor = Color.FromArgb(24, 76, 120)
        };
        pnlSplitCopy.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        pnlSplitCopy.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        pnlSplitCopy.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Math.Max(26f, 24f * dpiScale)));

        var btnCopyAll = new Button
        {
            Text = GetFormatButtonLabel(_lastUsedFormat),
            Font = new Font("Segoe UI", 9.25f, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(24, 76, 120),
            FlatStyle = FlatStyle.Flat,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(12, 6, 8, 6),
            Cursor = Cursors.Hand,
            Margin = Padding.Empty,
            Dock = DockStyle.Fill
        };
        btnCopyAll.FlatAppearance.BorderSize = 0;
        btnCopyAll.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 95, 150);
        toolTip.SetToolTip(btnCopyAll, "Copy complete molecule properties to clipboard");

        var btnCopyMenu = new Button
        {
            Text = "▾",
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(24, 76, 120),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Margin = Padding.Empty,
            Dock = DockStyle.Fill
        };
        btnCopyMenu.FlatAppearance.BorderSize = 0;
        btnCopyMenu.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 95, 150);
        toolTip.SetToolTip(btnCopyMenu, "Choose summary export format (Excel TSV, Multi-Line, Markdown, JSON)");

        pnlSplitCopy.Paint += (s, e) =>
        {
            try
            {
                using var pen = new Pen(Color.FromArgb(40, 110, 170), 1f);
                int x = btnCopyMenu.Left;
                e.Graphics.DrawLine(pen, x, 4, x, pnlSplitCopy.Height - 4);
            }
            catch { }
        };

        void ExecuteCopy(MoleculeSummaryFormat format)
        {
            _lastUsedFormat = format;
            btnCopyAll.Text = GetFormatButtonLabel(format);

            string text = _info.FormatClipboardSummary(format);
            string? html = _info.FormatHtmlForSummary(format);

            try { SafeClipboard.SetTextAndHtml(text, html); } catch { }

            string displayName = format switch
            {
                MoleculeSummaryFormat.ExcelTsvRow => "Excel TSV Row",
                MoleculeSummaryFormat.ExcelTsvWithHeader => "Excel Table with Headers",
                MoleculeSummaryFormat.MultiLine => "Multi-Line List",
                MoleculeSummaryFormat.MarkdownTable => "Markdown Table",
                MoleculeSummaryFormat.Json => "JSON",
                _ => "Compact Summary"
            };

            FlashCopied(btnCopyAll, displayName);
            _lblStatus.Text = $"✓ Copied full info as {displayName} to clipboard.";
            _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
            MoleculeNotification.NotifyCopied($"Molecule Info ({displayName})", _info.SubstanceName ?? _info.Formula ?? "All properties");
        }

        btnCopyAll.Click += (s, e) => ExecuteCopy(_lastUsedFormat);

        btnCopyMenu.Click += (s, e) =>
        {
            var menu = new ContextMenuStrip { Font = new Font("Segoe UI", 9.25f), ShowImageMargin = false };

            void AddMenuItem(string label, MoleculeSummaryFormat fmt)
            {
                var item = new ToolStripMenuItem(label, null, (ms, me) => ExecuteCopy(fmt))
                {
                    Checked = (_lastUsedFormat == fmt)
                };
                menu.Items.Add(item);
            }

            AddMenuItem("📋 Excel Row (TSV for Spreadsheets)", MoleculeSummaryFormat.ExcelTsvRow);
            AddMenuItem("📊 Excel Table (With Header Row)", MoleculeSummaryFormat.ExcelTsvWithHeader);
            AddMenuItem("📑 Multi-Line List (Rich Table for Word/Slides)", MoleculeSummaryFormat.MultiLine);
            AddMenuItem("📐 Markdown Table (Notion / GitHub / Obsidian)", MoleculeSummaryFormat.MarkdownTable);
            AddMenuItem("💻 JSON (for Python / Data Scripts)", MoleculeSummaryFormat.Json);
            menu.Items.Add(new ToolStripSeparator());
            AddMenuItem("💬 Compact Single-Line Summary", MoleculeSummaryFormat.CompactOneLine);

            menu.Show(pnlSplitCopy, new Point(0, pnlSplitCopy.Height));
        };

        pnlSplitCopy.Controls.Add(btnCopyAll, 0, 0);
        pnlSplitCopy.Controls.Add(btnCopyMenu, 1, 0);

        pnlButtons.Controls.Add(btnClose);
        pnlButtons.Controls.Add(pnlSplitCopy);

        pnlBottom.Controls.Add(_lblStatus);
        pnlBottom.Controls.Add(pnlButtons);

        // Main content panel containing modern, high-DPI aware DataGridView
        var pnlContent = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(249, 250, 251),
            Padding = new Padding((int)(16 * dpiScale), (int)(12 * dpiScale), (int)(16 * dpiScale), (int)(12 * dpiScale))
        };

        var pnlGridContainer = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(229, 231, 235), // Clean 1px border
            Padding = new Padding(1)
        };

        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            GridColor = Color.FromArgb(238, 240, 243),
            RowHeadersVisible = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AllowUserToResizeColumns = true,
            AllowUserToOrderColumns = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            EnableHeadersVisualStyles = false,
            AutoGenerateColumns = false,
            ShowCellToolTips = true,
            Font = new Font("Segoe UI", 9.25f, FontStyle.Regular)
        };

        try
        {
            typeof(DataGridView)
                .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
                .SetValue(dgv, true, null);
        }
        catch { }

        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        dgv.ColumnHeadersHeight = (int)Math.Max(38, 36 * dpiScale);
        dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(243, 244, 246),
            ForeColor = Color.FromArgb(55, 65, 81),
            Font = new Font("Segoe UI", 9.25f, FontStyle.Bold),
            Alignment = DataGridViewContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 0, 0),
            SelectionBackColor = Color.FromArgb(243, 244, 246),
            SelectionForeColor = Color.FromArgb(55, 65, 81)
        };

        var colProp = new DataGridViewTextBoxColumn
        {
            Name = "Property",
            HeaderText = "Property / Field",
            Width = (int)Math.Max(210, 185 * dpiScale),
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colProp.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        colProp.DefaultCellStyle.Padding = new Padding(14, 0, 8, 0);
        colProp.DefaultCellStyle.Font = new Font("Segoe UI", 9.25f, FontStyle.Bold);
        colProp.DefaultCellStyle.ForeColor = Color.FromArgb(55, 65, 81);
        colProp.DefaultCellStyle.BackColor = Color.FromArgb(250, 251, 252);
        colProp.DefaultCellStyle.SelectionBackColor = Color.FromArgb(239, 246, 255);
        colProp.DefaultCellStyle.SelectionForeColor = Color.FromArgb(30, 58, 138);

        var colVal = new DataGridViewTextBoxColumn
        {
            Name = "Value",
            HeaderText = "Value",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colVal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        colVal.DefaultCellStyle.Padding = new Padding(12, 0, 10, 0);
        colVal.DefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
        colVal.DefaultCellStyle.ForeColor = Color.FromArgb(17, 24, 39);
        colVal.DefaultCellStyle.SelectionBackColor = Color.FromArgb(239, 246, 255);
        colVal.DefaultCellStyle.SelectionForeColor = Color.FromArgb(17, 24, 39);

        var colAction = new DataGridViewButtonColumn
        {
            Name = "Action",
            HeaderText = "",
            Width = (int)Math.Max(88, 80 * dpiScale),
            FlatStyle = FlatStyle.Flat,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colAction.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        colAction.DefaultCellStyle.Padding = new Padding(6, 4, 6, 4);
        colAction.DefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Regular);
        colAction.DefaultCellStyle.ForeColor = Color.FromArgb(31, 41, 55);
        colAction.DefaultCellStyle.BackColor = Color.FromArgb(243, 244, 246);
        colAction.DefaultCellStyle.SelectionBackColor = Color.FromArgb(239, 246, 255);
        colAction.DefaultCellStyle.SelectionForeColor = Color.FromArgb(31, 41, 55);

        dgv.Columns.Add(colProp);
        dgv.Columns.Add(colVal);
        dgv.Columns.Add(colAction);

        dgv.RowTemplate.Height = (int)Math.Max(38, 36 * dpiScale);

        void AddPropertyRow(string propertyName, string displayValue, string copyPayload, bool isCode = false)
        {
            int rowIdx = dgv.Rows.Add(propertyName, string.IsNullOrWhiteSpace(displayValue) ? "—" : displayValue, "Copy");
            var row = dgv.Rows[rowIdx];
            row.Tag = (PropertyName: propertyName, Payload: copyPayload);

            if (isCode)
            {
                row.Cells[1].Style.Font = new Font("Consolas", 9.5f, FontStyle.Regular);
            }

            if (!string.IsNullOrWhiteSpace(displayValue) && displayValue != "—")
            {
                row.Cells[1].ToolTipText = $"{propertyName}: {displayValue}\n(Double-click or click 'Copy' to copy to clipboard)";
            }

            if (string.IsNullOrWhiteSpace(copyPayload))
            {
                row.Cells[2].Value = "";
            }
        }

        string nameVal = !string.IsNullOrWhiteSpace(_info.CompoundName) ? _info.CompoundName : "(Unregistered or hypothetical compound)";
        string nameCopy = !string.IsNullOrWhiteSpace(_info.CompoundName) ? _info.CompoundName : "";
        AddPropertyRow("Substance Name", nameVal, nameCopy);

        if (!string.IsNullOrWhiteSpace(_info.IupacName) && !string.Equals(_info.IupacName, _info.CompoundName, StringComparison.OrdinalIgnoreCase))
        {
            AddPropertyRow("IUPAC Name", _info.IupacName, _info.IupacName);
        }

        if (!_info.IsReaction)
        {
            AddPropertyRow("Molecular Weight", $"{_info.MolecularWeight:F2} g/mol", _info.MolecularWeight.ToString("F2"));
            AddPropertyRow("Chemical Formula", _info.Formula, _info.Formula);
            if (_info.ExactMass > 0)
                AddPropertyRow("Monoisotopic Mass", $"{_info.ExactMass:F4}", _info.ExactMass.ToString("F4"));
            AddPropertyRow("LogP / TPSA", $"LogP: {_info.LogP:F2}   |   TPSA: {_info.Tpsa:F1} Å²", $"LogP: {_info.LogP:F2}, TPSA: {_info.Tpsa:F1}");
            var lipinskiStatus = _info.PassesLipinski ? "Pass" : "Violation";
            AddPropertyRow("Lipinski Rule of 5", $"{lipinskiStatus} (HBD: {_info.HBondDonors}, HBA: {_info.HBondAcceptors}, RotBonds: {_info.RotatableBonds})", $"{lipinskiStatus} (HBD={_info.HBondDonors}, HBA={_info.HBondAcceptors}, RotBonds={_info.RotatableBonds})");
        }

        AddPropertyRow("Canonical SMILES", _info.Smiles, _info.Smiles, isCode: true);
        if (!string.IsNullOrWhiteSpace(_info.InChI))
            AddPropertyRow("InChI", _info.InChI, _info.InChI, isCode: true);
        if (!string.IsNullOrWhiteSpace(_info.InChIKey))
            AddPropertyRow("InChIKey", _info.InChIKey, _info.InChIKey, isCode: true);

        void CopyRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dgv.Rows.Count) return;
            var row = dgv.Rows[rowIndex];
            if (row.Tag is (string propName, string payload) && !string.IsNullOrEmpty(payload))
            {
                try { SafeClipboard.SetText(payload); } catch { }
                row.Cells[2].Value = "✓ Copied";
                dgv.InvalidateCell(2, rowIndex);
                _lblStatus.Text = $"✓ Copied {propName} to clipboard.";
                _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
                MoleculeNotification.NotifyCopied($"Molecule Info ({propName})", payload);

                var timer = new Timer { Interval = 1200 };
                timer.Tick += (s, e) =>
                {
                    try
                    {
                        timer.Stop();
                        timer.Dispose();
                        if (!dgv.IsDisposed && rowIndex < dgv.Rows.Count)
                        {
                            row.Cells[2].Value = "Copy";
                            dgv.InvalidateCell(2, rowIndex);
                        }
                    }
                    catch { }
                };
                timer.Start();
            }
        }

        dgv.CellContentClick += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 2)
            {
                CopyRow(e.RowIndex);
            }
        };

        dgv.CellDoubleClick += (s, e) =>
        {
            if (e.RowIndex >= 0)
            {
                CopyRow(e.RowIndex);
            }
        };

        dgv.KeyDown += (s, e) =>
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                if (dgv.CurrentRow != null)
                {
                    CopyRow(dgv.CurrentRow.Index);
                    e.Handled = true;
                }
            }
            else if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                if (dgv.CurrentRow != null)
                {
                    CopyRow(dgv.CurrentRow.Index);
                    e.Handled = true;
                }
            }
        };

        dgv.CellPainting += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 2)
            {
                e.PaintBackground(e.CellBounds, true);
                string btnText = e.Value?.ToString() ?? "Copy";
                if (!string.IsNullOrEmpty(btnText))
                {
                    int padX = (int)Math.Max(8, 6 * dpiScale);
                    int padY = (int)Math.Max(5, 4 * dpiScale);
                    var btnRect = new Rectangle(
                        e.CellBounds.X + padX,
                        e.CellBounds.Y + padY,
                        e.CellBounds.Width - padX * 2,
                        e.CellBounds.Height - padY * 2);

                    bool isCopied = btnText.Contains("✓");
                    Color btnBg = isCopied ? Color.FromArgb(220, 252, 231) : Color.FromArgb(243, 244, 246);
                    Color btnBorder = isCopied ? Color.FromArgb(187, 247, 208) : Color.FromArgb(209, 213, 219);
                    Color btnFg = isCopied ? Color.FromArgb(22, 101, 52) : Color.FromArgb(55, 65, 81);

                    using (var brush = new SolidBrush(btnBg))
                        e.Graphics.FillRectangle(brush, btnRect);

                    using (var pen = new Pen(btnBorder, 1f))
                        e.Graphics.DrawRectangle(pen, btnRect);

                    using (var font = new Font("Segoe UI", 8.75f, isCopied ? FontStyle.Bold : FontStyle.Regular))
                    {
                        TextRenderer.DrawText(
                            e.Graphics,
                            btnText,
                            font,
                            btnRect,
                            btnFg,
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                    }
                }
                e.Handled = true;
            }
        };

        dgv.CellMouseEnter += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 2)
            {
                dgv.Cursor = Cursors.Hand;
            }
        };

        dgv.CellMouseLeave += (s, e) =>
        {
            dgv.Cursor = Cursors.Default;
        };

        var rowContextMenu = new ContextMenuStrip { Font = new Font("Segoe UI", 9.25f), ShowImageMargin = false };
        rowContextMenu.Items.Add(new ToolStripMenuItem("📋 Copy Value", null, (s, e) =>
        {
            if (dgv.CurrentRow?.Index is int idx && idx >= 0) CopyRow(idx);
        }));
        rowContextMenu.Items.Add(new ToolStripMenuItem("🏷️ Copy Property & Value", null, (s, e) =>
        {
            if (dgv.CurrentRow?.Index is int idx && idx >= 0)
            {
                var row = dgv.Rows[idx];
                if (row.Tag is (string propName, string payload))
                {
                    string combined = $"{propName}: {payload}";
                    try { SafeClipboard.SetText(combined); } catch { }
                    _lblStatus.Text = $"✓ Copied '{combined}' to clipboard.";
                    _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
                    MoleculeNotification.NotifyCopied("Molecule Property", combined);
                }
            }
        }));
        rowContextMenu.Items.Add(new ToolStripSeparator());
        rowContextMenu.Items.Add(new ToolStripMenuItem("📊 Copy All (Excel Table)", null, (s, e) => ExecuteCopy(MoleculeSummaryFormat.ExcelTsvWithHeader)));
        rowContextMenu.Items.Add(new ToolStripMenuItem("📋 Copy All (Excel TSV Row)", null, (s, e) => ExecuteCopy(MoleculeSummaryFormat.ExcelTsvRow)));
        rowContextMenu.Items.Add(new ToolStripMenuItem("📑 Copy All (Multi-Line)", null, (s, e) => ExecuteCopy(MoleculeSummaryFormat.MultiLine)));
        rowContextMenu.Items.Add(new ToolStripMenuItem("📐 Copy All (Markdown Table)", null, (s, e) => ExecuteCopy(MoleculeSummaryFormat.MarkdownTable)));
        rowContextMenu.Items.Add(new ToolStripMenuItem("💻 Copy All (JSON)", null, (s, e) => ExecuteCopy(MoleculeSummaryFormat.Json)));

        dgv.CellMouseDown += (s, e) =>
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
            {
                dgv.ClearSelection();
                dgv.Rows[e.RowIndex].Selected = true;
                dgv.CurrentCell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex >= 0 ? e.ColumnIndex : 0];
                rowContextMenu.Show(Cursor.Position);
            }
        };

        pnlGridContainer.Controls.Add(dgv);
        pnlContent.Controls.Add(pnlGridContainer);

        Controls.Add(pnlContent);
        Controls.Add(pnlBottom);
        Controls.Add(pnlHeader);
        pnlContent.BringToFront();

        AcceptButton = btnCopyAll;
        CancelButton = btnClose;
    }

    private static string GetFormatButtonLabel(MoleculeSummaryFormat format) => format switch
    {
        MoleculeSummaryFormat.ExcelTsvRow => "Copy All (Excel TSV)",
        MoleculeSummaryFormat.ExcelTsvWithHeader => "Copy All (Excel Table)",
        MoleculeSummaryFormat.MultiLine => "Copy All (Multi-Line)",
        MoleculeSummaryFormat.MarkdownTable => "Copy All (Markdown)",
        MoleculeSummaryFormat.Json => "Copy All (JSON)",
        _ => "Copy All Summary"
    };

    private void FlashCopied(Button btn, string propName)
    {
        var origText = btn.Text;
        var origColor = btn.BackColor;
        var origForeColor = btn.ForeColor;

        btn.Text = "✓ Copied!";
        if (origForeColor == Color.White)
        {
            btn.BackColor = Color.FromArgb(22, 101, 52);
        }
        else
        {
            btn.BackColor = Color.FromArgb(220, 252, 231);
            btn.ForeColor = Color.FromArgb(22, 101, 52);
        }

        if (_lblStatus != null)
        {
            _lblStatus.Text = $"✓ Copied {propName.TrimEnd(':')} to clipboard.";
            _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
        }

        var t = new Timer { Interval = 1200 };
        t.Tick += (s, e) =>
        {
            try
            {
                t.Stop();
                t.Dispose();
                if (!btn.IsDisposed)
                {
                    btn.Text = origText;
                    btn.BackColor = origColor;
                    btn.ForeColor = origForeColor;
                }
            }
            catch { }
        };
        t.Start();
    }

    public static void ShowModal(MoleculeInfo info, string? title = null, IWin32Window? owner = null)
    {
        if (info == null) return;

        if (System.Threading.Thread.CurrentThread.GetApartmentState() == System.Threading.ApartmentState.STA)
        {
            ShowModalInternal(info, title, owner);
        }
        else
        {
            var hwnd = owner?.Handle ?? IntPtr.Zero;
            var t = new System.Threading.Thread(() =>
            {
                IWin32Window? winOwner = hwnd != IntPtr.Zero ? new Win32Window(hwnd) : null;
                ShowModalInternal(info, title, winOwner);
            });
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
            t.Join();
        }
    }

    private static void ShowModalInternal(MoleculeInfo info, string? title, IWin32Window? owner)
    {
        using var dlg = new MoleculeInfoDialog(info, title);
        if (owner != null && owner.Handle != IntPtr.Zero)
        {
            dlg.StartPosition = FormStartPosition.CenterParent;
            try
            {
                dlg.ShowDialog(owner);
                return;
            }
            catch
            {
                // Fallback to CenterScreen if owner window is invalid or disconnected
            }
        }
        dlg.StartPosition = FormStartPosition.CenterScreen;
        dlg.ShowDialog();
    }
}
