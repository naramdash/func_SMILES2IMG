using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Smiles2Img.Notifications;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

/// <summary>
/// Standard Windows Forms dialog displaying molecular properties, 2D structure preview,
/// and one-click individual or spreadsheet-ready summary copy actions.
/// </summary>
public sealed class MoleculeInfoDialog : Form
{
    private static MoleculeSummaryFormat _lastUsedFormat = MoleculeSummaryFormat.ExcelTsvWithHeader;

    private static readonly (MoleculeSummaryFormat Format, string Label, string DisplayName)[] SummaryFormats =
    [
        (MoleculeSummaryFormat.ExcelTsvWithHeader, "Excel Table (With Header)", "Excel Table with Headers"),
        (MoleculeSummaryFormat.ExcelTsvRow, "Excel Row (TSV)", "Excel TSV Row"),
        (MoleculeSummaryFormat.MultiLine, "Multi-Line List", "Multi-Line List"),
        (MoleculeSummaryFormat.MarkdownTable, "Markdown Table", "Markdown Table"),
        (MoleculeSummaryFormat.Json, "JSON", "JSON"),
        (MoleculeSummaryFormat.CompactOneLine, "Compact Summary", "Compact Summary")
    ];

    private readonly MoleculeInfo _info;
    private readonly Label _lblStatus;
    private readonly Timer _rowResetTimer;
    private readonly ContextMenuStrip _rowMenu;
    private Image? _previewImage;
    private int _copiedRowIndex = -1;

    public MoleculeInfoDialog(MoleculeInfo info, string? title = null)
    {
        _info = info ?? throw new ArgumentNullException(nameof(info));

        bool isCompactEntity = _info.IsReaction || _info.IsMacromolecule;
        float dpiScale = DialogHelper.ConfigureStandardForm(
            this,
            title ?? "Molecule Info & Chemical Properties",
            UiIcons.MoleculeInfoIcon,
            baseWidth: 1040,
            baseHeight: isCompactEntity ? 420 : 620,
            minWidth: 680,
            minHeight: isCompactEntity ? 340 : 480);

        int pad = (int)(14 * dpiScale);
        int btnH = Math.Max(28, (int)(28 * dpiScale));

        string heading = !string.IsNullOrWhiteSpace(_info.SubstanceName)
            ? _info.SubstanceName
            : (!string.IsNullOrWhiteSpace(_info.Formula) ? _info.Formula : "Chemical Structure");
        string subheading = _info.IsReaction
            ? "Reaction Structure (Indigo Engine)"
            : (_info.IsMacromolecule
                ? $"{_info.IupacName}  ·  Macromolecular Complex"
                : (_info.HasQueryOrDummyAtoms
                    ? "Query / R-group Structure"
                    : $"{_info.Formula}  ·  MW {_info.MolecularWeight:F2} g/mol  ·  LogP {_info.LogP:F2}  ·  TPSA {_info.Tpsa:F1} Å²"));

        PictureBox? picPreview = null;
        _previewImage = TryRenderPreview(_info);
        if (_previewImage != null)
        {
            picPreview = new PictureBox
            {
                Image = _previewImage,
                Size = new Size((int)(112 * dpiScale), (int)(72 * dpiScale)),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Margin = new Padding((int)(12 * dpiScale), 0, 0, 0)
            };
        }

        var headerContainer = DialogHelper.CreateHeaderStrip(dpiScale, heading, subheading, rightVisual: picPreview);

        _lblStatus = new Label
        {
            Text = "Double-click any row, click Copy, or press Ctrl+C / Enter to copy a value.",
            UseMnemonic = false,
            ForeColor = SystemColors.GrayText,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            AutoEllipsis = true,
            Margin = new Padding(0, 0, (int)(10 * dpiScale), 0)
        };

        var actionFlow = new FlowLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Anchor = AnchorStyles.Right,
            Margin = Padding.Empty
        };

        var cboFormat = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = (int)(188 * dpiScale),
            Margin = new Padding(0, 1, (int)(6 * dpiScale), 0)
        };
        foreach (var (_, label, _) in SummaryFormats)
            cboFormat.Items.Add(label);
        cboFormat.SelectedIndex = Math.Max(0, Array.FindIndex(SummaryFormats, f => f.Format == _lastUsedFormat));
        cboFormat.SelectedIndexChanged += (_, _) =>
        {
            if (cboFormat.SelectedIndex >= 0 && cboFormat.SelectedIndex < SummaryFormats.Length)
                _lastUsedFormat = SummaryFormats[cboFormat.SelectedIndex].Format;
        };

        var btnCopyAll = DialogHelper.CreateStandardButton("Copy All", (int)(92 * dpiScale), btnH, new Padding(0, 0, (int)(8 * dpiScale), 0));
        var btnClose = DialogHelper.CreateStandardButton("Close", (int)(84 * dpiScale), btnH);
        btnClose.DialogResult = DialogResult.OK;
        btnClose.Click += (_, _) => Close();

        actionFlow.Controls.Add(cboFormat);
        actionFlow.Controls.Add(btnCopyAll);
        actionFlow.Controls.Add(btnClose);

        var bottomContainer = DialogHelper.CreateBottomActionBar(dpiScale, _lblStatus, actionFlow);

        // 3. Standard DataGridView (no owner-draw CellPainting, clean wrapped rows)
        var contentHost = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(pad, (int)(10 * dpiScale), pad, (int)(8 * dpiScale))
        };

        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.FixedSingle,
            RowHeadersVisible = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AllowUserToOrderColumns = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = false,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            ShowCellToolTips = true
        };
        dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(6, 5, 6, 5);

        var colIcon = new DataGridViewImageColumn
        {
            Name = "Icon",
            HeaderText = "",
            Width = Math.Max(34, (int)(34 * dpiScale)),
            ImageLayout = DataGridViewImageCellLayout.Normal,
            ReadOnly = true,
            Resizable = DataGridViewTriState.False,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colIcon.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

        var colProp = new DataGridViewTextBoxColumn
        {
            Name = "Property",
            HeaderText = "Property",
            Width = Math.Max(165, (int)(175 * dpiScale)),
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colProp.DefaultCellStyle.Padding = new Padding(6, 5, 6, 5);
        colProp.DefaultCellStyle.Font = new Font(Font, FontStyle.Bold);

        var colVal = new DataGridViewTextBoxColumn
        {
            Name = "Value",
            HeaderText = "Value",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colVal.DefaultCellStyle.Padding = new Padding(6, 5, 8, 5);
        colVal.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

        var colAction = new DataGridViewButtonColumn
        {
            Name = "Action",
            HeaderText = "",
            Width = Math.Max(76, (int)(78 * dpiScale)),
            UseColumnTextForButtonValue = false,
            FlatStyle = FlatStyle.Standard,
            Resizable = DataGridViewTriState.False,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        colAction.DefaultCellStyle.Padding = new Padding(4, 3, 4, 3);

        dgv.Columns.AddRange(colIcon, colProp, colVal, colAction);

        void AddPropertyRow(string propertyName, string displayValue, string copyPayload, bool isCode = false)
        {
            var icon = UiIcons.GetPropertyIcon(propertyName, dpiScale);
            bool hasPayload = !string.IsNullOrWhiteSpace(copyPayload);
            int rowIdx = dgv.Rows.Add(icon, propertyName, string.IsNullOrWhiteSpace(displayValue) ? "—" : displayValue, hasPayload ? "Copy" : "");
            var row = dgv.Rows[rowIdx];
            row.Tag = (PropertyName: propertyName, Payload: copyPayload);

            if (!hasPayload)
                row.Cells[colAction.Index] = new DataGridViewTextBoxCell { Value = "" };

            if (isCode)
                row.Cells[colVal.Index].Style.Font = new Font("Consolas", 8.5f, FontStyle.Regular);

            if (!string.IsNullOrWhiteSpace(displayValue) && displayValue != "—")
                row.Cells[colVal.Index].ToolTipText = $"{propertyName}: {displayValue}\nDouble-click or press Ctrl+C to copy";
        }

        string nameVal = !string.IsNullOrWhiteSpace(_info.CompoundName) ? _info.CompoundName : "(Unregistered or hypothetical compound)";
        AddPropertyRow("Substance Name", nameVal, _info.CompoundName);

        if (!string.IsNullOrWhiteSpace(_info.IupacName) && !string.Equals(_info.IupacName, _info.CompoundName, StringComparison.OrdinalIgnoreCase))
            AddPropertyRow("IUPAC Name", _info.IupacName, _info.IupacName);

        if (!_info.IsReaction)
        {
            if (_info.MolecularWeight > 0 || !_info.IsMacromolecule)
                AddPropertyRow("Molecular Weight", $"{_info.MolecularWeight:F2} g/mol", _info.MolecularWeight.ToString("F2"));
            AddPropertyRow("Chemical Formula", _info.Formula, _info.Formula);
            if (!_info.IsMacromolecule)
            {
                if (_info.ExactMass > 0)
                    AddPropertyRow("Monoisotopic Mass", $"{_info.ExactMass:F4}", _info.ExactMass.ToString("F4"));
                AddPropertyRow("LogP / TPSA", $"LogP: {_info.LogP:F2}   |   TPSA: {_info.Tpsa:F1} Å²", $"LogP: {_info.LogP:F2}, TPSA: {_info.Tpsa:F1}");
                AddPropertyRow("Rings & Aromaticity",
                    $"Rings: {_info.RingCount} (Aromatic: {_info.AromaticRingCount})   |   Heavy Atoms: {_info.HeavyAtomCount}",
                    $"Rings={_info.RingCount}, AromaticRings={_info.AromaticRingCount}, HeavyAtoms={_info.HeavyAtomCount}");
                AddPropertyRow("3D Complexity (Fsp3)",
                    $"Fsp3: {_info.Fsp3:F2}",
                    _info.Fsp3.ToString("F2"));
            }
        }

        AddPropertyRow("Canonical SMILES", _info.Smiles, _info.Smiles, isCode: true);
        if (!string.IsNullOrWhiteSpace(_info.InChI))
            AddPropertyRow("InChI", _info.InChI, _info.InChI, isCode: true);
        if (!string.IsNullOrWhiteSpace(_info.InChIKey))
            AddPropertyRow("InChIKey", _info.InChIKey, _info.InChIKey, isCode: true);

        _rowResetTimer = new Timer { Interval = 1200 };
        _rowResetTimer.Tick += (_, _) =>
        {
            _rowResetTimer.Stop();
            if (!dgv.IsDisposed && _copiedRowIndex >= 0 && _copiedRowIndex < dgv.Rows.Count)
            {
                if (dgv.Rows[_copiedRowIndex].Cells[colAction.Index] is DataGridViewButtonCell)
                    dgv.Rows[_copiedRowIndex].Cells[colAction.Index].Value = "Copy";
            }
            _copiedRowIndex = -1;
        };

        void ExecuteSummaryCopy(MoleculeSummaryFormat format)
        {
            _lastUsedFormat = format;
            int idx = Array.FindIndex(SummaryFormats, f => f.Format == format);
            if (idx >= 0 && cboFormat.SelectedIndex != idx)
                cboFormat.SelectedIndex = idx;

            string text = _info.FormatClipboardSummary(format);
            string? html = _info.FormatHtmlForSummary(format);
            try { SafeClipboard.SetTextAndHtml(text, html); } catch { }

            string displayName = idx >= 0 ? SummaryFormats[idx].DisplayName : "Summary";
            _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
            _lblStatus.Text = $"Copied {displayName} to clipboard.";
            MoleculeNotification.NotifyCopied($"Molecule Info ({displayName})", _info.SubstanceName ?? _info.Formula ?? "All properties");
        }

        void CopyRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dgv.Rows.Count) return;
            if (dgv.Rows[rowIndex].Tag is not (string propName, string payload) || string.IsNullOrEmpty(payload))
                return;

            try { SafeClipboard.SetText(payload); } catch { }

            if (_copiedRowIndex >= 0 && _copiedRowIndex < dgv.Rows.Count &&
                dgv.Rows[_copiedRowIndex].Cells[colAction.Index] is DataGridViewButtonCell)
            {
                dgv.Rows[_copiedRowIndex].Cells[colAction.Index].Value = "Copy";
            }

            if (dgv.Rows[rowIndex].Cells[colAction.Index] is DataGridViewButtonCell)
            {
                dgv.Rows[rowIndex].Cells[colAction.Index].Value = "Copied";
                _copiedRowIndex = rowIndex;
                _rowResetTimer.Stop();
                _rowResetTimer.Start();
            }

            _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
            _lblStatus.Text = $"Copied {propName} to clipboard.";
            MoleculeNotification.NotifyCopied($"Molecule Info ({propName})", payload);
        }

        btnCopyAll.Click += (_, _) => ExecuteSummaryCopy(_lastUsedFormat);

        dgv.CellContentClick += (_, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == colAction.Index)
                CopyRow(e.RowIndex);
        };
        dgv.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex >= 0)
                CopyRow(e.RowIndex);
        };
        dgv.KeyDown += (_, e) =>
        {
            if (((e.Control && e.KeyCode == Keys.C) || e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) &&
                dgv.CurrentRow != null)
            {
                CopyRow(dgv.CurrentRow.Index);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        };

        _rowMenu = new ContextMenuStrip();
        _rowMenu.Items.Add("Copy Value", null, (_, _) =>
        {
            if (dgv.CurrentRow != null) CopyRow(dgv.CurrentRow.Index);
        });
        _rowMenu.Items.Add("Copy Property & Value", null, (_, _) =>
        {
            if (dgv.CurrentRow?.Tag is (string propName, string payload) && !string.IsNullOrEmpty(payload))
            {
                string combined = $"{propName}: {payload}";
                try { SafeClipboard.SetText(combined); } catch { }
                _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
                _lblStatus.Text = $"Copied {propName} with label to clipboard.";
                MoleculeNotification.NotifyCopied("Molecule Property", combined);
            }
        });
        _rowMenu.Items.Add(new ToolStripSeparator());
        foreach (var (fmt, label, _) in SummaryFormats)
        {
            var captured = fmt;
            _rowMenu.Items.Add($"Copy All ({label})", null, (_, _) => ExecuteSummaryCopy(captured));
        }

        dgv.CellMouseDown += (_, e) =>
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
            {
                dgv.ClearSelection();
                dgv.Rows[e.RowIndex].Selected = true;
                dgv.CurrentCell = dgv.Rows[e.RowIndex].Cells[Math.Max(0, e.ColumnIndex)];
                _rowMenu.Show(Cursor.Position);
            }
        };

        contentHost.Controls.Add(dgv);

        Shown += (_, _) => dgv.ClearSelection();

        Controls.Add(contentHost);
        Controls.Add(bottomContainer);
        Controls.Add(headerContainer);

        CancelButton = btnClose;
    }

    private static Image? TryRenderPreview(MoleculeInfo info)
    {
        if (info.IsReaction || info.IsMacromolecule || string.IsNullOrWhiteSpace(info.Smiles))
            return null;

        try
        {
            var png = MoleculeRenderer.Render(info.Smiles, new RenderOptions { Background = "white" });
            if (png is not { Length: > 0 }) return null;
            using var ms = new MemoryStream(png);
            using var temp = Image.FromStream(ms);
            return new Bitmap(temp);
        }
        catch
        {
            return null;
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _rowResetTimer.Dispose();
            _rowMenu.Dispose();
            _previewImage?.Dispose();
            _previewImage = null;
        }
        base.Dispose(disposing);
    }

    public static void ShowModal(MoleculeInfo info, string? title = null, IWin32Window? owner = null)
    {
        if (info == null) return;
        DialogHelper.RunOnStaThread(owner, winOwner =>
        {
            using var dlg = new MoleculeInfoDialog(info, title);
            DialogHelper.ShowModalDialog(dlg, winOwner);
        });
    }
}
