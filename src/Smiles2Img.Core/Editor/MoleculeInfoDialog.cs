using System;
using System.Drawing;
using System.Windows.Forms;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

/// <summary>
/// Modern, responsive, and resizable Windows Forms dialog displaying chemical substance info,
/// IUPAC name, molecular weight, formula, exact mass, SMILES, and InChI with individual
/// and all-in-one copy actions.
/// </summary>
public sealed class MoleculeInfoDialog : Form
{
    private readonly MoleculeInfo _info;
    private readonly Label _lblStatus;

    public MoleculeInfoDialog(MoleculeInfo info, string? title = null)
    {
        _info = info ?? throw new ArgumentNullException(nameof(info));

        Text = title ?? "Molecule Info & Chemical Properties";
        ClientSize = new Size(820, 580);
        MinimumSize = new Size(680, 480);
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

        // Header Panel
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 72,
            BackColor = Color.FromArgb(20, 50, 85),
            Padding = new Padding(24, 12, 20, 10)
        };

        var lblHeaderTitle = new Label
        {
            Text = "Molecular Properties & Substance Info",
            UseMnemonic = false,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 12.5f, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(22, 12)
        };

        var lblHeaderSub = new Label
        {
            Text = "Calculated via Indigo Chemical Engine & Resolved via PubChem / RCSB PDB",
            UseMnemonic = false,
            ForeColor = Color.FromArgb(203, 225, 245),
            Font = new Font("Segoe UI", 8.75f, FontStyle.Regular),
            AutoSize = true,
            Location = new Point(23, 40)
        };

        pnlHeader.Controls.Add(lblHeaderTitle);
        pnlHeader.Controls.Add(lblHeaderSub);

        // Bottom Action Panel
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 64,
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
            Size = new Size(95, 36),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 2, 0, 0)
        };
        btnClose.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
        btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(243, 244, 246);
        btnClose.Click += (s, e) => Close();

        _lblStatus = new Label
        {
            Text = "✓ Full summary automatically copied to clipboard. Press Esc to close.",
            Font = new Font("Segoe UI", 9f, FontStyle.Italic),
            ForeColor = Color.FromArgb(22, 101, 52),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            AutoEllipsis = true,
            Padding = new Padding(6, 0, 10, 0)
        };

        var btnCopyAll = new Button
        {
            Text = "📋 Copy All Summary",
            Font = new Font("Segoe UI", 9.25f, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(24, 76, 120),
            FlatStyle = FlatStyle.Flat,
            Size = new Size(185, 36),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 2, 12, 0)
        };
        btnCopyAll.FlatAppearance.BorderSize = 0;
        btnCopyAll.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 95, 150);
        btnCopyAll.Click += (s, e) =>
        {
            try { Clipboard.SetText(_info.FormatClipboardSummary()); } catch { }
            FlashCopied(btnCopyAll, "Full Summary");
            _lblStatus.Text = "✓ Full summary copied to clipboard.";
            _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
        };

        pnlButtons.Controls.Add(btnClose);
        pnlButtons.Controls.Add(btnCopyAll);

        pnlBottom.Controls.Add(_lblStatus);
        pnlBottom.Controls.Add(pnlButtons);

        // Main content panel containing responsive TableLayoutPanel
        var pnlContent = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(249, 250, 251),
            Padding = new Padding(12, 12, 12, 12)
        };

        var tblLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 3,
            Padding = new Padding(10, 8, 10, 8)
        };

        tblLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 185f));
        tblLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        tblLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95f));

        int rowIndex = 0;
        void AddPropertyRow(string labelText, string valueText, string copyPayload, bool isBold = false, bool isCode = false)
        {
            tblLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40f));

            var lbl = new Label
            {
                Text = labelText,
                UseMnemonic = false,
                Font = new Font("Segoe UI", 9.25f, isBold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = Color.FromArgb(55, 65, 81),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(0, 3, 12, 3),
                AutoEllipsis = true
            };

            var txt = new TextBox
            {
                Text = string.IsNullOrWhiteSpace(valueText) ? "—" : valueText,
                Font = isCode
                    ? new Font("Consolas", 9.75f, FontStyle.Regular)
                    : new Font("Segoe UI", 9.5f, isBold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = Color.FromArgb(17, 24, 39),
                BackColor = Color.White,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 5, 8, 5),
                BorderStyle = BorderStyle.FixedSingle
            };
            txt.DoubleClick += (s, e) => txt.SelectAll();
            if (!string.IsNullOrWhiteSpace(valueText) && valueText != "—")
            {
                toolTip.SetToolTip(txt, valueText);
            }

            var btnCopy = new Button
            {
                Text = "Copy",
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                ForeColor = Color.FromArgb(31, 41, 55),
                BackColor = Color.FromArgb(243, 244, 246),
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 0, 4),
                Height = 30,
                Cursor = Cursors.Hand
            };
            btnCopy.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btnCopy.FlatAppearance.MouseOverBackColor = Color.FromArgb(229, 231, 235);

            btnCopy.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(copyPayload))
                {
                    try { Clipboard.SetText(copyPayload); } catch { }
                    FlashCopied(btnCopy, labelText);
                }
            };

            if (string.IsNullOrWhiteSpace(copyPayload))
                btnCopy.Enabled = false;

            tblLayout.Controls.Add(lbl, 0, rowIndex);
            tblLayout.Controls.Add(txt, 1, rowIndex);
            tblLayout.Controls.Add(btnCopy, 2, rowIndex);
            rowIndex++;
        }

        string nameVal = !string.IsNullOrWhiteSpace(_info.CompoundName) ? _info.CompoundName : "(Unregistered or hypothetical compound)";
        string nameCopy = !string.IsNullOrWhiteSpace(_info.CompoundName) ? _info.CompoundName : "";
        AddPropertyRow("Substance Name:", nameVal, nameCopy, isBold: true);

        if (!string.IsNullOrWhiteSpace(_info.IupacName) && !string.Equals(_info.IupacName, _info.CompoundName, StringComparison.OrdinalIgnoreCase))
        {
            AddPropertyRow("IUPAC Name:", _info.IupacName, _info.IupacName);
        }

        if (!_info.IsReaction)
        {
            AddPropertyRow("Molecular Weight:", $"{_info.MolecularWeight:F2} g/mol", _info.MolecularWeight.ToString("F2"), isBold: true);
            AddPropertyRow("Chemical Formula:", _info.Formula, _info.Formula, isBold: true);
            if (_info.ExactMass > 0)
                AddPropertyRow("Monoisotopic Mass:", $"{_info.ExactMass:F4}", _info.ExactMass.ToString("F4"));
        }

        AddPropertyRow("Canonical SMILES:", _info.Smiles, _info.Smiles, isCode: true);
        if (!string.IsNullOrWhiteSpace(_info.InChI))
            AddPropertyRow("InChI:", _info.InChI, _info.InChI, isCode: true);
        if (!string.IsNullOrWhiteSpace(_info.InChIKey))
            AddPropertyRow("InChIKey:", _info.InChIKey, _info.InChIKey, isCode: true);

        pnlContent.Controls.Add(tblLayout);

        // WinForms docking order: Top and Bottom first, then Fill
        Controls.Add(pnlContent);
        Controls.Add(pnlBottom);
        Controls.Add(pnlHeader);
        pnlContent.BringToFront();

        AcceptButton = btnCopyAll;
        CancelButton = btnClose;

        // Auto-copy full summary on launch
        try { Clipboard.SetText(_info.FormatClipboardSummary()); } catch { }
    }

    private void FlashCopied(Button btn, string propName)
    {
        var origText = btn.Text;
        var origColor = btn.BackColor;
        btn.Text = "✓ Copied!";
        btn.BackColor = Color.FromArgb(220, 252, 231);

        if (_lblStatus != null)
        {
            _lblStatus.Text = $"✓ Copied {propName.TrimEnd(':')} to clipboard.";
            _lblStatus.ForeColor = Color.FromArgb(22, 101, 52);
        }

        var t = new Timer { Interval = 1200 };
        t.Tick += (s, e) =>
        {
            t.Stop();
            t.Dispose();
            if (!btn.IsDisposed)
            {
                btn.Text = origText;
                btn.BackColor = origColor;
            }
        };
        t.Start();
    }

    public static void ShowModal(MoleculeInfo info, string? title = null, IWin32Window? owner = null)
    {
        using var dlg = new MoleculeInfoDialog(info, title);
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
}
