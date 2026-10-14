using System;
using System.Drawing;
using System.Windows.Forms;
using Smiles2Img.Rendering;

namespace Smiles2Img.Editor;

/// <summary>
/// Modern Windows Forms dialog displaying chemical substance info, IUPAC name,
/// molecular weight, formula, exact mass, and SMILES with individual and all-in-one copy actions.
/// </summary>
public sealed class MoleculeInfoDialog : Form
{
    private readonly MoleculeInfo _info;
    private readonly Label _lblStatus;

    public MoleculeInfoDialog(MoleculeInfo info, string? title = null)
    {
        _info = info ?? throw new ArgumentNullException(nameof(info));

        Text = title ?? "Molecule Info & Chemical Properties";
        Width = 620;
        Height = 520;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = new Font("Segoe UI", 9.25f, FontStyle.Regular);
        BackColor = Color.FromArgb(248, 249, 250);

        // Header Panel
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 65,
            BackColor = Color.FromArgb(24, 76, 120),
            Padding = new Padding(18, 12, 18, 10)
        };

        var lblHeaderTitle = new Label
        {
            Text = "Molecular Properties & Substance Info",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(16, 10)
        };

        var lblHeaderSub = new Label
        {
            Text = "Calculated via Indigo Engine & Resolved via PubChem / PDB",
            ForeColor = Color.FromArgb(215, 230, 245),
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
            AutoSize = true,
            Location = new Point(17, 34)
        };

        pnlHeader.Controls.Add(lblHeaderTitle);
        pnlHeader.Controls.Add(lblHeaderSub);
        Controls.Add(pnlHeader);

        // Main content table
        var pnlContent = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20, 15, 20, 10),
            AutoScroll = true
        };

        int currentY = 15;
        void AddPropertyRow(string labelText, string valueText, string copyPayload, bool isBold = false)
        {
            var lbl = new Label
            {
                Text = labelText,
                Font = new Font("Segoe UI", 8.75f, isBold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 64, 67),
                Location = new Point(20, currentY + 3),
                Size = new Size(130, 22),
                TextAlign = ContentAlignment.MiddleRight
            };

            var txt = new TextBox
            {
                Text = string.IsNullOrWhiteSpace(valueText) ? "—" : valueText,
                Font = new Font("Segoe UI", 9.25f, isBold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = Color.FromArgb(32, 33, 36),
                BackColor = Color.White,
                ReadOnly = true,
                Location = new Point(155, currentY),
                Size = new Size(330, 26)
            };

            var btnCopy = new Button
            {
                Text = "Copy",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = Color.FromArgb(32, 33, 36),
                BackColor = Color.FromArgb(241, 243, 244),
                FlatStyle = FlatStyle.Flat,
                Location = new Point(492, currentY - 1),
                Size = new Size(82, 28),
                Cursor = Cursors.Hand
            };
            btnCopy.FlatAppearance.BorderColor = Color.FromArgb(218, 220, 224);

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

            pnlContent.Controls.Add(lbl);
            pnlContent.Controls.Add(txt);
            pnlContent.Controls.Add(btnCopy);

            currentY += 36;
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

        AddPropertyRow("Canonical SMILES:", _info.Smiles, _info.Smiles);
        if (!string.IsNullOrWhiteSpace(_info.InChI))
            AddPropertyRow("InChI:", _info.InChI, _info.InChI);
        if (!string.IsNullOrWhiteSpace(_info.InChIKey))
            AddPropertyRow("InChIKey:", _info.InChIKey, _info.InChIKey);

        Controls.Add(pnlContent);

        // Bottom Action Panel
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 65,
            BackColor = Color.FromArgb(241, 243, 244),
            Padding = new Padding(20, 12, 20, 12)
        };

        _lblStatus = new Label
        {
            Text = "✓ Full summary copied to clipboard.",
            Font = new Font("Segoe UI", 8.5f, FontStyle.Italic),
            ForeColor = Color.FromArgb(30, 115, 60),
            AutoSize = true,
            Location = new Point(20, 22)
        };

        var btnCopyAll = new Button
        {
            Text = "📋 Copy All Summary",
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(24, 76, 120),
            FlatStyle = FlatStyle.Flat,
            Size = new Size(160, 34),
            Location = new Point(310, 15),
            Cursor = Cursors.Hand
        };
        btnCopyAll.FlatAppearance.BorderSize = 0;
        btnCopyAll.Click += (s, e) =>
        {
            try { Clipboard.SetText(_info.FormatClipboardSummary()); } catch { }
            FlashCopied(btnCopyAll, "Full Summary");
            _lblStatus.Text = "✓ Full summary copied to clipboard.";
            _lblStatus.ForeColor = Color.FromArgb(30, 115, 60);
        };

        var btnClose = new Button
        {
            Text = "Close",
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            ForeColor = Color.FromArgb(60, 64, 67),
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Size = new Size(95, 34),
            Location = new Point(480, 15),
            Cursor = Cursors.Hand
        };
        btnClose.FlatAppearance.BorderColor = Color.FromArgb(218, 220, 224);
        btnClose.Click += (s, e) => Close();

        pnlBottom.Controls.Add(_lblStatus);
        pnlBottom.Controls.Add(btnCopyAll);
        pnlBottom.Controls.Add(btnClose);
        Controls.Add(pnlBottom);

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
        btn.BackColor = Color.FromArgb(230, 244, 234);

        if (_lblStatus != null)
        {
            _lblStatus.Text = $"✓ Copied {propName.TrimEnd(':')} to clipboard.";
            _lblStatus.ForeColor = Color.FromArgb(30, 115, 60);
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
            dlg.ShowDialog(owner);
        else
            dlg.ShowDialog();
    }
}
