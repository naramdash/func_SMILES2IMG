using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Lightweight, non-blocking floating toast positioned near the cursor using standard Windows colors,
/// clamped to the active monitor working area without stealing focus.
/// </summary>
public sealed class CopiedToast : Form
{
    private const int WS_EX_NOACTIVATE = 0x08000000;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_TOPMOST = 0x00000008;

    private readonly System.Windows.Forms.Timer _closeTimer;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST;
            return cp;
        }
    }

    protected override bool ShowWithoutActivation => true;

    private CopiedToast(string title, string value, bool isWarning = false)
    {
        float dpiScale = DialogHelper.GetDpiScale(this);

        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Font = SystemFonts.MessageBoxFont ?? new Font("Segoe UI", 9.0f);

        Color bgColor = isWarning ? Color.FromArgb(255, 244, 229) : SystemColors.Info;
        Color fgColor = SystemColors.InfoText;
        BackColor = SystemColors.ControlDark;
        Padding = new Padding(1);

        var inner = new Panel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = bgColor,
            Padding = new Padding((int)(12 * dpiScale), (int)(8 * dpiScale), (int)(12 * dpiScale), (int)(8 * dpiScale)),
            Dock = DockStyle.Fill
        };

        string header = isWarning ? title : $"Copied {title}";
        var lbl = new Label
        {
            Text = $"{header}: {Truncate(value, 60)}",
            UseMnemonic = false,
            ForeColor = fgColor,
            BackColor = bgColor,
            AutoSize = true,
            MaximumSize = new Size((int)(360 * dpiScale), 0),
            Dock = DockStyle.Fill
        };

        inner.Controls.Add(lbl);
        Controls.Add(inner);

        var preferred = GetPreferredSize(Size.Empty);
        var mouse = Cursor.Position;
        var area = Screen.FromPoint(mouse).WorkingArea;
        int margin = (int)(8 * dpiScale);
        int x = Math.Max(area.Left + margin, Math.Min(mouse.X + (int)(16 * dpiScale), area.Right - preferred.Width - margin));
        int y = Math.Max(area.Top + margin, Math.Min(mouse.Y + (int)(20 * dpiScale), area.Bottom - preferred.Height - margin));
        Location = new Point(x, y);

        _closeTimer = new System.Windows.Forms.Timer { Interval = isWarning ? 2000 : 1400 };
        _closeTimer.Tick += (_, _) =>
        {
            _closeTimer.Stop();
            Close();
        };
        _closeTimer.Start();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _closeTimer?.Dispose();
        }
        base.Dispose(disposing);
    }

    private static string Truncate(string text, int maxLen)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        text = text.Replace("\r", "").Replace("\n", " ").Trim();
        return text.Length <= maxLen ? text : text.Substring(0, maxLen - 3) + "...";
    }

    public static void Show(string title, string value, bool isWarning = false)
    {
        try
        {
            var thread = new Thread(() =>
            {
                try
                {
                    using var toast = new CopiedToast(title, value, isWarning);
                    Application.Run(toast);
                }
                catch { }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }
        catch { }
    }
}
