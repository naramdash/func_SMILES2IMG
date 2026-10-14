using System;
using System.Drawing;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Lightweight, non-blocking, borderless toast window that confirms clipboard copy actions
/// without stealing focus or requiring user interaction.
/// </summary>
public sealed class CopiedToast : Form
{
    private const int WS_EX_NOACTIVATE = 0x08000000;
    private const int WS_EX_TOOLWINDOW = 0x00000080;
    private const int WS_EX_TOPMOST = 0x00000008;

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
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        BackColor = isWarning ? Color.FromArgb(130, 60, 20) : Color.FromArgb(20, 50, 85);
        Padding = new Padding(16, 10, 16, 10);

        string prefix = isWarning ? "⚠ " : "✓ Copied ";
        var lbl = new Label
        {
            Text = $"{prefix}{title}:\n{Truncate(value, 55)}",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = true,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
        Controls.Add(lbl);

        var mouse = Cursor.Position;
        var screen = Screen.FromPoint(mouse).WorkingArea;
        int x = Math.Min(mouse.X + 16, screen.Right - 320);
        int y = Math.Min(mouse.Y + 20, screen.Bottom - 80);
        Location = new Point(Math.Max(screen.Left + 10, x), Math.Max(screen.Top + 10, y));

        var timer = new Timer { Interval = isWarning ? 2000 : 1400 };
        timer.Tick += (s, e) =>
        {
            timer.Stop();
            timer.Dispose();
            Close();
            Dispose();
        };
        timer.Start();
    }

    private static string Truncate(string text, int maxLen)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        text = text.Replace("\r", "").Replace("\n", " ");
        if (text.Length <= maxLen) return text;
        return text.Substring(0, maxLen - 3) + "...";
    }

    public static void Show(string title, string value, bool isWarning = false)
    {
        try
        {
            var toast = new CopiedToast(title, value, isWarning);
            toast.Show();
        }
        catch { }
    }
}
