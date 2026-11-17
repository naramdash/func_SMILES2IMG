using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Shared Windows Forms dialog configuration, owner-centered modal display,
/// DPI scaling, and standard MessageBox helpers across Excel, Word, and PowerPoint.
/// </summary>
public static class DialogHelper
{
    public const string AppTitle = "func_SMILES2IMG";

    public static float GetDpiScale(Control control)
    {
        try
        {
            using var g = control.CreateGraphics();
            return Math.Max(1.0f, g.DpiX / 96.0f);
        }
        catch
        {
            return 1.0f;
        }
    }

    public static IWin32Window? ToOwner(IntPtr hwnd) =>
        hwnd != IntPtr.Zero ? new Win32Window(hwnd) : null;

    public static float ConfigureStandardForm(
        Form form,
        string title,
        Icon? icon,
        int baseWidth,
        int baseHeight,
        int minWidth,
        int minHeight,
        bool sizable = true,
        double maxScreenFraction = 0.92)
    {
        float dpiScale = GetDpiScale(form);
        var screen = Screen.FromControl(form) ?? Screen.PrimaryScreen;
        var area = screen.WorkingArea;

        int targetW = Math.Min((int)(baseWidth * dpiScale), (int)(area.Width * maxScreenFraction));
        int targetH = Math.Min((int)(baseHeight * dpiScale), (int)(area.Height * maxScreenFraction));

        form.Text = title;
        form.Font = SystemFonts.MessageBoxFont ?? new Font("Segoe UI", 9.0f);
        form.BackColor = SystemColors.Control;
        form.ForeColor = SystemColors.ControlText;
        form.ClientSize = new Size(Math.Max(targetW, Math.Min(minWidth, area.Width)), Math.Max(targetH, Math.Min(minHeight, area.Height)));
        form.MinimumSize = new Size(Math.Min(minWidth, area.Width), Math.Min(minHeight, area.Height));
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = sizable ? FormBorderStyle.Sizable : FormBorderStyle.FixedDialog;
        form.MaximizeBox = sizable;
        form.MinimizeBox = false;
        form.ShowInTaskbar = false;
        form.KeyPreview = true;
        form.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                form.Close();
        };

        if (icon != null)
        {
            try { form.Icon = icon; } catch { }
        }

        return dpiScale;
    }

    public static Button CreateStandardButton(string text, int minWidth, int minHeight, Padding? margin = null) => new()
    {
        Text = text,
        UseVisualStyleBackColor = true,
        AutoSize = true,
        MinimumSize = new Size(minWidth, minHeight),
        Margin = margin ?? Padding.Empty
    };

    public static Panel CreateSeparator() => new()
    {
        Dock = DockStyle.Top,
        Height = 1,
        BackColor = SystemColors.ControlLight
    };

    public static Panel CreateHeaderStrip(
        float dpiScale,
        string title,
        string subtitle,
        Control? leftVisual = null,
        Control? rightVisual = null)
    {
        int pad = (int)(14 * dpiScale);
        var container = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = SystemColors.Window
        };

        int cols = 1 + (leftVisual != null ? 1 : 0) + (rightVisual != null ? 1 : 0);
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = cols,
            RowCount = 1,
            Padding = new Padding(pad, (int)(11 * dpiScale), pad, (int)(11 * dpiScale)),
            BackColor = SystemColors.Window
        };

        var titleStack = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            RowCount = 2,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        titleStack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        titleStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        titleStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var baseFont = SystemFonts.MessageBoxFont ?? Control.DefaultFont;
        var titleFont = new Font(baseFont.FontFamily, 12f, FontStyle.Bold);
        int titleH = TextRenderer.MeasureText("Ag", titleFont).Height + 4;
        int subH = TextRenderer.MeasureText("Ag", baseFont).Height + 4;

        titleStack.Controls.Add(new Label
        {
            Text = title,
            UseMnemonic = false,
            Font = titleFont,
            ForeColor = SystemColors.ControlText,
            AutoSize = false,
            AutoEllipsis = true,
            Dock = DockStyle.Top,
            Height = titleH,
            TextAlign = ContentAlignment.MiddleLeft,
            Margin = new Padding(0, 0, 0, 2)
        }, 0, 0);
        titleStack.Controls.Add(new Label
        {
            Text = subtitle,
            UseMnemonic = false,
            Font = baseFont,
            ForeColor = SystemColors.GrayText,
            AutoSize = false,
            AutoEllipsis = true,
            Dock = DockStyle.Top,
            Height = subH,
            TextAlign = ContentAlignment.MiddleLeft,
            Margin = Padding.Empty
        }, 0, 1);

        int col = 0;
        if (leftVisual != null)
        {
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.Controls.Add(leftVisual, col++, 0);
        }
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.Controls.Add(titleStack, col++, 0);
        if (rightVisual != null)
        {
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.Controls.Add(rightVisual, col, 0);
        }

        container.Controls.Add(layout);
        container.Controls.Add(CreateSeparator());
        return container;
    }

    public static Panel CreateBottomActionBar(float dpiScale, Control leftControl, Control rightControl)
    {
        int pad = (int)(14 * dpiScale);
        var container = new Panel
        {
            Dock = DockStyle.Bottom,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = SystemColors.Control
        };

        var bar = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(pad, (int)(10 * dpiScale), pad, (int)(10 * dpiScale))
        };
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.Controls.Add(leftControl, 0, 0);
        bar.Controls.Add(rightControl, 1, 0);

        container.Controls.Add(bar);
        container.Controls.Add(CreateSeparator());
        return container;
    }

    public static (Panel Container, Label Label) CreateScrollableWrappedLabel(string initialText = "")
    {
        var scrollHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var label = new Label
        {
            AutoSize = true,
            UseMnemonic = false,
            Location = Point.Empty,
            Text = initialText
        };
        void SyncWrap()
        {
            int maxW = Math.Max(200, scrollHost.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 8);
            label.MaximumSize = new Size(maxW, 0);
        }
        scrollHost.SizeChanged += (_, _) => SyncWrap();
        scrollHost.HandleCreated += (_, _) => SyncWrap();
        scrollHost.Controls.Add(label);
        return (scrollHost, label);
    }

    public static DialogResult ShowModalDialog(Form form, IntPtr ownerHwnd = default) =>
        ShowModalDialog(form, ToOwner(ownerHwnd));

    public static DialogResult ShowModalDialog(Form form, IWin32Window? owner)
    {
        if (owner != null && owner.Handle != IntPtr.Zero)
        {
            form.StartPosition = FormStartPosition.CenterParent;
            try
            {
                return form.ShowDialog(owner);
            }
            catch
            {
                // Fallback to CenterScreen if owner handle is disconnected across Office threads
            }
        }

        form.StartPosition = FormStartPosition.CenterScreen;
        return form.ShowDialog();
    }

    public static DialogResult ShowCommonDialog(CommonDialog dialog, IntPtr ownerHwnd = default)
    {
        var owner = ToOwner(ownerHwnd);
        if (owner != null)
        {
            try { return dialog.ShowDialog(owner); }
            catch { }
        }
        return dialog.ShowDialog();
    }

    public static void RunOnStaThread(IWin32Window? owner, Action<IWin32Window?> action)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
        {
            action(owner);
            return;
        }

        var hwnd = owner?.Handle ?? IntPtr.Zero;
        var thread = new Thread(() => action(ToOwner(hwnd)));
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        thread.Join();
    }

    public static void ShowInfo(string message, IntPtr ownerHwnd = default) =>
        ShowMessage(message, ToOwner(ownerHwnd), MessageBoxIcon.Information);

    public static void ShowWarning(string message, IntPtr ownerHwnd = default) =>
        ShowMessage(message, ToOwner(ownerHwnd), MessageBoxIcon.Warning);

    public static void ShowWarning(string message, IWin32Window? owner) =>
        ShowMessage(message, owner, MessageBoxIcon.Warning);

    public static void ShowError(string message, IWin32Window? owner = null) =>
        ShowMessage(message, owner, MessageBoxIcon.Error);

    public static DialogResult ConfirmYesNo(string message, IWin32Window? owner = null)
    {
        if (owner != null && owner.Handle != IntPtr.Zero)
        {
            try { return MessageBox.Show(owner, message, AppTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information); }
            catch { }
        }
        return MessageBox.Show(message, AppTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information);
    }

    private static void ShowMessage(string message, IWin32Window? owner, MessageBoxIcon icon)
    {
        if (owner != null && owner.Handle != IntPtr.Zero)
        {
            try
            {
                MessageBox.Show(owner, message, AppTitle, MessageBoxButtons.OK, icon);
                return;
            }
            catch { }
        }
        MessageBox.Show(message, AppTitle, MessageBoxButtons.OK, icon);
    }

    public static void ShowAbout(IntPtr ownerHwnd = default)
    {
        try
        {
            AboutDialog.ShowModal(ToOwner(ownerHwnd));
        }
        catch
        {
            AboutDialog.ShowModal();
        }
    }

    public static void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
        }
        catch
        {
            try { Process.Start(url); } catch { }
        }
    }
}
