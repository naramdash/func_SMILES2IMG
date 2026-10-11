using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Smiles2Img.Rendering;

namespace Smiles2Img;

internal sealed class Win32Window(IntPtr handle) : IWin32Window
{
    public IntPtr Handle => handle;
}

public sealed class KetcherForm : Form
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool DestroyIcon(IntPtr handle);

    private static readonly Icon? AppIcon = CreateAppIcon();

    private readonly WebView2 webView;
    private readonly string cellAddress;
    private readonly string initialSmiles;
    private bool initialized;

    public string? ResultSmiles { get; private set; }

    public KetcherForm(string cellAddress, string initialSmiles)
    {
        this.cellAddress = cellAddress ?? "";
        this.initialSmiles = initialSmiles ?? "";

        Text = $"EPAM Ketcher 분자 에디터 - [{this.cellAddress}]";
        Width = 1040;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(820, 560);
        BackColor = Color.FromArgb(30, 41, 59);

        if (AppIcon != null)
        {
            try { Icon = AppIcon; } catch { }
        }

        webView = new WebView2
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(webView);

        Shown += async (s, e) =>
        {
            if (!initialized)
            {
                initialized = true;
                await InitializeWebViewAsync();
            }
        };
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            NativeIndigo.EnsureWebView2Loader();
            var ketcherFolder = KetcherAssetManager.EnsureKetcherFolder();

            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Smiles2Img", "WebView2UserData");
            Directory.CreateDirectory(userDataFolder);

            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
            await webView.EnsureCoreWebView2Async(env);

            webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;

            webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "ketcher.local",
                ketcherFolder,
                CoreWebView2HostResourceAccessKind.Allow);

            webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

            webView.CoreWebView2.Navigate("https://ketcher.local/index.html");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                "Ketcher 웹 에디터를 초기화하는 중 오류가 발생했습니다:\n" + ex.Message,
                "SMILES2IMG 에디터 오류",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            Close();
        }
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var msg = e.TryGetWebMessageAsString();
            if (string.IsNullOrEmpty(msg)) return;

            if (msg == "ready")
            {
                var script = $"window.initEditor('{EscapeJs(cellAddress)}', '{EscapeJs(initialSmiles)}');";
                webView.ExecuteScriptAsync(script);
            }
            else if (msg.StartsWith("apply:", StringComparison.Ordinal))
            {
                ResultSmiles = msg.Substring("apply:".Length);
                DialogResult = DialogResult.OK;
                Close();
            }
            else if (msg == "cancel")
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }
        catch { }
    }

    private static string EscapeJs(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s!.Replace("\\", "\\\\")
                .Replace("'", "\\'")
                .Replace("\"", "\\\"")
                .Replace("\r", "")
                .Replace("\n", "\\n");
    }

    private static Icon? CreateAppIcon()
    {
        try
        {
            using var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Sleek dark circular background
            using var bgBrush = new SolidBrush(Color.FromArgb(15, 23, 42));
            g.FillEllipse(bgBrush, 1, 1, 30, 30);

            // Cyan benzene ring (hexagon)
            using var ringPen = new Pen(Color.FromArgb(56, 189, 248), 2f);
            var points = new PointF[]
            {
                new(16, 5),
                new(25, 10.5f),
                new(25, 21.5f),
                new(16, 27),
                new(7, 21.5f),
                new(7, 10.5f)
            };
            g.DrawPolygon(ringPen, points);

            // Inner aromatic ring
            using var innerPen = new Pen(Color.FromArgb(14, 165, 233), 1.2f);
            g.DrawEllipse(innerPen, 11, 11, 10, 10);

            // Colored atom nodes (Oxygen red, Nitrogen blue, Emerald green)
            using var redBrush = new SolidBrush(Color.FromArgb(239, 68, 68));
            using var blueBrush = new SolidBrush(Color.FromArgb(59, 130, 246));
            using var greenBrush = new SolidBrush(Color.FromArgb(16, 185, 129));

            g.FillEllipse(redBrush, 13.5f, 2.5f, 5, 5);
            g.FillEllipse(blueBrush, 22.5f, 19, 5, 5);
            g.FillEllipse(greenBrush, 4.5f, 19, 5, 5);

            var hIcon = bmp.GetHicon();
            try
            {
                using var temp = Icon.FromHandle(hIcon);
                return (Icon)temp.Clone();
            }
            finally
            {
                DestroyIcon(hIcon);
            }
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
            webView?.Dispose();
        }
        base.Dispose(disposing);
    }
}
