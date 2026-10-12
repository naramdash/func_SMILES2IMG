using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Smiles2Img.Rendering;

namespace Smiles2Img;

public sealed class Win32Window(IntPtr handle) : IWin32Window
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
    private readonly string initialStructure;
    private bool initialized;

    public string? ResultSmiles { get; private set; }
    public string? ResultKet { get; private set; }
    public string? ResultMolfile { get; private set; }
    public byte[]? ResultPngBytes { get; private set; }

    public KetcherForm(string cellAddress, string initialStructure)
    {
        this.cellAddress = cellAddress ?? "";
        this.initialStructure = initialStructure ?? "";

        string titleContext = this.cellAddress switch
        {
            "Word" => "Word 문서",
            "PowerPoint" => "PowerPoint 슬라이드",
            "" => "분자 구조식",
            _ => $"Excel [{this.cellAddress}]"
        };
        Text = $"EPAM Ketcher 분자 에디터 - {titleContext}";
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

            var baseUserDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Smiles2Img", "WebView2UserData");
            Directory.CreateDirectory(baseUserDataFolder);

            // Isolate per process (Excel / Word / PowerPoint) to prevent lock contention
            string procName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            var procUserDataFolder = Path.Combine(baseUserDataFolder, procName);
            Directory.CreateDirectory(procUserDataFolder);

            CoreWebView2Environment env;
            try
            {
                env = await CoreWebView2Environment.CreateAsync(null, procUserDataFolder, null);
                await webView.EnsureCoreWebView2Async(env);
            }
            catch (Exception ex) when (ex.HResult == unchecked((int)0x800700AA) || ex.Message.Contains("0x800700AA") || ex.Message.Contains("사용 중"))
            {
                // Fallback to a dedicated unique session directory if locked by an orphaned browser process
                string fallbackFolder = Path.Combine(Path.GetTempPath(), "Smiles2Img_WV2_" + procName + "_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(fallbackFolder);
                env = await CoreWebView2Environment.CreateAsync(null, fallbackFolder, null);
                await webView.EnsureCoreWebView2Async(env);
            }

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
            string msg = ex.Message;
            if (ex is DllNotFoundException || ex.GetType().Name.Contains("WebView2"))
            {
                msg += "\n\n(Microsoft Edge WebView2 런타임이 설치되어 있는지 확인해 주세요.)";
            }
            MessageBox.Show(
                this,
                "Ketcher 웹 에디터를 초기화하는 중 오류가 발생했습니다:\n" + msg,
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
                string buttonText = cellAddress switch
                {
                    "Word" => "문서에 적용 (Enter)",
                    "PowerPoint" => "슬라이드에 적용 (Enter)",
                    "Canvas" => "시트에 적용 (Enter)",
                    _ => "셀에 적용 (Enter)"
                };
                string badgeText = cellAddress switch
                {
                    "Word" => "Word 문서 (커서 위치)",
                    "PowerPoint" => "PowerPoint (현재 슬라이드)",
                    "Canvas" => "Excel 시트 (캔버스 형태)",
                    _ => string.IsNullOrEmpty(cellAddress) ? "분자 구조식" : $"셀: {cellAddress}"
                };

                var script = $@"
                    if (window.initEditor) {{
                        window.initEditor('{EscapeJs(cellAddress)}', '{EscapeJs(initialStructure)}');
                    }}
                    var b = document.getElementById('btn-apply');
                    if (b) b.innerText = '{EscapeJs(buttonText)}';
                    var ind = document.getElementById('cell-indicator');
                    if (ind) ind.innerText = '{EscapeJs(badgeText)}';
                ";
                webView.ExecuteScriptAsync(script);
            }
            else if (msg.StartsWith("{", StringComparison.Ordinal) && msg.Contains("\"action\":\"apply\""))
            {
                try
                {
                    var serializer = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                    var dict = serializer.Deserialize<Dictionary<string, object>>(msg);
                    if (dict != null)
                    {
                        if (dict.TryGetValue("smiles", out var s) && s is string smilesStr)
                            ResultSmiles = smilesStr.Trim();
                        if (dict.TryGetValue("ket", out var k) && k is string ketStr)
                            ResultKet = ketStr;
                        if (dict.TryGetValue("molfile", out var m) && m is string molStr)
                            ResultMolfile = molStr;
                        if (dict.TryGetValue("imageBase64", out var img) && img is string imgBase64 && !string.IsNullOrWhiteSpace(imgBase64))
                        {
                            try { ResultPngBytes = Convert.FromBase64String(imgBase64); } catch { }
                        }
                    }
                    DialogResult = DialogResult.OK;
                    Close();
                }
                catch
                {
                    DialogResult = DialogResult.Cancel;
                    Close();
                }
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
