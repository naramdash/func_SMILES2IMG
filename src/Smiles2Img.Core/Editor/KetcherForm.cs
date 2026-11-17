using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Smiles2Img.Editor;

namespace Smiles2Img;

public sealed class Win32Window(IntPtr handle) : IWin32Window
{
    public IntPtr Handle => handle;
}

public sealed class KetcherForm : Form
{
    private static readonly Icon? AppIcon = UiIcons.ChemicalIcon;
    private static CoreWebView2Environment? _sharedEnvironment;
    private static KetcherForm? _pooledInstance;

    private readonly WebView2 webView;
    private string cellAddress;
    private string initialStructure;
    private bool initialized;
    private bool isWebViewReady;
    private bool needsReinitOnShow;
    private bool _forceDispose;

    public bool IsPooled { get; set; }
    public string? ResultSmiles { get; private set; }
    public string? ResultKet { get; private set; }
    public string? ResultMolfile { get; private set; }
    public byte[]? ResultPngBytes { get; private set; }
    public byte[]? ResultUiViewportPngBytes { get; private set; }
    public string ResultBackground { get; private set; } = "transparent";
    public bool AutoApplyAfterLoadForTest { get; set; }
    public bool ToggleWhiteBackgroundForTest { get; set; }

    /// <summary>
    /// Rents a warm, reusable KetcherForm instance so subsequent editor launches open in &lt;50ms
    /// without re-initializing WebView2 or reloading the Ketcher React bundle.
    /// </summary>
    public static KetcherForm Rent(string cellAddress, string initialStructure)
    {
        if (_pooledInstance != null && !_pooledInstance.IsDisposed)
        {
            _pooledInstance.Reconfigure(cellAddress, initialStructure);
            return _pooledInstance;
        }

        _pooledInstance = new KetcherForm(cellAddress, initialStructure)
        {
            IsPooled = true
        };
        return _pooledInstance;
    }

    public KetcherForm(string cellAddress, string initialStructure)
    {
        this.cellAddress = cellAddress ?? "";
        this.initialStructure = initialStructure ?? "";

        UpdateWindowTitle();
        WebView2HostHelper.ApplyDpiScaledWindowBounds(this, 1320, 860, 1040, 700);
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

        VisibleChanged += async (s, e) =>
        {
            if (!Visible) return;
            if (!initialized)
            {
                initialized = true;
                await InitializeWebViewAsync();
            }
            else if (needsReinitOnShow && isWebViewReady)
            {
                needsReinitOnShow = false;
                SendInitScriptToWebView();
            }
        };
    }

    public void Reconfigure(string newCellAddress, string newInitialStructure)
    {
        cellAddress = newCellAddress ?? "";
        initialStructure = newInitialStructure ?? "";
        ResultSmiles = null;
        ResultKet = null;
        ResultMolfile = null;
        ResultPngBytes = null;
        ResultUiViewportPngBytes = null;
        ResultBackground = "transparent";
        DialogResult = DialogResult.None;
        UpdateWindowTitle();
        needsReinitOnShow = true;
    }

    private void UpdateWindowTitle()
    {
        string titleContext = cellAddress switch
        {
            "Word" => "Word Document",
            "PowerPoint" => "PowerPoint Slide",
            "" => "Chemical Structure",
            _ => $"Excel [{cellAddress}]"
        };
        Text = $"EPAM Ketcher Molecular Editor - {titleContext}";
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            var ketcherFolder = KetcherAssetManager.EnsureKetcherFolder();
            _sharedEnvironment = await WebView2HostHelper.InitializeWebViewCoreAsync(
                webView,
                _sharedEnvironment,
                procFolderSuffix: "",
                tempPrefix: "Smiles2Img_WV2_",
                virtualHostName: "ketcher.local",
                assetFolder: ketcherFolder);

            webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            webView.CoreWebView2.Navigate("https://ketcher.local/index.html");
        }
        catch (Exception ex)
        {
            _forceDispose = true;
            _pooledInstance = null;
            string msg = ex.Message;
            if (ex is DllNotFoundException || ex.GetType().Name.Contains("WebView2"))
            {
                msg += "\n\n(Please verify that the Microsoft Edge WebView2 Runtime is installed.)";
            }
            DialogHelper.ShowError("An error occurred while initializing Ketcher web editor:\n" + msg, this);
            Close();
        }
    }

    private void SendInitScriptToWebView()
    {
        if (webView.CoreWebView2 == null) return;

        string buttonText = cellAddress switch
        {
            "Word" => "Apply to Document (Ctrl+Enter)",
            "PowerPoint" => "Apply to Slide (Ctrl+Enter)",
            "Canvas" => "Apply to Sheet (Ctrl+Enter)",
            _ => "Apply to Cell (Ctrl+Enter)"
        };
        string badgeText = cellAddress switch
        {
            "Word" => "Word Document (Cursor)",
            "PowerPoint" => "PowerPoint (Active Slide)",
            "Canvas" => "Excel Sheet (Canvas)",
            _ => string.IsNullOrEmpty(cellAddress) ? "Chemical Structure" : $"Cell: {cellAddress}"
        };

        var clearOrLoadJs = string.IsNullOrWhiteSpace(initialStructure)
            ? "if (window.ketcher && window.ketcher.setMolecule) { window.ketcher.setMolecule(''); }"
            : "";

        var script = $@"
            currentBackground = 'transparent';
            var dot = document.getElementById('bg-dot'); if (dot) dot.className = 'bg-dot transparent';
            var lbl = document.getElementById('bg-label'); if (lbl) lbl.innerText = 'Background: Transparent';
            if (window.initEditor) {{
                window.initEditor('{WebView2HostHelper.EscapeJs(cellAddress)}', '{WebView2HostHelper.EscapeJs(initialStructure)}');
            }}
            {clearOrLoadJs}
            var b = document.getElementById('btn-apply');
            if (b) b.innerText = '{WebView2HostHelper.EscapeJs(buttonText)}';
            var ind = document.getElementById('cell-indicator');
            if (ind) ind.innerText = '{WebView2HostHelper.EscapeJs(badgeText)}';
        ";
        if (AutoApplyAfterLoadForTest)
        {
            var toggleJs = ToggleWhiteBackgroundForTest ? "if (window.toggleBackground) window.toggleBackground();" : "";
            script += $@"
                (function() {{
                    var attempts = 0;
                    var iv = setInterval(function() {{
                        attempts++;
                        if (window.ketcher && window.ketcher.editor && window.ketcher.editor.render && window.ketcher.editor.render.paper) {{
                            clearInterval(iv);
                            setTimeout(function() {{
                                {toggleJs}
                                if (window.onApply) window.onApply();
                            }}, 1000);
                        }} else if (attempts > 100) {{
                            clearInterval(iv);
                            if (window.onApply) window.onApply();
                        }}
                    }}, 100);
                }})();
            ";
        }
        webView.ExecuteScriptAsync(script);
    }

    private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var msg = e.TryGetWebMessageAsString();
            if (string.IsNullOrEmpty(msg)) return;

            if (msg == "ready")
            {
                isWebViewReady = true;
                needsReinitOnShow = false;
                SendInitScriptToWebView();
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
                        if (dict.TryGetValue("background", out var bg) && bg is string bgStr && !string.IsNullOrWhiteSpace(bgStr))
                            ResultBackground = bgStr.Trim();
                        if (dict.TryGetValue("imageBase64", out var img) && img is string imgBase64 && !string.IsNullOrWhiteSpace(imgBase64))
                        {
                            try { ResultPngBytes = Convert.FromBase64String(imgBase64); } catch { }
                        }
                    }
                    if (AutoApplyAfterLoadForTest)
                    {
                        ResultUiViewportPngBytes = await WebView2HostHelper.TryCaptureViewportPngAsync(webView);
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

    protected override void Dispose(bool disposing)
    {
        if (IsPooled && !_forceDispose)
        {
            return;
        }

        if (disposing)
        {
            webView?.Dispose();
        }
        base.Dispose(disposing);
    }
}
