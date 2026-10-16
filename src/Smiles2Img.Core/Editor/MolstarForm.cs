using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Smiles2Img.Rendering;

namespace Smiles2Img;

public sealed class MolstarForm : Form
{
    private static CoreWebView2Environment? _sharedEnvironment;
    private static MolstarForm? _pooledInstance;

    private readonly WebView2 webView;
    private string contextTitle;
    private string initialStructure;
    private string? initialSnapshotJson;
    private bool initialized;
    private bool isWebViewReady;
    private bool needsReinitOnShow;
    private bool _forceDispose;

    public bool IsPooled { get; set; }
    public string ResultMode { get; private set; } = "image"; // "image" or "3dmodel"
    public string ResultPdbId { get; private set; } = "4HHB";
    public byte[]? ResultImageBytes { get; private set; }
    public byte[]? ResultGlbBytes { get; private set; }
    public string? ResultSnapshotJson { get; private set; }
    public bool ResultEmbedInDoc { get; private set; } = true;
    public string ResultBackground { get; private set; } = "transparent";

    public static MolstarForm Rent(string contextTitle, string initialStructure, string? initialSnapshotJson = null)
    {
        if (_pooledInstance != null && !_pooledInstance.IsDisposed)
        {
            _pooledInstance.Reconfigure(contextTitle, initialStructure, initialSnapshotJson);
            return _pooledInstance;
        }

        _pooledInstance = new MolstarForm(contextTitle, initialStructure, initialSnapshotJson)
        {
            IsPooled = true
        };
        return _pooledInstance;
    }

    public MolstarForm(string contextTitle, string initialStructure, string? initialSnapshotJson = null)
    {
        this.contextTitle = contextTitle ?? "PowerPoint Slide";
        this.initialStructure = initialStructure ?? "";
        this.initialSnapshotJson = initialSnapshotJson;

        Text = $"Mol* 3D Molecular Viewer - {this.contextTitle}";

        float dpiScale = 1.0f;
        try
        {
            using var g = CreateGraphics();
            dpiScale = Math.Max(1.0f, g.DpiX / 96.0f);
        }
        catch { }

        var screen = Screen.FromControl(this) ?? Screen.PrimaryScreen;
        int targetW = Math.Min((int)(1340 * dpiScale), (int)(screen.WorkingArea.Width * 0.94));
        int targetH = Math.Min((int)(880 * dpiScale), (int)(screen.WorkingArea.Height * 0.94));
        ClientSize = new Size(Math.Max(targetW, 1060), Math.Max(targetH, 720));
        MinimumSize = new Size(Math.Min(880, screen.WorkingArea.Width), Math.Min(600, screen.WorkingArea.Height));
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(15, 23, 42);

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

    public void Reconfigure(string newContextTitle, string newInitialStructure, string? newInitialSnapshotJson = null)
    {
        contextTitle = newContextTitle ?? "PowerPoint Slide";
        initialStructure = newInitialStructure ?? "";
        initialSnapshotJson = newInitialSnapshotJson;
        ResultMode = "image";
        ResultPdbId = "4HHB";
        ResultImageBytes = null;
        ResultGlbBytes = null;
        ResultSnapshotJson = null;
        ResultEmbedInDoc = true;
        ResultBackground = "transparent";
        DialogResult = DialogResult.None;
        Text = $"Mol* 3D Molecular Viewer - {contextTitle}";
        needsReinitOnShow = true;
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            NativeIndigo.EnsureWebView2Loader();
            var molstarFolder = MolstarAssetManager.EnsureMolstarFolder();

            var baseUserDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Smiles2Img", "WebView2UserData");
            Directory.CreateDirectory(baseUserDataFolder);

            string procName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            var procUserDataFolder = Path.Combine(baseUserDataFolder, procName + "_Molstar");
            Directory.CreateDirectory(procUserDataFolder);

            try
            {
                _sharedEnvironment ??= await CoreWebView2Environment.CreateAsync(null, procUserDataFolder, null);
                await webView.EnsureCoreWebView2Async(_sharedEnvironment);
            }
            catch (Exception ex) when (ex.HResult == unchecked((int)0x800700AA) || ex.Message.Contains("0x800700AA") || ex.Message.IndexOf("in use", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string fallbackFolder = Path.Combine(Path.GetTempPath(), "Smiles2Img_WV2_Mol_" + procName + "_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(fallbackFolder);
                _sharedEnvironment = await CoreWebView2Environment.CreateAsync(null, fallbackFolder, null);
                await webView.EnsureCoreWebView2Async(_sharedEnvironment);
            }

            webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;

            webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "molstar.local",
                molstarFolder,
                CoreWebView2HostResourceAccessKind.Allow);

            webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

            webView.CoreWebView2.Navigate("https://molstar.local/index.html");
        }
        catch (Exception ex)
        {
            _forceDispose = true;
            string msg = ex.Message;
            if (ex is DllNotFoundException || ex.GetType().Name.Contains("WebView2"))
            {
                msg += "\n\n(Please verify that the Microsoft Edge WebView2 Runtime is installed.)";
            }
            MessageBox.Show(
                this,
                "An error occurred while initializing Mol* 3D viewer:\n" + msg,
                "Mol* 3D Viewer Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            Close();
        }
    }

    private void SendInitScriptToWebView()
    {
        if (webView.CoreWebView2 == null) return;
        var escapedContext = EscapeJs(contextTitle);
        var escapedInitial = EscapeJs(initialStructure);
        var escapedSnapshot = EscapeJs(initialSnapshotJson);

        var script = $@"
            (async function() {{
                try {{
                    if (typeof viewer !== 'undefined' && viewer && viewer.plugin && viewer.plugin.clear) {{
                        await viewer.plugin.clear();
                    }}
                }} catch (e) {{}}
                if (window.initMolstar) {{
                    window.initMolstar('{escapedContext}', '{escapedInitial}', '{escapedSnapshot}');
                }}
            }})();
        ";
        webView.ExecuteScriptAsync(script);
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
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
                        if (dict.TryGetValue("mode", out var m) && m is string modeStr)
                            ResultMode = modeStr.Trim().ToLowerInvariant();

                        if (dict.TryGetValue("pdbId", out var p) && p is string pdbStr)
                            ResultPdbId = pdbStr.Trim();

                        if (dict.TryGetValue("imageBase64", out var img) && img is string imgBase64 && !string.IsNullOrWhiteSpace(imgBase64))
                        {
                            try { ResultImageBytes = Convert.FromBase64String(imgBase64); } catch { }
                        }

                        if (dict.TryGetValue("glbBase64", out var glb) && glb is string glbBase64 && !string.IsNullOrWhiteSpace(glbBase64))
                        {
                            try { ResultGlbBytes = Convert.FromBase64String(glbBase64); } catch { }
                        }

                        if (dict.TryGetValue("snapshotJson", out var snap) && snap is string snapStr)
                            ResultSnapshotJson = snapStr;

                        if (dict.TryGetValue("embedInDoc", out var emb) && emb is bool embBool)
                            ResultEmbedInDoc = embBool;

                        if (dict.TryGetValue("bgMode", out var bg) && bg is string bgStr)
                            ResultBackground = bgStr;
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
