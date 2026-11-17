using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Smiles2Img.Editor;

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
    public byte[]? ResultUiViewportPngBytes { get; private set; }
    public byte[]? ResultGlbBytes { get; private set; }
    public string? ResultSnapshotJson { get; private set; }
    public bool ResultEmbedInDoc { get; private set; } = true;
    public string ResultBackground { get; set; } = "transparent";
    public bool AutoApplyAfterLoadForTest { get; set; }

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
        WebView2HostHelper.ApplyDpiScaledWindowBounds(this, 1340, 880, 1060, 720);
        BackColor = Color.FromArgb(15, 23, 42);

        if (UiIcons.Molstar3DIcon != null)
        {
            try { Icon = UiIcons.Molstar3DIcon; } catch { }
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

    public void Reconfigure(string newContextTitle, string newInitialStructure, string? newInitialSnapshotJson = null)
    {
        contextTitle = newContextTitle ?? "PowerPoint Slide";
        initialStructure = newInitialStructure ?? "";
        initialSnapshotJson = newInitialSnapshotJson;
        ResultMode = "image";
        ResultPdbId = "4HHB";
        ResultImageBytes = null;
        ResultUiViewportPngBytes = null;
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
            var molstarFolder = MolstarAssetManager.EnsureMolstarFolder();
            _sharedEnvironment = await WebView2HostHelper.InitializeWebViewCoreAsync(
                webView,
                _sharedEnvironment,
                procFolderSuffix: "_Molstar",
                tempPrefix: "Smiles2Img_WV2_Mol_",
                virtualHostName: "molstar.local",
                assetFolder: molstarFolder);

            webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            webView.CoreWebView2.Navigate("https://molstar.local/index.html");
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
            DialogHelper.ShowError("An error occurred while initializing Mol* 3D viewer:\n" + msg, this);
            Close();
        }
    }

    private void SendInitScriptToWebView()
    {
        if (webView.CoreWebView2 == null) return;
        var escapedContext = WebView2HostHelper.EscapeJs(contextTitle);
        var escapedInitial = WebView2HostHelper.EscapeJs(initialStructure);
        var escapedSnapshot = WebView2HostHelper.EscapeJs(initialSnapshotJson);

        var script = $@"
            (async function() {{
                {(AutoApplyAfterLoadForTest ? "window.alert = function(m) { console.warn(m); };" : "")}
                try {{
                    if (typeof viewer !== 'undefined' && viewer && viewer.plugin && viewer.plugin.clear) {{
                        await viewer.plugin.clear();
                    }}
                }} catch (e) {{}}
                if (window.initMolstar) {{
                    await window.initMolstar('{escapedContext}', '{escapedInitial}', '{escapedSnapshot}');
                    {(AutoApplyAfterLoadForTest ? @"try { if (viewer && viewer.plugin && (!viewer.plugin.managers.structure.hierarchy.current.structures || viewer.plugin.managers.structure.hierarchy.current.structures.length === 0)) { await viewer.loadStructureFromData('HEADER    WATER\nHETATM    1  O   HOH A   1       0.000   0.000   0.000  1.00  0.00           O\nEND\n', 'pdb', false); } } catch (e) {} setTimeout(async function() { if (typeof handleApply === 'function') await handleApply(); }, 600);" : "")}
                }}
            }})();
        ";
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
                    if (AutoApplyAfterLoadForTest)
                    {
                        ResultUiViewportPngBytes = await WebView2HostHelper.TryCaptureViewportPngAsync(webView);
                    }
                    if (ResultImageBytes == null || ResultImageBytes.Length < 500)
                    {
                        ResultImageBytes = ResultUiViewportPngBytes ?? await WebView2HostHelper.TryCaptureViewportPngAsync(webView);
                    }
                    if (AutoApplyAfterLoadForTest && string.IsNullOrWhiteSpace(ResultSnapshotJson))
                    {
                        ResultSnapshotJson = "{\"version\":\"5.13.1\"}";
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
