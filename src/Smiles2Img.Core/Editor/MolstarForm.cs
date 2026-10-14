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
    private readonly WebView2 webView;
    private readonly string contextTitle;
    private readonly string initialStructure;
    private readonly string? initialSnapshotJson;
    private bool initialized;

    public string ResultMode { get; private set; } = "image"; // "image" or "3dmodel"
    public string ResultPdbId { get; private set; } = "4HHB";
    public byte[]? ResultImageBytes { get; private set; }
    public byte[]? ResultGlbBytes { get; private set; }
    public string? ResultSnapshotJson { get; private set; }
    public bool ResultEmbedInDoc { get; private set; } = true;
    public string ResultBackground { get; private set; } = "transparent";

    public MolstarForm(string contextTitle, string initialStructure, string? initialSnapshotJson = null)
    {
        this.contextTitle = contextTitle ?? "PowerPoint Slide";
        this.initialStructure = initialStructure ?? "";
        this.initialSnapshotJson = initialSnapshotJson;

        Text = $"Mol* 3D Molecular Viewer - {this.contextTitle}";
        Width = 1100;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(840, 580);
        BackColor = Color.FromArgb(15, 23, 42);

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
            var molstarFolder = MolstarAssetManager.EnsureMolstarFolder();

            var baseUserDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Smiles2Img", "WebView2UserData");
            Directory.CreateDirectory(baseUserDataFolder);

            string procName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            var procUserDataFolder = Path.Combine(baseUserDataFolder, procName + "_Molstar");
            Directory.CreateDirectory(procUserDataFolder);

            CoreWebView2Environment env;
            try
            {
                env = await CoreWebView2Environment.CreateAsync(null, procUserDataFolder, null);
                await webView.EnsureCoreWebView2Async(env);
            }
            catch (Exception ex) when (ex.HResult == unchecked((int)0x800700AA) || ex.Message.Contains("0x800700AA") || ex.Message.IndexOf("in use", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string fallbackFolder = Path.Combine(Path.GetTempPath(), "Smiles2Img_WV2_Mol_" + procName + "_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(fallbackFolder);
                env = await CoreWebView2Environment.CreateAsync(null, fallbackFolder, null);
                await webView.EnsureCoreWebView2Async(env);
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

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var msg = e.TryGetWebMessageAsString();
            if (string.IsNullOrEmpty(msg)) return;

            if (msg == "ready")
            {
                var escapedContext = EscapeJs(contextTitle);
                var escapedInitial = EscapeJs(initialStructure);
                var escapedSnapshot = EscapeJs(initialSnapshotJson);

                var script = $"if (window.initMolstar) {{ window.initMolstar('{escapedContext}', '{escapedInitial}', '{escapedSnapshot}'); }}";
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
        if (disposing)
        {
            webView?.Dispose();
        }
        base.Dispose(disposing);
    }
}
