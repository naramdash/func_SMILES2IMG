<#
.SYNOPSIS
    Generates a slow, beginner-friendly 5-step mouse walkthrough video and screenshot
    for installing func_SMILES2IMG via PowerShell (common-install-walkthrough).
#>
param(
    [switch]$SaveDebugFrames
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$repoRoot = Split-Path -Parent $PSScriptRoot
$docsDir = Join-Path $repoRoot 'docs'
$distDocsDir = Join-Path $docsDir '.vitepress\dist'
$publicDir = Join-Path $docsDir 'public'
$mediaRoot = Join-Path $publicDir 'media'
$outDir = Join-Path $mediaRoot 'common-install-walkthrough'
$manifestPublicPath = Join-Path $mediaRoot 'manifest.json'
$manifestThemePath = Join-Path $docsDir '.vitepress\theme\mediaManifest.json'

if (-not (Test-Path (Join-Path $distDocsDir 'guide\install.html'))) {
    throw "Built install page not found at $distDocsDir\guide\install.html. Run 'npm.cmd run docs:build' first."
}

New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + $env:Path
$ffmpeg = (Get-Command ffmpeg -ErrorAction Stop).Source
$ffprobe = (Get-Command ffprobe -ErrorAction Stop).Source
$nodeExe = (Get-Command node -ErrorAction Stop).Source
$edgeExe = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
if (-not (Test-Path -LiteralPath $edgeExe)) {
    throw "Microsoft Edge executable not found at $edgeExe"
}

$addinVersion = '1.9.1'
try {
    $propsXml = [xml]([System.IO.File]::ReadAllText((Join-Path $repoRoot 'Directory.Build.props'), [System.Text.UTF8Encoding]::new($false)))
    if ($propsXml.Project.PropertyGroup.Version) {
        $addinVersion = [string]$propsXml.Project.PropertyGroup.Version
    }
} catch {}

# -----------------------------------------------------------------------------
# 1. Serve docs/.vitepress/dist locally and capture real /guide/install page
# -----------------------------------------------------------------------------
$workTemp = Join-Path ([System.IO.Path]::GetTempPath()) ("s2i_install_vid_" + [Guid]::NewGuid().ToString('N'))
$framesDir = Join-Path $workTemp 'frames'
New-Item -ItemType Directory -Path $framesDir -Force | Out-Null

$browserShotPath = Join-Path $workTemp 'edge_install_1680x1080.png'
$serverJsPath = Join-Path $workTemp 'serve.js'
$portFilePath = Join-Path $workTemp 'port.txt'

$serverJs = @'
const http = require('http');
const fs = require('fs');
const path = require('path');

const distRoot = path.resolve(process.argv[2]);
const portOutFile = process.argv[3];
const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.webp': 'image/webp',
  '.jpg': 'image/jpeg',
  '.json': 'application/json; charset=utf-8',
  '.woff2': 'font/woff2',
  '.woff': 'font/woff',
  '.vtt': 'text/vtt; charset=utf-8'
};

const server = http.createServer((req, res) => {
  const urlPath = (req.url || '/').split('?')[0].split('#')[0];
  let rel = urlPath.startsWith('/func_SMILES2IMG/') ? urlPath.slice('/func_SMILES2IMG/'.length) : urlPath.replace(/^\/+/, '');
  if (!rel || rel.endsWith('/')) rel += 'index.html';
  let filePath = path.resolve(distRoot, rel);
  if (!filePath.startsWith(distRoot)) {
    res.writeHead(403);
    res.end();
    return;
  }
  if (!fs.existsSync(filePath) && fs.existsSync(filePath + '.html')) {
    filePath += '.html';
  }
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    const idx = path.join(filePath, 'index.html');
    if (fs.existsSync(idx)) filePath = idx;
    else { res.writeHead(404); res.end('Not found'); return; }
  }
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.html' && filePath.replace(/\\/g, '/').endsWith('/guide/install.html')) {
    let html = fs.readFileSync(filePath, 'utf8');
    html = html.replace('</head>', '<style>.s2i-figure{display:none!important;}</style></head>');
    res.writeHead(200, { 'Content-Type': mimeTypes[ext] });
    res.end(html, 'utf8');
    return;
  }
  res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'application/octet-stream' });
  fs.createReadStream(filePath).pipe(res);
});

server.listen(0, '127.0.0.1', () => {
  const addr = server.address();
  fs.writeFileSync(portOutFile, String(addr.port), 'utf8');
});
'@

[System.IO.File]::WriteAllText($serverJsPath, $serverJs, [System.Text.UTF8Encoding]::new($false))

$serverProc = $null
try {
    $serverProc = Start-Process -FilePath $nodeExe -ArgumentList @("`"$serverJsPath`"", "`"$distDocsDir`"", "`"$portFilePath`"") -PassThru -WindowStyle Hidden
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while (-not (Test-Path -LiteralPath $portFilePath) -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 100
    }
    if (-not (Test-Path -LiteralPath $portFilePath)) {
        throw "Timed out waiting for ephemeral VitePress HTTP server."
    }
    $port = ([System.IO.File]::ReadAllText($portFilePath, [System.Text.UTF8Encoding]::new($false))).Trim()
    $targetUrl = "http://127.0.0.1:$port/func_SMILES2IMG/guide/install.html"
    Write-Host "[+] Capturing real Edge screenshot of $targetUrl (1680x1080)..." -ForegroundColor Cyan

    $edgeArgs = "--headless=new --disable-gpu --hide-scrollbars --lang=en-US --window-size=1680,1080 --force-device-scale-factor=1 --virtual-time-budget=3000 --screenshot=`"$browserShotPath`" `"$targetUrl`""
    $edgeProc = Start-Process -FilePath $edgeExe -ArgumentList $edgeArgs -PassThru -Wait
    if ($edgeProc.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $browserShotPath)) {
        throw "Edge screenshot capture failed (ExitCode=$($edgeProc.ExitCode))."
    }
} finally {
    if ($serverProc -and -not $serverProc.HasExited) {
        try { Stop-Process -Id $serverProc.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# -----------------------------------------------------------------------------
# 2. Inline C# compositor for crisp 1920x1280 5-step beginner walkthrough frames
# -----------------------------------------------------------------------------
$rendererCs = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;

public static class InstallWalkthroughRenderer
{
    private static bool IsTeal(Color c)
    {
        bool isLightBrand =
            Math.Abs(c.R - 15) <= 16 &&
            Math.Abs(c.G - 118) <= 16 &&
            Math.Abs(c.B - 110) <= 16;
        bool isDarkBrand =
            Math.Abs(c.R - 45) <= 16 &&
            Math.Abs(c.G - 212) <= 16 &&
            Math.Abs(c.B - 191) <= 16;
        return isLightBrand || isDarkBrand;
    }

    public static Rectangle LocateTealCopyButton(string pngPath)
    {
        using (var bmp = new Bitmap(pngPath))
        {
            int w = bmp.Width;
            int h = bmp.Height;
            int xMin = (int)(w * 0.30);
            int xMax = (int)(w * 0.90);
            int yMin = (int)(h * 0.25);
            int yMax = (int)(h * 0.85);

            var visited = new bool[w, h];
            Rectangle bestRect = Rectangle.Empty;
            int bestScore = 0;

            var queue = new Queue<Point>();
            for (int y = yMin; y < yMax; y++)
            {
                for (int x = xMin; x < xMax; x++)
                {
                    if (visited[x, y]) continue;
                    Color px = bmp.GetPixel(x, y);
                    if (!IsTeal(px)) continue;

                    visited[x, y] = true;
                    queue.Clear();
                    queue.Enqueue(new Point(x, y));

                    int minX = x, maxX = x, minY = y, maxY = y, count = 0;
                    while (queue.Count > 0)
                    {
                        Point pt = queue.Dequeue();
                        count++;
                        if (pt.X < minX) minX = pt.X;
                        if (pt.X > maxX) maxX = pt.X;
                        if (pt.Y < minY) minY = pt.Y;
                        if (pt.Y > maxY) maxY = pt.Y;

                        for (int dy = -1; dy <= 1; dy++)
                        {
                            for (int dx = -1; dx <= 1; dx++)
                            {
                                if (dx == 0 && dy == 0) continue;
                                int nx = pt.X + dx;
                                int ny = pt.Y + dy;
                                if (nx < xMin || nx >= xMax || ny < yMin || ny >= yMax) continue;
                                if (visited[nx, ny]) continue;
                                if (IsTeal(bmp.GetPixel(nx, ny)))
                                {
                                    visited[nx, ny] = true;
                                    queue.Enqueue(new Point(nx, ny));
                                }
                            }
                        }
                    }

                    int bw = maxX - minX + 1;
                    int bh = maxY - minY + 1;
                    if (bw >= 26 && bw <= 48 && bh >= 26 && bh <= 48 && Math.Abs(bw - bh) <= 6 && count >= 450)
                    {
                        if (count > bestScore)
                        {
                            bestScore = count;
                            bestRect = new Rectangle(minX, minY, bw, bh);
                        }
                    }
                }
            }

            if (bestRect.IsEmpty)
                throw new InvalidOperationException("Failed to locate the teal Copy button (#0f766e, ~34x34 px) in captured install page: " + pngPath);

            return bestRect;
        }
    }

    private static GraphicsPath RoundedRectPath(RectangleF bounds, float radius)
    {
        float d = radius * 2f;
        var path = new GraphicsPath();
        if (radius <= 0.1f)
        {
            path.AddRectangle(bounds);
            return path;
        }
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    private static void FillRoundedRect(Graphics g, Brush brush, RectangleF bounds, float radius)
    {
        using (var path = RoundedRectPath(bounds, radius))
        {
            g.FillPath(brush, path);
        }
    }

    private static void DrawRoundedRect(Graphics g, Pen pen, RectangleF bounds, float radius)
    {
        using (var path = RoundedRectPath(bounds, radius))
        {
            g.DrawPath(pen, path);
        }
    }

    private static double SmoothStep(double t)
    {
        if (t <= 0) return 0;
        if (t >= 1) return 1;
        return t * t * (3.0 - 2.0 * t);
    }

    private static PointF LerpPoint(PointF a, PointF b, double t)
    {
        double e = SmoothStep(t);
        return new PointF((float)(a.X + (b.X - a.X) * e), (float)(a.Y + (b.Y - a.Y) * e));
    }

    private static void DrawCursorAndClick(Graphics g, float x, float y, float clickProgress, bool isRightClick, string badgeText)
    {
        if (clickProgress > 0f && clickProgress < 1f)
        {
            int alpha = Math.Max(0, Math.Min(230, (int)((1f - clickProgress) * 230)));
            Color ringColor = isRightClick ? Color.FromArgb(alpha, 249, 115, 22) : Color.FromArgb(alpha, 37, 99, 235);
            float r1 = 10f + clickProgress * 34f;
            using (var pen1 = new Pen(ringColor, 3.2f))
            {
                g.DrawEllipse(pen1, x - r1, y - r1, r1 * 2f, r1 * 2f);
            }
            if (isRightClick)
            {
                float r2 = 6f + clickProgress * 22f;
                using (var pen2 = new Pen(Color.FromArgb(alpha, 251, 146, 60), 2.4f))
                {
                    g.DrawEllipse(pen2, x - r2, y - r2, r2 * 2f, r2 * 2f);
                }
            }
        }

        if (!string.IsNullOrEmpty(badgeText))
        {
            using (var fnt = new Font("Segoe UI", 10.5f, FontStyle.Bold))
            {
                var sz = g.MeasureString(badgeText, fnt);
                float bx = Math.Min(1920f - sz.Width - 28f, x + 22f);
                float by = Math.Max(76f, y - 34f);
                var rect = new RectangleF(bx, by, sz.Width + 18f, sz.Height + 8f);
                Color badgeBg = isRightClick ? Color.FromArgb(240, 234, 88, 12) : Color.FromArgb(240, 37, 99, 235);
                using (var sh = new SolidBrush(Color.FromArgb(80, 0, 0, 0)))
                {
                    FillRoundedRect(g, sh, new RectangleF(rect.X + 2f, rect.Y + 3f, rect.Width, rect.Height), 7f);
                }
                using (var bg = new SolidBrush(badgeBg))
                using (var border = new Pen(Color.White, 1.5f))
                using (var tb = new SolidBrush(Color.White))
                {
                    FillRoundedRect(g, bg, rect, 7f);
                    DrawRoundedRect(g, border, rect, 7f);
                    g.DrawString(badgeText, fnt, tb, bx + 9f, by + 4f);
                }
            }
        }

        // Crisp high-visibility arrow cursor (large for tutorial clarity)
        var pts = new[]
        {
            new PointF(x, y),
            new PointF(x, y + 26f),
            new PointF(x + 7.5f, y + 20f),
            new PointF(x + 12.5f, y + 30f),
            new PointF(x + 17f, y + 28f),
            new PointF(x + 12f, y + 18f),
            new PointF(x + 20.5f, y + 18f)
        };
        using (var sh = new SolidBrush(Color.FromArgb(80, 0, 0, 0)))
        {
            var shPts = new PointF[pts.Length];
            for (int i = 0; i < pts.Length; i++) shPts[i] = new PointF(pts[i].X + 2f, pts[i].Y + 2f);
            g.FillPolygon(sh, shPts);
        }
        using (var fill = new SolidBrush(Color.FromArgb(255, 15, 23, 42)))
        using (var border = new Pen(Color.White, 2.5f))
        {
            border.LineJoin = LineJoin.Round;
            g.FillPolygon(fill, pts);
            g.DrawPolygon(border, pts);
        }
    }

    private static void DrawTopStepBar(Graphics g, int activeStep, string instructionText, string actionHint, bool isRightHint)
    {
        using (var bg = new SolidBrush(Color.FromArgb(255, 15, 23, 42)))
        {
            g.FillRectangle(bg, 0, 0, 1920, 68);
        }
        using (var line = new Pen(Color.FromArgb(255, 51, 65, 85), 1.5f))
        {
            g.DrawLine(line, 0, 67, 1920, 67);
        }

        float cx = 28f;
        float cy = 17f;
        using (var numFont = new Font("Segoe UI", 12f, FontStyle.Bold))
        using (var stepFont = new Font("Segoe UI", 17f, FontStyle.Bold))
        using (var hintFont = new Font("Segoe UI", 11f, FontStyle.Bold))
        using (var whiteBrush = new SolidBrush(Color.White))
        {
            for (int s = 1; s <= 5; s++)
            {
                Color pillColor;
                Color textColor = Color.White;
                if (s < activeStep)
                    pillColor = Color.FromArgb(255, 16, 185, 129); // Emerald completed
                else if (s == activeStep)
                    pillColor = Color.FromArgb(255, 13, 148, 136); // Teal active
                else
                {
                    pillColor = Color.FromArgb(255, 30, 41, 59);
                    textColor = Color.FromArgb(255, 148, 163, 184);
                }

                var pillRect = new RectangleF(cx, cy, 34f, 34f);
                using (var pb = new SolidBrush(pillColor))
                {
                    g.FillEllipse(pb, pillRect);
                }
                if (s == activeStep)
                {
                    using (var ring = new Pen(Color.FromArgb(255, 94, 234, 212), 2f))
                    {
                        g.DrawEllipse(ring, pillRect);
                    }
                }
                string label = s < activeStep ? "\u2713" : s.ToString();
                var lsz = g.MeasureString(label, numFont);
                using (var tb = new SolidBrush(textColor))
                {
                    g.DrawString(label, numFont, tb, cx + (34f - lsz.Width) / 2f, cy + (34f - lsz.Height) / 2f);
                }

                if (s < 5)
                {
                    using (var sep = new Pen(Color.FromArgb(255, 71, 85, 105), 2f))
                    {
                        g.DrawLine(sep, cx + 38f, 34f, cx + 48f, 34f);
                    }
                }
                cx += 52f;
            }

            g.DrawString(instructionText, stepFont, whiteBrush, cx + 16f, 18f);

            if (!string.IsNullOrEmpty(actionHint))
            {
                var hsz = g.MeasureString(actionHint, hintFont);
                float hw = hsz.Width + 28f;
                var hRect = new RectangleF(1920f - hw - 28f, 17f, hw, 34f);
                Color hBg = isRightHint
                    ? Color.FromArgb(255, 234, 88, 12)
                    : (actionHint.StartsWith("Done") ? Color.FromArgb(255, 16, 185, 129) : Color.FromArgb(255, 37, 99, 235));
                using (var hb = new SolidBrush(hBg))
                {
                    FillRoundedRect(g, hb, hRect, 8f);
                }
                g.DrawString(actionHint, hintFont, whiteBrush, hRect.X + 14f, hRect.Y + (34f - hsz.Height) / 2f);
            }
        }
    }

    private static void DrawTaskbar(Graphics g, bool startHover, bool showTerminalIcon, string versionLabel)
    {
        using (var bg = new SolidBrush(Color.FromArgb(255, 20, 25, 35)))
        {
            g.FillRectangle(bg, 0, 1224, 1920, 56);
        }
        using (var topLine = new Pen(Color.FromArgb(255, 51, 65, 85), 1f))
        {
            g.DrawLine(topLine, 0, 1224, 1920, 1224);
        }

        // Start button at bottom-left (18..62, 1230..1274), center = (40, 1252)
        var startRect = new RectangleF(18f, 1230f, 44f, 44f);
        if (startHover)
        {
            using (var hb = new SolidBrush(Color.FromArgb(70, 255, 255, 255)))
            using (var hp = new Pen(Color.FromArgb(255, 249, 115, 22), 2f))
            {
                FillRoundedRect(g, hb, startRect, 7f);
                DrawRoundedRect(g, hp, startRect, 7f);
            }
        }
        // 4 Windows 11 blue panes centered inside startRect
        float px = 29f, py = 1241f, ps = 10f, gap = 2f;
        using (var p1 = new SolidBrush(Color.FromArgb(255, 56, 189, 248)))
        using (var p2 = new SolidBrush(Color.FromArgb(255, 14, 165, 233)))
        {
            g.FillRectangle(p1, px, py, ps, ps);
            g.FillRectangle(p1, px + ps + gap, py, ps, ps);
            g.FillRectangle(p2, px, py + ps + gap, ps, ps);
            g.FillRectangle(p2, px + ps + gap, py + ps + gap, ps, ps);
        }

        // Clean Edge taskbar icon (80..124) with active indicator
        var edgeRect = new RectangleF(80f, 1230f, 44f, 44f);
        using (var eb = new SolidBrush(Color.FromArgb(30, 255, 255, 255)))
        {
            FillRoundedRect(g, eb, edgeRect, 7f);
        }
        using (var ec = new SolidBrush(Color.FromArgb(255, 14, 165, 233)))
        using (var ef = new Font("Segoe UI", 12f, FontStyle.Bold))
        using (var wb = new SolidBrush(Color.White))
        {
            g.FillEllipse(ec, 89f, 1239f, 26f, 26f);
            g.DrawString("e", ef, wb, 95f, 1240f);
        }
        using (var ind = new SolidBrush(Color.FromArgb(255, 148, 163, 184)))
        {
            FillRoundedRect(g, ind, new RectangleF(95f, 1270f, 14f, 3f), 1.5f);
        }

        // PowerShell taskbar icon once opened (136..180)
        if (showTerminalIcon)
        {
            var termRect = new RectangleF(136f, 1230f, 44f, 44f);
            using (var tb = new SolidBrush(Color.FromArgb(55, 255, 255, 255)))
            {
                FillRoundedRect(g, tb, termRect, 7f);
            }
            using (var psBg = new SolidBrush(Color.FromArgb(255, 30, 58, 138)))
            using (var psFont = new Font("Consolas", 10.5f, FontStyle.Bold))
            using (var wb = new SolidBrush(Color.White))
            {
                FillRoundedRect(g, psBg, new RectangleF(145f, 1239f, 26f, 26f), 5f);
                g.DrawString(">_", psFont, wb, 147f, 1243f);
            }
            using (var ind = new SolidBrush(Color.FromArgb(255, 45, 212, 191)))
            {
                FillRoundedRect(g, ind, new RectangleF(149f, 1270f, 18f, 3f), 1.5f);
            }
        }

        // Clean right-side clock (no personal info)
        using (var tf = new Font("Segoe UI", 9.5f, FontStyle.Regular))
        using (var tb = new SolidBrush(Color.FromArgb(255, 226, 232, 240)))
        {
            g.DrawString("v" + versionLabel, tf, tb, 1760f, 1243f);
            g.DrawString("10:30 AM", tf, tb, 1835f, 1233f);
            g.DrawString("2026-11-23", tf, tb, 1830f, 1250f);
        }
    }

    private static void DrawBrowserWindow(Graphics g, Bitmap browserBmp, Rectangle copyBtnLocalRect, bool copyHover, bool copyDone, bool showCopiedToast)
    {
        var frameRect = new RectangleF(120f, 82f, 1680f, 1126f);
        using (var sh = new SolidBrush(Color.FromArgb(90, 0, 0, 0)))
        {
            FillRoundedRect(g, sh, new RectangleF(frameRect.X + 4f, frameRect.Y + 6f, frameRect.Width, frameRect.Height), 10f);
        }

        using (var chromeBg = new SolidBrush(Color.FromArgb(255, 241, 245, 249)))
        {
            FillRoundedRect(g, chromeBg, frameRect, 10f);
        }

        // Top address/title bar (y=82..126, h=44)
        using (var barFont = new Font("Segoe UI", 10.5f, FontStyle.Regular))
        using (var tabFont = new Font("Segoe UI", 10f, FontStyle.Bold))
        using (var textBrush = new SolidBrush(Color.FromArgb(255, 30, 41, 59)))
        using (var urlBrush = new SolidBrush(Color.FromArgb(255, 71, 85, 105)))
        {
            // Browser active tab pill
            using (var tabBg = new SolidBrush(Color.White))
            {
                FillRoundedRect(g, tabBg, new RectangleF(136f, 89f, 290f, 32f), 7f);
            }
            g.DrawString("Installation | func_SMILES2IMG", tabFont, textBrush, 152f, 96f);

            // Clean URL pill
            var urlRect = new RectangleF(450f, 89f, 820f, 30f);
            using (var urlBg = new SolidBrush(Color.White))
            using (var urlBorder = new Pen(Color.FromArgb(255, 203, 213, 225), 1f))
            {
                FillRoundedRect(g, urlBg, urlRect, 15f);
                DrawRoundedRect(g, urlBorder, urlRect, 15f);
            }
            g.DrawString("https://naramdash.github.io/func_SMILES2IMG/guide/install", barFont, urlBrush, 472f, 94f);
        }

        // 1:1 real headless Edge capture viewport (x=120, y=126, w=1680, h=1080)
        var dstRect = new Rectangle(120, 126, 1680, 1080);
        g.DrawImage(browserBmp, dstRect, 0, 0, browserBmp.Width, browserBmp.Height, GraphicsUnit.Pixel);

        // Chrome border
        using (var border = new Pen(Color.FromArgb(255, 148, 163, 184), 1.5f))
        {
            DrawRoundedRect(g, border, frameRect, 10f);
        }

        // Exact Copy Button overlay (mapped 1:1 from real pixel coordinates)
        var btnScreen = new RectangleF(
            120f + copyBtnLocalRect.X,
            126f + copyBtnLocalRect.Y,
            copyBtnLocalRect.Width,
            copyBtnLocalRect.Height);

        if (copyHover)
        {
            using (var ring = new Pen(Color.FromArgb(255, 45, 212, 191), 3f))
            {
                DrawRoundedRect(g, ring, new RectangleF(btnScreen.X - 4f, btnScreen.Y - 4f, btnScreen.Width + 8f, btnScreen.Height + 8f), 10f);
            }
        }

        if (copyDone)
        {
            using (var em = new SolidBrush(Color.FromArgb(255, 16, 185, 129)))
            {
                FillRoundedRect(g, em, btnScreen, 8f);
            }
            using (var checkPen = new Pen(Color.White, 2.8f))
            {
                checkPen.StartCap = LineCap.Round;
                checkPen.EndCap = LineCap.Round;
                checkPen.LineJoin = LineJoin.Round;
                float cx = btnScreen.X + btnScreen.Width / 2f;
                float cy = btnScreen.Y + btnScreen.Height / 2f;
                g.DrawLines(checkPen, new[]
                {
                    new PointF(cx - 7f, cy + 0.5f),
                    new PointF(cx - 2f, cy + 5.5f),
                    new PointF(cx + 7.5f, cy - 5.5f)
                });
            }
        }

        if (showCopiedToast)
        {
            using (var tf = new Font("Segoe UI", 10.5f, FontStyle.Bold))
            {
                string msg = "\u2713 Copied!";
                var sz = g.MeasureString(msg, tf);
                var pill = new RectangleF(
                    btnScreen.X + (btnScreen.Width - (sz.Width + 18f)) / 2f,
                    btnScreen.Y - 38f,
                    sz.Width + 18f,
                    30f);
                using (var bg = new SolidBrush(Color.FromArgb(245, 6, 95, 70)))
                using (var bdr = new Pen(Color.FromArgb(255, 52, 211, 153), 1.5f))
                using (var wb = new SolidBrush(Color.White))
                {
                    FillRoundedRect(g, bg, pill, 7f);
                    DrawRoundedRect(g, bdr, pill, 7f);
                    g.DrawString(msg, tf, wb, pill.X + 9f, pill.Y + (30f - sz.Height) / 2f);
                }
            }
        }
    }

    private struct StartMenuItem
    {
        public string Label;
        public string Hint;
        public bool IsDivider;
        public bool IsTarget;
        public StartMenuItem(string label, string hint, bool isDivider, bool isTarget)
        {
            Label = label;
            Hint = hint;
            IsDivider = isDivider;
            IsTarget = isTarget;
        }
    }

    private static void DrawStartContextMenu(Graphics g, bool hoverTerminal)
    {
        var items = new StartMenuItem[]
        {
            new StartMenuItem("Installed apps", "", false, false),
            new StartMenuItem("Power Options", "", false, false),
            new StartMenuItem("System", "", false, false),
            new StartMenuItem("Device Manager", "", false, false),
            new StartMenuItem("", "", true, false),
            new StartMenuItem("Terminal", "Click this (No Admin)", false, true),
            new StartMenuItem("Terminal (Admin)", "", false, false),
            new StartMenuItem("", "", true, false),
            new StartMenuItem("Task Manager", "", false, false),
            new StartMenuItem("Settings", "", false, false),
            new StartMenuItem("File Explorer", "", false, false),
            new StartMenuItem("Search", "", false, false),
            new StartMenuItem("Run", "", false, false)
        };

        float menuX = 20f;
        float menuW = 355f;
        float menuH = 446f;
        float menuY = 1218f - menuH; // 772..1218
        var menuRect = new RectangleF(menuX, menuY, menuW, menuH);

        using (var sh = new SolidBrush(Color.FromArgb(110, 0, 0, 0)))
        {
            FillRoundedRect(g, sh, new RectangleF(menuX + 4f, menuY + 6f, menuW, menuH), 10f);
        }
        using (var bg = new SolidBrush(Color.FromArgb(250, 31, 41, 55)))
        using (var border = new Pen(Color.FromArgb(255, 71, 85, 105), 1.5f))
        {
            FillRoundedRect(g, bg, menuRect, 10f);
            DrawRoundedRect(g, border, menuRect, 10f);
        }

        float curY = menuY + 12f;
        using (var fNorm = new Font("Segoe UI", 11.5f, FontStyle.Regular))
        using (var fBold = new Font("Segoe UI", 11.5f, FontStyle.Bold))
        using (var fHint = new Font("Segoe UI", 9.5f, FontStyle.Bold))
        using (var bNorm = new SolidBrush(Color.FromArgb(255, 226, 232, 240)))
        using (var bWhite = new SolidBrush(Color.White))
        using (var bTeal = new SolidBrush(Color.FromArgb(255, 16, 185, 129)))
        {
            foreach (var it in items)
            {
                if (it.IsDivider)
                {
                    using (var div = new Pen(Color.FromArgb(255, 71, 85, 105), 1f))
                    {
                        g.DrawLine(div, menuX + 12f, curY + 5f, menuX + menuW - 12f, curY + 5f);
                    }
                    curY += 12f;
                    continue;
                }

                float rowH = 36f;
                var rowRect = new RectangleF(menuX + 8f, curY, menuW - 16f, rowH);
                if (it.IsTarget)
                {
                    Color hl = hoverTerminal
                        ? Color.FromArgb(255, 37, 99, 235)
                        : Color.FromArgb(110, 15, 118, 110);
                    using (var hb = new SolidBrush(hl))
                    {
                        FillRoundedRect(g, hb, rowRect, 6f);
                    }
                    using (var bar = new SolidBrush(hoverTerminal ? Color.FromArgb(255, 96, 165, 250) : Color.FromArgb(255, 45, 212, 191)))
                    {
                        FillRoundedRect(g, bar, new RectangleF(rowRect.X + 2f, rowRect.Y + 7f, 4f, rowH - 14f), 2f);
                    }
                    g.DrawString(it.Label, fBold, bWhite, rowRect.X + 16f, rowRect.Y + 7f);

                    var hsz = g.MeasureString(it.Hint, fHint);
                    var pill = new RectangleF(rowRect.Right - hsz.Width - 18f, rowRect.Y + 6f, hsz.Width + 12f, 24f);
                    FillRoundedRect(g, bTeal, pill, 5f);
                    g.DrawString(it.Hint, fHint, bWhite, pill.X + 6f, pill.Y + 3f);
                }
                else
                {
                    g.DrawString(it.Label, fNorm, bNorm, rowRect.X + 16f, rowRect.Y + 7f);
                }
                curY += rowH;
            }
        }
    }

    private struct TermLine
    {
        public string Text;
        public Color Color;
        public TermLine(string t, Color c) { Text = t; Color = c; }
    }

    private static void DrawPowerShellWindow(
        Graphics g,
        string commandText,
        bool showBlinkingCursor,
        bool showRightClickHint,
        bool showPressEnterBadge,
        int revealedOutputLines,
        bool showDoneBanner,
        string versionLabel)
    {
        var winRect = new RectangleF(280f, 240f, 1360f, 750f);
        using (var sh = new SolidBrush(Color.FromArgb(130, 0, 0, 0)))
        {
            FillRoundedRect(g, sh, new RectangleF(winRect.X + 6f, winRect.Y + 10f, winRect.Width, winRect.Height), 12f);
        }

        // Terminal window background
        using (var body = new SolidBrush(Color.FromArgb(255, 12, 20, 37)))
        {
            FillRoundedRect(g, body, winRect, 12f);
        }

        // Title / Tab bar (y=180..226, h=46)
        using (var titleBg = new SolidBrush(Color.FromArgb(255, 30, 41, 59)))
        {
            FillRoundedRect(g, titleBg, new RectangleF(winRect.X, winRect.Y, winRect.Width, 50f), 12f);
            g.FillRectangle(titleBg, winRect.X, winRect.Y + 26f, winRect.Width, 24f);
        }
        using (var tabBg = new SolidBrush(Color.FromArgb(255, 12, 20, 37)))
        {
            FillRoundedRect(g, tabBg, new RectangleF(winRect.X + 16f, winRect.Y + 8f, 260f, 42f), 8f);
        }
        using (var psIcon = new SolidBrush(Color.FromArgb(255, 37, 99, 235)))
        using (var psIconFont = new Font("Consolas", 9.5f, FontStyle.Bold))
        using (var tabFont = new Font("Segoe UI", 11f, FontStyle.Bold))
        using (var ctrlFont = new Font("Segoe UI", 12f, FontStyle.Regular))
        using (var wb = new SolidBrush(Color.White))
        using (var mb = new SolidBrush(Color.FromArgb(255, 148, 163, 184)))
        {
            FillRoundedRect(g, psIcon, new RectangleF(winRect.X + 28f, winRect.Y + 18f, 24f, 20f), 4f);
            g.DrawString(">_", psIconFont, wb, winRect.X + 30f, winRect.Y + 20f);
            g.DrawString("Windows PowerShell", tabFont, wb, winRect.X + 60f, winRect.Y + 17f);
            g.DrawString("\u2014    \u25a1    \u2715", ctrlFont, mb, winRect.Right - 108f, winRect.Y + 14f);
        }

        using (var border = new Pen(Color.FromArgb(255, 56, 189, 248), 2f))
        {
            DrawRoundedRect(g, border, winRect, 12f);
        }

        // Terminal text body (starting y=252, Consolas 14.5pt -> ~10.8px char width, 1290px inner width)
        float textX = winRect.X + 32f;
        float textY = winRect.Y + 68f;
        float lineH = 26f;

        using (var mono = new Font("Consolas", 14.5f, FontStyle.Regular))
        using (var promptBrush = new SolidBrush(Color.FromArgb(255, 226, 232, 240)))
        using (var cmdBrush = new SolidBrush(Color.FromArgb(255, 250, 204, 21)))
        using (var sf = new StringFormat(StringFormat.GenericTypographic))
        {
            sf.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
            string promptPrefix = @"PS C:\Users\User> ";
            g.DrawString(promptPrefix, mono, promptBrush, textX, textY, sf);
            var prefixSz = g.MeasureString(promptPrefix, mono, 2000, sf);
            float cmdX = textX + prefixSz.Width;

            if (!string.IsNullOrEmpty(commandText))
            {
                g.DrawString(commandText, mono, cmdBrush, cmdX, textY, sf);
            }
            else if (showBlinkingCursor)
            {
                using (var curBrush = new SolidBrush(Color.FromArgb(255, 226, 232, 240)))
                {
                    g.FillRectangle(curBrush, cmdX + 2f, textY + 3f, 11f, 20f);
                }
            }

            if (showRightClickHint && string.IsNullOrEmpty(commandText))
            {
                using (var hf = new Font("Segoe UI", 13f, FontStyle.Bold))
                using (var hb = new SolidBrush(Color.FromArgb(255, 251, 146, 60)))
                {
                    g.DrawString("Right-click anywhere inside this window to paste automatically", hf, hb, textX, textY + 60f);
                }
            }

            var outputLines = new List<TermLine>
            {
                new TermLine("=== func_SMILES2IMG Installer (Excel, Word, PowerPoint) ===", Color.FromArgb(255, 34, 211, 238)),
                new TermLine("[+] Detected Office Architecture: x64 (Office 16.0)", Color.FromArgb(255, 74, 222, 128)),
                new TermLine("[+] Downloading func_SMILES2IMG package (v" + versionLabel + ")...", Color.FromArgb(255, 34, 211, 238)),
                new TermLine(@"[+] Excel Add-in registered (OPEN): ...\Smiles2Img-AddIn64.xll", Color.FromArgb(255, 74, 222, 128)),
                new TermLine("[+] Word COM Add-in registered: func_SMILES2IMG", Color.FromArgb(255, 74, 222, 128)),
                new TermLine("[+] PowerPoint COM Add-in registered: func_SMILES2IMG", Color.FromArgb(255, 74, 222, 128)),
                new TermLine("", Color.White),
                new TermLine("======================================================================", Color.FromArgb(255, 34, 211, 238)),
                new TermLine(" [SUMMARY] func_SMILES2IMG Install Summary", Color.FromArgb(255, 34, 211, 238)),
                new TermLine("----------------------------------------------------------------------", Color.FromArgb(255, 34, 211, 238)),
                new TermLine("  * Mode:               Clean Install / Upgrade", Color.White),
                new TermLine("  * Current Version:    v" + versionLabel, Color.FromArgb(255, 74, 222, 128)),
                new TermLine("  * Office Arch:        x64 (Office 16.0)", Color.White),
                new TermLine("  * Registered Apps:    Excel, Word, PowerPoint", Color.White),
                new TermLine(@"  * Install Folder:     C:\Users\User\AppData\Roaming\Microsoft\AddIns\Smiles2Img", Color.FromArgb(255, 148, 163, 184)),
                new TermLine("======================================================================", Color.FromArgb(255, 34, 211, 238)),
                new TermLine("", Color.White),
                new TermLine("[SUCCESS] func_SMILES2IMG setup complete!", Color.FromArgb(255, 52, 211, 153))
            };

            float curY = textY + lineH * 1.4f;
            int count = Math.Min(revealedOutputLines, outputLines.Count);
            for (int i = 0; i < count; i++)
            {
                if (!string.IsNullOrEmpty(outputLines[i].Text))
                {
                    using (var lb = new SolidBrush(outputLines[i].Color))
                    {
                        g.DrawString(outputLines[i].Text, mono, lb, textX, curY);
                    }
                }
                curY += lineH;
            }
        }

        if (showPressEnterBadge)
        {
            using (var kf = new Font("Segoe UI", 14f, FontStyle.Bold))
            {
                string ktxt = "Press Enter \u21b5";
                var ksz = g.MeasureString(ktxt, kf);
                var kRect = new RectangleF(winRect.Right - ksz.Width - 64f, winRect.Y + 130f, ksz.Width + 32f, 46f);
                using (var kBg = new SolidBrush(Color.FromArgb(255, 37, 99, 235)))
                using (var kBdr = new Pen(Color.White, 2f))
                using (var wb = new SolidBrush(Color.White))
                {
                    FillRoundedRect(g, kBg, kRect, 9f);
                    DrawRoundedRect(g, kBdr, kRect, 9f);
                    g.DrawString(ktxt, kf, wb, kRect.X + 16f, kRect.Y + (46f - ksz.Height) / 2f);
                }
            }
        }

        if (showDoneBanner)
        {
            using (var df = new Font("Segoe UI", 13.5f, FontStyle.Bold))
            {
                string dtxt = "\u2713 Done! Reopen Excel, Word, or PowerPoint to use the [SMILES] tab.";
                var dsz = g.MeasureString(dtxt, df);
                var dRect = new RectangleF(winRect.X + 32f, winRect.Bottom - 62f, dsz.Width + 32f, 42f);
                using (var dBg = new SolidBrush(Color.FromArgb(255, 6, 95, 70)))
                using (var dBdr = new Pen(Color.FromArgb(255, 52, 211, 153), 2f))
                using (var wb = new SolidBrush(Color.White))
                {
                    FillRoundedRect(g, dBg, dRect, 8f);
                    DrawRoundedRect(g, dBdr, dRect, 8f);
                    g.DrawString(dtxt, df, wb, dRect.X + 16f, dRect.Y + (42f - dsz.Height) / 2f);
                }
            }
        }
    }

    public static void RenderAllFrames(
        string browserPngPath,
        Rectangle copyBtnLocalRect,
        string framesDir,
        string shotOutPath,
        string versionLabel,
        string debugDirOrEmpty)
    {
        const int fps = 15;
        const int totalFrames = 405; // 27.0 seconds exact
        const string psCmd = "irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex";

        using (var browserBmp = new Bitmap(browserPngPath))
        {
            browserBmp.SetResolution(96f, 96f);

            PointF ptIdleStart = new PointF(1460f, 520f);
            PointF ptCopyBtn = new PointF(120f + copyBtnLocalRect.X + copyBtnLocalRect.Width / 2f, 126f + copyBtnLocalRect.Y + copyBtnLocalRect.Height / 2f);
            PointF ptStartBtn = new PointF(40f, 1252f);
            PointF ptMenuTerminal = new PointF(150f, 958f);
            PointF ptTerminalCenter = new PointF(760f, 520f);

            for (int f = 0; f < totalFrames; f++)
            {
                double t = f / (double)fps;

                int activeStep;
                string stepTitle;
                string stepHint;
                bool isRightHint = false;

                bool copyHover = false;
                bool copyDone = t >= 3.55;
                bool showCopiedToast = t >= 3.55 && t < 7.20;

                bool startHover = t >= 6.80 && t < 8.00;
                bool showStartMenu = t >= 8.00 && t < 12.90;
                bool hoverTerminal = t >= 11.70 && t < 12.90;

                bool showTerminal = t >= 12.90;
                bool blinkCursor = (t < 17.50) && (((f / 6) % 2) == 0);
                bool showRightClickHint = t >= 15.00 && t < 17.50;
                string typedCmd = t >= 17.50 ? psCmd : "";
                bool showPressEnter = t >= 19.80 && t < 21.40;
                int revealedLines = 0;
                if (t >= 21.40)
                {
                    double dt = t - 21.40;
                    revealedLines = Math.Min(18, 1 + (int)(dt / 0.16));
                }
                bool showDoneBanner = t >= 24.40;

                PointF cursor = ptIdleStart;
                float clickProg = -1f;
                bool clickIsRight = false;
                string clickBadge = "";

                if (t < 5.20)
                {
                    activeStep = 1;
                    stepTitle = "Step 1 / 5: Click the teal Copy button on the install command";
                    stepHint = "Left-Click";
                    if (t < 1.20)
                        cursor = ptIdleStart;
                    else if (t < 2.60)
                        cursor = LerpPoint(ptIdleStart, ptCopyBtn, (t - 1.20) / 1.40);
                    else
                    {
                        cursor = ptCopyBtn;
                        copyHover = t < 3.60;
                        if (t >= 3.15 && t < 4.15)
                        {
                            clickProg = (float)((t - 3.15) / 1.00);
                            clickBadge = t < 3.90 ? "Left-Click" : "";
                        }
                    }
                }
                else if (t < 10.60)
                {
                    activeStep = 2;
                    stepTitle = "Step 2 / 5: Right-click the Windows Start button";
                    stepHint = "Right-Click";
                    isRightHint = true;
                    if (t < 6.80)
                        cursor = LerpPoint(ptCopyBtn, ptStartBtn, (t - 5.20) / 1.60);
                    else
                    {
                        cursor = ptStartBtn;
                        if (t >= 7.45 && t < 8.45)
                        {
                            clickProg = (float)((t - 7.45) / 1.00);
                            clickIsRight = true;
                            clickBadge = t < 8.30 ? "Right-Click" : "";
                        }
                    }
                }
                else if (t < 15.00)
                {
                    activeStep = 3;
                    stepTitle = "Step 3 / 5: Click \"Terminal\" (Win 11) or \"Windows PowerShell\" (Win 10)";
                    stepHint = "Left-Click";
                    if (t < 11.80)
                        cursor = LerpPoint(ptStartBtn, ptMenuTerminal, (t - 10.60) / 1.20);
                    else
                    {
                        cursor = ptMenuTerminal;
                        if (t >= 12.45 && t < 13.35)
                        {
                            clickProg = (float)((t - 12.45) / 0.90);
                            clickBadge = t < 13.15 ? "Left-Click" : "";
                        }
                    }
                }
                else if (t < 19.80)
                {
                    activeStep = 4;
                    stepTitle = "Step 4 / 5: Right-click inside PowerShell to paste the command";
                    stepHint = "Right-Click";
                    isRightHint = true;
                    if (t < 16.40)
                        cursor = LerpPoint(ptMenuTerminal, ptTerminalCenter, (t - 15.00) / 1.40);
                    else
                    {
                        cursor = ptTerminalCenter;
                        if (t >= 17.00 && t < 18.00)
                        {
                            clickProg = (float)((t - 17.00) / 1.00);
                            clickIsRight = true;
                            clickBadge = t < 17.85 ? "Right-Click" : "";
                        }
                    }
                }
                else
                {
                    activeStep = 5;
                    stepTitle = "Step 5 / 5: Press Enter and wait until [SUCCESS] appears";
                    stepHint = showDoneBanner ? "Done!" : "Press Enter";
                    cursor = ptTerminalCenter;
                }

                using (var frame = new Bitmap(1920, 1280, PixelFormat.Format24bppRgb))
                {
                    frame.SetResolution(96f, 96f);
                    using (var g = Graphics.FromImage(frame))
                    {
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                        // Clean desktop background
                        using (var desk = new LinearGradientBrush(
                            new Rectangle(0, 0, 1920, 1280),
                            Color.FromArgb(255, 30, 41, 59),
                            Color.FromArgb(255, 15, 23, 42),
                            45f))
                        {
                            g.FillRectangle(desk, 0, 0, 1920, 1280);
                        }

                        DrawBrowserWindow(g, browserBmp, copyBtnLocalRect, copyHover, copyDone, showCopiedToast);

                        if (showStartMenu)
                            DrawStartContextMenu(g, hoverTerminal);

                        if (showTerminal)
                            DrawPowerShellWindow(g, typedCmd, blinkCursor, showRightClickHint, showPressEnter, revealedLines, showDoneBanner, versionLabel);

                        DrawTopStepBar(g, activeStep, stepTitle, stepHint, isRightHint);
                        DrawTaskbar(g, startHover, showTerminal, versionLabel);

                        // Keep mouse cursor clean over everything except hide subtle during final Done hold? Keep visible cleanly!
                        if (t < 24.40)
                        {
                            DrawCursorAndClick(g, cursor.X, cursor.Y, clickProg, clickIsRight, clickBadge);
                        }
                    }

                    string framePath = Path.Combine(framesDir, string.Format("f_{0:D4}.png", f + 1));
                    frame.Save(framePath, ImageFormat.Png);

                    if (f == totalFrames - 10)
                        frame.Save(shotOutPath, ImageFormat.Png);

                    if (!string.IsNullOrEmpty(debugDirOrEmpty))
                    {
                        if (f == 54)  frame.Save(Path.Combine(debugDirOrEmpty, "install_chk_step1.png"), ImageFormat.Png); // ~3.6s Copied!
                        if (f == 135) frame.Save(Path.Combine(debugDirOrEmpty, "install_chk_step2.png"), ImageFormat.Png); // ~9.0s Start Right-Click menu
                        if (f == 189) frame.Save(Path.Combine(debugDirOrEmpty, "install_chk_step3.png"), ImageFormat.Png); // ~12.6s Hover/Click Terminal
                        if (f == 270) frame.Save(Path.Combine(debugDirOrEmpty, "install_chk_step4.png"), ImageFormat.Png); // ~18.0s Pasted command
                        if (f == 380) frame.Save(Path.Combine(debugDirOrEmpty, "install_chk_step5.png"), ImageFormat.Png); // ~25.3s SUCCESS + Done
                    }
                }
            }
        }
    }
}
'@

Add-Type -TypeDefinition $rendererCs -ReferencedAssemblies @('System.Drawing') -ErrorAction Stop

$copyBtnLocalRect = [InstallWalkthroughRenderer]::LocateTealCopyButton($browserShotPath)
Write-Host "[+] Located teal Copy button at viewport rect: X=$($copyBtnLocalRect.X) Y=$($copyBtnLocalRect.Y) W=$($copyBtnLocalRect.Width) H=$($copyBtnLocalRect.Height)" -ForegroundColor Green

$rawShotPath = Join-Path $workTemp 'shot_raw.png'
$shotPath = Join-Path $outDir 'shot.webp'
$debugDir = if ($SaveDebugFrames) { [System.IO.Path]::GetTempPath() } else { '' }
Write-Host "[+] Rendering 405 slow beginner-friendly frames (1920x1280 @ 15fps = 27.0s)..." -ForegroundColor Cyan
[InstallWalkthroughRenderer]::RenderAllFrames($browserShotPath, $copyBtnLocalRect, $framesDir, $rawShotPath, $addinVersion, $debugDir)

& $ffmpeg -hide_banner -loglevel error -y -i $rawShotPath `
    -vf "scale=1920:1280:flags=lanczos" -c:v libwebp -quality 86 $shotPath
if ($LASTEXITCODE -ne 0) { throw "FFmpeg WebP screenshot encoding failed." }

# -----------------------------------------------------------------------------
# 3. Encode MP4 + WebM and generate bilingual WebVTT captions + manifest sync
# -----------------------------------------------------------------------------
$mp4Path = Join-Path $outDir 'clip.mp4'
$webmPath = Join-Path $outDir 'clip.webm'
$vttEnPath = Join-Path $outDir 'captions.en.vtt'
$vttKoPath = Join-Path $outDir 'captions.ko.vtt'
$inputPattern = Join-Path $framesDir 'f_%04d.png'

Write-Host "[+] Encoding MP4 & WebM videos..." -ForegroundColor Cyan
& $ffmpeg -hide_banner -loglevel error -y -framerate 15 -i $inputPattern `
    -vf "format=yuv420p" -c:v libx264 -preset medium -crf 22 -movflags +faststart $mp4Path
if ($LASTEXITCODE -ne 0) { throw "FFmpeg MP4 encoding failed." }

& $ffmpeg -hide_banner -loglevel error -y -framerate 15 -i $inputPattern `
    -vf "format=yuv420p" -c:v libvpx-vp9 -b:v 0 -crf 32 -row-mt 1 $webmPath
if ($LASTEXITCODE -ne 0) { throw "FFmpeg WebM encoding failed." }

$vttEn = @(
    'WEBVTT',
    '',
    '00:00.000 --> 00:05.200',
    'Step 1: Click the teal Copy button on the install command box.',
    '',
    '00:05.200 --> 00:10.600',
    'Step 2: Right-click the Windows Start button on your taskbar.',
    '',
    '00:10.600 --> 00:15.000',
    'Step 3: Click "Terminal" (Win 11) or "Windows PowerShell" (Win 10) - no Admin rights needed.',
    '',
    '00:15.000 --> 00:19.800',
    'Step 4: Right-click inside the PowerShell window to paste automatically.',
    '',
    '00:19.800 --> 00:27.000',
    'Step 5: Press Enter, wait until [SUCCESS] appears, then reopen Office.'
)

$vttKo = @(
    'WEBVTT',
    '',
    '00:00.000 --> 00:05.200',
    '1단계: 설치 명령어 상자 오른쪽의 청록색 [복사] 버튼을 클릭합니다.',
    '',
    '00:05.200 --> 00:10.600',
    '2단계: 작업표시줄의 Windows 시작 버튼(⊞)을 마우스로 우클릭합니다.',
    '',
    '00:10.600 --> 00:15.000',
    '3단계: [터미널(Terminal)] 또는 [Windows PowerShell]을 클릭합니다 (관리자 권한 불필요).',
    '',
    '00:15.000 --> 00:19.800',
    '4단계: 열린 PowerShell 창 안쪽을 마우스로 우클릭해 명령어를 자동 붙여넣기합니다.',
    '',
    '00:19.800 --> 00:27.000',
    '5단계: Enter를 누르고 [SUCCESS] 메시지가 나오면 Excel/Word/PowerPoint를 다시 엽니다.'
)

[System.IO.File]::WriteAllLines($vttEnPath, $vttEn, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllLines($vttKoPath, $vttKo, [System.Text.UTF8Encoding]::new($false))

$dur = [Math]::Round([double](& $ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $mp4Path), 1)

$manifest = [ordered]@{}
if (Test-Path -LiteralPath $manifestPublicPath) {
    $existingJson = [System.IO.File]::ReadAllText($manifestPublicPath, [System.Text.UTF8Encoding]::new($false))
    $existingObj = $existingJson | ConvertFrom-Json
    foreach ($prop in $existingObj.PSObject.Properties) {
        $manifest[$prop.Name] = $prop.Value
    }
}

$manifest['common-install-walkthrough'] = [ordered]@{
    id         = 'common-install-walkthrough'
    host       = 'common'
    titleEn    = 'Mouse-First 5-Step PowerShell Installation Walkthrough'
    titleKo    = '마우스로 따라하는 5단계 PowerShell 설치 가이드'
    shot       = '/media/common-install-walkthrough/shot.webp'
    webm       = '/media/common-install-walkthrough/clip.webm'
    mp4        = '/media/common-install-walkthrough/clip.mp4'
    captionEn  = 'Step-by-step installation video: Copy command -> Right-click Start -> Open Terminal / PowerShell -> Right-click to paste -> Press Enter.'
    captionKo  = '5단계 설치 안내 영상: 복사 버튼 클릭 -> 시작 버튼 우클릭 -> Terminal / PowerShell 열기 -> 창 안쪽 우클릭 붙여넣기 -> Enter 실행.'
    vttEn      = '/media/common-install-walkthrough/captions.en.vtt'
    vttKo      = '/media/common-install-walkthrough/captions.ko.vtt'
    version    = $addinVersion
    capturedAt = (Get-Date -Format 'yyyy-MM-dd')
    width      = 1920
    height     = 1280
    durationSec = $dur
}

$outJson = $manifest | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText($manifestPublicPath, $outJson, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText($manifestThemePath, $outJson, [System.Text.UTF8Encoding]::new($false))

Remove-Item -LiteralPath $workTemp -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ("  [OK] common-install-walkthrough | WebP 1920x1280 ({0} KB) | MP4 {1} KB | WebM {2} KB | {3}s" -f `
    [int]((Get-Item $shotPath).Length / 1KB),
    [int]((Get-Item $mp4Path).Length / 1KB),
    [int]((Get-Item $webmPath).Length / 1KB),
    $dur) -ForegroundColor Green
