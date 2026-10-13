using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using Smiles2Img.Rendering;

namespace Smiles2Img.Tests
{
    public static class DocScreenshots
    {
        public static void GenerateAll(string baseDir)
        {
            var assetDir = Path.Combine(baseDir, "assets");
            var publicDir = Path.Combine(baseDir, "docs/public");
            Directory.CreateDirectory(assetDir);
            Directory.CreateDirectory(publicDir);

            Console.WriteLine("Generating documentation screenshots...");
            GenerateRibbon(assetDir, publicDir);
            GenerateContextMenu(assetDir, publicDir);
            GenerateKetcher(assetDir, publicDir);
            GenerateDiagnostics(assetDir, publicDir);
            GenerateReactionScheme(assetDir, publicDir);
            Console.WriteLine("All documentation screenshots generated successfully!");
        }

        private static void SaveBoth(Bitmap bmp, string assetDir, string publicDir, string filename)
        {
            var p1 = Path.Combine(assetDir, filename);
            var p2 = Path.Combine(publicDir, filename);
            bmp.Save(p1, ImageFormat.Png);
            bmp.Save(p2, ImageFormat.Png);
            Console.WriteLine($"  [+] Created {filename} ({bmp.Width}x{bmp.Height})");
        }

        private static void GenerateRibbon(string assetDir, string publicDir)
        {
            const int w = 1260;
            const int h = 260;
            using var bmp = new Bitmap(w, h);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // Office Header
            using var topBg = new SolidBrush(Color.FromArgb(16, 124, 65));
            g.FillRectangle(topBg, 0, 0, w, 44);

            using var fontRegular = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            using var fontBold = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            using var fontGroup = new Font("Segoe UI", 8.5f, FontStyle.Regular);
            using var fontBtnTitle = new Font("Malgun Gothic", 9.5f, FontStyle.Bold);
            using var fontBtnSub = new Font("Segoe UI", 8.2f, FontStyle.Regular);

            var tabs = new[] { "File", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View" };
            var tx = 25;
            foreach (var t in tabs)
            {
                g.DrawString(t, fontRegular, Brushes.White, tx, 12);
                tx += (int)g.MeasureString(t, fontRegular).Width + 24;
            }

            // Active Tab: SMILES
            var activeRect = new Rectangle(tx - 6, 6, 95, 38);
            g.FillRectangle(Brushes.White, activeRect);
            using var greenBrush = new SolidBrush(Color.FromArgb(16, 124, 65));
            g.DrawString("SMILES", fontBold, greenBrush, tx + 12, 12);

            // Ribbon Body
            g.FillRectangle(Brushes.White, 0, 44, w, h - 44);
            using var divPen = new Pen(Color.FromArgb(229, 231, 235), 1.5f);
            using var borderPen = new Pen(Color.FromArgb(209, 213, 219), 1f);
            using var hoverBrush = new SolidBrush(Color.FromArgb(248, 250, 252));
            using var darkBrush = new SolidBrush(Color.FromArgb(17, 24, 39));
            using var mutedBrush = new SolidBrush(Color.FromArgb(107, 114, 128));
            using var blueBrush = new SolidBrush(Color.FromArgb(37, 99, 235));

            g.DrawLine(divPen, 0, 44, w, 44);

            // Group 1: Structure & Sketch
            DrawBigButton(g, 25, 52, 110, 160, "✏️", "구조식 삽입", "Ketcher 2D", "(Ctrl+Shift+K)", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, greenBrush, hoverBrush, borderPen);
            DrawBigButton(g, 145, 52, 120, 160, "🧪", "반응식 캔버스", "=KET2IMG", "화살표·주석보존", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, greenBrush, hoverBrush, borderPen);
            DrawBigButton(g, 275, 52, 110, 160, "⚡", "퀵 인서트", "선택 텍스트", "0.1초 즉시 치환", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, greenBrush, hoverBrush, borderPen);
            g.DrawLine(divPen, 400, 52, 400, 225);
            g.DrawString("Structure & Sketch (구조식 편집)", fontGroup, mutedBrush, 120, 232);

            // Group 2: Clipboard & Data
            DrawBigButton(g, 420, 52, 115, 160, "📋", "클립보드 삽입", "SMILES/KET", "즉시 그림 생성", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, greenBrush, hoverBrush, borderPen);
            DrawSmallButton(g, 545, 56, 210, 42, "📄 SMILES 복사", "순수 문자열 추출", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, hoverBrush, borderPen);
            DrawSmallButton(g, 545, 108, 210, 42, "⚖️ 분자 정보 (Formula·MW)", "화학식 및 분자량 복사", fontBtnTitle, fontBtnSub, darkBrush, blueBrush, hoverBrush, borderPen);
            DrawSmallButton(g, 545, 160, 210, 42, "🔬 정밀 질량 (Exact Mass)", "고분해능 단일동위원소", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, hoverBrush, borderPen);
            g.DrawLine(divPen, 770, 52, 770, 225);
            g.DrawString("Clipboard & Molecular Properties (물성 추출)", fontGroup, mutedBrush, 480, 232);

            // Group 3: Diagnostics & Help
            DrawBigButton(g, 790, 52, 115, 160, "🔍", "진단 창 보기", "화학 파서 로그", "오류 위치 표시", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, greenBrush, hoverBrush, borderPen);
            DrawBigButton(g, 915, 52, 115, 160, "📖", "사용자 문서", "GitHub Pages", "수식 완전 가이드", fontBtnTitle, fontBtnSub, darkBrush, mutedBrush, blueBrush, hoverBrush, borderPen);
            g.DrawLine(divPen, 1045, 52, 1045, 225);
            g.DrawString("System Diagnostics & Help (지원)", fontGroup, mutedBrush, 850, 232);

            // Right Badge
            var badgeRect = new Rectangle(1060, 56, 180, 155);
            using var badgeBg = new SolidBrush(Color.FromArgb(239, 246, 255));
            using var badgeBorder = new Pen(Color.FromArgb(191, 219, 254));
            g.FillRectangle(badgeBg, badgeRect);
            g.DrawRectangle(badgeBorder, badgeRect);
            g.DrawString("SMILES for Office", fontBold, blueBrush, 1075, 70);
            g.DrawString("v1.5.0 Suite", fontGroup, mutedBrush, 1075, 95);
            g.DrawString("• Excel Add-in (XLL)", fontGroup, darkBrush, 1075, 120);
            g.DrawString("• Word Add-in (COM)", fontGroup, darkBrush, 1075, 142);
            g.DrawString("• PowerPoint Add-in", fontGroup, darkBrush, 1075, 164);

            SaveBoth(bmp, assetDir, publicDir, "ribbon-guide.png");
        }

        private static void DrawBigButton(Graphics g, int x, int y, int bw, int bh, string emoji, string title, string sub1, string sub2, Font fTitle, Font fSub, Brush bDark, Brush bMuted, Brush bAccent, Brush bBg, Pen pBorder)
        {
            var r = new Rectangle(x, y, bw, bh);
            g.FillRectangle(bBg, r);
            g.DrawRectangle(pBorder, r);
            using var fEmoji = new Font("Segoe UI Emoji", 24f);
            using var sf = new StringFormat { Alignment = StringAlignment.Center };
            g.DrawString(emoji, fEmoji, bDark, x + bw / 2f, y + 14, sf);
            g.DrawString(title, fTitle, bDark, x + bw / 2f, y + 74, sf);
            g.DrawString(sub1, fSub, bMuted, x + bw / 2f, y + 98, sf);
            g.DrawString(sub2, fSub, bAccent, x + bw / 2f, y + 118, sf);
        }

        private static void DrawSmallButton(Graphics g, int x, int y, int bw, int bh, string title, string sub, Font fTitle, Font fSub, Brush bDark, Brush bMuted, Brush bBg, Pen pBorder)
        {
            var r = new Rectangle(x, y, bw, bh);
            g.FillRectangle(bBg, r);
            g.DrawRectangle(pBorder, r);
            g.DrawString(title, fTitle, bDark, x + 12, y + 6);
            g.DrawString(sub, fSub, bMuted, x + 14, y + 23);
        }

        private static void GenerateContextMenu(string assetDir, string publicDir)
        {
            const int w = 1260;
            const int h = 580;
            using var bmp = new Bitmap(w, h);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            g.Clear(Color.FromArgb(248, 250, 252));

            using var titleFont = new Font("Malgun Gothic", 13.5f, FontStyle.Bold);
            using var subFont = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            using var darkBrush = new SolidBrush(Color.FromArgb(17, 24, 39));
            using var mutedBrush = new SolidBrush(Color.FromArgb(100, 116, 139));
            using var greenBrush = new SolidBrush(Color.FromArgb(16, 124, 65));
            using var blueBrush = new SolidBrush(Color.FromArgb(37, 99, 235));
            using var divPen = new Pen(Color.FromArgb(226, 232, 240), 1f);
            using var borderPen = new Pen(Color.FromArgb(203, 213, 225), 1f);
            using var shadowBrush = new SolidBrush(Color.FromArgb(24, 0, 0, 0));

            g.DrawString("통합 우클릭 컨텍스트 메뉴 — 스마트 서브메뉴 시스템 (Office Context Menu)", titleFont, darkBrush, 35, 18);
            g.DrawString("셀 또는 영역에 화학 구조식이 이미 존재하는지 여부를 자동 감지하여 최적화된 하위 메뉴를 전개합니다.", subFont, mutedBrush, 37, 48);

            // Left Box: Structure Selected
            var leftBox = new Rectangle(35, 80, 580, 465);
            g.FillRectangle(Brushes.White, leftBox);
            g.DrawRectangle(borderPen, leftBox);
            g.DrawString("Case A. 분자 구조식이 들어있는 셀 / 그림 우클릭", new Font("Malgun Gothic", 10.5f, FontStyle.Bold), greenBrush, 55, 96);
            g.DrawString("Ketcher 제자리 수정 및 원클릭 물성 계산(화학식·분자량·Exact Mass) 메뉴 활성화", new Font("Segoe UI", 8.5f), mutedBrush, 55, 120);

            // Left Context Menu
            DrawContextMenuBlock(g, 55, 150, 210, 310, true,
                "🧪 Chemical Structure", true, 92,
                shadowBrush, borderPen, divPen, darkBrush, mutedBrush, greenBrush);

            // Left Submenu
            DrawSubmenuBlock(g, 267, 210, 315, 260, shadowBrush, borderPen, divPen, darkBrush, blueBrush, mutedBrush,
                new[] {
                    ("✏️  Edit Structure in Ketcher", true, false),
                    ("───", false, true),
                    ("📋  Copy SMILES", false, false),
                    ("🧪  Copy KET (Lossless Canvas)", false, false),
                    ("───", false, true),
                    ("🔢  Copy Molecular Formula", false, false),
                    ("⚖️  Copy Molecular Weight (MW)", false, false),
                    ("🔬  Copy Exact Mass", false, false),
                    ("───", false, true),
                    ("💡 Formula: C8H10N4O2 | MW: 194.19", false, false)
                });

            // Right Box: Empty Space Selected
            var rightBox = new Rectangle(645, 80, 580, 465);
            g.FillRectangle(Brushes.White, rightBox);
            g.DrawRectangle(borderPen, rightBox);
            g.DrawString("Case B. 빈 셀 또는 빈 공간 우클릭", new Font("Malgun Gothic", 10.5f, FontStyle.Bold), blueBrush, 665, 96);
            g.DrawString("단 2개의 필수 기능만 표시하여 일반 Office 컨텍스트 메뉴를 깔끔하게 유지", new Font("Segoe UI", 8.5f), mutedBrush, 665, 120);

            // Right Context Menu
            DrawContextMenuBlock(g, 665, 150, 210, 270, false,
                "🧪 Chemical Structure", false, 92,
                shadowBrush, borderPen, divPen, darkBrush, mutedBrush, blueBrush);

            // Right Submenu (2 items)
            DrawSubmenuBlock(g, 877, 220, 305, 115, shadowBrush, borderPen, divPen, darkBrush, greenBrush, mutedBrush,
                new[] {
                    ("✏️  Insert Structure in Ketcher", true, false),
                    ("───", false, true),
                    ("📋  Paste from Clipboard", false, false)
                });

            SaveBoth(bmp, assetDir, publicDir, "context-menu-guide.png");
        }

        private static void DrawContextMenuBlock(Graphics g, int x, int y, int bw, int bh, bool isGreen, string trigTitle, bool trigGreen, int trigY, Brush shadow, Pen border, Pen div, Brush dark, Brush muted, Brush trigBrush)
        {
            g.FillRectangle(shadow, x + 4, y + 4, bw, bh);
            g.FillRectangle(Brushes.White, x, y, bw, bh);
            g.DrawRectangle(border, x, y, bw, bh);

            using var f = new Font("Segoe UI", 9.2f);
            using var fb = new Font("Segoe UI", 9.2f, FontStyle.Bold);

            g.DrawString("잘라내기 (Cut)", f, muted, x + 25, y + 14);
            g.DrawString("복사 (Copy)", f, muted, x + 25, y + 40);
            g.DrawString("붙여넣기 옵션...", f, muted, x + 25, y + 66);
            g.DrawLine(div, x + 8, y + 92, x + bw - 8, y + 92);

            // Trigger rect
            var tr = new Rectangle(x + 4, y + 97, bw - 8, 32);
            using var trBg = new SolidBrush(isGreen ? Color.FromArgb(220, 252, 231) : Color.FromArgb(239, 246, 255));
            g.FillRectangle(trBg, tr);
            g.DrawString(trigTitle, fb, trigBrush, x + 15, y + 104);
            g.DrawString("▶", fb, trigBrush, x + bw - 26, y + 104);

            g.DrawLine(div, x + 8, y + 135, x + bw - 8, y + 135);
            g.DrawString("셀 서식 (Format Cells...)", f, muted, x + 25, y + 146);
            g.DrawString("메모 삽입 (New Note)", f, muted, x + 25, y + 174);
            g.DrawString("스마트 조회 (Smart Lookup)", f, muted, x + 25, y + 202);
            if (bh > 280)
            {
                g.DrawString("링크 (Link)", f, muted, x + 25, y + 230);
            }
        }

        private static void DrawSubmenuBlock(Graphics g, int x, int y, int bw, int bh, Brush shadow, Pen border, Pen div, Brush dark, Brush accent, Brush muted, (string Text, bool Bold, bool Sep)[] items)
        {
            g.FillRectangle(shadow, x + 4, y + 4, bw, bh);
            g.FillRectangle(Brushes.White, x, y, bw, bh);
            g.DrawRectangle(border, x, y, bw, bh);

            using var f = new Font("Segoe UI", 9.2f);
            using var fb = new Font("Segoe UI", 9.2f, FontStyle.Bold);
            using var fi = new Font("Segoe UI", 8.2f, FontStyle.Italic);

            var cy = y + 12;
            foreach (var item in items)
            {
                if (item.Sep)
                {
                    g.DrawLine(div, x + 10, cy + 2, x + bw - 10, cy + 2);
                    cy += 8;
                }
                else
                {
                    var font = item.Bold ? fb : (item.Text.StartsWith("💡") ? fi : f);
                    var brush = item.Bold ? dark : (item.Text.StartsWith("💡") ? muted : (item.Text.Contains("Formula") || item.Text.Contains("Weight") || item.Text.Contains("Exact") || item.Text.Contains("Clipboard") ? accent : dark));
                    g.DrawString(item.Text, font, brush, x + 15, cy);
                    cy += 24;
                }
            }
        }

        private static void GenerateKetcher(string assetDir, string publicDir)
        {
            const int w = 1200;
            const int h = 720;
            using var bmp = new Bitmap(w, h);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            g.Clear(Color.FromArgb(241, 245, 249));

            // Outer Form
            var winRect = new Rectangle(15, 12, w - 30, h - 24);
            g.FillRectangle(Brushes.White, winRect);
            using var borderPen = new Pen(Color.FromArgb(203, 213, 225), 2f);
            g.DrawRectangle(borderPen, winRect);

            // Form Title Bar
            using var titleBg = new SolidBrush(Color.FromArgb(15, 23, 42));
            g.FillRectangle(titleBg, 15, 12, w - 30, 42);
            using var fTitle = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            g.DrawString("🧬 EPAM Ketcher v3.18.0 — 화학 분자 및 반응식 2D 에디터 (100% Offline WebView2)", fTitle, Brushes.White, 32, 22);

            // Property Bar
            using var propBg = new SolidBrush(Color.FromArgb(248, 250, 252));
            using var divPen = new Pen(Color.FromArgb(226, 232, 240), 1.5f);
            g.FillRectangle(propBg, 15, 54, w - 30, 52);
            g.DrawLine(divPen, 15, 106, w - 15, 106);

            using var fPill = new Font("Segoe UI", 9.2f, FontStyle.Bold);
            using var darkBrush = new SolidBrush(Color.FromArgb(17, 24, 39));
            using var blueBrush = new SolidBrush(Color.FromArgb(29, 78, 216));
            using var greenBrush = new SolidBrush(Color.FromArgb(22, 101, 52));
            using var orangeBrush = new SolidBrush(Color.FromArgb(194, 65, 12));
            using var purpleBrush = new SolidBrush(Color.FromArgb(126, 34, 206));
            using var mutedBrush = new SolidBrush(Color.FromArgb(100, 116, 139));

            DrawPill(g, 30, 64, 155, 32, "🧪 Caffeine (카페인)", fPill, darkBrush, Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225));
            DrawPill(g, 195, 64, 160, 32, "Formula: C8H10N4O2", fPill, blueBrush, Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254));
            DrawPill(g, 365, 64, 155, 32, "MW: 194.19 g/mol", fPill, greenBrush, Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208));
            DrawPill(g, 530, 64, 160, 32, "Exact Mass: 194.08", fPill, orangeBrush, Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170));
            DrawPill(g, 990, 64, 180, 32, "🌓 배경: 투명 (Transparent)", fPill, purpleBrush, Color.FromArgb(243, 232, 255), Color.FromArgb(216, 180, 254));

            // Left Periodic Toolbar
            var atoms = new[] { "H", "C", "N", "O", "S", "P", "F", "Cl", "Br", "I" };
            var atomColors = new[] {
                Color.FromArgb(50, 50, 50), Color.FromArgb(30, 41, 59), Color.FromArgb(37, 99, 235),
                Color.FromArgb(220, 38, 38), Color.FromArgb(202, 138, 4), Color.FromArgb(234, 88, 12),
                Color.FromArgb(16, 185, 129), Color.FromArgb(5, 150, 105), Color.FromArgb(180, 83, 9), Color.FromArgb(147, 51, 234)
            };
            using var fAtom = new Font("Segoe UI", 11f, FontStyle.Bold);
            using var pToolBorder = new Pen(Color.FromArgb(209, 213, 219));
            for (var i = 0; i < atoms.Length; i++)
            {
                var ay = 120 + i * 44;
                var ar = new Rectangle(28, ay, 44, 35);
                g.FillRectangle(Brushes.White, ar);
                g.DrawRectangle(pToolBorder, ar);
                using var b = new SolidBrush(atomColors[i]);
                using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(atoms[i], fAtom, b, ar, sf);
            }

            // Top Drawing Tools
            var tools = new[] { "Single ─", "Double ═", "Triple ≡", "Wedge ◢", "Hash ▰", "Benzene ⬡", "Ring-5 ⬠", "Arrow ➔", "Plus ➕", "Clear 🗑️" };
            using var fTool = new Font("Segoe UI", 8.8f);
            for (var t = 0; t < tools.Length; t++)
            {
                var tr = new Rectangle(88 + t * 105, 118, 98, 30);
                g.FillRectangle(Brushes.White, tr);
                g.DrawRectangle(pToolBorder, tr);
                using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(tools[t], fTool, darkBrush, tr, sf);
            }

            // Grid Dots
            using var dotBrush = new SolidBrush(Color.FromArgb(226, 232, 240));
            for (var gx = 100; gx < w - 40; gx += 45)
            {
                for (var gy = 165; gy < 610; gy += 45)
                {
                    g.FillEllipse(dotBrush, gx, gy, 3, 3);
                }
            }

            // Render Caffeine Molecule in Center
            const string caffeineSmiles = "Cn1cnc2c1c(=O)n(C)c(=O)n2C";
            var pngBytes = MoleculeRenderer.Render(caffeineSmiles, color: true);
            using var ms = new MemoryStream(pngBytes);
            using var molImg = Image.FromStream(ms);
            g.DrawImage(molImg, 380, 180, 480, 410);

            // Bottom Action Footer
            g.FillRectangle(propBg, 15, 630, w - 30, 78);
            g.DrawLine(divPen, 15, 630, w - 15, 630);

            // Apply Button
            var applyRect = new Rectangle(940, 645, 230, 48);
            using var applyBg = new SolidBrush(Color.FromArgb(22, 163, 74));
            g.FillRectangle(applyBg, applyRect);
            using var fBtn = new Font("Malgun Gothic", 11f, FontStyle.Bold);
            using var sfBtn = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("✅ 문서에 적용 (Ctrl+Enter)", fBtn, Brushes.White, applyRect, sfBtn);

            // Cancel Button
            var cancelRect = new Rectangle(835, 645, 95, 48);
            g.FillRectangle(Brushes.White, cancelRect);
            g.DrawRectangle(pToolBorder, cancelRect);
            using var fSubBtn = new Font("Malgun Gothic", 10f);
            g.DrawString("취소 (Esc)", fSubBtn, darkBrush, cancelRect, sfBtn);

            // Copy SMILES & Copy KET
            var cp1 = new Rectangle(30, 648, 140, 42);
            g.FillRectangle(Brushes.White, cp1);
            g.DrawRectangle(pToolBorder, cp1);
            g.DrawString("📋 SMILES 복사", fSubBtn, darkBrush, cp1, sfBtn);

            var cp2 = new Rectangle(180, 648, 145, 42);
            g.FillRectangle(Brushes.White, cp2);
            g.DrawRectangle(pToolBorder, cp2);
            g.DrawString("🧪 KET 캔버스 복사", fSubBtn, darkBrush, cp2, sfBtn);

            g.DrawString("Ready | EPAM Ketcher v3.18.0 & Indigo 1.49.0 | Local WebView2 Virtual Host (100% Offline)", new Font("Segoe UI", 8.8f), mutedBrush, 345, 662);

            SaveBoth(bmp, assetDir, publicDir, "ketcher-guide.png");
        }

        private static void DrawPill(Graphics g, int x, int y, int bw, int bh, string text, Font f, Brush bText, Color bg, Color border)
        {
            var r = new Rectangle(x, y, bw, bh);
            using var bgBrush = new SolidBrush(bg);
            using var borderPen = new Pen(border);
            using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.FillRectangle(bgBrush, r);
            g.DrawRectangle(borderPen, r);
            g.DrawString(text, f, bText, r, sf);
        }

        private static void GenerateDiagnostics(string assetDir, string publicDir)
        {
            const int w = 960;
            const int h = 540;
            using var bmp = new Bitmap(w, h);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            g.Clear(Color.FromArgb(241, 245, 249));

            var win = new Rectangle(18, 16, w - 36, h - 32);
            g.FillRectangle(Brushes.White, win);
            using var borderPen = new Pen(Color.FromArgb(203, 213, 225), 2f);
            g.DrawRectangle(borderPen, win);

            using var titleBg = new SolidBrush(Color.FromArgb(15, 23, 42));
            g.FillRectangle(titleBg, 18, 16, w - 36, 42);
            using var fTitle = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            g.DrawString("🔍 SMILES Diagnostics & Chemical Engine Log (추가 기능 진단 창)", fTitle, Brushes.White, 35, 27);

            var cBox = new Rectangle(32, 72, w - 64, h - 105);
            using var cBg = new SolidBrush(Color.FromArgb(15, 23, 42));
            g.FillRectangle(cBg, cBox);

            using var fLog = new Font("Consolas", 10f);
            using var logGreen = new SolidBrush(Color.FromArgb(74, 222, 128));
            using var logYellow = new SolidBrush(Color.FromArgb(250, 204, 21));
            using var logCyan = new SolidBrush(Color.FromArgb(56, 189, 248));
            using var logRed = new SolidBrush(Color.FromArgb(248, 113, 113));
            using var logWhite = new SolidBrush(Color.FromArgb(241, 245, 249));
            using var logMuted = new SolidBrush(Color.FromArgb(100, 116, 139));

            var ly = 88;
            g.DrawString("[SYSTEM] Smiles2Img Unified Office Suite v1.5.0 (net48 x64 Architecture)", fLog, logCyan, 48, ly); ly += 24;
            g.DrawString("[INDIGO] Native Chemical Engine: EPAM Indigo 1.49.0-rc.2 loaded & initialized", fLog, logGreen, 48, ly); ly += 24;
            g.DrawString("[KETCHER] Embedded WebView2 Ketcher v3.18.0 bundle SHA256 verified (69 files cached)", fLog, logGreen, 48, ly); ly += 24;
            g.DrawString("[COM] Registered WordAddIn & PowerPointAddIn in HKCU:\\Software\\Classes", fLog, logGreen, 48, ly); ly += 24;
            g.DrawString("[CACHE] Excel in-cell picture cache sheet '__SMILES2IMG' active (0 orphan records)", fLog, logGreen, 48, ly); ly += 24;
            g.DrawString("--------------------------------------------------------------------------------------", fLog, logMuted, 48, ly); ly += 22;
            g.DrawString("[VALIDATION] Checking chemical structure parser error reporting:", fLog, logWhite, 48, ly); ly += 24;
            g.DrawString("  Input:  'C1=CC=CC=C' (Invalid benzene ring without closing digit)", fLog, logYellow, 48, ly); ly += 24;
            g.DrawString("  Result: #VALUE! (Indigo parse error: cycle 1 not closed at position 9)", fLog, logRed, 48, ly); ly += 24;
            g.DrawString("[SUCCESS] Corrected formula '=SMILES2IMG(\"c1ccccc1\")' -> 300x300 PNG rendered in 2.1 ms", fLog, logGreen, 48, ly); ly += 24;
            g.DrawString("[INFO] Property calculated: Formula=C6H6, MW=78.11 g/mol, ExactMass=78.0469", fLog, logCyan, 48, ly); ly += 24;
            g.DrawString("[STATUS] Ready. All Office hooks active (Ribbon, ContextMenu, Double-Click).", fLog, logWhite, 48, ly);

            SaveBoth(bmp, assetDir, publicDir, "diagnostics-guide.png");
        }

        private static void GenerateReactionScheme(string assetDir, string publicDir)
        {
            const int w = 1200;
            const int h = 420;
            using var bmp = new Bitmap(w, h);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            g.Clear(Color.FromArgb(248, 250, 252));

            using var titleFont = new Font("Malgun Gothic", 13.5f, FontStyle.Bold);
            using var subFont = new Font("Segoe UI", 9.5f);
            using var darkBrush = new SolidBrush(Color.FromArgb(17, 24, 39));
            using var mutedBrush = new SolidBrush(Color.FromArgb(100, 116, 139));
            using var borderPen = new Pen(Color.FromArgb(203, 213, 225), 1.5f);
            using var greenBrush = new SolidBrush(Color.FromArgb(16, 124, 65));

            g.DrawString("화학 반응식 (Reaction Scheme) & KET 캔버스 수식 (=KET2IMG)", titleFont, darkBrush, 35, 18);
            g.DrawString("반응 화살표, 촉매 및 조건 텍스트, 플러스 기호(+)가 포함된 전체 유기 합성 경로를 셀 안에 무손실 보존합니다.", subFont, mutedBrush, 37, 48);

            // Container Box
            var box = new Rectangle(35, 80, w - 70, 310);
            g.FillRectangle(Brushes.White, box);
            g.DrawRectangle(borderPen, box);

            // Render Benzene
            var r1Bytes = MoleculeRenderer.Render("c1ccccc1", color: true);
            using var ms1 = new MemoryStream(r1Bytes);
            using var img1 = Image.FromStream(ms1);
            g.DrawImage(img1, 70, 110, 200, 200);

            // Plus Sign
            using var fSym = new Font("Segoe UI", 28f, FontStyle.Bold);
            using var fCond = new Font("Segoe UI", 11f, FontStyle.Bold);
            using var fSub = new Font("Segoe UI", 9.5f);
            g.DrawString("+", fSym, darkBrush, 300, 180);

            // Nitric Acid
            var r2Bytes = MoleculeRenderer.Render("[N+](=O)(O)[O-]", color: true);
            using var ms2 = new MemoryStream(r2Bytes);
            using var img2 = Image.FromStream(ms2);
            g.DrawImage(img2, 360, 130, 180, 160);

            // Reaction Arrow & Conditions
            using var arrowPen = new Pen(Color.FromArgb(37, 99, 235), 3.5f) { EndCap = LineCap.ArrowAnchor };
            g.DrawLine(arrowPen, 570, 210, 760, 210);
            g.DrawString("H2SO4 (cat.)", fCond, darkBrush, 620, 168);
            g.DrawString("55 °C, 1 hr", fSub, mutedBrush, 630, 222);

            // Product: Nitrobenzene
            var prodBytes = MoleculeRenderer.Render("c1ccccc1[N+](=O)[O-]", color: true);
            using var ms3 = new MemoryStream(prodBytes);
            using var img3 = Image.FromStream(ms3);
            g.DrawImage(img3, 790, 105, 230, 210);

            // Formula callout on bottom
            var calloutRect = new Rectangle(55, 335, w - 110, 42);
            using var calloutBg = new SolidBrush(Color.FromArgb(236, 253, 245));
            using var calloutBorder = new Pen(Color.FromArgb(167, 243, 208));
            g.FillRectangle(calloutBg, calloutRect);
            g.DrawRectangle(calloutBorder, calloutRect);
            g.DrawString("✨ Excel 수식 입력: =KET2IMG(A2)  ➔  더블 클릭 시 화살표와 촉매 텍스트까지 제자리에서 즉시 재편집 가능!", new Font("Malgun Gothic", 10f, FontStyle.Bold), greenBrush, 75, 346);

            SaveBoth(bmp, assetDir, publicDir, "reaction-scheme-guide.png");
        }
    }
}
