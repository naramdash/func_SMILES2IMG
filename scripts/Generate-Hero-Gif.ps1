<#
.SYNOPSIS
    Generates a crisp looping Hero GIF (hero-demo.gif) demonstrating:
    1) Typing =NAME2IMG(A2) in Excel to render a molecule
    2) Dragging down to auto-render structures, MW, and Lipinski Ro5 across rows
    3) Double-clicking a cell to open the built-in Ketcher 2D Sketcher
    Uses pure .NET (System.Drawing + GIF89a muxer) with zero external dependencies.
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

public static class GifMuxer {
    public static void SaveLoopingGif(string outputPath, List<Bitmap> frames, List<int> delaysCs) {
        using (var outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        using (var bw = new BinaryWriter(outStream)) {
            for (int i = 0; i < frames.Count; i++) {
                using (var ms = new MemoryStream()) {
                    frames[i].Save(ms, ImageFormat.Gif);
                    byte[] bytes = ms.ToArray();

                    // Parse GIF header & Global Color Table size
                    // bytes[0..5] = GIF89a, bytes[6..12] = Logical Screen Descriptor
                    byte packed = bytes[10];
                    bool hasGct = (packed & 0x80) != 0;
                    int gctSize = hasGct ? (3 * (1 << ((packed & 0x07) + 1))) : 0;
                    int headerAndGctLen = 13 + gctSize;

                    if (i == 0) {
                        // Write Header + Logical Screen Descriptor + GCT of frame 0
                        bw.Write(bytes, 0, headerAndGctLen);
                        // Write NETSCAPE2.0 infinite loop extension
                        bw.Write(new byte[] {
                            0x21, 0xFF, 0x0B,
                            (byte)'N',(byte)'E',(byte)'T',(byte)'S',(byte)'C',(byte)'A',(byte)'P',(byte)'E',
                            (byte)'2',(byte)'.',(byte)'0',
                            0x03, 0x01, 0x00, 0x00, 0x00
                        });
                    }

                    // Find Image Descriptor (0x2C) in frame bytes
                    int pos = headerAndGctLen;
                    while (pos < bytes.Length - 1) {
                        if (bytes[pos] == 0x21) {
                            // Skip existing extension block
                            pos += 2;
                            while (pos < bytes.Length && bytes[pos] != 0) {
                                pos += 1 + bytes[pos];
                            }
                            pos++; // skip block terminator 0x00
                        } else if (bytes[pos] == 0x2C) {
                            break;
                        } else {
                            pos++;
                        }
                    }

                    // Write Graphic Control Extension with frame delay (in centiseconds, 1/100s)
                    int delay = delaysCs[i];
                    bw.Write(new byte[] {
                        0x21, 0xF9, 0x04, 0x04,
                        (byte)(delay & 0xFF), (byte)((delay >> 8) & 0xFF),
                        0x00, 0x00
                    });

                    // Write Image Descriptor with Local Color Table (copied from this frame's GCT)
                    bw.Write((byte)0x2C);
                    bw.Write(bytes, pos + 1, 8); // Left, Top, Width, Height
                    byte imgPacked = bytes[pos + 9];
                    if (hasGct) {
                        imgPacked = (byte)((imgPacked & 0x78) | 0x80 | (packed & 0x07));
                        bw.Write(imgPacked);
                        bw.Write(bytes, 13, gctSize); // Local Color Table
                    } else {
                        bw.Write(imgPacked);
                    }

                    // Write LZW image data up to before trailer (0x3B)
                    int dataStart = pos + 10;
                    int dataLen = (bytes.Length - 1) - dataStart;
                    bw.Write(bytes, dataStart, dataLen);
                }
            }
            bw.Write((byte)0x3B); // GIF Trailer
        }
    }
}
"@

$repoRoot = Split-Path -Parent $PSScriptRoot
$showcaseBmp = [System.Drawing.Bitmap]::new((Join-Path $repoRoot "docs\public\excel-showcase.png"))
$matrixBmp   = [System.Drawing.Bitmap]::new((Join-Path $repoRoot "docs\public\excel-matrix-showcase.png"))

function Copy-Crop {
    param([System.Drawing.Bitmap]$Src, [int]$X, [int]$Y, [int]$W, [int]$H)
    $rect = [System.Drawing.Rectangle]::new($X, $Y, $W, $H)
    return $Src.Clone($rect, $Src.PixelFormat)
}

$compounds = @(
    @{ Name = "Aspirin";     Smiles = "CC(=O)Oc1ccccc1C(=O)O";                 Mw = "180.16"; Ro5 = "PASS (0/4)"; Bmp = (Copy-Crop $matrixBmp   1155 112 200 94) },
    @{ Name = "Ibuprofen";   Smiles = "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O";       Mw = "206.28"; Ro5 = "PASS (0/4)"; Bmp = (Copy-Crop $showcaseBmp 1236 560 218 104) },
    @{ Name = "Caffeine";    Smiles = "Cn1cnc2c1c(=O)n(C)c(=O)n2C";            Mw = "194.19"; Ro5 = "PASS (0/4)"; Bmp = (Copy-Crop $showcaseBmp 1236 673 218 104) },
    @{ Name = "Pyridine";    Smiles = "c1ccncc1";                              Mw = "79.10";  Ro5 = "PASS (0/4)"; Bmp = (Copy-Crop $showcaseBmp 1236 447 218 104) },
    @{ Name = "Benzene";     Smiles = "c1ccccc1";                              Mw = "78.11";  Ro5 = "PASS (0/4)"; Bmp = (Copy-Crop $showcaseBmp 1236 221 218 104) }
)

$molBitmaps = $compounds | ForEach-Object { $_.Bmp }
$aspirinLargeBmp = Copy-Crop $matrixBmp 1155 112 200 94
$showcaseBmp.Dispose()
$matrixBmp.Dispose()

$ketcherBgPath = Join-Path $repoRoot "docs\public\ketcher-guide.png"
$ketcherBaseBmp = [System.Drawing.Bitmap]::new($ketcherBgPath)

function New-RoundedRectPath {
    param([int]$X, [int]$Y, [int]$W, [int]$H, [int]$R)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $R * 2
    $path.AddArc($X, $Y, $d, $d, 180, 90)
    $path.AddArc($X + $W - $d, $Y, $d, $d, 270, 90)
    $path.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0, 90)
    $path.AddArc($X, $Y + $H - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

$states = @(
    @{ Step = 1; Typed = '=NAME2';            Rows = 0; Ripple = 0; Ketcher = $false; DelayCs = 20 },
    @{ Step = 1; Typed = '=NAME2IMG(';        Rows = 0; Ripple = 0; Ketcher = $false; DelayCs = 20 },
    @{ Step = 1; Typed = '=NAME2IMG(A2)';     Rows = 0; Ripple = 0; Ketcher = $false; DelayCs = 35 },
    @{ Step = 1; Typed = '=NAME2IMG(A2)';     Rows = 1; Ripple = 0; Ketcher = $false; DelayCs = 85 },
    @{ Step = 2; Typed = '=NAME2IMG(A2)';     Rows = 2; Ripple = 0; Ketcher = $false; DelayCs = 28 },
    @{ Step = 2; Typed = '=NAME2IMG(A2)';     Rows = 3; Ripple = 0; Ketcher = $false; DelayCs = 28 },
    @{ Step = 2; Typed = '=NAME2IMG(A2)';     Rows = 4; Ripple = 0; Ketcher = $false; DelayCs = 28 },
    @{ Step = 2; Typed = '=NAME2IMG(A2)';     Rows = 5; Ripple = 0; Ketcher = $false; DelayCs = 120 },
    @{ Step = 3; Typed = '=NAME2IMG(A2)';     Rows = 5; Ripple = 1; Ketcher = $false; DelayCs = 22 },
    @{ Step = 3; Typed = '=NAME2IMG(A2)';     Rows = 5; Ripple = 2; Ketcher = $false; DelayCs = 22 },
    @{ Step = 3; Typed = '=NAME2IMG(A2)';     Rows = 5; Ripple = 0; Ketcher = $true;  DelayCs = 210 }
)

$width = 960
$height = 600
$framesList = New-Object 'System.Collections.Generic.List[System.Drawing.Bitmap]'
$delaysList = New-Object 'System.Collections.Generic.List[int]'
$emDash = [char]0x2014
$arrow  = [char]0x279C

foreach ($st in $states) {
    $bmp = New-Object System.Drawing.Bitmap $width, $height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $g.Clear([System.Drawing.Color]::White)

    # 1. Excel Title Bar
    $excelGreen = [System.Drawing.ColorTranslator]::FromHtml("#107C41")
    $excelBrush = New-Object System.Drawing.SolidBrush $excelGreen
    $g.FillRectangle($excelBrush, 0, 0, $width, 46)
    $excelBrush.Dispose()

    $fontTitle = New-Object System.Drawing.Font("Segoe UI", 11.5, [System.Drawing.FontStyle]::Bold)
    $fontSmall = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
    $fontSmallBold = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $fontMono = New-Object System.Drawing.Font("Consolas", 11.5, [System.Drawing.FontStyle]::Bold)

    $g.DrawString("Microsoft Excel  $emDash  SAR_Screening_Table.xlsx", $fontTitle, [System.Drawing.Brushes]::White, 18, 12)

    $stepText = if ($st.Step -eq 1) { "1/3  Type =NAME2IMG(A2) in Cell" }
                elseif ($st.Step -eq 2) { "2/3  Drag Down $arrow Auto-Render Structures & MW" }
                else { "3/3  Double-Click Cell $arrow Edit in 2D Sketcher" }
    $pillPath = New-RoundedRectPath 545 8 398 30 15
    $g.FillPath([System.Drawing.Brushes]::White, $pillPath)
    $pillPath.Dispose()
    $tealTextBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#0F766E"))
    $g.DrawString($stepText, $fontSmallBold, $tealTextBrush, 562, 14)
    $tealTextBrush.Dispose()

    # 2. Formula Bar (Y = 46..90)
    $fBarBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#F1F5F9"))
    $g.FillRectangle($fBarBrush, 0, 46, $width, 44)
    $fBarBrush.Dispose()
    $borderPen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#CBD5E1"), 1)
    $g.DrawLine($borderPen, 0, 90, $width, 90)

    $g.FillRectangle([System.Drawing.Brushes]::White, 12, 53, 64, 30)
    $g.DrawRectangle($borderPen, 12, 53, 64, 30)
    $g.DrawString("C2", $fontSmallBold, [System.Drawing.Brushes]::Black, 32, 59)

    $fxFont = New-Object System.Drawing.Font("Georgia", 11, ([System.Drawing.FontStyle]::Italic -bor [System.Drawing.FontStyle]::Bold))
    $g.DrawString("fx", $fxFont, [System.Drawing.Brushes]::DimGray, 88, 58)
    $fxFont.Dispose()

    $g.FillRectangle([System.Drawing.Brushes]::White, 118, 53, 828, 30)
    $g.DrawRectangle($borderPen, 118, 53, 828, 30)

    $fBarText = if ($st.Rows -eq 0) { $st.Typed + "|" } else { '=NAME2IMG(A2)     // + =SMILES2MW(B2) & =SMILES2LIPINSKI(B2)' }
    $formulaBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#0F766E"))
    $g.DrawString($fBarText, $fontMono, $formulaBrush, 128, 58)
    $formulaBrush.Dispose()

    # 3. Spreadsheet Grid
    $colX = @(0, 42, 200, 450, 650, 780, 960)
    $colHeaders = @("", "A : Compound Name", "B : SMILES", "C : 2D Structure", "D : MW (g/mol)", "E : Lipinski Ro5")

    $headerY = 90
    $headerH = 34
    $rowStartY = $headerY + $headerH
    $rowH = 94

    $colHeaderBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#E2E8F0"))
    $g.FillRectangle($colHeaderBrush, 0, $headerY, $width, $headerH)
    $colHeaderBrush.Dispose()

    for ($cIdx = 0; $cIdx -lt 6; $cIdx++) {
        $x1 = $colX[$cIdx]
        $w = $colX[$cIdx + 1] - $x1
        $g.DrawRectangle($borderPen, $x1, $headerY, $w, $headerH)
        if ($colHeaders[$cIdx] -ne "") {
            $headerTextBrush = if ($cIdx -ge 3) {
                New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#0F766E"))
            } else {
                New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#1E293B"))
            }
            $g.DrawString($colHeaders[$cIdx], $fontSmallBold, $headerTextBrush, $x1 + 10, $headerY + 8)
            $headerTextBrush.Dispose()
        }
    }

    for ($rIdx = 0; $rIdx -lt 5; $rIdx++) {
        $ry = $rowStartY + ($rIdx * $rowH)
        $comp = $compounds[$rIdx]

        $g.FillRectangle([System.Drawing.Brushes]::White, 42, $ry, $width - 42, $rowH)

        $rowNumBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#F1F5F9"))
        $g.FillRectangle($rowNumBrush, 0, $ry, 42, $rowH)
        $rowNumBrush.Dispose()
        $g.DrawRectangle($borderPen, 0, $ry, 42, $rowH)
        $g.DrawString("$($rIdx + 2)", $fontSmallBold, [System.Drawing.Brushes]::DimGray, 14, $ry + 36)

        for ($cIdx = 1; $cIdx -lt 6; $cIdx++) {
            $g.DrawRectangle($borderPen, $colX[$cIdx], $ry, $colX[$cIdx + 1] - $colX[$cIdx], $rowH)
        }

        $g.DrawString($comp.Name, $fontTitle, [System.Drawing.Brushes]::Black, $colX[1] + 14, $ry + 35)

        $smilesShort = if ($comp.Smiles.Length -gt 25) { $comp.Smiles.Substring(0, 23) + ".." } else { $comp.Smiles }
        $g.DrawString($smilesShort, $fontSmall, [System.Drawing.Brushes]::DarkSlateGray, $colX[2] + 10, $ry + 38)

        if ($rIdx -lt $st.Rows) {
            $molBmp = $molBitmaps[$rIdx]
            $destRect = [System.Drawing.Rectangle]::new($colX[3] + 18, $ry + 6, 164, 82)
            $g.DrawImage($molBmp, $destRect)

            $g.DrawString("$($comp.Mw)", $fontTitle, [System.Drawing.Brushes]::Black, $colX[4] + 32, $ry + 35)

            $passPath = New-RoundedRectPath ($colX[5] + 24) ($ry + 31) 128 30 15
            $passBg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#DCFCE7"))
            $g.FillPath($passBg, $passPath)
            $passBg.Dispose()
            $passPath.Dispose()
            $passTextBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#15803D"))
            $g.DrawString("PASS (0/4)", $fontSmallBold, $passTextBrush, $colX[5] + 46, $ry + 37)
            $passTextBrush.Dispose()
        } elseif ($rIdx -eq 0 -and $st.Rows -eq 0) {
            $g.DrawString($st.Typed, $fontMono, [System.Drawing.Brushes]::Teal, $colX[3] + 12, $ry + 35)
        }
    }

    # Selection Box & Fill Handle
    $selPen = New-Object System.Drawing.Pen($excelGreen, 3)
    $selRows = if ($st.Step -eq 2) { [Math]::Max(1, $st.Rows) } else { 1 }
    $selH = $selRows * $rowH
    $g.DrawRectangle($selPen, $colX[3], $rowStartY, ($colX[4] - $colX[3]), $selH)

    $fhX = $colX[4] - 5
    $fhY = $rowStartY + $selH - 5
    $fhBrush = New-Object System.Drawing.SolidBrush($excelGreen)
    $g.FillRectangle($fhBrush, $fhX, $fhY, 10, 10)
    $fhBrush.Dispose()
    $selPen.Dispose()

    if ($st.Ripple -gt 0) {
        $rippleR = 20 + ($st.Ripple * 14)
        $ripplePen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#0F766E"), 3)
        $cx = $colX[3] + 100
        $cy = $rowStartY + 46
        $g.DrawEllipse($ripplePen, $cx - $rippleR, $cy - $rippleR, $rippleR * 2, $rippleR * 2)
        $ripplePen.Dispose()
    }

    if ($st.Ketcher) {
        $mw = 780
        $mh = 468
        $mx = [int](($width - $mw) / 2)
        $my = 76

        # Solid dark slate backdrop border
        $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#1E293B"))
        $g.FillRectangle($shadowBrush, $mx - 4, $my - 4, $mw + 8, $mh + 8)
        $shadowBrush.Dispose()

        $g.DrawImage($ketcherBaseBmp, [System.Drawing.Rectangle]::new($mx, $my, $mw, $mh))
        $g.DrawImage($aspirinLargeBmp, [System.Drawing.Rectangle]::new($mx + 200, $my + 115, 380, 220))

        $modalPen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml("#0D9488"), 3)
        $g.DrawRectangle($modalPen, $mx, $my, $mw, $mh)
        $modalPen.Dispose()

        $bannerPath = New-RoundedRectPath ($mx + 135) ($my + $mh - 54) 510 38 19
        $bannerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml("#0F766E"))
        $g.FillPath($bannerBrush, $bannerPath)
        $bannerBrush.Dispose()
        $bannerPath.Dispose()
        $g.DrawString("Double-Click Any Cell  $arrow  Built-in Ketcher 2D Sketcher Opens!", $fontSmallBold, [System.Drawing.Brushes]::White, $mx + 168, $my + $mh - 43)
    }

    $borderPen.Dispose()
    $fontTitle.Dispose()
    $fontSmall.Dispose()
    $fontSmallBold.Dispose()
    $fontMono.Dispose()
    $g.Dispose()

    $framesList.Add($bmp)
    $delaysList.Add($st.DelayCs)
}

$outPublicGif = Join-Path $repoRoot "docs\public\hero-demo.gif"
$outAssetsGif = Join-Path $repoRoot "assets\hero-demo.gif"

[GifMuxer]::SaveLoopingGif($outPublicGif, $framesList, $delaysList)
Copy-Item $outPublicGif $outAssetsGif -Force

foreach ($b in $framesList) { $b.Dispose() }
foreach ($b in $molBitmaps) { $b.Dispose() }
$aspirinLargeBmp.Dispose()
$ketcherBaseBmp.Dispose()

Write-Host "[+] Generated $outPublicGif ($([math]::Round((Get-Item $outPublicGif).Length / 1KB, 1)) KB)" -ForegroundColor Green
