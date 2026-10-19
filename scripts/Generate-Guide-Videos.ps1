<#
.SYNOPSIS
    Generates edited high-definition (1080p) guide videos with subtitles and editing
    for Excel, Word, and PowerPoint add-ins using System.Drawing and FFmpeg.
#>
param(
    [string]$TargetApp = "All" # "Excel", "Word", "PowerPoint", or "All"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Ensure FFmpeg is in path
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") + ";" + $env:Path
$ffmpeg = (Get-Command ffmpeg -ErrorAction SilentlyContinue)?.Source
if (-not $ffmpeg) {
    throw "FFmpeg could not be found. Please ensure FFmpeg is installed."
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$assetsDir = Join-Path $repoRoot "assets"
$videosDir = Join-Path $assetsDir "videos"
if (-not (Test-Path $videosDir)) { New-Item -ItemType Directory -Path $videosDir -Force | Out-Null }

Write-Host "=== SMILES Office Suite Guide Video Generator ===" -ForegroundColor Cyan
Write-Host "FFmpeg path: $ffmpeg" -ForegroundColor Green
Write-Host "Output directory: $videosDir" -ForegroundColor Green

function New-SceneFrame {
    param(
        [string]$AppTitle,
        [string]$AppColorHex,
        [string]$AppIcon,
        [int]$StepNumber,
        [int]$TotalSteps,
        [string]$BadgeText,
        [string]$MainTitle,
        [string]$SubTitle,
        [string]$ImagePath,
        [string[]]$KeyPoints,
        [string]$FormulaCode,
        [string]$SubtitleText,
        [string]$OutputPath
    )

    $width = 1920
    $height = 1080
    $bmp = New-Object System.Drawing.Bitmap $width, $height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $appColor = [System.Drawing.ColorTranslator]::FromHtml($AppColorHex)
    $darkBg = [System.Drawing.ColorTranslator]::FromHtml("#0F172A")
    $cardBg = [System.Drawing.Color]::FromArgb(240, 255, 255, 255)
    $cardDark = [System.Drawing.Color]::FromArgb(230, 30, 41, 59)

    # 1. Background Gradient (Deep modern slate)
    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush (
        ([System.Drawing.Point]::new(0, 0)),
        ([System.Drawing.Point]::new(0, 1080)),
        [System.Drawing.ColorTranslator]::FromHtml("#0B132B"),
        [System.Drawing.ColorTranslator]::FromHtml("#1E293B")
    )
    $g.FillRectangle($bgBrush, 0, 0, $width, $height)
    $bgBrush.Dispose()

    # Subtle Accent Glow Top-Right
    $glowBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(35, $appColor.R, $appColor.G, $appColor.B))
    $g.FillEllipse($glowBrush, 1300, -100, 800, 500)
    $glowBrush.Dispose()

    # 2. Top Header Bar
    $headerRect = [System.Drawing.Rectangle]::new(0, 0, $width, 85)
    $headerBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 15, 23, 42))
    $g.FillRectangle($headerBrush, $headerRect)
    $headerBrush.Dispose()

    # Header Divider Line with Brand Color
    $headerLinePen = New-Object System.Drawing.Pen ($appColor, 4)
    $g.DrawLine($headerLinePen, 0, 85, $width, 85)
    $headerLinePen.Dispose()

    # Header App Brand Icon & Title
    $fontHeader = New-Object System.Drawing.Font ("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
    $fontHeaderSub = New-Object System.Drawing.Font ("Malgun Gothic", 14, [System.Drawing.FontStyle]::Regular)
    $g.DrawString("$AppIcon $AppTitle", $fontHeader, [System.Drawing.Brushes]::White, 50, 24)

    # Progress Indicator in Header
    if ($TotalSteps -gt 1) {
        $progressText = "STEP $StepNumber / $TotalSteps"
        $fontProgress = New-Object System.Drawing.Font ("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
        $badgePen = New-Object System.Drawing.Pen ($appColor, 2)
        $badgeBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(60, $appColor.R, $appColor.G, $appColor.B))
        $g.FillRectangle($badgeBrush, 1720, 24, 150, 36)
        $g.DrawRectangle($badgePen, 1720, 24, 150, 36)
        $sfCenter = New-Object System.Drawing.StringFormat
        $sfCenter.Alignment = [System.Drawing.StringAlignment]::Center
        $sfCenter.LineAlignment = [System.Drawing.StringAlignment]::Center
        $g.DrawString($progressText, $fontProgress, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(1720, 24, 150, 36), $sfCenter)
        $badgePen.Dispose()
        $badgeBrush.Dispose()
        $fontProgress.Dispose()
    }

    # 3. Main Title Area (Y: 110 ~ 210)
    $fontBadge = New-Object System.Drawing.Font ("Malgun Gothic", 12, [System.Drawing.FontStyle]::Bold)
    $fontMainTitle = New-Object System.Drawing.Font ("Malgun Gothic", 26, [System.Drawing.FontStyle]::Bold)
    $fontSubTitle = New-Object System.Drawing.Font ("Malgun Gothic", 15, [System.Drawing.FontStyle]::Regular)

    # Badge Pill
    $badgeW = [int][Math]::Max(160, ($BadgeText.Length * 18 + 30))
    $badgeRect = [System.Drawing.Rectangle]::new(60, 110, $badgeW, 32)
    $badgeBgBrush = New-Object System.Drawing.SolidBrush ($appColor)
    $g.FillRectangle($badgeBgBrush, $badgeRect)
    $badgeBgBrush.Dispose()
    $sfLeft = New-Object System.Drawing.StringFormat
    $sfLeft.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($BadgeText, $fontBadge, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]$badgeRect, $sfCenter)

    $g.DrawString($MainTitle, $fontMainTitle, [System.Drawing.Brushes]::White, 60, 152)
    $mutedTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#94A3B8"))
    $g.DrawString($SubTitle, $fontSubTitle, $mutedTextBrush, 64, 202)
    $mutedTextBrush.Dispose()

    # 4. Content Area Layout (Y: 250 ~ 910)
    # Left Side: Key Points & Formulas (Width: 700)
    # Right Side: Visual Showcase Image / Screenshot (Width: 1080)

    # --- LEFT PANEL: Cards, Key Points & Formula ---
    $leftCardX = 60
    $leftCardY = 255
    $leftCardW = 720
    $leftCardH = 650

    $cardBorderPen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#334155"), 2)
    $cardFillBrush = New-Object System.Drawing.SolidBrush ($cardDark)
    $g.FillRectangle($cardFillBrush, $leftCardX, $leftCardY, $leftCardW, $leftCardH)
    $g.DrawRectangle($cardBorderPen, $leftCardX, $leftCardY, $leftCardW, $leftCardH)

    # Card Header
    $fontCardHeader = New-Object System.Drawing.Font ("Malgun Gothic", 17, [System.Drawing.FontStyle]::Bold)
    $g.DrawString("💡 핵심 가이드 & 주요 포인트", $fontCardHeader, [System.Drawing.Brushes]::White, ($leftCardX + 25), ($leftCardY + 25))

    # Card Separator
    $g.DrawLine($cardBorderPen, ($leftCardX + 25), ($leftCardY + 65), ($leftCardX + $leftCardW - 25), ($leftCardY + 65))

    # Formula Box (if provided)
    $currY = $leftCardY + 85
    if ($FormulaCode) {
        $formulaBoxW = $leftCardW - 50
        $formulaBoxH = 100
        $codeBoxBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#090D16"))
        $codeBorderPen = New-Object System.Drawing.Pen ($appColor, 1.5)
        $g.FillRectangle($codeBoxBrush, ($leftCardX + 25), $currY, $formulaBoxW, $formulaBoxH)
        $g.DrawRectangle($codeBorderPen, ($leftCardX + 25), $currY, $formulaBoxW, $formulaBoxH)

        $fontCodeLabel = New-Object System.Drawing.Font ("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $fontCode = New-Object System.Drawing.Font ("Consolas", 14, [System.Drawing.FontStyle]::Bold)
        $labelBrush = New-Object System.Drawing.SolidBrush ($appColor)
        $codeBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#38BDF8"))
        $g.DrawString("실전 수식 / 실행 명령:", $fontCodeLabel, $labelBrush, ($leftCardX + 40), ($currY + 12))
        $g.DrawString($FormulaCode, $fontCode, $codeBrush, ($leftCardX + 40), ($currY + 38))
        
        $codeBoxBrush.Dispose()
        $codeBorderPen.Dispose()
        $fontCodeLabel.Dispose()
        $fontCode.Dispose()
        $labelBrush.Dispose()
        $codeBrush.Dispose()

        $currY += ($formulaBoxH + 25)
    }

    # Bullet Points
    $fontBulletTitle = New-Object System.Drawing.Font ("Malgun Gothic", 14, [System.Drawing.FontStyle]::Bold)
    $fontBulletDesc = New-Object System.Drawing.Font ("Malgun Gothic", 12.5, [System.Drawing.FontStyle]::Regular)
    $bulletNumBrush = New-Object System.Drawing.SolidBrush ($appColor)
    $bulletDescBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#E2E8F0"))

    $ptIndex = 1
    foreach ($pt in $KeyPoints) {
        if ($pt.Contains("::")) {
            $parts = $pt -split "::", 2
            $bTitle = $parts[0].Trim()
            $bDesc = $parts[1].Trim()
        } else {
            $bTitle = "포인트 $ptIndex"
            $bDesc = $pt
        }

        # Bullet Icon
        $g.FillEllipse($bulletNumBrush, ($leftCardX + 25), ($currY + 4), 16, 16)
        $g.DrawString($bTitle, $fontBulletTitle, [System.Drawing.Brushes]::White, ($leftCardX + 52), $currY)
        $currY += 32

        # Description wrapped
        $descRect = [System.Drawing.RectangleF]::new(($leftCardX + 52), $currY, ($leftCardW - 85), 65)
        $g.DrawString($bDesc, $fontBulletDesc, $bulletDescBrush, $descRect)
        $currY += 58
        $ptIndex++
    }

    $cardFillBrush.Dispose()
    $cardBorderPen.Dispose()
    $fontCardHeader.Dispose()
    $fontBulletTitle.Dispose()
    $fontBulletDesc.Dispose()
    $bulletNumBrush.Dispose()
    $bulletDescBrush.Dispose()

    # --- RIGHT PANEL: Visual Showcase Image (Width: 1060, Height: 650) ---
    $rightX = 810
    $rightY = 255
    $rightW = 1050
    $rightH = 650

    $imgFrameBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#090D16"))
    $imgFramePen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#334155"), 2)
    $g.FillRectangle($imgFrameBrush, $rightX, $rightY, $rightW, $rightH)
    $g.DrawRectangle($imgFramePen, $rightX, $rightY, $rightW, $rightH)

    if ($ImagePath -and (Test-Path $ImagePath)) {
        try {
            $srcImg = [System.Drawing.Image]::FromFile($ImagePath)
            
            # Compute aspect ratio fit
            $pad = 20
            $availW = $rightW - ($pad * 2)
            $availH = $rightH - ($pad * 2)

            $scaleW = $availW / $srcImg.Width
            $scaleH = $availH / $srcImg.Height
            $scale = [Math]::Min($scaleW, $scaleH)

            $destW = [int]($srcImg.Width * $scale)
            $destH = [int]($srcImg.Height * $scale)
            $destX = [int]($rightX + ($rightW - $destW) / 2)
            $destY = [int]($rightY + ($rightH - $destH) / 2)

            $g.DrawImage($srcImg, $destX, $destY, $destW, $destH)
            $srcImg.Dispose()
        } catch {
            Write-Warning "Could not render image $ImagePath : $_"
        }
    }

    $imgFrameBrush.Dispose()
    $imgFramePen.Dispose()

    # 5. BOTTOM SUBTITLE BAR (Y: 935 ~ 1050, Height: 115)
    # Translucent Frosted Dark Bar with Vibrant Highlight Pill & Crisp Subtitles
    $subBarRect = [System.Drawing.Rectangle]::new(50, 935, 1820, 115)
    $subBarBgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(245, 15, 23, 42))
    $subBarBorderPen = New-Object System.Drawing.Pen ($appColor, 2.5)
    $g.FillRectangle($subBarBgBrush, $subBarRect)
    $g.DrawRectangle($subBarBorderPen, $subBarRect)

    # Subtitle Left Pill: "자막 [CC]"
    $ccPillRect = [System.Drawing.Rectangle]::new(75, 955, 130, 40)
    $ccPillBrush = New-Object System.Drawing.SolidBrush ($appColor)
    $g.FillRectangle($ccPillBrush, $ccPillRect)
    $fontCc = New-Object System.Drawing.Font ("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $g.DrawString("자막 [CC]", $fontCc, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(75, 955, 130, 40), $sfCenter)
    $ccPillBrush.Dispose()
    $fontCc.Dispose()

    # Subtitle Text
    $fontSubText = New-Object System.Drawing.Font ("Malgun Gothic", 16.5, [System.Drawing.FontStyle]::Bold)
    $subTextBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#F8FAFC"))
    $subTextRect = [System.Drawing.RectangleF]::new(225, 950, 1610, 85)
    $g.DrawString($SubtitleText, $fontSubText, $subTextBrush, $subTextRect, $sfLeft)

    $subBarBgBrush.Dispose()
    $subBarBorderPen.Dispose()
    $fontSubText.Dispose()
    $subTextBrush.Dispose()

    # Cleanup top fonts
    $fontHeader.Dispose()
    $fontHeaderSub.Dispose()
    $fontBadge.Dispose()
    $fontMainTitle.Dispose()
    $fontSubTitle.Dispose()
    $sfCenter.Dispose()
    $sfLeft.Dispose()
    $g.Dispose()

    $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Build-VideoFromScenes {
    param(
        [string]$VideoName,
        [array]$Scenes,
        [string]$AppTitle,
        [string]$AppColor,
        [string]$AppIcon
    )

    $workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("smiles_vid_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    $outputMp4 = Join-Path $videosDir "$VideoName.mp4"

    Write-Host "`n--- Generating Frames for $VideoName ---" -ForegroundColor Cyan
    $concatListPath = Join-Path $workDir "concat.txt"
    $concatContent = @()

    $sceneIdx = 1
    $total = $Scenes.Count

    foreach ($sc in $Scenes) {
        $framePng = Join-Path $workDir ("scene_{0:D2}.png" -f $sceneIdx)
        
        New-SceneFrame `
            -AppTitle $AppTitle `
            -AppColorHex $AppColor `
            -AppIcon $AppIcon `
            -StepNumber $sceneIdx `
            -TotalSteps $total `
            -BadgeText $sc.Badge `
            -MainTitle $sc.Title `
            -SubTitle $sc.Subtitle `
            -ImagePath $sc.Image `
            -KeyPoints $sc.Points `
            -FormulaCode $sc.Formula `
            -SubtitleText $sc.SubtitleBar `
            -OutputPath $framePng

        $duration = if ($sc.Duration) { $sc.Duration } else { 14 }
        # Concat format: file 'path', duration X
        $escapedPng = $framePng.Replace("\", "/")
        $concatContent += "file '$escapedPng'"
        $concatContent += "duration $duration"

        Write-Host "  [+] Scene $sceneIdx / $total created ($duration s): $($sc.Title)" -ForegroundColor Green
        $sceneIdx++
    }

    # Add last image one more time for ffmpeg concat requirement
    $lastPng = (Join-Path $workDir ("scene_{0:D2}.png" -f ($sceneIdx - 1))).Replace("\", "/")
    $concatContent += "file '$lastPng'"

    [System.IO.File]::WriteAllLines($concatListPath, $concatContent)

    Write-Host "--- Encoding $VideoName.mp4 with FFmpeg ---" -ForegroundColor Cyan
    if (Test-Path $outputMp4) { Remove-Item $outputMp4 -Force }

    # FFmpeg command: high quality 1080p H.264 mp4
    $ffmpegArgs = @(
        "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", "`"$concatListPath`"",
        "-vf", "`"fps=30,format=yuv420p`"",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        "-movflags", "+faststart",
        "`"$outputMp4`""
    ) -join " "

    $cmd = "ffmpeg $ffmpegArgs"
    Invoke-Expression $cmd

    if (Test-Path $outputMp4) {
        $item = Get-Item $outputMp4
        $durSec = 0
        try {
            $probeCmd = "ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 `"$outputMp4`""
            $durSec = [Math]::Round([double](Invoke-Expression $probeCmd), 1)
        } catch {}
        Write-Host "[SUCCESS] Generated $outputMp4 ($([Math]::Round($item.Length / 1MB, 2)) MB, Duration: $durSec s)" -ForegroundColor Green
    } else {
        throw "Failed to generate $outputMp4!"
    }

    Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
}

# =========================================================================
# 1. EXCEL GUIDE SCENES
# =========================================================================
$excelScenes = @(
    @{
        Badge = "개요 및 핵심 소개"
        Title = "func_SMILES2IMG 핵심 기능 3분 마스터"
        Subtitle = "이미 설치된 Excel에서 네이티브 인셀 분자 구조식 & Ketcher 에디터 100% 활용하기"
        Image = (Join-Path $assetsDir "excel-showcase.png")
        Formula = "=SMILES2IMG(A2)"
        Points = @(
            "네이티브 인셀 이미지::셀 안에 쏙 들어가는 고해상도 벡터급 2D 분자 이미지 렌더링",
            "100% 로컬 오프라인 구동::외부 클라우드나 웹 서버 전송 없이 안전하게 기업 화학식 보호",
            "행/열 크기 자동 연동::셀 너비나 행 높이를 조절하면 비율을 유지하며 완벽히 스케일링",
            "Excel 전 버전 완벽 호환::Excel 2024/M365는 인셀(=SMILES2IMG), 구버전은 셀 연동(=FLOAT)"
        )
        SubtitleBar = "💡 본 영상은 func_SMILES2IMG이 설치된 환경에서 핵심 기능을 3분 내에 마스터하는 실전 가이드입니다."
        Duration = 13
    },
    @{
        Badge = "1단계: 기본 인셀 수식"
        Title = "기본 인셀 수식 =SMILES2IMG(A2) 입력하기"
        Subtitle = "SMILES 문자열 참조 또는 수식 직접 입력으로 0.1초 만에 분자 생성"
        Image = (Join-Path $assetsDir "excel-showcase.png")
        Formula = "=SMILES2IMG(A2)  또는  =SMILES2IMG(""CC(=O)Oc1ccccc1C(=O)O"")"
        Points = @(
            "수식 타이핑 IntelliSense::수식 입력 시 실시간 인자 설명과 버전 안내 툴팁 지원",
            "셀 참조 방식::A열에 SMILES 문자열을 넣고 B2에 =SMILES2IMG(A2)를 입력",
            "행 필터 & 정렬 완벽 유지::Excel 네이티브 인셀 이미지이므로 필터링/정렬 시에도 셀에 고정",
            "다중 행 일괄 채우기::수식 핸들을 더블 클릭하여 수백 개 화합물을 순식간에 일괄 생성"
        )
        SubtitleBar = "📌 1단계: 빈 셀에 =SMILES2IMG(A2)를 입력하면 셀 안에 고해상도 2D 분자 이미지가 즉시 생성됩니다."
        Duration = 15
    },
    @{
        Badge = "2단계: 전문 스타일 옵션"
        Title = "특허청(B/W), 학술지(Stereo), 다크모드 & 회전/반전"
        Subtitle = "3번째 style 인자와 4번째 transform 인자로 화학 전문가 규격 완벽 대응"
        Image = (Join-Path $assetsDir "excel-matrix-showcase.png")
        Formula = "=SMILES2IMG(A2, ""white"", ""bw num stereo h all-c"", ""flip rot90"")"
        Points = @(
            "특허청 출원 도면 (""bw"")::KIPO/USPTO 선화(Line Art) 100% 흑백 단색 출력",
            "NMR 귀속 번호 (""num"")::1H / 13C NMR 피크 귀속용 원자 일련번호(1, 2, 3...) 첨자 표기",
            "입체화학 라벨 (""stereo"")::비대칭 중심 카이랄 R/S 및 기하 E/Z 절대배열 표기",
            "궁극의 풀옵션 조합::공백이나 파이프(|)로 모든 옵션 결합 및 flip/rot90 회전·반전"
        )
        SubtitleBar = "🎨 2단계: 특허용(B/W), 저널용(Stereo), NMR 피크 귀속(Number), 다크모드 등 8가지 전문 스타일을 지원합니다."
        Duration = 16
    },
    @{
        Badge = "3단계: Ketcher 2D 에디터"
        Title = "내장 EPAM Ketcher 에디터로 직접 그리고 수정하기"
        Subtitle = "마우스 우클릭 또는 더블 클릭으로 100% 오프라인 웹 에디터 즉시 연동"
        Image = (Join-Path $assetsDir "ketcher-guide.png")
        Formula = "단축키: Ctrl + Shift + K  |  적용: Ctrl + Enter"
        Points = @(
            "원클릭 에디터 팝업::빈 셀을 우클릭하고 [분자 구조식 삽입 (셀 수식)] 선택",
            "직관적인 2D 드로잉::벤젠 고리, 결합선, 헤테로 원자(N, O, S)를 마우스로 자유롭게 작도",
            "더블 클릭 제자리 재편집::이미 삽입된 분자 셀을 더블 클릭하면 기존 구조식 그대로 재오픈",
            "기업 기밀 보안 100%::로컬 WebView2 가상 호스트 기술로 외부 통신 없이 완전 오프라인 구동"
        )
        SubtitleBar = "🧬 3단계: 셀을 우클릭하거나 더블 클릭하면 오프라인 Ketcher가 열려 마우스로 직접 그리고 수정할 수 있습니다."
        Duration = 16
    },
    @{
        Badge = "4단계: 화학 반응식 (KET)"
        Title = "=KET2IMG — 반응 화살표 & 텍스트 주석 보존"
        Subtitle = "반응물, 생성물, 촉매, 화살표가 포함된 복합 캔버스를 원본 그대로 렌더링"
        Image = (Join-Path $assetsDir "workflow-ketcher.png")
        Formula = "=KET2IMG(A2)  또는  =KET2IMG.FLOAT(A2)"
        Points = @(
            "네이티브 KET 포맷 지원::Ketcher 고유 캔버스 데이터(KET JSON)를 셀에 안전하게 보존",
            "반응식 완벽 표현::A + B -> C 형태의 다단계 유기 합성 경로와 시약 화살표 삽입",
            "더블 클릭 동기화::수식 셀을 더블 클릭하면 원본 반응식 그대로 Ketcher가 열려 수정 후 동기화",
            "SMILES 자동 감지::=SMILES2IMG 함수 또한 KET 포맷을 자동 감지하여 완벽히 렌더링"
        )
        SubtitleBar = "⚡ 4단계: 반응 화살표와 촉매 주석이 포함된 화학 반응식은 =KET2IMG 수식으로 무손실 보존됩니다."
        Duration = 15
    },
    @{
        Badge = "5단계: 퀵 인서트 & 물성 복사"
        Title = "클립보드 즉시 삽입 & 화학식·분자량(MW) 복사"
        Subtitle = "리본 메뉴 및 우클릭 컨텍스트 메뉴로 화학 연구 생산성 극대화"
        Image = (Join-Path $assetsDir "ribbon-guide.png")
        Formula = "상단 리본 [SMILES] 탭  |  우클릭 [분자 정보 복사]"
        Points = @(
            "[Quick Insert]::웹이나 논문에서 복사한 SMILES를 에디터 열지 않고 즉시 그림으로 삽입",
            "[Copy SMILES]::삽입된 분자 그림에서 표준 SMILES 문자열을 클립보드로 역추출",
            "[Copy Molecule Info]::화학식(Formula), 분자량(MW), 정밀 질량(Exact Mass) 원클릭 계산",
            "구버전 엑셀 호환(=FLOAT)::Excel 2016-2021에서도 셀 크기에 연동되는 플로팅 그림 지원"
        )
        SubtitleBar = "📋 5단계: 복사된 SMILES는 0.1초 만에 즉시 삽입되며, 화학식과 분자량(MW)도 원클릭으로 복사됩니다."
        Duration = 15
    },
    @{
        Badge = "핵심 요약 & 체크리스트"
        Title = "func_SMILES2IMG 실전 활용 체크리스트"
        Subtitle = "화학 연구와 데이터 분석의 속도를 혁신하는 차세대 오피스 스프레드시트"
        Image = (Join-Path $assetsDir "excel-showcase.png")
        Formula = "상세 가이드: docs/EXCEL_GUIDE.md"
        Points = @(
            "✔ =SMILES2IMG(A2)::가장 빠른 인셀 2D 화학 구조식 생성",
            "✔ bw / num / stereo::논문과 특허청 규격에 맞춘 원클릭 스타일링",
            "✔ 더블 클릭 수정::셀이나 그림을 언제든 더블 클릭해 Ketcher로 재작도",
            "✔ =KET2IMG(A2)::반응 화살표가 포함된 복합 반응식 무손실 보존"
        )
        SubtitleBar = "🎉 엑셀에서 모든 화학 구조식을 자유롭게 활용해보세요! 상세 가이드는 docs/EXCEL_GUIDE.md를 참고하세요."
        Duration = 13
    }
)

# =========================================================================
# 2. WORD GUIDE SCENES
# =========================================================================
$wordScenes = @(
    @{
        Badge = "개요 및 핵심 소개"
        Title = "func_SMILES2IMG 핵심 기능 3분 마스터"
        Subtitle = "학술 논문, 연구 보고서, 실험 노트를 위한 화학 구조식 자동 삽입 & Ketcher 연동"
        Image = (Join-Path $assetsDir "word-showcase.png")
        Formula = "Word 상단 메뉴 [SMILES] 탭 & 마우스 우클릭 메뉴"
        Points = @(
            "고해상도 인라인 그림::문서 본문과 표(Table) 셀 안에 깔끔하게 정렬되는 인라인 벡터급 PNG",
            "더블 클릭 제자리 수정::삽입된 그림을 더블 클릭하면 표 크기를 유지한 채 제자리 수정",
            "0.1초 즉시 삽입::본문에서 드래그한 SMILES 텍스트를 즉시 구조식 그림으로 치환",
            "100% 무손실 KET 보존::그림 속성(AlternativeText)에 전체 캔버스 데이터가 안전하게 저장"
        )
        SubtitleBar = "💡 본 영상은 Word용 SMILES 추가 기능이 설치된 상태에서 핵심 활용법을 빠르게 마스터하는 가이드입니다."
        Duration = 13
    },
    @{
        Badge = "1단계: 구조식 삽입"
        Title = "상단 리본 [SMILES] 탭에서 구조식 삽입하기"
        Subtitle = "100% 오프라인 Ketcher 에디터를 열고 마우스로 그려 즉시 문서에 배치"
        Image = (Join-Path $assetsDir "ribbon-guide.png")
        Formula = "[SMILES] 탭 -> [Insert Structure] -> [적용 (Ctrl+Enter)]"
        Points = @(
            "리본 메뉴 클릭::Word 상단 리본 메뉴 끝에 추가된 [SMILES] 탭 선택",
            "[Insert Structure] 클릭::EPAM Ketcher 분자 에디터 창이 팝업",
            "자유로운 2D 작도::단일 분자, 복합 반응식, 입체 결합선을 마우스로 손쉽게 드로잉",
            "배경 옵션 선택::상단 바에서 [배경: 투명 / 흰색]을 선택하고 Ctrl+Enter로 문서 삽입"
        )
        SubtitleBar = "📌 1단계: 상단 리본의 [SMILES] 탭에서 [Insert Structure]를 누르면 내장 Ketcher 에디터가 열립니다."
        Duration = 15
    },
    @{
        Badge = "2단계: 퀵 인서트"
        Title = "선택 텍스트 & 클립보드 즉시 삽입 (Quick Insert)"
        Subtitle = "드래그한 SMILES 텍스트를 에디터 실행 없이 0.1초 만에 그림으로 치환"
        Image = (Join-Path $assetsDir "word-showcase.png")
        Formula = "단축 동작: 텍스트 드래그 선택 -> [Quick Insert] 클릭"
        Points = @(
            "본문 텍스트 드래그::문서에 입력된 SMILES 문자열(예: C1=CC=C... )을 마우스로 드래그",
            "[Quick Insert] 클릭::리본 메뉴나 우클릭 메뉴의 [Quick Insert]를 누르면 즉시 변환",
            "클립보드 다이렉트 변환::웹 브라우저나 외부 논문에서 복사한 SMILES도 즉시 삽입",
            "보고서 작성 속도 극대화::수십 개의 화합물 코드를 순식간에 시각적 그림으로 변환"
        )
        SubtitleBar = "⚡ 2단계: 문서 내 SMILES 텍스트를 드래그하거나 클립보드 복사 후 [Quick Insert]로 즉시 그림으로 변환합니다."
        Duration = 15
    },
    @{
        Badge = "3단계: 더블 클릭 수정"
        Title = "그림 더블 클릭으로 레이아웃 유지 제자리 수정"
        Subtitle = "표 너비와 줄 간격을 흐트러뜨리지 않고 구조식을 실시간으로 업데이트"
        Image = (Join-Path $assetsDir "ketcher-guide.png")
        Formula = "동작: 그림 마우스 더블 클릭 -> Ketcher 수정 -> [적용]"
        Points = @(
            "원클릭 에디터 재호출::문서 내 삽입된 분자 그림을 마우스로 더블 클릭",
            "기존 구조 자동 복원::그림 안에 내장된 KET/SMILES 데이터를 읽어 Ketcher 캔버스에 로드",
            "치환기 및 결합선 수정::작용기를 변경하거나 원자를 수정한 뒤 [적용] 클릭",
            "제자리 완벽 치환::기존 그림의 너비와 높이, 표 서식을 유지하며 수정된 이미지로 교체"
        )
        SubtitleBar = "🔄 3단계: 삽입된 그림을 더블 클릭하면 언제든 Ketcher 에디터로 불러와 크기와 위치를 유지한 채 수정됩니다."
        Duration = 16
    },
    @{
        Badge = "4단계: 무손실 반응식"
        Title = "다단계 화학 반응식 & 무손실 메타데이터 보존"
        Subtitle = "반응 화살표, 텍스트 주석, 복수 분자가 그림 속성에 영구 보존"
        Image = (Join-Path $assetsDir "workflow-ketcher.png")
        Formula = "메타데이터 저장: AlternativeText / Title = KET:{...}"
        Points = @(
            "복합 반응식 지원::반응물 + 시약 -> 생성물 형태의 복수 분자 반응 경로 삽입",
            "Word 그림 속성 암호화::대체 텍스트(AlternativeText)에 전체 KET 데이터가 100% 보존",
            "파일 공유 시 무결성 유지::.docx 파일을 저장하고 타인에게 전달해도 메타데이터 완전 유지",
            "고해상도 PDF 변환::문서를 인쇄하거나 PDF로 변환해도 벡터급의 선명한 결합선 유지"
        )
        SubtitleBar = "🛡️ 4단계: 반응 화살표와 텍스트 주석이 그림 속성에 100% 무손실 보존되어 문서 저장 후에도 완벽히 복원됩니다."
        Duration = 15
    },
    @{
        Badge = "5단계: 물성 정보 복사"
        Title = "우클릭 메뉴로 SMILES 추출 & 화학 물성 복사"
        Subtitle = "화학식(Formula), 분자량(MW), 정밀 질량(Exact Mass) 원클릭 계산"
        Image = (Join-Path $assetsDir "word-showcase.png")
        Formula = "우클릭 메뉴: [Copy SMILES]  |  [Copy Molecule Info]"
        Points = @(
            "[Copy SMILES]::그림에서 표준 SMILES 문자열을 추출하여 클립보드로 복사",
            "[Copy Molecule Info]::화학식(Formula)과 분자량(MW, g/mol)을 계산해 대화상자에 표시",
            "클립보드 요약 텍스트 자동 복사::논문 작성 시 분자량 표에 바로 붙여넣기 가능",
            "캔버스 데이터 백업::[Copy Canvas Data (KET)]로 원본 JSON 문자열을 그대로 복사"
        )
        SubtitleBar = "📊 5단계: 그림 우클릭 메뉴로 SMILES를 다시 추출하거나 화학식·분자량(MW)을 즉시 계산해 복사합니다."
        Duration = 14
    },
    @{
        Badge = "핵심 요약 & 체크리스트"
        Title = "func_SMILES2IMG 실전 활용 체크리스트"
        Subtitle = "화학 논문 및 실험 보고서 작성의 생산성을 극대화하는 솔루션"
        Image = (Join-Path $assetsDir "word-showcase.png")
        Formula = "상세 가이드: docs/WORD_GUIDE.md"
        Points = @(
            "✔ 상단 [SMILES] 리본::Ketcher 에디터로 고해상도 인라인 분자 그림 삽입",
            "✔ [Quick Insert]::본문 드래그 텍스트 0.1초 즉시 치환",
            "✔ 그림 더블 클릭::문서 표 서식을 유지한 채 제자리 구조식 수정",
            "✔ 무손실 KET 보존::반응 화살표와 촉매 주석까지 완벽 복원"
        )
        SubtitleBar = "🎉 워드에서 전문적인 화학 논문과 보고서를 손쉽게 작성해보세요! 상세 가이드는 docs/WORD_GUIDE.md를 참고하세요."
        Duration = 13
    }
)

# =========================================================================
# 3. POWERPOINT GUIDE SCENES
# =========================================================================
$pptScenes = @(
    @{
        Badge = "개요 및 핵심 소개"
        Title = "func_SMILES2IMG 핵심 기능 3분 마스터"
        Subtitle = "연구 발표, 학회 세미나, 강의 슬라이드를 위한 고해상도 화학 구조식 & 반응식 프레젠테이션"
        Image = (Join-Path $assetsDir "powerpoint-showcase.png")
        Formula = "PowerPoint 상단 메뉴 [SMILES] 탭 & 슬라이드 우클릭 메뉴"
        Points = @(
            "슬라이드 원클릭 삽입::현재 활성 슬라이드 중앙에 선명한 고품질 벡터급 도형으로 삽입",
            "투명 vs 흰색 배경 선택::다크 테마 슬라이드 발표에 완벽 대응하는 투명 캔버스 지원",
            "더블 클릭 제자리 수정::슬라이드 레이아웃과 크기를 유지하며 Ketcher로 원클릭 수정",
            "다단계 반응식 보존::개체 태그(Tags['KET'])에 복잡한 반응 경로와 화살표 완전 보존"
        )
        SubtitleBar = "💡 본 영상은 PowerPoint용 SMILES 추가 기능이 설치된 상태에서 핵심 활용법을 빠르게 마스터하는 가이드입니다."
        Duration = 13
    },
    @{
        Badge = "1단계: 슬라이드 삽입"
        Title = "상단 리본 [SMILES] 탭에서 슬라이드에 구조식 배치"
        Subtitle = "현재 슬라이드 중앙에 고화질 분자 도형을 손쉽게 자동 배치"
        Image = (Join-Path $assetsDir "ribbon-guide.png")
        Formula = "[SMILES] 탭 -> [Insert Structure] -> [적용 (Ctrl+Enter)]"
        Points = @(
            "상단 리본 탭::PowerPoint 상단 리본 메뉴 끝에 추가된 [SMILES] 탭 선택",
            "[Insert Structure] 클릭::100% 오프라인 EPAM Ketcher 분자 에디터 창 팝업",
            "슬라이드 맞춤 작도::분자 구조나 반응식을 그린 뒤 상단 바에서 배경 스타일 선택",
            "자동 중앙 배치::[적용] 클릭 시 현재 슬라이드의 한가운데에 최적 비율로 즉시 삽입"
        )
        SubtitleBar = "📌 1단계: [Insert Structure] 버튼으로 에디터를 열고 구조식을 그린 뒤 적용하면 슬라이드에 즉시 배치됩니다."
        Duration = 15
    },
    @{
        Badge = "2단계: 다크 테마 배경"
        Title = "투명 vs. 흰색 배경 선택 & 다크 테마 슬라이드"
        Subtitle = "슬라이드 템플릿과 빔프로젝터 환경에 맞춰 최적의 캔버스 스타일 선택"
        Image = (Join-Path $assetsDir "powerpoint-showcase.png")
        Formula = "Ketcher 상단 바: [Background: Transparent]  vs  [Background: White]"
        Points = @(
            "투명 캔버스 (Transparent)::슬라이드의 어두운 배경색, 그라디언트 테마가 그대로 투과",
            "흰색 캔버스 (White)::깔끔한 화이트 카드 형태의 분자 레이아웃 (인쇄 및 라이트 테마)",
            "시인성 최적화::황(S) 원자는 눈부심 없는 황갈색(#A88013)으로 자동 보정",
            "고대비 프레젠테이션::빔프로젝터나 대형 모니터 발표 시 결합선이 매우 선명하게 표시"
        )
        SubtitleBar = "🎨 2단계: 다크 테마 슬라이드 발표를 위한 [투명 배경]과 출력용 [흰색 배경]을 원클릭으로 선택할 수 있습니다."
        Duration = 15
    },
    @{
        Badge = "3단계: 슬라이드 퀵 인서트"
        Title = "슬라이드 텍스트 상자 & 클립보드 즉시 삽입"
        Subtitle = "텍스트 상자의 SMILES를 에디터 실행 없이 0.1초 만에 그림으로 치환"
        Image = (Join-Path $assetsDir "ketcher-guide.png")
        Formula = "텍스트 상자 입력 -> [Quick Insert] 클릭"
        Points = @(
            "텍스트 상자 활용::슬라이드 안의 텍스트 상자에 SMILES 문자열을 입력",
            "[Quick Insert] 원클릭::리본 탭의 [Quick Insert]를 누르면 0.1초 만에 분자 도형으로 치환",
            "클립보드 다이렉트 변환::웹 브라우저나 논문에서 복사한 SMILES도 즉시 슬라이드 삽입",
            "발표 자료 준비 시간 단축::수십 개의 화합물 슬라이드를 순식간에 제작 가능"
        )
        SubtitleBar = "⚡ 3단계: 슬라이드 텍스트 상자나 클립보드의 SMILES/KET를 에디터 실행 없이 [Quick Insert]로 즉시 삽입합니다."
        Duration = 15
    },
    @{
        Badge = "4단계: 더블 클릭 수정"
        Title = "슬라이드 도형 더블 클릭으로 레이아웃 유지 수정"
        Subtitle = "슬라이드 위치, 크기, 종횡비를 100% 유지하며 Ketcher로 제자리 재편집"
        Image = (Join-Path $assetsDir "powerpoint-showcase.png")
        Formula = "동작: 슬라이드 분자 도형 더블 클릭 -> Ketcher 수정 -> [적용]"
        Points = @(
            "원클릭 에디터 호출::슬라이드 위 분자 그림을 마우스로 더블 클릭",
            "파워포인트 기본 서식창 차단::불필요한 그림 서식 창 대신 전용 Ketcher 에디터 즉시 오픈",
            "구조식 실시간 수정::원하는 결합선이나 치환기를 수정한 뒤 [적용] 클릭",
            "좌표 및 비율 유지::슬라이드 위 기존 크기, 위치, 카드 정렬을 그대로 유지하며 교체"
        )
        SubtitleBar = "🔄 4단계: 슬라이드 위 분자 그림을 더블 클릭하면 슬라이드 레이아웃과 크기를 유지한 채 즉시 수정됩니다."
        Duration = 16
    },
    @{
        Badge = "5단계: 반응식 & 물성 정보"
        Title = "복합 반응식 경로 카드 & Tags['KET'] 메타데이터"
        Subtitle = "다단계 유기 합성 경로와 시약 화살표 보존 및 Q&A 분자량 즉시 확인"
        Image = (Join-Path $assetsDir "workflow-ketcher.png")
        Formula = "우클릭 메뉴: [Copy SMILES]  |  [Copy Molecule Info]"
        Points = @(
            "다단계 반응 경로 슬라이드::반응물, 촉매, 부산물, 화살표를 한 화면에 직관적으로 배치",
            "PowerPoint Tags 영구 보존::도형 속성(Tags['KET'])에 원본 데이터가 안전하게 저장",
            "발표 중 Q&A 분자량 확인::도형 우클릭 -> [Copy Molecule Info]로 분자량/화학식 즉시 확인",
            "슬라이드 쇼 고화질 유지::4K 대형 스크린에서도 깨짐 없는 완벽한 고해상도 벡터 품질"
        )
        SubtitleBar = "🧪 5단계: 복잡한 유기 반응 경로와 시약 화살표까지 파워포인트 도형 메타데이터에 온전히 보존됩니다."
        Duration = 15
    },
    @{
        Badge = "핵심 요약 & 체크리스트"
        Title = "func_SMILES2IMG 실전 활용 체크리스트"
        Subtitle = "화학 학회, 연구 발표, 세미나를 위한 압도적 비주얼 프레젠테이션"
        Image = (Join-Path $assetsDir "powerpoint-showcase.png")
        Formula = "상세 가이드: docs/POWERPOINT_GUIDE.md"
        Points = @(
            "✔ 상단 [SMILES] 리본::슬라이드 중앙에 고화질 2D 화학 구조식 즉시 배치",
            "✔ [투명 배경] 토글::다크 테마 슬라이드 발표에 최적화된 높은 대비 제공",
            "✔ 도형 더블 클릭::슬라이드 레이아웃을 유지한 채 제자리 구조식 수정",
            "✔ 무손실 KET 보존::복잡한 유기 합성 메커니즘과 반응 경로 완벽 복원"
        )
        SubtitleBar = "🎉 완성도 높은 화학 발표 슬라이드를 손쉽게 제작해보세요! 상세 가이드는 docs/POWERPOINT_GUIDE.md를 참고하세요."
        Duration = 13
    }
)

# Execution
if ($TargetApp -eq "All" -or $TargetApp -eq "Excel") {
    Build-VideoFromScenes -VideoName "excel-guide-video" -Scenes $excelScenes -AppTitle "func_SMILES2IMG" -AppColor "#107C41" -AppIcon "📗"
}

if ($TargetApp -eq "All" -or $TargetApp -eq "Word") {
    Build-VideoFromScenes -VideoName "word-guide-video" -Scenes $wordScenes -AppTitle "func_SMILES2IMG" -AppColor "#185ABD" -AppIcon "📘"
}

if ($TargetApp -eq "All" -or $TargetApp -eq "PowerPoint") {
    Build-VideoFromScenes -VideoName "powerpoint-guide-video" -Scenes $pptScenes -AppTitle "func_SMILES2IMG" -AppColor "#C43E1C" -AppIcon "📙"
}

Write-Host "`n[ALL GUIDE VIDEOS GENERATED SUCCESSFULLY IN $videosDir]" -ForegroundColor Green
