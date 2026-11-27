<#
.SYNOPSIS
    Captures genuine live-driven screenshots (.png), usage videos (.webm + .mp4),
    bilingual WebVTT captions (.en.vtt / .ko.vtt), and hero GIF directly from running
    Microsoft Excel, Word, PowerPoint, EPAM Ketcher, Mol* 3D Viewer, MoleculeInfoDialog,
    and AboutDialog on Windows.
#>
param(
    [ValidateSet('All', 'Excel', 'Word', 'PowerPoint', 'Common', 'UseCases', 'Finalize')]
    [string]$Section = 'All'
)

$ErrorActionPreference = 'Stop'
$CaptureLogicalWidth  = 1230
$CaptureLogicalHeight = 820
Add-Type -AssemblyName System.Drawing, System.Windows.Forms, UIAutomationClient, UIAutomationTypes, WindowsBase

$repoRoot = Split-Path -Parent $PSScriptRoot
$docsDir = Join-Path $repoRoot 'docs'
$publicDir = Join-Path $docsDir 'public'
$mediaRoot = Join-Path $publicDir 'media'
$assetsDir = Join-Path $repoRoot 'assets'
$assetsVideosDir = Join-Path $assetsDir 'videos'
$manifestPublicPath = Join-Path $mediaRoot 'manifest.json'
$manifestThemePath = Join-Path $docsDir '.vitepress\theme\mediaManifest.json'

New-Item -ItemType Directory -Path $mediaRoot, $assetsDir, $assetsVideosDir -Force | Out-Null

# Load LiveCaptureEngine
$engineCs = Get-Content (Join-Path $PSScriptRoot 'LiveCaptureEngine.cs') -Raw
Add-Type -TypeDefinition $engineCs -ReferencedAssemblies @(
    'System.Drawing',
    'System.Windows.Forms',
    'UIAutomationClient',
    'UIAutomationTypes',
    'WindowsBase'
) -ErrorAction Stop

[LiveCaptureEngine]::InitDpi()

# Ensure English Office language preference keys are set
$installLangScript = Join-Path $PSScriptRoot 'Install-OfficeEnglishPack.ps1'
if (Test-Path $installLangScript) {
    & $installLangScript -RegistryOnly | Out-Null
}

# Verify FFmpeg
$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + $env:Path
$ffmpeg = (Get-Command ffmpeg -ErrorAction Stop).Source
$ffprobe = (Get-Command ffprobe -ErrorAction Stop).Source

# Load Smiles2Img.Core + initialize NativeIndigo
$coreDll = (Resolve-Path (Join-Path $repoRoot 'dist\x64\Word\Smiles2Img.Core.dll')).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))

$distXll = Join-Path $repoRoot 'dist\x64\Smiles2Img-AddIn64-packed.xll'
if (-not (Test-Path $distXll)) {
    $distXll = Join-Path $env:APPDATA 'Microsoft\AddIns\Smiles2Img\Smiles2Img-AddIn64.xll'
}

$addinVersion = '1.9.1'
try {
    $propsXml = [xml](Get-Content (Join-Path $repoRoot 'Directory.Build.props') -Raw)
    if ($propsXml.Project.PropertyGroup.Version) {
        $addinVersion = [string]$propsXml.Project.PropertyGroup.Version
    }
} catch {}

# Clean any orphaned headless Office processes from prior runs
Get-Process EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -eq '' } |
    Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

# Load existing manifest: keep all entries on partial section runs, and preserve externally managed entries (e.g. install walkthrough) on full runs
$manifest = [ordered]@{}
if (Test-Path $manifestPublicPath) {
    try {
        $existingJson = [System.IO.File]::ReadAllText($manifestPublicPath, [System.Text.UTF8Encoding]::new($false))
        $existing = $existingJson | ConvertFrom-Json
        foreach ($prop in $existing.PSObject.Properties) {
            if ($Section -ne 'All' -or $prop.Name -eq 'common-install-walkthrough') {
                $manifest[$prop.Name] = $prop.Value
            }
        }
    } catch {
        Write-Warning "Failed to load existing manifest.json: $($_.Exception.Message)"
    }
}

$runnerPid = [uint32]$PID
$aspirinSmiles = 'CC(=O)Oc1ccccc1C(=O)O'
$caffeineSmiles = 'Cn1cnc2c1c(=O)n(C)c(=O)n2C'
$ibuprofenSmiles = 'CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O'
$nicotineSmiles = 'c1cncc(c1)[C@@H]2CCCN2C'
$paracetamolSmiles = 'CC(=O)Nc1ccc(O)cc1'
$sampleReactionKet = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet('O=C(O)c1ccccc1O.CC(=O)OC(=O)C>>CC(=O)Oc1ccccc1C(=O)O')
$sampleSingleKet = [Smiles2Img.Rendering.MoleculeRenderer]::TryConvertToKet($aspirinSmiles)

function Invoke-ComRetry([scriptblock]$Operation, [int]$TimeoutSec = 30) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
    while ($true) {
        try { return (& $Operation) }
        catch {
            $hex = '{0:X8}' -f $_.Exception.HResult
            if (($hex -notmatch '80010001|8001010A|800AC472' -and
                 $_.Exception.ToString() -notmatch '80010001|8001010A|800AC472|rejected') -or
                [DateTime]::UtcNow -ge $deadline) { throw }
            Start-Sleep -Milliseconds 350
        }
    }
}

function Save-ManifestFiles {
    $json = $manifest | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($manifestPublicPath, $json, [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText($manifestThemePath, $json, [System.Text.UTF8Encoding]::new($false))
}

function Set-StandardHostWindow([IntPtr]$hwnd, [string]$label) {
    $sz = [LiveCaptureEngine]::SetMainWindowSizeLogical($hwnd, 4, 4, $CaptureLogicalWidth, $CaptureLogicalHeight)
    $dpiScale = [LiveCaptureEngine]::GetWindowDpiScale($hwnd)
    Write-Host ("  [Window] {0} sized to {1}x{2} px ({3}x{4} DIPs @ {5}% DPI)" -f `
        $label, $sz.Width, $sz.Height, $CaptureLogicalWidth, $CaptureLogicalHeight, [int][Math]::Round($dpiScale * 100)) -ForegroundColor DarkCyan
    return $sz
}

function Scale-RefX([IntPtr]$hwnd, [int]$refX) {
    $sz = [LiveCaptureEngine]::GetWindowSize($hwnd)
    if ($sz.Width -le 0) { return $refX }
    return [int][Math]::Round($refX * $sz.Width / 2768.0)
}

function Scale-RefY([IntPtr]$hwnd, [int]$refY) {
    $sz = [LiveCaptureEngine]::GetWindowSize($hwnd)
    if ($sz.Width -le 0) { return $refY }
    return [int][Math]::Round($refY * $sz.Width / 2766.0)
}

function Format-VttTime([double]$sec) {
    $clamped = [Math]::Max([double]0.0, [double]$sec)
    $ts = [TimeSpan]::FromSeconds($clamped)
    return ('{0:D2}:{1:D2}:{2:D2}.{3:D3}' -f $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds)
}

function Select-ExcelCell($excelObj, $sheetObj, [string]$cellAddr, [int]$waitMs = 380) {
    [void](Invoke-ComRetry { $sheetObj.Activate() })
    try { if ($excelObj.ActiveWindow.Zoom -ne 85) { $excelObj.ActiveWindow.Zoom = 85 } } catch {}
    [void](Invoke-ComRetry { $sheetObj.Range($cellAddr).Select() })
    try { [void]$excelObj.Calculate() } catch {}
    Start-Sleep -Milliseconds $waitMs
}

function Write-VttFile([string]$path, [array]$cues, [string]$langField, [double]$maxDur, [string]$fallbackEn, [string]$fallbackKo) {
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('WEBVTT')
    $lines.Add('')

    $norm = New-Object 'System.Collections.Generic.List[hashtable]'
    foreach ($cue in $cues) {
        $text = if ($langField -eq 'ko') { $cue.ko } else { $cue.en }
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        $norm.Add(@{
            tStart = [double]$cue.tStart
            tEnd   = [double]$cue.tEnd
            text   = $text
        })
    }

    if ($norm.Count -gt 0) {
        $sorted = @($norm | Sort-Object { $_.tStart })
        $cursorEnd = 0.0
        for ($i = 0; $i -lt $sorted.Count; $i++) {
            $cItem = $sorted[$i]
            if (-not $cItem -or [string]::IsNullOrWhiteSpace($cItem.text)) { continue }
            $ts = [Math]::Max([double]$cursorEnd, [double]$cItem.tStart)
            if ($ts -ge ($maxDur - 0.15)) {
                # Compress remaining cue into tail if multiple cues arrived late
                $ts = [Math]::Max([double]0.0, [double]($maxDur - 0.5))
                if ($i -gt 0 -and $ts -lt $cursorEnd) { continue }
            }
            $nextStart = if ($i + 1 -lt $sorted.Count) { [Math]::Max([double]($ts + 0.35), [double]$sorted[$i + 1].tStart) } else { $maxDur }
            $te = [Math]::Min([double]$maxDur, [Math]::Min([double]$cItem.tEnd, [double]($nextStart - 0.03)))
            if ($i + 1 -eq $sorted.Count) { $te = [Math]::Min([double]$maxDur, [Math]::Max([double]($ts + 0.5), [double]$cItem.tEnd)) }
            if ($te -le $ts) { $te = [Math]::Min([double]$maxDur, [double]($ts + 0.4)) }
            if ($te -le $ts) { continue }
            $lines.Add("$(Format-VttTime $ts) --> $(Format-VttTime $te)")
            $lines.Add($cItem.text)
            $lines.Add('')
            $cursorEnd = $te + 0.03
        }
    }
    if ($lines.Count -le 2) {
        $fb = if ($langField -eq 'ko') { $fallbackKo } else { $fallbackEn }
        if ([string]::IsNullOrWhiteSpace($fb)) {
            $fb = if ($langField -eq 'ko') { 'func_SMILES2IMG 실시간 실행 화면' } else { 'func_SMILES2IMG Live Workflow' }
        }
        $endSec = [Math]::Max([double]0.8, [double]$maxDur)
        $lines.Add("00:00:00.000 --> $(Format-VttTime $endSec)")
        $lines.Add($fb)
        $lines.Add('')
    }
    [System.IO.File]::WriteAllLines($path, $lines, [System.Text.UTF8Encoding]::new($false))
}

function Start-FeatureClip {
    param(
        [string]$Id,
        [string]$HostApp,
        [IntPtr]$MainHwnd,
        [uint32]$OfficePid,
        [string]$TitleEn,
        [string]$TitleKo,
        [string]$CaptionEn,
        [string]$CaptionKo
    )

    $outDir = Join-Path $mediaRoot $Id
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    $framesDir = Join-Path ([System.IO.Path]::GetTempPath()) ("s2i_cap_" + $Id + "_" + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $framesDir -Force | Out-Null

    $state = [pscustomobject]@{
        Id          = $Id
        HostApp     = $HostApp
        MainHwnd    = $MainHwnd
        OfficePid   = $OfficePid
        TitleEn     = $TitleEn
        TitleKo     = $TitleKo
        CaptionEn   = $CaptionEn
        CaptionKo   = $CaptionKo
        OutDir      = $outDir
        FramesDir   = $framesDir
        Fps         = 12
        FrameIdx    = 0
        LastPng     = ''
        CurX        = (Scale-RefX $MainHwnd 1120)
        CurY        = (Scale-RefY $MainHwnd 630)
        Cues        = (New-Object 'System.Collections.Generic.List[hashtable]')
        ShotSaved   = $false
        ShotPath    = (Join-Path $framesDir 'shot_raw.png')
    }

    $state | Add-Member -MemberType ScriptMethod -Name Snap -Value {
        param([int]$count = 1, [int]$cx = -1, [int]$cy = -1, [double]$ripple = 0)
        for ($i = 0; $i -lt $count; $i++) {
            $this.FrameIdx++
            $p = Join-Path $this.FramesDir ('f_{0:D4}.png' -f $this.FrameIdx)
            if ($i -eq 0 -or $ripple -gt 0) {
                [LiveCaptureEngine]::CaptureCompositeFrame($this.MainHwnd, $this.OfficePid, $script:runnerPid, $cx, $cy, [float]$ripple, $p)
                $this.LastPng = $p
            } else {
                Copy-Item -LiteralPath $this.LastPng -Destination $p -Force
            }
        }
    }

    $state | Add-Member -MemberType ScriptMethod -Name MoveTo -Value {
        param([int]$tx, [int]$ty, [int]$steps = 6)
        $sx = $this.CurX
        $sy = $this.CurY
        for ($i = 1; $i -le $steps; $i++) {
            $t = $i / [double]$steps
            $ease = $t * $t * (3.0 - 2.0 * $t)
            $nx = [int]($sx + ($tx - $sx) * $ease)
            $ny = [int]($sy + ($ty - $sy) * $ease)
            $this.FrameIdx++
            $p = Join-Path $this.FramesDir ('f_{0:D4}.png' -f $this.FrameIdx)
            [LiveCaptureEngine]::CaptureCompositeFrame($this.MainHwnd, $this.OfficePid, $script:runnerPid, $nx, $ny, 0, $p)
            $this.LastPng = $p
        }
        $this.CurX = $tx
        $this.CurY = $ty
    }

    $state | Add-Member -MemberType ScriptMethod -Name ClickAt -Value {
        param([int]$tx, [int]$ty, [string]$cueEn = '', [string]$cueKo = '')
        $this.MoveTo($tx, $ty, 6)
        $tStart = $this.FrameIdx / [double]$this.Fps
        foreach ($r in @(0.25, 0.55, 0.85)) {
            $this.FrameIdx++
            $p = Join-Path $this.FramesDir ('f_{0:D4}.png' -f $this.FrameIdx)
            [LiveCaptureEngine]::CaptureCompositeFrame($this.MainHwnd, $this.OfficePid, $script:runnerPid, $tx, $ty, [float]$r, $p)
            $this.LastPng = $p
        }
        if ($cueEn -or $cueKo) {
            $this.Cues.Add(@{ tStart = $tStart; tEnd = ($tStart + 2.2); en = $cueEn; ko = $cueKo })
        }
    }

    $state | Add-Member -MemberType ScriptMethod -Name AddCue -Value {
        param([double]$durSec, [string]$cueEn, [string]$cueKo)
        $tStart = $this.FrameIdx / [double]$this.Fps
        $this.Cues.Add(@{ tStart = $tStart; tEnd = ($tStart + $durSec); en = $cueEn; ko = $cueKo })
    }

    $state | Add-Member -MemberType ScriptMethod -Name CaptureShot -Value {
        param([string[]]$extraCopyPaths = @())
        [LiveCaptureEngine]::CaptureCompositeFrame($this.MainHwnd, $this.OfficePid, $script:runnerPid, -1, -1, 0, $this.ShotPath)
        [LiveCaptureEngine]::VerifyImageDiverse($this.ShotPath, 800, 450)
        $this.ShotSaved = $true
        foreach ($ep in $extraCopyPaths) {
            $dir = Split-Path -Parent $ep
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Copy-Item -LiteralPath $this.ShotPath -Destination $ep -Force
        }
    }

    $state | Add-Member -MemberType ScriptMethod -Name RunWithModal -Value {
        param([string]$titleSub, [int]$holdMs, [scriptblock]$invokeModal, [string]$cueEn = '', [string]$cueKo = '', [string[]]$extraShotCopies = @())
        $tStart = $this.FrameIdx / [double]$this.Fps
        if ($cueEn -or $cueKo) {
            $this.Cues.Add(@{ tStart = $tStart; tEnd = ($tStart + 2.8); en = $cueEn; ko = $cueKo })
        }
        $prevIdx = $this.FrameIdx
        [LiveCaptureEngine]::StartModalSnapLoop($this.MainHwnd, $this.OfficePid, $script:runnerPid, $this.FramesDir, $this.FrameIdx, $titleSub, $holdMs)
        & $invokeModal
        $newIdx = [LiveCaptureEngine]::StopModalSnapLoop()
        $captured = $newIdx - $prevIdx
        $bestIdx = [LiveCaptureEngine]::LastActiveModalShotFrameIndex
        Write-Host "    [Modal] '$titleSub': +$captured frames (best=$bestIdx)" -ForegroundColor DarkCyan
        if ($captured -le 0 -and $extraShotCopies.Count -gt 0) {
            throw "Modal capture '$titleSub' recorded 0 frames; refusing to leave stale extraShotCopies."
        }
        if ($newIdx -gt $this.FrameIdx) {
            $this.FrameIdx = $newIdx
            if ($bestIdx -lt 1) { $bestIdx = [Math]::Max(1, $this.FrameIdx - 4) }
            $midFrame = Join-Path $this.FramesDir ('f_{0:D4}.png' -f $bestIdx)
            if (Test-Path $midFrame) {
                $this.LastPng = $midFrame
                if (-not $this.ShotSaved -or $extraShotCopies.Count -gt 0) {
                    Copy-Item -LiteralPath $midFrame -Destination $this.ShotPath -Force
                    $this.ShotSaved = $true
                    foreach ($ep in $extraShotCopies) {
                        $dir = Split-Path -Parent $ep
                        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                        Copy-Item -LiteralPath $midFrame -Destination $ep -Force
                    }
                }
            }
        }
    }

    return $state
}

function Complete-FeatureClip {
    param([pscustomobject]$Clip)

    if (-not $Clip.ShotSaved) {
        $Clip.CaptureShot(@())
    }
    [LiveCaptureEngine]::VerifyImageDiverse($Clip.ShotPath, 800, 450)

    if ($Clip.FrameIdx -lt 36) {
        $Clip.Snap((36 - $Clip.FrameIdx), -1, -1, 0)
    }

    $webmPath = Join-Path $Clip.OutDir 'clip.webm'
    $mp4Path  = Join-Path $Clip.OutDir 'clip.mp4'
    $vttEnPath = Join-Path $Clip.OutDir 'captions.en.vtt'
    $vttKoPath = Join-Path $Clip.OutDir 'captions.ko.vtt'

    $inputPattern = Join-Path $Clip.FramesDir 'f_%04d.png'
    & $ffmpeg -hide_banner -loglevel error -y -framerate $Clip.Fps -i $inputPattern `
        -vf "scale=1920:1280:flags=lanczos,format=yuv420p" `
        -c:v libx264 -preset medium -crf 23 -movflags +faststart $mp4Path
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg MP4 encoding failed for $($Clip.Id)" }

    & $ffmpeg -hide_banner -loglevel error -y -framerate $Clip.Fps -i $inputPattern `
        -vf "scale=1920:1280:flags=lanczos,format=yuv420p" `
        -c:v libvpx-vp9 -b:v 0 -crf 34 -row-mt 1 $webmPath
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg WebM encoding failed for $($Clip.Id)" }

    $shotWebpPath = Join-Path $Clip.OutDir 'shot.webp'
    & $ffmpeg -hide_banner -loglevel error -y -i $Clip.ShotPath `
        -vf "scale=1920:1280:flags=lanczos" `
        -c:v libwebp -quality 86 $shotWebpPath
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg WebP screenshot encoding failed for $($Clip.Id)" }

    $dur = 0.0
    try {
        $dur = [Math]::Round([double](& $ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $mp4Path), 2)
    } catch {}
    if ($dur -le 0.1) { $dur = [Math]::Round($Clip.FrameIdx / [double]$Clip.Fps, 2) }

    Write-VttFile -path $vttEnPath -cues $Clip.Cues -langField 'en' -maxDur $dur -fallbackEn $Clip.CaptionEn -fallbackKo $Clip.CaptionKo
    Write-VttFile -path $vttKoPath -cues $Clip.Cues -langField 'ko' -maxDur $dur -fallbackEn $Clip.CaptionEn -fallbackKo $Clip.CaptionKo

    $durRounded = [Math]::Round($dur, 1)

    $w = 1920
    $h = 1280

    $manifest[$Clip.Id] = [ordered]@{
        id         = $Clip.Id
        host       = $Clip.HostApp
        titleEn    = $Clip.TitleEn
        titleKo    = $Clip.TitleKo
        shot       = "/media/$($Clip.Id)/shot.webp"
        webm       = "/media/$($Clip.Id)/clip.webm"
        mp4        = "/media/$($Clip.Id)/clip.mp4"
        captionEn  = $Clip.CaptionEn
        captionKo  = $Clip.CaptionKo
        vttEn      = "/media/$($Clip.Id)/captions.en.vtt"
        vttKo      = "/media/$($Clip.Id)/captions.ko.vtt"
        version    = $addinVersion
        capturedAt = (Get-Date).ToString('yyyy-MM-dd')
        width      = $w
        height     = $h
        durationSec = $durRounded
    }
    Save-ManifestFiles

    Write-Host ("  [OK] {0,-34} | WebP {1}x{2} ({3}KB) | MP4 {4}KB | WebM {5}KB | {6}s" -f `
        $Clip.Id, $w, $h, [int]((Get-Item $shotWebpPath).Length/1KB), [int]((Get-Item $mp4Path).Length/1KB), [int]((Get-Item $webmPath).Length/1KB), $durRounded) -ForegroundColor Green

    Remove-Item -LiteralPath $Clip.FramesDir -Recurse -Force -ErrorAction SilentlyContinue
}

function Show-LiveKetcherOverHost {
    param([IntPtr]$OwnerHwnd, [string]$Context, [string]$InitialInput, [bool]$ToggleWhiteBg = $false, [int]$HoldMs = 2000)
    $form = New-Object Smiles2Img.KetcherForm($Context, $InitialInput)
    $form.AutoApplyAfterLoadForTest = $false
    $owner = New-Object Smiles2Img.Win32Window($OwnerHwnd)
    $t = New-Object System.Windows.Forms.Timer
    $t.Tag = $form
    $t.Interval = $HoldMs + 15500
    $t.Add_Tick({
        param($sender, $e)
        $sender.Stop()
        try { $sender.Tag.Close() } catch {}
    })
    $t.Start()
    [void]$form.ShowDialog($owner)
    $t.Dispose()
    $form.Dispose()
}

function Convert-DarkBgToTransparent([byte[]]$pngBytes) {
    $msIn = New-Object System.IO.MemoryStream(, $pngBytes)
    $src = [System.Drawing.Bitmap]::FromStream($msIn)
    $dst = New-Object System.Drawing.Bitmap($src.Width, $src.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $tl = $src.GetPixel(2, 2)
        $isDarkSolid = ($tl.R -le 18 -and $tl.G -le 18 -and $tl.B -le 18 -and $tl.A -ge 240)
        for ($y = 0; $y -lt $src.Height; $y++) {
            for ($x = 0; $x -lt $src.Width; $x++) {
                $px = $src.GetPixel($x, $y)
                if ($isDarkSolid -and $px.R -le 18 -and $px.G -le 18 -and $px.B -le 18) {
                    $dst.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, 0, 0, 0))
                } else {
                    $dst.SetPixel($x, $y, $px)
                }
            }
        }
        $msOut = New-Object System.IO.MemoryStream
        $dst.Save($msOut, [System.Drawing.Imaging.ImageFormat]::Png)
        $res = $msOut.ToArray()
        $msOut.Dispose()
        return ,$res
    } finally {
        $dst.Dispose()
        $src.Dispose()
        $msIn.Dispose()
    }
}

$script:cached1CrnPng = $null
function Get-Live1CrnTransparentPng([IntPtr]$OwnerHwnd, [string]$Context) {
    $diskCache = Join-Path ([System.IO.Path]::GetTempPath()) 's2i_cached_1crn_trans.png'
    if ($script:cached1CrnPng -and $script:cached1CrnPng.Length -gt 8000) {
        return ,$script:cached1CrnPng
    }
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $res = Show-LiveMolstarOverHost -OwnerHwnd $OwnerHwnd -Context $Context -PdbId '1CRN' -HoldMs (3000 + $attempt * 800) -AutoApply $true
        if ($res -and $res.ImageBytes -and $res.ImageBytes.Length -gt 8000) {
            $script:cached1CrnPng = Convert-DarkBgToTransparent $res.ImageBytes
            [System.IO.File]::WriteAllBytes($diskCache, $script:cached1CrnPng)
            return ,$script:cached1CrnPng
        }
        Start-Sleep -Milliseconds 600
    }
    if (Test-Path $diskCache) {
        $diskBytes = [System.IO.File]::ReadAllBytes($diskCache)
        if ($diskBytes.Length -gt 8000) {
            $script:cached1CrnPng = $diskBytes
            return ,$script:cached1CrnPng
        }
    }
    throw "Failed to capture valid 1CRN transparent PNG from MolstarForm after 3 live attempts."
}

function Show-LiveMolstarOverHost {
    param([IntPtr]$OwnerHwnd, [string]$Context, [string]$PdbId, [int]$HoldMs = 2200, [bool]$AutoApply = $true)
    $form = New-Object Smiles2Img.MolstarForm($Context, $PdbId, $null)
    $form.AutoApplyAfterLoadForTest = $AutoApply
    $owner = New-Object Smiles2Img.Win32Window($OwnerHwnd)
    $t = New-Object System.Windows.Forms.Timer
    $t.Tag = $form
    $t.Interval = $HoldMs + 12000
    $t.Add_Tick({
        param($sender, $e)
        $sender.Stop()
        try { $sender.Tag.Close() } catch {}
    })
    $t.Start()
    [void]$form.ShowDialog($owner)
    $t.Dispose()
    $bytes = $form.ResultImageBytes
    if (-not $bytes) { $bytes = $form.ResultUiViewportPngBytes }
    $glb = $form.ResultGlbBytes
    $form.Dispose()
    return @{ ImageBytes = $bytes; GlbBytes = $glb }
}

function Show-LiveMoleculeInfoOverHost {
    param([IntPtr]$OwnerHwnd, [string]$SmilesOrKet, [int]$HoldMs = 2000)
    $info = [Smiles2Img.Rendering.MoleculeRenderer]::TryGetMoleculeInfo($SmilesOrKet, $true)
    $dlg = New-Object Smiles2Img.Editor.MoleculeInfoDialog($info, 'Molecule Info & Chemical Properties')
    $owner = New-Object Smiles2Img.Win32Window($OwnerHwnd)
    $t = New-Object System.Windows.Forms.Timer
    $t.Tag = $dlg
    $t.Interval = $HoldMs + 1500
    $t.Add_Tick({
        param($sender, $e)
        $sender.Stop()
        try { $sender.Tag.Close() } catch {}
    })
    $t.Start()
    [void]$dlg.ShowDialog($owner)
    $t.Dispose()
    $dlg.Dispose()
}

function Show-LiveAboutOverHost {
    param([IntPtr]$OwnerHwnd, [int]$HoldMs = 2000)
    $dlg = New-Object Smiles2Img.Editor.AboutDialog('func_SMILES2IMG')
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $mr = New-Object LiveCaptureEngine+RECT
    if ($OwnerHwnd -ne [IntPtr]::Zero -and [LiveCaptureEngine]::GetWindowRect($OwnerHwnd, [ref]$mr) -and ($mr.R -gt $mr.L)) {
        $cx = [int]($mr.L + (($mr.R - $mr.L) - $dlg.Width) / 2)
        $cy = [int]($mr.T + (($mr.B - $mr.T) - $dlg.Height) / 2)
        $dlg.Location = New-Object System.Drawing.Point([Math]::Max(40, $cx), [Math]::Max(40, $cy))
    } else {
        $dlg.Location = New-Object System.Drawing.Point(400, 220)
    }
    $dlg.TopMost = $true
    $t = New-Object System.Windows.Forms.Timer
    $t.Tag = $dlg
    $t.Interval = $HoldMs + 3500
    $t.Add_Tick({
        param($sender, $e)
        $sender.Stop()
        try { $sender.Tag.Close() } catch {}
    })
    $dlg.Tag = $t
    $dlg.Add_Shown({
        param($sender, $e)
        $sender.Activate()
        $sender.BringToFront()
        if ($sender.Tag) { $sender.Tag.Start() }
    })
    [void]$dlg.ShowDialog()
    $t.Stop()
    $t.Dispose()
    $dlg.Dispose()
}

# =============================================================================
# 1. MICROSOFT EXCEL LIVE CAPTURE (11 Features + Ribbon + Showcase + Matrix)
# =============================================================================
if ($Section -eq 'All' -or $Section -eq 'Excel' -or $Section -eq 'UseCases') {
    Write-Host "`n=== [1/4] Capturing Live Microsoft Excel Features ===" -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    try {
        [void]$excel.RegisterXLL($distXll)
        $xHwnd = [IntPtr]$excel.Hwnd
        $xPid = [LiveCaptureEngine]::GetPidForHwnd($xHwnd)
        [void](Set-StandardHostWindow $xHwnd 'Excel')
        Start-Sleep -Milliseconds 700
        [void][LiveCaptureEngine]::SelectRibbonTab($xHwnd, 'SMILES')
        Start-Sleep -Milliseconds 600

        $ptInsertSmiles = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Insert / Edit (SMILES)', 'Insert / Edit'))
        $ptInsertKet    = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Insert / Edit (KET)'))
        $ptMolstar      = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('3D Structure (Mol*)'))
        $ptPaste        = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Paste Structure'))
        $ptCopySvg      = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Copy SVG'))
        $ptMolInfo      = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Molecule Info'))
        $ptBatch        = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Batch Style'))
        $ptSdf          = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Import SDF'))
        $ptDiag         = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('Diagnostics'))
        $ptAbout        = [LiveCaptureEngine]::FindControlRelativeCenter($xHwnd, @('About func_SMILES2IMG'))

        $wb = $excel.Workbooks.Add()
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        if ($Section -eq 'All' -or $Section -eq 'Excel') {
        # --- 1.1 excel-overview ---
        $ws = $wb.ActiveSheet
        $ws.Name = 'Compounds'
        $ws.Range('A1:F1').Value2 = @('Compound', 'SMILES', '2D Structure (=SMILES2IMG)', 'Formula', 'MW (g/mol)', 'LogP')
        $ws.Range('A1:F1').Font.Bold = $true
        $ws.Range('A1:F1').Interior.Color = 0xD1FAE5
        $ws.Columns.Item(1).ColumnWidth = 12
        $ws.Columns.Item(2).ColumnWidth = 36
        $ws.Columns.Item(3).ColumnWidth = 22
        $ws.Columns.Item(4).ColumnWidth = 11
        $ws.Columns.Item(5).ColumnWidth = 12
        $ws.Columns.Item(6).ColumnWidth = 10
        $rows = @(
            @('Aspirin', $aspirinSmiles),
            @('Caffeine', $caffeineSmiles),
            @('Ibuprofen', $ibuprofenSmiles),
            @('Paracetamol', $paracetamolSmiles)
        )
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $r = $i + 2
            $ws.Cells.Item($r, 1).Value2 = $rows[$i][0]
            $ws.Cells.Item($r, 2).Value2 = $rows[$i][1]
            $ws.Rows.Item($r).RowHeight = 52
        }

        $c = Start-FeatureClip -Id 'excel-overview' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Excel Overview: In-Cell Structure & Property Tables' `
            -TitleKo 'Excel 개요: 셀 안 화학 구조식 & 물성 표 자동 생성' `
            -CaptionEn 'Type =SMILES2IMG(B2), =SMILES2MW(B2), and =SMILES2LOGP(B2) in Microsoft Excel.' `
            -CaptionKo '엑셀 셀에 =SMILES2IMG(B2), =SMILES2MW(B2), =SMILES2LOGP(B2)를 입력해 구조식 표를 완성합니다.'
        $c.Snap(4, $ptInsertSmiles.X, $ptInsertSmiles.Y)
        $c.AddCue(1.6, 'Enter =SMILES2IMG(B2) to render crisp 2D structures inside Excel cells.', '=SMILES2IMG(B2) 수식으로 셀 안에 2D 화학 구조식을 렌더링합니다.')
        [void]$ws.Range('C2').Select()
        $ws.Range('C2').Formula2 = '=SMILES2IMG(B2)'
        $ws.Range('D2').Formula2 = '=SMILES2FORMULA(B2)'
        $ws.Range('E2').Formula2 = '=SMILES2MW(B2)'
        $ws.Range('F2').Formula2 = '=SMILES2LOGP(B2)'
        Start-Sleep -Seconds 2
        $c.Snap(5, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520))
        $c.AddCue(1.6, 'Fill formulas down to populate Aspirin, Caffeine, Ibuprofen & Paracetamol.', '수식을 아래로 채우면 모든 화합물의 구조식·분자량·LogP가 즉시 채워집니다.')
        for ($r = 3; $r -le 5; $r++) {
            $ws.Range("C$r").Formula2 = "=SMILES2IMG(B$r)"
            $ws.Range("D$r").Formula2 = "=SMILES2FORMULA(B$r)"
            $ws.Range("E$r").Formula2 = "=SMILES2MW(B$r)"
            $ws.Range("F$r").Formula2 = "=SMILES2LOGP(B$r)"
            Start-Sleep -Milliseconds 600
            $c.Snap(4, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd (520 + ($r - 2) * 85)))
        }
        Start-Sleep -Seconds 1
        Select-ExcelCell $excel $ws 'C2' 400
        $c.CaptureShot(@(
            (Join-Path $publicDir 'excel-showcase.png'),
            (Join-Path $assetsDir 'excel-showcase.png'),
            (Join-Path $publicDir 'excel-ribbon-guide.png'),
            (Join-Path $assetsDir 'excel-ribbon-guide.png'),
            (Join-Path $publicDir 'ribbon-guide.png'),
            (Join-Path $assetsDir 'ribbon-guide.png')
        ))
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.2 excel-formula-table ---
        $c = Start-FeatureClip -Id 'excel-formula-table' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: Build a SMILES Structure Table (=SMILES2IMG)' `
            -TitleKo '레시피: SMILES 목록으로 구조식 표 만들기 (=SMILES2IMG)' `
            -CaptionEn 'Scale cell row height and fill down =SMILES2IMG(B2) across rows.' `
            -CaptionKo '행 높이를 조절하고 =SMILES2IMG(B2) 수식을 아래로 채워 구조식 표를 만듭니다.'
        $ws.Range('C2:C5').ClearContents()
        Select-ExcelCell $excel $ws 'C2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520))
        $c.ClickAt((Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520), 'Select cell C2 and type =SMILES2IMG(B2).', 'C2 셀에 =SMILES2IMG(B2)를 입력합니다.')
        $ws.Range('C2').Formula2 = '=SMILES2IMG(B2)'
        Start-Sleep -Seconds 2
        $c.Snap(5, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520))
        $c.AddCue(1.6, 'Fill down C2:C5 for Aspirin, Caffeine, Ibuprofen, and Paracetamol.', 'C2:C5 범위로 수식을 복사해 모든 행을 렌더링합니다.')
        for ($r = 3; $r -le 5; $r++) {
            $ws.Range("C$r").Formula2 = "=SMILES2IMG(B$r)"
            Start-Sleep -Milliseconds 500
            $c.Snap(4, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd (520 + ($r - 2) * 85)))
        }
        Start-Sleep -Seconds 1
        Select-ExcelCell $excel $ws 'C3' 400
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.3 excel-name-cas ---
        $wsName = $wb.Worksheets.Add()
        $wsName.Name = 'NameAndCAS'
        $wsName.Range('A1:D1').Value2 = @('Input (Name / CAS / Solvent)', 'Structure (=NAME2IMG)', 'Canonical SMILES (=NAME2SMILES)', 'MW')
        $wsName.Range('A1:D1').Font.Bold = $true
        $wsName.Range('A1:D1').Interior.Color = 0xE0F2FE
        $wsName.Columns.Item(1).ColumnWidth = 22
        $wsName.Columns.Item(2).ColumnWidth = 22
        $wsName.Columns.Item(3).ColumnWidth = 30
        $wsName.Columns.Item(4).ColumnWidth = 12
        $nameInputs = @('Aspirin', '50-78-2', 'Caffeine', 'THF')
        for ($i = 0; $i -lt $nameInputs.Count; $i++) {
            $r = $i + 2
            $wsName.Cells.Item($r, 1).Value2 = $nameInputs[$i]
            $wsName.Rows.Item($r).RowHeight = 52
        }
        $c = Start-FeatureClip -Id 'excel-name-cas' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: Render Structures from Chemical Names & CAS Numbers' `
            -TitleKo '레시피: 화합물 이름·CAS 번호·실험실 약어로 구조식 그리기' `
            -CaptionEn 'Use =NAME2IMG(A2) and =NAME2SMILES(A2) for offline solvents, drugs, and CAS numbers.' `
            -CaptionKo '=NAME2IMG(A2)와 =NAME2SMILES(A2)로 시약 이름·CAS 번호를 즉시 변환합니다.'
        Select-ExcelCell $excel $wsName 'B2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 480), (Scale-RefY $xHwnd 520))
        $c.AddCue(1.8, 'Type =NAME2IMG(A2) for Aspirin, CAS 50-78-2, Caffeine, and THF.', 'Aspirin, 50-78-2, Caffeine, THF에 =NAME2IMG(A2)를 적용합니다.')
        for ($r = 2; $r -le 5; $r++) {
            $wsName.Range("B$r").Formula2 = "=NAME2IMG(A$r)"
            $wsName.Range("C$r").Formula2 = "=NAME2SMILES(A$r)"
            $wsName.Range("D$r").Formula2 = "=SMILES2MW(C$r)"
            Start-Sleep -Milliseconds 600
            $c.Snap(4, (Scale-RefX $xHwnd 480), (Scale-RefY $xHwnd (520 + ($r - 2) * 85)))
        }
        Start-Sleep -Seconds 1
        Select-ExcelCell $excel $wsName 'B2' 400
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.4 excel-style-matrix ---
        $wsStyle = $wb.Worksheets.Add()
        $wsStyle.Name = 'StyleMatrix'
        $wsStyle.Range('A1:F1').Value2 = @('Style Preset', 'Formula Syntax', 'Aspirin', 'Caffeine', 'Ibuprofen', 'Nicotine')
        $wsStyle.Range('A1:F1').Font.Bold = $true
        $wsStyle.Range('A1:F1').Interior.Color = 0xEDE9FE
        $wsStyle.Columns.Item(1).ColumnWidth = 20
        $wsStyle.Columns.Item(2).ColumnWidth = 38
        $wsStyle.Range('B2:B6').Font.Size = 9.5
        for ($col = 3; $col -le 6; $col++) { $wsStyle.Columns.Item($col).ColumnWidth = 15 }
        $styleMols = @($aspirinSmiles, $caffeineSmiles, $ibuprofenSmiles, $nicotineSmiles)
        $styleDefs = @(
            @{ Label = 'Standard Color'; Bg = ''; Style = '' },
            @{ Label = 'ACS Publication B&W'; Bg = ''; Style = 'bw' },
            @{ Label = 'Highlight Benzene'; Bg = ''; Style = 'hl:c1ccccc1' },
            @{ Label = 'Dark Mode (White Ink)'; Bg = '#1E293B'; Style = 'white' },
            @{ Label = 'NMR Atom # & Stereo'; Bg = ''; Style = 'num stereo' }
        )
        for ($rIdx = 0; $rIdx -lt $styleDefs.Count; $rIdx++) {
            $rowNum = $rIdx + 2
            $sd = $styleDefs[$rIdx]
            $wsStyle.Rows.Item($rowNum).RowHeight = 44
            $wsStyle.Cells.Item($rowNum, 1).Value2 = $sd.Label
            $wsStyle.Cells.Item($rowNum, 2).NumberFormat = '@'
            $wsStyle.Cells.Item($rowNum, 2).Value2 = "=SMILES2IMG(smiles, `"$($sd.Bg)`", `"$($sd.Style)`")"
        }
        $c = Start-FeatureClip -Id 'excel-style-matrix' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: ACS B&W, Highlighting, Dark Mode & Atom Numbering' `
            -TitleKo '레시피: 논문 흑백(ACS), 작용기 강조, 다크 모드, 원자 번호 스타일' `
            -CaptionEn 'Customize background [bg] and [style] tokens per cell.' `
            -CaptionKo '2번째 [배경색]과 3번째 [스타일] 인수로 셀별 스타일을 지정합니다.'
        Select-ExcelCell $excel $wsStyle 'C2' 300
        $c.Snap(4, (Scale-RefX $xHwnd 650), (Scale-RefY $xHwnd 500))
        $c.AddCue(1.8, 'Compare Standard Color, ACS B&W, Benzene Highlight, Dark Mode, and Atom #.', '표준 컬러, ACS 흑백, 벤젠 강조, 다크 모드, 원자 번호 5종 스타일을 렌더링합니다.')
        for ($rIdx = 0; $rIdx -lt $styleDefs.Count; $rIdx++) {
            $rowNum = $rIdx + 2
            $sd = $styleDefs[$rIdx]
            for ($mIdx = 0; $mIdx -lt $styleMols.Count; $mIdx++) {
                $colNum = $mIdx + 3
                $wsStyle.Cells.Item($rowNum, $colNum).Formula2 = "=SMILES2IMG(`"$($styleMols[$mIdx])`", `"$($sd.Bg)`", `"$($sd.Style)`")"
            }
            if ($sd.Bg -eq '#1E293B') {
                $wsStyle.Range("C${rowNum}:F${rowNum}").Interior.Color = 0x3B291E # Dark slate BGR
            }
            Start-Sleep -Milliseconds 650
            $c.Snap(4, (Scale-RefX $xHwnd 700), (Scale-RefY $xHwnd (460 + $rIdx * 70)))
        }
        Start-Sleep -Seconds 1
        $excel.ActiveWindow.ScrollRow = 1
        $excel.ActiveWindow.ScrollColumn = 1
        Select-ExcelCell $excel $wsStyle 'C3' 420
        $c.CaptureShot(@(
            (Join-Path $publicDir 'excel-matrix-showcase.png'),
            (Join-Path $assetsDir 'excel-matrix-showcase.png')
        ))
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.5 excel-reactions-ket ---
        $wsKet = $wb.Worksheets.Add()
        $wsKet.Name = 'ReactionsKET'
        $wsKet.Range('A1:C1').Value2 = @('Reaction / Aspirin KET JSON', 'Rendered Diagram (=KET2IMG)', 'Formula (=KET2FORMULA)')
        $wsKet.Range('A1:C1').Font.Bold = $true
        $wsKet.Columns.Item(1).ColumnWidth = 28
        $wsKet.Columns.Item(2).ColumnWidth = 44
        $wsKet.Columns.Item(3).ColumnWidth = 22
        $wsKet.Rows.Item(2).RowHeight = 105
        $wsKet.Rows.Item(3).RowHeight = 90
        $wsKet.Range('A2').Value2 = $sampleReactionKet
        $wsKet.Range('A3').Value2 = $sampleSingleKet
        Select-ExcelCell $excel $wsKet 'B2' 350
        $c = Start-FeatureClip -Id 'excel-reactions-ket' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: Multi-Step Reaction Schemes (=KET2IMG & Ketcher)' `
            -TitleKo '레시피: 화살표·촉매가 포함된 화학 반응식 (=KET2IMG & Ketcher)' `
            -CaptionEn 'Launch Ketcher from Excel to draw reaction arrows and persist lossless KET JSON.' `
            -CaptionKo '엑셀에서 Ketcher를 열어 반응 화살표를 그리고 =KET2IMG로 무손실 저장합니다.'
        $c.Snap(4, (Scale-RefX $xHwnd 550), (Scale-RefY $xHwnd 560))
        $c.ClickAt($ptInsertKet.X, $ptInsertKet.Y, 'Open Ketcher 2D editor to draw or inspect the KET reaction scheme.', 'Ketcher 2D 편집기를 열어 살리실산 아세틸화 반응식을 확인합니다.')
        $c.RunWithModal('Ketcher', 2000, {
            Show-LiveKetcherOverHost -OwnerHwnd $xHwnd -Context 'B2' -InitialInput $sampleReactionKet -HoldMs 2000
        }, 'Apply the salicylic acid + acetic anhydride reaction back to B2.', 'Apply를 눌러 B2 셀에 =KET2IMG(A2) 반응식을 삽입합니다.')
        $wsKet.Range('B2').Formula2 = '=KET2IMG(A2)'
        $wsKet.Range('C2').Value2 = 'Aspirin Synthesis (KET)'
        $wsKet.Range('B3').Formula2 = '=KET2IMG(A3)'
        $wsKet.Range('C3').Formula2 = '=KET2FORMULA(A3)'
        Start-Sleep -Seconds 2
        Select-ExcelCell $excel $wsKet 'B2' 400
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.6 excel-ribbon-tools ---
        Select-ExcelCell $excel $ws 'C2' 350
        $c = Start-FeatureClip -Id 'excel-ribbon-tools' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Excel Ribbon Tools: Copy SVG, Batch Style & About func_SMILES2IMG' `
            -TitleKo 'Excel 리본 도구: SVG 복사, Batch Style, About func_SMILES2IMG' `
            -CaptionEn 'Restyle sheet structures to ACS B&W, copy vector SVG, or inspect diagnostics and About.' `
            -CaptionKo '시트 전체를 ACS 흑백 스타일로 일괄 전환하고 SVG 복사와 정보 창을 확인합니다.'
        $c.Snap(4, $ptCopySvg.X, $ptCopySvg.Y)
        $c.ClickAt($ptCopySvg.X, $ptCopySvg.Y, 'Click Copy SVG to copy lossless vector graphics for Illustrator.', 'Copy SVG를 눌러 Illustrator용 벡터 그래픽을 복사합니다.')
        [void]$excel.Run('CopySelectedMoleculeSvgCommand')
        Start-Sleep -Milliseconds 400
        $c.Snap(4, $ptCopySvg.X, $ptCopySvg.Y)
        $c.ClickAt($ptBatch.X, $ptBatch.Y, 'Click Batch Style -> Publication B&W to restyle the sheet.', 'Batch Style -> Publication B&W로 시트 전체를 흑백 변환합니다.')
        [void]$excel.Run('BatchStyleBwCommand')
        Start-Sleep -Seconds 2
        $c.CaptureShot(@())
        $c.Snap(6, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520))
        $c.ClickAt($ptAbout.X, $ptAbout.Y, "Open About func_SMILES2IMG to verify version v$addinVersion and update status.", "About func_SMILES2IMG 창에서 v$addinVersion 버전과 라이선스를 확인합니다.")
        $c.RunWithModal('About func_SMILES2IMG', 2200, {
            Show-LiveAboutOverHost -OwnerHwnd $xHwnd -HoldMs 2200
        }, 'Inspect Third-Party Components, License Notices & Check for Updates.', '구성요소·라이선스 전문·업데이트 확인을 실행할 수 있습니다.')
        $c.Snap(5, -1, -1)
        Complete-FeatureClip $c
        # Restore color on sheet 1
        $ws.Range('C2:C5').Select()
        [void]$excel.Run('BatchStyleColorTransCommand')
        Start-Sleep -Seconds 1

        # --- 1.7 excel-legacy-float ---
        $wsFloat = $wb.Worksheets.Add()
        $wsFloat.Name = 'LegacyFloat'
        $wsFloat.Range('A1:D1').Value2 = @('Compound / SMILES', 'Floating Picture (=SMILES2IMG.FLOAT)', 'Name Lookup (.FLOAT)', 'MW')
        $wsFloat.Range('A1:D1').Font.Bold = $true
        $wsFloat.Columns.Item(1).ColumnWidth = 26
        $wsFloat.Columns.Item(2).ColumnWidth = 28
        $wsFloat.Columns.Item(3).ColumnWidth = 26
        $wsFloat.Columns.Item(4).ColumnWidth = 14
        $wsFloat.Range('A2').Value2 = $aspirinSmiles
        $wsFloat.Range('A3').Value2 = $caffeineSmiles
        $wsFloat.Rows.Item(2).RowHeight = 76
        $wsFloat.Rows.Item(3).RowHeight = 76
        $c = Start-FeatureClip -Id 'excel-legacy-float' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Excel 2016–2021 Compatibility (.FLOAT Formulas)' `
            -TitleKo '구버전 Excel 2016–2021 호환 (.FLOAT 수식)' `
            -CaptionEn 'Append .FLOAT to any image formula (=SMILES2IMG.FLOAT, =NAME2IMG.FLOAT) on Excel 2016-2021.' `
            -CaptionKo 'Excel 2016~2021에서는 함수명 뒤에 .FLOAT를 붙여 셀 크기에 맞춰 움직이는 그림을 만듭니다.'
        Select-ExcelCell $excel $wsFloat 'B2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520))
        $c.ClickAt((Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 520), 'Enter =SMILES2IMG.FLOAT(A2) and =NAME2IMG.FLOAT("Aspirin") in Row 2.', '2행에 =SMILES2IMG.FLOAT(A2)와 =NAME2IMG.FLOAT("Aspirin")을 입력합니다.')
        Invoke-ComRetry { $wsFloat.Range('B2').Formula2 = '=SMILES2IMG.FLOAT(A2)' }
        Invoke-ComRetry { $wsFloat.Range('C2').Formula2 = '=NAME2IMG.FLOAT("Aspirin")' }
        Invoke-ComRetry { $wsFloat.Range('D2').Formula2 = '=SMILES2MW(A2)' }
        Start-Sleep -Seconds 1
        $c.Snap(5, (Scale-RefX $xHwnd 860), (Scale-RefY $xHwnd 520))
        $c.ClickAt((Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 620), 'Fill Row 3 with =SMILES2IMG.FLOAT(A3) and =NAME2IMG.FLOAT("Caffeine").', '3행에도 =SMILES2IMG.FLOAT(A3)와 =NAME2IMG.FLOAT("Caffeine")을 채웁니다.')
        Invoke-ComRetry { $wsFloat.Range('B3').Formula2 = '=SMILES2IMG.FLOAT(A3)' }
        Invoke-ComRetry { $wsFloat.Range('C3').Formula2 = '=NAME2IMG.FLOAT("Caffeine")' }
        Invoke-ComRetry { $wsFloat.Range('D3').Formula2 = '=SMILES2MW(A3)' }
        Start-Sleep -Seconds 2
        Select-ExcelCell $excel $wsFloat 'B2' 400
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.8 excel-cheminformatics ---
        $wsChem = $wb.Worksheets.Add()
        $wsChem.Name = 'Cheminformatics'
        $wsChem.Range('A1:J1').Value2 = @('Name', 'SMILES', 'Structure', 'MW', 'ExactMass', 'LogP', 'TPSA', 'Rings', 'Aromatic', 'Fsp3')
        $wsChem.Range('A1:J1').Font.Bold = $true
        $wsChem.Range('A1:J1').Interior.Color = 0xDCFCE7
        $wsChem.Columns.Item(1).ColumnWidth = 10
        $wsChem.Columns.Item(2).ColumnWidth = 23
        $wsChem.Range('B2:B5').Font.Size = 8.5
        $wsChem.Columns.Item(3).ColumnWidth = 14
        $wsChem.Columns.Item(4).ColumnWidth = 8.5
        $wsChem.Columns.Item(5).ColumnWidth = 9.5
        $wsChem.Columns.Item(6).ColumnWidth = 6.5
        $wsChem.Columns.Item(7).ColumnWidth = 6.5
        $wsChem.Columns.Item(8).ColumnWidth = 6
        $wsChem.Columns.Item(9).ColumnWidth = 8
        $wsChem.Columns.Item(10).ColumnWidth = 6.5
        $chemMols = @(
            @('Aspirin', $aspirinSmiles),
            @('Caffeine', $caffeineSmiles),
            @('Ibuprofen', $ibuprofenSmiles),
            @('Nicotine', $nicotineSmiles)
        )
        for ($i = 0; $i -lt $chemMols.Count; $i++) {
            $r = $i + 2
            $wsChem.Rows.Item($r).RowHeight = 52
            $wsChem.Cells.Item($r, 1).Value2 = $chemMols[$i][0]
            $wsChem.Cells.Item($r, 2).Value2 = $chemMols[$i][1]
        }
        $c = Start-FeatureClip -Id 'excel-cheminformatics' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: MW, Exact Mass, LogP, TPSA & 3D Complexity' `
            -TitleKo '레시피: 분자량, 정밀질량, LogP, TPSA & 3D 복잡도' `
            -CaptionEn 'Compute physicochemical and 3D complexity descriptors locally offline.' `
            -CaptionKo '로컬 오프라인으로 분자량·LogP·TPSA·고리 수·Fsp3를 계산합니다.'
        Select-ExcelCell $excel $wsChem 'D2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 650), (Scale-RefY $xHwnd 500))
        $c.AddCue(1.8, 'Compute structures, MW, Exact Mass, LogP, TPSA & Fsp3 row by row.', '구조식과 분자량·정밀질량·LogP·TPSA·Fsp3를 순차 계산합니다.')
        for ($r = 2; $r -le 5; $r++) {
            $wsChem.Range("C$r").Formula2 = "=SMILES2IMG(B$r)"
            $wsChem.Range("D$r").Formula2 = "=SMILES2MW(B$r)"
            $wsChem.Range("E$r").Formula2 = "=SMILES2EXACTMASS(B$r)"
            $wsChem.Range("F$r").Formula2 = "=SMILES2LOGP(B$r)"
            $wsChem.Range("G$r").Formula2 = "=SMILES2TPSA(B$r)"
            $wsChem.Range("H$r").Formula2 = "=SMILES2RINGS(B$r)"
            $wsChem.Range("I$r").Formula2 = "=SMILES2AROMATIC(B$r)"
            $wsChem.Range("J$r").Formula2 = "=SMILES2FSP3(B$r)"
            Start-Sleep -Milliseconds 500
            $c.Snap(4, (Scale-RefX $xHwnd (580 + ($r - 2) * 140)), (Scale-RefY $xHwnd (490 + ($r - 2) * 80)))
        }
        Start-Sleep -Seconds 1
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        Select-ExcelCell $excel $wsChem 'D2' 420
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.9 excel-sar-screening ---
        $wsSar = $wb.Worksheets.Add()
        $wsSar.Name = 'SARScreening'
        $wsSar.Range('A1:E1').Value2 = @('Compound', 'SMILES', 'Highlight (c1ccccc1)', 'Has Benzene (=SMILES.MATCH)', 'Has Carboxyl (=SMILES.MATCH)')
        $wsSar.Range('A1:E1').Font.Bold = $true
        $wsSar.Columns.Item(1).ColumnWidth = 11
        $wsSar.Columns.Item(2).ColumnWidth = 35
        $wsSar.Columns.Item(3).ColumnWidth = 20
        $wsSar.Columns.Item(4).ColumnWidth = 22
        $wsSar.Columns.Item(5).ColumnWidth = 22
        for ($i = 0; $i -lt $chemMols.Count; $i++) {
            $r = $i + 2
            $wsSar.Rows.Item($r).RowHeight = 52
            $wsSar.Cells.Item($r, 1).Value2 = $chemMols[$i][0]
            $wsSar.Cells.Item($r, 2).Value2 = $chemMols[$i][1]
        }
        $c = Start-FeatureClip -Id 'excel-sar-screening' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: Substructure Search (SMILES.MATCH) & Scaffold Highlighting' `
            -TitleKo '레시피: 부분구조 검색(SMILES.MATCH) & 골격 하이라이트 스크리닝' `
            -CaptionEn 'Filter scaffolds with =SMILES.MATCH and highlight matching rings with hl:SMARTS.' `
            -CaptionKo '=SMILES.MATCH로 골격 포함 여부를 판별하고 hl: 옵션으로 매칭 고리를 강조합니다.'
        Select-ExcelCell $excel $wsSar 'C2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 600), (Scale-RefY $xHwnd 500))
        $c.AddCue(1.8, 'Run =SMILES.MATCH(B2,"c1ccccc1") and highlight benzene rings across rows.', '=SMILES.MATCH 부분구조 검색과 하이라이트 이미지를 순차 생성합니다.')
        for ($r = 2; $r -le 5; $r++) {
            $wsSar.Range("C$r").Formula2 = "=SMILES2IMG(B$r, `"`", `"hl:c1ccccc1`")"
            $wsSar.Range("D$r").Formula2 = "=SMILES.MATCH(B$r, `"c1ccccc1`")"
            $wsSar.Range("E$r").Formula2 = "=SMILES.MATCH(B$r, `"C(=O)O`")"
            Start-Sleep -Milliseconds 500
            $c.Snap(4, (Scale-RefX $xHwnd 720), (Scale-RefY $xHwnd (490 + ($r - 2) * 80)))
        }
        Start-Sleep -Seconds 1
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        Select-ExcelCell $excel $wsSar 'D2' 420
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.10 excel-conversions ---
        $wsConv = $wb.Worksheets.Add()
        $wsConv.Name = 'Conversions'
        $wsConv.Range('A1:E1').Value2 = @('Input Name/CAS', 'SMILES (=NAME2SMILES)', 'Standard InChI (=SMILES2INCHI)', 'InChIKey (=SMILES2INCHIKEY)', 'From InChI (=INCHI2IMG)')
        $wsConv.Range('A1:E1').Font.Bold = $true
        $wsConv.Columns.Item(1).ColumnWidth = 14
        $wsConv.Columns.Item(2).ColumnWidth = 22
        $wsConv.Columns.Item(3).ColumnWidth = 34
        $wsConv.Columns.Item(4).ColumnWidth = 28
        $wsConv.Columns.Item(5).ColumnWidth = 17
        $wsConv.Range('C:D').Font.Size = 9
        $wsConv.Range('C:D').WrapText = $true
        $wsConv.Range('A2').Value2 = '50-78-2'
        $wsConv.Range('A3').Value2 = 'Ibuprofen'
        $wsConv.Range('A4').Value2 = 'Caffeine'
        for ($r = 2; $r -le 4; $r++) { $wsConv.Rows.Item($r).RowHeight = 68 }
        $c = Start-FeatureClip -Id 'excel-conversions' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: Convert Between Name, CAS, SMILES, InChI & InChIKey' `
            -TitleKo '레시피: 이름 · CAS · SMILES · InChI · InChIKey 상호 변환' `
            -CaptionEn 'Cross-convert chemical identifiers and render structures directly from InChI strings.' `
            -CaptionKo 'CAS 번호, SMILES, InChI, InChIKey를 상호 변환하고 =INCHI2IMG로 렌더링합니다.'
        Select-ExcelCell $excel $wsConv 'B2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 480), (Scale-RefY $xHwnd 500))
        $c.AddCue(1.8, 'Convert CAS 50-78-2, Ibuprofen & Caffeine into SMILES, InChI & InChIKey.', 'CAS 50-78-2, Ibuprofen, Caffeine을 SMILES·InChI·InChIKey로 순차 변환합니다.')
        for ($r = 2; $r -le 4; $r++) {
            $wsConv.Range("B$r").Formula2 = "=NAME2SMILES(A$r)"
            $wsConv.Range("C$r").Formula2 = "=SMILES2INCHI(B$r)"
            $wsConv.Range("D$r").Formula2 = "=SMILES2INCHIKEY(B$r)"
            $wsConv.Range("E$r").Formula2 = "=INCHI2IMG(C$r)"
            Start-Sleep -Milliseconds 600
            $c.Snap(4, (Scale-RefX $xHwnd 780), (Scale-RefY $xHwnd (490 + ($r - 2) * 90)))
        }
        Start-Sleep -Seconds 1
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        Select-ExcelCell $excel $wsConv 'D2' 420
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 1.11 excel-sdf-workflow ---
        $sdfPath = (Resolve-Path (Join-Path $publicDir 'sample_compounds.sdf')).Path
        $c = Start-FeatureClip -Id 'excel-sdf-workflow' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Recipe: Import .SDF Chemical Database Tables' `
            -TitleKo '레시피: .SDF 화학 데이터베이스 11열 표 가져오기' `
            -CaptionEn 'Click [SMILES] -> [Import SDF] to generate an 11-column live SAR sheet.' `
            -CaptionKo '리본 [SMILES] -> [Import SDF]로 11개 물성 열이 갖춰진 엑셀 표를 즉시 생성합니다.'
        $c.Snap(4, $ptSdf.X, $ptSdf.Y)
        $c.ClickAt($ptSdf.X, $ptSdf.Y, 'Click [Import SDF] to load sample_compounds.sdf.', '[Import SDF]를 클릭해 sample_compounds.sdf 파일을 불러옵니다.')
        $wsSdf = Invoke-ComRetry { $wb.Worksheets.Add() }
        Invoke-ComRetry { $wsSdf.Name = 'SDFImport' }
        Invoke-ComRetry { $wsSdf.Activate() }
        $sdfRecords = [Smiles2Img.Rendering.SdfManager]::ReadSdfFile($sdfPath, 500)
        $sdfCols = @(
            @('Name',           $null,              12.5),
            @('Structure',      'SMILES2IMG',       11.5),
            @('SMILES',         $null,              17.0),
            @('Formula',        'SMILES2FORMULA',    7.5),
            @('MW (g/mol)',     'SMILES2MW',         9.0),
            @('Exact Mass',     'SMILES2EXACTMASS',  8.5),
            @('LogP',           'SMILES2LOGP',       5.0),
            @('TPSA',           'SMILES2TPSA',       5.5),
            @('Rings',          'SMILES2RINGS',      5.5),
            @('Aromatic Rings', 'SMILES2AROMATIC',  11.5),
            @('Fsp3',           'SMILES2FSP3',       5.5)
        )
        for ($ci = 0; $ci -lt $sdfCols.Count; $ci++) {
            $colNum = $ci + 1
            Invoke-ComRetry {
                $hc = $wsSdf.Cells.Item(1, $colNum)
                $hc.Value2 = $sdfCols[$ci][0]
                $hc.Font.Bold = $true
                $hc.Font.Size = 9.0
                $hc.Interior.Color = 0xFFE7E0
                $wsSdf.Columns.Item($colNum).ColumnWidth = $sdfCols[$ci][2]
            }
        }
        Invoke-ComRetry { $wsSdf.Range('C2:C20').Font.Size = 8.0 }
        $imported = 0
        for ($i = 0; $i -lt $sdfRecords.Count; $i++) {
            $rec = $sdfRecords[$i]
            $row = $i + 2
            $compName = if (-not [string]::IsNullOrWhiteSpace($rec.Name)) { $rec.Name } else { "Compound #$($i + 1)" }
            $compSmiles = $rec.Smiles
            Invoke-ComRetry {
                $wsSdf.Rows.Item($row).RowHeight = 52
                $wsSdf.Cells.Item($row, 1).Value2 = $compName
                $wsSdf.Cells.Item($row, 3).Value2 = $compSmiles
                for ($ci = 0; $ci -lt $sdfCols.Count; $ci++) {
                    $fn = $sdfCols[$ci][1]
                    if ($fn) {
                        $wsSdf.Cells.Item($row, $ci + 1).Formula2 = "=$fn(C$row)"
                    }
                }
            }
            $imported++
            if (($i % 2 -eq 1) -or ($i -eq $sdfRecords.Count - 1)) {
                Start-Sleep -Milliseconds 350
                $c.Snap(3, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd (480 + [Math]::Min($i, 3) * 75)))
            }
        }
        Start-Sleep -Seconds 2
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        Select-ExcelCell $excel $wsSdf 'B2' 450
        $c.AddCue(1.8, "Imported $imported compounds into an 11-column structure & complexity worksheet.", "$imported 개 화합물을 11열 구조·물성·3D 복잡도 시트로 불러왔습니다.")
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c
        }

        if ($Section -eq 'All' -or $Section -eq 'Excel' -or $Section -eq 'UseCases') {
        # --- 1.12 excel-usecase-eln-stoichiometry ---
        $wsEln = Invoke-ComRetry { $wb.Worksheets.Add() }
        Invoke-ComRetry { $wsEln.Name = 'ELNStoichiometry' }
        Invoke-ComRetry { $wsEln.Activate() }
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        $wsEln.Range('A1:G1').Value2 = @('Reagent / CAS', 'SMILES (=NAME2SMILES)', 'Structure (=SMILES2IMG)', 'MW (=SMILES2MW)', 'Eq.', 'mmol', 'Mass (mg)')
        $wsEln.Range('A1:G1').Font.Bold = $true
        $wsEln.Range('A1:G1').Interior.Color = 0xEFE7DC
        $wsEln.Columns.Item(1).ColumnWidth = 17
        $wsEln.Columns.Item(2).ColumnWidth = 23
        $wsEln.Columns.Item(3).ColumnWidth = 22
        $wsEln.Columns.Item(4).ColumnWidth = 17
        $wsEln.Columns.Item(5).ColumnWidth = 7
        $wsEln.Columns.Item(6).ColumnWidth = 9
        $wsEln.Columns.Item(7).ColumnWidth = 13
        $wsEln.Range('B2:B4').Font.Size = 9
        $wsEln.Range('A2').Value2 = 'Salicylic Acid'
        $wsEln.Range('A3').Value2 = 'Acetic Anhydride'
        $wsEln.Range('A4').Value2 = '50-78-2'
        $wsEln.Range('E2').Value2 = 1.0
        $wsEln.Range('E3').Value2 = 1.5
        $wsEln.Range('E4').Value2 = 1.0
        $wsEln.Range('F2').Value2 = 5.0
        $wsEln.Range('F2').Interior.Color = 0xCCF5FF
        $wsEln.Range('F2').Font.Bold = $true
        $wsEln.Range('G2:G4').NumberFormat = '0.0 "mg"'
        $wsEln.Range('D2:D4').NumberFormat = '0.00'
        for ($r = 2; $r -le 4; $r++) { $wsEln.Rows.Item($r).RowHeight = 68 }
        $c = Start-FeatureClip -Id 'excel-usecase-eln-stoichiometry' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Use Case: Synthesis ELN & Stoichiometry Calculator' `
            -TitleKo '실전 시나리오: 합성 실험노트(ELN) & 당량·질량 자동 계산기' `
            -CaptionEn 'Populate SMILES, 2D structures, numeric MW, and reagent mg weights from names or CAS, then scale mmol dynamically.' `
            -CaptionKo '시약명·CAS 번호로 SMILES·구조식·분자량·질량(mg)을 채우고 mmol 배치를 변경해 무게를 자동 재계산합니다.'
        Select-ExcelCell $excel $wsEln 'B2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 520), (Scale-RefY $xHwnd 480))
        $c.AddCue(1.8, 'Resolve Salicylic Acid, Acetic Anhydride & CAS 50-78-2 into structures and mg weights.', '시약명과 CAS 번호(50-78-2)로 구조식·분자량·질량(mg)을 자동 계산합니다.')
        for ($r = 2; $r -le 4; $r++) {
            $wsEln.Range("B$r").Formula2 = "=NAME2SMILES(A$r)"
            $wsEln.Range("C$r").Formula2 = "=SMILES2IMG(B$r)"
            $wsEln.Range("D$r").Formula2 = "=SMILES2MW(B$r)"
            if ($r -gt 2) { $wsEln.Range("F$r").Formula2 = "=`$F`$2*E$r" }
            $wsEln.Range("G$r").Formula2 = "=ROUND(D$r*F$r, 1)"
            Start-Sleep -Milliseconds 600
            $c.Snap(4, (Scale-RefX $xHwnd 780), (Scale-RefY $xHwnd (490 + ($r - 2) * 90)))
        }
        Start-Sleep -Seconds 1
        Select-ExcelCell $excel $wsEln 'F2' 450
        $c.Snap(6, (Scale-RefX $xHwnd 1340), (Scale-RefY $xHwnd 490))
        $c.AddCue(1.8, 'Change reaction scale F2 from 5.0 to 10.0 mmol to recalculate every reagent weight.', '기준 스케일 F2를 5.0에서 10.0 mmol로 변경해 모든 시약 질량을 갱신합니다.')
        $wsEln.Range('F2').Value2 = 10.0
        Select-ExcelCell $excel $wsEln 'G4' 600
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        $c.CaptureShot(@())
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        # --- 1.13 excel-usecase-deduplication ---
        $wsDedup = Invoke-ComRetry { $wb.Worksheets.Add() }
        Invoke-ComRetry { $wsDedup.Name = 'Deduplication' }
        Invoke-ComRetry { $wsDedup.Activate() }
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        $wsDedup.Range('A1:E1').Value2 = @('Input SMILES / Name', 'Structure', 'Canonical SMILES', 'InChIKey (=SMILES2INCHIKEY)', 'Duplicate Flag')
        $wsDedup.Range('A1:E1').Font.Bold = $true
        $wsDedup.Range('A1:E1').Interior.Color = 0xEFE7DC
        $wsDedup.Columns.Item(1).ColumnWidth = 22
        $wsDedup.Columns.Item(2).ColumnWidth = 14
        $wsDedup.Columns.Item(3).ColumnWidth = 22
        $wsDedup.Columns.Item(4).ColumnWidth = 29
        $wsDedup.Columns.Item(5).ColumnWidth = 14
        $wsDedup.Range('A2:A5').Font.Size = 9.5
        $wsDedup.Range('C2:D5').Font.Size = 9
        $wsDedup.Range('A2').Value2 = 'c1ccccc1C(=O)O'
        $wsDedup.Range('A3').Value2 = 'O=C(O)c1ccccc1'
        $wsDedup.Range('A4').Value2 = 'Aspirin'
        $wsDedup.Range('A5').Value2 = 'Caffeine'
        for ($r = 2; $r -le 5; $r++) { $wsDedup.Rows.Item($r).RowHeight = 64 }
        $c = Start-FeatureClip -Id 'excel-usecase-deduplication' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Use Case: Reagent & Library Deduplication via InChIKey' `
            -TitleKo '실전 시나리오: InChIKey 기반 시약·화합물 중복 제거' `
            -CaptionEn 'Normalize different SMILES strings of the same molecule into identical InChIKeys and flag duplicates with COUNTIF.' `
            -CaptionKo '서로 다르게 표기된 동일 분자 SMILES를 InChIKey로 정규화해 COUNTIF로 중복 항목을 판별합니다.'
        Select-ExcelCell $excel $wsDedup 'C2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 600), (Scale-RefY $xHwnd 490))
        $c.AddCue(1.8, 'Generate Canonical SMILES and 27-character InChIKeys across mixed SMILES and names.', '서로 다른 SMILES 표기와 화합물명에서 표준 SMILES와 InChIKey를 생성합니다.')
        $wsDedup.Range('B2').Formula2 = '=SMILES2IMG(A2)'
        $wsDedup.Range('C2').Formula2 = '=INCHI2SMILES(SMILES2INCHI(A2))'
        $wsDedup.Range('D2').Formula2 = '=SMILES2INCHIKEY(A2)'
        $wsDedup.Range('B3').Formula2 = '=SMILES2IMG(A3)'
        $wsDedup.Range('C3').Formula2 = '=INCHI2SMILES(SMILES2INCHI(A3))'
        $wsDedup.Range('D3').Formula2 = '=SMILES2INCHIKEY(A3)'
        Start-Sleep -Milliseconds 700
        $c.Snap(5, (Scale-RefX $xHwnd 1280), (Scale-RefY $xHwnd 560))
        $wsDedup.Range('B4').Formula2 = '=NAME2IMG(A4)'
        $wsDedup.Range('C4').Formula2 = '=NAME2SMILES(A4)'
        $wsDedup.Range('D4').Formula2 = '=SMILES2INCHIKEY(C4)'
        $wsDedup.Range('B5').Formula2 = '=NAME2IMG(A5)'
        $wsDedup.Range('C5').Formula2 = '=NAME2SMILES(A5)'
        $wsDedup.Range('D5').Formula2 = '=SMILES2INCHIKEY(C5)'
        Start-Sleep -Milliseconds 800
        $c.Snap(5, (Scale-RefX $xHwnd 1280), (Scale-RefY $xHwnd 720))
        $c.AddCue(1.8, 'Evaluate COUNTIF on InChIKey: both Benzoic acid variants are flagged as Duplicate.', 'InChIKey 열에 COUNTIF를 적용해 표기가 다른 두 벤조산 항목을 중복(Duplicate)으로 검출합니다.')
        for ($r = 2; $r -le 5; $r++) {
            $wsDedup.Range("E$r").Formula2 = "=IF(COUNTIF(`$D`$2:`$D`$5, D$r)>1, `"Duplicate`", `"Unique`")"
        }
        $wsDedup.Range('E2:E3').Interior.Color = 0xD7D7FE
        $wsDedup.Range('E2:E3').Font.Bold = $true
        $wsDedup.Range('E4:E5').Interior.Color = 0xE7F7DC
        $wsDedup.Range('E2:E5').HorizontalAlignment = -4108
        Start-Sleep -Seconds 1
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        Select-ExcelCell $excel $wsDedup 'E2' 450
        $c.CaptureShot(@())
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        # --- 1.14 excel-usecase-rgroup-enumeration ---
        $wsRgrp = Invoke-ComRetry { $wb.Worksheets.Add() }
        Invoke-ComRetry { $wsRgrp.Name = 'RGroupEnum' }
        Invoke-ComRetry { $wsRgrp.Activate() }
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        $wsRgrp.Range('A1:G1').Value2 = @('R-Group', 'Template ($B$2)', 'Analog SMILES (=SUBSTITUTE)', '2D Structure (=SMILES2IMG)', 'MW', 'Formula', 'LogP')
        $wsRgrp.Range('A1:G1').Font.Bold = $true
        $wsRgrp.Range('A1:G1').Interior.Color = 0xEFE7DC
        $wsRgrp.Columns.Item(1).ColumnWidth = 10
        $wsRgrp.Columns.Item(2).ColumnWidth = 23
        $wsRgrp.Columns.Item(3).ColumnWidth = 25
        $wsRgrp.Columns.Item(4).ColumnWidth = 23
        $wsRgrp.Columns.Item(5).ColumnWidth = 9
        $wsRgrp.Columns.Item(6).ColumnWidth = 11
        $wsRgrp.Columns.Item(7).ColumnWidth = 8
        $wsRgrp.Range('B2:C6').Font.Size = 9
        $wsRgrp.Range('B2').Value2 = 'CC(=O)Nc1ccc([R])cc1'
        $wsRgrp.Range('B2').Interior.Color = 0xCCF5FF
        $wsRgrp.Range('B2').Font.Bold = $true
        $rGroups = @('O', 'F', 'C(F)(F)F', 'Cl', 'OC')
        for ($i = 0; $i -lt 4; $i++) {
            $r = $i + 2
            $wsRgrp.Rows.Item($r).RowHeight = 62
            $wsRgrp.Range("A$r").Value2 = $rGroups[$i]
        }
        $wsRgrp.Rows.Item(6).RowHeight = 62
        $c = Start-FeatureClip -Id 'excel-usecase-rgroup-enumeration' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Use Case: R-Group Analog Enumeration with SUBSTITUTE' `
            -TitleKo '실전 시나리오: SUBSTITUTE 기반 R-치환기 유도체 자동 생성' `
            -CaptionEn 'Swap [R] placeholders via Excel SUBSTITUTE to generate analog 2D structures, MW, Formula, and LogP offline.' `
            -CaptionKo '엑셀 SUBSTITUTE 함수로 [R] 치환기를 자동 치환해 유도체 구조식·분자량·화학식·LogP를 오프라인으로 생성합니다.'
        Select-ExcelCell $excel $wsRgrp 'C2' 350
        $c.Snap(4, (Scale-RefX $xHwnd 680), (Scale-RefY $xHwnd 490))
        $c.AddCue(1.8, 'Substitute R-groups (O, F, CF3, Cl) into template $B$2 and compute structures & properties.', '치환기(O, F, CF3, Cl)를 템플릿 $B$2에 치환해 구조식과 물성을 자동 계산합니다.')
        for ($r = 2; $r -le 5; $r++) {
            $wsRgrp.Range("C$r").Formula2 = "=SUBSTITUTE(`$B`$2, `"[R]`", A$r)"
            $wsRgrp.Range("D$r").Formula2 = "=SMILES2IMG(C$r)"
            $wsRgrp.Range("E$r").Formula2 = "=SMILES2MW(C$r)"
            $wsRgrp.Range("F$r").Formula2 = "=SMILES2FORMULA(C$r)"
            $wsRgrp.Range("G$r").Formula2 = "=SMILES2LOGP(C$r)"
            Start-Sleep -Milliseconds 500
            $c.Snap(4, (Scale-RefX $xHwnd 980), (Scale-RefY $xHwnd (490 + ($r - 2) * 85)))
        }
        $c.AddCue(1.8, 'Add a new methoxy fragment (OC) in A6 to render the 5th analog and its descriptors.', 'A6에 새 메톡시 치환기(OC)를 추가하면 5번째 유도체 구조와 물성이 바로 생성됩니다.')
        $wsRgrp.Range('A6').Value2 = 'OC'
        $wsRgrp.Range('C6').Formula2 = '=SUBSTITUTE($B$2, "[R]", A6)'
        $wsRgrp.Range('D6').Formula2 = '=SMILES2IMG(C6)'
        $wsRgrp.Range('E6').Formula2 = '=SMILES2MW(C6)'
        $wsRgrp.Range('F6').Formula2 = '=SMILES2FORMULA(C6)'
        $wsRgrp.Range('G6').Formula2 = '=SMILES2LOGP(C6)'
        Start-Sleep -Seconds 1
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        Select-ExcelCell $excel $wsRgrp 'C6' 450
        $c.CaptureShot(@())
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        # --- 1.15 excel-usecase-dynamic-filter ---
        $wsFilt = Invoke-ComRetry { $wb.Worksheets.Add() }
        Invoke-ComRetry { $wsFilt.Name = 'DynamicFilter' }
        Invoke-ComRetry { $wsFilt.Activate() }
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        $wsFilt.Rows.Item(1).RowHeight = 24
        $wsFilt.Rows.Item(2).RowHeight = 22
        $wsFilt.Range('A1').Value2 = 'Master Library (6 Compounds)'
        $wsFilt.Range('A1').Font.Bold = $true
        $wsFilt.Range('G1').Value2 = 'Cutoff LogP <='
        $wsFilt.Range('G1').Font.Bold = $true
        $wsFilt.Range('H1').Value2 = 4.2
        $wsFilt.Range('H1').Interior.Color = 0xCCF5FF
        $wsFilt.Range('H1').Font.Bold = $true
        $wsFilt.Range('H1').HorizontalAlignment = -4108
        $wsFilt.Range('A2:E2').Value2 = @('Compound', 'Structure', 'MW', 'LogP', 'SMILES')
        $wsFilt.Range('A2:E2').Font.Bold = $true
        $wsFilt.Range('A2:E2').Interior.Color = 0xEFE7DC
        $wsFilt.Range('G2:J2').Value2 = @('Filtered Lead', 'Live Structure (=FILTER)', 'MW', 'LogP')
        $wsFilt.Range('G2:J2').Font.Bold = $true
        $wsFilt.Range('G2:J2').Interior.Color = 0xE7F7DC
        $wsFilt.Columns.Item(1).ColumnWidth = 12
        $wsFilt.Columns.Item(2).ColumnWidth = 13
        $wsFilt.Columns.Item(3).ColumnWidth = 8
        $wsFilt.Columns.Item(4).ColumnWidth = 7
        $wsFilt.Columns.Item(5).ColumnWidth = 18
        $wsFilt.Columns.Item(6).ColumnWidth = 3
        $wsFilt.Columns.Item(7).ColumnWidth = 13
        $wsFilt.Columns.Item(8).ColumnWidth = 22
        $wsFilt.Columns.Item(9).ColumnWidth = 8
        $wsFilt.Columns.Item(10).ColumnWidth = 7
        $wsFilt.Range('E3:E8').Font.Size = 8
        $filtData = @(
            @('Caffeine',    $caffeineSmiles),
            @('Paracetamol', $paracetamolSmiles),
            @('Nicotine',    $nicotineSmiles),
            @('Aspirin',     $aspirinSmiles),
            @('Ibuprofen',   $ibuprofenSmiles),
            @('Diclofenac',  'O=C(O)Cc1ccccc1Nc2c(Cl)cccc2Cl')
        )
        for ($i = 0; $i -lt $filtData.Count; $i++) {
            $r = $i + 3
            $wsFilt.Rows.Item($r).RowHeight = 56
            $wsFilt.Range("A$r").Value2 = $filtData[$i][0]
            $wsFilt.Range("E$r").Value2 = $filtData[$i][1]
            $wsFilt.Range("B$r").Formula2 = "=SMILES2IMG(E$r)"
            $wsFilt.Range("C$r").Formula2 = "=SMILES2MW(E$r)"
            $wsFilt.Range("D$r").Formula2 = "=SMILES2LOGP(E$r)"
        }
        $wsFilt.Range('G3').Formula2 = '=FILTER(A3:D8, D3:D8<=H1)'
        Start-Sleep -Seconds 2
        Select-ExcelCell $excel $wsFilt 'G3' 500
        [void][LiveCaptureEngine]::DismissUiaButton($xHwnd, $xPid, 'Got it')
        Start-Sleep -Milliseconds 350
        $c = Start-FeatureClip -Id 'excel-usecase-dynamic-filter' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Use Case: Dynamic In-Cell Structure Filtering (=FILTER)' `
            -TitleKo '실전 시나리오: Microsoft 365 동적 배열(=FILTER) 구조식 자동 필터링' `
            -CaptionEn 'Spill live in-cell 2D chemical structures alongside MW and LogP dynamically as threshold cell H1 changes.' `
            -CaptionKo 'LogP 기준 셀(H1)을 변경하면 =FILTER 함수가 인셀 2D 구조식과 물성 행을 동적으로 다시 추출합니다.'
        Select-ExcelCell $excel $wsFilt 'G3' 400
        [void][LiveCaptureEngine]::DismissUiaButton($xHwnd, $xPid, 'Got it')
        $c.Snap(8, (Scale-RefX $xHwnd 1520), (Scale-RefY $xHwnd 560))
        $c.AddCue(1.8, 'Lower LogP threshold H1 from 4.2 to 2.0: high-LogP compounds drop out automatically.', '기준 셀 H1을 4.2에서 2.0으로 낮추면 고지용성 화합물이 구조식과 함께 자동 제외됩니다.')
        $wsFilt.Range('H1').Value2 = 2.0
        Select-ExcelCell $excel $wsFilt 'H1' 800
        $c.Snap(8, (Scale-RefX $xHwnd 1680), (Scale-RefY $xHwnd 420))
        $c.AddCue(1.6, 'Tighten cutoff H1 to 1.0 to narrow down polar leads, then restore to 2.0.', '기준을 1.0으로 더 조여 친수성 후보만 추린 뒤 2.0으로 복원합니다.')
        $wsFilt.Range('H1').Value2 = 1.0
        Select-ExcelCell $excel $wsFilt 'H1' 800
        $c.Snap(8, (Scale-RefX $xHwnd 1680), (Scale-RefY $xHwnd 420))
        $wsFilt.Range('H1').Value2 = 2.0
        Select-ExcelCell $excel $wsFilt 'G3' 700
        [void][LiveCaptureEngine]::DismissUiaButton($xHwnd, $xPid, 'Got it')
        Start-Sleep -Milliseconds 300
        $excel.ActiveWindow.ScrollColumn = 1
        $excel.ActiveWindow.ScrollRow = 1
        $c.CaptureShot(@())
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        # --- 1.16 excel-diagnostics-dialog ---
        $wsDiag = Invoke-ComRetry { $wb.Worksheets.Add() }
        Invoke-ComRetry { $wsDiag.Name = 'Diagnostics' }
        Invoke-ComRetry { $wsDiag.Activate() }
        try { $excel.ActiveWindow.Zoom = 85 } catch {}
        $wsDiag.Range('A1:D1').Value2 = @('Test Case', 'Input SMILES', 'Formula (=SMILES2IMG)', 'Expected Cause')
        $wsDiag.Range('A1:D1').Font.Bold = $true
        $wsDiag.Range('A1:D1').Interior.Color = 0xEFE7DC
        $wsDiag.Columns.Item(1).ColumnWidth = 18
        $wsDiag.Columns.Item(2).ColumnWidth = 24
        $wsDiag.Columns.Item(3).ColumnWidth = 18
        $wsDiag.Columns.Item(4).ColumnWidth = 26
        $wsDiag.Range('A2').Value2 = 'Valid Aspirin'
        $wsDiag.Range('B2').Value2 = $aspirinSmiles
        $wsDiag.Range('D2').Value2 = 'Valid SMILES -> 2D image'
        $wsDiag.Range('A3').Value2 = 'Unclosed Ring'
        $wsDiag.Range('B3').Value2 = 'c1ccccc'
        $wsDiag.Range('D3').Value2 = 'Unclosed ring #1 (#VALUE!)'
        $wsDiag.Range('A4').Value2 = 'Unmatched Bracket'
        $wsDiag.Range('B4').Value2 = 'CC(=O'
        $wsDiag.Range('D4').Value2 = 'Syntax / parenthesis (#VALUE!)'
        $wsDiag.Range('A5').Value2 = 'Bad Valence'
        $wsDiag.Range('B5').Value2 = 'C(C)(C)(C)(C)C'
        $wsDiag.Range('D5').Value2 = 'Pentavalent carbon (#VALUE!)'
        $wsDiag.Rows.Item(2).RowHeight = 64
        for ($r = 3; $r -le 5; $r++) { $wsDiag.Rows.Item($r).RowHeight = 32 }
        for ($r = 2; $r -le 5; $r++) { $wsDiag.Range("C$r").Formula2 = "=SMILES2IMG(B$r)" }
        Start-Sleep -Seconds 1
        $c = Start-FeatureClip -Id 'excel-diagnostics-dialog' -HostApp 'excel' -MainHwnd $xHwnd -OfficePid $xPid `
            -TitleEn 'Excel Diagnostics: Inspect #VALUE! Indigo Parser Errors' `
            -TitleKo 'Excel 진단 로그(Diagnostics): #VALUE! 화학 문법 오류 원인 확인' `
            -CaptionEn 'Click [SMILES] -> [Diagnostics] to read exact EPAM Indigo parser messages behind #VALUE! cells.' `
            -CaptionKo '#VALUE! 오류가 발생했을 때 리본 [Diagnostics]로 정확한 Indigo 파서 오류 원인을 확인합니다.'
        Select-ExcelCell $excel $wsDiag 'C3' 450
        $c.Snap(6, $ptDiag.X, $ptDiag.Y)
        $c.ClickAt($ptDiag.X, $ptDiag.Y, 'Invalid SMILES returns #VALUE!. Click [Diagnostics] on the [SMILES] ribbon.', '잘못된 SMILES가 #VALUE!를 반환했습니다. 리본의 [Diagnostics] 버튼을 클릭합니다.')
        [void](Invoke-ComRetry { $excel.Run('ShowDiagnostics') })
        Start-Sleep -Milliseconds 900
        $diagHwnd = [LiveCaptureEngine]::FindChildWindowByTitle($xHwnd, $xPid, 'Diagnostic')
        if ($diagHwnd -eq [IntPtr]::Zero) { $diagHwnd = [LiveCaptureEngine]::FindChildWindowByTitle($xHwnd, $xPid, 'Log') }
        if ($diagHwnd -eq [IntPtr]::Zero) { $diagHwnd = [LiveCaptureEngine]::FindChildWindowByTitle($xHwnd, $xPid, '') }
        if ($diagHwnd -ne [IntPtr]::Zero) {
            [LiveCaptureEngine]::PositionChildWindowRelative(
                $xHwnd, $diagHwnd,
                (Scale-RefX $xHwnd 820), (Scale-RefY $xHwnd 920),
                (Scale-RefX $xHwnd 1820), (Scale-RefY $xHwnd 780)
            )
            Start-Sleep -Milliseconds 500
        }
        $c.AddCue(2.4, 'The Diagnostics window lists exact Indigo parser errors for each failing formula.', 'Diagnostics 창에 오류를 일으킨 함수명과 정확한 EPAM Indigo 파서 메시지가 기록됩니다.')
        $c.CaptureShot(@())
        $c.Snap(16, -1, -1)
        if ($diagHwnd -ne [IntPtr]::Zero) {
            [LiveCaptureEngine]::CloseWindow($diagHwnd)
            Start-Sleep -Milliseconds 400
        }
        $c.Snap(6, -1, -1)
        Complete-FeatureClip $c
        }

        $wb.Close($false)
    } finally {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# =============================================================================
# 2. MICROSOFT WORD LIVE CAPTURE (7 Features + Ribbon + Showcase)
# =============================================================================
if ($Section -eq 'All' -or $Section -eq 'Word' -or $Section -eq 'UseCases') {
    Write-Host "`n=== [2/4] Capturing Live Microsoft Word Manuscript Tutorial (6 Chapters) ===" -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    try {
        $wordAddIn = $null
        foreach ($ai in $word.COMAddIns) { if ($ai.ProgId -eq 'Smiles2Img.Word.AddIn') { $wordAddIn = $ai; break } }
        if (-not $wordAddIn) { throw 'Smiles2Img.Word.AddIn not registered' }
        if (-not $wordAddIn.Connect) { $wordAddIn.Connect = $true }
        $wordObj = $wordAddIn.Object

        $doc = $word.Documents.Add()
        $doc.PageSetup.TopMargin = 36
        $doc.PageSetup.BottomMargin = 36
        $doc.PageSetup.LeftMargin = 48
        $doc.PageSetup.RightMargin = 48
        $word.ActiveWindow.View.Zoom.Percentage = 85
        $wHwnd = [IntPtr]$word.ActiveWindow.Hwnd
        $wPid = [LiveCaptureEngine]::GetPidForHwnd($wHwnd)
        [void](Set-StandardHostWindow $wHwnd 'Word')
        Start-Sleep -Milliseconds 700
        [void][LiveCaptureEngine]::SelectRibbonTab($wHwnd, 'SMILES')
        Start-Sleep -Milliseconds 600

        $ptInsert   = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Insert / Edit Structure'))
        $ptMolstar  = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('3D Structure (Mol*)'))
        $ptPaste    = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Paste Structure'))
        $ptCopySvg  = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Copy SVG'))
        $ptMolInfo  = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Molecule Info'))
        $ptBatch    = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Batch Style'))

        if ($Section -eq 'All' -or $Section -eq 'Word') {
        # --- Helper to build manuscript section cleanly ---
        $p1 = $doc.Paragraphs.Add()
        $p1.Range.Text = "Synthesis and Structural Characterization of Aspirin Analogs"
        $p1.Range.Font.Name = "Segoe UI"
        $p1.Range.Font.Size = 18
        $p1.Range.Font.Bold = 1
        $p1.Range.InsertParagraphAfter()

        $pSub = $doc.Paragraphs.Add()
        $pSub.Range.Text = "Research Manuscript Draft - Prepared with func_SMILES2IMG for Microsoft Word"
        $pSub.Range.Font.Name = "Segoe UI"
        $pSub.Range.Font.Size = 10
        $pSub.Range.Font.Color = 0x64748B
        $pSub.Range.InsertParagraphAfter()

        # --- 2.1 Ch.1 Insert Structure with Ketcher ---
        $pCh1 = $doc.Paragraphs.Add()
        $pCh1.Range.Text = "1. Introduction: Acetylsalicylic acid (Aspirin) inhibits COX-1/2 via irreversible acetylation:"
        $pCh1.Range.Font.Name = "Segoe UI"
        $pCh1.Range.Font.Size = 11
        $pCh1.Range.Font.Bold = 0
        $pCh1.Range.Font.Color = 0x1E293B
        $pCh1.Range.InsertParagraphAfter()

        $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).Select()
        $c = Start-FeatureClip -Id 'word-ch1-insert-ketcher' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ch. 1: Insert 2D Structure into Manuscript Prose' `
            -TitleKo 'Word 1장: 논문 서론 본문에 Ketcher로 2D 구조식 삽입' `
            -CaptionEn 'Place the cursor in Word, click [Insert / Edit Structure], sketch in Ketcher, and apply.' `
            -CaptionKo '원고 본문에 커서를 두고 [Insert / Edit Structure]로 구조식을 그려 삽입합니다.'
        $c.Snap(4, $ptInsert.X, $ptInsert.Y)
        $c.ClickAt($ptInsert.X, $ptInsert.Y, 'Click [Insert / Edit Structure] on the Word SMILES ribbon.', 'Word SMILES 리본에서 [Insert / Edit Structure]를 클릭합니다.')
        $c.RunWithModal('Ketcher', 2000, {
            Show-LiveKetcherOverHost -OwnerHwnd $wHwnd -Context 'Word' -InitialInput $aspirinSmiles -HoldMs 2000
        }, 'Press Ctrl+Enter (Apply) to insert Aspirin inline at the cursor.', 'Ctrl+Enter(적용)를 눌러 커서 위치에 고화질 인라인 구조식을 삽입합니다.')
        [void][Smiles2Img.Editor.SafeClipboard]::SetText($aspirinSmiles)
        $rCh1Ins = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $rCh1Ins.Select()
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($doc.InlineShapes.Count -gt 0) {
            $picCh1 = $doc.InlineShapes.Item($doc.InlineShapes.Count)
            $picCh1.LockAspectRatio = -1
            $picCh1.Width = 85
        }
        $word.Selection.InsertParagraphAfter()
        $word.ActiveWindow.VerticalPercentScrolled = 0
        Start-Sleep -Milliseconds 450
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 2.2 Ch.2 Quick Insert from Selected SMILES Text ---
        $pCh2 = $doc.Paragraphs.Add()
        $pCh2.Range.Text = "2. Lead Analog (Caffeine):"
        $pCh2.Range.Font.Bold = 1
        $pCh2.Range.InsertParagraphAfter()
        $pCh2Mol = $doc.Paragraphs.Add()
        $pCh2Mol.Range.Text = $caffeineSmiles
        $pCh2Mol.Range.Font.Bold = 0
        $pCh2Mol.Range.Select()
        $c = Start-FeatureClip -Id 'word-ch2-quick-insert' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ch. 2: Convert Highlighted SMILES Text In-Place (Paste Structure)' `
            -TitleKo 'Word 2장: 본문 속 SMILES 텍스트 선택 후 구조식으로 변환' `
            -CaptionEn 'Highlight a SMILES string in Word prose and click [Paste Structure] to replace it in-place.' `
            -CaptionKo '본문에서 SMILES 문자열을 드래그 선택하고 [Paste Structure]를 누르면 해당 위치에서 구조식으로 치환됩니다.'
        $word.ActiveWindow.VerticalPercentScrolled = 0
        $c.Snap(6, $ptPaste.X, $ptPaste.Y)
        $c.ClickAt($ptPaste.X, $ptPaste.Y, 'Highlight the SMILES string and click [Paste Structure].', '선택된 SMILES 상태에서 [Paste Structure]를 클릭합니다.')
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($doc.InlineShapes.Count -gt 0) {
            $picCh2 = $doc.InlineShapes.Item($doc.InlineShapes.Count)
            $picCh2.LockAspectRatio = -1
            $picCh2.Width = 85
        }
        $word.Selection.InsertParagraphAfter()
        $word.ActiveWindow.VerticalPercentScrolled = 0
        Start-Sleep -Milliseconds 500
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 2.3 Ch.3 Compound Table & Molecule Info ---
        $pCh3 = $doc.Paragraphs.Add()
        $pCh3.Range.Text = "3. Compound Summary Table (Properties & 2D Structures):"
        $pCh3.Range.Font.Bold = 1
        $pCh3.Range.InsertParagraphAfter()
        $tblRange = $doc.Paragraphs.Add().Range
        $tbl = $doc.Tables.Add($tblRange, 3, 4)
        $tbl.Borders.Enable = 1
        $tbl.Cell(1, 1).Range.Text = 'Compound'
        $tbl.Cell(1, 2).Range.Text = '2D Structure'
        $tbl.Cell(1, 3).Range.Text = 'Formula'
        $tbl.Cell(1, 4).Range.Text = 'MW (g/mol)'
        $tbl.Rows.Item(1).Range.Font.Bold = 1
        $tbl.Rows.Item(1).Shading.BackgroundPatternColor = 0xF1F5F9
        $tbl.Cell(2, 1).Range.Text = 'Aspirin'
        $tbl.Cell(2, 3).Range.Text = 'C9H8O4'
        $tbl.Cell(2, 4).Range.Text = '180.16'
        $tbl.Cell(2, 2).Range.Text = $aspirinSmiles
        $rC2 = $tbl.Cell(2, 2).Range
        $rC2.End = $rC2.End - 1
        $rC2.Select()
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($tbl.Cell(2, 2).Range.InlineShapes.Count -gt 0) {
            $picT2 = $tbl.Cell(2, 2).Range.InlineShapes.Item(1)
            $picT2.LockAspectRatio = -1
            $picT2.Width = 72
        }
        $tbl.Cell(3, 1).Range.Text = 'Ibuprofen'
        $tbl.Cell(3, 3).Range.Text = 'C13H18O2'
        $tbl.Cell(3, 4).Range.Text = '206.28'
        $tbl.Cell(3, 2).Range.Text = $ibuprofenSmiles
        $rC3 = $tbl.Cell(3, 2).Range
        $rC3.End = $rC3.End - 1
        $rC3.Select()
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($tbl.Cell(3, 2).Range.InlineShapes.Count -gt 0) {
            $picT3 = $tbl.Cell(3, 2).Range.InlineShapes.Item(1)
            $picT3.LockAspectRatio = -1
            $picT3.Width = 72
        }

        $doc.InlineShapes.Item($doc.InlineShapes.Count).Range.Select()
        $word.ActiveWindow.VerticalPercentScrolled = 0
        $c = Start-FeatureClip -Id 'word-ch3-compound-table' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ch. 3: Compound Tables & Molecule Info Copy (TSV / Markdown)' `
            -TitleKo 'Word 3장: 논문 화합물 표 삽입 및 Molecule Info 물성 복사' `
            -CaptionEn 'Insert structures inside Word table cells and open Molecule Info to copy 12-col TSV rows.' `
            -CaptionKo '워드 표 셀 안에 구조식을 넣고 Molecule Info 창에서 계산된 물성과 12열 TSV를 복사합니다.'
        $c.Snap(4, $ptMolInfo.X, $ptMolInfo.Y)
        $c.ClickAt($ptMolInfo.X, $ptMolInfo.Y, 'Select a structure in the table and open Molecule Info.', '표 안의 구조식을 선택하고 Molecule Info를 엽니다.')
        $c.RunWithModal('Molecule Info', 2200, {
            Show-LiveMoleculeInfoOverHost -OwnerHwnd $wHwnd -SmilesOrKet $ibuprofenSmiles -HoldMs 2200
        }, 'Inspect calculated property rows or export a 12-column TSV row straight into Word.', '물성 행을 확인하고 12열 TSV 행이나 개별 물성을 복사할 수 있습니다.')
        $c.Snap(5, -1, -1)
        Complete-FeatureClip $c

        # --- Save word-overview & word-showcase right here cleanly framed from top of Page 1 ---
        $doc.Range(0, 0).Select()
        $word.ActiveWindow.VerticalPercentScrolled = 0
        Start-Sleep -Milliseconds 350
        $cOv = Start-FeatureClip -Id 'word-overview' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Overview: Manuscript Paragraphs, Tables & SMILES Ribbon' `
            -TitleKo 'Word 개요: 학술 논문 본문·요약 표 및 SMILES 리본' `
            -CaptionEn 'Full Microsoft Word manuscript powered by func_SMILES2IMG inline structures and table integration.' `
            -CaptionKo 'func_SMILES2IMG 인라인 구조식과 요약 표로 작성된 실제 Microsoft Word 논문 원고 화면입니다.'
        $cOv.CaptureShot(@(
            (Join-Path $publicDir 'word-showcase.png'),
            (Join-Path $assetsDir 'word-showcase.png'),
            (Join-Path $publicDir 'word-ribbon-guide.png'),
            (Join-Path $assetsDir 'word-ribbon-guide.png')
        ))
        $cOv.MoveTo($ptInsert.X, $ptInsert.Y, 6)
        $cOv.AddCue(1.8, 'Insert 2D structures, reaction schemes, and 3D proteins right inside Word.', 'Word 안에서 2D 구조식·반응식·3D 단백질을 자유롭게 삽입하고 편집합니다.')
        $cOv.MoveTo($ptBatch.X, $ptBatch.Y, 8)
        $cOv.Snap(6, $ptBatch.X, $ptBatch.Y)
        Complete-FeatureClip $cOv

        # --- 2.4 Ch.4 Reaction Scheme (KET Metadata) ---
        $pCh4 = $doc.Paragraphs.Add()
        $pCh4.Range.Text = "4. Synthetic Scheme 1 (Salicylic Acid + Acetic Anhydride -> Aspirin, KET):"
        $pCh4.Range.Font.Bold = 1
        $pCh4.Range.InsertParagraphAfter()
        $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).Select()
        $word.ActiveWindow.ScrollIntoView($pCh4.Range, $true)
        $c = Start-FeatureClip -Id 'word-ch4-reaction-scheme' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ch. 4: Insert Multi-Step Reaction Schemes with Lossless KET' `
            -TitleKo 'Word 4장: 다단계 합성 반응식 삽입 및 KET 무손실 보존' `
            -CaptionEn 'Draw reaction schemes in Ketcher; arrows and catalyst annotations survive save & reopen.' `
            -CaptionKo '반응 화살표와 촉매 주석이 포함된 반응식도 AlternativeText KET 데이터로 함께 보존됩니다.'
        $c.Snap(4, $ptInsert.X, $ptInsert.Y)
        $c.ClickAt($ptInsert.X, $ptInsert.Y, 'Open Ketcher from Word to sketch or inspect the Aspirin reaction scheme.', 'Word에서 Ketcher를 열어 살리실산 -> 아스피린 반응식을 작성합니다.')
        $c.RunWithModal('Ketcher', 2000, {
            Show-LiveKetcherOverHost -OwnerHwnd $wHwnd -Context 'Word' -InitialInput $sampleReactionKet -HoldMs 2000
        }, 'Apply Scheme 1 inline into Word with lossless KET metadata.', '적용(Apply)을 눌러 KET 좌표가 보존된 Scheme 1을 본문에 삽입합니다.')
        [void][Smiles2Img.Editor.SafeClipboard]::SetText("KET:$sampleReactionKet")
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($doc.InlineShapes.Count -gt 0) {
            $picRxnW = $doc.InlineShapes.Item($doc.InlineShapes.Count)
            $picRxnW.LockAspectRatio = -1
            $picRxnW.Width = 310
        }
        $word.Selection.InsertParagraphAfter()
        $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).Select()
        $word.ActiveWindow.ScrollIntoView($pCh4.Range, $true)
        Start-Sleep -Milliseconds 500
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 2.5 Ch.5 Double-Click Re-Edit & 3D Protein (Mol*) ---
        $crn1Bytes = Get-Live1CrnTransparentPng -OwnerHwnd $wHwnd -Context 'Word Document'
        $rBreak = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $rBreak.InsertBreak(7) # wdPageBreak so Section 5 + 3D Figure 2 sit together cleanly at top of Page 2
        $pCh5 = $doc.Paragraphs.Add()
        $pCh5.Range.Text = "5. Macromolecular Target Structure (Figure 2: PDB 1CRN via Mol* 3D Viewer):"
        $pCh5.Range.Font.Bold = 1
        $pCh5.Range.InsertParagraphAfter()
        $rFig5 = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $rFig5.Select()
        $word.ActiveWindow.ScrollIntoView($pCh5.Range, $true)
        $c = Start-FeatureClip -Id 'word-ch5-edit-and-molstar' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ch. 5: In-Place Re-Editing & 3D Protein Figures (Mol*)' `
            -TitleKo 'Word 5장: 그림 더블클릭 수정 & 3D 단백질 Figure 삽입 (Mol*)' `
            -CaptionEn 'Click [3D Structure (Mol*)] to load PDB ID 1CRN and embed a publication figure.' `
            -CaptionKo '[3D Structure (Mol*)]로 1CRN 단백질을 불러와 논문 Figure로 삽입합니다.'
        $c.Snap(4, $ptMolstar.X, $ptMolstar.Y)
        $c.ClickAt($ptMolstar.X, $ptMolstar.Y, 'Click [3D Structure (Mol*)] to open the 3D protein viewer.', '[3D Structure (Mol*)]를 눌러 3D 단백질 뷰어를 엽니다.')
        $c.RunWithModal('Mol*', 2200, {
            [void](Show-LiveMolstarOverHost -OwnerHwnd $wHwnd -Context 'Word Document' -PdbId '1CRN' -HoldMs 2200 -AutoApply $false)
        }, 'Apply camera angle & Cartoon representation to embed Figure 2 in Word.', '원하는 각도와 Cartoon 스타일을 맞춘 뒤 Apply로 원고에 삽입합니다.')
        $tmpMol = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "w_mol_$([Guid]::NewGuid().ToString('N')).png")
        [System.IO.File]::WriteAllBytes($tmpMol, $crn1Bytes)
        $pMolImg = $doc.Paragraphs.Add()
        $pic3d = $pMolImg.Range.InlineShapes.AddPicture($tmpMol, $false, $true)
        $pic3d.LockAspectRatio = -1
        $pic3d.Width = 220
        $pic3d.AlternativeText = 'MOLSTAR:{"pdb":"1CRN"}'
        $pMolImg.Range.InsertParagraphAfter()
        Remove-Item -LiteralPath $tmpMol -Force -ErrorAction SilentlyContinue
        $pFigCap = $doc.Paragraphs.Add()
        $pFigCap.Range.Text = "Figure 2. Crystal structure of Crambin (PDB: 1CRN) rendered with Mol* 3D."
        $pFigCap.Range.Font.Size = 9.5
        $pFigCap.Range.Font.Bold = 0
        $doc.Range($pCh5.Range.Start, $pCh5.Range.Start).Select()
        $word.ActiveWindow.ScrollIntoView($pCh5.Range, $true)
        Start-Sleep -Milliseconds 600
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 2.6 Ch.6 Batch Style ACS B&W & Copy SVG ---
        $doc.Range($pCh3.Range.Start, $pCh3.Range.Start).Select()
        $word.ActiveWindow.ScrollIntoView($pCh3.Range, $true)
        $word.ActiveWindow.VerticalPercentScrolled = 12
        Start-Sleep -Milliseconds 400
        $c = Start-FeatureClip -Id 'word-ch6-publication-finish' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ch. 6: Pre-Submission ACS B&W Batch Style & Vector SVG' `
            -TitleKo 'Word 6장: 투고 전 ACS 흑백 일괄 통일(Batch Style) & SVG 복사' `
            -CaptionEn 'Convert every structure in the manuscript to ACS monochrome B&W and copy vector SVG.' `
            -CaptionKo '원고 전체 구조식을 ACS 흑백 선화로 일괄 전환하고 SVG 벡터를 복사합니다.'
        $c.Snap(5, $ptBatch.X, $ptBatch.Y)
        $c.ClickAt($ptBatch.X, $ptBatch.Y, 'Run [Batch Style -> Publication B&W] across the entire Word manuscript.', '[Batch Style -> Publication B&W]를 실행해 문서 전체를 흑백 선화로 변환합니다.')
        [void]$wordObj.ExecuteRibbonAction('OnBatchStyleBw')
        $doc.Range($pCh3.Range.Start, $pCh3.Range.Start).Select()
        $word.ActiveWindow.ScrollIntoView($pCh3.Range, $true)
        $word.ActiveWindow.VerticalPercentScrolled = 15
        Start-Sleep -Milliseconds 900
        $c.Snap(6, $ptBatch.X, $ptBatch.Y)
        $c.CaptureShot(@())
        $c.ClickAt($ptCopySvg.X, $ptCopySvg.Y, 'Select any figure and click [Copy SVG] for Adobe Illustrator.', '그림을 선택하고 [Copy SVG]로 Illustrator용 무손실 벡터를 복사합니다.')
        $doc.InlineShapes.Item(1).Range.Select()
        [void]$wordObj.ExecuteRibbonAction('OnCopySvg')
        $c.Snap(6, -1, -1)
        Complete-FeatureClip $c

        # --- 2.7 Word Ribbon Tools Tour ---
        $ptAboutW = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('About func_SMILES2IMG'))
        [void]$wordObj.ExecuteRibbonAction('OnBatchStyleColorTrans')
        Start-Sleep -Milliseconds 600
        $doc.Range($p1.Range.Start, $p1.Range.Start).Select()
        $word.ActiveWindow.ScrollIntoView($p1.Range, $true)
        $word.ActiveWindow.VerticalPercentScrolled = 0
        Start-Sleep -Milliseconds 400
        $c = Start-FeatureClip -Id 'word-ribbon-tools' -HostApp 'word' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Word Ribbon: [SMILES] Insert, Copy SVG, Molecule Info & Batch Style' `
            -TitleKo 'Word 리본: [SMILES] 구조 삽입 · SVG 복사 · 물성 · 일괄 스타일' `
            -CaptionEn 'Explore every command on the Microsoft Word [SMILES] ribbon tab.' `
            -CaptionKo 'Microsoft Word 상단 [SMILES] 리본 탭의 주요 도구를 한눈에 확인합니다.'
        $c.Snap(4, $ptInsert.X, $ptInsert.Y)
        $c.MoveTo($ptInsert.X, $ptInsert.Y, 6)
        $c.AddCue(1.2, 'Insert / Edit Structure (Ketcher 2D) and 3D Structure (Mol*).', 'Insert / Edit Structure(Ketcher 2D)와 3D Structure(Mol*)로 2D·3D 구조를 삽입합니다.')
        $c.MoveTo($ptMolstar.X, $ptMolstar.Y, 6)
        $c.MoveTo($ptPaste.X, $ptPaste.Y, 6)
        $c.AddCue(1.2, 'Paste Structure converts selected SMILES or clipboard data in place.', 'Paste Structure로 선택 텍스트나 클립보드 SMILES를 제자리에서 구조식으로 변환합니다.')
        $c.MoveTo($ptCopySvg.X, $ptCopySvg.Y, 6)
        $c.MoveTo($ptMolInfo.X, $ptMolInfo.Y, 6)
        $c.AddCue(1.2, 'Copy SVG, Molecule Info, Batch Style, and About func_SMILES2IMG.', 'Copy SVG, Molecule Info, Batch Style, About func_SMILES2IMG를 바로 실행합니다.')
        $c.MoveTo($ptBatch.X, $ptBatch.Y, 6)
        $c.MoveTo($ptAboutW.X, $ptAboutW.Y, 6)
        $c.CaptureShot(@(
            (Join-Path $publicDir 'word-ribbon-guide.png'),
            (Join-Path $assetsDir 'word-ribbon-guide.png')
        ))
        $c.Snap(6, -1, -1)
        Complete-FeatureClip $c
        }

        try { $doc.Close($false) } catch {}

        if ($Section -eq 'All' -or $Section -eq 'Word' -or $Section -eq 'UseCases') {
        # --- 2.8 word-usecase-thesis-captions ---
        $docThesis = $word.Documents.Add()
        $docThesis.PageSetup.TopMargin = 36
        $docThesis.PageSetup.BottomMargin = 36
        $docThesis.PageSetup.LeftMargin = 48
        $docThesis.PageSetup.RightMargin = 48
        $word.ActiveWindow.View.Zoom.Percentage = 85
        $wHwndT = [IntPtr]$word.ActiveWindow.Hwnd
        [void](Set-StandardHostWindow $wHwndT 'Word (Thesis)')
        Start-Sleep -Milliseconds 500
        [void][LiveCaptureEngine]::SelectRibbonTab($wHwndT, 'SMILES')
        Start-Sleep -Milliseconds 400
        $ptBatchT = [LiveCaptureEngine]::FindControlRelativeCenter($wHwndT, @('Batch Style'))

        $pThTitle = $docThesis.Paragraphs.Add()
        $pThTitle.Range.Text = 'Chapter 3. Synthesis & Structural Verification of Acetylsalicylic Acid'
        $pThTitle.Range.Font.Name = 'Segoe UI'
        $pThTitle.Range.Font.Size = 16
        $pThTitle.Range.Font.Bold = 1
        $pThTitle.Range.Font.Color = 0x1E293B
        $pThTitle.Range.InsertParagraphAfter()

        $pTh1 = $docThesis.Paragraphs.Add()
        $pTh1.Range.Text = '3.1 Reaction Scheme (Word Captions & Cross-Reference Anchors)'
        $pTh1.Range.Font.Name = 'Segoe UI'
        $pTh1.Range.Font.Size = 12
        $pTh1.Range.Font.Bold = 1
        $pTh1.Range.InsertParagraphAfter()

        $pThBody1 = $docThesis.Paragraphs.Add()
        $pThBody1.Range.Text = 'Color master thesis draft for digital defense PDF and monochrome bound print copy (Scheme 3.1):'
        $pThBody1.Range.Font.Name = 'Segoe UI'
        $pThBody1.Range.Font.Size = 10.5
        $pThBody1.Range.Font.Bold = 0
        $pThBody1.Range.InsertParagraphAfter()

        $docThesis.Range($docThesis.Content.End - 1, $docThesis.Content.End - 1).Select()
        [void][Smiles2Img.Editor.SafeClipboard]::SetText($sampleReactionKet)
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($docThesis.InlineShapes.Count -ge 1) {
            $picTh1 = $docThesis.InlineShapes.Item($docThesis.InlineShapes.Count)
            $picTh1.LockAspectRatio = -1
            $picTh1.Width = 370
        }
        $word.Selection.Collapse(0)
        $word.Selection.InsertParagraphAfter()
        $pCap1 = $docThesis.Paragraphs.Add()
        $pCap1.Range.Text = 'Scheme 3.1. Acid-catalyzed acetylation of salicylic acid to aspirin (KET editable scheme).'
        $pCap1.Range.Font.Name = 'Segoe UI'
        $pCap1.Range.Font.Size = 9.5
        $pCap1.Range.Font.Italic = 1
        $pCap1.Range.Font.Color = 0x64748B
        $pCap1.Range.InsertParagraphAfter()

        $pTh2 = $docThesis.Paragraphs.Add()
        $pTh2.Range.Text = '3.2 Target Compound Characterization'
        $pTh2.Range.Font.Name = 'Segoe UI'
        $pTh2.Range.Font.Size = 12
        $pTh2.Range.Font.Bold = 1
        $pTh2.Range.Font.Italic = 0
        $pTh2.Range.Font.Color = 0x1E293B
        $pTh2.Range.InsertParagraphAfter()

        $docThesis.Range($docThesis.Content.End - 1, $docThesis.Content.End - 1).Select()
        [void][Smiles2Img.Editor.SafeClipboard]::SetText($aspirinSmiles)
        [void]$wordObj.ExecuteRibbonAction('OnQuickInsert')
        if ($docThesis.InlineShapes.Count -ge 2) {
            $picTh2 = $docThesis.InlineShapes.Item($docThesis.InlineShapes.Count)
            $picTh2.LockAspectRatio = -1
            $picTh2.Width = 145
        }
        $word.Selection.Collapse(0)
        $word.Selection.InsertParagraphAfter()
        if ($docThesis.InlineShapes.Count -lt 2) { throw "Thesis capture expected 2 inline structures, got $($docThesis.InlineShapes.Count)" }
        $pCap2 = $docThesis.Paragraphs.Add()
        $pCap2.Range.Text = 'Figure 3.1. Target compound 1a (Acetylsalicylic acid, C9H8O4, MW 180.16 g/mol).'
        $pCap2.Range.Font.Name = 'Segoe UI'
        $pCap2.Range.Font.Size = 9.5
        $pCap2.Range.Font.Italic = 1
        $pCap2.Range.Font.Color = 0x64748B
        $pCap2.Range.InsertParagraphAfter()

        $docThesis.Range($pThBody1.Range.Start, $pThBody1.Range.Start).Select()
        $word.ActiveWindow.ScrollIntoView($pThTitle.Range, $true)
        $word.ActiveWindow.VerticalPercentScrolled = 0
        Start-Sleep -Milliseconds 600

        $c = Start-FeatureClip -Id 'word-usecase-thesis-captions' -HostApp 'word' -MainHwnd $wHwndT -OfficePid $wPid `
            -TitleEn 'Use Case: Long Thesis Dual Color/B&W & Preserved Captions' `
            -TitleKo '실전 시나리오: 학위논문 캡션·앵커 유지 및 인쇄용 흑백 일괄 전환' `
            -CaptionEn 'Switch a multi-section thesis draft from color PDF to ACS B&W print copy without disturbing Word captions or anchors.' `
            -CaptionKo '학위논문 초안의 컬러 구조식과 반응식을 캡션 번호·문단 앵커 손상 없이 인쇄용 ACS 흑백으로 일괄 전환합니다.'
        $c.Snap(8, $ptBatchT.X, $ptBatchT.Y)
        $c.ClickAt($ptBatchT.X, $ptBatchT.Y, 'Color thesis master with Scheme 3.1 & Figure 3.1. Run [Batch Style -> Publication B&W].', 'Scheme 3.1과 Figure 3.1 캡션이 연결된 컬러 논문 초안에서 [Batch Style -> Publication B&W]를 실행합니다.')
        [void]$wordObj.ExecuteRibbonAction('OnBatchStyleBw')
        $docThesis.Range($pThBody1.Range.Start, $pThBody1.Range.Start).Select()
        $word.ActiveWindow.ScrollIntoView($pThTitle.Range, $true)
        $word.ActiveWindow.VerticalPercentScrolled = 0
        Start-Sleep -Milliseconds 900
        $c.AddCue(2.2, 'Both figures convert to monochrome line art while Scheme 3.1 and Figure 3.1 captions remain intact.', '두 그림이 흑백 선화로 변환되면서도 Scheme 3.1·Figure 3.1 캡션과 배치가 그대로 유지됩니다.')
        $c.CaptureShot(@())
        $c.Snap(14, -1, -1)
        Complete-FeatureClip $c
        $docThesis.Close($false)
        }
    } finally {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# =============================================================================
# 3. MICROSOFT POWERPOINT LIVE CAPTURE (8 Features + Ribbon + Showcase)
# =============================================================================
if ($Section -eq 'All' -or $Section -eq 'PowerPoint' -or $Section -eq 'UseCases') {
    Write-Host "`n=== [3/4] Capturing Live Microsoft PowerPoint Slide Gallery ===" -ForegroundColor Cyan
    $ppt = New-Object -ComObject PowerPoint.Application
    try {
        $pptAddIn = $null
        foreach ($ai in $ppt.COMAddIns) { if ($ai.ProgId -eq 'Smiles2Img.PowerPoint.AddIn') { $pptAddIn = $ai; break } }
        if (-not $pptAddIn) { throw 'Smiles2Img.PowerPoint.AddIn not registered' }
        if (-not $pptAddIn.Connect) { $pptAddIn.Connect = $true }
        $pptObj = $pptAddIn.Object

        $pres = $ppt.Presentations.Add(-1)
        $pres.PageSetup.SlideWidth = 960
        $pres.PageSetup.SlideHeight = 540
        $pHwnd = [IntPtr]$ppt.HWND
        $pPid = [LiveCaptureEngine]::GetPidForHwnd($pHwnd)
        try { $ppt.ActiveWindow.WindowState = 1 } catch {} # ppWindowNormal
        [void](Set-StandardHostWindow $pHwnd 'PowerPoint')
        Start-Sleep -Milliseconds 700
        [void][LiveCaptureEngine]::SelectRibbonTab($pHwnd, 'SMILES')
        Start-Sleep -Milliseconds 600

        function Reset-PowerPointSlideZoom {
            try {
                # Flush WM_PRINT once so PowerPoint DirectX viewport syncs to resized HiDPI bounds even when display is idle
                $tmpBmp = New-Object System.Drawing.Bitmap(16, 16)
                $tmpG = [System.Drawing.Graphics]::FromImage($tmpBmp)
                $tmpHdc = $tmpG.GetHdc()
                [void][LiveCaptureEngine]::PrintWindow($pHwnd, $tmpHdc, 0)
                $tmpG.ReleaseHdc($tmpHdc)
                $tmpG.Dispose()
                $tmpBmp.Dispose()
                Start-Sleep -Milliseconds 250
                $ppt.ActiveWindow.ViewType = 9 # ppViewNormal
                $ppt.ActiveWindow.Panes.Item(2).Activate()
                $ppt.ActiveWindow.View.ZoomToFit = -1
                Start-Sleep -Milliseconds 250
            } catch {}
        }

        $ptInsert  = [LiveCaptureEngine]::FindControlRelativeCenter($pHwnd, @('Insert / Edit Structure'))
        $ptMolstar = [LiveCaptureEngine]::FindControlRelativeCenter($pHwnd, @('3D Structure (Mol*)'))
        $ptPaste   = [LiveCaptureEngine]::FindControlRelativeCenter($pHwnd, @('Paste Structure'))
        $ptMolInfo = [LiveCaptureEngine]::FindControlRelativeCenter($pHwnd, @('Molecule Info'))
        $ptCopySvgP = [LiveCaptureEngine]::FindControlRelativeCenter($pHwnd, @('Copy SVG'))
        $ptBatch   = [LiveCaptureEngine]::FindControlRelativeCenter($pHwnd, @('Batch Style'))

        function New-DarkSlide($presObj, $titleText, $subText) {
            $s = $presObj.Slides.Add($presObj.Slides.Count + 1, 12)
            $s.FollowMasterBackground = 0
            $s.Background.Fill.Solid()
            $s.Background.Fill.ForeColor.RGB = 0x2A170F # BGR for #0F172A dark slate
            $tbTitle = $s.Shapes.AddTextbox(1, 36, 22, 880, 42)
            $tbTitle.TextFrame.TextRange.Text = $titleText
            $tbTitle.TextFrame.TextRange.Font.Name = 'Segoe UI'
            $tbTitle.TextFrame.TextRange.Font.Size = 22
            $tbTitle.TextFrame.TextRange.Font.Bold = -1
            $tbTitle.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
            $tbSub = $s.Shapes.AddTextbox(1, 36, 62, 880, 24)
            $tbSub.TextFrame.TextRange.Text = $subText
            $tbSub.TextFrame.TextRange.Font.Name = 'Segoe UI'
            $tbSub.TextFrame.TextRange.Font.Size = 12
            $tbSub.TextFrame.TextRange.Font.Color.RGB = 0xB8A394
            return $s
        }

        if ($Section -eq 'All' -or $Section -eq 'PowerPoint') {
        # --- 3.1 powerpoint-slide-candidate-grid ---
        $sGrid = New-DarkSlide $pres 'Lead Compound Candidates - Dark Slide Screening' 'Inserted via func_SMILES2IMG PowerPoint Add-in • Transparent Dark Mode White Ink'
        $sGrid.Select()
        Reset-PowerPointSlideZoom
        $cards = @(
            @{ Name = 'Aspirin'; Formula = 'C9H8O4'; Smiles = $aspirinSmiles },
            @{ Name = 'Caffeine'; Formula = 'C8H10N4O2'; Smiles = $caffeineSmiles },
            @{ Name = 'Nicotine'; Formula = 'C10H14N2'; Smiles = $nicotineSmiles },
            @{ Name = 'Paracetamol'; Formula = 'C8H9NO2'; Smiles = $paracetamolSmiles }
        )
        $cGrid = Start-FeatureClip -Id 'powerpoint-slide-candidate-grid' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: 4-Card Lead Candidate Grid' `
            -TitleKo 'PowerPoint 갤러리: 4종 후보 물질 비교 카드 슬라이드' `
            -CaptionEn 'Arrange 4 compounds on a dark conference slide with transparent backgrounds.' `
            -CaptionKo '어두운 발표 슬라이드 위에 하얀 배경 사각형 없이 4종 후보 물질 카드를 배치합니다.'
        $cGrid.Snap(4, $ptPaste.X, $ptPaste.Y)
        $cGrid.AddCue(2.0, 'Paste SMILES into each card and convert to Dark Mode (White Ink).', '각 후보 카드에 구조식을 삽입하고 투명 흰색 잉크(Dark Mode)로 통일합니다.')
        for ($i = 0; $i -lt $cards.Count; $i++) {
            $card = $cards[$i]
            $cx = 36 + $i * 224
            $box = $sGrid.Shapes.AddShape(1, $cx, 104, 208, 390)
            $box.Fill.Solid()
            $box.Fill.ForeColor.RGB = 0x3B291E # BGR for #1E293B
            $box.Line.ForeColor.RGB = 0x554133
            $t1 = $sGrid.Shapes.AddTextbox(1, $cx + 12, 114, 184, 30)
            $t1.TextFrame.TextRange.Text = $card.Name
            $t1.TextFrame.TextRange.Font.Size = 15
            $t1.TextFrame.TextRange.Font.Bold = -1
            $t1.TextFrame.TextRange.Font.Color.RGB = 0xF8BD38 # BGR Sky Blue #38BDF8
            $t2 = $sGrid.Shapes.AddTextbox(1, $cx + 12, 142, 184, 24)
            $t2.TextFrame.TextRange.Text = $card.Formula
            $t2.TextFrame.TextRange.Font.Size = 11
            $t2.TextFrame.TextRange.Font.Color.RGB = 0xE1D5CB

            $tbMol = $sGrid.Shapes.AddTextbox(1, $cx + 12, 180, 184, 40)
            $tbMol.TextFrame.TextRange.Text = $card.Smiles
            $tbMol.Select()
            [void]$pptObj.ExecuteRibbonAction('OnQuickInsert')
            try { $tbMol.Delete() } catch {}
            $pic = $sGrid.Shapes.Item($sGrid.Shapes.Count)
            $pic.LockAspectRatio = -1
            $pic.Width = 176
            if ($pic.Height -gt 176) { $pic.Height = 176 }
            $pic.Left = $cx + 14 + ((180 - $pic.Width) / 2)
            $pic.Top = 185 + ((180 - $pic.Height) / 2)
            $pic.Select()
            [void]$pptObj.ExecuteRibbonAction('OnBatchStyleDark')
            $cGrid.Snap(5, (Scale-RefX $pHwnd (520 + $i * 360)), (Scale-RefY $pHwnd 650))
        }
        $sGrid.Shapes.Item($sGrid.Shapes.Count).Select()
        Start-Sleep -Milliseconds 1200
        $cGrid.CaptureShot(@())
        $cGrid.Snap(6, -1, -1)
        Complete-FeatureClip $cGrid

        # --- 3.2 powerpoint-slide-dark-mode ---
        $sDark = New-DarkSlide $pres 'Dark Mode Presentation Structures (White Ink)' 'Transparent background - high-contrast white carbon bonds'
        $sDark.Select()
        $c = Start-FeatureClip -Id 'powerpoint-slide-dark-mode' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: Dark Mode (White Ink + Transparent)' `
            -TitleKo 'PowerPoint 갤러리: 다크 테마용 투명 흰색 결합선 슬라이드' `
            -CaptionEn 'Switch inserted slide molecules to white ink over dark backgrounds.' `
            -CaptionKo '어두운 슬라이드 배경에 맞춰 흰색 잉크 다크 모드로 일괄 전환합니다.'
        [void][Smiles2Img.Editor.SafeClipboard]::SetText($aspirinSmiles)
        [void]$pptObj.ExecuteRibbonAction('OnQuickInsert')
        $picDark = $sDark.Shapes.Item($sDark.Shapes.Count)
        $picDark.Select()
        Start-Sleep -Milliseconds 400
        $c.Snap(5, $ptBatch.X, $ptBatch.Y)
        $c.ClickAt($ptBatch.X, $ptBatch.Y, 'Click [Batch Style -> Dark Mode] to convert bonds to crisp white ink.', '[Batch Style -> Dark Mode]를 눌러 흰색 결합선으로 전환합니다.')
        [void]$pptObj.ExecuteRibbonAction('OnBatchStyleDark')
        Start-Sleep -Milliseconds 900
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 3.3 powerpoint-slide-reaction ---
        $sRxn = New-DarkSlide $pres 'Synthetic Mechanism & Reaction Pathway Slide' 'Lossless KET coordinates preserved inside PowerPoint Shape.Tags'
        $sRxn.Select()
        $cardRxn = $sRxn.Shapes.AddShape(1, 60, 110, 840, 380)
        $cardRxn.Fill.Solid()
        $cardRxn.Fill.ForeColor.RGB = 0xFFFFFF
        $c = Start-FeatureClip -Id 'powerpoint-slide-reaction' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: Multi-Step Reaction Pathway Slide' `
            -TitleKo 'PowerPoint 갤러리: 유기 합성 반응 메커니즘 슬라이드' `
            -CaptionEn 'Sketch reaction arrows and catalysts in Ketcher and place cleanly onto slides.' `
            -CaptionKo 'Ketcher에서 반응 화살표와 촉매를 그려 슬라이드에 배치하고 KET 데이터를 보존합니다.'
        $c.Snap(4, $ptInsert.X, $ptInsert.Y)
        $c.ClickAt($ptInsert.X, $ptInsert.Y, 'Open Ketcher from PowerPoint to sketch the Aspirin synthesis reaction.', 'PowerPoint에서 Ketcher를 열어 살리실산 -> 아스피린 반응식을 작성합니다.')
        $c.RunWithModal('Ketcher', 2000, {
            Show-LiveKetcherOverHost -OwnerHwnd $pHwnd -Context 'PowerPoint' -InitialInput $sampleReactionKet -HoldMs 2000
        }, 'Apply reaction scheme onto the active slide without overflowing bounds.', 'Apply를 눌러 카드 안에 반응식을 깔끔하게 배치합니다.')
        [void][Smiles2Img.Editor.SafeClipboard]::SetText("KET:$sampleReactionKet")
        [void]$pptObj.ExecuteRibbonAction('OnQuickInsert')
        $rxnPic = $sRxn.Shapes.Item($sRxn.Shapes.Count)
        $rxnPic.LockAspectRatio = -1
        $rxnPic.Width = 700
        if ($rxnPic.Height -gt 310) { $rxnPic.Height = 310 }
        $rxnPic.Left = 60 + ((840 - $rxnPic.Width) / 2)
        $rxnPic.Top = 115 + ((370 - $rxnPic.Height) / 2)
        try { $ppt.ActiveWindow.Selection.Unselect() } catch {}
        Start-Sleep -Milliseconds 600
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 3.4 powerpoint-slide-molstar-3d ---
        $crn1PptBytes = Get-Live1CrnTransparentPng -OwnerHwnd $pHwnd -Context 'PowerPoint Slide'
        $s3d = New-DarkSlide $pres '3D Macromolecular Target (PDB: 1CRN Crambin)' 'Interactive 360° rotatable GLB & high-DPI transparent 2D figures via Mol*'
        $s3d.Select()
        $c = Start-FeatureClip -Id 'powerpoint-slide-molstar-3d' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: Interactive 3D Protein Slide (Mol*)' `
            -TitleKo 'PowerPoint 갤러리: 발표 중 360° 회전하는 3D 단백질 슬라이드' `
            -CaptionEn 'Launch [3D Structure (Mol*)], fetch PDB 1CRN, and place a transparent figure onto your slide.' `
            -CaptionKo '[3D Structure (Mol*)]로 1CRN 단백질을 불러와 슬라이드에 투명하게 삽입합니다.'
        $c.Snap(4, $ptMolstar.X, $ptMolstar.Y)
        $c.ClickAt($ptMolstar.X, $ptMolstar.Y, 'Click [3D Structure (Mol*)] on the PowerPoint ribbon.', '리본의 [3D Structure (Mol*)]를 클릭합니다.')
        $c.RunWithModal('Mol*', 2200, {
            [void](Show-LiveMolstarOverHost -OwnerHwnd $pHwnd -Context 'PowerPoint Slide' -PdbId '1CRN' -HoldMs 2200 -AutoApply $false)
        }, 'Apply 3D Model (GLB) or transparent 2D figure straight onto the dark slide.', '3D 모델(GLB) 또는 투명 2D 이미지를 다크 슬라이드 위에 바로 삽입합니다.')
        if ($crn1PptBytes -and $crn1PptBytes.Length -gt 500) {
            $tmpMolP = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "p_mol_$([Guid]::NewGuid().ToString('N')).png")
            [System.IO.File]::WriteAllBytes($tmpMolP, $crn1PptBytes)
            $shp3d = $s3d.Shapes.AddPicture($tmpMolP, 0, -1, 220, 105, 520, 390)
            Remove-Item -LiteralPath $tmpMolP -Force -ErrorAction SilentlyContinue
        }
        try { $ppt.ActiveWindow.Selection.Unselect() } catch {}
        Start-Sleep -Milliseconds 600
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 3.5 powerpoint-slide-textbox-quick-insert ---
        $sTb = New-DarkSlide $pres 'Convert Slide TextBox SMILES In-Place' 'Select any TextBox containing SMILES or a chemical name and click Paste Structure'
        $sTb.Select()
        $tbDemo = $sTb.Shapes.AddTextbox(1, 220, 180, 520, 70)
        $tbDemo.TextFrame.TextRange.Text = 'CC(=O)Oc1ccccc1C(=O)O'
        $tbDemo.TextFrame.TextRange.Font.Size = 24
        $tbDemo.TextFrame.TextRange.Font.Color.RGB = 0xF8BD38
        $tbDemo.Select()
        $c = Start-FeatureClip -Id 'powerpoint-slide-textbox-quick-insert' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: TextBox SMILES -> Structure In-Place' `
            -TitleKo 'PowerPoint 갤러리: 슬라이드 텍스트 상자 SMILES 제자리 변환' `
            -CaptionEn 'Select a TextBox containing SMILES or "Aspirin" and click [Paste Structure].' `
            -CaptionKo 'SMILES가 적힌 텍스트 상자를 선택하고 [Paste Structure]를 누르면 그 자리에서 도형으로 변환됩니다.'
        $c.Snap(6, $ptPaste.X, $ptPaste.Y)
        $c.ClickAt($ptPaste.X, $ptPaste.Y, 'Click [Paste Structure] while the SMILES TextBox is selected.', 'SMILES 텍스트 상자가 선택된 상태에서 [Paste Structure]를 클릭합니다.')
        [void]$pptObj.ExecuteRibbonAction('OnQuickInsert')
        try { $tbDemo.Delete() } catch {}
        $picTb = $sTb.Shapes.Item($sTb.Shapes.Count)
        $picTb.Select()
        [void]$pptObj.ExecuteRibbonAction('OnBatchStyleDark')
        Start-Sleep -Milliseconds 800
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 3.6 powerpoint-slide-live-qa-info ---
        $sGrid.Select()
        $sGrid.Shapes.Item($sGrid.Shapes.Count).Select()
        $c = Start-FeatureClip -Id 'powerpoint-slide-live-qa-info' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: Check MW, LogP & TPSA During Live Q&A' `
            -TitleKo 'PowerPoint 갤러리: 발표 질의응답 중 분자량·LogP·TPSA 즉시 확인' `
            -CaptionEn 'Select any slide molecule and click [Molecule Info] to inspect all descriptors live.' `
            -CaptionKo '발표 도중 구조식 도형을 선택하고 [Molecule Info]를 눌러 분자량·물성을 즉시 확인합니다.'
        $c.Snap(4, $ptMolInfo.X, $ptMolInfo.Y)
        $c.ClickAt($ptMolInfo.X, $ptMolInfo.Y, 'Click [Molecule Info] for the selected slide molecule.', '선택한 도형에 대해 [Molecule Info]를 클릭합니다.')
        $c.RunWithModal('Molecule Info', 2200, {
            Show-LiveMoleculeInfoOverHost -OwnerHwnd $pHwnd -SmilesOrKet $paracetamolSmiles -HoldMs 2200
        }, 'Inspect Molecular Weight, Exact Mass, LogP, and copy multi-line slide bullet summaries.', '분자량·정밀질량·LogP를 즉시 확인하고 슬라이드 불릿 목록으로 복사할 수 있습니다.')
        $c.Snap(5, -1, -1)
        Complete-FeatureClip $c

        # --- 3.7 powerpoint-slide-batch-style ---
        $sBatch = New-DarkSlide $pres 'Slide Deck Batch Style Switcher' 'Toggle between Standard Color, Publication B&W, White Background & Dark Mode'
        $sBatch.Select()
        [void][Smiles2Img.Editor.SafeClipboard]::SetText($ibuprofenSmiles)
        [void]$pptObj.ExecuteRibbonAction('OnQuickInsert')
        $picBatch = $sBatch.Shapes.Item($sBatch.Shapes.Count)
        $picBatch.Select()
        $c = Start-FeatureClip -Id 'powerpoint-slide-batch-style' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Gallery: Switch Deck Theme with Batch Style' `
            -TitleKo 'PowerPoint 갤러리: Batch Style로 슬라이드 테마 일괄 전환' `
            -CaptionEn 'Cycle a slide deck between White Background, Publication B&W, and Dark Mode.' `
            -CaptionKo '흰색 배경, 논문용 흑백, 다크 모드로 슬라이드 전체를 일괄 전환합니다.'
        $c.Snap(4, $ptBatch.X, $ptBatch.Y)
        $c.ClickAt($ptBatch.X, $ptBatch.Y, 'Apply White Background card style for mixed slide themes.', '먼저 White Background 스타일을 적용합니다.')
        [void]$pptObj.ExecuteRibbonAction('OnBatchStyleWhiteBg')
        $sBatch.Shapes.Item($sBatch.Shapes.Count).Select()
        Start-Sleep -Milliseconds 600
        $c.Snap(6, $ptBatch.X, $ptBatch.Y)
        $c.ClickAt($ptBatch.X, $ptBatch.Y, 'Switch to Dark Mode (White Ink) for dark conference slides.', '이어서 다크 테마 슬라이드용 Dark Mode (White Ink)로 전환합니다.')
        [void]$pptObj.ExecuteRibbonAction('OnBatchStyleDark')
        Start-Sleep -Milliseconds 900
        $c.CaptureShot(@())
        $c.Snap(8, -1, -1)
        Complete-FeatureClip $c

        # --- 3.8 powerpoint-overview & powerpoint-showcase (captured AFTER all 6 slides exist!) ---
        $sGrid.Select()
        try { $ppt.ActiveWindow.Selection.Unselect() } catch {}
        Reset-PowerPointSlideZoom
        Start-Sleep -Milliseconds 400
        $cOv = Start-FeatureClip -Id 'powerpoint-overview' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'PowerPoint Overview: Dark Mode Slides, Reactions & 3D Proteins' `
            -TitleKo 'PowerPoint 개요: 다크 모드 슬라이드, 반응식 & 3D 단백질 발표' `
            -CaptionEn 'Browse the complete 6-slide presentation deck with dark-mode grids, reactions, and 3D proteins.' `
            -CaptionKo '다크 모드 그리드, 합성 반응식, 3D 단백질이 모두 담긴 6장 PowerPoint 발표 자료 전체 화면입니다.'
        $cOv.CaptureShot(@(
            (Join-Path $publicDir 'powerpoint-showcase.png'),
            (Join-Path $assetsDir 'powerpoint-showcase.png'),
            (Join-Path $publicDir 'powerpoint-ribbon-guide.png'),
            (Join-Path $assetsDir 'powerpoint-ribbon-guide.png')
        ))
        $cOv.MoveTo($ptInsert.X, $ptInsert.Y, 5)
        $cOv.AddCue(1.5, 'Preview Slide 1 (Grid), Slide 3 (Reaction KET) & Slide 4 (Mol* 3D Protein).', '후보 그리드, 반응식 슬라이드, Mol* 3D 단백질 슬라이드를 순차 확인합니다.')
        $sRxn.Select(); Start-Sleep -Milliseconds 300; $cOv.Snap(6, -1, -1)
        $s3d.Select(); Start-Sleep -Milliseconds 400; $cOv.Snap(6, -1, -1)
        $sGrid.Select(); Start-Sleep -Milliseconds 300; $cOv.Snap(4, $ptBatch.X, $ptBatch.Y)
        Complete-FeatureClip $cOv
        }

        if ($Section -eq 'All' -or $Section -eq 'PowerPoint' -or $Section -eq 'UseCases') {
        # --- 3.9 powerpoint-slide-svg-animation ---
        $sSvg = $pres.Slides.Add($pres.Slides.Count + 1, 12)
        $sSvg.FollowMasterBackground = 0
        $sSvg.Background.Fill.Solid()
        $sSvg.Background.Fill.ForeColor.RGB = 0xFAF6F4 # Soft off-white #F4F6FA
        $tbSvgTitle = $sSvg.Shapes.AddTextbox(1, 36, 20, 888, 38)
        $tbSvgTitle.TextFrame.TextRange.Text = 'Step-by-Step SVG Vector Editing & Stage Animation'
        $tbSvgTitle.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $tbSvgTitle.TextFrame.TextRange.Font.Size = 22
        $tbSvgTitle.TextFrame.TextRange.Font.Bold = -1
        $tbSvgTitle.TextFrame.TextRange.Font.Color.RGB = 0x2A170F
        $tbSvgSub = $sSvg.Shapes.AddTextbox(1, 36, 58, 888, 22)
        $tbSvgSub.TextFrame.TextRange.Text = 'Copy SVG -> Paste onto Slide -> Convert to Shape -> Recolor Key Product & Arrow Paths'
        $tbSvgSub.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $tbSvgSub.TextFrame.TextRange.Font.Size = 11.5
        $tbSvgSub.TextFrame.TextRange.Font.Color.RGB = 0x695547

        # Top-left original KET reference box
        $boxOrig = $sSvg.Shapes.AddShape(1, 36, 94, 888, 152)
        $boxOrig.Fill.Solid()
        $boxOrig.Fill.ForeColor.RGB = 0xFFFFFF
        $boxOrig.Line.ForeColor.RGB = 0xE2D5CB
        $lblOrig = $sSvg.Shapes.AddTextbox(1, 50, 100, 360, 22)
        $lblOrig.TextFrame.TextRange.Text = '1. Original Editable KET Scheme (Source)'
        $lblOrig.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $lblOrig.TextFrame.TextRange.Font.Size = 10.5
        $lblOrig.TextFrame.TextRange.Font.Bold = -1
        $lblOrig.TextFrame.TextRange.Font.Color.RGB = 0x8A5A2B

        # Bottom editable vector stage box
        $boxVec = $sSvg.Shapes.AddShape(1, 36, 260, 888, 250)
        $boxVec.Fill.Solid()
        $boxVec.Fill.ForeColor.RGB = 0xFFFFFF
        $boxVec.Line.ForeColor.RGB = 0xC8A882
        $lblVec = $sSvg.Shapes.AddTextbox(1, 50, 268, 520, 22)
        $lblVec.TextFrame.TextRange.Text = '2. Pasted SVG -> Convert to Native PowerPoint Shapes (Highlight & Click Build)'
        $lblVec.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $lblVec.TextFrame.TextRange.Font.Size = 10.5
        $lblVec.TextFrame.TextRange.Font.Bold = -1
        $lblVec.TextFrame.TextRange.Font.Color.RGB = 0x597E0D

        $sSvg.Select()
        Reset-PowerPointSlideZoom
        Start-Sleep -Milliseconds 400

        [void][Smiles2Img.Editor.SafeClipboard]::SetText($sampleReactionKet)
        [void]$pptObj.ExecuteRibbonAction('OnQuickInsert')
        $picOrig = $sSvg.Shapes.Item($sSvg.Shapes.Count)
        $picOrig.LockAspectRatio = -1
        $picOrig.Height = 106
        if ($picOrig.Width -gt 780) { $picOrig.Width = 780 }
        $picOrig.Left = 480 - ($picOrig.Width / 2)
        $picOrig.Top = 126
        $picOrig.Select()
        Start-Sleep -Milliseconds 400

        $c = Start-FeatureClip -Id 'powerpoint-slide-svg-animation' -HostApp 'powerpoint' -MainHwnd $pHwnd -OfficePid $pPid `
            -TitleEn 'Use Case: Copy SVG -> Convert to Editable PowerPoint Shapes' `
            -TitleKo '실전 시나리오: Copy SVG -> 파워포인트 네이티브 도형 변환 및 단계 강조' `
            -CaptionEn 'Copy lossless SVG from a reaction scheme, paste onto a slide, convert to native vector shapes, and recolor key regions.' `
            -CaptionKo '반응식에서 [Copy SVG]로 벡터를 복사해 슬라이드에 붙여넣은 뒤 네이티브 도형으로 변환해 핵심 결합·생성물을 강조합니다.'
        $c.Snap(6, $ptCopySvgP.X, $ptCopySvgP.Y)
        $c.ClickAt($ptCopySvgP.X, $ptCopySvgP.Y, 'Select the KET scheme and click [Copy SVG] on the [SMILES] ribbon.', '반응식을 선택하고 [SMILES] 리본의 [Copy SVG]를 클릭합니다.')
        [void]$pptObj.ExecuteRibbonAction('OnCopySvg')
        Start-Sleep -Milliseconds 400

        # Render the exact SVG XML from MoleculeRenderer and place it as a native Office SVG graphic
        $svgXml = [Smiles2Img.Rendering.MoleculeRenderer]::RenderSvg($sampleReactionKet, [Smiles2Img.Rendering.RenderOptions]::Default)
        $tmpSvgPath = Join-Path ([System.IO.Path]::GetTempPath()) ("s2i_slide_svg_" + [Guid]::NewGuid().ToString('N') + ".svg")
        [System.IO.File]::WriteAllText($tmpSvgPath, $svgXml, [System.Text.UTF8Encoding]::new($false))
        $svgPic = $sSvg.Shapes.AddPicture($tmpSvgPath, 0, -1, 120, 304)
        $svgPic.LockAspectRatio = -1
        $svgPic.Height = 148
        if ($svgPic.Width -gt 800) { $svgPic.Width = 800 }
        $svgPic.Left = 480 - ($svgPic.Width / 2)
        $svgPic.Top = 304
        $svgLeft = $svgPic.Left
        $svgWidth = $svgPic.Width
        $svgTop = $svgPic.Top
        $svgHeight = $svgPic.Height
        $svgPic.Select()
        Start-Sleep -Milliseconds 500
        $c.AddCue(1.8, 'Paste the scalable SVG graphic onto the slide and run Convert to Shape.', '슬라이드에 무손실 SVG 그래픽을 붙여넣고 도형으로 변환(Convert to Shape)을 실행합니다.')
        $c.Snap(6, (Scale-RefX $pHwnd 1380), (Scale-RefY $pHwnd 1250))

        # Convert native SVG graphic into individual editable PowerPoint Freeform vector shapes via SVGEdit
        $firstVecIdx = $sSvg.Shapes.Count
        try {
            $ppt.CommandBars.ExecuteMso('SVGEdit')
            Start-Sleep -Milliseconds 1200
        } catch {
            Write-Warning "SVGEdit failed: $($_.Exception.Message)"
        }

        # SVGEdit splits the SVG Graphic directly into native Freeform shapes starting at $firstVecIdx
        $recoloredCount = 0
        $prodNames = New-Object 'System.Collections.Generic.List[string]'
        $arrowMinX = $svgLeft + ($svgWidth * 0.598)
        $prodMinX  = $svgLeft + ($svgWidth * 0.66)
        $vecTotal = $sSvg.Shapes.Count - $firstVecIdx + 1
        for ($si = $firstVecIdx; $si -le $sSvg.Shapes.Count; $si++) {
            $sh = $sSvg.Shapes.Item($si)
            $centerX = $sh.Left + ($sh.Width / 2.0)
            if ($centerX -ge $prodMinX) {
                try {
                    if ($sh.Fill.Visible -ne 0) { $sh.Fill.ForeColor.RGB = 0x3D3DE1 } # Vivid crimson/coral (BGR #E13D3D)
                    if ($sh.Line.Visible -ne 0) { $sh.Line.ForeColor.RGB = 0x3D3DE1 }
                    $prodNames.Add($sh.Name)
                    $recoloredCount++
                } catch {}
            } elseif ($centerX -ge $arrowMinX) {
                try {
                    if ($sh.Fill.Visible -ne 0) { $sh.Fill.ForeColor.RGB = 0xD8631D } # Cobalt blue arrow (BGR #1D63D8)
                    if ($sh.Line.Visible -ne 0) { $sh.Line.ForeColor.RGB = 0xD8631D }
                    $recoloredCount++
                } catch {}
            }
        }
        Write-Host "    [SVG Convert] Split into $vecTotal native PowerPoint Freeform vector shapes; recolored $recoloredCount shapes." -ForegroundColor DarkCyan

        # Add clean stage footer pills beneath the 3 reaction zones
        $stages = @(
            @{ Text = 'Stage 1: Reactants (Click 1)';      X = 90;  W = 240; Bg = 0xF0E6DC; Fg = 0x473626 },
            @{ Text = 'Stage 2: Arrow (Wipe)';             X = 390; W = 175; Bg = 0xFBE9DA; Fg = 0xB84E10 },
            @{ Text = 'Stage 3: Highlighted Product (Click 3)'; X = 595; W = 290; Bg = 0xE2E2FE; Fg = 0x2626B9 }
        )
        foreach ($st in $stages) {
            $pill = $sSvg.Shapes.AddShape(5, $st.X, 468, $st.W, 28) # msoShapeRoundedRectangle
            $pill.Fill.Solid()
            $pill.Fill.ForeColor.RGB = $st.Bg
            $pill.Line.Visible = 0
            $pill.TextFrame.TextRange.Text = $st.Text
            $pill.TextFrame.TextRange.Font.Name = 'Segoe UI'
            $pill.TextFrame.TextRange.Font.Size = 10.5
            $pill.TextFrame.TextRange.Font.Bold = -1
            $pill.TextFrame.TextRange.Font.Color.RGB = $st.Fg
            $pill.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignCenter
        }
        try {
            if ($prodNames.Count -gt 1) {
                $prodGrp = $sSvg.Shapes.Range(@($prodNames)).Group()
                $prodGrp.Select()
            } else {
                $ppt.ActiveWindow.Selection.Unselect()
            }
        } catch {}
        Remove-Item -LiteralPath $tmpSvgPath -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 700
        $c.AddCue(2.4, 'Converted SVG paths are now native PowerPoint vector shapes—recolored and grouped by stage for animation.', '변환된 SVG 경로가 네이티브 파워포인트 도형으로 분리되어 단계별 색상 강조와 그룹 애니메이션이 가능합니다.')
        $c.CaptureShot(@())
        $c.Snap(14, -1, -1)
        Complete-FeatureClip $c
        }

        $pres.Saved = -1
        $pres.Close()
    } finally {
        try { $ppt.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
        Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# =============================================================================
# 4. COMMON EDITORS & DIALOGS OVER LIVE OFFICE (4 Features + Core Screenshots)
# =============================================================================
if ($Section -eq 'All' -or $Section -eq 'Common') {
    Write-Host "`n=== [4/4] Capturing Common Editors & Dialogs Over Live Office ===" -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    try {
        $doc = $word.Documents.Add()
        $pIntro = $doc.Paragraphs.Add()
        $pIntro.Range.Text = "func_SMILES2IMG Shared Editors & Dialogs (Excel · Word · PowerPoint)"
        $pIntro.Range.Font.Size = 14
        $pIntro.Range.Font.Bold = 1
        $wHwnd = [IntPtr]$word.ActiveWindow.Hwnd
        $wPid = [LiveCaptureEngine]::GetPidForHwnd($wHwnd)
        try { $word.ActiveWindow.View.Zoom.Percentage = 85 } catch {}
        [void](Set-StandardHostWindow $wHwnd 'Word (Common)')
        Start-Sleep -Milliseconds 600
        [void][LiveCaptureEngine]::SelectRibbonTab($wHwnd, 'SMILES')
        Start-Sleep -Milliseconds 500
        $ptWInsert  = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Insert / Edit Structure'))
        $ptWMolInfo = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('Molecule Info'))
        $ptWAbout   = [LiveCaptureEngine]::FindControlRelativeCenter($wHwnd, @('About func_SMILES2IMG'))

        # --- 4.1 common-ketcher-2d ---
        $c = Start-FeatureClip -Id 'common-ketcher-2d' -HostApp 'common' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'EPAM Ketcher 2D Molecular & Reaction Sketcher' `
            -TitleKo 'EPAM Ketcher 2D 분자·반응식 편집기 (공통)' `
            -CaptionEn 'Sketch molecules and multi-step reactions locally offline with ribbon launch and double-click editing.' `
            -CaptionKo '리본 버튼 또는 더블클릭으로 오프라인 Ketcher 2D 편집기를 열어 분자와 반응식을 그립니다.'
        $c.Snap(8, $ptWInsert.X, $ptWInsert.Y)
        $c.ClickAt($ptWInsert.X, $ptWInsert.Y, 'Click [Insert / Edit Structure] on the [SMILES] ribbon to launch Ketcher.', '리본의 [Insert / Edit Structure] 버튼으로 Ketcher를 엽니다.')
        $c.RunWithModal('Ketcher', 2200, {
            Show-LiveKetcherOverHost -OwnerHwnd $wHwnd -Context 'Word' -InitialInput $aspirinSmiles -HoldMs 2200
        }, 'Live WebView2 Ketcher 2D editor running offline inside Office.', '오프라인 WebView2 환경에서 실행 중인 실제 Ketcher 2D 편집기입니다.', @(
            (Join-Path $publicDir 'ketcher-guide.png'),
            (Join-Path $assetsDir 'ketcher-guide.png'),
            (Join-Path $publicDir 'workflow-ketcher.png'),
            (Join-Path $assetsDir 'workflow-ketcher.png')
        ))
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        # --- 4.3 common-molecule-info ---
        $c = Start-FeatureClip -Id 'common-molecule-info' -HostApp 'common' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'Detailed Molecule Info Window & 5 Copy Formats' `
            -TitleKo '상세 분자 정보 창 (Molecule Info) & 5가지 서식 내보내기' `
            -CaptionEn 'Inspect calculated property rows with a 2D preview and copy 12-column TSV, Multi-Line, Markdown, or JSON.' `
            -CaptionKo '2D 미리보기와 함께 계산된 물성 행을 확인하고 12열 TSV·마크다운·JSON으로 즉시 복사합니다.'
        $c.Snap(8, $ptWMolInfo.X, $ptWMolInfo.Y)
        $c.ClickAt($ptWMolInfo.X, $ptWMolInfo.Y, 'Click [Molecule Info] to inspect calculated properties and export formats.', '리본의 [Molecule Info]를 클릭해 물성과 복사 서식을 확인합니다.')
        $c.RunWithModal('Molecule Info', 2200, {
            Show-LiveMoleculeInfoOverHost -OwnerHwnd $wHwnd -SmilesOrKet $aspirinSmiles -HoldMs 2200
        }, 'Live MoleculeInfoDialog with GDI+ property icons and multi-format copy export.', 'GDI+ 벡터 아이콘과 다중 서식 복사 버튼을 갖춘 실제 Molecule Info 대화상자입니다.', @(
            (Join-Path $publicDir 'molecule-info-dialog.png'),
            (Join-Path $assetsDir 'molecule-info-dialog.png')
        ))
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        # --- 4.4 common-about-dialog ---
        $c = Start-FeatureClip -Id 'common-about-dialog' -HostApp 'common' -MainHwnd $wHwnd -OfficePid $wPid `
            -TitleEn 'About func_SMILES2IMG & Update Check' `
            -TitleKo 'About func_SMILES2IMG 정보 창 & 업데이트 확인' `
            -CaptionEn "Standard 3-tab About dialog showing version v$addinVersion, third-party engines, and license notices." `
            -CaptionKo "v$addinVersion 버전, 내장 엔진 버전, 오픈소스 라이선스와 업데이트 확인을 제공하는 표준 3-탭 정보 창입니다."
        $c.Snap(8, $ptWAbout.X, $ptWAbout.Y)
        $c.ClickAt($ptWAbout.X, $ptWAbout.Y, 'Click [About func_SMILES2IMG] to view version info and check for updates.', '[About func_SMILES2IMG]를 눌러 버전 정보 및 업데이트 확인을 실행합니다.')
        $c.RunWithModal('About func_SMILES2IMG', 2200, {
            Show-LiveAboutOverHost -OwnerHwnd $wHwnd -HoldMs 2200
        }, 'Inspect Third-Party Components & License Notices or click [Check for Updates].', '외장 엔진 버전·오픈소스 라이선스 확인 및 [Check for Updates]를 바로 실행할 수 있습니다.', @(
            (Join-Path $publicDir 'about-dialog.png'),
            (Join-Path $assetsDir 'about-dialog.png')
        ))
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        $doc.Close($false)
    } finally {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
    }

    # --- 4.2 common-molstar-3d over real Microsoft PowerPoint! ---
    $pptC = New-Object -ComObject PowerPoint.Application
    try {
        $presC = $pptC.Presentations.Add(-1)
        $sC = $presC.Slides.Add(1, 12)
        $sC.FollowMasterBackground = 0
        $sC.Background.Fill.Solid()
        $sC.Background.Fill.ForeColor.RGB = 0x2A170F
        $pHwndC = [IntPtr]$pptC.HWND
        $pPidC = [LiveCaptureEngine]::GetPidForHwnd($pHwndC)
        [void](Set-StandardHostWindow $pHwndC 'PowerPoint (Common)')
        Start-Sleep -Milliseconds 600
        [void][LiveCaptureEngine]::SelectRibbonTab($pHwndC, 'SMILES')
        Start-Sleep -Milliseconds 500
        $ptPMolstar = [LiveCaptureEngine]::FindControlRelativeCenter($pHwndC, @('3D Structure (Mol*)'))

        $c = Start-FeatureClip -Id 'common-molstar-3d' -HostApp 'common' -MainHwnd $pHwndC -OfficePid $pPidC `
            -TitleEn 'Mol* 3D Macromolecular & Protein Viewer (PDB / GLB)' `
            -TitleKo 'Mol* 3D 단백질 · 고분자 뷰어 (공통)' `
            -CaptionEn 'Load PDB codes like 1CRN or 4HHB, style ribbons/surfaces, and insert rotatable GLB models or PNG figures.' `
            -CaptionKo '1CRN, 4HHB 같은 PDB 코드를 불러와 3D 리본/표면 스타일을 꾸미고 문서에 삽입합니다.'
        $c.Snap(8, $ptPMolstar.X, $ptPMolstar.Y)
        $c.ClickAt($ptPMolstar.X, $ptPMolstar.Y, 'Click [3D Structure (Mol*)] on the PowerPoint ribbon to open the 3D viewer.', 'PowerPoint 리본에서 [3D Structure (Mol*)]를 클릭해 3D 뷰어를 엽니다.')
        $c.RunWithModal('Mol*', 2200, {
            [void](Show-LiveMolstarOverHost -OwnerHwnd $pHwndC -Context 'PowerPoint Slide' -PdbId '1CRN' -HoldMs 2200 -AutoApply $false)
        }, 'Live WebGL Mol* 3D viewer displaying Crambin (1CRN) cartoon structure over PowerPoint.', 'PowerPoint 위에서 1CRN 3D 단백질을 실시간 렌더링하는 WebGL Mol* 뷰어입니다.')
        $c.Snap(10, -1, -1)
        Complete-FeatureClip $c

        $presC.Saved = -1
        $presC.Close()
    } finally {
        try { $pptC.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptC) | Out-Null
        Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# =============================================================================
# 5. FINALIZE COMPILATION VIDEOS & HERO LIVE GIF
# =============================================================================
if ($Section -eq 'All' -or $Section -eq 'Finalize') {
    Write-Host "`n=== [5/5] Building App Overview Compilation Videos & Live Hero GIF ===" -ForegroundColor Cyan

    function Join-ClipList([string[]]$ids, [string]$outMp4) {
        $listTxt = Join-Path ([System.IO.Path]::GetTempPath()) ("concat_" + [Guid]::NewGuid().ToString('N') + ".txt")
        $lines = @()
        foreach ($cid in $ids) {
            $p = Join-Path $mediaRoot "$cid\clip.mp4"
            if (Test-Path $p) {
                $lines += "file '" + $p.Replace('\', '/') + "'"
            }
        }
        if ($lines.Count -eq 0) { return }
        [System.IO.File]::WriteAllLines($listTxt, $lines, [System.Text.UTF8Encoding]::new($false))
        & $ffmpeg -hide_banner -loglevel error -y -f concat -safe 0 -i $listTxt -c copy -movflags +faststart $outMp4
        Remove-Item -LiteralPath $listTxt -Force -ErrorAction SilentlyContinue
        Write-Host "  [OK] Compiled $outMp4 ($([int]((Get-Item $outMp4).Length/1KB)) KB)" -ForegroundColor Green
    }

    Join-ClipList @('excel-overview', 'excel-formula-table', 'excel-style-matrix', 'excel-reactions-ket', 'excel-cheminformatics', 'excel-sdf-workflow') (Join-Path $assetsVideosDir 'excel-guide-video.mp4')
    Join-ClipList @('word-overview', 'word-ch1-insert-ketcher', 'word-ch2-quick-insert', 'word-ch3-compound-table', 'word-ch4-reaction-scheme', 'word-ch5-edit-and-molstar', 'word-ch6-publication-finish') (Join-Path $assetsVideosDir 'word-guide-video.mp4')
    Join-ClipList @('powerpoint-overview', 'powerpoint-slide-candidate-grid', 'powerpoint-slide-dark-mode', 'powerpoint-slide-molstar-3d', 'powerpoint-slide-textbox-quick-insert', 'powerpoint-slide-batch-style') (Join-Path $assetsVideosDir 'powerpoint-guide-video.mp4')

    # Generate live hero demo media from real Excel + Ketcher MP4 clips:
    # - VitePress docs: hero-demo.webm (VP9 crf 34) + hero-demo-poster.webp (rendered excel-overview shot)
    # - GitHub README: assets/hero-demo.webp (Animated WebP)
    $heroSrc1 = Join-Path $mediaRoot 'excel-overview\clip.mp4'
    $heroSrc2 = Join-Path $mediaRoot 'excel-reactions-ket\clip.mp4'
    $heroPosterSrc = Join-Path $mediaRoot 'excel-overview\shot.webp'
    $heroPublicWebm = Join-Path $publicDir 'hero-demo.webm'
    $heroPublicPoster = Join-Path $publicDir 'hero-demo-poster.webp'
    $heroAssetsWebp = Join-Path $assetsDir 'hero-demo.webp'
    $heroConcatTxt = Join-Path ([System.IO.Path]::GetTempPath()) ("hero_" + [Guid]::NewGuid().ToString('N') + ".txt")
    $heroClipLines = @()
    if (Test-Path $heroSrc1) { $heroClipLines += "file '$($heroSrc1.Replace('\','/'))'" }
    if (Test-Path $heroSrc2) { $heroClipLines += "file '$($heroSrc2.Replace('\','/'))'" }
    if ($heroClipLines.Count -gt 0) {
        [System.IO.File]::WriteAllLines($heroConcatTxt, $heroClipLines, [System.Text.UTF8Encoding]::new($false))
        $heroJoined = Join-Path ([System.IO.Path]::GetTempPath()) ("hero_joined_" + [Guid]::NewGuid().ToString('N') + ".mp4")
        & $ffmpeg -hide_banner -loglevel error -y -f concat -safe 0 -i $heroConcatTxt -c copy $heroJoined
        & $ffmpeg -hide_banner -loglevel error -y -i $heroJoined `
            -vf "scale=1280:854:flags=lanczos" `
            -c:v libvpx-vp9 -b:v 0 -crf 34 -row-mt 1 -an `
            $heroPublicWebm
        if (Test-Path $heroPosterSrc) {
            & $ffmpeg -hide_banner -loglevel error -y -i $heroPosterSrc `
                -vf "scale=1280:854:flags=lanczos" `
                -c:v libwebp -quality 84 `
                $heroPublicPoster
        } else {
            & $ffmpeg -hide_banner -loglevel error -y -i $heroJoined `
                -frames:v 1 -vf "scale=1280:854:flags=lanczos" `
                -c:v libwebp -quality 84 `
                $heroPublicPoster
        }
        & $ffmpeg -hide_banner -loglevel error -y -i $heroJoined `
            -vf "scale=1280:854:flags=lanczos" `
            -c:v libwebp_anim -lossless 0 -compression_level 6 -q:v 80 -preset default -loop 0 `
            $heroAssetsWebp
        Remove-Item -LiteralPath $heroConcatTxt, $heroJoined -Force -ErrorAction SilentlyContinue
        Write-Host "  [OK] Generated live hero-demo.webm ($([int]((Get-Item $heroPublicWebm).Length/1KB)) KB), hero-demo-poster.webp ($([int]((Get-Item $heroPublicPoster).Length/1KB)) KB), assets/hero-demo.webp ($([int]((Get-Item $heroAssetsWebp).Length/1KB)) KB)" -ForegroundColor Green
    }
}

Save-ManifestFiles
Write-Host "`n[ALL LIVE MEDIA CAPTURE COMPLETED SUCCESSFULLY]" -ForegroundColor Green

