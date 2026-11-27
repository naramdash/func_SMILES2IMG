<#
.SYNOPSIS
    Generates lightweight 3:2 (1380x920 WebP) carousel preview images directly from the
    3 official self-contained Playbook files (Excel .xlsx, Word .docx, PowerPoint .pptx)
    into docs/public/examples/previews/ for the Homepage Playbook Preview Carousel.
#>
param(
    [ValidateSet('All', 'Excel', 'Word', 'PowerPoint')]
    [string]$App = 'All'
)

$ErrorActionPreference = 'Stop'
$PreviewLogicalWidth  = 1230
$PreviewLogicalHeight = 820 # Exact 3:2 aspect ratio (1230x820 DIPs = 2766x1844 px @ 225% DPI)
Add-Type -AssemblyName System.Drawing, System.Windows.Forms, UIAutomationClient, UIAutomationTypes, WindowsBase

$repoRoot    = Split-Path -Parent $PSScriptRoot
$examplesDir = Join-Path $repoRoot 'examples'
$outDir      = Join-Path $repoRoot 'docs\public\examples\previews'
$tmpDir      = Join-Path ([System.IO.Path]::GetTempPath()) ("s2i_pb_prev_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $outDir, $tmpDir -Force | Out-Null

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + $env:Path
$ffmpeg = (Get-Command ffmpeg -ErrorAction Stop).Source

# Ensure English Office UI
$langScript = Join-Path $PSScriptRoot 'Install-OfficeEnglishPack.ps1'
if (Test-Path $langScript) { & $langScript -RegistryOnly | Out-Null }

# Load LiveCaptureEngine for crisp HiDPI Excel & Word live capture
$engineCs = Get-Content (Join-Path $PSScriptRoot 'LiveCaptureEngine.cs') -Raw
Add-Type -TypeDefinition $engineCs -ReferencedAssemblies @(
    'System.Drawing', 'System.Windows.Forms', 'UIAutomationClient', 'UIAutomationTypes', 'WindowsBase'
) -ErrorAction Stop
[LiveCaptureEngine]::InitDpi()

function Convert-ToWebP([string]$srcPng, [string]$dstWebp) {
    & $ffmpeg -hide_banner -loglevel error -y -i $srcPng -vf 'scale=1380:920:flags=lanczos' -c:v libwebp -quality 86 $dstWebp
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg WebP conversion failed: $srcPng -> $dstWebp" }
    $kb = [math]::Round((Get-Item $dstWebp).Length / 1KB, 1)
    Write-Host ("  [OK] {0,-18} ({1,5:N1} KB)" -f (Split-Path $dstWebp -Leaf), $kb) -ForegroundColor Green
}

try {
    # =========================================================================
    # 1. EXCEL PLAYBOOK PREVIEWS (10 tabs captured live in 3:2 aspect ratio)
    # =========================================================================
    if ($App -eq 'All' -or $App -eq 'Excel') {
        Write-Host "`n[1/3] Capturing Excel Playbook Tab Previews (3:2)..." -ForegroundColor Cyan
        $xlsxPath = (Resolve-Path (Join-Path $examplesDir 'func_SMILES2IMG_Excel_Playbook.xlsx')).Path
        $xllPath  = Join-Path $repoRoot 'dist\x64\Smiles2Img-AddIn64-packed.xll'
        if (-not (Test-Path $xllPath)) { $xllPath = Join-Path $env:APPDATA 'Microsoft\AddIns\Smiles2Img\Smiles2Img-AddIn64.xll' }

        Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $wb = $null
        try {
            if (Test-Path $xllPath) { [void]$excel.RegisterXLL((Resolve-Path $xllPath).Path) }
            $wb = $excel.Workbooks.Open($xlsxPath, 0, $true)
            $xHwnd = [IntPtr]$excel.Hwnd
            $xPid  = [LiveCaptureEngine]::GetPidForHwnd($xHwnd)
            [void][LiveCaptureEngine]::SetMainWindowSizeLogical($xHwnd, 4, 4, $PreviewLogicalWidth, $PreviewLogicalHeight)
            Start-Sleep -Milliseconds 900
            [void][LiveCaptureEngine]::SelectRibbonTab($xHwnd, 'SMILES')
            Start-Sleep -Milliseconds 800

            $excelTabs = @(
                @{ Name = '0. Start Here';            Zoom = 76; Cell = 'E6'; Out = 'excel-01.webp' },
                @{ Name = '1. Structure Rendering';   Zoom = 74; Cell = 'E4'; Out = 'excel-02.webp' },
                @{ Name = '2. Properties & 3D';       Zoom = 72; Cell = 'D4'; Out = 'excel-03.webp' },
                @{ Name = '3. SAR & SMARTS';          Zoom = 78; Cell = 'E5'; Out = 'excel-04.webp' },
                @{ Name = '4. Notation Conversions';  Zoom = 70; Cell = 'F4'; Out = 'excel-05.webp' },
                @{ Name = '5. Styles & Transforms';   Zoom = 78; Cell = 'F5'; Out = 'excel-06.webp' },
                @{ Name = '6. Reactions & KET';       Zoom = 78; Cell = 'E4'; Out = 'excel-07.webp' },
                @{ Name = '7. Real-World Workflows';  Zoom = 82; Cell = 'D6'; Out = 'excel-08.webp' },
                @{ Name = '8. Legacy Excel (.FLOAT)'; Zoom = 74; Cell = 'G4'; Out = 'excel-09.webp' },
                @{ Name = '9. Ribbon, SDF & Ref';     Zoom = 80; Cell = 'B6'; Out = 'excel-10.webp' }
            )
            foreach ($tab in $excelTabs) {
                $ws = $wb.Worksheets.Item($tab.Name)
                $ws.Activate()
                $excel.ActiveWindow.Zoom = $tab.Zoom
                $excel.ActiveWindow.ScrollRow = 1
                $excel.ActiveWindow.ScrollColumn = 1
                $ws.Range($tab.Cell).Select() | Out-Null
                Start-Sleep -Milliseconds 1050
                $rawPng = Join-Path $tmpDir ($tab.Out -replace '\.webp$', '.png')
                [LiveCaptureEngine]::CaptureCompositeFrame($xHwnd, $xPid, [uint32]$PID, -1, -1, 0, $rawPng)
                Convert-ToWebP $rawPng (Join-Path $outDir $tab.Out)
            }
        } finally {
            if ($wb) { try { $wb.Close($false) } catch {} }
            if ($excel) { try { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
        }
    }

    # =========================================================================
    # 2. WORD PLAYBOOK PREVIEWS (6 pages captured live in 3:2 single-page view)
    # =========================================================================
    if ($App -eq 'All' -or $App -eq 'Word') {
        Write-Host "`n[2/3] Capturing Word Playbook Live Page Previews (3:2)..." -ForegroundColor Cyan
        $docxPath    = (Resolve-Path (Join-Path $examplesDir 'func_SMILES2IMG_Word_Playbook.docx')).Path
        $tmpDocxPath = Join-Path $tmpDir 'func_SMILES2IMG_Word_Playbook.docx'
        Copy-Item -LiteralPath $docxPath -Destination $tmpDocxPath -Force

        Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $word.DisplayAlerts = 0
        $doc = $null
        try {
            $word.WindowState = 0 # wdWindowStateNormal
            $doc = $word.Documents.Open($tmpDocxPath, $false, $false, $false)
            $wHwnd = [IntPtr]$word.ActiveWindow.Hwnd
            $wPid  = [LiveCaptureEngine]::GetPidForHwnd($wHwnd)
            [void][LiveCaptureEngine]::SetMainWindowSizeLogical($wHwnd, 4, 4, $PreviewLogicalWidth, $PreviewLogicalHeight)
            Start-Sleep -Milliseconds 700

            $doc.ShowSpellingErrors = $false
            $doc.ShowGrammaticalErrors = $false

            $view = $word.ActiveWindow.View
            $view.Type = 3 # wdPrintView
            $view.ShowAll = $false
            $view.DisplayPageBoundaries = $true
            $view.ShowRevisionsAndComments = $false
            $word.ActiveWindow.DisplayRulers = $false
            $word.ActiveWindow.DocumentMap = $false
            $view.Zoom.PageColumns = 1
            $view.Zoom.Percentage = 86
            Start-Sleep -Milliseconds 400
            [void][LiveCaptureEngine]::SelectRibbonTab($wHwnd, 'SMILES')
            Start-Sleep -Milliseconds 700

            # Capture each of the 6 pages cleanly at scroll=0 (avoids Word PrintWindow scroll cache bug)
            $pageOuts = @('word-01.webp', 'word-02.webp', 'word-03.webp', 'word-04.webp', 'word-05.webp', 'word-06.webp')
            for ($pi = 0; $pi -lt $pageOuts.Count; $pi++) {
                if ($pi -gt 0) {
                    # Delete the current first page of our temp copy so the next page sits at top of page 1
                    $r1 = $word.Selection.GoTo(1, 1, 1)
                    $r2 = $word.Selection.GoTo(1, 1, 2)
                    if ($r2.Start -gt $r1.Start) {
                        $delRng = $doc.Range($r1.Start, $r2.Start)
                        $delRng.Delete() | Out-Null
                        while ($doc.Paragraphs.Count -gt 1 -and ($doc.Paragraphs.Item(1).Range.Text -replace '[\r\n\f\s]', '') -eq '' -and $doc.Paragraphs.Item(1).Range.InlineShapes.Count -eq 0 -and $doc.Paragraphs.Item(1).Range.Tables.Count -eq 0) {
                            $doc.Paragraphs.Item(1).Range.Delete() | Out-Null
                        }
                    }
                    $view.Zoom.PageColumns = 1
                    $view.Zoom.Percentage = 86
                    $word.ActiveWindow.VerticalPercentScrolled = 0
                    $word.ScreenRefresh()
                    Start-Sleep -Milliseconds 850
                }
                # Park selection on a non-table paragraph on page 1 so no table move-handle appears
                $parked = $false
                foreach ($para in $doc.Paragraphs) {
                    if ($para.Range.Tables.Count -eq 0 -and ($para.Range.Text -replace '[\r\n\f\s]', '') -ne '') {
                        $para.Range.Select() | Out-Null
                        $word.Selection.Collapse(1)
                        $parked = $true
                        break
                    }
                }
                if (-not $parked) { $doc.Range(0, 0).Select() | Out-Null }
                $word.ActiveWindow.VerticalPercentScrolled = 0
                $word.ScreenRefresh()
                Start-Sleep -Milliseconds 600
                $rawPng = Join-Path $tmpDir ($pageOuts[$pi] -replace '\.webp$', '.png')
                [LiveCaptureEngine]::CaptureCompositeFrame($wHwnd, $wPid, [uint32]$PID, -1, -1, 0, $rawPng)
                Convert-ToWebP $rawPng (Join-Path $outDir $pageOuts[$pi])
            }
        } finally {
            if ($doc) { try { $doc.Close($false) } catch {} }
            if ($word) { try { $word.Quit(0); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {} }
            Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }

    # =========================================================================
    # 3. POWERPOINT PLAYBOOK PREVIEWS (all 8 slides captured live in 3:2 PPT UI)
    # =========================================================================
    if ($App -eq 'All' -or $App -eq 'PowerPoint') {
        Write-Host "`n[3/3] Capturing PowerPoint Playbook Live Editor Previews (8 Slides, 3:2)..." -ForegroundColor Cyan
        $pptxPath = (Resolve-Path (Join-Path $examplesDir 'func_SMILES2IMG_PowerPoint_Playbook.pptx')).Path
        Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
        $ppt = New-Object -ComObject PowerPoint.Application
        $pres = $null
        try {
            $pres = $ppt.Presentations.Open($pptxPath, -1, 0, -1) # ReadOnly=True, Untitled=False, WithWindow=True
            $pHwnd = [IntPtr]$ppt.HWND
            $pPid  = [LiveCaptureEngine]::GetPidForHwnd($pHwnd)
            try { $ppt.ActiveWindow.WindowState = 1 } catch {} # ppWindowNormal
            [void][LiveCaptureEngine]::SetMainWindowSizeLogical($pHwnd, 4, 4, $PreviewLogicalWidth, $PreviewLogicalHeight)
            Start-Sleep -Milliseconds 900
            [void][LiveCaptureEngine]::SelectRibbonTab($pHwnd, 'SMILES')
            Start-Sleep -Milliseconds 800

            try {
                $ppt.ActiveWindow.ViewType = 9 # ppViewNormal
                $ppt.ActiveWindow.SplitHorizontal = 15
                $ppt.ActiveWindow.SplitVertical = 88
            } catch {}

            for ($idx = 1; $idx -le $pres.Slides.Count; $idx++) {
                $pres.Slides.Item($idx).Select()
                try { $ppt.ActiveWindow.Selection.Unselect() } catch {}
                # Flush WM_PRINT so PowerPoint DirectX slide viewport syncs to HiDPI bounds cleanly
                try {
                    $tmpBmp = New-Object System.Drawing.Bitmap(16, 16)
                    $tmpG = [System.Drawing.Graphics]::FromImage($tmpBmp)
                    $tmpHdc = $tmpG.GetHdc()
                    [void][LiveCaptureEngine]::PrintWindow($pHwnd, $tmpHdc, 0)
                    $tmpG.ReleaseHdc($tmpHdc)
                    $tmpG.Dispose()
                    $tmpBmp.Dispose()
                    $ppt.ActiveWindow.Panes.Item(2).Activate()
                    $ppt.ActiveWindow.View.ZoomToFit = -1
                } catch {}
                Start-Sleep -Milliseconds 650
                $outName = "ppt-{0:D2}.webp" -f $idx
                $rawPng  = Join-Path $tmpDir ($outName -replace '\.webp$', '.png')
                [LiveCaptureEngine]::CaptureCompositeFrame($pHwnd, $pPid, [uint32]$PID, -1, -1, 0, $rawPng)
                Convert-ToWebP $rawPng (Join-Path $outDir $outName)
            }
        } finally {
            if ($pres) { try { $pres.Close() } catch {} }
            if ($ppt) { try { $ppt.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {} }
        }
    }

    Write-Host "`n[PASS] Generated Playbook preview WebP images in docs/public/examples/previews/" -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}
