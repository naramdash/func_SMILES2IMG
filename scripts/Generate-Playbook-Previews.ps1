<#
.SYNOPSIS
    Generates lightweight 16:9 (1280x720 WebP) carousel preview images directly from the
    3 official self-contained Playbook files (Excel .xlsx, Word .docx, PowerPoint .pptx)
    into docs/public/examples/previews/ for the Homepage Playbook Preview Carousel.
#>
param(
    [ValidateSet('All', 'Excel', 'Word', 'PowerPoint')]
    [string]$App = 'All'
)

$ErrorActionPreference = 'Stop'
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
    & $ffmpeg -hide_banner -loglevel error -y -i $srcPng -vf 'scale=1280:720:flags=lanczos' -c:v libwebp -quality 86 $dstWebp
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg WebP conversion failed: $srcPng -> $dstWebp" }
    $kb = [math]::Round((Get-Item $dstWebp).Length / 1KB, 1)
    Write-Host ("  [OK] {0,-18} ({1,5:N1} KB)" -f (Split-Path $dstWebp -Leaf), $kb) -ForegroundColor Green
}

try {
    # =========================================================================
    # 1. EXCEL PLAYBOOK PREVIEWS (5 key tabs captured live from .xlsx)
    # =========================================================================
    if ($App -eq 'All' -or $App -eq 'Excel') {
        Write-Host "`n[1/3] Capturing Excel Playbook Tab Previews..." -ForegroundColor Cyan
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
            [LiveCaptureEngine]::SetMainWindowSize($xHwnd, 8, 8, 1920, 1080)
            Start-Sleep -Milliseconds 900
            [void][LiveCaptureEngine]::SelectRibbonTab($xHwnd, 'SMILES')
            Start-Sleep -Milliseconds 800

            $excelTabs = @(
                @{ Name = '0. Start Here';            Zoom = 100; Cell = 'E6'; Out = 'excel-01.webp' },
                @{ Name = '1. Structure Rendering';   Zoom = 96;  Cell = 'E4'; Out = 'excel-02.webp' },
                @{ Name = '2. Properties & Lipinski'; Zoom = 78;  Cell = 'D4'; Out = 'excel-03.webp' },
                @{ Name = '3. SAR & Similarity';      Zoom = 105; Cell = 'E5'; Out = 'excel-04.webp' },
                @{ Name = '4. Notation Conversions';  Zoom = 91;  Cell = 'F4'; Out = 'excel-05.webp' },
                @{ Name = '5. Styles & Transforms';   Zoom = 104; Cell = 'F5'; Out = 'excel-06.webp' },
                @{ Name = '6. Reactions & KET';       Zoom = 104; Cell = 'E4'; Out = 'excel-07.webp' },
                @{ Name = '7. Real-World Workflows';  Zoom = 110; Cell = 'D6'; Out = 'excel-08.webp' },
                @{ Name = '8. Legacy Excel (.FLOAT)'; Zoom = 96;  Cell = 'G4'; Out = 'excel-09.webp' },
                @{ Name = '9. Ribbon, SDF & Ref';     Zoom = 108; Cell = 'B6'; Out = 'excel-10.webp' }
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
    # 2. WORD PLAYBOOK PREVIEWS (3 live 2-page spreads inside Microsoft Word)
    # =========================================================================
    if ($App -eq 'All' -or $App -eq 'Word') {
        Write-Host "`n[2/3] Capturing Word Playbook Live 2-Page Spread Previews..." -ForegroundColor Cyan
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
            # 2304x1296 (exact 16:9) keeps Word in true symmetric 2-column spread centering (no 3-column left shift)
            [LiveCaptureEngine]::SetMainWindowSize($wHwnd, 8, 8, 2304, 1296)
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
            $view.Zoom.PageColumns = 2
            $view.Zoom.PageRows = 1
            Start-Sleep -Milliseconds 600
            # Back off auto-fit zoom slightly so full page height + bottom margin stay cleanly inside frame
            $targetWordZoom = [Math]::Max(88, [int]$view.Zoom.Percentage - 12)
            $view.Zoom.Percentage = $targetWordZoom
            Start-Sleep -Milliseconds 400
            [void][LiveCaptureEngine]::SelectRibbonTab($wHwnd, 'SMILES')
            Start-Sleep -Milliseconds 700

            # Capture each 2-page spread cleanly at scroll=0 (avoids Word PrintWindow scroll cache bug)
            $spreadOuts = @('word-01.webp', 'word-02.webp', 'word-03.webp')
            for ($si = 0; $si -lt $spreadOuts.Count; $si++) {
                if ($si -gt 0) {
                    # Delete the current first 2 pages of our temp copy so the next 2-page spread sits at page 1-2
                    $r1 = $word.Selection.GoTo(1, 1, 1)
                    $r3 = $word.Selection.GoTo(1, 1, 3)
                    if ($r3.Start -gt $r1.Start) {
                        $delRng = $doc.Range($r1.Start, $r3.Start)
                        $delRng.Delete() | Out-Null
                        while ($doc.Paragraphs.Count -gt 1 -and ($doc.Paragraphs.Item(1).Range.Text -replace '[\r\n\f\s]', '') -eq '' -and $doc.Paragraphs.Item(1).Range.InlineShapes.Count -eq 0 -and $doc.Paragraphs.Item(1).Range.Tables.Count -eq 0) {
                            $doc.Paragraphs.Item(1).Range.Delete() | Out-Null
                        }
                    }
                    $view.Zoom.PageColumns = 2
                    $view.Zoom.PageRows = 1
                    $view.Zoom.Percentage = $targetWordZoom
                    $word.ActiveWindow.VerticalPercentScrolled = 0
                    $word.ScreenRefresh()
                    Start-Sleep -Milliseconds 900
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
                $rawPng = Join-Path $tmpDir ($spreadOuts[$si] -replace '\.webp$', '.png')
                [LiveCaptureEngine]::CaptureCompositeFrame($wHwnd, $wPid, [uint32]$PID, -1, -1, 0, $rawPng)
                Convert-ToWebP $rawPng (Join-Path $outDir $spreadOuts[$si])
            }
        } finally {
            if ($doc) { try { $doc.Close($false) } catch {} }
            if ($word) { try { $word.Quit(0); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {} }
            Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }

    # =========================================================================
    # 3. POWERPOINT PLAYBOOK PREVIEWS (all 8 slides captured live in PPT Editor UI)
    # =========================================================================
    if ($App -eq 'All' -or $App -eq 'PowerPoint') {
        Write-Host "`n[3/3] Capturing PowerPoint Playbook Live Editor Previews (8 Slides)..." -ForegroundColor Cyan
        $pptxPath = (Resolve-Path (Join-Path $examplesDir 'func_SMILES2IMG_PowerPoint_Playbook.pptx')).Path
        Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue
        $ppt = New-Object -ComObject PowerPoint.Application
        $pres = $null
        try {
            $pres = $ppt.Presentations.Open($pptxPath, -1, 0, -1) # ReadOnly=True, Untitled=False, WithWindow=True
            $pHwnd = [IntPtr]$ppt.HWND
            $pPid  = [LiveCaptureEngine]::GetPidForHwnd($pHwnd)
            try { $ppt.ActiveWindow.WindowState = 1 } catch {} # ppWindowNormal
            [LiveCaptureEngine]::SetMainWindowSize($pHwnd, 8, 8, 1920, 1080)
            Start-Sleep -Milliseconds 900
            [void][LiveCaptureEngine]::SelectRibbonTab($pHwnd, 'SMILES')
            Start-Sleep -Milliseconds 800

            try {
                $ppt.ActiveWindow.ViewType = 9 # ppViewNormal
                $ppt.ActiveWindow.SplitHorizontal = 15
                $ppt.ActiveWindow.SplitVertical = 86
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
