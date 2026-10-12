param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '../assets/word-showcase.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 1. Load Smiles2Img.Core for high-res molecule rendering
$coreDll = (Resolve-Path (Join-Path $PSScriptRoot "../dist/x64/Word/Smiles2Img.Core.dll")).Path
[System.Reflection.Assembly]::LoadFrom($coreDll) | Out-Null
[Smiles2Img.Rendering.NativeIndigo]::Initialize((Split-Path $coreDll -Parent))

# Render sample molecules to temp files
$samples = @(
    @{ Name = "Aspirin (아스피린)"; Formula = "C9H8O4"; Smiles = "CC(=O)Oc1ccccc1C(=O)O"; Category = "NSAID 진통소염제" },
    @{ Name = "Caffeine (카페인)"; Formula = "C8H10N4O2"; Smiles = "Cn1cnc2c1c(=O)n(C)c(=O)n2C"; Category = "크산틴 알칼로이드" },
    @{ Name = "Paracetamol (아세트아미노펜)"; Formula = "C8H9NO2"; Smiles = "CC(=O)Nc1ccc(O)cc1"; Category = "해열진통제" },
    @{ Name = "Nicotine (니코틴)"; Formula = "C10H14N2"; Smiles = "c1cncc(c1)[C@@H]2CCCN2C"; Category = "천연 알칼로이드" }
)

$tempFiles = @()
foreach ($s in $samples) {
    $bytes = [Smiles2Img.Rendering.MoleculeRenderer]::Render($s.Smiles)
    $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "word_smiles_$([Guid]::NewGuid().ToString('N')).png")
    [System.IO.File]::WriteAllBytes($tmp, $bytes)
    $s.TempPath = $tmp
    $tempFiles += $tmp
}

# 2. Automate Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Add()

    # Title
    $titlePara = $doc.Paragraphs.Add()
    $titlePara.Range.Text = "화학 화합물 연구 보고서 (SMILES for Word)"
    $titlePara.Range.Font.Name = "Malgun Gothic"
    $titlePara.Range.Font.Size = 18
    $titlePara.Range.Font.Bold = 1
    $titlePara.Range.Font.Color = 0x1E293B
    $titlePara.Range.InsertParagraphAfter()

    # Subtitle
    $subPara = $doc.Paragraphs.Add()
    $subPara.Range.Text = "EPAM Indigo 분자 구조식 자동 렌더링 및 Ketcher 에디터 연동 샘플 문서"
    $subPara.Range.Font.Name = "Malgun Gothic"
    $subPara.Range.Font.Size = 10.5
    $subPara.Range.Font.Color = 0x64748B
    $subPara.Range.InsertParagraphAfter()

    # Table (5 rows, 4 columns)
    $tableRange = $doc.Paragraphs.Add().Range
    $table = $doc.Tables.Add($tableRange, $samples.Count + 1, 4)
    $table.Borders.Enable = 1
    $table.Borders.OutsideColor = 0xCBD5E1
    $table.Borders.InsideColor = 0xE2E8F0

    # Header Row
    $headers = @("화합물명", "화학식 / 용도", "2D 화학 구조식 (Double-Click)", "SMILES 코드")
    for ($c = 1; $c -le 4; $c++) {
        $cell = $table.Cell(1, $c)
        $cell.Range.Text = $headers[$c - 1]
        $cell.Range.Font.Name = "Malgun Gothic"
        $cell.Range.Font.Size = 10
        $cell.Range.Font.Bold = 1
        $cell.Shading.BackgroundPatternColor = 0xF1F5F9
        $cell.VerticalAlignment = 1
    }

    # Data Rows
    for ($i = 0; $i -lt $samples.Count; $i++) {
        $row = $i + 2
        $item = $samples[$i]

        $c1 = $table.Cell($row, 1)
        $c1.Range.Text = $item.Name
        $c1.Range.Font.Name = "Malgun Gothic"
        $c1.Range.Font.Size = 9.5
        $c1.Range.Font.Bold = 1

        $c2 = $table.Cell($row, 2)
        $c2.Range.Text = "$($item.Formula)`r`n$($item.Category)"
        $c2.Range.Font.Name = "Segoe UI"
        $c2.Range.Font.Size = 9

        # Structure Image in Column 3
        $c3 = $table.Cell($row, 3)
        $pic = $c3.Range.InlineShapes.AddPicture($item.TempPath, $false, $true)
        $pic.Width = 140
        $pic.Height = 105
        $pic.AlternativeText = "SMILES:" + $item.Smiles
        $pic.Title = "SMILES:" + $item.Smiles

        $c4 = $table.Cell($row, 4)
        $c4.Range.Text = $item.Smiles
        $c4.Range.Font.Name = "Consolas"
        $c4.Range.Font.Size = 8.5
        $c4.Range.Font.Color = 0x475569

        for ($col = 1; $col -le 4; $col++) {
            $table.Cell($row, $col).VerticalAlignment = 1
        }
    }

    # Column widths
    $table.Columns.Item(1).Width = 120
    $table.Columns.Item(2).Width = 110
    $table.Columns.Item(3).Width = 150
    $table.Columns.Item(4).Width = 140

    # Note
    $notePara = $doc.Paragraphs.Add()
    $notePara.Range.Text = "※ 문서 내의 임의의 분자 구조식 그림을 마우스로 더블 클릭하면 내장된 EPAM Ketcher 2D 에디터가 열려 바로 편집할 수 있습니다."
    $notePara.Range.Font.Name = "Malgun Gothic"
    $notePara.Range.Font.Size = 9
    $notePara.Range.Font.Italic = 1
    $notePara.Range.Font.Color = 0x2563EB

    # Copy Table as Picture to Clipboard or Export to High-Res PNG
    $exportPath = [System.IO.Path]::GetFullPath($OutputPath)
    $doc.Range().CopyAsPicture()
    Start-Sleep -Milliseconds 500

    $img = [System.Windows.Forms.Clipboard]::GetImage()
    if ($img) {
        $img.Save($exportPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host "[+] Saved Word Showcase: $exportPath ($($img.Width)x$($img.Height))" -ForegroundColor Green
    } else {
        $pdfPath = [System.IO.Path]::ChangeExtension($exportPath, ".pdf")
        $doc.ExportAsFixedFormat($pdfPath, 17)
        Write-Host "[+] Saved Word Showcase PDF: $pdfPath" -ForegroundColor Green

        # Render PDF to 2x High-Res PNG via WinRT
        try {
            Add-Type -AssemblyName "System.Runtime.WindowsRuntime"
            [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
            [Windows.Storage.StorageFolder, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
            [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime] | Out-Null

            $asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | 
                Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.ContainsGenericParameters } |
                Select-Object -First 1

            $asTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() | 
                Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and -not $_.ContainsGenericParameters } |
                Select-Object -First 1

            $fileTask = $asTaskGeneric.MakeGenericMethod([Windows.Storage.StorageFile]).Invoke($null, @([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath)))
            $fileTask.Wait()
            $docTask = $asTaskGeneric.MakeGenericMethod([Windows.Data.Pdf.PdfDocument]).Invoke($null, @([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($fileTask.Result)))
            $docTask.Wait()
            $pdfDoc = $docTask.Result

            $page = $pdfDoc.GetPage(0)
            $folderTask = $asTaskGeneric.MakeGenericMethod([Windows.Storage.StorageFolder]).Invoke($null, @([Windows.Storage.StorageFolder]::GetFolderFromPathAsync((Split-Path $exportPath))))
            $folderTask.Wait()
            $outTask = $asTaskGeneric.MakeGenericMethod([Windows.Storage.StorageFile]).Invoke($null, @($folderTask.Result.CreateFileAsync((Split-Path $exportPath -Leaf), [Windows.Storage.CreationCollisionOption]::ReplaceExisting)))
            $outTask.Wait()
            $streamTask = $asTaskGeneric.MakeGenericMethod([Windows.Storage.Streams.IRandomAccessStream]).Invoke($null, @($outTask.Result.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)))
            $streamTask.Wait()
            $stream = $streamTask.Result

            $renderOptions = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $renderOptions.DestinationWidth = [uint32]($page.Size.Width * 2)
            $renderOptions.DestinationHeight = [uint32]($page.Size.Height * 2)

            $rTask = $asTaskAction.Invoke($null, @($page.RenderToStreamAsync($stream, $renderOptions)))
            $rTask.Wait()
            $fTask = $asTaskGeneric.MakeGenericMethod([bool]).Invoke($null, @($stream.FlushAsync()))
            $fTask.Wait()
            $stream.Dispose()
            $page.Dispose()
            Write-Host "[+] Rendered High-Res Word Showcase PNG: $exportPath" -ForegroundColor Green
        } catch {
            Write-Warning "WinRT PDF rendering failed: $_"
        }
    }

    $doc.Close($false)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    foreach ($tmp in $tempFiles) {
        try { if (Test-Path $tmp) { Remove-Item $tmp -Force } } catch { }
    }
}
