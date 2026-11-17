<#
.SYNOPSIS
    Shared GDI+ renderer for authentic Office 365 cascading right-click context menus
    matching RibbonXmlBuilder.cs in v1.8.0+.
#>

function Draw-OfficeContextMenuPanel {
    param(
        [System.Drawing.Graphics]$g,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [array]$Items,
        [System.Drawing.Color]$AccentColor,
        [hashtable]$MiniToolbar = $null,
        [bool]$ShowSearchBox = $false
    )

    $rowH = 24
    $sepH = 9
    $padY = 6
    $headerH = 0
    if ($MiniToolbar) { $headerH += 46 }
    if ($ShowSearchBox) { $headerH += 34 }

    $contentH = 0
    foreach ($it in $Items) {
        $contentH += if ($it.Kind -eq 'sep') { $sepH } else { $rowH }
    }
    $totalH = $headerH + $contentH + $padY * 2

    # Soft multi-layer shadow
    for ($s = 7; $s -ge 1; $s--) {
        $alpha = [int](20 - $s * 2.2)
        $sb = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($alpha, 0, 0, 0))
        $g.FillRectangle($sb, ($X - $s + 3), ($Y - $s + 4), ($Width + $s * 2), ($totalH + $s * 2))
        $sb.Dispose()
    }

    $borderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(203, 213, 225), 1)
    $menuRect = New-Object System.Drawing.Rectangle($X, $Y, $Width, $totalH)
    $g.FillRectangle([System.Drawing.Brushes]::White, $menuRect)
    $g.DrawRectangle($borderPen, $menuRect)

    $fontMenu = New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Regular)
    $fontBold = New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Bold)
    $fontSmall = New-Object System.Drawing.Font('Segoe UI', 8.0, [System.Drawing.FontStyle]::Regular)
    $brushText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 41, 59))
    $brushMuted = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100, 116, 139))
    $sepPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(226, 232, 240), 1)

    $curY = $Y

    if ($MiniToolbar) {
        $miniRect = New-Object System.Drawing.Rectangle($X, $curY, $Width, 44)
        $miniBg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 250, 252))
        $g.FillRectangle($miniBg, $miniRect)
        $g.DrawLine($borderPen, $X, ($curY + 44), ($X + $Width), ($curY + 44))
        if ($MiniToolbar.Mode -eq 'text') {
            $g.FillRectangle([System.Drawing.Brushes]::White, ($X + 10), ($curY + 6), 68, 18)
            $g.DrawRectangle($borderPen, ($X + 10), ($curY + 6), 68, 18)
            $g.DrawString('Calibri', $fontSmall, $brushText, ($X + 14), ($curY + 7))
            $g.FillRectangle([System.Drawing.Brushes]::White, ($X + 82), ($curY + 6), 28, 18)
            $g.DrawRectangle($borderPen, ($X + 82), ($curY + 6), 28, 18)
            $g.DrawString('11', $fontSmall, $brushText, ($X + 88), ($curY + 7))
            $g.DrawString('B', $fontBold, $brushText, ($X + 122), ($curY + 6))
            $g.DrawString('I', (New-Object System.Drawing.Font('Segoe UI', 9.0, [System.Drawing.FontStyle]::Italic)), $brushText, ($X + 142), ($curY + 6))
            $g.DrawString('A', $fontBold, (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 38, 38))), ($X + 164), ($curY + 6))
            $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 38, 38))), ($X + 164), ($curY + 20), 10, 2)
            $g.DrawString('Fill', $fontSmall, $brushMuted, ($X + 12), ($curY + 27))
            $g.DrawString('Borders', $fontSmall, $brushMuted, ($X + 46), ($curY + 27))
            $g.DrawString('Align', $fontSmall, $brushMuted, ($X + 100), ($curY + 27))
            $g.DrawString('$   %', $fontSmall, $brushText, ($X + 144), ($curY + 27))
        } else {
            $g.DrawString('Style', $fontSmall, $brushMuted, ($X + 14), ($curY + 14))
            $g.DrawString('Crop', $fontSmall, $brushMuted, ($X + 60), ($curY + 14))
            $g.DrawString('Arrange', $fontSmall, $brushMuted, ($X + 104), ($curY + 14))
            $g.DrawString('Format', $fontSmall, $brushMuted, ($X + 162), ($curY + 14))
        }
        $curY += 46
    }

    if ($ShowSearchBox) {
        $searchRect = New-Object System.Drawing.Rectangle(($X + 8), ($curY + 4), ($Width - 16), 24)
        $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(248, 250, 252))), $searchRect)
        $g.DrawRectangle($borderPen, $searchRect)
        $penMag = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(148, 163, 184), 1.4)
        $g.DrawEllipse($penMag, ($X + 16), ($curY + 11), 7, 7)
        $g.DrawLine($penMag, ($X + 22), ($curY + 17), ($X + 25), ($curY + 20))
        $g.DrawString('Search the menus', $fontSmall, $brushMuted, ($X + 30), ($curY + 9))
        $curY += 34
    }

    $curY += $padY
    $anchorRects = @{}

    foreach ($it in $Items) {
        if ($it.Kind -eq 'sep') {
            $g.DrawLine($sepPen, ($X + 32), ($curY + 4), ($X + $Width - 10), ($curY + 4))
            $curY += $sepH
            continue
        }

        $rowRect = New-Object System.Drawing.Rectangle(($X + 3), $curY, ($Width - 6), $rowH)
        if ($it.Highlight) {
            $hiBg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(24, $AccentColor.R, $AccentColor.G, $AccentColor.B))
            $hiBorder = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(95, $AccentColor.R, $AccentColor.G, $AccentColor.B), 1)
            $g.FillRectangle($hiBg, $rowRect)
            $g.DrawRectangle($hiBorder, $rowRect)
        }

        if ($it.Key) {
            $anchorRects[$it.Key] = $rowRect
        }

        # Crisp vector menu icon (14x14 centered at X+12, curY+5)
        $ix = $X + 12
        $iy = $curY + 5
        $iconColor = if ($it.Highlight) { $AccentColor } elseif ($it.IconColor) { $it.IconColor } else { [System.Drawing.Color]::FromArgb(71, 85, 105) }
        Draw-MenuIcon -g $g -Kind $it.Icon -X $ix -Y $iy -Color $iconColor

        $textBrush = if ($it.Muted) { $brushMuted } elseif ($it.Highlight) { New-Object System.Drawing.SolidBrush($AccentColor) } else { $brushText }
        $fontUse = if ($it.Highlight) { $fontBold } else { $fontMenu }
        $g.DrawString($it.Label, $fontUse, $textBrush, ($X + 34), ($curY + 4))

        if ($it.Shortcut) {
            $g.DrawString($it.Shortcut, $fontSmall, $brushMuted, ($X + $Width - 46), ($curY + 5))
        }
        if ($it.HasSub) {
            $g.DrawString([char]0x203A, $fontBold, $textBrush, ($X + $Width - 16), ($curY + 3))
        }

        $curY += $rowH
    }

    return @{ Bounds = $menuRect; Anchors = $anchorRects }
}

function Draw-MenuIcon {
    param(
        [System.Drawing.Graphics]$g,
        [string]$Kind,
        [int]$X,
        [int]$Y,
        [System.Drawing.Color]$Color
    )
    if ([string]::IsNullOrEmpty($Kind)) { return }
    $pen = New-Object System.Drawing.Pen($Color, 1.35)
    $brush = New-Object System.Drawing.SolidBrush($Color)
    switch ($Kind) {
        'cut' {
            $g.DrawEllipse($pen, ($X + 1), ($Y + 8), 4, 4)
            $g.DrawEllipse($pen, ($X + 1), ($Y + 2), 4, 4)
            $g.DrawLine($pen, ($X + 4), ($Y + 8), ($X + 12), ($Y + 2))
            $g.DrawLine($pen, ($X + 4), ($Y + 4), ($X + 12), ($Y + 11))
        }
        'copy' {
            $g.DrawRectangle($pen, ($X + 4), ($Y + 1), 8, 9)
            $g.FillRectangle([System.Drawing.Brushes]::White, ($X + 1), ($Y + 4), 8, 9)
            $g.DrawRectangle($pen, ($X + 1), ($Y + 4), 8, 9)
        }
        'paste' {
            $g.DrawRectangle($pen, ($X + 2), ($Y + 3), 10, 11)
            $g.FillRectangle($brush, ($X + 4), ($Y + 1), 6, 3)
        }
        'pen' {
            $g.DrawLine($pen, ($X + 2), ($Y + 10), ($X + 10), ($Y + 2))
            $g.DrawLine($pen, ($X + 10), ($Y + 2), ($X + 12), ($Y + 4))
            $g.DrawLine($pen, ($X + 12), ($Y + 4), ($X + 4), ($Y + 12))
            $g.DrawLine($pen, ($X + 2), ($Y + 10), ($X + 1), ($Y + 13))
            $g.DrawLine($pen, ($X + 1), ($Y + 13), ($X + 4), ($Y + 12))
        }
        'chem' {
            # Benzene hexagon icon representing Chemical Structure
            $pts = @(
                (New-Object System.Drawing.PointF(($X + 7), ($Y + 1))),
                (New-Object System.Drawing.PointF(($X + 13), ($Y + 4))),
                (New-Object System.Drawing.PointF(($X + 13), ($Y + 10))),
                (New-Object System.Drawing.PointF(($X + 7), ($Y + 13))),
                (New-Object System.Drawing.PointF(($X + 1), ($Y + 10))),
                (New-Object System.Drawing.PointF(($X + 1), ($Y + 4)))
            )
            $g.DrawPolygon($pen, $pts)
            $g.DrawEllipse($pen, ($X + 4), ($Y + 4), 6, 6)
        }
        'fx' {
            $f = New-Object System.Drawing.Font('Georgia', 8.5, [System.Drawing.FontStyle]::Italic)
            $g.DrawString('fx', $f, $brush, ($X + 1), ($Y + 1))
        }
        'ket' {
            $g.DrawRectangle($pen, ($X + 1), ($Y + 2), 12, 10)
            $g.DrawLine($pen, ($X + 4), ($Y + 6), ($X + 10), ($Y + 6))
            $g.DrawLine($pen, ($X + 4), ($Y + 9), ($X + 8), ($Y + 9))
        }
        'cube' {
            $g.DrawRectangle($pen, ($X + 2), ($Y + 5), 7, 7)
            $g.DrawLine($pen, ($X + 2), ($Y + 5), ($X + 6), ($Y + 2))
            $g.DrawLine($pen, ($X + 9), ($Y + 5), ($X + 13), ($Y + 2))
            $g.DrawLine($pen, ($X + 9), ($Y + 12), ($X + 13), ($Y + 9))
            $g.DrawLine($pen, ($X + 6), ($Y + 2), ($X + 13), ($Y + 2))
            $g.DrawLine($pen, ($X + 13), ($Y + 2), ($X + 13), ($Y + 9))
        }
        'inchi' {
            $g.DrawRectangle($pen, ($X + 1), ($Y + 2), 12, 10)
            $g.DrawLine($pen, ($X + 3), ($Y + 9), ($X + 11), ($Y + 9))
            $g.DrawLine($pen, ($X + 4), ($Y + 7), ($X + 9), ($Y + 3))
        }
        'svg' {
            $g.DrawRectangle($pen, ($X + 1), ($Y + 2), 12, 10)
            $g.DrawLine($pen, ($X + 2), ($Y + 10), ($X + 6), ($Y + 6))
            $g.DrawLine($pen, ($X + 6), ($Y + 6), ($X + 9), ($Y + 9))
            $g.DrawLine($pen, ($X + 9), ($Y + 9), ($X + 12), ($Y + 5))
        }
        'info' {
            $g.DrawRectangle($pen, ($X + 2), ($Y + 1), 10, 12)
            $g.DrawLine($pen, ($X + 4), ($Y + 5), ($X + 10), ($Y + 5))
            $g.DrawLine($pen, ($X + 4), ($Y + 8), ($X + 10), ($Y + 8))
            $g.DrawLine($pen, ($X + 4), ($Y + 11), ($X + 8), ($Y + 11))
        }
        'calc' {
            $g.DrawRectangle($pen, ($X + 2), ($Y + 1), 10, 12)
            $g.FillRectangle($brush, ($X + 4), ($Y + 3), 6, 2)
            $g.FillRectangle($brush, ($X + 4), ($Y + 7), 2, 2)
            $g.FillRectangle($brush, ($X + 8), ($Y + 7), 2, 2)
            $g.FillRectangle($brush, ($X + 4), ($Y + 10), 2, 2)
            $g.FillRectangle($brush, ($X + 8), ($Y + 10), 2, 2)
        }
        'table' {
            $g.DrawRectangle($pen, ($X + 1), ($Y + 2), 12, 10)
            $g.DrawLine($pen, ($X + 1), ($Y + 6), ($X + 13), ($Y + 6))
            $g.DrawLine($pen, ($X + 6), ($Y + 2), ($X + 6), ($Y + 12))
        }
        'bullets' {
            $g.FillRectangle($brush, ($X + 2), ($Y + 3), 2, 2)
            $g.DrawLine($pen, ($X + 6), ($Y + 4), ($X + 12), ($Y + 4))
            $g.FillRectangle($brush, ($X + 2), ($Y + 7), 2, 2)
            $g.DrawLine($pen, ($X + 6), ($Y + 8), ($X + 12), ($Y + 8))
            $g.FillRectangle($brush, ($X + 2), ($Y + 11), 2, 2)
            $g.DrawLine($pen, ($X + 6), ($Y + 12), ($X + 12), ($Y + 12))
        }
        'tag' {
            $f = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
            $g.DrawString('A', $f, $brush, ($X + 2), ($Y - 1))
        }
        'key' {
            $g.DrawEllipse($pen, ($X + 1), ($Y + 4), 5, 5)
            $g.DrawLine($pen, ($X + 6), ($Y + 7), ($X + 13), ($Y + 7))
            $g.DrawLine($pen, ($X + 11), ($Y + 7), ($X + 11), ($Y + 10))
        }
    }
    $pen.Dispose()
    $brush.Dispose()
}

function Get-SharedMoleculeInfoContextItems {
    $blue = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $teal = [System.Drawing.Color]::FromArgb(13, 148, 136)
    $amber = [System.Drawing.Color]::FromArgb(217, 119, 6)
    return @(
        @{ Label = 'Copy Substance Name'; Icon = 'tag'; IconColor = $blue },
        @{ Label = 'Copy Chemical Formula'; Icon = 'fx'; IconColor = $teal },
        @{ Label = 'Copy Molecular Weight (MW)'; Icon = 'calc'; IconColor = $blue },
        @{ Label = 'Copy SMILES'; Icon = 'copy'; IconColor = $blue },
        @{ Label = 'Copy InChI'; Icon = 'copy'; IconColor = $blue },
        @{ Label = 'Copy InChIKey'; Icon = 'key'; IconColor = $amber },
        @{ Label = 'Copy All Summary (Text)'; Icon = 'info'; IconColor = $teal },
        @{ Label = 'Copy as Excel Row (TSV Columns)'; Icon = 'table'; IconColor = $teal },
        @{ Label = 'Copy as Multi-Line List'; Icon = 'bullets'; IconColor = $blue },
        @{ Label = 'Copy as Markdown Table'; Icon = 'table'; IconColor = $blue },
        @{ Kind = 'sep' },
        @{ Label = 'Detailed Molecule Info Window...'; Icon = 'info'; Highlight = $true }
    )
}
