<#
.SYNOPSIS
    Generates AppIcon.ico from the Asurgent brand palette - a simple
    meeting-room display glyph (screen + stand) on the navy rounded-square
    mark used throughout UI/MainWindow.xaml.

.DESCRIPTION
    Run this after changing the design below and commit the resulting
    AppIcon.ico - it is not generated automatically. Build-Exe.ps1
    embeds it into Start-MeetingRoomProvisioning.exe, and
    UI/MainWindow.xaml references it as the running window's icon, so a
    rebuild of the .exe (see Build-Exe.ps1) is needed after regenerating
    this file for the .exe's icon specifically to pick up any change; the
    window icon picks it up on the next launch with no rebuild needed.

    Pure System.Drawing (GDI+) - no external image tools or downloads.
    Hand-assembles a proper multi-resolution .ico (PNG-compressed entries
    per size, the format Windows has supported since Vista) rather than
    relying on Bitmap.Save(..., ImageFormat.Icon), which only writes a
    single resolution.
#>
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'AppIcon.ico')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Asurgent brand palette - see UI/MainWindow.xaml's Window.Resources for
# the source of truth these are copied from.
$navy   = [System.Drawing.Color]::FromArgb(255, 0x02, 0x00, 0x38)   # NavyDeepBrush
$white  = [System.Drawing.Color]::White
$accent = [System.Drawing.Color]::FromArgb(255, 0x29, 0x62, 0xFF)   # AccentBrush

function New-IconBitmap([int]$Size) {
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    # Background: rounded navy square, matching the in-app sidebar logo mark.
    $radius = [Math]::Max(2, [int]($Size * 0.22))
    $rect = New-Object System.Drawing.Rectangle 0, 0, ($Size - 1), ($Size - 1)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
    $path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
    $path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $navyBrush = New-Object System.Drawing.SolidBrush $navy
    $g.FillPath($navyBrush, $path)

    # Foreground glyph: a simple meeting-room display screen + stand -
    # degrades to a plain white rounded shape on navy at 16px, which still
    # reads as a distinct, legible icon rather than a muddy blob.
    $screenW = $Size * 0.56
    $screenH = $Size * 0.38
    $screenX = ($Size - $screenW) / 2
    $screenY = $Size * 0.22
    $screenRadius = [Math]::Max(1, $Size * 0.06)

    $screenRect = New-Object System.Drawing.RectangleF $screenX, $screenY, $screenW, $screenH
    $screenPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $sd = $screenRadius * 2
    $screenPath.AddArc($screenRect.X, $screenRect.Y, $sd, $sd, 180, 90)
    $screenPath.AddArc($screenRect.Right - $sd, $screenRect.Y, $sd, $sd, 270, 90)
    $screenPath.AddArc($screenRect.Right - $sd, $screenRect.Bottom - $sd, $sd, $sd, 0, 90)
    $screenPath.AddArc($screenRect.X, $screenRect.Bottom - $sd, $sd, $sd, 90, 90)
    $screenPath.CloseFigure()
    $whiteBrush = New-Object System.Drawing.SolidBrush $white
    $g.FillPath($whiteBrush, $screenPath)

    # Small accent dot (camera) at top-center of the screen.
    $dotR = [Math]::Max(1, $Size * 0.035)
    $dotX = $Size / 2 - $dotR
    $dotY = $screenY + $screenH * 0.22 - $dotR
    $accentBrush = New-Object System.Drawing.SolidBrush $accent
    $g.FillEllipse($accentBrush, $dotX, $dotY, $dotR * 2, $dotR * 2)

    # Stand: small accent rectangle beneath the screen.
    $standW = $Size * 0.20
    $standH = $Size * 0.07
    $standX = ($Size - $standW) / 2
    $standY = $screenY + $screenH + $Size * 0.03
    $g.FillRectangle($accentBrush, $standX, $standY, $standW, $standH)

    # Base: slightly wider accent bar under the stand.
    $baseW = $Size * 0.34
    $baseH = $Size * 0.045
    $baseX = ($Size - $baseW) / 2
    $baseY = $standY + $standH
    $g.FillRectangle($accentBrush, $baseX, $baseY, $baseW, $baseH)

    $g.Dispose()
    return $bmp
}

# Standard Windows icon sizes, largest first.
$sizes = @(256, 64, 48, 32, 16)

# String keys, not int: an [ordered] hashtable is backed by
# OrderedDictionary, whose indexer treats an *integer* key as a positional
# index rather than a dictionary key lookup (confirmed by hitting exactly
# this "index out of range" error with an int key before enough entries
# existed for that position to exist). Keying by "$s" (string) sidesteps
# the ambiguous overload entirely.
$pngBytesBySize = [ordered]@{}
foreach ($s in $sizes) {
    $bmp = New-IconBitmap -Size $s
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $pngBytesBySize["$s"] = $ms.ToArray()
    $ms.Dispose()
    $bmp.Dispose()
}

# Hand-assemble a multi-resolution .ico (PNG-compressed entries per size).
$fs = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create)
$bw = New-Object System.IO.BinaryWriter $fs
try {
    # ICONDIR header
    $bw.Write([UInt16]0)      # reserved
    $bw.Write([UInt16]1)      # type = icon
    $bw.Write([UInt16]$sizes.Count)

    $headerSize = 6
    $dirEntrySize = 16
    $offset = $headerSize + ($dirEntrySize * $sizes.Count)

    foreach ($s in $sizes) {
        $bytes = $pngBytesBySize["$s"]
        $sizeByte = if ($s -ge 256) { 0 } else { $s }  # 0 means 256 in ICO format
        $bw.Write([Byte]$sizeByte)   # width
        $bw.Write([Byte]$sizeByte)   # height
        $bw.Write([Byte]0)           # color palette
        $bw.Write([Byte]0)           # reserved
        $bw.Write([UInt16]1)         # color planes
        $bw.Write([UInt16]32)        # bits per pixel
        $bw.Write([UInt32]$bytes.Length)
        $bw.Write([UInt32]$offset)
        $offset += $bytes.Length
    }
    foreach ($s in $sizes) {
        $bw.Write($pngBytesBySize["$s"])
    }
    $bw.Flush()
} finally {
    $bw.Close()
    $fs.Close()
}

Write-Host "Wrote $OutputPath" -ForegroundColor Green
