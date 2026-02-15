$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

function New-DebugBadgedBitmap {
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Icon]$Icon,

        [Parameter(Mandatory = $true)]
        [int]$Size
    )

    $bitmap = $null
    $graphics = $null
    $bgBrush = $null
    $borderPen = $null
    $textBrush = $null
    $font = $null
    $stringFormat = $null
    try {
        $bitmap = $Icon.ToBitmap()
        if ($bitmap.Width -ne $Size -or $bitmap.Height -ne $Size) {
            $scaled = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $graphics = [System.Drawing.Graphics]::FromImage($scaled)
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.DrawImage($bitmap, 0, 0, $Size, $Size)
            $bitmap.Dispose()
            $bitmap = $scaled
            $graphics.Dispose()
            $graphics = $null
        }

        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

        $padding = [Math]::Max(1, [int][Math]::Round($Size * 0.05))
        $diameter = [Math]::Max(8, [int][Math]::Round($Size * 0.44))
        $x = $Size - $diameter - $padding
        $y = $Size - $diameter - $padding
        $rect = [System.Drawing.Rectangle]::new($x, $y, $diameter, $diameter)

        $bgColor = [System.Drawing.Color]::FromArgb(235, 255, 191, 0) # Amber-ish
        $bgBrush = [System.Drawing.SolidBrush]::new($bgColor)

        $borderWidth = [Math]::Max(1, [int][Math]::Round($Size * 0.03))
        $borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(230, 0, 0, 0), $borderWidth)

        $text = if ($Size -lt 32) { "D" } else { "DBG" }
        $fontSize = if ($text -eq "D") { [float]($diameter * 0.62) } else { [float]($diameter * 0.36) }
        $font = [System.Drawing.Font]::new("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $textBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(245, 0, 0, 0))

        $stringFormat = [System.Drawing.StringFormat]::new()
        $stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
        $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center

        $graphics.FillEllipse($bgBrush, $rect)
        $graphics.DrawEllipse($borderPen, $rect)
        $rectF = [System.Drawing.RectangleF]::new($rect.X, $rect.Y, $rect.Width, $rect.Height)
        $graphics.DrawString($text, $font, $textBrush, $rectF, $stringFormat)

        return $bitmap
    }
    catch {
        if ($bitmap) { $bitmap.Dispose() }
        throw
    }
    finally {
        if ($graphics) { $graphics.Dispose() }
        if ($bgBrush) { $bgBrush.Dispose() }
        if ($borderPen) { $borderPen.Dispose() }
        if ($textBrush) { $textBrush.Dispose() }
        if ($font) { $font.Dispose() }
        if ($stringFormat) { $stringFormat.Dispose() }
    }
}

function Write-IcoFileFromPngEntries {
    param(
        [Parameter(Mandatory = $true)]
        [int[]]$Sizes,

        [Parameter(Mandatory = $true)]
        [byte[][]]$PngEntries,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    if ($Sizes.Length -ne $PngEntries.Length) {
        throw "Sizes and PngEntries length mismatch."
    }

    $count = $Sizes.Length
    $headerSize = 6 + (16 * $count)
    $offset = $headerSize

    $outDir = Split-Path -Parent $OutputPath
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
    }

    $fs = $null
    $bw = $null
    try {
        $fs = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $bw = [System.IO.BinaryWriter]::new($fs)

        $bw.Write([UInt16]0) # reserved
        $bw.Write([UInt16]1) # type (icon)
        $bw.Write([UInt16]$count)

        for ($i = 0; $i -lt $count; $i++) {
            $size = $Sizes[$i]
            $png = $PngEntries[$i]

            $w = if ($size -ge 256) { 0 } else { $size }
            $h = if ($size -ge 256) { 0 } else { $size }

            $bw.Write([byte]$w)
            $bw.Write([byte]$h)
            $bw.Write([byte]0) # color count
            $bw.Write([byte]0) # reserved
            $bw.Write([UInt16]1) # planes
            $bw.Write([UInt16]32) # bit count
            $bw.Write([UInt32]$png.Length)
            $bw.Write([UInt32]$offset)
            $offset += $png.Length
        }

        for ($i = 0; $i -lt $count; $i++) {
            $bw.Write($PngEntries[$i])
        }
    }
    finally {
        if ($bw) { $bw.Dispose() }
        if ($fs) { $fs.Dispose() }
    }
}

function New-DebugBadgedIconFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputIconPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputIconPath
    )

    if (-not (Test-Path $InputIconPath)) {
        throw "Input icon not found: $InputIconPath"
    }

    $sizes = @(16, 32, 48, 256)
    $pngEntries = [System.Collections.Generic.List[byte[]]]::new()

    foreach ($size in $sizes) {
        $icon = $null
        $bitmap = $null
        $ms = $null
        try {
            $icon = [System.Drawing.Icon]::new($InputIconPath, $size, $size)
            $bitmap = New-DebugBadgedBitmap -Icon $icon -Size $size
            $ms = [System.IO.MemoryStream]::new()
            $bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $pngEntries.Add($ms.ToArray()) | Out-Null
        }
        finally {
            if ($ms) { $ms.Dispose() }
            if ($bitmap) { $bitmap.Dispose() }
            if ($icon) { $icon.Dispose() }
        }
    }

    Write-IcoFileFromPngEntries -Sizes $sizes -PngEntries $pngEntries.ToArray() -OutputPath $OutputIconPath
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

$pairs = @(
    @{
        In  = (Join-Path $repoRoot "RedSalamander\\res\\logo.ico")
        Out = (Join-Path $repoRoot "RedSalamander\\res\\logo_debug.ico")
    },
    @{
        In  = (Join-Path $repoRoot "RedSalamanderMonitor\\res\\logo_background.ico")
        Out = (Join-Path $repoRoot "RedSalamanderMonitor\\res\\logo_background_debug.ico")
    }
)

foreach ($pair in $pairs) {
    Write-Host "Generating: $($pair.Out)" -ForegroundColor Yellow
    New-DebugBadgedIconFile -InputIconPath $pair.In -OutputIconPath $pair.Out
}
