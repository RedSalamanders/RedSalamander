<#
.SYNOPSIS
    Generate MSIX tile assets from RedSalamander logo.png.
.DESCRIPTION
    Resizes `RedSalamander\res\logo.png` into the set of PNG assets referenced by
    `Installer\msix\Package.appxmanifest` (Square44x44, Square150x150, Wide310x150,
    Square310x310, StoreLogo).
.EXAMPLE
    .\Installer\msix\GenerateAssets.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\\..")).Path
$sourceLogo = Join-Path $repoRoot "RedSalamander\\res\\logo.png"
$assetsDir = Join-Path $PSScriptRoot "Assets"

if (-not (Test-Path $sourceLogo)) {
    throw "Source logo not found: $sourceLogo"
}

New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null

$targets = @(
    @{ Name = "Square44x44Logo.png";   Width = 44;  Height = 44  }
    @{ Name = "Square150x150Logo.png"; Width = 150; Height = 150 }
    @{ Name = "Wide310x150Logo.png";   Width = 310; Height = 150 }
    @{ Name = "Square310x310Logo.png"; Width = 310; Height = 310 }
    @{ Name = "StoreLogo.png";         Width = 50;  Height = 50  }
)

function New-ResizedPng {
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Image]$Source,

        [Parameter(Mandatory = $true)]
        [int]$Width,

        [Parameter(Mandatory = $true)]
        [int]$Height,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $bitmap.SetResolution(96, 96)

        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([System.Drawing.Color]::Transparent)
            $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

            $scale = [Math]::Min(($Width / [double]$Source.Width), ($Height / [double]$Source.Height))
            $drawWidth = [Math]::Max([int][Math]::Round($Source.Width * $scale), 1)
            $drawHeight = [Math]::Max([int][Math]::Round($Source.Height * $scale), 1)

            $x = [int][Math]::Floor(($Width - $drawWidth) / 2.0)
            $y = [int][Math]::Floor(($Height - $drawHeight) / 2.0)

            $destRect = New-Object System.Drawing.Rectangle $x, $y, $drawWidth, $drawHeight
            $graphics.DrawImage($Source, $destRect)
        }
        finally {
            $graphics.Dispose()
        }

        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }
}

$sourceImage = [System.Drawing.Image]::FromFile($sourceLogo)
try {
    foreach ($target in $targets) {
        $outputPath = Join-Path $assetsDir $target.Name
        New-ResizedPng -Source $sourceImage -Width $target.Width -Height $target.Height -Path $outputPath
        Write-Host "Wrote: $outputPath" -ForegroundColor Cyan
    }
}
finally {
    $sourceImage.Dispose()
}
