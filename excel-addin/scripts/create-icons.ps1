$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$assetDir = Join-Path $PSScriptRoot "..\assets"
New-Item -ItemType Directory -Force -Path $assetDir | Out-Null

function New-IconPng {
    param(
        [int] $Size,
        [string] $Path
    )

    $bitmap = New-Object System.Drawing.Bitmap $Size, $Size
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::FromArgb(15, 108, 90))

    $fontSize = [Math]::Max(8, [Math]::Floor($Size * 0.42))
    $font = New-Object System.Drawing.Font "Segoe UI", $fontSize, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = New-Object System.Drawing.RectangleF 0, 0, $Size, $Size

    $graphics.DrawString("X", $font, $brush, $rect, $format)
    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

    $format.Dispose()
    $brush.Dispose()
    $font.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()
}

New-IconPng -Size 16 -Path (Join-Path $assetDir "icon-16.png")
New-IconPng -Size 32 -Path (Join-Path $assetDir "icon-32.png")
New-IconPng -Size 64 -Path (Join-Path $assetDir "icon-64.png")
New-IconPng -Size 80 -Path (Join-Path $assetDir "icon-80.png")
