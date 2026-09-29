Param(
    [string]$SourceDir = (Join-Path $PSScriptRoot "..\..\branding\ibp\source"),
    [string]$OutputDir = (Join-Path $PSScriptRoot "..\..\branding\ibp")
)

# Prepares private branding images (kept out of git in branding/) for the add-in build:
#   header.png  dialog banner: cropped to the logo, 96 px high (two times the 48 px header for high DPI)
#   app.png     window/ribbon icon: cropped square, 256 x 256 px
# Content is found as the pixels darker than the background; a margin keeps the logo off the edges.

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

function Get-ContentBounds([Drawing.Bitmap]$bitmap) {
    $minX = $bitmap.Width; $minY = $bitmap.Height; $maxX = -1; $maxY = -1
    for ($y = 0; $y -lt $bitmap.Height; $y += 2) {
        for ($x = 0; $x -lt $bitmap.Width; $x += 2) {
            $p = $bitmap.GetPixel($x, $y)
            if ($p.A -gt 32 -and (($p.R + $p.G + $p.B) / 3) -lt 200) {
                if ($x -lt $minX) { $minX = $x }; if ($x -gt $maxX) { $maxX = $x }
                if ($y -lt $minY) { $minY = $y }; if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }
    if ($maxX -lt 0) { throw "No logo content found." }
    return New-Object Drawing.Rectangle($minX, $minY, ($maxX - $minX + 1), ($maxY - $minY + 1))
}

function Save-Scaled([Drawing.Bitmap]$source, [Drawing.Rectangle]$crop, [int]$width, [int]$height, [Drawing.Color]$background, [string]$path) {
    $target = New-Object Drawing.Bitmap($width, $height, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $graphics = [Drawing.Graphics]::FromImage($target)
        try {
            $graphics.Clear($background)
            $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.DrawImage($source, (New-Object Drawing.Rectangle(0, 0, $width, $height)), $crop, [Drawing.GraphicsUnit]::Pixel)
        }
        finally { $graphics.Dispose() }
        $target.Save($path, [Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $target.Dispose() }
}

$SourceDir = (Resolve-Path $SourceDir).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

# Banner
$header = New-Object Drawing.Bitmap (Join-Path $SourceDir "header-solid-blue-164x48.png")
try {
    $background = $header.GetPixel(2, 2)
    $content = Get-ContentBounds $header
    $margin = [int]($content.Height * 0.18)
    $crop = [Drawing.Rectangle]::Intersect(
        (New-Object Drawing.Rectangle(($content.X - $margin), ($content.Y - $margin), ($content.Width + 2 * $margin), ($content.Height + 2 * $margin))),
        (New-Object Drawing.Rectangle(0, 0, $header.Width, $header.Height)))
    $height = 96
    $width = [int][Math]::Round($crop.Width * $height / $crop.Height)
    Save-Scaled $header $crop $width $height $background (Join-Path $OutputDir "header.png")
    Write-Host ("header.png: {0} x {1} px, Hintergrund #{2:X2}{3:X2}{4:X2}" -f $width, $height, $background.R, $background.G, $background.B)
}
finally { $header.Dispose() }

# Icon
$app = New-Object Drawing.Bitmap (Join-Path $SourceDir "app.png")
try {
    $background = $app.GetPixel(2, 2)
    $content = Get-ContentBounds $app
    $side = [Math]::Max($content.Width, $content.Height) + 8
    $cx = $content.X + [int]($content.Width / 2); $cy = $content.Y + [int]($content.Height / 2)
    $crop = [Drawing.Rectangle]::Intersect(
        (New-Object Drawing.Rectangle(($cx - [int]($side / 2)), ($cy - [int]($side / 2)), $side, $side)),
        (New-Object Drawing.Rectangle(0, 0, $app.Width, $app.Height)))
    Save-Scaled $app $crop 256 256 $background (Join-Path $OutputDir "app.png")
    Write-Host "app.png: 256 x 256 px"
}
finally { $app.Dispose() }
