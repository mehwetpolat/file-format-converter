$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$root = Split-Path $PSScriptRoot -Parent
$assetDirectory = Join-Path $root "src/DosyaTipCevirici.WinForms/Assets"
New-Item -ItemType Directory -Path $assetDirectory -Force | Out-Null
$iconPath = Join-Path $assetDirectory "app.ico"
$previewPath = Join-Path $assetDirectory "app-preview.png"

function New-RoundedPath([single]$x, [single]$y, [single]$width, [single]$height, [single]$radius) {
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $diameter = $radius * 2
    $path.AddArc($x, $y, $diameter, $diameter, 180, 90)
    $path.AddArc($x + $width - $diameter, $y, $diameter, $diameter, 270, 90)
    $path.AddArc($x + $width - $diameter, $y + $height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($x, $y + $height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-IconImage([int]$size) {
    $bitmap = [System.Drawing.Bitmap]::new($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.Clear([System.Drawing.Color]::Transparent)
        $graphics.ScaleTransform($size / 256.0, $size / 256.0)

        $backgroundPath = New-RoundedPath 10 10 236 236 55
        $backgroundBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            [System.Drawing.Rectangle]::new(10, 10, 236, 236),
            [System.Drawing.Color]::FromArgb(27, 56, 130),
            [System.Drawing.Color]::FromArgb(42, 113, 221),
            [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
        $graphics.FillPath($backgroundBrush, $backgroundPath)

        $pageShadow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(62, 6, 20, 60))
        $shadow = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $shadow.AddPolygon([System.Drawing.Point[]]@(
            [System.Drawing.Point]::new(70, 51), [System.Drawing.Point]::new(162, 51),
            [System.Drawing.Point]::new(196, 85), [System.Drawing.Point]::new(196, 210),
            [System.Drawing.Point]::new(70, 210)))
        $graphics.FillPath($pageShadow, $shadow)

        $page = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $page.AddPolygon([System.Drawing.Point[]]@(
            [System.Drawing.Point]::new(62, 43), [System.Drawing.Point]::new(153, 43),
            [System.Drawing.Point]::new(187, 77), [System.Drawing.Point]::new(187, 202),
            [System.Drawing.Point]::new(62, 202)))
        $pageBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(250, 252, 255))
        $graphics.FillPath($pageBrush, $page)

        $fold = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $fold.AddPolygon([System.Drawing.Point[]]@(
            [System.Drawing.Point]::new(153, 43), [System.Drawing.Point]::new(153, 77),
            [System.Drawing.Point]::new(187, 77)))
        $foldBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(188, 215, 249))
        $graphics.FillPath($foldBrush, $fold)

        $linePen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(104, 146, 201), 7)
        $linePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
        $linePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $graphics.DrawLine($linePen, 82, 106, 159, 106)
        $graphics.DrawLine($linePen, 82, 126, 140, 126)

        $arrow = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $arrow.AddPolygon([System.Drawing.Point[]]@(
            [System.Drawing.Point]::new(80, 153), [System.Drawing.Point]::new(145, 153),
            [System.Drawing.Point]::new(145, 140), [System.Drawing.Point]::new(180, 166),
            [System.Drawing.Point]::new(145, 192), [System.Drawing.Point]::new(145, 179),
            [System.Drawing.Point]::new(80, 179)))
        $arrowBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(10, 196, 180))
        $graphics.FillPath($arrowBrush, $arrow)
        return $bitmap
    }
    finally {
        $graphics.Dispose()
        $backgroundPath.Dispose(); $backgroundBrush.Dispose()
        $pageShadow.Dispose(); $shadow.Dispose(); $page.Dispose(); $pageBrush.Dispose()
        $fold.Dispose(); $foldBrush.Dispose(); $linePen.Dispose(); $arrow.Dispose(); $arrowBrush.Dispose()
    }
}

$sizes = @(16, 24, 32, 48, 64, 128, 256)
$images = @()
foreach ($size in $sizes) {
    $bitmap = New-IconImage $size
    try {
        if ($size -eq 256) { $bitmap.Save($previewPath, [System.Drawing.Imaging.ImageFormat]::Png) }
        $stream = [System.IO.MemoryStream]::new()
        try {
            $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
            $images += ,$stream.ToArray()
        }
        finally { $stream.Dispose() }
    }
    finally { $bitmap.Dispose() }
}

$file = [System.IO.File]::Create($iconPath)
$writer = [System.IO.BinaryWriter]::new($file)
try {
    $writer.Write([uint16]0)
    $writer.Write([uint16]1)
    $writer.Write([uint16]$images.Count)
    $offset = 6 + 16 * $images.Count
    for ($index = 0; $index -lt $images.Count; $index++) {
        $size = $sizes[$index]
        $writer.Write([byte]($size % 256))
        $writer.Write([byte]($size % 256))
        $writer.Write([byte]0)
        $writer.Write([byte]0)
        $writer.Write([uint16]1)
        $writer.Write([uint16]32)
        $writer.Write([uint32]$images[$index].Length)
        $writer.Write([uint32]$offset)
        $offset += $images[$index].Length
    }
    foreach ($image in $images) { $writer.Write([byte[]]$image) }
}
finally { $writer.Dispose() }

Write-Host "Simge oluşturuldu: $iconPath"
