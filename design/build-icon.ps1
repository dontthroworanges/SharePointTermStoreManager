<#
.SYNOPSIS
    Builds the app icon (Assets\AppIcon.png and Assets\AppIcon.ico) from design\AppIcon.svg.
.DESCRIPTION
    Renders the SVG at 512 px with headless Microsoft Edge (using a throwaway browser profile), then scales it
    down with high-quality resampling to each icon size. The .ico holds 16-256 px images as PNGs, which Windows
    Vista and later read.
.EXAMPLE
    pwsh design\build-icon.ps1
    pwsh design\build-icon.ps1 -OutDir $env:TEMP\icon-preview
#>
param(
    [string]$Svg = (Join-Path $PSScriptRoot 'AppIcon.svg'),
    [string]$OutDir = (Join-Path $PSScriptRoot '..\SharePointTermStoreManager\Assets'),
    [string]$Edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$work = Join-Path ([IO.Path]::GetTempPath()) ("appicon-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
try {
    # 1. Render the SVG at 512 px on a transparent page.
    $master = Join-Path $work 'master.png'
    $svgUri = ([Uri](Resolve-Path $Svg).Path).AbsoluteUri
    $html = Join-Path $work 'icon.html'
    Set-Content $html "<!DOCTYPE html><html><body style='margin:0;background:transparent'><img src='$svgUri' width='512' height='512' style='display:block'></body></html>"
    $args = @(
        '--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run',
        "--user-data-dir=$(Join-Path $work 'profile')",
        '--default-background-color=00000000', '--window-size=512,512',
        "--screenshot=$master", ([Uri]$html).AbsoluteUri
    )
    # Headless Edge sometimes writes the screenshot but leaves its helper processes running, so wait for the
    # file rather than for Edge to exit, then close Edge's whole process tree. Try a few times.
    for ($attempt = 1; $attempt -le 3 -and -not (Test-Path $master); $attempt++) {
        $edgeProcess = Start-Process -FilePath $Edge -ArgumentList $args -PassThru -WindowStyle Hidden
        $deadline = (Get-Date).AddSeconds(45)
        while ((Get-Date) -lt $deadline -and -not ((Test-Path $master) -and (Get-Item $master).Length -gt 0)) {
            Start-Sleep -Milliseconds 250
        }
        Start-Sleep -Milliseconds 500 # let Edge finish writing
        & taskkill.exe /PID $edgeProcess.Id /T /F 2>&1 | Out-Null
    }
    if (-not (Test-Path $master)) { throw 'Edge did not render the icon (tried 3 times).' }

    # 2. Scale to each size.
    $sizes = 16, 20, 24, 32, 40, 48, 64, 128, 256
    $source = [System.Drawing.Image]::FromFile($master)
    $pngs = @{}
    foreach ($size in $sizes) {
        $bitmap = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [System.Drawing.Graphics]::FromImage($bitmap)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.DrawImage($source, 0, 0, $size, $size)
        $g.Dispose()
        $stream = New-Object IO.MemoryStream
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        $pngs[$size] = $stream.ToArray()
        $bitmap.Dispose()
    }
    $source.Dispose()

    # 3. AppIcon.png (64 px, used in the title bar) and AppIcon.ico (all sizes).
    [IO.File]::WriteAllBytes((Join-Path $OutDir 'AppIcon.png'), $pngs[64])

    $ico = New-Object IO.MemoryStream
    $writer = New-Object IO.BinaryWriter $ico
    $writer.Write([UInt16]0); $writer.Write([UInt16]1); $writer.Write([UInt16]$sizes.Count)
    $offset = 6 + 16 * $sizes.Count
    foreach ($size in $sizes) {
        $dim = if ($size -ge 256) { 0 } else { $size }
        $writer.Write([byte]$dim); $writer.Write([byte]$dim)
        $writer.Write([byte]0); $writer.Write([byte]0)
        $writer.Write([UInt16]1); $writer.Write([UInt16]32)
        $writer.Write([UInt32]$pngs[$size].Length); $writer.Write([UInt32]$offset)
        $offset += $pngs[$size].Length
    }
    foreach ($size in $sizes) { $writer.Write($pngs[$size]) }
    $writer.Flush()
    [IO.File]::WriteAllBytes((Join-Path $OutDir 'AppIcon.ico'), $ico.ToArray())
    $writer.Dispose()

    Copy-Item $master (Join-Path $OutDir 'AppIcon-512-preview.png') -ErrorAction SilentlyContinue
    if ((Resolve-Path $OutDir).Path -like '*\Assets') { Remove-Item (Join-Path $OutDir 'AppIcon-512-preview.png') -ErrorAction SilentlyContinue }
    "Wrote AppIcon.png and AppIcon.ico ($($sizes -join ', ') px) to $((Resolve-Path $OutDir).Path)"
}
finally {
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}
