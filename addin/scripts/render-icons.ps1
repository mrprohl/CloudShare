# Renders the CloudShare add-in icon (addin/assets/icon.svg – an own design: a cloud with a padlock on a blue tile)
# as PNG in the sizes the manifest uses. Microsoft Edge draws the SVG once at 1024 px on a transparent background;
# System.Drawing scales it down in steps (halving, then the last step) for smooth edges.
#   powershell -File addin/scripts/render-icons.ps1 -OutDir addin/assets
# After new icons: raise the ?v= of the icon URLs in addin/manifest.xml (Outlook caches icons by URL).
param(
  [string]$OutDir = ".",
  [string]$Svg = (Join-Path $PSScriptRoot "..\assets\icon.svg"),
  [int[]]$Sizes = @(16, 32, 64, 80, 128)
)

Add-Type -AssemblyName System.Drawing
$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edge)) { throw "Microsoft Edge not found at $edge" }

$work = Join-Path ([IO.Path]::GetTempPath()) "cloudshare-icon-$(Get-Random)"
New-Item -ItemType Directory -Path $work | Out-Null
$master = 1024
$svgUrl = "file:///" + ((Resolve-Path $Svg).Path -replace '\\', '/')
$page = Join-Path $work "icon.html"
Set-Content -Path $page -Encoding UTF8 -Value "<!doctype html><html><body style=`"margin:0;background:transparent`"><img src=`"$svgUrl`" width=`"$master`" height=`"$master`" style=`"display:block`"></body></html>"

$big = Join-Path $work "icon-$master.png"
$opts = "--headless=new", "--hide-scrollbars", "--disable-gpu", "--no-first-run", "--allow-file-access-from-files",
        "--default-background-color=00000000", "--user-data-dir=$work\profile", "--window-size=$master,$master",
        "--screenshot=$big", ("file:///" + ($page -replace '\\', '/'))
$p = Start-Process $edge $opts -PassThru
if (-not $p.WaitForExit(60000)) { $p.Kill(); throw "Edge did not finish" }
if (-not (Test-Path $big)) { throw "Edge wrote no picture" }

function Scale([System.Drawing.Bitmap]$src, [int]$size) {
  $dst = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($dst)
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)
  $attr = New-Object System.Drawing.Imaging.ImageAttributes
  $attr.SetWrapMode([System.Drawing.Drawing2D.WrapMode]::TileFlipXY)   # no grey seam at the edges
  $g.DrawImage($src, (New-Object System.Drawing.Rectangle 0, 0, $size, $size), 0, 0, $src.Width, $src.Height, [System.Drawing.GraphicsUnit]::Pixel, $attr)
  $g.Dispose()
  $dst
}

$source = New-Object System.Drawing.Bitmap $big
foreach ($size in $Sizes) {
  $current = $source
  while ($current.Width / 2 -ge $size * 2) { $current = Scale $current ([int]($current.Width / 2)) }
  $out = Scale $current $size
  $file = Join-Path $OutDir "icon-$size.png"
  $out.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)
  $out.Dispose()
  Write-Output "wrote $file"
}
$source.Dispose()
Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue
