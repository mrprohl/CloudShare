# Renders the CloudShare add-in icon as PNG in the sizes the manifest uses: the Nextcloud mark, taken unchanged from
# the official logo (https://upload.wikimedia.org/wikipedia/commons/6/60/Nextcloud_Logo.svg, path "path1052" – the
# three rings without the word mark), with an upload arrow in the hole of the middle ring. Drawn at 8x and scaled down
# for smooth edges.
#   powershell -File addin/scripts/render-icons.ps1 -OutDir addin/assets
param([string]$OutDir = ".", [int[]]$Sizes = @(16, 32, 64, 80, 128))

Add-Type -AssemblyName System.Drawing

$blue = [System.Drawing.Color]::FromArgb(255, 0, 130, 201)   # #0082C9 (fill of the logo)
$dark = [System.Drawing.Color]::FromArgb(255, 66, 66, 66)    # #424242

# The rings exactly as in the logo (SVG path data, logo coordinates).
$rings = 'm 66.407896,9.375 c -11.805271,0 -21.811217,8.003196 -24.912392,18.846621 -2.695245,-5.751517 -8.535934,-9.780938 -15.263394,-9.780938 -9.25185,0 -16.85711,7.605263 -16.85711,16.857108 0,9.251833 7.60526,16.860567 16.85711,16.860567 6.72746,0 12.568149,-4.031885 15.263395,-9.784412 3.101175,10.84425 13.10712,18.850106 24.912391,18.850106 11.717964,0 21.67289,-7.885111 24.853382,-18.607048 2.745036,5.621934 8.513436,9.541354 15.145342,9.541354 9.25185,0 16.86057,-7.608734 16.86057,-16.860567 0,-9.251845 -7.60872,-16.857108 -16.86057,-16.857108 -6.631906,0 -12.400306,3.916965 -15.145342,9.537891 C 88.080786,17.257475 78.12586,9.375 66.407896,9.375 Z m 0,9.895518 c 8.911648,0 16.030748,7.115653 16.030748,16.027273 0,8.911605 -7.1191,16.030737 -16.030748,16.030737 -8.911593,0 -16.027247,-7.119132 -16.027247,-16.030737 0,-8.91162 7.115653,-16.027271 16.027247,-16.027273 z M 26.23211,28.336202 c 3.90438,0 6.96505,3.057188 6.96505,6.961589 0,3.904386 -3.06067,6.965049 -6.96505,6.965049 -3.90439,0 -6.96161,-3.060663 -6.96161,-6.965049 0,-3.904401 3.05722,-6.961589 6.96161,-6.961589 z m 80.17451,0 c 3.90442,0 6.96506,3.057188 6.96506,6.961589 0,3.904386 -3.06066,6.965049 -6.96506,6.965049 -3.90436,0 -6.961576,-3.060663 -6.961576,-6.965049 0,-3.904401 3.057226,-6.961589 6.961576,-6.961589 z'
$left = 9.375; $right = 123.263666; $top = 9.375; $bottom = 61.220669   # bounding box of the rings
$holeX = 66.407896; $holeY = 35.297791; $holeR = 16.027247                # hole of the middle ring

# Minimal SVG path reader for the commands the logo uses (M/m, C/c, Z/z) → GraphicsPath in logo coordinates.
function Read-SvgPath([string]$d) {
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath ([System.Drawing.Drawing2D.FillMode]::Winding)
  $tokens = [regex]::Matches($d, '[MmCcZz]|-?\d*\.?\d+(?:e-?\d+)?') | ForEach-Object { $_.Value }
  $i = 0; $cmd = ''; $x = 0.0; $y = 0.0; $startX = 0.0; $startY = 0.0
  while ($i -lt $tokens.Count) {
    if ($tokens[$i] -cmatch '^[MmCcZz]$') { $cmd = $tokens[$i]; $i++ }
    switch -CaseSensitive ($cmd) {
      { $_ -in 'M', 'm' } {
        $nx = [double]$tokens[$i]; $ny = [double]$tokens[$i + 1]; $i += 2
        if ($cmd -ceq 'm') { $nx += $x; $ny += $y }
        $path.StartFigure(); $x = $nx; $y = $ny; $startX = $x; $startY = $y
        $cmd = if ($cmd -ceq 'm') { 'c' } else { 'C' }   # the logo only follows a move with curves
      }
      { $_ -in 'C', 'c' } {
        $p = 0..5 | ForEach-Object { [double]$tokens[$i + $_] }; $i += 6
        if ($cmd -ceq 'c') { $p = @(($p[0] + $x), ($p[1] + $y), ($p[2] + $x), ($p[3] + $y), ($p[4] + $x), ($p[5] + $y)) }
        $path.AddBezier([single]$x, [single]$y, [single]$p[0], [single]$p[1], [single]$p[2], [single]$p[3], [single]$p[4], [single]$p[5])
        $x = $p[4]; $y = $p[5]
      }
      { $_ -in 'Z', 'z' } { $path.CloseFigure(); $x = $startX; $y = $startY; $cmd = '' }
    }
  }
  $path
}

foreach ($size in $Sizes) {
  $scale = 8
  $n = $size * $scale
  $big = New-Object System.Drawing.Bitmap $n, $n
  $g = [System.Drawing.Graphics]::FromImage($big)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.Color]::Transparent)

  # The rings fill the width and sit in the vertical middle.
  $f = $n / ($right - $left)
  $g.TranslateTransform([single](-$left * $f), [single](($n - ($bottom - $top) * $f) / 2 - $top * $f))
  $g.ScaleTransform([single]$f, [single]$f)
  $g.FillPath((New-Object System.Drawing.SolidBrush $blue), (Read-SvgPath $rings))

  # Upload arrow in the hole of the middle ring (shape in units where it is 38 high incl. its round caps).
  $k = 0.86 * $holeR / 19
  # At 16 px the arrow would drop below 1.2 px – keep it legible.
  $arrow = New-Object System.Drawing.Pen $dark, ([Math]::Max(9.5 * $k, 1.2 * $scale / $f))
  $arrow.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
  $arrow.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
  $arrow.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
  $tip = [System.Drawing.PointF]::new($holeX, $holeY - 14.5 * $k)
  $g.DrawLine($arrow, $tip, [System.Drawing.PointF]::new($holeX, $holeY + 14.5 * $k))
  $g.DrawLines($arrow, [System.Drawing.PointF[]]@(
    [System.Drawing.PointF]::new($holeX - 11.5 * $k, $holeY - 2.5 * $k), $tip,
    [System.Drawing.PointF]::new($holeX + 11.5 * $k, $holeY - 2.5 * $k)))
  $g.Dispose()

  $small = New-Object System.Drawing.Bitmap $size, $size
  $gs = [System.Drawing.Graphics]::FromImage($small)
  $gs.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $gs.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $gs.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $gs.DrawImage($big, 0, 0, $size, $size)
  $gs.Dispose()
  $file = Join-Path $OutDir "icon-$size.png"
  $small.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)
  $small.Dispose(); $big.Dispose()
  "written $file"
}
