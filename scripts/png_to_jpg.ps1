param(
  [Parameter(Mandatory=$true)][string]$Src,
  [Parameter(Mandatory=$true)][string]$Dst
)

Add-Type -AssemblyName System.Drawing

$img = [System.Drawing.Image]::FromFile($Src)
try {
  $bmp = New-Object System.Drawing.Bitmap $img.Width, $img.Height
  try {
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    try {
      $graphics.Clear([System.Drawing.Color]::White)
      $graphics.DrawImage($img, 0, 0, $img.Width, $img.Height)
    } finally {
      $graphics.Dispose()
    }

    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
      Where-Object { $_.MimeType -eq 'image/jpeg' }
    $params = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
      [System.Drawing.Imaging.Encoder]::Quality,
      94L
    )
    $bmp.Save($Dst, $codec, $params)
  } finally {
    $bmp.Dispose()
  }
} finally {
  $img.Dispose()
}
