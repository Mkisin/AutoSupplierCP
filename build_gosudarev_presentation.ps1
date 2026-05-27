$OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = 'Stop'

$cwd = Get-Location
$template = Join-Path $cwd 'Шаблон ЦЗС.pptx'
$out = Join-Path $cwd 'Государев_стандарт_ЦЗС_переменные_v2.pptx'
Copy-Item -LiteralPath $template -Destination $out -Force

$payload = (python .\build_payload.py | ConvertFrom-Json)

$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1
$pres = $app.Presentations.Open($out, $false, $false, $false)

function Replace-InShape($shape, $map) {
    try {
        if ($shape.Type -eq 6) {
            for ($i=1; $i -le $shape.GroupItems.Count; $i++) {
                Replace-InShape $shape.GroupItems.Item($i) $map
            }
            return
        }
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            $text = $shape.TextFrame.TextRange.Text
            foreach ($property in $map.PSObject.Properties) {
                $text = $text.Replace($property.Name, [string]$property.Value)
            }
            $shape.TextFrame.TextRange.Text = $text
        }
    } catch {}
}

function Replace-AllText($presentation, $map) {
    for ($s=1; $s -le $presentation.Slides.Count; $s++) {
        $slide = $presentation.Slides.Item($s)
        for ($i=1; $i -le $slide.Shapes.Count; $i++) {
            Replace-InShape $slide.Shapes.Item($i) $map
        }
    }
}

function Find-PlaceholderShapes($shape, [string]$needle, $found) {
    try {
        if ($shape.Type -eq 6) {
            for ($i=1; $i -le $shape.GroupItems.Count; $i++) {
                Find-PlaceholderShapes $shape.GroupItems.Item($i) $needle $found
            }
            return
        }
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            if ($shape.TextFrame.TextRange.Text -like "*$needle*") {
                [void]$found.Add($shape)
            }
        }
    } catch {}
}

function Insert-ImageForPlaceholder($presentation, [string]$needle, [string]$imagePath) {
    if (-not (Test-Path -LiteralPath $imagePath)) { return 0 }
    $count = 0
    for ($s=1; $s -le $presentation.Slides.Count; $s++) {
        $slide = $presentation.Slides.Item($s)
        $found = New-Object System.Collections.ArrayList
        for ($i=1; $i -le $slide.Shapes.Count; $i++) {
            Find-PlaceholderShapes $slide.Shapes.Item($i) $needle $found
        }
        foreach ($shape in @($found)) {
            $left = $shape.Left
            $top = $shape.Top
            $width = $shape.Width
            $height = $shape.Height
            try { $shape.Delete() } catch {}
            [void]$slide.Shapes.AddPicture($imagePath, $false, $true, $left, $top, $width, $height)
            $count++
        }
    }
    return $count
}

Replace-AllText $pres $payload.replacements

$pic1 = (Get-ChildItem -File -Filter "*pic1*" | Select-Object -First 1).FullName
$photos = Get-ChildItem -File -LiteralPath "Фотки переговоров" | Where-Object { $_.Extension -match "\.(jpg|jpeg|png)$" } | Sort-Object Name
$logos = Get-ChildItem -File -LiteralPath "Логотипы поставщиков" | Where-Object { $_.Extension -match "\.(jpg|jpeg|png)$" } | Sort-Object Name

if ($pic1) { [void](Insert-ImageForPlaceholder $pres "{{pic1}}" $pic1) }
if ($photos.Count -ge 1) { [void](Insert-ImageForPlaceholder $pres "{{pic2}}" $photos[0].FullName) }
if ($photos.Count -ge 2) { [void](Insert-ImageForPlaceholder $pres "{{pic3}}" $photos[1].FullName) }
if ($photos.Count -ge 3) { [void](Insert-ImageForPlaceholder $pres "{{pic4}}" $photos[2].FullName) }
if ($logos.Count -ge 1) { [void](Insert-ImageForPlaceholder $pres "{{logo1}}" $logos[0].FullName) }

$pres.Save()
$pres.Close()
$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null

Write-Output "created=$out"
Write-Output "client=$($payload.company); contact=$($payload.contact); position=$($payload.position); category=$($payload.category); stats_source=$($payload.stats_source)"


