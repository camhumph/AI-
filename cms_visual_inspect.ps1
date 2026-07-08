param(
    [Parameter(Mandatory=$true)][string]$JobFolder,
    [string]$ImagePath = "",
    [string]$CadCsv = ""
)

$ErrorActionPreference = "Stop"

function Clean-Text([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace '[\r\n,]+',' ' -replace '\s+',' ').Trim()
}

function Find-FirstFile([string]$folder, [string[]]$patterns) {
    foreach ($pat in $patterns) {
        $f = Get-ChildItem -LiteralPath $folder -File -Filter $pat -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notmatch '(?i)\bBACK\b' } |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1
        if ($f) { return $f.FullName }
    }
    foreach ($pat in $patterns) {
        $f = Get-ChildItem -LiteralPath $folder -File -Filter $pat -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1
        if ($f) { return $f.FullName }
    }
    return ""
}

function Get-StdRole([int]$pos, [int]$count, [string]$name) {
    $u = (" " + $name.ToUpperInvariant() + " ")
    if ($u -match '\bTCP\b|TOP CLAMP') { return "Top Clamp Plate" }
    if ($u -match '\bBCP\b|BOTTOM CLAMP|BOT CLAMP') { return "Bottom Clamp Plate" }
    if ($u -match 'CAVITY') { return "Cavity Plate" }
    if ($u -match 'CORE') { return "Core Plate" }
    if ($u -match 'STRIPPER') { return "Stripper Plate" }
    if ($u -match 'DIE BACK') { return "Die Backup Plate" }
    if ($u -match 'DIE') { return "Die Plate" }
    if ($u -match 'SUPPORT') { return "Support Plate" }

    if ($pos -eq 1) { return "Top Clamp Plate" }
    if ($pos -eq $count) { return "Bottom Clamp Plate" }
    switch ($count) {
        3 { if ($pos -eq 2) { return "Stripper Plate" } }
        4 {
            if ($pos -eq 2) { return "Cavity Plate" }
            if ($pos -eq 3) { return "Core Plate" }
        }
        5 {
            if ($pos -eq 2) { return "Cavity Plate" }
            if ($pos -eq 3) { return "Core Plate" }
            if ($pos -eq 4) { return "Support Plate" }
        }
        6 {
            if ($pos -eq 2) { return "Cavity Plate" }
            if ($pos -eq 3) { return "Stripper Plate" }
            if ($pos -eq 4) { return "Core Plate" }
            if ($pos -eq 5) { return "Support Plate" }
        }
        7 {
            if ($pos -eq 2) { return "Cavity Plate" }
            if ($pos -eq 3) { return "Stripper Plate" }
            if ($pos -eq 4) { return "Core Plate" }
            if ($pos -eq 5) { return "Die Plate" }
            if ($pos -eq 6) { return "Die Backup Plate" }
        }
    }
    return "Plate $pos"
}

function Get-PcsSeriesGuess([object[]]$roleRows, [object[]]$allParts, [int]$fullCount, [int]$visualBands, [bool]$potBlockLike) {
    $roles = @($roleRows | ForEach-Object { [string]$_.Role })
    $roleText = (" " + ($roles -join " ") + " ").ToUpperInvariant()
    $compText = (" " + (($allParts | ForEach-Object { [string]$_.Component }) -join " ") + " ").ToUpperInvariant()

    if ($potBlockLike) {
        return "Not one of the six PCS standard series - BMS/pot-block style"
    }

    $hasStripper = ($roleText -match ' STRIPPER ' -or $compText -match 'STRIPPER')
    $hasDieBackup = ($roleText -match 'DIE BACKUP' -or $compText -match 'DIE BACK')
    $hasDie = ($roleText -match ' DIE ' -or $compText -match 'DIE PLATE|DIE PLT')
    $hasX = ($roleText -match ' X PLATE ' -or $compText -match '\bX\s*(PLATE|PLT)\b')
    $hasAcp = ($compText -match '\bACP\b|A CLAMPING|A-CLAMPING')
    $hasTopClamp = ($roleText -match 'TOP CLAMP')

    if ($hasStripper) {
        if ($fullCount -ge 7 -or $hasDieBackup) { return "PCS 6 Plate Stripper Series" }
        if ($fullCount -ge 5 -or $visualBands -ge 5) { return "PCS 5 Plate Stripper Series" }
        return "PCS T Series / stripper-style"
    }

    if ($hasX -or $hasDie) { return "PCS AX Series" }
    if ($hasAcp -or (-not $hasTopClamp -and $fullCount -le 4)) { return "PCS B Series" }
    if ($fullCount -ge 4) { return "PCS A Series" }
    if ($fullCount -ge 3 -and $visualBands -ge 3) { return "PCS T Series" }

    return "Unknown PCS series"
}

function Analyze-JpegBands([string]$path) {
    $result = [ordered]@{
        ImageWidth = 0
        ImageHeight = 0
        ObjectTop = 0
        ObjectBottom = 0
        VisibleBandCount = 0
        BandRows = @()
        Notes = @()
    }
    if (-not $path -or -not (Test-Path -LiteralPath $path)) {
        $result.Notes += "No JPEG available."
        return [pscustomobject]$result
    }

    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($path)
    try {
        $w = $bmp.Width
        $h = $bmp.Height
        $result.ImageWidth = $w
        $result.ImageHeight = $h

        $x1 = [Math]::Max(0, [int]($w * 0.22))
        $x2 = [Math]::Min($w - 1, [int]($w * 0.78))
        $stepX = [Math]::Max(1, [int](($x2 - $x1) / 90))
        $stepY = [Math]::Max(1, [int]($h / 700))

        $rows = New-Object System.Collections.Generic.List[object]
        for ($y = 0; $y -lt $h; $y += $stepY) {
            $n = 0
            $r = 0.0; $g = 0.0; $b = 0.0
            for ($x = $x1; $x -le $x2; $x += $stepX) {
                $c = $bmp.GetPixel($x, $y)
                $maxc = [Math]::Max($c.R, [Math]::Max($c.G, $c.B))
                $minc = [Math]::Min($c.R, [Math]::Min($c.G, $c.B))
                $isBg = ($c.R -gt 238 -and $c.G -gt 238 -and $c.B -gt 238) -or (($maxc - $minc) -lt 8 -and $maxc -gt 225)
                if (-not $isBg) {
                    $n++
                    $r += $c.R; $g += $c.G; $b += $c.B
                }
            }
            if ($n -gt 0) {
                $rows.Add([pscustomobject]@{
                    Y = $y
                    Coverage = $n
                    R = $r / $n
                    G = $g / $n
                    B = $b / $n
                })
            }
        }

        if ($rows.Count -lt 5) {
            $result.Notes += "JPEG object coverage was too low for reliable band detection."
            return [pscustomobject]$result
        }

        $covMax = ($rows | Measure-Object Coverage -Maximum).Maximum
        $objRows = @($rows | Where-Object { $_.Coverage -ge [Math]::Max(3, $covMax * 0.18) })
        if ($objRows.Count -lt 5) {
            $result.Notes += "JPEG object region was too sparse for reliable band detection."
            return [pscustomobject]$result
        }

        $top = ($objRows | Select-Object -First 1).Y
        $bottom = ($objRows | Select-Object -Last 1).Y
        $result.ObjectTop = $top
        $result.ObjectBottom = $bottom

        $edges = New-Object System.Collections.Generic.List[int]
        for ($i = 1; $i -lt $objRows.Count; $i++) {
            $a = $objRows[$i-1]
            $brow = $objRows[$i]
            $dr = [Math]::Abs($a.R - $brow.R)
            $dg = [Math]::Abs($a.G - $brow.G)
            $db = [Math]::Abs($a.B - $brow.B)
            $dc = [Math]::Sqrt(($dr*$dr)+($dg*$dg)+($db*$db))
            $dcov = [Math]::Abs($a.Coverage - $brow.Coverage)
            if ($dc -gt 18 -or $dcov -gt ($covMax * 0.18)) {
                $edges.Add([int](($a.Y + $brow.Y) / 2))
            }
        }

        $merged = New-Object System.Collections.Generic.List[int]
        foreach ($e in $edges) {
            if ($merged.Count -eq 0 -or [Math]::Abs($e - $merged[$merged.Count-1]) -gt [Math]::Max(8, $h * 0.015)) {
                $merged.Add($e)
            }
        }

        $bandCount = [Math]::Max(1, $merged.Count + 1)
        $result.VisibleBandCount = $bandCount
        $result.BandRows = @($merged)
        if ($bandCount -gt 12) {
            $result.Notes += "High band count likely includes holes/straps/shadows; use as a warning, not exact plate count."
        }
    }
    finally {
        $bmp.Dispose()
    }
    return [pscustomobject]$result
}

if (-not (Test-Path -LiteralPath $JobFolder)) {
    throw "Job folder not found: $JobFolder"
}
if (-not $CadCsv) { $CadCsv = Join-Path $JobFolder "XT_Export_CAD_Dimensions.csv" }
if (-not $ImagePath) {
    $ImagePath = Find-FirstFile $JobFolder @("* ISO.jpg", "*ISO.jpg", "*.jpg", "*.jpeg")
}

$parts = @()
if (Test-Path -LiteralPath $CadCsv) {
    $parts = @(Import-Csv -LiteralPath $CadCsv)
}

$totalQty = 0
foreach ($p in $parts) {
    $q = 1
    [void][int]::TryParse([string]$p.Qty, [ref]$q)
    if ($q -lt 1) { $q = 1 }
    $totalQty += $q
}

$maxFoot = 0.0
foreach ($p in $parts) {
    $w = [double]$p.Width
    $l = [double]$p.Length
    $foot = $w * $l
    if ($foot -gt $maxFoot) { $maxFoot = $foot }
}

$full = @()
$smallHardware = 0
$midLargeCount = 0
$rodLikeCount = 0
foreach ($p in $parts) {
    $t = [double]$p.Thickness
    $w = [double]$p.Width
    $l = [double]$p.Length
    $foot = $w * $l
    if ($t -ge 0.4 -and $maxFoot -gt 0 -and $foot -ge $maxFoot * 0.82) {
        $full += $p
    }
    if ($foot -lt 20 -or ($w -lt 2 -and $l -lt 4)) {
        $smallHardware++
    }
    if ($maxFoot -gt 0 -and $t -ge 1.0 -and $foot -ge ($maxFoot * 0.08) -and $foot -lt ($maxFoot * 0.82)) {
        $midLargeCount++
    }
    if ($t -ge 0.5 -and [Math]::Abs($w - $t) -lt 0.2 -and $l -gt ($w * 4.0)) {
        $rodLikeCount++
    }
}

$axis = "CenterZ"
if ($full.Count -ge 2) {
    $ranges = @{}
    foreach ($a in "CenterX","CenterY","CenterZ") {
        $vals = @($full | ForEach-Object { [double]$_.$a })
        $ranges[$a] = (($vals | Measure-Object -Maximum).Maximum - ($vals | Measure-Object -Minimum).Minimum)
    }
    $axis = ($ranges.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
}

$orderedFull = @($full | Sort-Object { -[double]$_.$axis })
if ($orderedFull.Count -ge 2) {
    $first = $orderedFull[0]
    $last = $orderedFull[$orderedFull.Count-1]
    if ([double]$first.Thickness -gt ([double]$last.Thickness + 0.25)) {
        [array]::Reverse($orderedFull)
    }
}

$names = @()
for ($i = 0; $i -lt $orderedFull.Count; $i++) {
    $p = $orderedFull[$i]
    $role = Get-StdRole ($i+1) $orderedFull.Count ([string]$p.Component)
    $names += [pscustomobject]@{
        Role = $role
        Component = [string]$p.Component
        Qty = [string]$p.Qty
        Thickness = [string]$p.Thickness
        Width = [string]$p.Width
        Length = [string]$p.Length
        StackPosition = $i + 1
        Evidence = "full-footprint CAD + stack order"
        Confidence = if ($p.Component -match '(?i)TCP|BCP|CAVITY|CORE|STRIPPER|DIE|CLAMP') { "High" } else { "Medium" }
    }
}

$visual = Analyze-JpegBands $ImagePath
$namedPotSignals = @($parts | Where-Object { $_.Component -match '(?i)POT|HOLDER|SMED' }).Count
$potBlockLike = ($namedPotSignals -ge 2) -or `
                ($orderedFull.Count -eq 2 -and $midLargeCount -ge 2 -and $visual.VisibleBandCount -ge 3) -or `
                ($orderedFull.Count -le 2 -and $visual.VisibleBandCount -ge 4 -and ($rodLikeCount -ge 3 -or $midLargeCount -ge 1))
$type = if ($orderedFull.Count -ge 3) { "Standard mold base" } else { "Pot/block or non-standard base" }
if ($potBlockLike) {
    $type = "BMS/pot-block style"
}
$pcsSeries = Get-PcsSeriesGuess $names $parts $orderedFull.Count $visual.VisibleBandCount $potBlockLike

$parting = "Unknown"
if ($orderedFull.Count -ge 4) {
    $cavity = $names | Where-Object Role -eq "Cavity Plate" | Select-Object -First 1
    $core = $names | Where-Object Role -eq "Core Plate" | Select-Object -First 1
    $strip = $names | Where-Object Role -eq "Stripper Plate" | Select-Object -First 1
    if ($cavity -and $core) { $parting = "Between Cavity Plate and Core Plate area" }
    elseif ($cavity -and $strip) { $parting = "Between Cavity Plate and Stripper Plate area" }
} elseif ($potBlockLike) {
    $parting = "Around the middle split between upper holder/pot blocks and lower holder/pot blocks"
}

$warnings = New-Object System.Collections.Generic.List[string]
if ($visual.VisibleBandCount -gt 0 -and $orderedFull.Count -gt 0) {
    $diff = [Math]::Abs($visual.VisibleBandCount - $orderedFull.Count)
    if ($diff -ge 3) {
        $warnings.Add("JPEG visible layer count ($($visual.VisibleBandCount)) differs from CAD full-plate count ($($orderedFull.Count)). Check straps/shadows/hidden plates.")
    }
}
if ($rodLikeCount -ge 3) {
    $warnings.Add("Several near-square long parts detected; likely rods/pins/pillars, not rails.")
}
if ($pcsSeries -eq "Unknown PCS series") {
    $warnings.Add("PCS six-series classifier could not choose A/B/T/AX/5 Plate Stripper/6 Plate Stripper from this view; review CAD/BOM.")
}
foreach ($n in $visual.Notes) { $warnings.Add($n) }

$csvOut = Join-Path $JobFolder "Visual_Mold_Inspection.csv"
$txtOut = Join-Path $JobFolder "Visual_Mold_Inspection.txt"

$names | Export-Csv -LiteralPath $csvOut -NoTypeInformation -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("VISUAL MOLD INSPECTION")
$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$lines.Add("Job folder: $JobFolder")
$lines.Add("Image: $ImagePath")
$lines.Add("CAD CSV: $CadCsv")
$lines.Add("")
$lines.Add("Summary")
$lines.Add("  Mold type guess: $type")
$lines.Add("  PCS six-series guess: $pcsSeries")
$lines.Add("  CAD component rows: $($parts.Count)")
$lines.Add("  CAD quantity total: $totalQty")
$lines.Add("  CAD full-footprint plate count: $($orderedFull.Count)")
$lines.Add("  JPEG visible layer/band estimate: $($visual.VisibleBandCount)")
$lines.Add("  Parting line guess: $parting")
$lines.Add("  PCS series supported: A Series; B Series; T Series; AX Series; 5 Plate Stripper Series; 6 Plate Stripper Series")
$lines.Add("")
$lines.Add("Likely Outside/Stack Parts")
foreach ($n in $names) {
    $lines.Add(("  {0}. {1}: {2}  Qty {3}  {4} x {5} x {6}  Confidence={7}" -f $n.StackPosition,$n.Role,(Clean-Text $n.Component),$n.Qty,$n.Thickness,$n.Width,$n.Length,$n.Confidence))
}
$lines.Add("")
$lines.Add("Warnings")
if ($warnings.Count -eq 0) {
    $lines.Add("  None")
} else {
    foreach ($w in $warnings) { $lines.Add("  - $w") }
}
[System.IO.File]::WriteAllLines($txtOut, $lines, [System.Text.Encoding]::UTF8)

Write-Output "Visual inspection written:"
Write-Output "  $txtOut"
Write-Output "  $csvOut"
