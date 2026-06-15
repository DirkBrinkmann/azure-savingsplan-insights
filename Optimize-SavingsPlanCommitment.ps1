<#
.SYNOPSIS
    Azure Savings Plan (SP) commitment optimizer.

.DESCRIPTION
    Scans a folder (or a single file) for Azure hourly PayGo usage CSV
    exports (the *HourlyUsage.csv files produced by
    Export-BenefitRecommendations.ps1) and computes, per file:

      (a) The maximum hourly SP commitment that produces ZERO unused SP
          across the observed period (K_max_no_waste = min hourly usage).
      (b) The OPTIMUM hourly SP commitment that maximises savings vs.
          pure PayGo for each requested average SP discount
          (default 30 / 35 / 40 %).

    Results are reported as per-month savings/waste and whole-term
    (1- or 3-year) projections, with the term taken from the filename
    (`...P1Y...` or `...P3Y...`). Per-file locale (de-DE vs en-US) is
    auto-detected for both date format and decimal separator.

    The combined Markdown report groups files by lookback window
    (`## 30 day recommendation` / `## 60 day recommendation`), with one
    sub-section per file headed by the export date and the billing scope
    parsed from the filename pattern:

        {yyyy-MM-dd}-{EA|MCA|SUB|RG}-{id1}[-{id2}]-Last{30|60}Days-P{1|3}Y-HourlyUsage.csv

    Commitment K is expressed in PayGo-equivalent USD/hour, directly
    comparable to the hourly usage column in the CSV.

.PARAMETER Path
    CSV file or folder containing *HourlyUsage.csv files.
    Default: 'HourlyUsage' subfolder next to this script.

.PARAMETER Discounts
    Comma-separated *average* SP discounts to model.

    The default range `30,32,34,36,38,40,42` reflects a span of *common
    average* Savings Plan discounts seen in practice — from conservative
    (~30 %) to optimistic (~42 %). **Actual SP discounts vary by SKU**,
    region, and reservation footprint, so the report shows the
    recommended commitment and projected savings/waste across the entire
    range rather than a single point estimate.

    Whole percent (`30`) and fractions (`0.30`) are both accepted.

.PARAMETER Out
    Output markdown file. When omitted, the filename is auto-derived from
    the input metadata in the form
    `{yyyy-MM-dd}-{scope}-Last{N}Days-P{T}Y-SavingsPlanRecommendation.md`
    (mixed inputs collapse shared tokens; differing tokens become
    `Multi` / joined with `+`).

.EXAMPLE
    .\Optimize-SavingsPlanCommitment.ps1
    .\Optimize-SavingsPlanCommitment.ps1 -Path .\HourlyUsage
    .\Optimize-SavingsPlanCommitment.ps1 -Path . -Discounts 20,22,24,30,35,40
    .\Optimize-SavingsPlanCommitment.ps1 -Path .\HourlyUsage\2026-06-11-SUB-...-Last30Days-P1Y-HourlyUsage.csv

.NOTES
    Version: v3.2026-06-15.0
    Author : Dirk Brinkmann
#>

[CmdletBinding()]
param(
    [string]$Path = $(Join-Path $PSScriptRoot 'HourlyUsage'),
    # Default range = common average SP discounts, conservative (~30%) -> optimistic (~42%).
    # Real per-SKU discounts vary; the range gives a sensitivity band rather than a single point.
    [string]$Discounts = '30,32,34,36,38,40,42',
    [string]$Out
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$HoursPerMonth = 8766 / 12          # 730.5
$HoursPerYear  = 8766                # average Gregorian year

$FullPattern = '^(?<date>\d{4}-\d{2}-\d{2})-(?<kind>EA|MCA|SUB|RG)-(?<ids>.+?)-Last(?<days>\d+)Days-P(?<term>\d+)Y-HourlyUsage\.csv$'
$TailPattern = 'Last(?<days>\d+)Days-P(?<term>\d+)Y-HourlyUsage\.csv$'

$ScopeKindLabels = @{
    'EA'  = 'Billing Account'
    'MCA' = 'Billing Account / Billing Profile'
    'SUB' = 'Subscription'
    'RG'  = 'Subscription / Resource Group'
}

$DeDateFormats = @('dd.MM.yyyy HH:mm', 'dd.MM.yyyy HH:mm:ss')
$UsDateFormats = @(
    'MM/dd/yyyy HH:mm', 'MM/dd/yyyy h:mm tt', 'MM/dd/yyyy HH:mm:ss',
    'yyyy-MM-dd HH:mm', 'yyyy-MM-ddTHH:mm',
    'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss'
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

$Inv = [System.Globalization.CultureInfo]::InvariantCulture

function Format-Money {
    param([double]$Value)
    if ($Value -lt 0) { return ('-$' + ([math]::Abs($Value)).ToString('N2', $Inv)) }
    return ('$' + $Value.ToString('N2', $Inv))
}

function Format-Pct {
    param([double]$Value)
    return ([int][math]::Round($Value * 100)).ToString($Inv) + '%'
}

function Format-Int {
    param([double]$Value)
    return $Value.ToString('N0', $Inv)
}

function ConvertTo-DiscountList {
    param([string]$Text)
    $out = New-Object System.Collections.Generic.List[double]
    foreach ($tok in ($Text -split ',')) {
        $t = $tok.Trim().TrimEnd('%')
        if ([string]::IsNullOrEmpty($t)) { continue }
        $v = [double]::Parse($t, [System.Globalization.CultureInfo]::InvariantCulture)
        if ($v -gt 1) { $v = $v / 100.0 }
        if ($v -le 0 -or $v -ge 1) {
            throw "Discount '$tok' out of range (0,1)."
        }
        $out.Add($v)
    }
    if ($out.Count -eq 0) { throw 'No valid discounts.' }
    return ($out | Sort-Object)
}

function Get-Quantile {
    # Linear interpolation between order statistics — matches numpy default
    # and the Python script's `quantile()` helper.
    param(
        [double[]]$SortedValues,
        [double]$Q
    )
    if (-not $SortedValues -or $SortedValues.Count -eq 0) { throw 'Empty input.' }
    if ($Q -le 0) { return $SortedValues[0] }
    if ($Q -ge 1) { return $SortedValues[$SortedValues.Count - 1] }
    $pos = $Q * ($SortedValues.Count - 1)
    $lo = [math]::Floor($pos)
    $frac = $pos - $lo
    if ($lo + 1 -ge $SortedValues.Count) { return $SortedValues[$lo] }
    return $SortedValues[$lo] + $frac * ($SortedValues[$lo + 1] - $SortedValues[$lo])
}

function Get-CostWithSp {
    param(
        [double[]]$Usage,
        [double]$K,
        [double]$D
    )
    $overflow = 0.0
    foreach ($u in $Usage) {
        if ($u -gt $K) { $overflow += ($u - $K) }
    }
    return $Usage.Count * $K * (1.0 - $D) + $overflow
}

function Get-WasteDollars {
    param(
        [double[]]$Usage,
        [double]$K,
        [double]$D
    )
    $waste = 0.0
    foreach ($u in $Usage) {
        if ($u -lt $K) { $waste += ($K - $u) }
    }
    return $waste * (1.0 - $D)
}

# ---------------------------------------------------------------------------
# File parsing
# ---------------------------------------------------------------------------

function Get-FilenameMeta {
    param([string]$FileName)
    $meta = [ordered]@{
        ExportDate   = $null
        ScopeKind    = $null
        ScopeIds     = $null
        LookbackDays = $null
        TermYears    = $null
    }
    $m = [regex]::Match($FileName, $FullPattern, 'IgnoreCase')
    if ($m.Success) {
        $meta.ExportDate   = $m.Groups['date'].Value
        $meta.ScopeKind    = $m.Groups['kind'].Value.ToUpper()
        $meta.ScopeIds     = $m.Groups['ids'].Value
        $meta.LookbackDays = [int]$m.Groups['days'].Value
        $meta.TermYears    = [int]$m.Groups['term'].Value
        return $meta
    }
    $m = [regex]::Match($FileName, $TailPattern, 'IgnoreCase')
    if ($m.Success) {
        $meta.LookbackDays = [int]$m.Groups['days'].Value
        $meta.TermYears    = [int]$m.Groups['term'].Value
    }
    return $meta
}

function Get-LocaleAndDelimiter {
    param([string]$FilePath)

    $reader = [System.IO.StreamReader]::new($FilePath, [System.Text.Encoding]::UTF8, $true)
    try {
        $headerLine = $null
        $dataLine = $null
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($null -eq $headerLine) { $headerLine = $line; continue }
            $dataLine = $line
            break
        }
    } finally {
        $reader.Dispose()
    }

    if (-not $headerLine -or -not $dataLine) {
        throw "$FilePath`: cannot read header + first data row."
    }

    # Delimiter sniff: pick whichever of ';' or ',' is more frequent on header.
    $delimiter = if ($headerLine.Split(';').Count -ge $headerLine.Split(',').Count -and $headerLine.Contains(';')) {
        ';'
    } else {
        ','
    }

    $cells = $dataLine.Split($delimiter)
    if ($cells.Count -lt 2) {
        throw "$FilePath`: first data row has <2 cells with delimiter '$delimiter'."
    }
    $dateCell = $cells[0].Trim().Trim('"')
    $valCell  = $cells[1].Trim().Trim('"')

    # Locale heuristic (decimal separator first, then date pattern).
    $isDe = ($valCell.Contains(',') -and -not $valCell.Contains('.'))
    $isUs = ($valCell.Contains('.') -and -not $valCell.Contains(','))
    if (-not ($isDe -or $isUs)) {
        $isDe = ($dateCell -match '^\d{1,2}\.\d{1,2}\.\d{4}')
        $isUs = -not $isDe
    }

    $primary  = if ($isDe) { $DeDateFormats } else { $UsDateFormats }
    $fallback = if ($isDe) { $UsDateFormats } else { $DeDateFormats }

    $dateFormat = $null
    foreach ($fmt in $primary) {
        $tmp = [datetime]::MinValue
        if ([datetime]::TryParseExact($dateCell, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$tmp)) {
            $dateFormat = $fmt; break
        }
    }
    if (-not $dateFormat) {
        foreach ($fmt in $fallback) {
            $tmp = [datetime]::MinValue
            if ([datetime]::TryParseExact($dateCell, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$tmp)) {
                $dateFormat = $fmt
                $isDe = -not $isDe
                $isUs = -not $isUs
                break
            }
        }
    }
    if (-not $dateFormat) {
        throw "$FilePath`: unrecognised date format '$dateCell'."
    }

    return [pscustomobject]@{
        Locale     = if ($isDe) { 'de-DE' } else { 'en-US' }
        Delimiter  = $delimiter
        DateFormat = $dateFormat
    }
}

function Get-UsageData {
    param([pscustomobject]$Meta)

    $rows = New-Object System.Collections.Generic.List[object]
    $decimalComma = ($Meta.Locale -eq 'de-DE')
    $reader = [System.IO.StreamReader]::new($Meta.Path, [System.Text.Encoding]::UTF8, $true)
    try {
        $headerSeen = $false
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if (-not $headerSeen) {
                $headerSeen = $true
                $headerCells = $line.Split($Meta.Delimiter) | ForEach-Object { $_.Trim().Trim('"').ToLower() }
                if ($headerCells.Count -lt 2 -or $headerCells[0] -ne 'datetime' -or $headerCells[1] -ne 'hourlypaygousage') {
                    throw "$($Meta.Path): unexpected header '$line'."
                }
                continue
            }
            $cells = $line.Split($Meta.Delimiter)
            if ($cells.Count -lt 2) { continue }
            $dt = [datetime]::ParseExact($cells[0].Trim().Trim('"'), $Meta.DateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
            $v = $cells[1].Trim().Trim('"')
            if ($decimalComma) {
                $v = $v.Replace('.', '').Replace(',', '.')
            }
            $val = [double]::Parse($v, [System.Globalization.CultureInfo]::InvariantCulture)
            $rows.Add([pscustomobject]@{ Time = $dt; Value = $val })
        }
    } finally {
        $reader.Dispose()
    }
    if ($rows.Count -eq 0) {
        throw "$($Meta.Path): no usage rows parsed."
    }
    return ($rows | Sort-Object Time)
}

function Get-Inputs {
    param([string]$InputPath)
    if (Test-Path -LiteralPath $InputPath -PathType Leaf) {
        return @((Get-Item -LiteralPath $InputPath))
    }
    if (Test-Path -LiteralPath $InputPath -PathType Container) {
        $files = Get-ChildItem -LiteralPath $InputPath -Filter '*HourlyUsage.csv' -File | Sort-Object Name
        if ($files.Count -eq 0) {
            throw "No *HourlyUsage.csv files in '$InputPath'."
        }
        return $files
    }
    throw "Path not found: '$InputPath'."
}

# ---------------------------------------------------------------------------
# Math
# ---------------------------------------------------------------------------

function Get-Analysis {
    param(
        [double[]]$Usage,
        [double[]]$Discounts
    )
    $sorted = ($Usage | Sort-Object)
    $n = $Usage.Count
    $total = ($Usage | Measure-Object -Sum).Sum
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($d in $Discounts) {
        $k = Get-Quantile -SortedValues $sorted -Q $d
        $cost = Get-CostWithSp -Usage $Usage -K $k -D $d
        $savHr  = ($total - $cost) / $n
        $wasteHr = (Get-WasteDollars -Usage $Usage -K $k -D $d) / $n
        $overflow = 0; $wasteHrs = 0
        foreach ($u in $Usage) {
            if ($u -gt $k) { $overflow++ }
            elseif ($u -lt $k) { $wasteHrs++ }
        }
        $out.Add([pscustomobject]@{
            Discount      = $d
            KStar         = $k
            SavingsPerHr  = $savHr
            WastePerHr    = $wasteHr
            HrsOverflow   = $overflow
            HrsWaste      = $wasteHrs
        })
    }
    return $out
}

function Get-WasteFreeMetrics {
    param(
        [double[]]$Usage,
        [double[]]$Discounts
    )
    $k = ($Usage | Measure-Object -Minimum).Minimum
    $n = $Usage.Count
    $total = ($Usage | Measure-Object -Sum).Sum
    $rates = [ordered]@{}
    foreach ($d in $Discounts) {
        $cost = Get-CostWithSp -Usage $Usage -K $k -D $d
        $savHr = ($total - $cost) / $n
        $wasteHr = (Get-WasteDollars -Usage $Usage -K $k -D $d) / $n
        $rates[[string]$d] = [pscustomobject]@{ Discount = $d; SavingsPerHr = $savHr; WastePerHr = $wasteHr }
    }
    return [pscustomobject]@{ KNoWaste = $k; Rates = $rates }
}

# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

function Get-SectionTitle {
    param([pscustomobject]$Meta)
    $parts = @()
    if ($Meta.ExportDate) { $parts += $Meta.ExportDate }
    if ($Meta.ScopeKind -and $Meta.ScopeIds) {
        $label = $ScopeKindLabels[$Meta.ScopeKind]
        if (-not $label) { $label = $Meta.ScopeKind }
        $parts += "$label $($Meta.ScopeIds)"
    }
    if ($parts.Count -eq 0) { return $Meta.FileName }
    return ($parts -join ' — ')
}

function Get-FileSectionMarkdown {
    param(
        [pscustomobject]$Meta,
        [object[]]$Data,
        [double[]]$Discounts
    )
    $usage = @($Data | ForEach-Object { $_.Value })
    $n = $usage.Count
    $totalPaygo = ($usage | Measure-Object -Sum).Sum
    $stats = $usage | Measure-Object -Minimum -Maximum -Average
    $uMin = $stats.Minimum; $uMax = $stats.Maximum; $uMean = $stats.Average
    $variance = 0.0
    foreach ($u in $usage) { $variance += [math]::Pow($u - $uMean, 2) }
    $uStd = [math]::Sqrt($variance / $n)
    $first = $Data[0].Time
    $last  = $Data[-1].Time
    $obsDays = $n / 24.0
    $termH = if ($Meta.TermYears) { $HoursPerYear * $Meta.TermYears } else { $HoursPerYear * 3 }
    $termLabel = if ($Meta.TermYears) { "$($Meta.TermYears)-year" } else { '(term unknown)' }

    $wf = Get-WasteFreeMetrics -Usage $usage -Discounts $Discounts
    $opts = Get-Analysis -Usage $usage -Discounts $Discounts

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("### $(Get-SectionTitle -Meta $Meta)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('| Metadata | Value |')
    [void]$sb.AppendLine('|---|---|')
    [void]$sb.AppendLine("| Source file | ``$($Meta.FileName)`` |")
    if ($Meta.ExportDate) {
        [void]$sb.AppendLine("| Export date (filename) | $($Meta.ExportDate) |")
    }
    if ($Meta.ScopeKind) {
        $kindLabel = $ScopeKindLabels[$Meta.ScopeKind]
        if (-not $kindLabel) { $kindLabel = $Meta.ScopeKind }
        [void]$sb.AppendLine("| Scope (filename) | $kindLabel — ``$($Meta.ScopeIds)`` |")
    }
    [void]$sb.AppendLine("| Locale detected | ``$($Meta.Locale)`` (delimiter ``$($Meta.Delimiter)``, date ``$($Meta.DateFormat)``) |")
    [void]$sb.AppendLine("| Lookback (filename) | $($Meta.LookbackDays) days |")
    [void]$sb.AppendLine("| SP term (filename) | $($Meta.TermYears) year(s) — $(Format-Int $termH) h |")
    [void]$sb.AppendLine("| Observed rows | $(Format-Int $n) hours ($($obsDays.ToString('N2', $Inv)) days) |")
    [void]$sb.AppendLine("| Date range | $($first.ToString('yyyy-MM-dd HH:mm')) → $($last.ToString('yyyy-MM-dd HH:mm')) |")
    [void]$sb.AppendLine("| PayGo total (observed) | $(Format-Money $totalPaygo) |")
    [void]$sb.AppendLine("| PayGo /hour | min $(Format-Money $uMin) · mean $(Format-Money $uMean) · max $(Format-Money $uMax) · σ $(Format-Money $uStd) |")
    [void]$sb.AppendLine("| Projected PayGo over term | $(Format-Money ($uMean * $termH)) |")
    [void]$sb.AppendLine('')

    # ---- (a) Waste-free max ----
    [void]$sb.AppendLine('### (a) Waste-free max commitment')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("``K_max_no_waste`` = min hourly PayGo usage = **$(Format-Money $wf.KNoWaste) / hour** (PayGo-equivalent).")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Commitment value over $termLabel`: **$(Format-Money ($wf.KNoWaste * $termH))**.")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("| Discount | Monthly savings | Monthly waste | $termLabel savings | $termLabel waste |")
    [void]$sb.AppendLine('|---|---:|---:|---:|---:|')
    foreach ($d in $Discounts) {
        $r = $wf.Rates[[string]$d]
        [void]$sb.AppendLine("| $(Format-Pct $d) | $(Format-Money ($r.SavingsPerHr * $HoursPerMonth)) | $(Format-Money ($r.WastePerHr * $HoursPerMonth)) | $(Format-Money ($r.SavingsPerHr * $termH)) | $(Format-Money ($r.WastePerHr * $termH)) |")
    }
    [void]$sb.AppendLine('')

    # ---- (b) Optimum per discount ----
    [void]$sb.AppendLine('### (b) Optimum commitment per discount')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('``K*(d) = quantile(U, d)`` — analytic optimum.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("| Discount | K* /hr | Commitment value over term | Monthly savings | Monthly waste | $termLabel savings | $termLabel waste | Hrs overflow | Hrs with waste |")
    [void]$sb.AppendLine('|---|---:|---:|---:|---:|---:|---:|---:|---:|')
    foreach ($r in $opts) {
        [void]$sb.AppendLine("| $(Format-Pct $r.Discount) | $(Format-Money $r.KStar) | $(Format-Money ($r.KStar * $termH)) | $(Format-Money ($r.SavingsPerHr * $HoursPerMonth)) | $(Format-Money ($r.WastePerHr * $HoursPerMonth)) | $(Format-Money ($r.SavingsPerHr * $termH)) | $(Format-Money ($r.WastePerHr * $termH)) | $(Format-Int $r.HrsOverflow) | $(Format-Int $r.HrsWaste) |")
    }
    [void]$sb.AppendLine('')
    return $sb.ToString().TrimEnd()
}

function Write-ConsoleSection {
    param(
        [pscustomobject]$Meta,
        [object[]]$Data,
        [double[]]$Discounts
    )
    $usage = @($Data | ForEach-Object { $_.Value })
    $n = $usage.Count
    $termH = if ($Meta.TermYears) { $HoursPerYear * $Meta.TermYears } else { $HoursPerYear * 3 }
    $wf = Get-WasteFreeMetrics -Usage $usage -Discounts $Discounts
    $opts = Get-Analysis -Usage $usage -Discounts $Discounts

    Write-Host ''
    Write-Host ('=' * 88)
    Write-Host " $(Get-SectionTitle -Meta $Meta)"
    Write-Host "   source: $($Meta.FileName)"
    Write-Host ('=' * 88)
    Write-Host ("  locale={0}  lookback={1}d  term={2}y ({3} h)  rows={4}" -f $Meta.Locale, $Meta.LookbackDays, $Meta.TermYears, (Format-Int $termH), (Format-Int $n))
    Write-Host ("  (a) K_max_no_waste = {0} /hr  -> {1} over term" -f (Format-Money $wf.KNoWaste), (Format-Money ($wf.KNoWaste * $termH)))
    foreach ($d in $Discounts) {
        $r = $wf.Rates[[string]$d]
        Write-Host ("      {0,4}: monthly save {1,10}  waste {2,8}  | term save {3,12}  waste {4,10}" -f (Format-Pct $d), (Format-Money ($r.SavingsPerHr * $HoursPerMonth)), (Format-Money ($r.WastePerHr * $HoursPerMonth)), (Format-Money ($r.SavingsPerHr * $termH)), (Format-Money ($r.WastePerHr * $termH)))
    }
    Write-Host '  (b) Optimum K*(d):'
    foreach ($r in $opts) {
        Write-Host ("      {0,4}: K*={1,8}/hr ({2,12} over term)  monthly save {3,10}  waste {4,7}" -f (Format-Pct $r.Discount), (Format-Money $r.KStar), (Format-Money ($r.KStar * $termH)), (Format-Money ($r.SavingsPerHr * $HoursPerMonth)), (Format-Money ($r.WastePerHr * $HoursPerMonth)))
    }
}

# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

function New-OutputFileName {
    # Auto-derive output filename from input metadata, mirroring the
    # input file naming convention:
    #   {yyyy-MM-dd}-{kind}-{ids}-Last{N}Days-P{T}Y-SavingsPlanRecommendation.md
    # Shared tokens are reused; differing tokens are joined with '+';
    # missing tokens fall back to today's date / 'MultiScope' / 'LastUnknown'
    # / 'PUnknownY' so the filename is always well-formed.
    param([object[]]$Entries)

    $dates = @($Entries | ForEach-Object { $_.Meta.ExportDate } | Where-Object { $_ } | Sort-Object -Unique)
    $kinds = @($Entries | ForEach-Object { $_.Meta.ScopeKind }  | Where-Object { $_ } | Sort-Object -Unique)
    $ids   = @($Entries | ForEach-Object { $_.Meta.ScopeIds }   | Where-Object { $_ } | Sort-Object -Unique)
    $lbs   = @($Entries | ForEach-Object { $_.Meta.LookbackDays } | Where-Object { $null -ne $_ } | Sort-Object -Unique)
    $tms   = @($Entries | ForEach-Object { $_.Meta.TermYears }    | Where-Object { $null -ne $_ } | Sort-Object -Unique)

    $datePart = if ($dates.Count -eq 1) { $dates[0] }
                elseif ($dates.Count -gt 1) { $dates[-1] }
                else { (Get-Date).ToString('yyyy-MM-dd') }

    if ($kinds.Count -eq 1 -and $ids.Count -eq 1) {
        $scopePart = "$($kinds[0])-$($ids[0])"
    } elseif ($kinds.Count -eq 1) {
        $scopePart = "$($kinds[0])-MultiScope"
    } else {
        $scopePart = 'MultiScope'
    }

    $lookbackPart = if ($lbs.Count -eq 1) { "Last$($lbs[0])Days" }
                    elseif ($lbs.Count -gt 1) { "Last$(($lbs -join '+'))Days" }
                    else { 'LastUnknownDays' }

    $termPart = if ($tms.Count -eq 1) { "P$($tms[0])Y" }
                elseif ($tms.Count -gt 1) { "P$(($tms -join '+'))Y" }
                else { 'PUnknownY' }

    # Sanitise: strip filesystem-hostile chars (the scope ids may contain
    # forward slashes if the user fed an odd filename through).
    $name = "$datePart-$scopePart-$lookbackPart-$termPart-SavingsPlanRecommendation.md"
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalid) { $name = $name.Replace([string]$ch, '_') }
    return $name
}

$discountList = [double[]](ConvertTo-DiscountList -Text $Discounts)
$inputs = Get-Inputs -InputPath $Path

$metas = New-Object System.Collections.Generic.List[object]
foreach ($f in $inputs) {
    try {
        $locInfo = Get-LocaleAndDelimiter -FilePath $f.FullName
        $fname   = Get-FilenameMeta -FileName $f.Name
        $meta = [pscustomobject]@{
            Path         = $f.FullName
            FileName     = $f.Name
            Locale       = $locInfo.Locale
            Delimiter    = $locInfo.Delimiter
            DateFormat   = $locInfo.DateFormat
            ExportDate   = $fname.ExportDate
            ScopeKind    = $fname.ScopeKind
            ScopeIds     = $fname.ScopeIds
            LookbackDays = $fname.LookbackDays
            TermYears    = $fname.TermYears
        }
        $data = @(Get-UsageData -Meta $meta)
        $metas.Add([pscustomobject]@{ Meta = $meta; Data = $data })
    } catch {
        Write-Warning "skipping $($f.Name): $($_.Exception.Message)"
    }
}

if ($metas.Count -eq 0) {
    Write-Error 'No usable input files.'
    exit 2
}

$md = [System.Text.StringBuilder]::new()
[void]$md.AppendLine('# Azure Savings Plan — Optimization Results')
[void]$md.AppendLine('')
[void]$md.AppendLine("_Generated: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))_  ")
[void]$md.AppendLine('_Script: `Optimize-SavingsPlanCommitment.ps1` v3.2026-06-15.0_  ')
$discountPctList = (@($discountList | ForEach-Object { Format-Pct $_ })) -join ', '
$dMin = Format-Pct ($discountList | Measure-Object -Minimum).Minimum
$dMax = Format-Pct ($discountList | Measure-Object -Maximum).Maximum
[void]$md.AppendLine("_Discounts analysed: ${discountPctList}_  ")
[void]$md.AppendLine('_Commitment K reported in **PayGo-equivalent USD/hour** (matches Azure SP recommender output)._')
[void]$md.AppendLine('')
[void]$md.AppendLine("> **About the discount range.** Azure Savings Plan discounts vary by SKU, region, and reservation footprint — there is no single ""SP discount"" for a given scope. The default range above spans **common average** SP discounts from **conservative ($dMin)** to **optimistic ($dMax)**, so the recommendation reads as a *sensitivity band* rather than a single point estimate. Pass your own list via ``-Discounts`` (e.g. ``-Discounts 20,25,30``) if your workload mix justifies a different range.")
[void]$md.AppendLine('')
[void]$md.AppendLine('## Inputs')
[void]$md.AppendLine('')
[void]$md.AppendLine('| File | Date | Scope | Lookback | Term | Locale |')
[void]$md.AppendLine('|---|---|---|---|---|---|')
foreach ($entry in $metas) {
    $m = $entry.Meta
    if ($m.ScopeKind -and $m.ScopeIds) {
        $kindLabel = $ScopeKindLabels[$m.ScopeKind]
        if (-not $kindLabel) { $kindLabel = $m.ScopeKind }
        $scopeCell = "$kindLabel ``$($m.ScopeIds)``"
    } else {
        $scopeCell = '?'
    }
    $date = if ($m.ExportDate) { $m.ExportDate } else { '?' }
    $lb   = if ($null -ne $m.LookbackDays) { "$($m.LookbackDays) d" } else { '?' }
    $tm   = if ($null -ne $m.TermYears)    { "$($m.TermYears) y" } else { '?' }
    [void]$md.AppendLine("| ``$($m.FileName)`` | $date | $scopeCell | $lb | $tm | $($m.Locale) |")
}
[void]$md.AppendLine('')

# Group by lookback days; unknowns last.
$groups = @{}
foreach ($entry in $metas) {
    $key = $entry.Meta.LookbackDays
    if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
    $groups[$key].Add($entry)
}
$keys = $groups.Keys | Sort-Object { if ($null -eq $_) { [int]::MaxValue } else { [int]$_ } }
foreach ($k in $keys) {
    $heading = if ($null -ne $k) { "## $k day recommendation" } else { '## Recommendations (lookback unknown)' }
    [void]$md.AppendLine($heading)
    [void]$md.AppendLine('')
    foreach ($entry in $groups[$k]) {
        [void]$md.AppendLine((Get-FileSectionMarkdown -Meta $entry.Meta -Data $entry.Data -Discounts $discountList))
        [void]$md.AppendLine('')
        Write-ConsoleSection -Meta $entry.Meta -Data $entry.Data -Discounts $discountList
    }
}

[void]$md.AppendLine('---')
[void]$md.AppendLine('')
[void]$md.AppendLine('### Notes')
[void]$md.AppendLine('')
[void]$md.AppendLine('* **Azure SP discounts vary by SKU**, region, and reservation footprint. The ``Discount`` columns in the tables above are *average blended* rates; treat them as a sensitivity band, not a single quoted rate.')
[void]$md.AppendLine('* Hours per month = 8 766 / 12 = 730.5; hours per year = 8 766 (average Gregorian year).')
[void]$md.AppendLine('* Term hours = years × 8 766 (1 y = 8 766 h; 3 y = 26 298 h).')
[void]$md.AppendLine('* Monthly and term figures are linear extrapolations of the **observed per-hour rate** during the lookback window. They assume the lookback is representative of the term.')
[void]$md.AppendLine('* `K_max_no_waste` is the largest commitment with zero waste in the observed window; future usage dips below the observed minimum will still cause waste.')
[void]$md.AppendLine('* `K*(d) = quantile(U, d)` is the analytic optimum of `Cost(K)=N·K·(1−d) + Σ max(0, U−K)`.')
[void]$md.AppendLine('')

if ([string]::IsNullOrWhiteSpace($Out)) {
    $Out = Join-Path $PSScriptRoot (New-OutputFileName -Entries $metas)
}

Set-Content -LiteralPath $Out -Value $md.ToString() -Encoding UTF8
Write-Host ''
Write-Host "Wrote $Out"
