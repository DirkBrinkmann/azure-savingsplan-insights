<#
.SYNOPSIS
    Interactively exports Azure Cost Management Benefit Recommendations to CSV.

.DESCRIPTION
    Prompts the user for billing scope, lookback period, term, output folder
    and CSV format, then calls the Azure Cost Management
    "Benefit Recommendations - List" REST API
    (api-version=2025-03-01) with both $expand=properties/usage and
    $expand=properties/allRecommendationDetails.

    Produces two CSV files:
      1. yyyy-MM-dd-<scope>-<lookback>-<term>-AllRecommendationDetails.csv
         Flattened properties.allRecommendationDetails across all returned
         recommendations.
      2. yyyy-MM-dd-<scope>-<lookback>-<term>-HourlyUsage.csv
         Hourly Pay-As-You-Go usage for the recommended (top) recommendation,
         expanded to one row per hour across the lookback window.
         Columns: DateTime,HourlyPayGoUsage

    Authenticates using the Az PowerShell module (Get-AzAccessToken).

.EXAMPLE
    PS> .\Export-BenefitRecommendations.ps1

.NOTES
    Version : v1.0.2026-06-11.11
    Author  : Dirk Brinkmann
    Requires: PowerShell 5.1+ and the Az.Accounts module (Connect-AzAccount first).
    API ref : https://learn.microsoft.com/rest/api/cost-management/benefit-recommendations/list?view=rest-cost-management-2025-03-01
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$script:ApiVersion = '2025-03-01'
$script:ArmBase    = 'https://management.azure.com'

#region Helpers ---------------------------------------------------------------

function Read-Choice {
    <#
        Prompts the user with a numbered menu and returns the chosen value.
    #>
    param(
        [Parameter(Mandatory)] [string]   $Title,
        [Parameter(Mandatory)] [string[]] $Options,
        [int] $DefaultIndex = 0
    )
    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    for ($i = 0; $i -lt $Options.Count; $i++) {
        $marker = if ($i -eq $DefaultIndex) { '*' } else { ' ' }
        Write-Host ("  {0} [{1}] {2}" -f $marker, ($i + 1), $Options[$i])
    }
    while ($true) {
        $raw = Read-Host ("Enter choice 1-{0} (default {1})" -f $Options.Count, ($DefaultIndex + 1))
        if ([string]::IsNullOrWhiteSpace($raw)) { return $Options[$DefaultIndex] }
        if ($raw -match '^\d+$') {
            $idx = [int]$raw - 1
            if ($idx -ge 0 -and $idx -lt $Options.Count) { return $Options[$idx] }
        }
        Write-Warning "Invalid selection."
    }
}

function Read-NonEmpty {
    param([Parameter(Mandatory)][string] $Prompt)
    while ($true) {
        $v = Read-Host $Prompt
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
        Write-Warning "Value is required."
    }
}

function ConvertTo-SafeSlug {
    param([Parameter(Mandatory)][string] $Value)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + @(' ', '/', '\', ':', '?', '*', '"', '<', '>', '|')
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Value.ToCharArray()) {
        if ($invalid -contains $c) { [void]$sb.Append('_') } else { [void]$sb.Append($c) }
    }
    ($sb.ToString()).Trim('_')
}

function Get-ArmAccessToken {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        throw "Az.Accounts module is not installed. Run: Install-Module Az.Accounts -Scope CurrentUser"
    }
    Import-Module Az.Accounts -ErrorAction Stop | Out-Null

    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Host "No Az context found. Launching Connect-AzAccount..." -ForegroundColor Yellow
        Connect-AzAccount | Out-Null
    }

    # Get-AzAccessToken changed in Az.Accounts 2.x/3.x to return SecureString by default.
    $tokenObj = $null
    try {
        $tokenObj = Get-AzAccessToken -ResourceUrl $script:ArmBase -AsSecureString:$false -ErrorAction Stop
    } catch {
        # Older versions don't support -AsSecureString.
        $tokenObj = Get-AzAccessToken -ResourceUrl $script:ArmBase -ErrorAction Stop
    }

    $tokenStr = $null
    if ($tokenObj.Token -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
        try   { $tokenStr = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } else {
        $tokenStr = [string]$tokenObj.Token
    }
    return $tokenStr
}

function Invoke-ArmGet {
    <#
        Calls an ARM URL with Authorization header. Returns parsed JSON.
        Throws a descriptive error with the HTTP status code on failure.
    #>
    param(
        [Parameter(Mandatory)][string] $Url,
        [Parameter(Mandatory)][string] $Token
    )
    try {
        return Invoke-RestMethod -Method GET -Uri $Url -Headers @{
            Authorization = "Bearer $Token"
            'Content-Type' = 'application/json'
        } -ErrorAction Stop
    } catch {
        $resp   = $_.Exception.Response
        $status = if ($resp) { [int]$resp.StatusCode } else { 0 }
        $body   = $null
        if ($resp) {
            try {
                $stream = $resp.GetResponseStream()
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $body   = $reader.ReadToEnd()
                }
            } catch { }
        }
        $hint = switch ($status) {
            401 { "Authentication failed. Run Connect-AzAccount and retry." }
            403 { "Authorization failed. Your account lacks the required role (e.g. Cost Management Reader / Billing Reader) on the selected scope." }
            404 { "Scope not found. Verify the IDs you entered and your tenant context." }
            default { "Unexpected HTTP $status from ARM." }
        }
        throw "ARM call failed: $hint`nURL: $Url`nBody: $body"
    }
}

function ConvertTo-FlatItems {
    <#
        Recursively walks a value and emits leaf items (anything that is not
        an array/IEnumerable; strings are treated as scalars). Used to robustly
        flatten responses where an array property may be returned nested one
        or more levels deep (e.g. [[{...},{...}]]).
    #>
    param($Value)
    if ($null -eq $Value) { return }
    if ($Value -is [string] -or -not ($Value -is [System.Collections.IEnumerable])) {
        Write-Output -InputObject $Value -NoEnumerate
        return
    }
    foreach ($child in $Value) {
        ConvertTo-FlatItems $child
    }
}

function ConvertTo-CsvCell {
    # Make a value safe for CSV output: scalars pass through, anything complex
    # becomes a compact JSON string (so an array doesn't render as
    # "System.Object[]").
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string])  { return $Value }
    if ($Value -is [bool])    { return [string]$Value }
    if ($Value.GetType().IsPrimitive) { return $Value }
    if ($Value -is [datetime]) { return $Value.ToString('o') }
    if ($Value -is [System.Collections.IEnumerable]) {
        try { return ($Value | ConvertTo-Json -Compress -Depth 10) } catch { return [string]$Value }
    }
    try { return ($Value | ConvertTo-Json -Compress -Depth 10) } catch { return [string]$Value }
}

function Format-MdValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    # Strings first — they are IEnumerable in .NET, so check before the array branch.
    if ($Value -is [string])  { return $Value }
    if ($Value -is [bool])    { return [string]$Value }
    if ($Value.GetType().IsPrimitive) { return [string]$Value }
    if ($Value -is [datetime]) { return $Value.ToString('o') }
    # Arrays / collections.
    if ($Value -is [System.Collections.IEnumerable]) {
        try { return ($Value | ConvertTo-Json -Compress -Depth 10) } catch { return [string]$Value }
    }
    # Complex objects (PSCustomObject from JSON, etc.).
    try { return ($Value | ConvertTo-Json -Compress -Depth 10) } catch { return [string]$Value }
}

function Resolve-DetailsArray {
    <#
        The Cost Management API returns `properties.allRecommendationDetails`
        as an object wrapper { "value": [...] } in 2025-03-01 responses, but
        older variants returned a plain array. Accept both shapes and return
        the inner array.
    #>
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return @($Value)
    }
    if ($Value -is [psobject] -and ($Value.PSObject.Properties.Name -contains 'value')) {
        $inner = $Value.value
        if ($inner -is [System.Collections.IEnumerable] -and -not ($inner -is [string])) {
            return @($inner)
        }
    }
    return @($Value)
}

#endregion Helpers

#region Input collection -------------------------------------------------------

Write-Host "=== Azure Benefit Recommendations Export ===" -ForegroundColor Green

$scopeKind = Read-Choice -Title 'Select billing scope kind:' -Options @(
    'Billing Account (EA)',
    'Billing Profile (MCA)',
    'Subscription',
    'Resource Group'
)

$scopePath = $null
$scopeSlug = $null
switch ($scopeKind) {
    'Billing Account (EA)' {
        $ba        = Read-NonEmpty 'Billing Account ID'
        $scopePath = "/providers/Microsoft.Billing/billingAccounts/$ba"
        $scopeSlug = "EA-$ba"
    }
    'Billing Profile (MCA)' {
        $ba        = Read-NonEmpty 'Billing Account ID'
        $bp        = Read-NonEmpty 'Billing Profile ID'
        $scopePath = "/providers/Microsoft.Billing/billingAccounts/$ba/billingProfiles/$bp"
        $scopeSlug = "MCA-$ba-$bp"
    }
    'Subscription' {
        $sub       = Read-NonEmpty 'Subscription ID'
        $scopePath = "/subscriptions/$sub"
        $scopeSlug = "SUB-$sub"
    }
    'Resource Group' {
        $sub       = Read-NonEmpty 'Subscription ID'
        $rg        = Read-NonEmpty 'Resource Group name'
        $scopePath = "/subscriptions/$sub/resourceGroups/$rg"
        $scopeSlug = "RG-$sub-$rg"
    }
}
$scopeSlug = ConvertTo-SafeSlug $scopeSlug

$lookback = Read-Choice -Title 'Select lookback period:' -Options @('Last7Days','Last30Days','Last60Days') -DefaultIndex 1
$term     = Read-Choice -Title 'Select term:' -Options @('P1Y','P3Y') -DefaultIndex 0

$defaultDir = (Get-Location).Path
$outDirRaw  = Read-Host "Output folder (default: $defaultDir)"
$outDir     = if ([string]::IsNullOrWhiteSpace($outDirRaw)) { $defaultDir } else { $outDirRaw.Trim() }
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$csvFormat = Read-Choice -Title 'CSV format:' -Options @(
    'US (comma, ISO datetime yyyy-MM-ddTHH:mm)',
    'DE (semicolon, datetime dd.MM.yyyy HH:mm)'
) -DefaultIndex 0

$delimiter  = if ($csvFormat -like 'US*') { ',' } else { ';' }
$dtFormat   = if ($csvFormat -like 'US*') { 'yyyy-MM-ddTHH:mm' } else { 'dd.MM.yyyy HH:mm' }

$exportRawChoice = Read-Choice -Title 'Also export the raw JSON response?' -Options @('No','Yes') -DefaultIndex 0
$exportRaw       = ($exportRawChoice -eq 'Yes')

#endregion Input collection

#region API call --------------------------------------------------------------

Write-Host "`nAcquiring ARM access token..." -ForegroundColor Cyan
$token = Get-ArmAccessToken

$expand = 'properties/usage,properties/allRecommendationDetails'
$filter = "properties/lookBackPeriod eq '$lookback' and properties/term eq '$term'"

$query  = @(
    "api-version=$script:ApiVersion",
    ("`$expand=" + [System.Uri]::EscapeDataString($expand)),
    ("`$filter=" + [System.Uri]::EscapeDataString($filter))
) -join '&'

$url = "$script:ArmBase$scopePath/providers/Microsoft.CostManagement/benefitRecommendations?$query"

Write-Host "Verifying API access and fetching recommendations..." -ForegroundColor Cyan
Write-Verbose "GET $url"

$all = New-Object System.Collections.Generic.List[object]
$next = $url
$page = 0
while ($next) {
    $page++
    Write-Host ("  Page {0}..." -f $page)
    $resp = Invoke-ArmGet -Url $next -Token $token
    $respProps = $resp.PSObject.Properties.Name
    if ($respProps -contains 'value' -and $resp.value) { $all.AddRange([object[]]$resp.value) }
    $next = if ($respProps -contains 'nextLink') { $resp.nextLink } else { $null }
}

Write-Host ("Retrieved {0} recommendation(s)." -f $all.Count) -ForegroundColor Green
if ($all.Count -eq 0) {
    Write-Warning "No recommendations returned for the selected scope/lookback/term. Nothing to export."
    return
}

#endregion API call

#region Output filenames ------------------------------------------------------

$datePrefix = (Get-Date).ToString('yyyy-MM-dd')
$base       = "$datePrefix-$scopeSlug-$lookback-$term"
$detailsCsv = Join-Path $outDir "$base-AllRecommendationDetails.csv"
$topMd      = Join-Path $outDir "$base-TopRecommendation.md"
$hourlyCsv  = Join-Path $outDir "$base-HourlyUsage.csv"
$rawJson    = Join-Path $outDir "$base-Raw.json"

#endregion

#region Optional raw JSON export ---------------------------------------------

if ($exportRaw) {
    Write-Host "`nWriting raw JSON response..." -ForegroundColor Cyan
    $rawWrapper = [pscustomobject]@{
        exportedAt    = (Get-Date).ToString('o')
        apiVersion    = $script:ApiVersion
        scope         = $scopePath
        scopeKind     = $scopeKind
        lookBackPeriod= $lookback
        term          = $term
        expand        = $expand
        count         = $all.Count
        value         = $all.ToArray()
    }
    $json = $rawWrapper | ConvertTo-Json -Depth 100
    [System.IO.File]::WriteAllText($rawJson, $json, [System.Text.UTF8Encoding]::new($true))
    Write-Host ("  Wrote {0}" -f $rawJson) -ForegroundColor Green
}

#endregion

#region Top recommendation detection ------------------------------------------

# Pick the recommended/top item: prefer the recommendation whose
# allRecommendationDetails contains an entry with recommended=true; otherwise
# fall back to the first item (the service returns best-fit first).
$top              = $null
$topRecommendedId = $null   # id of the inner allRecommendationDetails entry flagged recommended
foreach ($rec in $all) {
    $p = $rec.properties
    if ($p -and $p.PSObject.Properties.Name -contains 'allRecommendationDetails' -and $p.allRecommendationDetails) {
        foreach ($entry in (Resolve-DetailsArray $p.allRecommendationDetails)) {
            $entryProps = $entry.PSObject.Properties.Name
            if ($entryProps -contains 'recommended' -and [bool]$entry.recommended) {
                $top              = $rec
                $topRecommendedId = if ($entryProps -contains 'recommendationId') { [string]$entry.recommendationId } else { $null }
                break
            }
        }
    }
    if ($top) { break }
}
if (-not $top) { $top = $all[0] }

#endregion

#region CSV 1: AllRecommendationDetails (flat) -------------------------------

Write-Host "`nWriting AllRecommendationDetails CSV (top recommendation's allSavingsBenefitDetails collection)..." -ForegroundColor Cyan

$pTop = $top.properties
$allDetails = @()
if ($pTop -and $pTop.PSObject.Properties.Name -contains 'allRecommendationDetails' -and $pTop.allRecommendationDetails) {
    # 2025-03-01 wraps the collection in { "value": [...] }; older variants returned a plain array.
    $resolved = Resolve-DetailsArray $pTop.allRecommendationDetails
    $allDetails = @(ConvertTo-FlatItems $resolved | Where-Object { $_ -is [psobject] -and -not ($_ -is [string]) -and -not ($_ -is [ValueType]) })
}

if ($allDetails.Count -eq 0) {
    Write-Warning "Top recommendation has no allRecommendationDetails (allSavingsBenefitDetails) collection; CSV not written."
} else {
    # Collect the union of all property names across the collection so the CSV
    # has a stable, complete header even if some elements omit fields.
    $allKeys = [ordered]@{}
    foreach ($d in $allDetails) {
        foreach ($prop in $d.PSObject.Properties) {
            if (-not $allKeys.Contains($prop.Name)) { $allKeys[$prop.Name] = $true }
        }
    }

    $detailRows = New-Object System.Collections.Generic.List[object]
    foreach ($d in $allDetails) {
        $present = $d.PSObject.Properties.Name
        $row = [ordered]@{}
        foreach ($k in $allKeys.Keys) {
            $raw = if ($present -contains $k) { $d.$k } else { $null }
            $row[$k] = ConvertTo-CsvCell $raw
        }
        $detailRows.Add([pscustomobject]$row)
    }

    $detailRows | Export-Csv -Path $detailsCsv -NoTypeInformation -Delimiter $delimiter -Encoding UTF8
    Write-Host ("  Wrote {0} row(s) to {1}" -f $detailRows.Count, $detailsCsv) -ForegroundColor Green
}

#endregion

#region MD: TopRecommendation -------------------------------------------------

Write-Host "`nWriting TopRecommendation markdown..." -ForegroundColor Cyan

# Identify the recommended commitment inside allRecommendationDetails.
$recommendedDetail = $null
foreach ($entry in $allDetails) {
    if ($entry.PSObject.Properties.Name -contains 'recommended' -and [bool]$entry.recommended) {
        $recommendedDetail = $entry; break
    }
}
if (-not $recommendedDetail -and $allDetails.Count -gt 0) {
    $recommendedDetail = $allDetails[0]
}

$today = (Get-Date).ToString('yyyy-MM-dd')

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine("# Top Benefit Recommendation")
[void]$md.AppendLine()
[void]$md.AppendLine("Generated by ``Export-BenefitRecommendations.ps1`` on $today.")
[void]$md.AppendLine()

[void]$md.AppendLine("## Context")
[void]$md.AppendLine()
[void]$md.AppendLine("| Field | Value |")
[void]$md.AppendLine("|---|---|")
[void]$md.AppendLine("| Scope | ``$scopePath`` |")
[void]$md.AppendLine("| Scope kind | $scopeKind |")
[void]$md.AppendLine("| Lookback period | $lookback |")
[void]$md.AppendLine("| Term | $term |")
[void]$md.AppendLine("| API version | $script:ApiVersion |")
[void]$md.AppendLine("| Recommendations returned | $($all.Count) |")
[void]$md.AppendLine()

[void]$md.AppendLine("## Recommendation")
[void]$md.AppendLine()
[void]$md.AppendLine("| Field | Value |")
[void]$md.AppendLine("|---|---|")
[void]$md.AppendLine("| id | ``$($top.id)`` |")
[void]$md.AppendLine("| name | $($top.name) |")
if ($top.PSObject.Properties.Name -contains 'kind')     { [void]$md.AppendLine("| kind | $($top.kind) |") }
if ($top.PSObject.Properties.Name -contains 'type')     { [void]$md.AppendLine("| type | $($top.type) |") }
if ($top.PSObject.Properties.Name -contains 'location') { [void]$md.AppendLine("| location | $($top.location) |") }
if ($pTop) {
    foreach ($prop in $pTop.PSObject.Properties) {
        if ($prop.Name -in @('allRecommendationDetails','usage')) { continue }
        [void]$md.AppendLine(("| {0} | {1} |" -f $prop.Name, (Format-MdValue $prop.Value)))
    }
}
[void]$md.AppendLine()

if ($recommendedDetail) {
    [void]$md.AppendLine("## Recommended commitment")
    [void]$md.AppendLine()
    [void]$md.AppendLine("This is the entry from ``allRecommendationDetails`` (an ``AllSavingsBenefitDetails`` object) flagged as ``recommended = true``.")
    [void]$md.AppendLine()
    [void]$md.AppendLine("| Field | Value |")
    [void]$md.AppendLine("|---|---|")
    foreach ($prop in $recommendedDetail.PSObject.Properties) {
        [void]$md.AppendLine(("| {0} | {1} |" -f $prop.Name, (Format-MdValue $prop.Value)))
    }
    [void]$md.AppendLine()
}

if ($allDetails.Count -gt 0) {
    # Build a wide table over the union of fields, in a deterministic order.
    $headerKeys = [ordered]@{}
    foreach ($d in $allDetails) {
        foreach ($prop in $d.PSObject.Properties) {
            if (-not $headerKeys.Contains($prop.Name)) { $headerKeys[$prop.Name] = $true }
        }
    }
    $keysArr = @($headerKeys.Keys)

    [void]$md.AppendLine("## All commitment alternatives")
    [void]$md.AppendLine()
    [void]$md.AppendLine("Source: ``properties.allRecommendationDetails`` (collection of ``AllSavingsBenefitDetails``). Also exported as CSV: ``$([System.IO.Path]::GetFileName($detailsCsv))``.")
    [void]$md.AppendLine()
    [void]$md.AppendLine("| " + ($keysArr -join " | ") + " |")
    [void]$md.AppendLine("|" + (($keysArr | ForEach-Object { '---' }) -join "|") + "|")
    foreach ($d in $allDetails) {
        $present = $d.PSObject.Properties.Name
        $cells = foreach ($k in $keysArr) {
            $v = if ($present -contains $k) { $d.$k } else { $null }
            (Format-MdValue $v) -replace '\|','\|' -replace '\r?\n',' '
        }
        [void]$md.AppendLine("| " + ($cells -join " | ") + " |")
    }
    [void]$md.AppendLine()
}

[System.IO.File]::WriteAllText($topMd, $md.ToString(), [System.Text.UTF8Encoding]::new($true))
Write-Host ("  Wrote {0}" -f $topMd) -ForegroundColor Green

#endregion

#region Print overall recommendation -----------------------------------------

Write-Host "`n=== Overall Recommendation ===" -ForegroundColor Yellow

$topSummary = [ordered]@{
    RecommendationId = $top.id
    Name             = $top.name
    Kind             = if ($top.PSObject.Properties.Name -contains 'kind') { $top.kind } else { $null }
}
if ($pTop) {
    foreach ($prop in $pTop.PSObject.Properties) {
        if ($prop.Name -in @('allRecommendationDetails','usage')) { continue }
        $topSummary[$prop.Name] = $prop.Value
    }
}
$topSummary.GetEnumerator() | ForEach-Object {
    $val = if ($null -eq $_.Value) { '' } else { $_.Value }
    Write-Host ("  {0,-30} : {1}" -f $_.Name, $val)
}

if ($recommendedDetail) {
    Write-Host "`n  -- Recommended commitment --" -ForegroundColor Yellow
    foreach ($prop in $recommendedDetail.PSObject.Properties) {
        $val = if ($null -eq $prop.Value) { '' } else { $prop.Value }
        Write-Host ("  {0,-30} : {1}" -f $prop.Name, $val)
    }
}

#endregion

#region CSV 2: Hourly usage for top recommendation ---------------------------

Write-Host "`nWriting HourlyUsage CSV (top recommendation)..." -ForegroundColor Cyan

$usage = $null
if ($top.properties -and $top.properties.PSObject.Properties.Name -contains 'usage') {
    $usage = $top.properties.usage
}

$usageProps = if ($usage) { $usage.PSObject.Properties.Name } else { @() }
if (-not $usage -or ($usageProps -notcontains 'charges') -or -not $usage.charges) {
    Write-Warning "Top recommendation has no usage.charges data; skipping hourly CSV."
    return
}

$charges = @($usage.charges)

$dtStyles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# firstConsumptionDate / lastConsumptionDate live on the recommendation's
# properties object in api-version 2025-03-01, but some older variants placed
# them under usage. Probe both, falling back to the lookback window.
$pTopProps   = if ($top.properties) { $top.properties.PSObject.Properties.Name } else { @() }
function Resolve-DateProp {
    param([string] $Name)
    if ($pTopProps -contains $Name -and $top.properties.$Name) { return [string]$top.properties.$Name }
    if ($usageProps -contains $Name -and $usage.$Name)         { return [string]$usage.$Name }
    return $null
}
$firstRaw = Resolve-DateProp 'firstConsumptionDate'
$lastRaw  = Resolve-DateProp 'lastConsumptionDate'

$lookbackDays  = switch ($lookback) { 'Last7Days' { 7 } 'Last30Days' { 30 } 'Last60Days' { 60 } }
$fallbackEnd   = [datetime]::UtcNow
$fallbackEnd   = [datetime]::new($fallbackEnd.Year, $fallbackEnd.Month, $fallbackEnd.Day, $fallbackEnd.Hour, 0, 0, [System.DateTimeKind]::Utc)
$fallbackStart = $fallbackEnd.AddDays(-$lookbackDays)

if ($firstRaw) {
    $firstDt = [datetime]::Parse($firstRaw, $invariant, $dtStyles)
} else {
    Write-Warning "firstConsumptionDate not present; assuming series starts at the lookback window start."
    $firstDt = $fallbackStart
}

if ($lastRaw) {
    $lastDt = [datetime]::Parse($lastRaw, $invariant, $dtStyles)
} else {
    Write-Warning "lastConsumptionDate not present; deriving from firstConsumptionDate + charges count."
    $lastDt = $firstDt.AddHours([math]::Max(0, $charges.Count - 1))
}

# Iterate the actual API-reported data range [firstConsumptionDate,
# lastConsumptionDate], not "now - lookback ... now". The API only emits one
# value per hour it has data for, and lastConsumptionDate is typically a few
# hours/days behind wall-clock time, so walking up to "now" would produce
# trailing empty rows.
$rows = New-Object System.Collections.Generic.List[object]
$rowEnd = $lastDt
# If lastConsumptionDate appears stale relative to charges.Count, prefer the
# count-derived end so we don't truncate real data.
$countDerivedEnd = $firstDt.AddHours([math]::Max(0, $charges.Count - 1))
if ($countDerivedEnd -gt $rowEnd) { $rowEnd = $countDerivedEnd }

for ($t = $firstDt; $t -le $rowEnd; $t = $t.AddHours(1)) {
    $idx = [int][math]::Round(($t - $firstDt).TotalHours)
    if ($idx -lt 0 -or $idx -ge $charges.Count) { continue }
    $value = $charges[$idx]
    $rows.Add([pscustomobject]@{
        DateTime         = $t.ToString($dtFormat, [System.Globalization.CultureInfo]::InvariantCulture)
        HourlyPayGoUsage = $value
    })
}

$rows | Export-Csv -Path $hourlyCsv -NoTypeInformation -Delimiter $delimiter -Encoding UTF8
Write-Host ("  Wrote {0} hourly row(s) to {1}" -f $rows.Count, $hourlyCsv) -ForegroundColor Green

Write-Host "`nDone." -ForegroundColor Green

#endregion
