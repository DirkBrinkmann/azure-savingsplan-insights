v1.2026-06-22.0 Dirk Brinkmann

# Azure Savings Plan Insights

Two PowerShell scripts that, together, take you from raw Azure Cost
Management data to an actionable Savings Plan (SP) commitment
recommendation:

| Step | Script | What it does |
|-----:|--------|--------------|
| 1 | [`Export-BenefitRecommendations.ps1`](./Export-BenefitRecommendations.ps1) | Calls the Azure Cost Management **Benefit Recommendations – List** REST API for a chosen scope (Billing Account / Billing Profile / Subscription / Resource Group), lookback (`Last7Days` / `Last30Days` / `Last60Days`), and term (`P1Y` / `P3Y`). Writes flat CSVs, a Markdown summary, and a per-hour Pay-As-You-Go usage series (`*-HourlyUsage.csv`). |
| 2 | [`Optimize-SavingsPlanCommitment.ps1`](./Optimize-SavingsPlanCommitment.ps1) | Reads one or more of those `*-HourlyUsage.csv` files and computes (a) the largest **waste-free** hourly commitment and (b) the **optimum** commitment across a range of common average SP discounts (default ≈30 % conservative → ≈42 % optimistic — actual per-SKU discounts vary). Output is a Markdown report named after the inputs: `{date}-{scope}-Last{N}Days-P{T}Y-SavingsPlanRecommendation.md`. |

**Typical flow:**

```powershell
# 1. Export hourly PayGo usage for the scope/lookback/term you care about
.\Export-BenefitRecommendations.ps1
#   -> writes 2026-06-11-SUB-…-Last60Days-P1Y-HourlyUsage.csv (and friends)

# 2. Turn that usage into a commitment recommendation
.\Optimize-SavingsPlanCommitment.ps1 -Path .
#   -> writes 2026-06-11-SUB-…-Last60Days-P1Y-SavingsPlanRecommendation.md
```

The two scripts share a filename convention
(`{yyyy-MM-dd}-{EA|MCA|SUB|RG}-{ids}-Last{30|60}Days-P{1|3}Y-…`) so the
Optimizer can pick up exports directly from the folder the Exporter
wrote to.

---

# Azure Benefit Recommendations Export

Interactive PowerShell script that calls the Azure Cost Management
**Benefit Recommendations - List** REST API
([`2025-03-01`](https://learn.microsoft.com/rest/api/cost-management/benefit-recommendations/list?view=rest-cost-management-2025-03-01))
and exports the results to CSV — including a per-hour Pay-As-You-Go usage
series suitable for further analysis (Excel, Power BI, etc.).

Script: [`Export-BenefitRecommendations.ps1`](./Export-BenefitRecommendations.ps1)

---

## Purpose

Azure Cost Management generates Reserved Instance and Savings Plan benefit
recommendations across different scopes (Billing Account, Billing Profile,
Subscription, Resource Group). The REST response is JSON, nested, and contains
several alternative commitment levels per recommendation plus an embedded
hourly usage time series.

This script provides a one-command way to:

1. Verify that you (the signed-in user) actually have access to the API for the
   selected scope.
2. Pull recommendations for a chosen **lookback period** and **term**.
3. Flatten the alternative commitment details into a tabular CSV.
4. Materialise the embedded hourly usage of the **recommended (top)**
   recommendation into a second CSV with one row per hour across the API's
   reported data range — ready for charting or capacity planning.

## Functionality

- Interactive prompts for all inputs (no parameters required).
- Authenticates via the **Az PowerShell module** (`Get-AzAccessToken`).
  Handles both the legacy plain-text token and the newer SecureString return
  type, and runs `Connect-AzAccount` automatically if no Az context is found.
- Always requests both `$expand` options the API supports:
  - `$expand=properties/usage`
  - `$expand=properties/allRecommendationDetails`
- Filters server-side via
  `$filter=properties/lookBackPeriod eq '<lb>' and properties/term eq '<term>'`.
- Follows `nextLink` pagination until exhausted.
- Maps non-2xx HTTP responses to clear, actionable messages:
  | Status | Meaning                                                                          |
  |--------|----------------------------------------------------------------------------------|
  | 401    | Not authenticated — run `Connect-AzAccount`.                                     |
  | 403    | Missing role on the scope (e.g. Cost Management Reader, Billing Reader).         |
  | 404    | Scope path/IDs are wrong or not visible in your current tenant.                  |
  | other  | Raw status + response body included.                                             |

## Inputs (prompted)

| Prompt           | Values                                                                                          |
|------------------|-------------------------------------------------------------------------------------------------|
| Scope kind       | Billing Account (EA), Billing Profile (MCA), Subscription, Resource Group                       |
| Scope IDs        | Depends on scope kind: BillingAccountId [+ BillingProfileId], SubscriptionId [+ ResourceGroup]  |
| Lookback period  | `Last7Days`, `Last30Days` (default), `Last60Days`                                               |
| Term             | `P1Y` (default), `P3Y`                                                                          |
| Output folder    | Any path — default is the current directory; created if missing                                 |
| CSV format       | **US** (comma, ISO datetime `yyyy-MM-ddTHH:mm`) — default; **DE** (semicolon, `dd.MM.yyyy HH:mm`) |
| Export raw JSON  | **No** (default) or **Yes** — when Yes, writes the full API response to a `.json` file alongside the other outputs |

Scope path is built as follows:

| Scope kind            | ARM path                                                                                  |
|-----------------------|-------------------------------------------------------------------------------------------|
| Billing Account (EA)  | `/providers/Microsoft.Billing/billingAccounts/{billingAccountId}`                         |
| Billing Profile (MCA) | `/providers/Microsoft.Billing/billingAccounts/{billingAccountId}/billingProfiles/{bp}`    |
| Subscription          | `/subscriptions/{subscriptionId}`                                                         |
| Resource Group        | `/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}`                      |

## Outputs

Three files are written to the chosen output folder, UTF-8 encoded; CSVs use
the chosen delimiter.

### 1. `yyyy-MM-dd-<scope>-<lookback>-<term>-AllRecommendationDetails.csv`

Flat CSV built from the **top recommendation's `properties.allRecommendationDetails`**
collection (one element of `AllSavingsBenefitDetails` per row). Columns are the
union of all fields seen across the collection, e.g. `commitmentAmount`,
`overageCost`, `benefitCost`, `totalCost`, `savingsAmount`, `coveragePercentage`,
`wastageCost`, `recommended`, …

### 2. `yyyy-MM-dd-<scope>-<lookback>-<term>-TopRecommendation.md`

Markdown file describing the top recommendation in detail:

- **Context** — scope, lookback, term, API version, page count.
- **Recommendation** — every property of the top recommendation (id, name,
  kind, and every field under `properties` except `usage` and
  `allRecommendationDetails`).
- **Recommended commitment** — every field of the `allRecommendationDetails`
  entry flagged `recommended = true`.
- **All commitment alternatives** — the entire `allRecommendationDetails`
  collection rendered as a Markdown table.

### 3. `yyyy-MM-dd-<scope>-<lookback>-<term>-HourlyUsage.csv`

Hourly Pay-As-You-Go usage for the **recommended (top)** recommendation.

### 4. `yyyy-MM-dd-<scope>-<lookback>-<term>-Raw.json` *(optional)*

Written only when the "Also export the raw JSON response?" prompt is answered
**Yes**. Contains the full, unmodified set of recommendation objects returned
by the API (all pages combined), wrapped with a small metadata header
(`exportedAt`, `apiVersion`, `scope`, `scopeKind`, `lookBackPeriod`, `term`,
`expand`, `count`, `value`). Useful for debugging, archival, or feeding the
data into other tooling without re-querying ARM.

Selection rule: the script picks the recommendation whose
`allRecommendationDetails` contains an entry with `recommended = true`; if none
is flagged, it falls back to the first item returned (the service returns
best-fit first).

Columns:

| Column            | Description                                                              |
|-------------------|--------------------------------------------------------------------------|
| `DateTime`        | One row per hour across the API-reported data window `[properties.firstConsumptionDate, properties.lastConsumptionDate]` (UTC, hour-truncated). No trailing blanks. |
| `HourlyPayGoUsage`| Value from `properties.usage.charges`, aligned via `firstConsumptionDate`. This is the **hourly PayGo cost in USD/hour only** — it does **not** include the underlying resource *usage* (consumed quantity) that produced the cost. |

> **⚠ Cost, not usage.** Despite the column name, `HourlyPayGoUsage` is a
> **cost** series (USD/hour from `usage.charges`), not a consumption-quantity
> series. Azure's own recommendations under `properties.allRecommendationDetails`
> (the `allSavingsBenefitDetails` objects) are computed from **both cost and the
> usage that created it**. Any downstream analysis built purely on this hourly
> cost column — including `Optimize-SavingsPlanCommitment.ps1` — is therefore an
> approximation of Azure's usage-aware recommendation. See
> [Analyzer assumptions and caveats](#analyzer-assumptions-and-caveats).

Example (US format):

```
DateTime,HourlyPayGoUsage
2026-05-12T08:00,1.6
2026-05-12T09:00,1.72
...
```

Example (DE format):

```
DateTime;HourlyPayGoUsage
12.05.2026 08:00;1,6
12.05.2026 09:00;1,72
...
```

The `<scope>` slug in the filename is built from the scope kind + the last
scope IDs (sanitised for filesystem-safe characters).

## Requirements

- **PowerShell** 5.1+ (or PowerShell 7+).
- **Az.Accounts** module:
  ```powershell
  Install-Module Az.Accounts -Scope CurrentUser
  ```
- Azure RBAC permission to read benefit recommendations on the selected scope.
  Typical roles: **Cost Management Reader**, **Billing Reader**, or higher.

## Usage

```powershell
git clone https://github.com/DirkBrinkmann/azure-savingsplan-insights.git
cd azure-savingsplan-insights

# Optional: pre-authenticate against the right tenant/subscription
Connect-AzAccount
Set-AzContext -Subscription '<subscription-id-or-name>'   # if relevant

# Run the script — it will prompt for everything
.\Export-BenefitRecommendations.ps1
```

Get the built-in help:

```powershell
Get-Help .\Export-BenefitRecommendations.ps1 -Full
```

### Example session

```
=== Azure Benefit Recommendations Export ===

Select billing scope kind:
  * [1] Billing Account (EA)
    [2] Billing Profile (MCA)
    [3] Subscription
    [4] Resource Group
Enter choice 1-4 (default 1): 3
Subscription ID: 00000000-0000-0000-0000-000000000000

Select lookback period:
    [1] Last7Days
  * [2] Last30Days
    [3] Last60Days
Enter choice 1-3 (default 2): <Enter>

Select term:
  * [1] P1Y
    [2] P3Y
Enter choice 1-2 (default 1): <Enter>

Output folder (default: C:\Data\Github\savingsplan-api): <Enter>

CSV format:
  * [1] US (comma, ISO datetime yyyy-MM-ddTHH:mm)
    [2] DE (semicolon, datetime dd.MM.yyyy HH:mm)
Enter choice 1-2 (default 1): <Enter>

Acquiring ARM access token...
Verifying API access and fetching recommendations...
  Page 1...
Retrieved 12 recommendation(s).

Writing AllRecommendationDetails CSV (top recommendation's allSavingsBenefitDetails collection)...
  Wrote 8 row(s) to C:\...\2026-06-11-SUB-0000...-Last30Days-P1Y-AllRecommendationDetails.csv

Writing TopRecommendation markdown...
  Wrote C:\...\2026-06-11-SUB-0000...-Last30Days-P1Y-TopRecommendation.md

Writing HourlyUsage CSV (top recommendation)...
  Wrote 720 hourly row(s) to C:\...\2026-06-11-SUB-0000...-Last30Days-P1Y-HourlyUsage.csv

Done.
```

## Troubleshooting

- **`Az.Accounts module is not installed.`**
  Install it: `Install-Module Az.Accounts -Scope CurrentUser`.

- **`ARM call failed: Authentication failed.` (HTTP 401)**
  Token expired or you're not signed in: run `Connect-AzAccount`.

- **`ARM call failed: Authorization failed.` (HTTP 403)**
  Your account lacks the required role on the chosen scope. Ask for
  *Cost Management Reader* (or *Billing Reader* at the billing account level).

- **`ARM call failed: Scope not found.` (HTTP 404)**
  Double-check the IDs you entered, and that your current Az context points at
  the right tenant (`Get-AzContext`, `Connect-AzAccount -Tenant <tenantId>`).

- **`No recommendations returned ... Nothing to export.`**
  The API simply has nothing to recommend for that combination of scope,
  lookback and term right now. Try a longer lookback (`Last60Days`) or a
  broader scope.

- **Hourly CSV shorter than expected.**
  Expected: the CSV covers exactly
  `properties.firstConsumptionDate` → `properties.lastConsumptionDate`. The API
  typically lags wall-clock time by several hours/days, so a `Last60Days` export
  may legitimately contain ~1,400+ rows rather than the full 60 × 24 = 1,440.
  Check `properties.totalHours` and `properties.lastConsumptionDate` in the raw
  JSON (export with the "Export raw JSON?" prompt set to *Yes*) to confirm the
  window your data actually covers.

## Versioning

Script version is in the header of `Export-BenefitRecommendations.ps1`
(format `v1.0.YYYY-MM-DD.0`). Bump the major version on incompatible CLI/output
changes; the date segment tracks the last edit.

## References

- [Benefit Recommendations - List (REST API 2025-03-01)](https://learn.microsoft.com/rest/api/cost-management/benefit-recommendations/list?view=rest-cost-management-2025-03-01)
- [Az.Accounts module](https://learn.microsoft.com/powershell/module/az.accounts/)

---

## Azure Savings Plan Optimizer (`Optimize-SavingsPlanCommitment.ps1`)

`Optimize-SavingsPlanCommitment.ps1` analyses Azure **hourly PayGo usage**
exports produced by `Export-BenefitRecommendations.ps1` (the
`*-HourlyUsage.csv` files) and recommends two Savings Plan (SP) commitment
values per file:

1. **Waste-free max commitment** — the largest hourly commitment that produces
   *zero unused SP* in the observed window (= the minimum hourly usage value).
2. **Optimum commitment per average SP discount** — the hourly commitment that
   maximises total savings versus pure PayGo. For an average discount `d`
   over `N` hours of usage `U(h)` the cost is

   ```
   Cost(K, d) = N · K · (1 − d) + Σ max(0, U(h) − K)
   ```

   The analytic optimum is `K*(d) = quantile(U, d)` (i.e. the `d`-th
   percentile of hourly usage).

All commitment values are expressed in **PayGo-equivalent USD/hour**, which
is directly comparable to the hourly usage column in the input CSV and
matches how the Azure SP recommender reports its suggestion.

### Analyzer inputs

Place one or more **`*HourlyUsage.csv`** exports (as produced by
`Export-BenefitRecommendations.ps1`) in a folder. Expected filename pattern:

```
{yyyy-MM-dd}-{EA|MCA|SUB|RG}-{id1}[-{id2}]-Last{30|60}Days-P{1|3}Y-HourlyUsage.csv
```

The script extracts:

| Token         | Meaning                                                                  |
|---------------|--------------------------------------------------------------------------|
| `yyyy-MM-dd`  | Export date (used in section sub-headings and output filename).          |
| `EA` / `MCA` / `SUB` / `RG` | Scope kind (billing account, billing profile, subscription, resource group). |
| `id1[-id2]`   | Scope IDs (billing account [+ billing profile], subscription [+ RG]).    |
| `Last30Days` / `Last60Days` | Lookback window — drives the top-level report grouping (`## 30 day recommendation` / `## 60 day recommendation`). |
| `P1Y` / `P3Y` | SP **term** length (used to project term-level savings).                 |

By convention the analyzer's default input folder is `HourlyUsage/` next to the
script (the script accepts any folder/file via `-Path`).

#### Locale auto-detection

Each file is inspected independently and parsed accordingly:

| Locale  | Decimal sep | Date format            | Default delimiter |
|---------|-------------|------------------------|-------------------|
| `de-DE` | comma       | `dd.MM.yyyy HH:mm`     | `;`               |
| `en-US` | dot         | `MM/dd/yyyy HH:mm` or ISO `yyyy-MM-dd HH:mm` | `,` |

Delimiter is sniffed from the first non-empty line; decimal/date locale is
inferred from the first data row.

CSV header must be `DateTime;HourlyPayGoUsage` (or the comma-delimited
equivalent).

### Analyzer CLI usage

```powershell
.\Optimize-SavingsPlanCommitment.ps1 [-Path <path>] [-Discounts 30,32,34,36,38,40,42] [-Out <file>]
```

* `-Path` — a single CSV file or a folder containing `*HourlyUsage.csv`
  files. Default: the `HourlyUsage` subfolder next to the script.
* `-Discounts` — comma-separated *average* SP discounts to model. Whole
  percent (`30`) and fractions (`0.30`) are both accepted. Default:
  `30,32,34,36,38,40,42` — a range of **common average** SP discounts
  spanning **conservative (≈30 %)** to **optimistic (≈42 %)**.

  > Azure Savings Plan discounts vary by SKU, region, and reservation
  > footprint — there is no single "SP discount" for a scope. The default
  > range therefore produces a sensitivity band rather than a single point
  > estimate; pass your own list (e.g. `-Discounts 20,25,30`) if your
  > workload mix justifies a different range.
* `-Out` — output Markdown file **or destination folder**. When omitted,
  the filename is auto-derived from input metadata in the form
  `{yyyy-MM-dd}-{scope}-Last{N}Days-P{T}Y-SavingsPlanRecommendation.md`
  (mirroring the input file convention) and dropped next to the script.
  Shared tokens across multiple inputs are kept; differing tokens are
  joined with `+` (e.g. `Last30+60Days`, `P1+3Y`); missing tokens fall
  back to today's date / `MultiScope` / `LastUnknownDays` / `PUnknownY`.
  If `-Out` is an existing directory, ends with `\` / `/`, or has no
  file extension, it is treated as a folder and the auto-derived
  filename is written inside it (the folder is created if missing).

Examples:

```powershell
.\Optimize-SavingsPlanCommitment.ps1                                      # scan default folder
.\Optimize-SavingsPlanCommitment.ps1 -Path .\HourlyUsage
.\Optimize-SavingsPlanCommitment.ps1 -Path . -Discounts 20,22,24,30,35,40
.\Optimize-SavingsPlanCommitment.ps1 -Path .\HourlyUsage\2026-06-11-SUB-…-Last30Days-P1Y-HourlyUsage.csv
.\Optimize-SavingsPlanCommitment.ps1 -Path .\HourlyUsage -Out .\my-report.md   # explicit output path
```

### Sample output

<details>
<summary>Click to expand a sample <code>*-SavingsPlanRecommendation.md</code> report</summary>

```markdown
# Azure Savings Plan — Optimization Results

_Generated: 2026-06-17 10:45_
_Script: `Optimize-SavingsPlanCommitment.ps1` v5.2026-06-17.0_
_Discounts analysed: 30%, 32%, 34%, 36%, 38%, 40%, 42%_
_Commitment K reported in **PayGo-equivalent USD/hour** (matches Azure SP recommender output)._

> **About the discount range.** Azure Savings Plan discounts vary by SKU, region, and
> reservation footprint — there is no single "SP discount" for a given scope. The default
> range spans common average SP discounts from **conservative (30%)** to **optimistic (42%)**,
> so the recommendation reads as a *sensitivity band* rather than a single point estimate.

> **How to read the Savings and Waste columns.**
>
> * **Savings** = PayGo cost − SP cost. This is **already net** (waste is baked in).
> * **Waste (USD)** = unused commitment capacity × (1 − discount).
> * Savings and waste are **not additive** — waste is a subset of the commitment cost.
> * **Hrs w/ overflow /mo** and **Hrs w/ waste /mo** are frequency counts (how often,
>   not duration).

## Inputs

| File | Date | Scope | Lookback | Term | Locale |
|---|---|---|---|---|---|
| `2026-06-11-MCA-12345678-BP1-Last60Days-P3Y-HourlyUsage.csv` | 2026-06-11 | Billing Account / Billing Profile `12345678-BP1` | 60 d | 3 y | en-US |

## 60 day recommendation

### 2026-06-11 — Billing Account / Billing Profile 12345678-BP1

| Metadata | Value |
|---|---|
| Source file | `2026-06-11-MCA-12345678-BP1-Last60Days-P3Y-HourlyUsage.csv` |
| Export date (filename) | 2026-06-11 |
| Scope (filename) | Billing Account / Billing Profile — `12345678-BP1` |
| Locale detected | `en-US` (delimiter `,`, date `yyyy-MM-ddTHH:mm`) |
| Lookback (filename) | 60 days |
| SP term (filename) | 3 year(s) — 26,298 h |
| Observed rows | 1,440 hours (60.00 days) |
| Date range | 2026-04-11 00:00 → 2026-06-09 23:00 |
| PayGo total (observed) | $36,803.89 |
| PayGo /hour | min $13.50 · mean $25.56 · max $36.47 · σ $7.80 |
| Projected PayGo over term | $672,130.99 |
| Reporting month | 8,766 / 12 = 730.5 hours (average Gregorian year / 12) |

### (a) Waste-free max commitment

`K_max_no_waste` = min hourly PayGo usage = **$13.50 / hour** (PayGo-equivalent).

Commitment value over 3-year: **$355,033.52**.

| Discount | Monthly savings | Monthly waste | 3-year savings | 3-year waste |
|---|---:|---:|---:|---:|
| 30% | $2,958.61 | $0.00 | $106,510.06 | $0.00 |
| 34% | $3,353.09 | $0.00 | $120,711.40 | $0.00 |
| 38% | $3,747.58 | $0.00 | $134,912.74 | $0.00 |
| 42% | $4,142.06 | $0.00 | $149,114.08 | $0.00 |

### (b) Optimum commitment per discount

`K*(d) = quantile(U, d)` — analytic optimum.

| Discount | K* /hr | Commitment /term | Monthly savings | Monthly waste | 3-yr savings | 3-yr waste | Hrs w/ overflow /mo | Overflow % | Hrs w/ waste /mo | Waste % |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| _(waste-free)_ | $13.50 | $355,034 | — | — | — | — | 730.0 | 99.9% | 0.0 | 0.0% |
| 30% | $18.72 | $492,283 | $3,519 | $409 | $126,676 | $14,706 | 511.4 | 70.0% | 219.2 | 30.0% |
| 34% | $19.47 | $511,958 | $4,076 | $501 | $146,719 | $18,049 | 481.9 | 66.0% | 248.6 | 34.0% |
| 38% | $20.24 | $532,242 | $4,656 | $596 | $167,621 | $21,471 | 453.0 | 62.0% | 277.5 | 38.0% |
| 42% | $20.92 | $550,036 | $5,257 | $673 | $189,257 | $24,220 | 423.6 | 58.0% | 306.9 | 42.0% |
```

**Key takeaway from this sample:** at a 36 % average SP discount, committing **$19.88 /hr**
(~$540 k over 3 years) yields **$4,363 /month net savings** versus pure PayGo — with
about 36 % of hours experiencing some waste (worth ~$554 /month, already deducted from
the savings figure).

</details>

### Analyzer outputs

* **`{date}-{scope}-Last{N}Days-P{T}Y-SavingsPlanRecommendation.md`** —
  combined Markdown report. Inputs are grouped by lookback window: a
  top-level `## 30 day recommendation` and/or `## 60 day recommendation`
  heading, then one `### {yyyy-MM-dd} — {scope label}` sub-section per
  file containing:
  * Metadata block (filename, locale, lookback, term, observed hours,
    PayGo totals, scope kind / IDs).
  * Table (a): waste-free max commitment with monthly and per-term savings
    and waste at every requested discount.
  * Table (b): optimum commitment per discount, with `K*`, total commitment
    value over the term, monthly and per-term savings and waste, and
    **monthly overflow/waste hours with percentages**. The waste-free max
    commitment from table (a) appears as the first row for reference.
* **Console** — compact mirror of the same per-file numbers, with the date
  + scope label on the leading line and the filename on the line below for
  traceability.

#### Reporting basis

* Hours per month = `8766 / 12 = 730.5`.
* Hours per year = `8766` (average Gregorian year).
* Hours per term = `years × 8766` (`1Y = 8766`, `3Y = 26 298`).
* Per-hour rates are computed from the observed lookback window and then
  scaled to monthly and term totals. This is a linear extrapolation and
  assumes the lookback is representative of the full term.

### Analyzer assumptions and caveats

* **Cost-only basis (not Azure's native recommendation).** The analyzer works
  **purely from the hourly PayGo cost** (`HourlyPayGoUsage` = `properties.usage.charges`).
  That series is cost in USD/hour and **ignores the underlying resource usage**
  (consumed quantity) that produced the cost. Azure's native Savings Plan
  recommendations — the `allSavingsBenefitDetails` entries under
  `properties.allRecommendationDetails` — are computed from **both cost and
  usage**. A recommendation derived from cost alone is therefore an
  **approximation** and may, for some workload mixes, point to a different
  commitment than Azure's usage-aware recommendation. Treat the output as
  **directional** and cross-check against the `AllRecommendationDetails` export
  and Azure's native recommendation before purchasing.
* Commitment `K` is treated as **PayGo-equivalent USD/hour**.
* **Azure SP discounts vary by SKU**, region, and reservation footprint.
  The average discount `d` is a flat blended rate applied to every covered
  hour; the default `-Discounts` range (≈30 % → ≈42 %) is a *sensitivity
  band* of common average discounts (conservative → optimistic), not a
  set of quotable rates.
* `K_max_no_waste` is waste-free **only over the observed window**; future
  hours below the observed minimum will produce some waste.
* No purchase-side constraints (Azure SP commitments are entered to three
  decimals). Round the recommendation as needed when buying.
* The 30- or 60-day lookback may not capture seasonal peaks/troughs — treat
  the term projection as directional, not contractual.

### Analyzer files

```
Optimize-SavingsPlanCommitment.ps1            — the analyser
*-SavingsPlanRecommendation.md                — generated reports (auto-named; gitignored)
HourlyUsage/                                  — optional folder of Azure SP recommender exports (any lookback/term; gitignored)
```
