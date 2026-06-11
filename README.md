v1.0.0 Dirk Brinkmann 2026-06-11

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
| `HourlyPayGoUsage`| Value from `properties.usage.charges`, aligned via `firstConsumptionDate`. |

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
