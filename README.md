# Entra ID Guest Report

Generates an interactive, self-contained HTML cleanup report for every guest user (and external B2B member) in a Microsoft Entra ID tenant. Each account gets a **Remove / Review / Keep** recommendation with the reasons behind it. The report is designed for guest lifecycle reviews and access governance, and can run as a weekly automated report in Azure DevOps.

## What it does

The script connects to Microsoft Graph, retrieves every guest account in the tenant and analyzes each one across several dimensions:

- **Invitation state**: accepted, pending (with the number of days since the invitation was sent), or no invitation state (cross-tenant sync, on-premises synced)
- **Sign-in activity**: last successful sign-in, last interactive and non-interactive sign-in attempts, days since last activity. This needs an Entra ID P1/P2 license; see [Sign-in activity](#sign-in-activity).
- **Access**:
  - **Groups**: direct memberships, with group type (Microsoft 365, Security, Mail-enabled security, Distribution, dynamic) and a role-assignable flag
  - **Teams**: Microsoft 365 groups with a team provisioned, with visibility
  - **Applications**: direct app role assignments, with the role name resolved
  - **Directory roles**: Entra ID roles assigned directly to the guest, or through a role-assignable group the guest is a member of, including administrative-unit scope
- **Sponsors**: the users or groups recorded as responsible for the guest
- **Identity provider**: Entra ID, Microsoft Account, Email one-time passcode, Google, other federated IdPs, or not redeemed
- **Guest domain**: the guest's home domain, read from the mail address or the `#EXT#` user principal name
- **User type**: B2B guests (`userType = Guest`) and external members (`userType = Member` with an `#EXT#` UPN or an external user state, e.g. from cross-tenant synchronization)
- **Account state**: enabled/disabled, on-premises synced, creation date (New Guest / Last Year / +1 Year buckets)
- **Cleanup recommendation**: Remove / Review / Keep (or On-Prem). See [Cleanup recommendation](#cleanup-recommendation).

The output is a single `.html` file that works offline, with no external dependencies and no server required.

## Cleanup recommendation

Guest accounts don't carry permissions the way applications do, so the report doesn't compute a risk score. Instead, each account gets one actionable recommendation, built from the rules below and evaluated in order. Every rule that applies is listed in the recommendation modal, so the reasoning stays visible.

| Situation | Recommendation |
|-----------|----------------|
| Invitation pending for `-PendingInvitationDays` (default 30) days or more | **Remove** |
| Invitation pending for less than that | Keep |
| Created less than 30 days ago | Keep (grace period for new guests) |
| Sign-in activity unavailable (no P1/P2 license) | Review |
| Never signed in, `-InactiveDays` (default 90) or more after redeeming the invitation / being created | **Remove** |
| Never signed in, less than that | Review |
| No sign-in for `-RemoveAfterDays` (default 365) days or more | **Remove** |
| No sign-in for `-InactiveDays` days or more | Review |
| Signed in within `-InactiveDays` days | Keep |

Modifiers applied afterwards:

| Signal | Effect |
|--------|--------|
| Account disabled and inactive / never signed in | Review → **Remove** |
| Account disabled but recently active or new | Keep → Review |
| Holds one or more directory roles (directly or through a role-assignable group) | Keep → Review, and always listed as a reason |
| Synchronized from on-premises Active Directory | **On-Prem**: changes must be made in the source directory |
| No sponsor, or no group/Teams/app/role access at all | Listed as information; doesn't change the recommendation |

The recommendation modal also states the suggested action. For **Remove**: delete the account; deleted users can be restored from **Deleted users** within 30 days. For **Review**: confirm with the sponsor or resource owner, or use an Access Review.

### Sign-in activity

Activity is based on `signInActivity.lastSuccessfulSignInDateTime`. Entra ID only records successful sign-ins since December 2023. For accounts with no successful sign-in recorded, the report falls back to the most recent interactive or non-interactive sign-in **attempt**, which may have failed. Those dates are marked with `*` in the table and explained in the sign-in modal.

Reading `signInActivity` requires an **Entra ID P1 or P2** license and the `AuditLog.Read.All` permission. If the tenant can't return sign-in data, the script logs a warning and continues without it. The report then shows a banner, accounts that would be assessed on inactivity are marked **Review**, and invitation-based rules still apply.

## Report features

| Feature | Detail |
|---------|--------|
| Summary cards | Clickable cards for total guests, Remove, Review, Keep, pending invitations, never signed in, inactive, disabled, guests with directory roles, and guests without a sponsor, each filtering the table |
| Search | Debounced real-time text search across name, UPN, mail, domain, company, sponsor, group, team, app and role names |
| Filter panel | Multi-dimensional filter tags for recommendation, last sign-in bucket, invitation state, enabled state, creation date bucket, user type, identity provider, access type, sponsor and source |
| Column sort | Click any sortable column header to sort ascending; click again to reverse |
| Export CSV | Downloads the currently visible (filtered) rows as a `.csv` file, UTF-8 with BOM for correct Excel rendering. The export includes extra columns not shown in the table: object ID, UPN, mail, company, days since last sign-in, group/team/app/role/sponsor names and the recommendation reasons |
| Dark / light mode | Toggle persisted to `localStorage`; a GitHub repo link icon sits next to the toggle in the header |
| Portal deep links | Guest names link to the user in the Entra portal; groups, Teams and apps in the modals link to their portal pages |
| Detail modals | Clickable badges open a detail panel; see [Interactive badges](#interactive-badges) |

## Interactive badges

Badges that only act as filters (User Type, Identity, Invitation, Enabled, New, No sponsor) add a filter when clicked. Modal tables are sortable: click any column header to sort ascending, click again to reverse.

| Column | Badge | Modal content |
|--------|-------|---------------|
| Last Sign-In | (date) | Last successful sign-in, last interactive and non-interactive attempts, created date, invitation state and change date, identity provider, creation type, company, UPN, mail and object ID |
| Access | Groups: N | Direct group memberships with type and role-assignable flag; each group links to the Entra portal |
| Access | Teams: N | Teams the guest is a member of, with visibility |
| Access | Apps: N | Applications assigned directly to the guest, with the assigned role and assignment date |
| Access | Roles: N | Directory roles with assignment path (Direct or via group) and scope; each role links to the Microsoft Learn built-in roles reference |
| Sponsors | N sponsor(s) | Sponsor name, UPN/mail and type (user or group) |
| Recommendation | Remove / Review / Keep / On-Prem | Suggested action, every reason that contributed, and the thresholds used |
| Guest | New | Plain filter badge for guests created in the last 30 days |

## Prerequisites

### PowerShell version

PowerShell 5.1 or PowerShell 7+.

### Microsoft Graph module

The script calls Microsoft Graph directly through `Invoke-MgGraphRequest`, so only the authentication module is required (version 2.0.0 or later):

```powershell
Install-Module -Name Microsoft.Graph.Authentication -MinimumVersion '2.0.0' -Scope CurrentUser -AllowClobber
```

The full `Microsoft.Graph` umbrella module works too.

### Entra ID permissions

The identity used to run the script (user, service principal, or managed identity) needs these **Microsoft Graph application permissions** with admin consent:

| Permission | Purpose |
|------------|---------|
| `User.Read.All` | Read guest and external member accounts |
| `AuditLog.Read.All` | Read sign-in activity (also requires an Entra ID P1/P2 license) |
| `Directory.Read.All` | Read group memberships, app role assignments, sponsors and application roles |
| `RoleManagement.Read.Directory` | Read directory role assignments |

For **interactive (user) runs**, these are requested as delegated scopes during sign-in. Reading sign-in activity with delegated permissions also requires the signed-in user to hold a role that can read sign-in reports, such as Reports Reader, Security Reader or Global Reader.

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-OutputPath` | String | Auto-generated | Path for the HTML report. Defaults to `EntraIDGuestReport__{TenantName}_{Date}.html` in the current directory. |
| `-TenantId` | String | None | Entra ID tenant ID. Required for Service Principal authentication. |
| `-AccessToken` | SecureString | None | Pre-acquired Microsoft Graph token. Used in Azure DevOps pipelines via `Get-AzAccessToken`. |
| `-ClientId` | String | None | App (client) ID for Service Principal authentication. Must be combined with `-CertificateThumbprint` and `-TenantId`, or with `-UseManagedIdentity` to select a user-assigned Managed Identity. |
| `-CertificateThumbprint` | String | None | Certificate thumbprint for Service Principal authentication. Requires `-ClientId`. |
| `-UseManagedIdentity` | Switch | None | Authenticate using the Managed Identity of the hosting environment. System-assigned by default; add `-ClientId` for a user-assigned identity. |
| `-NonInteractive` | Switch | None | Suppress all prompts and skip automatic browser launch. Required for unattended or pipeline runs. |
| `-TargetUser` | String | None | Analyze a single guest or external member by object ID or user principal name. Skips tenant enumeration and the confirmation prompt. |
| `-ExcludeExternalMembers` | Switch | None | Only report `userType = Guest` accounts; skip external members. |
| `-SkipAccessDetails` | Switch | None | Skip group, Teams, application and sponsor collection. Directly assigned directory roles are still reported. Useful on very large tenants. |
| `-InactiveDays` | Int | `90` | Days without a sign-in after which a guest is marked for Review. |
| `-RemoveAfterDays` | Int | `365` | Days without a sign-in after which a guest is marked for Remove. Must be greater than `-InactiveDays`. |
| `-PendingInvitationDays` | Int | `30` | Days after which an unredeemed invitation is marked for Remove. |
| `-DryRun` | Switch | None | Run the full analysis but skip writing the report to disk and opening the browser. |

## Manual usage

### Interactive sign-in (default tenant)

```powershell
.\Export-EntraGuestReport.ps1
```

A browser window opens for Microsoft sign-in. The report is saved to the current directory and opened automatically when complete.

### Interactive sign-in for a specific tenant

```powershell
.\Export-EntraGuestReport.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

### Service Principal with certificate

```powershell
.\Export-EntraGuestReport.ps1 `
    -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12" `
    -NonInteractive `
    -OutputPath "C:\Reports\GuestReport.html"
```

### Managed Identity

```powershell
# System-assigned Managed Identity
.\Export-EntraGuestReport.ps1 -UseManagedIdentity -NonInteractive -OutputPath ".\report.html"

# User-assigned Managed Identity (select by its client ID)
.\Export-EntraGuestReport.ps1 -UseManagedIdentity -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" -NonInteractive -OutputPath ".\report.html"
```

### Custom thresholds

```powershell
# Review after 60 days without sign-in, remove after 180 days, expire invitations after 14 days
.\Export-EntraGuestReport.ps1 -InactiveDays 60 -RemoveAfterDays 180 -PendingInvitationDays 14
```

### Analyze a single guest

```powershell
.\Export-EntraGuestReport.ps1 -TargetUser "john_contoso.com#EXT#@fabrikam.onmicrosoft.com"
```

### Large tenants

Group, application and sponsor details are read through Microsoft Graph JSON batching (20 requests per call, three requests per guest). On tenants with tens of thousands of guests, use `-SkipAccessDetails` for a faster run that still covers activity, invitation state and direct directory roles.

## Azure DevOps pipeline

The included `azure-pipelines.yml` automates weekly report generation using a federated Azure service connection. No secrets or passwords are stored in ADO.

### How it works

1. The pipeline runs on a schedule (every Monday at 06:00 UTC) and on every push to `main`.
2. The `Microsoft.Graph.Authentication` module is restored from a pipeline cache (`Cache@2`). On the first run it is downloaded once with `Save-PSResource` and saved back to the cache.
3. The `AzurePowerShell@5` task authenticates using the ADO service connection (`azureSubscription`), which uses **Workload Identity Federation**. The task retries once on transient failures.
4. The task obtains a Graph access token from the authenticated Az context and passes it to the script via `-AccessToken`.
5. The HTML report is saved to `$(Build.ArtifactStagingDirectory)` and published as a pipeline artifact named `EntraIDGuestReport`.
6. The run gets a **markdown summary tab** with the recommendation and signal counts. If any guest holds an Entra ID directory role, the script logs a pipeline warning and the run completes as **SucceededWithIssues** (orange). These signals are emitted only when the script detects an ADO agent via `TF_BUILD`; local runs are unaffected.

> **Token lifetime:** the Graph token obtained via the service connection is valid for roughly an hour and can't be refreshed mid-run. Directory role assignments are pre-fetched in bulk and per-guest details use JSON batching to stay well within that window. On very large tenants, add `-SkipAccessDetails` to the script call in the pipeline.

> **Personal data:** the report contains names and email addresses of external users. Set an artifact retention period that matches your data protection requirements.

### Import the repository into Azure DevOps

1. In your ADO project: **Repos → Files**, then click **Import** (on a brand-new empty repo you'll see an **Import a repository** button instead).
2. **Clone URL**: `https://github.com/matejklemencic/EntraID-GuestReport`. Leave **Requires authorization** unchecked; this repo is public.
3. Click **Import**.

### Setup steps

**1. Create an Azure service connection in ADO**

**Project Settings → Service Connections → New service connection → Azure Resource Manager**. Select **Workload Identity Federation** with **automatic** app registration, and pick the subscription to scope the connection to.

**2. Grant the service principal Graph permissions**

Find the service principal created for your service connection in Entra ID (**App registrations**) and add these **Application permissions** with admin consent:

- `User.Read.All`
- `AuditLog.Read.All`
- `Directory.Read.All`
- `RoleManagement.Read.Directory`

**3. Update the pipeline file**

Edit `azure-pipelines.yml`, set `azureSubscription` to the name of your service connection, then commit and push the change to `main`:

```yaml
azureSubscription: "Your-Service-Connection-Name"
```

**4. Create the pipeline in ADO**

**Pipelines → New pipeline → Azure Repos Git → select the repo → Existing Azure Pipelines YAML file → `/azure-pipelines.yml`** → **Save** (or **Save and run**).

**5. Run or schedule**

Push to `main` to trigger immediately, or let the Monday schedule run automatically. The report is available under the pipeline run's **Artifacts** tab.

> **First-run authorization:** if the service connection wasn't granted access to all pipelines, the first run pauses with a **"This pipeline needs permission to access a resource"** banner. Open the run and click **Permit**.

> **Hosted agent parallelism:** new ADO organizations often start with 0 free hosted parallel jobs, so the run queues indefinitely. Link an Azure subscription for billing under **Organization Settings → Billing** to get the free tier, or use the [request form](https://aka.ms/azpipelines-parallelism-request).

### Pipeline output

The artifact `EntraIDGuestReport` contains:

```
EntraIDGuestReport_YYYY-MM-DD.html
```

Download the artifact from the pipeline run and open the HTML file in any browser. No internet connection is required.

## Upgrading from version 2.0

Version v26.10.26 is a rewrite. Breaking changes:

- **Module**: uses `Microsoft.Graph.Authentication` instead of `Microsoft.Entra`.
- **Output**: a single interactive HTML report. The server-side CSV and `_Summary.txt` files are gone; use the **Export CSV** button in the report instead (it exports the filtered rows).
- **Parameters removed**: `-CSVOnly`, `-HTMLOnly`, `-FileName`, `-ExportPath` (use `-OutputPath`) and `-PendingAcceptanceOnly` (use the Invitation filter in the report).
- **Activity** is now based on the last *successful* sign-in, not the last sign-in attempt.
- **External members** (`userType = Member`) are included by default; use `-ExcludeExternalMembers` for the previous behavior.

## File reference

| File | Purpose |
|------|---------|
| `Export-EntraGuestReport.ps1` | Main script |
| `azure-pipelines.yml` | Azure DevOps pipeline definition |

## Author

[Matej Klemencic](https://www.matej.guru)
