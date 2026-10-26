<#
.SYNOPSIS
    Generates an interactive HTML cleanup report for Microsoft Entra ID guest users.
.DESCRIPTION
    This script connects to Microsoft Graph, retrieves every guest user (and, by default, every external
    B2B member), analyzes invitation state, sign-in activity, group / Teams / application access, directory
    roles and sponsors, and produces a filterable interactive HTML report with a cleanup recommendation
    (Remove / Review / Keep) for each account, detail modals and deep links to the Entra portal.
.AUTHOR
    Matej Klemencic (www.matej.guru)
.NOTES
    Version:        v26.10.26

    Installation: Install the Microsoft Graph authentication module (v2.0.0 or later). The script talks to
    Microsoft Graph directly through Invoke-MgGraphRequest, so no other Graph module is required:
    Install-Module -Name Microsoft.Graph.Authentication -MinimumVersion '2.0.0' -Scope CurrentUser -AllowClobber
.OUTPUTS
    None. Writes an HTML report to the file specified by -OutputPath.
.PARAMETER OutputPath
    Path to save the generated HTML report. Defaults to "EntraIDGuestReport__{TenantName}_{Date}.html".
.PARAMETER TenantId
    Optional. The Entra ID tenant ID to connect to. Required when using Service Principal authentication.
.PARAMETER AccessToken
    Optional. A pre-acquired Microsoft Graph access token (SecureString). Use this when the pipeline
    already has an authenticated Az context (e.g. AzurePowerShell@5 with azureSubscription).
    Obtain with: (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString).Token
.PARAMETER ClientId
    Optional. The Application (client) ID of a Service Principal for unattended/pipeline authentication.
    Must be combined with -CertificateThumbprint and -TenantId. Alternatively, combine with
    -UseManagedIdentity to select a user-assigned Managed Identity by its client ID.
.PARAMETER CertificateThumbprint
    Optional. Certificate thumbprint for Service Principal authentication. Use with -ClientId and -TenantId.
.PARAMETER UseManagedIdentity
    Optional. Use the Azure Managed Identity of the hosting environment (e.g. Azure VM, Azure Function,
    Azure Automation). No credentials required. Uses the system-assigned identity by default;
    combine with -ClientId to select a user-assigned Managed Identity.
.PARAMETER NonInteractive
    Optional. Suppress all interactive prompts. Required for unattended/pipeline runs.
    Also suppresses automatic browser launch after report generation.
.PARAMETER TargetUser
    Optional. Analyze a single guest or external member instead of the whole tenant. Accepts the
    object ID or the user principal name. Skips the tenant enumeration and the confirmation prompt.
.PARAMETER ExcludeExternalMembers
    Optional. Only report users with userType 'Guest'. By default, external B2B users with userType
    'Member' (e.g. from cross-tenant synchronization or invited as members) are included as well.
.PARAMETER SkipAccessDetails
    Optional. Skip collecting group, Teams, application and sponsor details. Directly assigned directory
    roles are still reported. Use on very large tenants to shorten the run.
.PARAMETER InactiveDays
    Number of days without a sign-in after which a guest is considered inactive and marked for Review.
    Default is 90.
.PARAMETER RemoveAfterDays
    Number of days without a sign-in after which a guest is marked for Remove. Must be greater than
    -InactiveDays. Default is 365.
.PARAMETER PendingInvitationDays
    Number of days after which an unredeemed invitation is marked for Remove. Default is 30.
.PARAMETER DryRun
    Optional. When set, skips writing the HTML report to disk and skips launching the browser.
    Useful for testing parameter validation and authentication without side effects.
.EXAMPLE
    # Interactive: generate a report for the default tenant
    .\Export-EntraGuestReport.ps1
.EXAMPLE
    # Interactive: generate a report for a specific tenant
    .\Export-EntraGuestReport.ps1 -TenantId "your-tenant-id"
.EXAMPLE
    # Azure DevOps with azureSubscription service connection (recommended for ADO)
    $token = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString).Token
    .\Export-EntraGuestReport.ps1 -AccessToken $token -NonInteractive -OutputPath "report.html"
.EXAMPLE
    # Service Principal with certificate
    .\Export-EntraGuestReport.ps1 -TenantId "tenant-id" -ClientId "app-id" -CertificateThumbprint "thumbprint" -NonInteractive
.EXAMPLE
    # Managed Identity (Azure-hosted automation)
    .\Export-EntraGuestReport.ps1 -UseManagedIdentity -NonInteractive -OutputPath ".\report.html"
.EXAMPLE
    # Stricter thresholds: review after 60 days, remove after 180 days
    .\Export-EntraGuestReport.ps1 -InactiveDays 60 -RemoveAfterDays 180
.EXAMPLE
    # Quickly analyze a single guest
    .\Export-EntraGuestReport.ps1 -TargetUser "john_contoso.com#EXT#@fabrikam.onmicrosoft.com"
    #>

[CmdletBinding()]
param(
    [string]$OutputPath = "",
    [string]$TenantId = "",
    # Pre-acquired token (e.g. from Get-AzAccessToken in AzurePowerShell@5 task)
    [System.Security.SecureString]$AccessToken,
    # Service Principal / pipeline authentication
    [string]$ClientId = "",
    [string]$CertificateThumbprint = "",
    [switch]$UseManagedIdentity,
    [switch]$NonInteractive,
    # Scope
    [string]$TargetUser = "",
    [switch]$ExcludeExternalMembers,
    [switch]$SkipAccessDetails,
    # Cleanup thresholds (days)
    [ValidateRange(1, 3650)]
    [int]$InactiveDays = 90,
    [ValidateRange(1, 3650)]
    [int]$RemoveAfterDays = 365,
    [ValidateRange(1, 3650)]
    [int]$PendingInvitationDays = 30,
    [switch]$DryRun
)

# Validate mutually exclusive / dependent parameters
if ($RemoveAfterDays -le $InactiveDays) {
    Write-Error "-RemoveAfterDays ($RemoveAfterDays) must be greater than -InactiveDays ($InactiveDays)."
    exit 1
}
if ($CertificateThumbprint -and -not $ClientId) {
    Write-Error "-CertificateThumbprint requires -ClientId (Service Principal authentication)."
    exit 1
}
if ($ClientId -and -not $CertificateThumbprint -and -not $UseManagedIdentity) {
    Write-Error "-ClientId requires either -CertificateThumbprint (Service Principal authentication) or -UseManagedIdentity (user-assigned Managed Identity)."
    exit 1
}
if ($ClientId -and $CertificateThumbprint -and -not $TenantId) {
    Write-Error "-TenantId is required when using Service Principal authentication (-ClientId/-CertificateThumbprint)."
    exit 1
}

# Make all unhandled errors terminating. The main body below is wrapped in try/finally with
# no catch, so an unhandled error still flows through the finally block (guaranteeing
# Disconnect-MgGraph runs) before it terminates the script.
$ErrorActionPreference = 'Stop'

# Unattended runs: rendering Write-Progress is pure overhead and noise in pipeline logs
if ($NonInteractive) { $ProgressPreference = 'SilentlyContinue' }

# Script version - keep in sync with the .NOTES header above; surfaced in the report footer
$ScriptVersion = 'v26.10.26'

# Guests created within this many days are never recommended for removal (still settling in)
$NewGuestDays = 30

# Properties requested for every reported user. signInActivity is appended when the tenant
# supports it (Entra ID P1/P2 + AuditLog.Read.All); see Invoke-UserQuery.
$userSelectBase = 'id,displayName,userPrincipalName,mail,otherMails,userType,createdDateTime,creationType,externalUserState,externalUserStateChangeDateTime,accountEnabled,onPremisesSyncEnabled,companyName,identities'
$script:signInDataAvailable = $true
$script:accessReadFailures  = 0

# Function to safely import modules
function Import-GraphModuleSafely {
    param([string]$ModuleName)

    # Reuse an already-loaded module if it meets the minimum version requirement
    $loaded = Get-Module -Name $ModuleName
    if ($loaded) {
        if ($loaded.Version -ge [version]'2.0.0') {
            Write-Verbose "$ModuleName $($loaded.Version) already loaded - skipping re-import"
            return
        }
        Write-Warning "$ModuleName $($loaded.Version) is loaded but v2.0.0+ is required - re-importing"
    }

    try {
        Import-Module $ModuleName -MinimumVersion '2.0.0' -ErrorAction Stop
        Write-Host "Successfully imported $ModuleName" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not import $ModuleName v2+. Attempting Microsoft.Graph umbrella module..."
        try {
            $umbrella = Get-Module -Name 'Microsoft.Graph'
            if (-not $umbrella -or $umbrella.Version -lt [version]'2.0.0') {
                Import-Module Microsoft.Graph -MinimumVersion '2.0.0' -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Failed to import required Graph modules (v2.0.0+).`nRun: Install-Module Microsoft.Graph.Authentication -MinimumVersion '2.0.0' -Scope CurrentUser -AllowClobber"
            exit 1
        }
    }
}

# Retry wrapper for Graph API calls that may be throttled (HTTP 429).
# Reads the Retry-After response header when available; falls back to exponential back-off.
function Invoke-MgWithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3
    )
    $attempt = 0
    while ($true) {
        try {
            return (& $ScriptBlock)
        }
        catch {
            $attempt++
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            if ($statusCode -eq 429 -and $attempt -le $MaxRetries) {
                $retryAfter = $null
                try { $retryAfter = [int]$_.Exception.Response.Headers['Retry-After'] } catch {}
                $delay = if ($retryAfter -gt 0) { $retryAfter } else { [Math]::Pow(2, $attempt) * 2 }
                Write-Warning "Graph API throttled (429). Retrying in $delay s (attempt $attempt/$MaxRetries)..."
                Start-Sleep -Seconds $delay
            } else {
                throw
            }
        }
    }
}

# HTTP status code of a failed Invoke-MgGraphRequest call (0 when not available)
function Get-HttpStatusCode {
    param($ErrorRecord)
    try { return [int]$ErrorRecord.Exception.Response.StatusCode } catch { return 0 }
}

# GET a Graph collection and follow @odata.nextLink until every page is read
function Get-GraphCollection {
    param([string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-MgWithRetry { Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop }
        foreach ($item in @($page.value)) { if ($null -ne $item) { $items.Add($item) } }
        $next = $page.'@odata.nextLink'
    }
    return $items
}

# Send GET requests through the Graph JSON batching endpoint, 20 per call. Individual requests
# that are throttled (429) or fail transiently (5xx) are retried in later rounds, honouring the
# largest Retry-After seen. Returns a hashtable: request id -> batch response (status, headers, body).
function Invoke-GraphBatch {
    param(
        [System.Collections.Generic.List[hashtable]]$Requests,
        [string]$Label,
        [int]$MaxRounds = 5
    )
    $results = @{}
    $requestsById = @{}
    foreach ($r in $Requests) { $requestsById[$r.id] = $r }

    $pending = @($Requests)
    $round = 0
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($pending.Count -gt 0) {
        $round++
        $retry = [System.Collections.Generic.List[hashtable]]::new()
        $wait = 0
        $batchCount = [Math]::Ceiling($pending.Count / 20)
        # Unconditional heartbeat (~20 lines per round) so long unattended runs stay visible in logs
        $heartbeatInterval = [Math]::Max(1, [Math]::Ceiling($batchCount / 20))
        $batchNumber = 0

        for ($i = 0; $i -lt $pending.Count; $i += 20) {
            $batchNumber++
            $chunk = @($pending[$i..([Math]::Min($i + 19, $pending.Count - 1))])
            $body = @{ requests = @($chunk | ForEach-Object { @{ id = $_.id; method = 'GET'; url = $_.url } }) } | ConvertTo-Json -Depth 4 -Compress
            $response = Invoke-MgWithRetry { Invoke-MgGraphRequest -Method POST -Uri 'v1.0/$batch' -Body $body -ContentType 'application/json' -ErrorAction Stop }

            foreach ($resp in @($response.responses)) {
                $status = [int]$resp.status
                if (($status -eq 429 -or $status -ge 500) -and $round -lt $MaxRounds) {
                    $retry.Add($requestsById[[string]$resp.id])
                    $retryAfter = 0
                    try { $retryAfter = [int]$resp.headers['Retry-After'] } catch {}
                    if ($retryAfter -gt $wait) { $wait = $retryAfter }
                    continue
                }
                $results[[string]$resp.id] = $resp
            }

            Write-Progress -Activity "Reading $Label" -Status "Batch $batchNumber of $batchCount" -PercentComplete (($batchNumber / $batchCount) * 100)
            if ($batchNumber % $heartbeatInterval -eq 0 -or $batchNumber -eq $batchCount) {
                Write-Host "  ...$Label batch $batchNumber/$batchCount ($([Math]::Round($stopwatch.Elapsed.TotalSeconds))s elapsed)" -ForegroundColor DarkGray
            }
        }
        Write-Progress -Activity "Reading $Label" -Completed

        if ($retry.Count -gt 0) {
            $delay = if ($wait -gt 0) { $wait } else { [Math]::Pow(2, $round) * 2 }
            Write-Warning "$($retry.Count) $Label request(s) throttled or failed transiently. Retrying in $delay s (round $round/$MaxRounds)..."
            Start-Sleep -Seconds $delay
        }
        $pending = @($retry)
    }
    return $results
}

# Read every item of a collection returned in a batch response, following @odata.nextLink.
# Ok is $false when the request failed (missing response or HTTP status >= 400).
function Get-BatchItems {
    param($Response)
    $items = [System.Collections.Generic.List[object]]::new()
    if (-not $Response -or [int]$Response.status -ge 400) {
        return @{ Ok = $false; Items = $items }
    }
    foreach ($item in @($Response.body.value)) { if ($null -ne $item) { $items.Add($item) } }
    $next = $Response.body.'@odata.nextLink'
    if ($next) {
        foreach ($item in @(Get-GraphCollection $next)) { $items.Add($item) }
    }
    return @{ Ok = $true; Items = $items }
}

# Run a user query with signInActivity first; if the tenant cannot return sign-in data
# (no Entra ID P1/P2 license, or AuditLog.Read.All missing) fall back to the base properties
# once and remember that for the rest of the run.
function Invoke-UserQuery {
    param([scriptblock]$Query)
    if ($script:signInDataAvailable) {
        try {
            return (& $Query "$userSelectBase,signInActivity")
        }
        catch {
            $status = Get-HttpStatusCode $_
            if ($status -ne 403 -and $_.Exception.Message -notmatch 'NonPremium|Premium|Forbidden') { throw }
            Write-Warning "Sign-in activity is not available: $($_.Exception.Message)"
            Write-Warning "Reading sign-in activity requires an Entra ID P1/P2 license and AuditLog.Read.All. Continuing without sign-in data."
            $script:signInDataAvailable = $false
        }
    }
    return (& $Query $userSelectBase)
}

# Escape special HTML characters to prevent injection / markup breakage
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if (-not $Text) { return '' }
    $Text.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&#39;')
}

# Encode a string as a JavaScript string literal for the embedded reportDetails object.
# Faster than ConvertTo-Json per value (thousands of guests x several modals), and also
# escapes every < so a value can never close or confuse the surrounding <script> block.
function ConvertTo-JsString {
    param([string]$Text)
    if (-not $Text) { return '""' }
    $t = $Text.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t').Replace('<', '\u003c')
    $t = $t.Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')
    $t = [regex]::Replace($t, '[\x00-\x1F]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
    return '"' + $t + '"'
}

# Graph dates arrive as ISO-8601 strings or, depending on PowerShell / SDK version, as already
# parsed DateTime values. Normalize both to UTC.
function ConvertTo-UtcDateTime {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Local) { return $Value.ToUniversalTime() }
        return [datetime]::SpecifyKind($Value, [System.DateTimeKind]::Utc)
    }
    if ($Value -is [System.DateTimeOffset]) { return $Value.UtcDateTime }
    $text = [string]$Value
    if (-not $text) { return $null }
    try {
        return [datetime]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]'AdjustToUniversal, AssumeUniversal')
    }
    catch {
        Write-Verbose "Could not parse date value '$text'"
        return $null
    }
}

# Whole days elapsed since a UTC date ($null when the date is unknown)
function Get-DaysSince {
    param($Date, [datetime]$Now)
    if ($null -eq $Date) { return $null }
    return [int][Math]::Max(0, [Math]::Floor(($Now - $Date).TotalDays))
}

function Format-ReportDate {
    param($Date)
    if ($null -eq $Date) { return '' }
    return $Date.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-ReportDateTime {
    param($Date)
    if ($null -eq $Date) { return '' }
    return $Date.ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture) + ' UTC'
}

# Home domain of a guest: mail first, then other mail addresses, then the domain encoded in
# the #EXT# user principal name (john_contoso.com#EXT#@tenant.onmicrosoft.com -> contoso.com).
function Get-GuestDomain {
    param([string]$Mail, $OtherMails, [string]$UserPrincipalName)
    foreach ($address in @($Mail) + @($OtherMails)) {
        if ($address -and $address -match '@([^@]+)$') { return $Matches[1].ToLower() }
    }
    if ($UserPrincipalName -match '^(.+)#EXT#@') {
        $local = $Matches[1]
        $index = $local.LastIndexOf('_')
        if ($index -ge 0) { return $local.Substring($index + 1).ToLower() }
    }
    if ($UserPrincipalName -match '@([^@]+)$') { return $Matches[1].ToLower() }
    return ''
}

# Identity provider the guest redeemed the invitation with, read from the identities collection
function Get-IdentityProvider {
    param($Identities, [string]$Invitation)
    $issuers = @(@($Identities) | Where-Object { $_ -and $_.signInType -ne 'userPrincipalName' } | ForEach-Object { [string]$_.issuer } | Where-Object { $_ })
    if ($issuers -contains 'ExternalAzureAD')  { return @{ Key = 'entra';  Label = 'Entra ID' } }
    if ($issuers -contains 'MicrosoftAccount') { return @{ Key = 'msa';    Label = 'Microsoft Account' } }
    if ($issuers -contains 'mail')             { return @{ Key = 'otp';    Label = 'Email OTP' } }
    if ($issuers -contains 'google.com')       { return @{ Key = 'google'; Label = 'Google' } }
    if ($issuers.Count -gt 0)                  { return @{ Key = 'other';  Label = $issuers[0] } }
    if ($Invitation -eq 'pending')             { return @{ Key = 'unredeemed'; Label = 'Not Redeemed' } }
    return @{ Key = 'unknown'; Label = 'Unknown' }
}

function Get-GroupTypeLabel {
    param($Group)
    $types = @($Group.groupTypes)
    $label = if ($types -contains 'Unified') { 'Microsoft 365' }
             elseif ($Group.securityEnabled -and $Group.mailEnabled) { 'Mail-enabled security' }
             elseif ($Group.securityEnabled) { 'Security' }
             elseif ($Group.mailEnabled) { 'Distribution' }
             else { 'Group' }
    if ($types -contains 'DynamicMembership') { $label += ' (dynamic)' }
    return $label
}

function Get-RoleScopeLabel {
    param([string]$DirectoryScopeId)
    if (-not $DirectoryScopeId -or $DirectoryScopeId -eq '/') { return 'Tenant' }
    if ($DirectoryScopeId -like '/administrativeUnits/*') { return "Administrative unit ($($DirectoryScopeId.Substring(21)))" }
    return "Scoped ($DirectoryScopeId)"
}

# Build the Microsoft Learn deep link for an Entra built-in role.
# The reference page uses GitHub-style anchors: role name lowercased, non-alphanumerics
# dropped, and spaces collapsed to single hyphens (e.g. 'Global Administrator' -> '#global-administrator').
function Get-EntraRoleLearnUrl {
    param([string]$RoleName)
    $anchor = (($RoleName.ToLower() -replace '[^a-z0-9 -]', '') -replace '\s+', '-')
    return "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference#$anchor"
}

function New-Reason {
    param([string]$Level, [string]$Text)
    [PSCustomObject]@{ Level = $Level; Text = $Text }
}

# Cleanup recommendation for one guest: Remove / Review / Keep (or On-Prem for synced accounts),
# with the list of reasons that led to it. Thresholds come from the script parameters.
function Get-CleanupRecommendation {
    param($Guest)
    $reasons = [System.Collections.Generic.List[object]]::new()
    $rec = 'keep'

    if ($Guest.Invitation -eq 'pending') {
        if ($Guest.PendingDays -ge $PendingInvitationDays) {
            $rec = 'remove'
            $reasons.Add((New-Reason 'remove' "Invitation not redeemed for $($Guest.PendingDays) days (threshold: $PendingInvitationDays days)"))
        } else {
            $reasons.Add((New-Reason 'keep' "Invitation sent $($Guest.PendingDays) days ago, still within the $PendingInvitationDays-day redemption window"))
        }
    }
    elseif ($null -ne $Guest.AgeDays -and $Guest.AgeDays -lt $NewGuestDays) {
        $reasons.Add((New-Reason 'keep' "Created $($Guest.AgeDays) days ago, within the $NewGuestDays-day grace period for new guests"))
    }
    elseif (-not $script:signInDataAvailable) {
        $rec = 'review'
        $reasons.Add((New-Reason 'review' "Sign-in activity is unavailable in this tenant, so inactivity cannot be assessed"))
    }
    elseif ($null -eq $Guest.LastActivity) {
        $since = $Guest.DaysSinceRedemption
        $sinceText = if ($Guest.Invitation -eq 'accepted') { "the invitation was redeemed $since days ago" } else { "the account was created $since days ago" }
        if ($null -eq $since -or $since -ge $InactiveDays) {
            $rec = 'remove'
            $reasons.Add((New-Reason 'remove' "Never signed in, although $sinceText (threshold: $InactiveDays days)"))
        } else {
            $rec = 'review'
            $reasons.Add((New-Reason 'review' "Never signed in since $sinceText"))
        }
    }
    elseif ($Guest.DaysSinceActivity -ge $RemoveAfterDays) {
        $rec = 'remove'
        $reasons.Add((New-Reason 'remove' "No sign-in for $($Guest.DaysSinceActivity) days (remove threshold: $RemoveAfterDays days)"))
    }
    elseif ($Guest.DaysSinceActivity -ge $InactiveDays) {
        $rec = 'review'
        $reasons.Add((New-Reason 'review' "No sign-in for $($Guest.DaysSinceActivity) days (inactive threshold: $InactiveDays days)"))
    }
    else {
        $reasons.Add((New-Reason 'keep' "Last sign-in $($Guest.DaysSinceActivity) days ago (inactive threshold: $InactiveDays days)"))
    }

    if ($Guest.LastActivityIsApproximate) {
        $reasons.Add((New-Reason 'info' "No successful sign-in is recorded. Last activity is the most recent sign-in attempt, which may have failed"))
    }

    # A disabled account that is also inactive has no reason to exist; a disabled account that
    # was recently active needs a human decision.
    if (-not $Guest.AccountEnabled) {
        if ($rec -eq 'keep') {
            $rec = 'review'
            $reasons.Add((New-Reason 'review' "Account is disabled although it was recently active or created. Confirm whether it should be deleted"))
        } elseif ($rec -eq 'review' -and $Guest.ActivityBucket -in @('inactive', 'never')) {
            $rec = 'remove'
            $reasons.Add((New-Reason 'remove' "Account is disabled and inactive"))
        } else {
            $reasons.Add((New-Reason 'info' "Account is disabled (sign-in blocked)"))
        }
    }

    # Privileged external accounts always deserve a second look, even when active
    if ($Guest.Roles.Count -gt 0) {
        $roleNames = (@($Guest.Roles | ForEach-Object { $_.Role }) | Sort-Object -Unique) -join ', '
        $reasons.Add((New-Reason 'review' "Holds $($Guest.Roles.Count) directory role assignment(s): $roleNames. Confirm this external account still needs privileged access"))
        if ($rec -eq 'keep') { $rec = 'review' }
    }

    if ($Guest.AccessCollected) {
        if (($Guest.Groups.Count + $Guest.Teams.Count + $Guest.Apps.Count + $Guest.Roles.Count) -eq 0) {
            $reasons.Add((New-Reason 'info' "No direct group, Teams, application or directory role access found"))
        }
        if ($Guest.Sponsors.Count -eq 0) {
            $reasons.Add((New-Reason 'info' "No sponsor assigned. Nobody in the organization is recorded as responsible for this guest"))
        }
    }

    if ($Guest.OnPremisesSynced) {
        $rec = 'onprem'
        $reasons.Add((New-Reason 'info' "Synchronized from on-premises Active Directory. Changes made in Entra ID are overwritten by the next sync"))
    }

    return @{ Recommendation = $rec; Reasons = $reasons }
}

function New-ModalTableHtml {
    param([string]$TableId, [string[]]$Headers, [string]$RowsHtml)
    $headerCells = for ($i = 0; $i -lt $Headers.Count; $i++) {
        "<th onclick=`"sortTable('$TableId',$i,'string')`" style='text-align:left;padding:4px 8px;border-bottom:1px solid var(--border)'>$($Headers[$i])</th>"
    }
    "<table id='$TableId' style='width:100%;border-collapse:collapse;font-size:13px'><thead><tr>$($headerCells -join '')</tr></thead><tbody>$RowsHtml</tbody></table>"
}

function New-ModalRowHtml {
    param([string[]]$Cells)
    '<tr>' + (($Cells | ForEach-Object { "<td style='padding:4px 8px;border-bottom:1px solid var(--border-row)'>$_</td>" }) -join '') + '</tr>'
}

function Get-ModalOnClick {
    param([string]$Key, [string]$Name)
    "openDetailModal(reportDetails['$Key'].${Name}Title,reportDetails['$Key'].${Name}Html)"
}

# Import required modules
Write-Host "Importing Microsoft Graph modules..." -ForegroundColor Green
Import-GraphModuleSafely "Microsoft.Graph.Authentication"

# Connect to Microsoft Graph
$scopes = @(
    "User.Read.All",
    "AuditLog.Read.All",
    "Directory.Read.All",
    "RoleManagement.Read.Directory"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
try {
    if ($AccessToken) {
        # Token passed in - e.g. from AzurePowerShell@5 azureSubscription task
        Write-Host "  Auth method: Pre-acquired access token" -ForegroundColor Cyan
        Connect-MgGraph -AccessToken $AccessToken -NoWelcome -ErrorAction Stop
    } elseif ($ClientId -and $CertificateThumbprint) {
        # Service Principal with certificate - recommended for pipelines
        Write-Host "  Auth method: Service Principal (certificate)" -ForegroundColor Cyan
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
    } elseif ($UseManagedIdentity) {
        # Managed Identity - for Azure-hosted automation.
        # -ClientId selects a user-assigned Managed Identity; omit it for the system-assigned one.
        if ($ClientId) {
            Write-Host "  Auth method: Managed Identity (user-assigned)" -ForegroundColor Cyan
            Connect-MgGraph -Identity -ClientId $ClientId -NoWelcome -ErrorAction Stop
        } else {
            Write-Host "  Auth method: Managed Identity" -ForegroundColor Cyan
            Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
        }
    } elseif ($TenantId) {
        Connect-MgGraph -Scopes $scopes -TenantId $TenantId -NoWelcome -ErrorAction Stop
    } else {
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    }
    if (-not (Get-MgContext)) {
        throw 'Connect-MgGraph returned no context - authentication may have been cancelled or denied.'
    }
}
catch {
    Write-Error "Authentication failed: $($_.Exception.Message)"
    exit 1
}

try {

Write-Host "Gathering guest users..." -ForegroundColor Green

if ($TargetUser) {
    # Targeted single-user scan - accepts the object ID or the user principal name
    Write-Host "Targeted scan requested for '$TargetUser' - skipping tenant enumeration" -ForegroundColor Cyan
    try {
        $rawUsers = @(Invoke-UserQuery {
            param([string]$Select)
            Invoke-MgWithRetry { Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$([uri]::EscapeDataString($TargetUser))?`$select=$Select" -ErrorAction Stop }
        })
    }
    catch {
        if ((Get-HttpStatusCode $_) -eq 404) {
            Write-Error "No user found matching -TargetUser '$TargetUser' (checked object ID and user principal name)."
            exit 1
        }
        throw
    }
    $target = $rawUsers[0]
    if ($target.userType -ne 'Guest' -and $target.userPrincipalName -notlike '*#EXT#*' -and -not $target.externalUserState) {
        Write-Error "'$TargetUser' is an internal member account, not a guest or external member."
        exit 1
    }
} else {
    $rawUsers = @(Invoke-UserQuery {
        param([string]$Select)
        # signInActivity limits the page size, so $top is only raised when it is not requested
        $top = if ($Select -like '*signInActivity*') { '' } else { '&$top=999' }
        Get-GraphCollection "v1.0/users?`$filter=$([uri]::EscapeDataString("userType eq 'Guest'"))&`$select=$Select$top"
    })
    Write-Host "Found $($rawUsers.Count) guest users" -ForegroundColor Green

    # External members: B2B users with userType Member (invited as members or created by
    # cross-tenant synchronization). They can't be filtered server-side without advanced queries,
    # so all members are listed with a minimal projection and matched client-side, then the
    # full properties are read for the matches through $batch.
    if (-not $ExcludeExternalMembers) {
        Write-Host "Looking for external members (B2B users with userType Member)..." -ForegroundColor Green
        $memberCandidates = Get-GraphCollection "v1.0/users?`$filter=$([uri]::EscapeDataString("userType eq 'Member'"))&`$select=id,userPrincipalName,externalUserState&`$top=999"
        $externalMemberIds = @($memberCandidates | Where-Object { $_.userPrincipalName -like '*#EXT#*' -or $_.externalUserState } | ForEach-Object { $_.id })
        if ($externalMemberIds.Count -gt 0) {
            $select = if ($script:signInDataAvailable) { "$userSelectBase,signInActivity" } else { $userSelectBase }
            $memberRequests = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($id in $externalMemberIds) {
                $memberRequests.Add(@{ id = "u_$id"; url = "/users/$($id)?`$select=$select" })
            }
            $memberResults = Invoke-GraphBatch -Requests $memberRequests -Label 'external members'
            foreach ($result in $memberResults.Values) {
                if ([int]$result.status -eq 200) { $rawUsers += $result.body }
                else { Write-Warning "Could not read external member '$($result.id.Substring(2))' (HTTP $($result.status))." }
            }
        }
        Write-Host "Found $($externalMemberIds.Count) external members" -ForegroundColor Green
    }
}

# Confirmation prompt (skipped in non-interactive / pipeline mode, or for a targeted single-user scan)
if (-not $NonInteractive -and -not $TargetUser -and -not $SkipAccessDetails -and $rawUsers.Count -gt 0) {
    Write-Host "`n" -NoNewline
    $confirmation = Read-Host "Continue with access analysis of $($rawUsers.Count) users? (Y/N)"
    if ($confirmation -notmatch '^[Yy]') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
} else {
    Write-Host "Processing $($rawUsers.Count) users..." -ForegroundColor Cyan
}

# BULK PRE-FETCH: directory role assignments and definitions are read once tenant-wide and
# matched to guests (directly, or through role-assignable groups) instead of per user.
$roleDefinitionsById = @{}
$roleAssignmentsByPrincipal = @{}
try {
    Write-Host "Pre-fetching directory role assignments (bulk)..." -ForegroundColor Green
    foreach ($definition in (Get-GraphCollection 'v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName')) {
        $roleDefinitionsById[$definition.id] = $definition.displayName
    }
    foreach ($assignment in (Get-GraphCollection 'v1.0/roleManagement/directory/roleAssignments')) {
        if (-not $roleAssignmentsByPrincipal.ContainsKey($assignment.principalId)) {
            $roleAssignmentsByPrincipal[$assignment.principalId] = [System.Collections.Generic.List[object]]::new()
        }
        $roleAssignmentsByPrincipal[$assignment.principalId].Add($assignment)
    }
    Write-Host "  Cached role assignments for $($roleAssignmentsByPrincipal.Count) principals" -ForegroundColor Cyan
}
catch {
    Write-Warning "Could not read directory role assignments ($($_.Exception.Message)). Directory roles will not be reported."
}

# Per-guest access details through $batch: direct group memberships (incl. Teams), app role
# assignments and sponsors - three requests per guest, twenty requests per Graph call.
$accessResults = @{}
$servicePrincipalsById = @{}
if (-not $SkipAccessDetails -and $rawUsers.Count -gt 0) {
    Write-Host "Reading group, Teams, application and sponsor details..." -ForegroundColor Green
    $accessRequests = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($u in $rawUsers) {
        $accessRequests.Add(@{ id = "m_$($u.id)"; url = "/users/$($u.id)/memberOf?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,isAssignableToRole,resourceProvisioningOptions,visibility&`$top=999" })
        $accessRequests.Add(@{ id = "a_$($u.id)"; url = "/users/$($u.id)/appRoleAssignments?`$select=id,appRoleId,resourceId,resourceDisplayName,createdDateTime" })
        $accessRequests.Add(@{ id = "s_$($u.id)"; url = "/users/$($u.id)/sponsors?`$select=id,displayName,userPrincipalName,mail" })
    }
    $accessResults = Invoke-GraphBatch -Requests $accessRequests -Label 'access details'

    # Resolve app role names: read appRoles once per resource service principal
    $resourceIds = @($accessResults.Keys | Where-Object { $_ -like 'a_*' } | ForEach-Object {
        @($accessResults[$_].body.value) | Where-Object { $_ } | ForEach-Object { $_.resourceId }
    } | Sort-Object -Unique)
    if ($resourceIds.Count -gt 0) {
        $spRequests = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($resourceId in $resourceIds) {
            $spRequests.Add(@{ id = "sp_$resourceId"; url = "/servicePrincipals/$($resourceId)?`$select=id,appId,appRoles" })
        }
        $spResults = Invoke-GraphBatch -Requests $spRequests -Label 'application roles'
        foreach ($result in $spResults.Values) {
            if ([int]$result.status -ne 200) { continue }
            $roleNames = @{}
            foreach ($appRole in @($result.body.appRoles)) { if ($appRole) { $roleNames[[string]$appRole.id] = $appRole.displayName } }
            $servicePrincipalsById[$result.body.id] = @{ AppId = $result.body.appId; Roles = $roleNames }
        }
    }
}

Write-Host "`nProceeding with analysis..." -ForegroundColor Green

$now = [DateTime]::UtcNow
$report = [System.Collections.Generic.List[object]]::new()

foreach ($u in $rawUsers) {
    $invitation = switch ($u.externalUserState) {
        'Accepted'          { 'accepted' }
        'PendingAcceptance' { 'pending' }
        default             { 'none' }
    }
    $created      = ConvertTo-UtcDateTime $u.createdDateTime
    $stateChanged = ConvertTo-UtcDateTime $u.externalUserStateChangeDateTime
    $ageDays      = Get-DaysSince $created $now
    $redeemedOrCreated = if ($invitation -eq 'accepted' -and $stateChanged) { $stateChanged } else { $created }
    $pendingDays  = if ($invitation -eq 'pending') { Get-DaysSince $(if ($stateChanged) { $stateChanged } else { $created }) $now } else { $null }

    # Sign-in activity: lastSuccessfulSignInDateTime is authoritative. It is only recorded since
    # December 2023, so older activity falls back to the last interactive / non-interactive
    # sign-in attempt, which may include failed attempts (flagged as approximate).
    $signIn             = $u.signInActivity
    $lastSuccess        = ConvertTo-UtcDateTime $signIn.lastSuccessfulSignInDateTime
    $lastInteractive    = ConvertTo-UtcDateTime $signIn.lastSignInDateTime
    $lastNonInteractive = ConvertTo-UtcDateTime $signIn.lastNonInteractiveSignInDateTime
    $lastActivity = $lastSuccess
    $isApproximate = $false
    if (-not $lastActivity) {
        $lastActivity = @($lastInteractive, $lastNonInteractive) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
        $isApproximate = [bool]$lastActivity
    }
    $daysSinceActivity = Get-DaysSince $lastActivity $now

    $activityBucket = if (-not $script:signInDataAvailable) { 'unknown' }
                      elseif ($invitation -eq 'pending') { 'pending' }
                      elseif (-not $lastActivity) { 'never' }
                      elseif ($daysSinceActivity -lt $InactiveDays) { 'active' }
                      elseif ($daysSinceActivity -lt $RemoveAfterDays) { 'inactive' }
                      else { 'stale' }

    $createdBucket = if ($null -eq $ageDays) { '' }
                     elseif ($ageDays -le $NewGuestDays) { 'new' }
                     elseif ($ageDays -le 365) { 'lastyear' }
                     else { 'over1y' }

    $groups   = [System.Collections.Generic.List[object]]::new()
    $teams    = [System.Collections.Generic.List[object]]::new()
    $apps     = [System.Collections.Generic.List[object]]::new()
    $sponsors = [System.Collections.Generic.List[object]]::new()
    $roles    = [System.Collections.Generic.List[object]]::new()
    $accessCollected = -not $SkipAccessDetails

    if ($accessCollected) {
        $membership = Get-BatchItems $accessResults["m_$($u.id)"]
        $appAssignments = Get-BatchItems $accessResults["a_$($u.id)"]
        $sponsorItems = Get-BatchItems $accessResults["s_$($u.id)"]
        if (-not ($membership.Ok -and $appAssignments.Ok -and $sponsorItems.Ok)) {
            $script:accessReadFailures++
            Write-Verbose "Some access details could not be read for $($u.userPrincipalName)"
        }

        foreach ($item in $membership.Items) {
            if ($item['@odata.type'] -ne '#microsoft.graph.group') { continue }
            $entry = [PSCustomObject]@{
                Id             = [string]$item.id
                Name           = [string]$item.displayName
                Type           = Get-GroupTypeLabel $item
                RoleAssignable = [bool]$item.isAssignableToRole
                Visibility     = [string]$item.visibility
            }
            if (@($item.resourceProvisioningOptions) -contains 'Team') { $teams.Add($entry) } else { $groups.Add($entry) }
        }

        foreach ($item in $appAssignments.Items) {
            $resource = $servicePrincipalsById[[string]$item.resourceId]
            $roleName = if ([string]$item.appRoleId -eq '00000000-0000-0000-0000-000000000000') { 'Default Access' }
                        elseif ($resource -and $resource.Roles.ContainsKey([string]$item.appRoleId)) { $resource.Roles[[string]$item.appRoleId] }
                        else { [string]$item.appRoleId }
            $apps.Add([PSCustomObject]@{
                ResourceId = [string]$item.resourceId
                AppId      = if ($resource) { $resource.AppId } else { '' }
                Name       = [string]$item.resourceDisplayName
                Role       = $roleName
                Assigned   = ConvertTo-UtcDateTime $item.createdDateTime
            })
        }

        foreach ($item in $sponsorItems.Items) {
            $sponsors.Add([PSCustomObject]@{
                Name = [string]$item.displayName
                Upn  = if ($item.userPrincipalName) { [string]$item.userPrincipalName } else { [string]$item.mail }
                Type = if ($item['@odata.type'] -eq '#microsoft.graph.group') { 'Group' } else { 'User' }
            })
        }
    }

    # Directory roles: assigned to the guest directly, or to a role-assignable group the guest is in
    foreach ($assignment in @($roleAssignmentsByPrincipal[[string]$u.id])) {
        if (-not $assignment) { continue }
        $roles.Add([PSCustomObject]@{
            Role  = if ($roleDefinitionsById.ContainsKey($assignment.roleDefinitionId)) { $roleDefinitionsById[$assignment.roleDefinitionId] } else { $assignment.roleDefinitionId }
            Via   = 'Direct'
            Scope = Get-RoleScopeLabel $assignment.directoryScopeId
        })
    }
    foreach ($group in @($groups) + @($teams)) {
        if (-not $group -or -not $group.RoleAssignable) { continue }
        foreach ($assignment in @($roleAssignmentsByPrincipal[$group.Id])) {
            if (-not $assignment) { continue }
            $roles.Add([PSCustomObject]@{
                Role  = if ($roleDefinitionsById.ContainsKey($assignment.roleDefinitionId)) { $roleDefinitionsById[$assignment.roleDefinitionId] } else { $assignment.roleDefinitionId }
                Via   = "Group: $($group.Name)"
                Scope = Get-RoleScopeLabel $assignment.directoryScopeId
            })
        }
    }

    $identityProvider = Get-IdentityProvider $u.identities $invitation

    $guest = [PSCustomObject]@{
        Id                        = [string]$u.id
        DisplayName               = [string]$u.displayName
        UserPrincipalName         = [string]$u.userPrincipalName
        Mail                      = [string]$u.mail
        Domain                    = Get-GuestDomain $u.mail $u.otherMails $u.userPrincipalName
        CompanyName               = [string]$u.companyName
        CreationType              = [string]$u.creationType
        UserType                  = if ($u.userType -eq 'Guest') { 'guest' } else { 'member' }
        Created                   = $created
        AgeDays                   = $ageDays
        CreatedBucket             = $createdBucket
        Invitation                = $invitation
        InvitationStateChanged    = $stateChanged
        PendingDays               = $pendingDays
        DaysSinceRedemption       = Get-DaysSince $redeemedOrCreated $now
        AccountEnabled            = [bool]$u.accountEnabled
        OnPremisesSynced          = [bool]$u.onPremisesSyncEnabled
        IdentityKey               = $identityProvider.Key
        IdentityLabel             = $identityProvider.Label
        LastSuccessfulSignIn      = $lastSuccess
        LastInteractiveSignIn     = $lastInteractive
        LastNonInteractiveSignIn  = $lastNonInteractive
        LastActivity              = $lastActivity
        LastActivityIsApproximate = $isApproximate
        DaysSinceActivity         = $daysSinceActivity
        ActivityBucket            = $activityBucket
        AccessCollected           = $accessCollected
        Groups                    = $groups
        Teams                     = $teams
        Apps                      = $apps
        Roles                     = $roles
        Sponsors                  = $sponsors
        Recommendation            = $null
        Reasons                   = $null
    }
    $assessment = Get-CleanupRecommendation $guest
    $guest.Recommendation = $assessment.Recommendation
    $guest.Reasons = $assessment.Reasons
    $report.Add($guest)
}

if ($script:accessReadFailures -gt 0) {
    Write-Warning "Access details could not be read completely for $($script:accessReadFailures) user(s). Their group, app or sponsor counts may be incomplete (run with -Verbose for names)."
}

Write-Host "Final report contains $($report.Count) users" -ForegroundColor Green

# Calculate summary statistics
# Note: wrap pipelines in @(...) so .Count is reliable in Windows PowerShell 5.1
# (a single matched object otherwise yields $null for .Count).
$totalGuests        = @($report).Count
$removeCount        = @($report | Where-Object { $_.Recommendation -eq 'remove' }).Count
$reviewCount        = @($report | Where-Object { $_.Recommendation -eq 'review' }).Count
$keepCount          = @($report | Where-Object { $_.Recommendation -eq 'keep' }).Count
$onPremCount        = @($report | Where-Object { $_.Recommendation -eq 'onprem' }).Count
$pendingCount       = @($report | Where-Object { $_.Invitation -eq 'pending' }).Count
$neverSignedInCount = @($report | Where-Object { $_.ActivityBucket -eq 'never' }).Count
$inactiveCount      = @($report | Where-Object { $_.ActivityBucket -in @('inactive', 'stale') }).Count
$staleCount         = @($report | Where-Object { $_.ActivityBucket -eq 'stale' }).Count
$disabledCount      = @($report | Where-Object { -not $_.AccountEnabled }).Count
$privilegedCount    = @($report | Where-Object { $_.Roles.Count -gt 0 }).Count
$noSponsorCount     = @($report | Where-Object { $_.AccessCollected -and $_.Sponsors.Count -eq 0 }).Count
$externalMemberCount = @($report | Where-Object { $_.UserType -eq 'member' }).Count
$guestTypeCount     = $totalGuests - $externalMemberCount

# Get tenant information
$tenantInfo = @((Invoke-MgWithRetry { Invoke-MgGraphRequest -Method GET -Uri 'v1.0/organization?$select=id,displayName' -ErrorAction Stop }).value)[0]
$tenantName = [string]$tenantInfo.displayName
$tenantId = [string]$tenantInfo.id

if (-not $OutputPath) {
    $safeName = $tenantName -replace '[^\w]', '_'
    $OutputPath = "EntraIDGuestReport__${safeName}_$(Get-Date -Format 'yyyy-MM-dd').html"
}

# Resolve a relative output path against the current PowerShell location.
# .NET file APIs (WriteAllText / Path.GetFullPath) resolve relative paths against
# [Environment]::CurrentDirectory (the process start directory, e.g. C:\Windows\System32
# for elevated shells), NOT PowerShell's $PWD. Making the path absolute here ensures the
# report is written where the user expects - next to their current location.
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path (Get-Location).Path $OutputPath
}

# Ensure output directory exists
$outputDir = Split-Path $OutputPath -Parent
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Cyan
}

# Verify the output path is writable before spending time on HTML generation
$_writeTestDir  = if ($outputDir) { $outputDir } else { '.' }
$_writeTestFile = Join-Path $_writeTestDir ([System.IO.Path]::GetRandomFileName())
try {
    [System.IO.File]::WriteAllText($_writeTestFile, '')
    Remove-Item $_writeTestFile -Force -ErrorAction SilentlyContinue
} catch {
    Write-Error "Output path '$OutputPath' is not writable: $($_.Exception.Message)"
    exit 1
}

# Generate HTML report. A StringBuilder keeps rendering linear for tenants with many guests.
$sb = [System.Text.StringBuilder]::new()

[void]$sb.Append(@'
<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Entra ID Guest User Report</title>
    <style>
        /* -- Light Theme Variables -- */
        :root {
            --blue:        #0078d4;
            --blue-dark:   #005a9e;
            --blue-light:  #cce4f6;
            --blue-xs:     #f0f6fc;
            --text:        #1a1a1a;
            --text-muted:  #888;
            --gray:        #666;
            --gray-dark:   #444;
            --bg-page:     #edf2f7;
            --bg-panel:    #ffffff;
            --bg-header:   #0078d4;
            --bg-input:    #ffffff;
            --bg-subtle:   #f7fafc;
            --border:      #ddd;
            --border-row:  #e2e8f0;
            --shadow:      0 2px 10px rgba(0,0,0,0.08);
            --radius:      8px;

            --rec-remove-bg:   #fed7d7;
            --rec-remove-text: #c53030;
            --rec-review-bg:   #fefcbf;
            --rec-review-text: #b7791f;
            --rec-keep-bg:     #c6f6d5;
            --rec-keep-text:   #2f855a;

            --good:          #2f855a;
            --bad:           #c53030;
            --warn:          #b7791f;
        }

        /* -- Dark Theme Variables -- */
        [data-theme="dark"] {
            --blue:        #3aa0f3;
            --blue-dark:   #2b8fd4;
            --blue-light:  #1a3a55;
            --blue-xs:     #152030;
            --text:        #e2e8f0;
            --text-muted:  #6a7585;
            --gray:        #9aa5b3;
            --gray-dark:   #c8d0da;
            --bg-page:     #111520;
            --bg-panel:    #1c1f2e;
            --bg-header:   #0d1117;
            --bg-input:    #252840;
            --bg-subtle:   #232738;
            --border:      #343848;
            --border-row:  #2a2e3e;
            --shadow:      0 2px 12px rgba(0,0,0,0.35);

            --rec-remove-bg:   #3a1518;
            --rec-remove-text: #f98a8a;
            --rec-review-bg:   #332b0d;
            --rec-review-text: #e6c45c;
            --rec-keep-bg:     #122a1a;
            --rec-keep-text:   #6ee7a0;

            --good:          #68d391;
            --bad:           #f98a8a;
            --warn:          #e6c45c;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            font-size: 14px;
            color: var(--text);
            background: var(--bg-page);
            min-height: 100vh;
            transition: background 0.25s, color 0.25s;
        }

        a { color: var(--blue); }

        /* -- Header -- */
        header {
            background: var(--bg-header);
            color: #fff;
            padding: 16px 32px;
            display: flex;
            align-items: center;
            gap: 16px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.3);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        header .header-text { flex: 1; }
        header h1 { font-size: 16px; font-weight: 600; letter-spacing: 0.2px; }
        header .meta { font-size: 11px; opacity: 0.75; margin-top: 1px; line-height: 1.7; }
        header .meta strong { font-weight: 600; }

        #themeToggle {
            background: rgba(255,255,255,0.12);
            border: 1px solid rgba(255,255,255,0.25);
            color: #fff;
            font-size: 17px;
            width: 36px; height: 36px;
            border-radius: 50%;
            cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            transition: background 0.2s, transform 0.2s;
            flex-shrink: 0;
            line-height: 1;
        }
        #themeToggle:hover { background: rgba(255,255,255,0.25); transform: rotate(20deg); }
        .github-link {
            background: rgba(255,255,255,0.12);
            border: 1px solid rgba(255,255,255,0.25);
            width: 36px; height: 36px;
            border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            flex-shrink: 0;
            transition: background 0.2s, transform 0.2s;
            line-height: 0;
        }
        .github-link:hover { background: rgba(255,255,255,0.25); transform: scale(1.1); }

        /* -- Layout -- */
        .container { max-width: 1600px; margin: 24px auto; padding: 0 24px 48px; }

        /* -- Summary cards -- */
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .summary-card {
            background: var(--bg-panel);
            padding: 18px 20px;
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            border-top: 4px solid var(--blue);
            text-align: center;
            cursor: pointer;
            transition: background 0.25s, box-shadow 0.25s, transform 0.15s;
        }
        .summary-card:hover { transform: translateY(-2px); box-shadow: 0 4px 14px rgba(0,0,0,0.12); }
        .summary-card h3 { margin: 0 0 6px; color: var(--gray-dark); font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
        .summary-card .number { font-size: 30px; font-weight: 700; color: var(--blue); margin: 6px 0; }
        /* Number and border colors matching badge colors per card type */
        .summary-card[data-fv="remove"] .number         { color: #c50f1f; }
        .summary-card[data-fv="review"] .number         { color: #b06000; }
        .summary-card[data-fv="keep"] .number           { color: #107c10; }
        .summary-card[data-fv="pending"] .number        { color: #b06000; }
        .summary-card[data-fv="never"] .number          { color: #d83b01; }
        .summary-card[data-fv="inactive,stale"] .number { color: #d83b01; }
        .summary-card[data-fv="no"][data-fg="enabled"] .number { color: #767676; }
        .summary-card[data-fv="roles"] .number          { color: #7054a0; }
        .summary-card[data-fv="nosponsor"] .number      { color: #b06000; }
        .summary-card[data-fv="remove"]                 { border-top-color: #c50f1f; }
        .summary-card[data-fv="review"]                 { border-top-color: #b06000; }
        .summary-card[data-fv="keep"]                   { border-top-color: #107c10; }
        .summary-card[data-fv="pending"]                { border-top-color: #b06000; }
        .summary-card[data-fv="never"]                  { border-top-color: #d83b01; }
        .summary-card[data-fv="inactive,stale"]         { border-top-color: #d83b01; }
        .summary-card[data-fv="no"][data-fg="enabled"]  { border-top-color: #767676; }
        .summary-card[data-fv="roles"]                  { border-top-color: #7054a0; }
        .summary-card[data-fv="nosponsor"]              { border-top-color: #b06000; }
        [data-theme="dark"] .summary-card[data-fv="remove"] .number         { color: #e3223a; }
        [data-theme="dark"] .summary-card[data-fv="review"] .number         { color: #c87000; }
        [data-theme="dark"] .summary-card[data-fv="keep"] .number           { color: #2a9d2a; }
        [data-theme="dark"] .summary-card[data-fv="pending"] .number        { color: #c87000; }
        [data-theme="dark"] .summary-card[data-fv="never"] .number          { color: #f0571f; }
        [data-theme="dark"] .summary-card[data-fv="inactive,stale"] .number { color: #f0571f; }
        [data-theme="dark"] .summary-card[data-fv="roles"] .number          { color: #9070c0; }
        [data-theme="dark"] .summary-card[data-fv="nosponsor"] .number      { color: #c87000; }
        [data-theme="dark"] .summary-card[data-fv="remove"]                 { border-top-color: #e3223a; }
        [data-theme="dark"] .summary-card[data-fv="review"]                 { border-top-color: #c87000; }
        [data-theme="dark"] .summary-card[data-fv="keep"]                   { border-top-color: #2a9d2a; }
        [data-theme="dark"] .summary-card[data-fv="pending"]                { border-top-color: #c87000; }
        [data-theme="dark"] .summary-card[data-fv="never"]                  { border-top-color: #f0571f; }
        [data-theme="dark"] .summary-card[data-fv="inactive,stale"]         { border-top-color: #f0571f; }
        [data-theme="dark"] .summary-card[data-fv="roles"]                  { border-top-color: #9070c0; }
        [data-theme="dark"] .summary-card[data-fv="nosponsor"]              { border-top-color: #c87000; }
        .summary-card .subtitle { color: var(--text-muted); font-size: 12.5px; }

        /* -- Controls -- */
        .controls {
            background: var(--bg-panel);
            padding: 20px;
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            margin-bottom: 24px;
            transition: background 0.25s, box-shadow 0.25s;
        }

        /* -- Tag filters -- */
        .filter-bar { display: flex; align-items: center; gap: 12px; margin: 4px 0 8px; }
        .filter-toggle-btn {
            display: inline-flex; align-items: center; gap: 8px;
            background: var(--bg-panel); border: 1px solid var(--border); color: var(--text);
            padding: 8px 16px; border-radius: 6px;
            font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
            cursor: pointer;
            transition: background 0.15s, border-color 0.15s;
        }
        .filter-toggle-btn:hover { border-color: var(--blue); background: var(--blue-xs); }
        .filter-toggle-btn[aria-expanded="true"] { border-color: var(--blue); }
        .filter-toggle-btn .caret { font-size: 10px; color: var(--gray); transition: transform 0.15s; }
        .export-csv-btn {
            display: inline-flex; align-items: center; gap: 8px;
            background: var(--bg-panel); border: 1px solid var(--border); color: var(--text);
            padding: 8px 16px; border-radius: 6px;
            font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
            cursor: pointer; margin-left: auto;
            transition: background 0.15s, border-color 0.15s;
        }
        .export-csv-btn:hover { border-color: var(--blue); background: var(--blue-xs); }
        .filter-summary { font-size: 13px; font-weight: 600; color: var(--text); }

        .clickable-badge { cursor: pointer; transition: filter 0.15s, box-shadow 0.15s; }

        .filter-search {
            width: 100%;
            margin: 4px 0 28px;
            padding: 11px 14px;
            border: 1px solid var(--border);
            border-radius: 6px;
            font-size: 14px;
            color: var(--text);
            background: var(--bg-input);
            transition: border-color 0.15s, box-shadow 0.15s;
        }
        .filter-search:focus { outline: none; border-color: var(--blue); box-shadow: 0 0 0 2px var(--blue-light); }

        .filter-group { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; flex-wrap: wrap; }
        .filter-group-label {
            font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
            color: var(--gray); width: 130px; flex-shrink: 0;
        }
        .filter-tag {
            display: inline-flex; align-items: center; gap: 5px;
            padding: 4px 12px; border: 1px solid var(--border); border-radius: 14px;
            background: var(--bg-panel); color: var(--text); font-size: 12px; font-weight: 600;
            cursor: pointer; user-select: none;
            transition: background 0.15s, border-color 0.15s, color 0.15s;
        }
        .filter-tag:hover { border-color: var(--blue); background: var(--blue-xs); }
        .filter-tag .cnt { font-size: 10px; font-weight: 600; color: var(--text-muted); }
        .filter-tag.active { color: #fff; border-color: transparent; }
        .filter-tag.active .cnt { color: rgba(255,255,255,0.85); }

        .filter-tag.active.c-blue   { background: #0078d4; }
        .filter-tag.active.c-amber  { background: #b06000; }
        .filter-tag.active.c-red    { background: #c50f1f; }
        .filter-tag.active.c-orange { background: #d83b01; }
        .filter-tag.active.c-green  { background: #107c10; }
        .filter-tag.active.c-purple { background: #7054a0; }
        .filter-tag.active.c-gray   { background: #767676; }
        [data-theme="dark"] .filter-tag.active.c-green  { background: #2a9d2a; }
        [data-theme="dark"] .filter-tag.active.c-red    { background: #e3223a; }
        [data-theme="dark"] .filter-tag.active.c-orange { background: #f0571f; }
        [data-theme="dark"] .filter-tag.active.c-amber  { background: #c87000; }
        [data-theme="dark"] .filter-tag.active.c-purple { background: #9070c0; }

        .filter-results {
            display: flex; align-items: center; flex-wrap: wrap; gap: 8px;
            margin-top: 16px; padding-top: 14px; border-top: 1px solid var(--border);
        }
        .result-count { font-size: 13px; font-weight: 600; color: var(--gray-dark); margin-right: 4px; }
        .active-chips { display: inline-flex; flex-wrap: wrap; gap: 6px; }
        .chip {
            display: inline-flex; align-items: center; gap: 4px;
            background: #0078d4; color: #fff; border: none;
            padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; cursor: pointer;
        }
        .chip.c-green  { background: #107c10; }
        .chip.c-red    { background: #c50f1f; }
        .chip.c-orange { background: #d83b01; }
        .chip.c-amber  { background: #b06000; }
        .chip.c-purple { background: #7054a0; }
        .chip.c-gray   { background: #767676; }
        .chip:hover { filter: brightness(1.15); }
        [data-theme="dark"] .chip.c-green  { background: #2a9d2a; }
        [data-theme="dark"] .chip.c-red    { background: #e3223a; }
        [data-theme="dark"] .chip.c-orange { background: #f0571f; }
        [data-theme="dark"] .chip.c-amber  { background: #c87000; }
        [data-theme="dark"] .chip.c-purple { background: #9070c0; }
        .clear-all {
            display: inline-flex; align-items: center; gap: 8px;
            background: var(--bg-panel); border: 1px solid var(--border); color: var(--text);
            padding: 8px 16px; border-radius: 6px;
            font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
            cursor: pointer; margin-left: auto;
            transition: background 0.15s, border-color 0.15s;
        }
        .clear-all:hover { border-color: var(--blue); background: var(--blue-xs); }

        /* -- Table -- */
        .table-wrap {
            background: var(--bg-panel);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            overflow: auto;
            transition: background 0.25s, box-shadow 0.25s;
        }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 11px 12px; text-align: left; border-bottom: 1px solid var(--border-row); vertical-align: top; }
        td { font-size: 11px; font-weight: 400; color: var(--gray-dark); letter-spacing: 0.5px; }
        th {
            background: var(--bg-subtle); font-weight: 700; cursor: pointer; user-select: none;
            position: sticky; top: 0; z-index: 1; color: var(--gray-dark);
            text-transform: uppercase; letter-spacing: 0.5px; font-size: 11px;
        }
        th:hover { background: var(--blue-xs); }
        th.sort-asc::after  { content: ' \25B2'; font-size: 0.8em; opacity: 0.7; }
        th.sort-desc::after { content: ' \25BC'; font-size: 0.8em; opacity: 0.7; }
        tbody tr { transition: background 0.12s; }
        tbody tr:hover { background: var(--bg-subtle); }
        .csv-only { display: none; }

        /* -- Recommendation row tints -- */
        .rec-remove { background: var(--rec-remove-bg) !important; color: var(--rec-remove-text); }
        .rec-review { background: var(--rec-review-bg) !important; color: var(--rec-review-text); }
        .rec-keep   { background: var(--rec-keep-bg) !important; color: var(--rec-keep-text); }

        /* -- Badges -- */
        .badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            background: #767676;
            color: #fff;
            font-size: 9px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 10px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            vertical-align: middle;
            text-align: center;
            line-height: 1.3;
            white-space: nowrap;
        }
        .badge.green   { background: #107c10; }
        .badge.red     { background: #c50f1f; }
        .badge.orange  { background: #d83b01; }
        .badge.amber   { background: #b06000; }
        .badge.blue    { background: #0078d4; }
        .badge.purple  { background: #7054a0; }
        .badge.gray    { background: #767676; }

        [data-theme="dark"] .badge.green  { background: #2a9d2a; }
        [data-theme="dark"] .badge.red    { background: #e3223a; }
        [data-theme="dark"] .badge.orange { background: #f0571f; }
        [data-theme="dark"] .badge.amber  { background: #c87000; }
        [data-theme="dark"] .badge.purple { background: #9070c0; }

        .perm-badges { display: flex; flex-direction: column; align-items: flex-start; gap: 4px; }

        .activity-inactive { color: #d83b01; font-weight: 700; }
        .activity-stale    { color: #c50f1f; font-weight: 700; }
        [data-theme="dark"] .activity-inactive { color: #f0571f; }
        [data-theme="dark"] .activity-stale    { color: #e3223a; }
        .activity-link { cursor: pointer; }
        .activity-link:hover { text-decoration: underline; }

        /* -- Recommendation modal -- */
        .rec-header {
            display: flex; align-items: center; justify-content: space-between; gap: 16px;
            margin: 0 0 14px 0; padding-bottom: 12px; border-bottom: 1px solid var(--border);
        }
        .rec-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); }
        .rec-action { font-size: 13px; margin-top: 4px; color: var(--text); }
        .rec-header .badge { font-size: 12px; padding: 5px 14px; }
        .reason-list { margin: 5px 0; padding-left: 4px; list-style: none; }
        .reason-list li { display: flex; gap: 10px; align-items: baseline; margin: 6px 0; font-size: 13px; }
        .reason-list li .badge { flex: 0 0 64px; }
        .info-banner {
            background: rgba(176, 96, 0, 0.08); border: 1px solid #b06000; border-left: 3px solid #b06000;
            border-radius: 6px; padding: 8px 12px; margin: 0 0 14px 0; font-size: 12.5px; color: var(--text-muted);
        }
        [data-theme="dark"] .info-banner {
            background: rgba(200, 112, 0, 0.14); border-color: #c87000;
        }
        .page-banner { margin-bottom: 24px; font-size: 13px; }
        .modal-note { margin-top: 12px; font-size: 12px; color: var(--text-muted); }

        .mono { font-family: 'Inter', sans-serif; font-size: 11px; font-weight: 400; color: var(--gray-dark); letter-spacing: 0.5px; }
        .muted { color: var(--text-muted); }
        .guest-name { font-family: 'Inter', sans-serif; font-size: 11px; font-weight: 600; color: var(--gray-dark); letter-spacing: 0.5px; text-decoration: none; }
        .guest-name:hover { text-decoration: underline; color: var(--blue); }
        .cell-sm { font-size: 11px; }

        /* -- Modal -- */
        .modal-overlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,0.55); z-index: 1000;
            align-items: center; justify-content: center;
        }
        .modal-overlay.active { display: flex; }
        .modal-box {
            position: relative;
            background: var(--bg-panel); color: var(--text);
            border: 1px solid var(--border); border-radius: 10px;
            max-width: 720px; width: 90%; max-height: 80vh;
            overflow-y: auto; padding: 24px 28px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.35);
        }
        .modal-box h3 { color: var(--text); margin: 0 0 14px 0; padding-right: 30px; font-size: 15px; }
        .modal-close {
            position: absolute; top: 12px; right: 14px;
            background: none; border: none; color: var(--text-muted);
            font-size: 22px; cursor: pointer; line-height: 1; padding: 0;
        }
        .modal-close:hover { color: var(--text); }
        @media (max-width: 600px) { .modal-box { width: 95%; max-height: 90vh; padding: 18px; } }

        /* -- Footer -- */
        footer.report-footer {
            text-align: center;
            background: var(--bg-header);
            color: rgba(255,255,255,0.85);
            font-size: 11px;
            padding: 16px 24px;
            margin-top: 0;
        }
        footer.report-footer a { color: rgba(255,255,255,0.85); text-decoration: none; }
        footer.report-footer a:hover { text-decoration: underline; }
    </style>
</head>
'@)

$signInBanner = if (-not $script:signInDataAvailable) {
    "<div class='info-banner page-banner'><strong>Sign-in activity unavailable.</strong> Reading sign-in activity requires an Entra ID P1/P2 license and the AuditLog.Read.All permission. Recommendations below are based on invitation state, account age and account status only, and guests that would otherwise be assessed on inactivity are marked for Review.</div>"
} else { '' }
$accessBanner = if ($SkipAccessDetails) {
    "<div class='info-banner page-banner'><strong>Access details skipped.</strong> The report was generated with -SkipAccessDetails: group, Teams, application and sponsor details were not collected. Only directly assigned directory roles are shown.</div>"
} else { '' }
$unknownActivityTag = if (-not $script:signInDataAvailable) {
    "<span class='filter-tag c-gray' data-group='activity' data-value='unknown' onclick='toggleTag(this)'>Unknown <span class='cnt'></span></span>"
} else { '' }

[void]$sb.Append(@"
<body>
    <header>
        <div class="header-text">
            <h1>Entra ID Guest User Report</h1>
            <div class="meta">
                <strong>Tenant:</strong> $(ConvertTo-HtmlSafe $tenantName) &nbsp;&middot;&nbsp;
                <strong>Tenant ID:</strong> $(ConvertTo-HtmlSafe $tenantId) &nbsp;&middot;&nbsp;
                <strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") &nbsp;&middot;&nbsp;
                <strong>Thresholds:</strong> inactive $InactiveDays d, remove $RemoveAfterDays d, pending invitation $PendingInvitationDays d
            </div>
        </div>
        <a class="github-link" href="https://github.com/matejklemencic/EntraID-GuestReport" target="_blank" rel="noopener" title="View source on GitHub"><svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 16 16" fill="#ffffff"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0 0 16 8c0-4.42-3.58-8-8-8z"/></svg></a>
        <button id="themeToggle" onclick="toggleTheme()" title="Toggle dark / light mode">&#127769;</button>
    </header>

    <div class="container">

    $signInBanner
    $accessBanner

    <div class="summary">
        <div class="summary-card" data-clear="1" title="Clear all filters">
            <h3>Total Guests</h3>
            <div class="number">$totalGuests</div>
            <div class="subtitle">$guestTypeCount guests &middot; $externalMemberCount external members</div>
        </div>
        <div class="summary-card" data-fg="rec" data-fv="remove" title="Strong cleanup candidates. Click to filter.">
            <h3>Remove</h3>
            <div class="number">$removeCount</div>
            <div class="subtitle">Recommended for deletion</div>
        </div>
        <div class="summary-card" data-fg="rec" data-fv="review" title="Needs a decision by the sponsor or resource owner. Click to filter.">
            <h3>Review</h3>
            <div class="number">$reviewCount</div>
            <div class="subtitle">Confirm whether access is still needed</div>
        </div>
        <div class="summary-card" data-fg="rec" data-fv="keep" title="Active or recently invited. Click to filter.">
            <h3>Keep</h3>
            <div class="number">$keepCount</div>
            <div class="subtitle">No action needed</div>
        </div>
        <div class="summary-card" data-fg="invitation" data-fv="pending" title="Invitation sent but never redeemed. Click to filter.">
            <h3>Pending Invitations</h3>
            <div class="number">$pendingCount</div>
            <div class="subtitle">Invitation not redeemed yet</div>
        </div>
        <div class="summary-card" data-fg="activity" data-fv="never" title="Redeemed or created, but never signed in. Click to filter.">
            <h3>Never Signed In</h3>
            <div class="number">$neverSignedInCount</div>
            <div class="subtitle">No sign-in recorded</div>
        </div>
        <div class="summary-card" data-fg="activity" data-fv="inactive,stale" title="No sign-in for $InactiveDays days or more. Click to filter.">
            <h3>Inactive</h3>
            <div class="number">$inactiveCount</div>
            <div class="subtitle">No sign-in for $InactiveDays+ days ($staleCount for $RemoveAfterDays+)</div>
        </div>
        <div class="summary-card" data-fg="enabled" data-fv="no" title="Sign-in blocked. Click to filter.">
            <h3>Disabled</h3>
            <div class="number">$disabledCount</div>
            <div class="subtitle">Guests with sign-in disabled</div>
        </div>
        <div class="summary-card" data-fg="access" data-fv="roles" title="External accounts holding Entra ID directory roles. Click to filter.">
            <h3>Directory Roles</h3>
            <div class="number">$privilegedCount</div>
            <div class="subtitle">Guests with privileged access</div>
        </div>
        <div class="summary-card" data-fg="sponsor" data-fv="nosponsor" title="Nobody in the organization is recorded as responsible. Click to filter.">
            <h3>No Sponsor</h3>
            <div class="number">$noSponsorCount</div>
            <div class="subtitle">Guests without an accountable sponsor</div>
        </div>
    </div>

    <div class="filter-bar">
        <button id="filterToggle" class="filter-toggle-btn" onclick="toggleFilterPanel()" aria-expanded="false">
            <span class="caret">&#9656;</span> Search &amp; Filter
        </button>
        <span id="filterSummary" class="filter-summary"></span>
        <button class="export-csv-btn" onclick="exportCsv()" title="Download visible rows as CSV">&#8595; Export CSV</button>
    </div>

    <div class="controls" id="controlsPanel" hidden>
        <input type="text" id="searchInput" class="filter-search" placeholder="Search by name, email, domain, company, sponsor, group, team, app or role..." oninput="queueApplyFilters()">

        <div class="filter-group">
            <span class="filter-group-label">Recommendation</span>
            <span class="filter-tag c-red"   data-group="rec" data-value="remove" onclick="toggleTag(this)">Remove <span class="cnt"></span></span>
            <span class="filter-tag c-amber" data-group="rec" data-value="review" onclick="toggleTag(this)">Review <span class="cnt"></span></span>
            <span class="filter-tag c-green" data-group="rec" data-value="keep"   onclick="toggleTag(this)">Keep <span class="cnt"></span></span>
            <span class="filter-tag c-gray"  data-group="rec" data-value="onprem" onclick="toggleTag(this)">On-Prem <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Last Sign-In</span>
            <span class="filter-tag c-green"  data-group="activity" data-value="active"   onclick="toggleTag(this)">Active (&lt; $InactiveDays d) <span class="cnt"></span></span>
            <span class="filter-tag c-orange" data-group="activity" data-value="inactive" onclick="toggleTag(this)">Inactive ($InactiveDays+ d) <span class="cnt"></span></span>
            <span class="filter-tag c-red"    data-group="activity" data-value="stale"    onclick="toggleTag(this)">Stale ($RemoveAfterDays+ d) <span class="cnt"></span></span>
            <span class="filter-tag c-orange" data-group="activity" data-value="never"    onclick="toggleTag(this)">Never <span class="cnt"></span></span>
            $unknownActivityTag
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Invitation</span>
            <span class="filter-tag c-green" data-group="invitation" data-value="accepted" onclick="toggleTag(this)">Accepted <span class="cnt"></span></span>
            <span class="filter-tag c-amber" data-group="invitation" data-value="pending"  onclick="toggleTag(this)">Pending <span class="cnt"></span></span>
            <span class="filter-tag c-gray"  data-group="invitation" data-value="none"     onclick="toggleTag(this)">No Invitation State <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Enabled</span>
            <span class="filter-tag c-green" data-group="enabled" data-value="yes" onclick="toggleTag(this)">Enabled <span class="cnt"></span></span>
            <span class="filter-tag c-gray"  data-group="enabled" data-value="no"  onclick="toggleTag(this)">Disabled <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Created</span>
            <span class="filter-tag c-orange" data-group="created" data-value="new"      onclick="toggleTag(this)">New Guest <span class="cnt"></span></span>
            <span class="filter-tag c-blue"   data-group="created" data-value="lastyear" onclick="toggleTag(this)">Last Year <span class="cnt"></span></span>
            <span class="filter-tag c-gray"   data-group="created" data-value="over1y"   onclick="toggleTag(this)">+1 Year <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">User Type</span>
            <span class="filter-tag c-blue"   data-group="usertype" data-value="guest"  onclick="toggleTag(this)">Guest <span class="cnt"></span></span>
            <span class="filter-tag c-purple" data-group="usertype" data-value="member" onclick="toggleTag(this)">External Member <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Identity</span>
            <span class="filter-tag c-blue"  data-group="identity" data-value="entra"      onclick="toggleTag(this)">Entra ID <span class="cnt"></span></span>
            <span class="filter-tag c-blue"  data-group="identity" data-value="msa"        onclick="toggleTag(this)">Microsoft Account <span class="cnt"></span></span>
            <span class="filter-tag c-blue"  data-group="identity" data-value="otp"        onclick="toggleTag(this)">Email OTP <span class="cnt"></span></span>
            <span class="filter-tag c-blue"  data-group="identity" data-value="google"     onclick="toggleTag(this)">Google <span class="cnt"></span></span>
            <span class="filter-tag c-gray"  data-group="identity" data-value="other"      onclick="toggleTag(this)">Other <span class="cnt"></span></span>
            <span class="filter-tag c-amber" data-group="identity" data-value="unredeemed" onclick="toggleTag(this)">Not Redeemed <span class="cnt"></span></span>
            <span class="filter-tag c-gray"  data-group="identity" data-value="unknown"    onclick="toggleTag(this)">Unknown <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Access</span>
            <span class="filter-tag c-blue"   data-group="access" data-value="groups"   onclick="toggleTag(this)">Groups <span class="cnt"></span></span>
            <span class="filter-tag c-blue"   data-group="access" data-value="teams"    onclick="toggleTag(this)">Teams <span class="cnt"></span></span>
            <span class="filter-tag c-blue"   data-group="access" data-value="apps"     onclick="toggleTag(this)">Apps <span class="cnt"></span></span>
            <span class="filter-tag c-purple" data-group="access" data-value="roles"    onclick="toggleTag(this)">Directory Roles <span class="cnt"></span></span>
            <span class="filter-tag c-gray"   data-group="access" data-value="noaccess" onclick="toggleTag(this)">No Access <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Sponsor</span>
            <span class="filter-tag c-green" data-group="sponsor" data-value="has"       onclick="toggleTag(this)">Has Sponsor <span class="cnt"></span></span>
            <span class="filter-tag c-amber" data-group="sponsor" data-value="nosponsor" onclick="toggleTag(this)">No Sponsor <span class="cnt"></span></span>
        </div>
        <div class="filter-group">
            <span class="filter-group-label">Source</span>
            <span class="filter-tag c-blue" data-group="source" data-value="cloud"  onclick="toggleTag(this)">Cloud <span class="cnt"></span></span>
            <span class="filter-tag c-gray" data-group="source" data-value="onprem" onclick="toggleTag(this)">On-Premises Synced <span class="cnt"></span></span>
        </div>

        <div class="filter-results">
            <span id="resultCount" class="result-count"></span>
            <span id="activeChips" class="active-chips"></span>
            <button id="clearAll" class="clear-all" onclick="clearAllFilters()" style="display:none;">Clear all</button>
        </div>
    </div>

    <div class="table-wrap">
    <table id="reportTable">
        <thead>
            <tr>
                <th onclick="sortTable('reportTable', 0, 'string')">Guest</th>
                <th onclick="sortTable('reportTable', 1, 'string')">Domain</th>
                <th onclick="sortTable('reportTable', 2, 'string')">User Type</th>
                <th onclick="sortTable('reportTable', 3, 'string')">Identity</th>
                <th onclick="sortTable('reportTable', 4, 'string')">Created</th>
                <th onclick="sortTable('reportTable', 5, 'string')">Invitation</th>
                <th onclick="sortTable('reportTable', 6, 'string')">Enabled</th>
                <th onclick="sortTable('reportTable', 7, 'string')">Last Sign-In</th>
                <th>Access</th>
                <th onclick="sortTable('reportTable', 9, 'string')">Sponsors</th>
                <th onclick="sortTable('reportTable', 10, 'string')">Recommendation</th>
                <th class="csv-only">Object ID</th>
                <th class="csv-only">User Principal Name</th>
                <th class="csv-only">Mail</th>
                <th class="csv-only">Company</th>
                <th class="csv-only">Days Since Last Sign-In</th>
                <th class="csv-only">Groups</th>
                <th class="csv-only">Teams</th>
                <th class="csv-only">Applications</th>
                <th class="csv-only">Directory Roles</th>
                <th class="csv-only">Sponsor Names</th>
                <th class="csv-only">Reasons</th>
            </tr>
        </thead>
        <tbody>
"@)

# Sort: Remove first, then Review, Keep, On-Prem; within a group the longest inactive first
$recOrder = @{ remove = 1; review = 2; keep = 3; onprem = 4 }
$sortedReport = @($report | Sort-Object @{ Expression = { $recOrder[$_.Recommendation] } },
                                        @{ Expression = { if ($null -eq $_.DaysSinceActivity) { [int]::MaxValue } else { $_.DaysSinceActivity } }; Descending = $true },
                                        @{ Expression = { $_.DisplayName } })
$modalDataEntries = [System.Collections.Generic.List[string]]::new()

$recBadgeColor = @{ remove = 'red'; review = 'amber'; keep = 'green'; onprem = 'gray' }
$recLabel      = @{ remove = 'Remove'; review = 'Review'; keep = 'Keep'; onprem = 'On-Prem' }
$recAction     = @{
    remove = 'Delete this guest account. Deleted users can be restored from Deleted users within 30 days.'
    review = 'Confirm with the sponsor or the owners of the resources this guest can access whether access is still needed. Consider an Access Review for recurring checks.'
    keep   = 'No action needed at this time.'
    onprem = 'This account is synchronized from on-premises Active Directory. Disable or delete it in the source directory.'
}
$reasonBadge = @{ remove = 'red'; review = 'amber'; keep = 'green'; info = 'gray' }

foreach ($g in $sortedReport) {
    # Object ID from Graph - escaped for consistency with every other API-sourced field.
    # ConvertTo-HtmlSafe is a no-op on a well-formed GUID, so the key and portal URL are unaffected.
    $key = ConvertTo-HtmlSafe $g.Id
    $nameText = if ($g.DisplayName) { $g.DisplayName } else { $g.UserPrincipalName }
    $safeName = ConvertTo-HtmlSafe $nameText
    $contact = if ($g.Mail) { $g.Mail } else { $g.UserPrincipalName }
    $portalUrl = "https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/$key"
    $newBadge = if ($g.CreatedBucket -eq 'new') { "<span class='badge orange clickable-badge csv-exclude' data-fg='created' data-fv='new' title='Created within the last $NewGuestDays days.'>New</span> " } else { '' }

    # User type / source
    $userTypeText = if ($g.UserType -eq 'guest') {
        "<span class='badge blue clickable-badge' data-fg='usertype' data-fv='guest' title='B2B guest (userType Guest).'>Guest</span>"
    } else {
        "<span class='badge purple clickable-badge' data-fg='usertype' data-fv='member' title='External B2B user with userType Member, e.g. invited as a member or created by cross-tenant synchronization. Members get the same default directory permissions as employees.'>External Member</span>"
    }
    if ($g.OnPremisesSynced) {
        $userTypeText += "<div style='margin-top:4px'><span class='badge gray clickable-badge' data-fg='source' data-fv='onprem' title='Synchronized from on-premises Active Directory.'>On-Prem Synced</span></div>"
    }

    $identityText = "<span class='badge gray clickable-badge' data-fg='identity' data-fv='$($g.IdentityKey)' title='Identity provider used to redeem the invitation.'>$(ConvertTo-HtmlSafe $g.IdentityLabel)</span>"

    $invitationText = switch ($g.Invitation) {
        'accepted' { "<span class='badge green clickable-badge' data-fg='invitation' data-fv='accepted' title='Invitation redeemed.'>Accepted</span>" }
        'pending'  { "<span class='badge amber clickable-badge' data-fg='invitation' data-fv='pending' title='Invitation sent but not redeemed yet.'>Pending</span><div class='muted' style='margin-top:4px'>$($g.PendingDays) days</div>" }
        default    { "<span class='badge gray clickable-badge' data-fg='invitation' data-fv='none' title='No B2B invitation state, e.g. created by cross-tenant synchronization or synced from on-premises.'>N/A</span>" }
    }

    $enabledText = if ($g.AccountEnabled) { "<span class='badge green clickable-badge' data-fg='enabled' data-fv='yes' title='Sign-in allowed.'>Yes</span>" } else { "<span class='badge gray clickable-badge' data-fg='enabled' data-fv='no' title='Sign-in blocked.'>No</span>" }

    # Last sign-in cell (opens the activity modal)
    $activityOnClick = Get-ModalOnClick $key 'activity'
    $activitySort = if ($g.LastActivity) { Format-ReportDate $g.LastActivity } elseif ($g.ActivityBucket -eq 'never') { '0000-00-00' } else { '' }
    $activityInner = switch ($g.ActivityBucket) {
        'unknown' { "<span class='muted'>Unknown</span>" }
        'pending' { "<span class='muted'>Not redeemed</span>" }
        'never'   { "<span class='activity-inactive'>Never</span>" }
        default {
            $cls = if ($g.ActivityBucket -eq 'stale') { 'activity-stale' } elseif ($g.ActivityBucket -eq 'inactive') { 'activity-inactive' } else { '' }
            $approxMark = if ($g.LastActivityIsApproximate) { '*' } else { '' }
            "<span class='$cls'>$(Format-ReportDate $g.LastActivity)$approxMark</span><div class='muted' style='margin-top:4px'>$($g.DaysSinceActivity) days ago</div>"
        }
    }
    $activityTitle = if ($g.LastActivityIsApproximate) { 'Based on the last sign-in attempt (no successful sign-in recorded). Click for sign-in details.' } else { 'Click for sign-in details.' }
    $activityText = "<span class='activity-link' onclick=`"$activityOnClick`" title='$activityTitle'>$activityInner</span>"

    # Access badges (each opens its own modal)
    $accessBadges = [System.Collections.Generic.List[string]]::new()
    if ($g.AccessCollected) {
        if ($g.Groups.Count -gt 0) { $accessBadges.Add("<span class='badge blue' onclick=`"$(Get-ModalOnClick $key 'groups')`" style='cursor:pointer' title='Direct group memberships. Click to view details.'>Groups: $($g.Groups.Count)</span>") }
        if ($g.Teams.Count -gt 0)  { $accessBadges.Add("<span class='badge blue' onclick=`"$(Get-ModalOnClick $key 'teams')`" style='cursor:pointer' title='Microsoft Teams the guest is a member of. Click to view details.'>Teams: $($g.Teams.Count)</span>") }
        if ($g.Apps.Count -gt 0)   { $accessBadges.Add("<span class='badge blue' onclick=`"$(Get-ModalOnClick $key 'apps')`" style='cursor:pointer' title='Applications assigned directly to the guest. Click to view details.'>Apps: $($g.Apps.Count)</span>") }
    }
    if ($g.Roles.Count -gt 0) { $accessBadges.Add("<span class='badge purple' onclick=`"$(Get-ModalOnClick $key 'roles')`" style='cursor:pointer' title='Entra ID directory roles, assigned directly or through a role-assignable group. Click to view details.'>Roles: $($g.Roles.Count)</span>") }
    if (-not $g.AccessCollected) {
        $accessBadges.Add("<span class='badge gray' title='Not collected (-SkipAccessDetails).'>Not collected</span>")
    } elseif ($accessBadges.Count -eq 0) {
        $accessBadges.Add("<span class='badge gray' title='No direct group, Teams, application or directory role access found.'>None</span>")
    }
    $accessText = $accessBadges -join ' '

    $sponsorText = if (-not $g.AccessCollected) {
        "<span class='badge gray' title='Not collected (-SkipAccessDetails).'>Not collected</span>"
    } elseif ($g.Sponsors.Count -gt 0) {
        "<span class='badge blue' onclick=`"$(Get-ModalOnClick $key 'sponsors')`" style='cursor:pointer' title='People or groups responsible for this guest. Click to view details.'>$($g.Sponsors.Count) sponsor(s)</span>"
    } else {
        "<span class='badge gray clickable-badge' data-fg='sponsor' data-fv='nosponsor' title='Nobody in the organization is recorded as responsible for this guest.'>No sponsor</span>"
    }
    $sponsorValue = if (-not $g.AccessCollected) { 'na' } elseif ($g.Sponsors.Count -gt 0) { 'has' } else { 'none' }

    $recText = "<span class='badge $($recBadgeColor[$g.Recommendation])' onclick=`"$(Get-ModalOnClick $key 'rec')`" style='cursor:pointer' title='Click to see why.'>$($recLabel[$g.Recommendation])</span>"

    # ----- Modal content -----
    $recItems = ($g.Reasons | ForEach-Object {
        "<li><span class='badge $($reasonBadge[$_.Level])'>$($_.Level)</span><span>$(ConvertTo-HtmlSafe $_.Text)</span></li>"
    }) -join ''
    $recHtml = "<div class='rec-header'><div><div class='rec-label'>Cleanup recommendation</div><div class='rec-action'>$($recAction[$g.Recommendation])</div></div><span class='badge $($recBadgeColor[$g.Recommendation])'>$($recLabel[$g.Recommendation])</span></div>" +
               "<ul class='reason-list'>$recItems</ul>" +
               "<p class='modal-note'>Thresholds: inactive after $InactiveDays days, remove after $RemoveAfterDays days without sign-in, unredeemed invitations removed after $PendingInvitationDays days, new guests (under $NewGuestDays days) are kept.</p>"

    $none = "<span class='muted'>None recorded</span>"
    $activityRows = @(
        @('Last successful sign-in',              $(if ($g.LastSuccessfulSignIn) { Format-ReportDateTime $g.LastSuccessfulSignIn } else { $none })),
        @('Last interactive sign-in attempt',     $(if ($g.LastInteractiveSignIn) { Format-ReportDateTime $g.LastInteractiveSignIn } else { $none })),
        @('Last non-interactive sign-in attempt', $(if ($g.LastNonInteractiveSignIn) { Format-ReportDateTime $g.LastNonInteractiveSignIn } else { $none })),
        @('Days since last activity',             $(if ($null -ne $g.DaysSinceActivity) { "$($g.DaysSinceActivity)" } else { $none })),
        @('Created',                              $(if ($g.Created) { Format-ReportDateTime $g.Created } else { $none })),
        @('Invitation state',                     $(switch ($g.Invitation) { 'accepted' { 'Accepted' } 'pending' { 'Pending acceptance' } default { 'N/A' } })),
        @('Invitation state changed',             $(if ($g.InvitationStateChanged) { Format-ReportDateTime $g.InvitationStateChanged } else { $none })),
        @('Identity provider',                    (ConvertTo-HtmlSafe $g.IdentityLabel)),
        @('Creation type',                        $(if ($g.CreationType) { ConvertTo-HtmlSafe $g.CreationType } else { $none })),
        @('Company',                              $(if ($g.CompanyName) { ConvertTo-HtmlSafe $g.CompanyName } else { $none })),
        @('User principal name',                  (ConvertTo-HtmlSafe $g.UserPrincipalName)),
        @('Mail',                                 $(if ($g.Mail) { ConvertTo-HtmlSafe $g.Mail } else { $none })),
        @('Object ID',                            "<code class='mono'>$key</code>")
    )
    $activityTable = "<table style='width:100%;border-collapse:collapse;font-size:13px'><tbody>" +
        (($activityRows | ForEach-Object { "<tr><td style='padding:4px 8px;border-bottom:1px solid var(--border-row);width:45%'><strong>$($_[0])</strong></td><td style='padding:4px 8px;border-bottom:1px solid var(--border-row)'>$($_[1])</td></tr>" }) -join '') +
        "</tbody></table>"
    $activityNote = if (-not $script:signInDataAvailable) {
        "<div class='info-banner'>Sign-in activity is unavailable in this tenant (requires Entra ID P1/P2 and AuditLog.Read.All).</div>"
    } elseif ($g.LastActivityIsApproximate) {
        "<div class='info-banner'>No successful sign-in is recorded. Entra ID only tracks successful sign-ins since December 2023, so last activity falls back to the most recent sign-in attempt, which may have failed.</div>"
    } else { '' }
    $activityHtml = $activityNote + $activityTable

    $groupsHtml = ''
    if ($g.Groups.Count -gt 0) {
        $rows = ($g.Groups | Sort-Object Name | ForEach-Object {
            $gid = ConvertTo-HtmlSafe $_.Id
            $link = "<a href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/$gid' target='_blank' rel='noopener'>$(ConvertTo-HtmlSafe $_.Name)</a>"
            $flag = if ($_.RoleAssignable) { "<span class='badge purple' title='Role-assignable group: membership can grant Entra ID directory roles.'>Role-assignable</span>" } else { '' }
            New-ModalRowHtml @($link, (ConvertTo-HtmlSafe $_.Type), $flag)
        }) -join ''
        $groupsHtml = (New-ModalTableHtml 'groupsModalTable' @('Group', 'Type', 'Notes') $rows) +
            "<p class='modal-note'>Direct memberships only. Memberships of dynamic groups are removed automatically when the guest is deleted or no longer matches the rule.</p>"
    }

    $teamsHtml = ''
    if ($g.Teams.Count -gt 0) {
        $rows = ($g.Teams | Sort-Object Name | ForEach-Object {
            $gid = ConvertTo-HtmlSafe $_.Id
            $link = "<a href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/$gid' target='_blank' rel='noopener'>$(ConvertTo-HtmlSafe $_.Name)</a>"
            $visibility = if ($_.Visibility) { ConvertTo-HtmlSafe $_.Visibility } else { '' }
            New-ModalRowHtml @($link, $visibility)
        }) -join ''
        $teamsHtml = (New-ModalTableHtml 'teamsModalTable' @('Team', 'Visibility') $rows) +
            "<p class='modal-note'>Teams are Microsoft 365 groups with a team provisioned. Guest access to individual shared channels is not included.</p>"
    }

    $appsHtml = ''
    if ($g.Apps.Count -gt 0) {
        $rows = ($g.Apps | Sort-Object Name | ForEach-Object {
            $appName = ConvertTo-HtmlSafe $_.Name
            $nameCell = if ($_.AppId) {
                "<a href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$(ConvertTo-HtmlSafe $_.ResourceId)/appId/$(ConvertTo-HtmlSafe $_.AppId)' target='_blank' rel='noopener'>$appName</a>"
            } else { $appName }
            New-ModalRowHtml @($nameCell, (ConvertTo-HtmlSafe $_.Role), (Format-ReportDate $_.Assigned))
        }) -join ''
        $appsHtml = (New-ModalTableHtml 'appsModalTable' @('Application', 'Role', 'Assigned') $rows) +
            "<p class='modal-note'>Direct assignments only. Access granted through group assignments is visible in the Groups list.</p>"
    }

    $rolesHtml = ''
    if ($g.Roles.Count -gt 0) {
        $rows = ($g.Roles | Sort-Object Role | ForEach-Object {
            $roleLink = "<a href='$(Get-EntraRoleLearnUrl $_.Role)' target='_blank' rel='noopener'>$(ConvertTo-HtmlSafe $_.Role)</a>"
            New-ModalRowHtml @($roleLink, (ConvertTo-HtmlSafe $_.Via), (ConvertTo-HtmlSafe $_.Scope))
        }) -join ''
        $rolesHtml = "<div class='info-banner'>External accounts with directory roles are a common target. Confirm the role is still required and prefer PIM eligible assignments over active ones.</div>" +
            (New-ModalTableHtml 'rolesModalTable' @('Role', 'Assignment', 'Scope') $rows) +
            "<p class='modal-note'>Active role assignments (including activated PIM assignments). PIM eligible assignments are not included.</p>"
    }

    $sponsorsHtml = ''
    if ($g.Sponsors.Count -gt 0) {
        $rows = ($g.Sponsors | Sort-Object Name | ForEach-Object {
            New-ModalRowHtml @((ConvertTo-HtmlSafe $_.Name), (ConvertTo-HtmlSafe $_.Upn), $_.Type)
        }) -join ''
        $sponsorsHtml = New-ModalTableHtml 'sponsorsModalTable' @('Name', 'UPN / Mail', 'Type') $rows
    }

    $modalDataEntries.Add("`"$key`": { recTitle: $(ConvertTo-JsString "Cleanup Recommendation for $nameText"), recHtml: $(ConvertTo-JsString $recHtml), activityTitle: $(ConvertTo-JsString "Sign-In Activity for $nameText"), activityHtml: $(ConvertTo-JsString $activityHtml), groupsTitle: $(ConvertTo-JsString "Groups for $nameText ($($g.Groups.Count))"), groupsHtml: $(ConvertTo-JsString $groupsHtml), teamsTitle: $(ConvertTo-JsString "Teams for $nameText ($($g.Teams.Count))"), teamsHtml: $(ConvertTo-JsString $teamsHtml), appsTitle: $(ConvertTo-JsString "Applications for $nameText ($($g.Apps.Count))"), appsHtml: $(ConvertTo-JsString $appsHtml), rolesTitle: $(ConvertTo-JsString "Directory Roles for $nameText ($($g.Roles.Count))"), rolesHtml: $(ConvertTo-JsString $rolesHtml), sponsorsTitle: $(ConvertTo-JsString "Sponsors for $nameText ($($g.Sponsors.Count))"), sponsorsHtml: $(ConvertTo-JsString $sponsorsHtml) }")

    # Search text and CSV-only values
    $groupNames   = @($g.Groups | ForEach-Object { $_.Name }) -join ', '
    $teamNames    = @($g.Teams | ForEach-Object { $_.Name }) -join ', '
    $appNames     = @($g.Apps | ForEach-Object { "$($_.Name) ($($_.Role))" }) -join ', '
    $roleNames    = @($g.Roles | ForEach-Object { "$($_.Role) ($($_.Via))" }) -join ', '
    $sponsorNames = @($g.Sponsors | ForEach-Object { $_.Name }) -join ', '
    $reasonText   = @($g.Reasons | ForEach-Object { $_.Text }) -join ' | '
    $searchText = ConvertTo-HtmlSafe ((@($g.DisplayName, $g.UserPrincipalName, $g.Mail, $g.Domain, $g.CompanyName, $sponsorNames, $groupNames, $teamNames, $appNames, $roleNames) | Where-Object { $_ }) -join ' ').ToLower()
    $daysSinceText = if ($null -ne $g.DaysSinceActivity) { "$($g.DaysSinceActivity)" } else { '' }

    $rowClass = if ($g.Recommendation -eq 'onprem') { '' } else { "rec-$($g.Recommendation)" }
    $collectedValue = if ($g.AccessCollected) { 'yes' } else { 'no' }
    $enabledValue = if ($g.AccountEnabled) { 'yes' } else { 'no' }
    $sourceValue = if ($g.OnPremisesSynced) { 'onprem' } else { 'cloud' }

    [void]$sb.Append(@"
            <tr class="$rowClass" data-search="$searchText" data-rec="$($g.Recommendation)" data-usertype="$($g.UserType)" data-invitation="$($g.Invitation)" data-enabled="$enabledValue" data-createdbucket="$($g.CreatedBucket)" data-activity="$($g.ActivityBucket)" data-identity="$($g.IdentityKey)" data-groupcount="$($g.Groups.Count)" data-teamcount="$($g.Teams.Count)" data-appcount="$($g.Apps.Count)" data-rolecount="$($g.Roles.Count)" data-collected="$collectedValue" data-sponsor="$sponsorValue" data-source="$sourceValue">
                <td data-sort="$($safeName.ToLower())">$newBadge<a class="guest-name" href="$portalUrl" target="_blank" title="Open in Entra portal">$safeName</a><div class="muted" style="margin-top:4px">$(ConvertTo-HtmlSafe $contact)</div></td>
                <td>$(ConvertTo-HtmlSafe $g.Domain)</td>
                <td class="cell-sm">$userTypeText</td>
                <td class="cell-sm">$identityText</td>
                <td class="cell-sm">$(Format-ReportDate $g.Created)</td>
                <td class="cell-sm">$invitationText</td>
                <td class="cell-sm">$enabledText</td>
                <td class="cell-sm" data-sort="$activitySort">$activityText</td>
                <td class="cell-sm"><span class="perm-badges">$accessText</span></td>
                <td class="cell-sm" data-sort="$($g.Sponsors.Count)">$sponsorText</td>
                <td class="cell-sm" data-sort="$($recOrder[$g.Recommendation])">$recText</td>
                <td class="csv-only">$key</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $g.UserPrincipalName)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $g.Mail)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $g.CompanyName)</td>
                <td class="csv-only">$daysSinceText</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $groupNames)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $teamNames)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $appNames)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $roleNames)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $sponsorNames)</td>
                <td class="csv-only">$(ConvertTo-HtmlSafe $reasonText)</td>
            </tr>
"@)
}

$reportDetailsBlock = "const reportDetails = {`n" + ($modalDataEntries -join ",`n") + "`n};"

[void]$sb.Append(@"
        </tbody>
    </table>
    </div>

    <script>
        $reportDetailsBlock
"@)

[void]$sb.Append(@'

        // Active filters: group -> Set of selected values (OR within a group, AND across groups)
        const activeFilters = {};
        const groupLabels = {
            rec: 'Recommendation', activity: 'Last Sign-In', invitation: 'Invitation', enabled: 'Enabled', created: 'Created',
            usertype: 'User Type', identity: 'Identity', access: 'Access', sponsor: 'Sponsor', source: 'Source'
        };
        const valueLabels = {
            remove: 'Remove', review: 'Review', keep: 'Keep', onprem: 'On-Prem',
            active: 'Active', inactive: 'Inactive', stale: 'Stale', never: 'Never', unknown: 'Unknown',
            accepted: 'Accepted', pending: 'Pending', none: 'No Invitation State',
            yes: 'Yes', no: 'No',
            new: 'New Guest', lastyear: 'Last Year', over1y: '+1 Year',
            guest: 'Guest', member: 'External Member',
            entra: 'Entra ID', msa: 'Microsoft Account', otp: 'Email OTP', google: 'Google', other: 'Other', unredeemed: 'Not Redeemed',
            groups: 'Groups', teams: 'Teams', apps: 'Apps', roles: 'Directory Roles', noaccess: 'No Access',
            has: 'Has Sponsor', nosponsor: 'No Sponsor',
            cloud: 'Cloud'
        };

        function rowMatchesTag(row, group, value) {
            const d = row.dataset;
            switch (group) {
                case 'rec':        return d.rec === value;
                case 'activity':   return d.activity === value;
                case 'invitation': return d.invitation === value;
                case 'enabled':    return d.enabled === value;
                case 'created':    return d.createdbucket === value;
                case 'usertype':   return d.usertype === value;
                case 'identity':   return d.identity === value;
                case 'source':     return d.source === value;
                case 'sponsor':
                    if (value === 'has')       return d.sponsor === 'has';
                    if (value === 'nosponsor') return d.sponsor === 'none';
                    return false;
                case 'access': {
                    const g = parseInt(d.groupcount) || 0;
                    const t = parseInt(d.teamcount) || 0;
                    const a = parseInt(d.appcount) || 0;
                    const r = parseInt(d.rolecount) || 0;
                    if (value === 'groups')   return g > 0;
                    if (value === 'teams')    return t > 0;
                    if (value === 'apps')     return a > 0;
                    if (value === 'roles')    return r > 0;
                    if (value === 'noaccess') return d.collected === 'yes' && g === 0 && t === 0 && a === 0 && r === 0;
                    return false;
                }
            }
            return false;
        }

        let _searchDebounceHandle = null;

        function queueApplyFilters() {
            if (_searchDebounceHandle) {
                clearTimeout(_searchDebounceHandle);
            }
            _searchDebounceHandle = setTimeout(applyFilters, 150);
        }

        function rowVisible(row) {
            const search = document.getElementById('searchInput').value.toLowerCase().trim();
            if (search && !(row.dataset.search || '').includes(search)) return false;
            for (const group in activeFilters) {
                const vals = activeFilters[group];
                if (!vals || vals.size === 0) continue;
                let any = false;
                vals.forEach(v => { if (rowMatchesTag(row, group, v)) any = true; });
                if (!any) return false;
            }
            return true;
        }

        function applyFilters() {
            const rows = document.querySelectorAll('#reportTable tbody tr');
            let shown = 0;
            rows.forEach(row => {
                const ok = rowVisible(row);
                row.style.display = ok ? '' : 'none';
                if (ok) shown++;
            });
            document.getElementById('resultCount').textContent =
                'Showing ' + shown + ' of ' + rows.length + ' guests';
            renderActiveChips();

            const search = document.getElementById('searchInput').value.trim();
            let activeCount = (search ? 1 : 0);
            for (const g in activeFilters) activeCount += activeFilters[g].size;
            const summary = document.getElementById('filterSummary');
            if (activeCount > 0) {
                summary.textContent = activeCount + ' filter' + (activeCount > 1 ? 's' : '') +
                    ' \u00b7 showing ' + shown + ' of ' + rows.length;
                summary.style.display = '';
            } else {
                summary.style.display = 'none';
            }
        }

        function toggleFilterPanel(forceShow) {
            const panel = document.getElementById('controlsPanel');
            const btn = document.getElementById('filterToggle');
            const show = forceShow === true ? true : panel.hasAttribute('hidden');
            if (show) { panel.removeAttribute('hidden'); } else { panel.setAttribute('hidden', ''); }
            btn.setAttribute('aria-expanded', show ? 'true' : 'false');
            btn.querySelector('.caret').innerHTML = show ? '&#9662;' : '&#9656;';
        }

        function addFilter(group, value) {
            if (!activeFilters[group]) activeFilters[group] = new Set();
            if (!activeFilters[group].has(value)) {
                activeFilters[group].add(value);
                const tag = document.querySelector('.filter-tag[data-group="' + group + '"][data-value="' + value + '"]');
                if (tag) tag.classList.add('active');
            }
            applyFilters();
        }

        function toggleTag(el) {
            const group = el.dataset.group, value = el.dataset.value;
            if (!activeFilters[group]) activeFilters[group] = new Set();
            if (activeFilters[group].has(value)) {
                activeFilters[group].delete(value);
                el.classList.remove('active');
            } else {
                activeFilters[group].add(value);
                el.classList.add('active');
            }
            applyFilters();
        }

        function renderActiveChips() {
            const wrap = document.getElementById('activeChips');
            const chips = [];
            for (const group in activeFilters) {
                activeFilters[group].forEach(v => {
                    const tagEl = document.querySelector('.filter-tag[data-group="' + group + '"][data-value="' + v + '"]');
                    const colorClass = tagEl ? (Array.from(tagEl.classList).find(c => c.startsWith('c-')) || '') : '';
                    chips.push('<span class="chip ' + colorClass + '" onclick="removeFilter(\'' + group + '\',\'' + v + '\')">' + groupLabels[group] + ': ' + (valueLabels[v] || v) + ' &times;</span>');
                });
            }
            wrap.innerHTML = chips.join('');
            document.getElementById('clearAll').style.display = chips.length ? '' : 'none';
        }

        function removeFilter(group, value) {
            if (activeFilters[group]) activeFilters[group].delete(value);
            const el = document.querySelector('.filter-tag[data-group="' + group + '"][data-value="' + value + '"]');
            if (el) el.classList.remove('active');
            applyFilters();
        }

        function clearAllFilters() {
            for (const group in activeFilters) activeFilters[group].clear();
            document.querySelectorAll('.filter-tag.active').forEach(el => el.classList.remove('active'));
            document.getElementById('searchInput').value = '';
            applyFilters();
        }

        function computeTagCounts() {
            const rows = Array.from(document.querySelectorAll('#reportTable tbody tr'));
            document.querySelectorAll('.filter-tag').forEach(el => {
                const group = el.dataset.group, value = el.dataset.value;
                const n = rows.filter(r => rowMatchesTag(r, group, value)).length;
                el.querySelector('.cnt').textContent = '(' + n + ')';
            });
        }

        // Cells can carry a data-sort value (ISO date, recommendation priority, ...) that
        // sorts better than their displayed text
        let _sortCol = -1, _sortAsc = true;
        const _modalSortState = {};
        function sortTable(tableId, columnIndex, type) {
            const table = document.getElementById(tableId);
            if (!table) return;
            let state;
            if (tableId === 'reportTable') {
                if (_sortCol === columnIndex) { _sortAsc = !_sortAsc; } else { _sortCol = columnIndex; _sortAsc = type !== 'number'; }
                state = { col: _sortCol, asc: _sortAsc };
            } else {
                state = _modalSortState[tableId] || { col: -1, asc: true };
                if (state.col === columnIndex) { state.asc = !state.asc; } else { state.col = columnIndex; state.asc = type !== 'number'; }
                _modalSortState[tableId] = state;
            }
            const dir = state.asc ? 1 : -1;
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const cellValue = cell => (cell.dataset.sort !== undefined ? cell.dataset.sort : cell.textContent).trim();

            rows.sort((a, b) => {
                let aVal = cellValue(a.cells[columnIndex]);
                let bVal = cellValue(b.cells[columnIndex]);
                if (type === 'number') {
                    aVal = parseInt(aVal.replace(/[^0-9]/g, '')) || 0;
                    bVal = parseInt(bVal.replace(/[^0-9]/g, '')) || 0;
                    return (aVal - bVal) * dir;
                }
                return aVal.localeCompare(bVal) * dir;
            });

            rows.forEach(row => tbody.appendChild(row));

            table.querySelectorAll('thead th').forEach(th => {
                th.classList.remove('sort-asc', 'sort-desc');
            });
            table.querySelectorAll('thead th')[columnIndex]
                ?.classList.add(state.asc ? 'sort-asc' : 'sort-desc');
        }

        function exportCsv() {
            const headers = Array.from(document.querySelectorAll('#reportTable thead th'))
                .map(th => th.textContent.replace(/[\u25B2\u25BC]/g, '').trim());

            const visibleRows = Array.from(document.querySelectorAll('#reportTable tbody tr'))
                .filter(tr => tr.style.display !== 'none');

            // Neutralize leading formula characters so a hostile value (e.g. a guest display
            // name starting with '=') cannot execute as a formula when opened in Excel.
            const esc = v => {
                let s = String(v);
                if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
                return '"' + s.replace(/"/g, '""') + '"';
            };

            const lines = [headers.map(esc).join(';')];
            visibleRows.forEach(tr => {
                const cells = Array.from(tr.cells).map(td => {
                    const clone = td.cloneNode(true);
                    clone.querySelectorAll('.csv-exclude').forEach(n => n.remove());
                    return esc(clone.textContent.trim().replace(/\s+/g, ' '));
                });
                lines.push(cells.join(';'));
            });

            // UTF-8 BOM ensures Excel opens the file with correct encoding
            const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement('a');
            a.href     = url;
            a.download = 'EntraIDGuestReport_' + new Date().toISOString().slice(0, 10) + '.csv';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        let _modalTrigger = null;

        function openDetailModal(title, html) {
            _modalTrigger = document.activeElement;
            document.getElementById('detailModalTitle').textContent = title;
            document.getElementById('detailModalBody').innerHTML = html;
            document.getElementById('detailModalOverlay').classList.add('active');
            document.querySelector('#detailModalOverlay .modal-box').scrollTop = 0;
            document.querySelector('.modal-close').focus();
            for (const id in _modalSortState) delete _modalSortState[id];
        }

        function closeDetailModal() {
            document.getElementById('detailModalOverlay').classList.remove('active');
            if (_modalTrigger) { _modalTrigger.focus(); _modalTrigger = null; }
        }

        function toggleTheme() {
            const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
            const next = isDark ? 'light' : 'dark';
            document.documentElement.setAttribute('data-theme', next);
            document.getElementById('themeToggle').innerHTML = next === 'dark' ? '&#9728;&#65039;' : '&#127769;';
            localStorage.setItem('entra-report-theme', next);
        }

        (function () {
            const saved = localStorage.getItem('entra-report-theme') || 'light';
            document.documentElement.setAttribute('data-theme', saved);
            document.getElementById('themeToggle').innerHTML = saved === 'dark' ? '&#9728;&#65039;' : '&#127769;';

            document.getElementById('reportTable').addEventListener('click', function (e) {
                const badge = e.target.closest('.clickable-badge');
                if (!badge) return;
                const group = badge.dataset.fg, value = badge.dataset.fv;
                if (!group || !value) return;
                addFilter(group, value);
                toggleFilterPanel(true);
            });

            // A card can map to several values of one group (e.g. Inactive = inactive + stale)
            document.querySelector('.summary').addEventListener('click', function (e) {
                const card = e.target.closest('.summary-card');
                if (!card) return;
                if (card.dataset.clear) { clearAllFilters(); return; }
                const group = card.dataset.fg, value = card.dataset.fv;
                if (!group || !value) return;
                value.split(',').forEach(v => addFilter(group, v));
                toggleFilterPanel(true);
                document.getElementById('reportTable').scrollIntoView({ behavior: 'smooth', block: 'start' });
            });

            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') closeDetailModal();
            });

            computeTagCounts();
            applyFilters();
        })();
    </script>

    <div id="detailModalOverlay" class="modal-overlay" onclick="if(event.target===this) closeDetailModal()">
        <div class="modal-box" role="dialog" aria-modal="true" aria-labelledby="detailModalTitle" aria-describedby="detailModalBody">
            <button class="modal-close" onclick="closeDetailModal()" aria-label="Close">&times;</button>
            <h3 id="detailModalTitle"></h3>
            <div id="detailModalBody"></div>
        </div>
    </div>
    </div>
'@)

[void]$sb.Append(@"

    <footer class="report-footer">
        Copyright &copy; 2026 <a href="https://www.matej.guru/about" target="_blank" rel="noopener">Matej Klemencic</a> &middot; $ScriptVersion
    </footer>
</body>
</html>
"@)

$html = $sb.ToString()

# Save the HTML report - UTF-8 without BOM for consistent output across PS 5.1 and PS 7
if ($DryRun) {
    Write-Host "DryRun: skipping report write to '$OutputPath'" -ForegroundColor Yellow
} else {
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Host "Report generated successfully: $OutputPath" -ForegroundColor Green

    # Open the report in default browser (interactive mode only)
    if (-not $NonInteractive) {
        Write-Host "Opening report in default browser..." -ForegroundColor Green
        Start-Process ([System.IO.Path]::GetFullPath($OutputPath))
    }
}

# Display summary statistics in console
Write-Host "`n=== REPORT SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total users: $totalGuests" -ForegroundColor White
Write-Host "  - Guests: $guestTypeCount" -ForegroundColor White
Write-Host "  - External members: $externalMemberCount" -ForegroundColor White
Write-Host "`nCleanup Recommendation:" -ForegroundColor White
Write-Host "  - Remove: $removeCount" -ForegroundColor Red
Write-Host "  - Review: $reviewCount" -ForegroundColor Yellow
Write-Host "  - Keep: $keepCount" -ForegroundColor Green
Write-Host "  - On-premises synced: $onPremCount" -ForegroundColor Gray
Write-Host "`nActivity Analysis:" -ForegroundColor White
if ($script:signInDataAvailable) {
    Write-Host "  - Never signed in: $neverSignedInCount" -ForegroundColor DarkYellow
    Write-Host "  - Inactive ($InactiveDays+ days): $inactiveCount" -ForegroundColor DarkYellow
    Write-Host "  - Stale ($RemoveAfterDays+ days): $staleCount" -ForegroundColor Red
} else {
    Write-Host "  - Sign-in activity unavailable (Entra ID P1/P2 and AuditLog.Read.All required)" -ForegroundColor DarkYellow
}
Write-Host "  - Pending invitations: $pendingCount" -ForegroundColor DarkYellow
Write-Host "  - Disabled accounts: $disabledCount" -ForegroundColor Gray
Write-Host "`nGovernance Analysis:" -ForegroundColor White
Write-Host "  - Guests with directory roles: $privilegedCount" -ForegroundColor Red
if (-not $SkipAccessDetails) {
    Write-Host "  - Guests without a sponsor: $noSponsorCount" -ForegroundColor DarkYellow
}

# Azure DevOps integration - emit pipeline-native signals when running on an ADO agent
# (TF_BUILD is set by the agent). No effect on local or interactive runs.
if ($env:TF_BUILD -eq 'True') {
    if ($privilegedCount -gt 0) {
        Write-Host "##vso[task.logissue type=warning]$privilegedCount guest(s) hold Entra ID directory roles - review the EntraIDGuestReport artifact."
        Write-Host "##vso[task.complete result=SucceededWithIssues;]Guests with directory roles found"
    }

    # Markdown summary shown as a tab on the pipeline run (same numbers as the console summary)
    $adoSummaryMd = @"
# Entra ID Guest Report - $tenantName

Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') &middot; **$totalGuests** guests analyzed ($guestTypeCount guests, $externalMemberCount external members)

| Recommendation | Guests |
|----------------|--------|
| Remove | $removeCount |
| Review | $reviewCount |
| Keep | $keepCount |
| On-premises synced | $onPremCount |

| Signal | Guests |
|--------|--------|
| Pending invitations | $pendingCount |
| Never signed in | $(if ($script:signInDataAvailable) { $neverSignedInCount } else { 'n/a' }) |
| Inactive ($InactiveDays+ days) | $(if ($script:signInDataAvailable) { $inactiveCount } else { 'n/a' }) |
| Disabled | $disabledCount |
| Directory roles | $privilegedCount |
| No sponsor | $(if ($SkipAccessDetails) { 'n/a' } else { $noSponsorCount }) |

Thresholds: inactive $InactiveDays days, remove $RemoveAfterDays days, pending invitation $PendingInvitationDays days.

Download the full interactive HTML report from the **EntraIDGuestReport** pipeline artifact.
"@
    $adoSummaryPath = Join-Path ([System.IO.Path]::GetTempPath()) 'EntraIDGuestReport-summary.md'
    [System.IO.File]::WriteAllText($adoSummaryPath, $adoSummaryMd, [System.Text.UTF8Encoding]::new($false))
    Write-Host "##vso[task.uploadsummary]$adoSummaryPath"
}

    Write-Host "`nScript completed successfully!" -ForegroundColor Green

} finally {
    # Guaranteed to run even if the script throws - keeps Graph sessions clean
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

exit 0
