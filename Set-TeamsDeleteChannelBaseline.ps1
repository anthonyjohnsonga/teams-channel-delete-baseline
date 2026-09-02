#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams

<#
.SYNOPSIS
    Disables "Allow members to delete and restore channels" on all Teams in the tenant.

.DESCRIPTION
    One-off remediation pass. Enumerates every team and sets
    memberSettings.allowDeleteChannels to false where it is currently true.

    Note that this sets a point-in-time baseline, not a permanent one: there is no
    tenant-level policy that pins allowDeleteChannels, and a team owner can re-enable
    it from the Teams client afterwards. The script is idempotent, so it is safe to
    run again if you need to re-establish the baseline.

    Two auth modes:
      Interactive     - signs you in with a browser. For ad-hoc runs and testing.
      ManagedIdentity - unattended execution, e.g. an Azure Automation runbook.
                        No secrets to store or rotate.

.NOTES
    Requirements.

    Three Microsoft Graph modules. Groups is easy to overlook - it supplies Get-MgGroup,
    which this script uses to enumerate teams.
        Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams -Scope CurrentUser

    Interactive mode: sign in as Teams Administrator or Global Administrator.

    Managed identity mode: grant the automation account's managed identity these
    application roles (via Graph, they cannot be assigned in the portal UI):
        TeamSettings.ReadWrite.All
        Group.Read.All

.EXAMPLE
    # Report only - run this first
    .\Set-TeamsDeleteChannelBaseline.ps1 -WhatIf

.EXAMPLE
    # Apply, signed in interactively
    .\Set-TeamsDeleteChannelBaseline.ps1

.EXAMPLE
    # Inside an Azure Automation runbook
    .\Set-TeamsDeleteChannelBaseline.ps1 -AuthMode ManagedIdentity
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Interactive', 'ManagedIdentity')]
    [string] $AuthMode = 'Interactive',

    # Reports land in a Reports folder under the working directory the script is run from.
    # Resolved to an absolute path here so the path echoed at the end is unambiguous.
    [string] $LogPath = (Join-Path `
        (Join-Path (Get-Location).Path 'Reports') `
        "TeamsChannelBaseline_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv")
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()

# --- Connect -----------------------------------------------------------------

switch ($AuthMode) {
    'Interactive' {
        Connect-MgGraph -Scopes 'TeamSettings.ReadWrite.All', 'Group.Read.All' -NoWelcome
    }
    'ManagedIdentity' {
        Connect-MgGraph -Identity -NoWelcome
    }
}

$context  = Get-MgContext
$tenantId = $context.TenantId
$runStamp = Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'

Write-Host "Connected to tenant $tenantId" -ForegroundColor Cyan

# Pre-flight the permissions. Connect-MgGraph succeeds even when admin consent was never
# granted, and without this the first sign of trouble is a 403 on every team in turn.
# Each requirement lists the scopes that satisfy it, so holding a stronger scope than the
# minimum does not trip the check.
$requiredScopes = [ordered]@{
    'Group.Read.All'             = @('Group.Read.All', 'Group.ReadWrite.All',
                                     'Directory.Read.All', 'Directory.ReadWrite.All')
    'TeamSettings.ReadWrite.All' = @('TeamSettings.ReadWrite.All')
}

if (-not $context.Scopes) {
    # App-only tokens do not always surface their roles here. Warn rather than block:
    # refusing to run on a check we cannot perform would be worse than a clear 403.
    Write-Warning 'Could not read granted permissions from the Graph context - skipping the pre-flight check.'
}
else {
    $missing = foreach ($requirement in $requiredScopes.Keys) {
        if (-not ($requiredScopes[$requirement] | Where-Object { $_ -in $context.Scopes })) {
            $requirement
        }
    }

    if ($missing) {
        throw ("Missing required Graph permission(s): {0}. " -f ($missing -join ', ')) +
              'The sign-in succeeded but these were not consented, so every team would fail ' +
              'with a 403. An administrator needs to grant them before this script can run.'
    }
}

# Every row goes through here so the CSV cannot go ragged, and so the tenant and run
# timestamp are stamped on each row rather than living only in the filename and the
# console. Before is a nullable boolean - left null where the value could not be read,
# so the column stays one type instead of mixing booleans with a magic string.
function Add-Result {
    param(
        [Parameter(Mandatory)] $Team,
        [Parameter(Mandatory)] [string] $Action,
        [nullable[bool]] $Before = $null,
        [nullable[int]] $StatusCode = $null,
        [string] $Detail = ''
    )
    $results.Add([pscustomobject]@{
        RunTimestamp = $runStamp
        TenantId     = $tenantId
        Team         = $Team.DisplayName
        TeamId       = $Team.Id
        Before       = $Before
        Action       = $Action
        StatusCode   = $StatusCode
        Detail       = $Detail
    })
}

# Graph surfaces the HTTP status in different places depending on which layer threw, so
# probe the common shapes rather than assuming one. Returns null if none of them carry it.
function Get-HttpStatus {
    param($ErrorRecord)
    $ex = $ErrorRecord.Exception
    $raw = if ($null -ne $ex.Response -and $null -ne $ex.Response.StatusCode) { $ex.Response.StatusCode }
           elseif ($null -ne $ex.HttpStatus) { $ex.HttpStatus }
           elseif ($null -ne $ex.StatusCode) { $ex.StatusCode }
           else { $null }

    if ($null -eq $raw) { return $null }
    try { [int] $raw } catch { $null }
}

# A 429 or a 5xx means the team was not remediated but nothing is wrong with the team -
# the SDK already retried and gave up, and a later re-run will pick it up. Filing those
# as Error makes a transient blip indistinguishable from a genuinely broken team, so
# they are labelled and counted separately.
function Add-Failure {
    param(
        [Parameter(Mandatory)] $Team,
        [Parameter(Mandatory)] $ErrorRecord,
        [nullable[bool]] $Before = $null
    )
    $status  = Get-HttpStatus $ErrorRecord
    $message = $ErrorRecord.Exception.Message

    if ($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) {
        $script:retryable++
        Write-Warning "[RETRY] $($Team.DisplayName) (HTTP $status): $message"
        Add-Result -Team $Team -Action 'Retryable' -Before $Before -StatusCode $status -Detail $message
    }
    else {
        $script:failed++
        Write-Warning "[FAIL] $($Team.DisplayName): $message"
        Add-Result -Team $Team -Action 'Error' -Before $Before -StatusCode $status -Detail $message
    }
}

# --- Enumerate ---------------------------------------------------------------

# Filtering on groups rather than Get-MgTeam: pages predictably and returns
# displayName in the same call, so no second lookup per team.
$teams = Get-MgGroup -All `
    -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
    -Property 'id,displayName'

Write-Host "Found $($teams.Count) team(s).`n" -ForegroundColor Gray

# --- Remediate ---------------------------------------------------------------

$changed = 0; $already = 0; $skipped = 0; $retryable = 0; $failed = 0

foreach ($team in $teams) {

    try {
        $current = Get-MgTeam -TeamId $team.Id -ErrorAction Stop
    }
    catch {
        # Groups still provisioning commonly fail here.
        Add-Failure -Team $team -ErrorRecord $_
        continue
    }

    # No memberSettings means there is nothing to read and nothing safe to write back.
    # Without this guard a null would fall through to the update below and be recorded
    # as remediated.
    if ($null -eq $current.MemberSettings) {
        Write-Warning "[SKIP] $($team.DisplayName): team returned no memberSettings"
        $failed++
        Add-Result -Team $team -Action 'Error' -Detail 'Team returned no memberSettings'
        continue
    }

    if ($current.MemberSettings.AllowDeleteChannels -eq $false) {
        $already++
        Add-Result -Team $team -Action 'AlreadyCompliant' -Before $false
        continue
    }

    # Archived teams reject writes with a permanent 403, so they are an expected skip
    # rather than a failure. Checked after the compliance test above, so an archived team
    # that is already compliant is reported as compliant - which it is - and only teams
    # that would actually need a write land here.
    if ($current.IsArchived) {
        Write-Host "[SKIP] $($team.DisplayName): archived" -ForegroundColor DarkGray
        $skipped++
        Add-Result -Team $team -Action 'Skipped-Archived' -Before $true -Detail 'Team is archived'
        continue
    }

    # Send all six memberSettings back, not just the one being changed. Graph normally
    # merges a partial complex type, but if it ever stopped doing so a partial PATCH
    # would reset the other five to defaults on every team in the tenant, silently, and
    # the report would still say Remediated. The full object is already in hand from the
    # read above, so preserving the siblings explicitly costs nothing.
    $newSettings = @{
        allowCreateUpdateChannels         = $current.MemberSettings.AllowCreateUpdateChannels
        allowCreatePrivateChannels        = $current.MemberSettings.AllowCreatePrivateChannels
        allowDeleteChannels               = $false
        allowAddRemoveApps                = $current.MemberSettings.AllowAddRemoveApps
        allowCreateUpdateRemoveTabs       = $current.MemberSettings.AllowCreateUpdateRemoveTabs
        allowCreateUpdateRemoveConnectors = $current.MemberSettings.AllowCreateUpdateRemoveConnectors
    }

    if ($PSCmdlet.ShouldProcess($team.DisplayName, 'Set allowDeleteChannels = false')) {
        try {
            Update-MgTeam -TeamId $team.Id -MemberSettings $newSettings
            Write-Host "[FIXED] $($team.DisplayName)" -ForegroundColor Yellow
            $changed++
            Add-Result -Team $team -Action 'Remediated' -Before $true
        }
        catch {
            Add-Failure -Team $team -ErrorRecord $_ -Before $true
        }
    }
    else {
        Write-Host "[WOULD FIX] $($team.DisplayName)" -ForegroundColor DarkYellow
        Add-Result -Team $team -Action 'WhatIf' -Before $true
    }
}

# --- Report ------------------------------------------------------------------

Write-Host "`nRemediated: $changed | Already compliant: $already | Skipped (archived): $skipped | Retryable: $retryable | Errors: $failed" -ForegroundColor Green

if ($retryable -gt 0) {
    Write-Host "$retryable team(s) hit throttling or a server error and were not changed. Re-run to pick them up." -ForegroundColor Yellow
}

# Create the report folder if it does not exist yet. New-Item supports ShouldProcess,
# so it needs -WhatIf:$false for the same reason Export-Csv does - see below.
$reportDir = Split-Path -Parent $LogPath
if ($reportDir -and -not (Test-Path -LiteralPath $reportDir)) {
    New-Item -ItemType Directory -Path $reportDir -Force -WhatIf:$false | Out-Null
}

# -WhatIf:$false on the export below. Export-Csv supports ShouldProcess, so it
# inherits $WhatIfPreference from the script scope and gets suppressed during a
# dry run. The report should write in both modes.
$results | Export-Csv -Path $LogPath -NoTypeInformation -WhatIf:$false
Write-Host "Report written to $LogPath" -ForegroundColor Cyan

# Disconnect-MgGraph does not implement ShouldProcess - no -WhatIf to suppress.
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

# Exit non-zero if any team was left unremediated by a problem, so an unattended caller
# does not read a completed job as a clean one. Archived skips are expected and do not
# count; retryable failures do, because that work is genuinely still outstanding.
if ($failed -gt 0 -or $retryable -gt 0) { exit 1 }
exit 0
