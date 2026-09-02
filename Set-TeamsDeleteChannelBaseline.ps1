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

$tenantId = (Get-MgContext).TenantId
$runStamp = Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'

Write-Host "Connected to tenant $tenantId" -ForegroundColor Cyan

# Every row goes through here so the CSV cannot go ragged, and so the tenant and run
# timestamp are stamped on each row rather than living only in the filename and the
# console. Before is a nullable boolean - left null where the value could not be read,
# so the column stays one type instead of mixing booleans with a magic string.
function Add-Result {
    param(
        [Parameter(Mandatory)] $Team,
        [Parameter(Mandatory)] [string] $Action,
        [nullable[bool]] $Before = $null,
        [string] $Detail = ''
    )
    $results.Add([pscustomobject]@{
        RunTimestamp = $runStamp
        TenantId     = $tenantId
        Team         = $Team.DisplayName
        TeamId       = $Team.Id
        Before       = $Before
        Action       = $Action
        Detail       = $Detail
    })
}

# --- Enumerate ---------------------------------------------------------------

# Filtering on groups rather than Get-MgTeam: pages predictably and returns
# displayName in the same call, so no second lookup per team.
$teams = Get-MgGroup -All `
    -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
    -Property 'id,displayName'

Write-Host "Found $($teams.Count) team(s).`n" -ForegroundColor Gray

# --- Remediate ---------------------------------------------------------------

$changed = 0; $already = 0; $skipped = 0; $failed = 0

foreach ($team in $teams) {

    try {
        $current = Get-MgTeam -TeamId $team.Id -ErrorAction Stop
    }
    catch {
        # Groups still provisioning commonly fail here.
        Write-Warning "[SKIP] $($team.DisplayName): $($_.Exception.Message)"
        $failed++
        Add-Result -Team $team -Action 'Error' -Detail $_.Exception.Message
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
            Write-Warning "[FAIL] $($team.DisplayName): $($_.Exception.Message)"
            $failed++
            Add-Result -Team $team -Action 'Error' -Before $true -Detail $_.Exception.Message
        }
    }
    else {
        Write-Host "[WOULD FIX] $($team.DisplayName)" -ForegroundColor DarkYellow
        Add-Result -Team $team -Action 'WhatIf' -Before $true
    }
}

# --- Report ------------------------------------------------------------------

Write-Host "`nRemediated: $changed | Already compliant: $already | Skipped (archived): $skipped | Errors: $failed" -ForegroundColor Green

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
