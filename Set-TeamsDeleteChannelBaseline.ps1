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

.REQUIREMENTS
    Microsoft.Graph.Authentication and Microsoft.Graph.Teams modules.
        Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Teams -Scope CurrentUser

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

Write-Host "Connected to tenant $((Get-MgContext).TenantId)" -ForegroundColor Cyan

# --- Enumerate ---------------------------------------------------------------

# Filtering on groups rather than Get-MgTeam: pages predictably and returns
# displayName in the same call, so no second lookup per team.
$teams = Get-MgGroup -All `
    -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
    -Property 'id,displayName'

Write-Host "Found $($teams.Count) team(s).`n" -ForegroundColor Gray

# --- Remediate ---------------------------------------------------------------

$changed = 0; $already = 0; $failed = 0

foreach ($team in $teams) {

    try {
        $current = Get-MgTeam -TeamId $team.Id -ErrorAction Stop
    }
    catch {
        # Archived teams and groups still provisioning commonly fail here.
        Write-Warning "[SKIP] $($team.DisplayName): $($_.Exception.Message)"
        $failed++
        $results.Add([pscustomobject]@{
            Team = $team.DisplayName; TeamId = $team.Id
            Before = 'unknown'; Action = 'Error'; Detail = $_.Exception.Message
        })
        continue
    }

    if ($current.MemberSettings.AllowDeleteChannels -eq $false) {
        $already++
        $results.Add([pscustomobject]@{
            Team = $team.DisplayName; TeamId = $team.Id
            Before = $false; Action = 'AlreadyCompliant'; Detail = ''
        })
        continue
    }

    if ($PSCmdlet.ShouldProcess($team.DisplayName, 'Set allowDeleteChannels = false')) {
        try {
            Update-MgTeam -TeamId $team.Id -MemberSettings @{ allowDeleteChannels = $false }
            Write-Host "[FIXED] $($team.DisplayName)" -ForegroundColor Yellow
            $changed++
            $results.Add([pscustomobject]@{
                Team = $team.DisplayName; TeamId = $team.Id
                Before = $true; Action = 'Remediated'; Detail = ''
            })
        }
        catch {
            Write-Warning "[FAIL] $($team.DisplayName): $($_.Exception.Message)"
            $failed++
            $results.Add([pscustomobject]@{
                Team = $team.DisplayName; TeamId = $team.Id
                Before = $true; Action = 'Error'; Detail = $_.Exception.Message
            })
        }
    }
    else {
        Write-Host "[WOULD FIX] $($team.DisplayName)" -ForegroundColor DarkYellow
        $results.Add([pscustomobject]@{
            Team = $team.DisplayName; TeamId = $team.Id
            Before = $true; Action = 'WhatIf'; Detail = ''
        })
    }
}

# --- Report ------------------------------------------------------------------

Write-Host "`nRemediated: $changed | Already compliant: $already | Errors: $failed" -ForegroundColor Green

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
