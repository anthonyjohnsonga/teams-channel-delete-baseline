# teams-channel-delete-baseline

Disables **"Allow members to delete and restore channels"** across every team in a
Microsoft 365 tenant.

`Set-TeamsDeleteChannelBaseline.ps1` enumerates all Teams-provisioned groups, checks each
team's `memberSettings.allowDeleteChannels`, and sets it to `false` wherever it is
currently `true`. It reports on every team it touches and writes a CSV audit trail. The
script is idempotent — teams already compliant are recorded and skipped, so it is safe to
run more than once.

---

## Requirements

**PowerShell modules**

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams -Scope CurrentUser
```

All three are required. `Microsoft.Graph.Groups` supplies `Get-MgGroup`, which enumerates
the teams. The script declares these with `#Requires`, so a missing module stops it
immediately rather than failing partway through.

**Permissions**

| Mode | What you need |
|---|---|
| `Interactive` | Sign in as **Teams Administrator** or **Global Administrator**. The script requests the `TeamSettings.ReadWrite.All` and `Group.Read.All` delegated scopes at sign-in. |
| `ManagedIdentity` | Grant the managed identity the `TeamSettings.ReadWrite.All` and `Group.Read.All` **application** roles. These must be assigned through Graph — the Azure portal UI cannot assign them. |

The script verifies these immediately after sign-in and stops with a clear message if they
were not consented, rather than failing with a 403 on every team in turn. Holding a
stronger scope than the minimum (for example `Group.ReadWrite.All` in place of
`Group.Read.All`) satisfies the check.

---

## Usage

### 1. Dry run first

Always start here. `-WhatIf` makes no changes but still writes the full CSV report, so you
get an accurate picture of what would be modified before committing to anything.

```powershell
.\Set-TeamsDeleteChannelBaseline.ps1 -WhatIf
```

Teams that would be changed are listed as `[WOULD FIX]`. Review the CSV, confirm the scope
looks right, then proceed.

### 2. Apply

```powershell
.\Set-TeamsDeleteChannelBaseline.ps1
```

Signs you in through the browser and applies the change. Each remediated team prints as
`[FIXED]`.

### 3. Apply unattended

For an Azure Automation runbook or any other non-interactive host:

```powershell
.\Set-TeamsDeleteChannelBaseline.ps1 -AuthMode ManagedIdentity
```

Uses the host's system-assigned managed identity — no secrets to store or rotate.

### Options

| Parameter | Default | Description |
|---|---|---|
| `-AuthMode` | `Interactive` | `Interactive` or `ManagedIdentity`. |
| `-LogPath` | `.\Reports\TeamsChannelBaseline_<timestamp>.csv` | Where to write the CSV report. The parent folder is created if it does not exist. |
| `-WhatIf` | — | Report only; makes no changes. |

---

## Output

A summary line on the console:

```
Remediated: 12 | Already compliant: 84 | Skipped (archived): 3 | Retryable: 0 | Errors: 2
```

And a CSV report saved to a `Reports` folder under the directory you run the script from —
created automatically if it does not exist — with one row per team:

| Column | Description |
|---|---|
| `RunTimestamp` | When the run started, ISO 8601. Identical on every row. |
| `TenantId` | Tenant the run was made against |
| `Team` | Team display name |
| `TeamId` | Group / team object ID |
| `Before` | Value of `allowDeleteChannels` before the run. Empty where it could not be read. |
| `Action` | See below |
| `StatusCode` | HTTP status where the failure carried one. Empty otherwise. |
| `Detail` | Error message or skip reason, where applicable |

`Action` is one of:

| Value | Meaning |
|---|---|
| `Remediated` | Was `true`, has been set to `false` |
| `AlreadyCompliant` | Already `false`, left alone |
| `Skipped-Archived` | Would need changing, but the team is archived |
| `WhatIf` | Would be changed — dry run only, nothing was written |
| `Retryable` | Throttled (429) or a server error (5xx). Not changed; re-run picks it up. |
| `Error` | Could not be read or written; see `StatusCode` and `Detail` |

The CSV is written in both dry-run and apply modes.

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Every team was either remediated, already compliant, or an expected archived skip |
| `1` | At least one team was left unremediated by an error or throttling |

Archived skips do not cause a non-zero exit — they are expected. `Retryable` results do,
because that work is genuinely still outstanding.

---

## Notes

- **This sets a point-in-time baseline, not a permanent one.** No tenant-level policy pins
  `allowDeleteChannels`, so a team owner can re-enable it from the Teams client afterwards.
  Re-run the script to re-establish the baseline.
- **Archived teams are skipped, not failed.** They reject writes permanently, so they are
  reported as `Skipped-Archived` and counted separately from errors. An archived team that
  is already compliant is reported as `AlreadyCompliant` — no change was needed either way.
  To bring an archived team into line, unarchive it and re-run.
- **The `Reports` folder is created under your current working directory**, so run the
  script from the same place each time if you want the reports collected together. Reports
  are excluded from this repository by `.gitignore` — they contain team display names and
  object IDs.
- **In an Azure Automation runbook the working directory is the job sandbox**, which is
  discarded when the job ends, taking the report with it. Capture the console output, or
  point `-LogPath` somewhere durable.
