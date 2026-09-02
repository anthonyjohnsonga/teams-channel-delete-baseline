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

> All three are required. `Microsoft.Graph.Groups` supplies `Get-MgGroup`, which the script
> uses to enumerate teams — it is easy to miss because the script's own header only lists
> the other two.

**Permissions**

| Mode | What you need |
|---|---|
| `Interactive` | Sign in as **Teams Administrator** or **Global Administrator**. The script requests the `TeamSettings.ReadWrite.All` and `Group.Read.All` delegated scopes at sign-in. |
| `ManagedIdentity` | Grant the managed identity the `TeamSettings.ReadWrite.All` and `Group.Read.All` **application** roles. These must be assigned through Graph — the Azure portal UI cannot assign them. |

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
Remediated: 12 | Already compliant: 84 | Errors: 2
```

And a CSV report saved to a `Reports` folder under the directory you run the script from —
created automatically if it does not exist — with one row per team:

| Column | Description |
|---|---|
| `Team` | Team display name |
| `TeamId` | Group / team object ID |
| `Before` | Value of `allowDeleteChannels` before the run |
| `Action` | `Remediated`, `AlreadyCompliant`, `WhatIf`, or `Error` |
| `Detail` | Error message, where applicable |

The CSV is written in both dry-run and apply modes.

---

## Notes

- **This sets a point-in-time baseline, not a permanent one.** No tenant-level policy pins
  `allowDeleteChannels`, so a team owner can re-enable it from the Teams client afterwards.
  Re-run the script to re-establish the baseline.
- **Archived teams are expected to fail** and will appear in the CSV as `Error`. This is
  normal and does not require action.
- **The `Reports` folder is created under your current working directory**, so run the
  script from the same place each time if you want the reports collected together. Reports
  are excluded from this repository by `.gitignore` — they contain team display names and
  object IDs.
- **In an Azure Automation runbook the working directory is the job sandbox**, which is
  discarded when the job ends, taking the report with it. Capture the console output, or
  point `-LogPath` somewhere durable.
