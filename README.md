# SPFolderMover.ps1

Interactive WinForms tool for migrating folders between SharePoint Online document libraries via Microsoft Graph API. Built for the MedPro Healthcare Staffing tenant but usable against any SharePoint Online site.

---

## Features

- **Two-step GUI wizard** — Step 1 selects source folders; Step 2 picks destination and executes.
- **Multi-select with Shift+Click** range support for batch operations.
- **Dry-run mode** (`-DryRun`) — previews all planned moves without making any changes.
- **Permission preservation** — detects unique (non-inherited) permissions on source folders and re-applies them after each move.
- **Self-skip / circular-move guard** — blocks moving a folder into itself or its own subfolder.
- **Retry logic** — configurable retry count/delay with automatic handling of 401 (token refresh), 429 (throttle/`Retry-After`), and 5xx transient errors.
- **Sortable folder tree** — toggle between "oldest first" (default) and "A–Z" via clickable column headers.
- **Drive/library selector** — switch between document libraries in both pickers.
- **In-tool progress log** — color-coded RichTextBox showing real-time move status.
- **CSV audit log** — every operation (SUCCESS / FAILED / DRY-RUN / SKIP) is written to a timestamped CSV.

---

## Requirements

| Requirement | Detail |
|---|---|
| PowerShell | 5.1 or later |
| OS | Windows (WinForms) |
| Module | [MSAL.PS](https://github.com/AzureAD/MSAL.NET) — auto-installed from PSGallery if missing |
| Authentication | Delegated / interactive (Graph Explorer client ID — no app registration required) |
| Permissions | `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Sites.FullControl.All` |

> **Note:** The delegated flow requires the signed-in user to have at least Contribute access on the target SharePoint site. `Sites.FullControl.All` is requested to support permission restoration; the actual effective permission is capped by what the signed-in user holds in the tenant.

---

## Installation

```powershell
# Option A – run directly
.\SPFolderMover.ps1 -SiteUrl "https://medprostaffing.sharepoint.com/sites/HR"

# Option B – if execution policy blocks unsigned scripts, strip the ADS Zone.Identifier first
Unblock-File .\SPFolderMover.ps1
.\SPFolderMover.ps1 -SiteUrl "https://medprostaffing.sharepoint.com/sites/HR"
```

MSAL.PS will be installed automatically to the current user scope on first run if not already present.

---

## Parameters

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `-SiteUrl` | `string` | **Yes** | — | Full SharePoint site URL, e.g. `https://medprostaffing.sharepoint.com/sites/HR` |
| `-TenantId` | `string` | No | `medprostaffing.onmicrosoft.com` | Entra tenant ID (GUID or domain) |
| `-DryRun` | `switch` | No | off | Preview all planned moves without making any changes |
| `-MaxRetries` | `int` | No | `3` | Number of retry attempts on transient Graph API errors |
| `-RetryDelaySeconds` | `int` | No | `5` | Seconds to wait between retries |
| `-LogPath` | `string` | No | `C:\temp\SPMove_<timestamp>.csv` | Path for the CSV audit log |

---

## Usage Examples

```powershell
# Dry-run – preview moves with no changes
.\SPFolderMover.ps1 -SiteUrl "https://medprostaffing.sharepoint.com/sites/HR" -DryRun

# Live run with increased retries
.\SPFolderMover.ps1 -SiteUrl "https://medprostaffing.sharepoint.com/sites/HR" -MaxRetries 5

# Different tenant / site, custom log location
.\SPFolderMover.ps1 `
    -TenantId  "contoso.onmicrosoft.com" `
    -SiteUrl   "https://contoso.sharepoint.com/sites/Finance" `
    -LogPath   "D:\logs\sp_migration.csv"
```

---

## Workflow

```
Launch script
     │
     ▼
MSAL interactive sign-in (browser pop-up)
     │
     ▼
Step 1 – Source Picker
  • Browse / expand folder tree
  • Check one or more folders  (Shift+Click for range)
  • Toggle sort: Name A-Z  |  Date Modified (oldest first)
  • Switch document library via the Library dropdown
  • Click "Move To >"
     │
     ▼
Step 2 – Destination Picker
  • Navigate to destination folder, single-click to select
  • Source folders summarized in amber bar at top
  • Click "Begin Move" → confirmation dialog
     │
     ▼
Execution (per folder)
  ├─ SKIP if source == destination
  ├─ SKIP if destination is a subfolder of source
  ├─ Read existing permissions
  ├─ PATCH parentReference (Graph API move)
  ├─ If unique permissions existed → POST each back to moved item
  └─ Log result (SUCCESS / FAILED / DRY-RUN / SKIP)
     │
     ▼
CSV log saved to LogPath
```

---

## CSV Log Format

| Column | Description |
|---|---|
| `Timestamp` | Local time of the operation |
| `SourcePath` | Display name of the source folder |
| `DestPath` | Display name of the destination folder |
| `Status` | `SUCCESS`, `FAILED`, `DRY-RUN`, or `SKIP` |
| `Detail` | `perms restored`, `inherited`, reason for skip, or error message |

Default log location: `C:\temp\SPMove_yyyyMMdd_HHmmss.csv`

---

## Authentication Details

The script uses the well-known **Graph Explorer client ID** (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) for delegated interactive auth — no Azure app registration is required. On subsequent runs within the same PowerShell session, MSAL attempts a silent token refresh from cache before falling back to interactive.

Token is automatically re-acquired mid-run on 401 responses.

---

## Security Notes

- **No credentials are stored or hardcoded.** Authentication is fully delegated; tokens are held in memory only for the lifetime of the session.
- **Least-privilege caveat:** `Sites.FullControl.All` is included in the scope request to enable permission re-application after moves. If your environment restricts that scope, moves will succeed but unique permission restoration will fail gracefully with a WARN log entry.
- **Audit trail:** Every operation is written to a CSV regardless of outcome. Retain logs for compliance or rollback triage.
- **Dry-run first:** Always run with `-DryRun` before executing against production libraries to validate the planned moves.

---

## Known Limitations

- Moves are **within the same tenant and drive**. Cross-tenant or cross-site moves are not supported by the Graph `PATCH parentReference` approach.
- Sharing-link–type permissions are intentionally skipped during restoration; only explicit role assignments (`grantedTo` + `roles`) are re-applied.
- The `lastModifiedDateTime` sort reflects the folder metadata date, not the most recent file modification inside the folder.
- WinForms requires an interactive Windows session — this script cannot run headless or as a scheduled task.

---

## Related Tools

| Tool | Purpose |
|---|---|
| `HDAgent.ps1` | New-hire provisioning across AD and M365 |
| `ExchangeAdminTool.ps1` | Exchange Online administration GUI |
| `Win11_Tuner.ps1` | Windows 11 endpoint optimization |

---

*Internal tool — MedPro Healthcare Staffing / StackPoint IT*
