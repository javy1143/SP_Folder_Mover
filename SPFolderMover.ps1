#Requires -Version 5.1
<#
.SYNOPSIS
    SPFolderMover.ps1 - Interactive SharePoint folder migration tool for MedPro.

.DESCRIPTION
    Authenticates via MSAL delegated (Graph Explorer client ID - no app registration needed).
    Presents a WinForms tree picker to select source folder(s) and a destination.
    Supports dry-run preview, permission preservation, retry logic, and CSV logging.

.PARAMETER TenantId
    Your Entra tenant ID (GUID or domain). Defaults to medprostaffing.onmicrosoft.com.

.PARAMETER SiteUrl
    Full SharePoint site URL, e.g. https://medprostaffing.sharepoint.com/sites/HR

.PARAMETER DryRun
    Switch. Preview all planned moves without making any changes.

.PARAMETER MaxRetries
    Number of retry attempts on transient Graph API errors. Default: 3.

.PARAMETER RetryDelaySeconds
    Seconds to wait between retries. Default: 5.

.PARAMETER LogPath
    Path for the CSV move log. Defaults to C:\temp\SPMove_<timestamp>.csv

.EXAMPLE
    .\SPFolderMover.ps1 -SiteUrl "https://medprostaffing.sharepoint.com/sites/HR" -DryRun
    .\SPFolderMover.ps1 -SiteUrl "https://medprostaffing.sharepoint.com/sites/HR" -MaxRetries 5
#>

[CmdletBinding()]
param(
    [string]$TenantId     = "medprostaffing.onmicrosoft.com",
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    [switch]$DryRun,
    [int]$MaxRetries        = 3,
    [int]$RetryDelaySeconds = 5,
    [string]$LogPath        = "C:\temp\SPMove_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -----------------------------------------------
# CONSTANTS
# -----------------------------------------------
$GraphExplorerClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
$GraphBaseUrl          = "https://graph.microsoft.com/v1.0"
$Scopes = @(
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "https://graph.microsoft.com/Files.ReadWrite.All",
    "https://graph.microsoft.com/Sites.FullControl.All",
    "offline_access"
)

# -----------------------------------------------
# LOGGING
# -----------------------------------------------
$LogEntries = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DRY-RUN","SUCCESS")]
        [string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "WARN"    { "Yellow" }
        "ERROR"   { "Red" }
        "DRY-RUN" { "Cyan" }
        "SUCCESS" { "Green" }
        default   { "White" }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}

function Add-LogEntry {
    param([string]$SourcePath,[string]$DestPath,[string]$Status,[string]$Detail="")
    $LogEntries.Add([PSCustomObject]@{
        Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        SourcePath = $SourcePath
        DestPath   = $DestPath
        Status     = $Status
        Detail     = $Detail
    })
}

function Save-Log {
    try {
        $dir = Split-Path $LogPath -Parent
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $LogEntries | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
        Write-Log "Log saved to $LogPath" -Level SUCCESS
    } catch {
        Write-Log "Could not save log: $_" -Level WARN
    }
}

# -----------------------------------------------
# PREREQUISITE CHECK
# -----------------------------------------------
function Assert-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "Module '$Name' not found. Installing from PSGallery..." -Level WARN
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Installed '$Name' successfully." -Level SUCCESS
        } catch {
            Write-Log "Failed to install '$Name': $_" -Level ERROR
            throw
        }
    }
    Import-Module $Name -ErrorAction Stop
}

# -----------------------------------------------
# AUTHENTICATION (MSAL - delegated interactive)
# -----------------------------------------------
$script:AccessToken = $null

function Get-GraphToken {
    Write-Log "Acquiring token via MSAL interactive flow..."
    try {
        $tokenResult = Get-MsalToken `
            -ClientId   $GraphExplorerClientId `
            -TenantId   $TenantId `
            -Scopes     $Scopes `
            -Interactive `
            -ErrorAction Stop
        $script:AccessToken = $tokenResult.AccessToken
        Write-Log "Authentication successful. Token acquired." -Level SUCCESS
    } catch {
        Write-Log "Authentication failed: $_" -Level ERROR
        throw
    }
}

# Silent token refresh - tries cached token first, falls back to interactive
function Refresh-GraphToken {
    try {
        $tokenResult = Get-MsalToken `
            -ClientId   $GraphExplorerClientId `
            -TenantId   $TenantId `
            -Scopes     $Scopes `
            -Silent `
            -ErrorAction Stop
        $script:AccessToken = $tokenResult.AccessToken
        Write-Log "Token refreshed silently." -Level SUCCESS
    } catch {
        Write-Log "Silent refresh failed, falling back to interactive..." -Level WARN
        Get-GraphToken
    }
}

function Get-AuthHeaders {
    if (-not $script:AccessToken) { Get-GraphToken }
    return @{
        Authorization  = "Bearer $script:AccessToken"
        "Content-Type" = "application/json"
    }
}

# -----------------------------------------------
# GRAPH API HELPER - with retry
# -----------------------------------------------
function Invoke-GraphRequest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body   = $null,
        [int]$Retries   = $MaxRetries
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Uri     = $Uri
                Method  = $Method
                Headers = Get-AuthHeaders
                ErrorAction = "Stop"
            }
            if ($Body) {
                $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress)
            }
            $response = Invoke-RestMethod @params
            return $response
        } catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # Re-auth on 401
            if ($statusCode -eq 401 -and $attempt -le $Retries) {
                Write-Log "Token expired (401). Re-authenticating..." -Level WARN
                Get-GraphToken
                continue
            }

            # Throttling - respect Retry-After if present
            if ($statusCode -eq 429 -and $attempt -le $Retries) {
                $retryAfter = $RetryDelaySeconds
                try {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                } catch {}
                Write-Log "Throttled (429). Waiting $retryAfter s before retry $attempt/$Retries..." -Level WARN
                Start-Sleep -Seconds $retryAfter
                continue
            }

            # Transient server errors
            if ($statusCode -in @(500,502,503,504) -and $attempt -le $Retries) {
                Write-Log "Transient error ($statusCode). Retry $attempt/$Retries in $RetryDelaySeconds s..." -Level WARN
                Start-Sleep -Seconds $RetryDelaySeconds
                continue
            }

            Write-Log "Graph API error on $Method $Uri - $_" -Level ERROR
            throw
        }
    }
}

# -----------------------------------------------
# SHAREPOINT SITE + DRIVE RESOLUTION
# -----------------------------------------------
$script:SiteId  = $null
$script:DriveId = $null

function Resolve-Site {
    Write-Log "Resolving site: $SiteUrl"

    # Normalize - add https:// if missing
    $normalizedUrl = $SiteUrl
    if ($normalizedUrl -notmatch '^https?://') {
        $normalizedUrl = "https://$normalizedUrl"
    }

    $uri      = [System.Uri]$normalizedUrl
    $siteHost = $uri.Host
    $sitePath = $uri.AbsolutePath.TrimStart("/").TrimEnd("/")

    # Root site vs named site need different Graph endpoints
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $graphUri = "$GraphBaseUrl/sites/$siteHost"
    } else {
        $graphUri = "$GraphBaseUrl/sites/${siteHost}:/${sitePath}"
    }

    Write-Log "Graph site URI: $graphUri"
    $site = Invoke-GraphRequest -Uri $graphUri

    if (-not $site -or [string]::IsNullOrWhiteSpace($site.id)) {
        throw "Graph API returned no site for '$normalizedUrl'. Check the SiteUrl and that admin consent is granted."
    }

    $script:SiteId = $site.id
    Write-Log "Site resolved - '$($site.displayName)' (ID: $($script:SiteId))" -Level SUCCESS
}

function Get-Drives {
    $response = Invoke-GraphRequest -Uri "$GraphBaseUrl/sites/$($script:SiteId)/drives"
    return @($response.value)
}

function Set-Drive {
    param([string]$DriveId)
    $script:DriveId = $DriveId
}

# -----------------------------------------------
# FOLDER ENUMERATION (with paging support)
# -----------------------------------------------
function Get-FolderChildren {
    param([string]$ItemId = "root")
    $uri = "$GraphBaseUrl/drives/$($script:DriveId)/items/$ItemId/children" +
           "?`$select=id,name,folder,parentReference,lastModifiedDateTime&`$filter=folder ne null&`$top=200"
    $allItems = [System.Collections.Generic.List[object]]::new()

    while ($uri) {
        $response = Invoke-GraphRequest -Uri $uri
        if ($response.value) {
            foreach ($item in $response.value) {
                $allItems.Add($item)
            }
        }
        # Handle paging - Graph returns @odata.nextLink for large result sets
        # Must use PSObject.Properties check because Set-StrictMode -Latest
        # throws on accessing nonexistent properties directly
        if ($response.PSObject.Properties.Match('@odata.nextLink').Count -gt 0) {
            $uri = $response.'@odata.nextLink'
        } else {
            $uri = $null
        }
    }

    return @($allItems)
}

function Format-SPDate {
    param([string]$UtcString)
    if ([string]::IsNullOrWhiteSpace($UtcString)) { return "Unknown" }
    try {
        $utc   = [datetime]::Parse($UtcString, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
        $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, [System.TimeZoneInfo]::Local)
        return $local.ToString("M/d/yyyy h:mm tt")
    } catch {
        return $UtcString
    }
}

function Get-ItemPath {
    param([string]$ItemId)
    $item = Invoke-GraphRequest -Uri "$GraphBaseUrl/drives/$($script:DriveId)/items/$ItemId"
    return $item.parentReference.path + "/" + $item.name
}

# -----------------------------------------------
# PERMISSIONS
# -----------------------------------------------
function Get-FolderPermissions {
    param([string]$ItemId)
    try {
        $uri = "$GraphBaseUrl/drives/$($script:DriveId)/items/$ItemId/permissions"
        $response = Invoke-GraphRequest -Uri $uri
        return $response.value
    } catch {
        Write-Log "Could not read permissions for item $ItemId : $_" -Level WARN
        return @()
    }
}

function Test-HasUniquePermissions {
    param([object[]]$Permissions)
    # Inherited permissions have inheritedFrom set
    return ($Permissions | Where-Object { -not $_.inheritedFrom }) -ne $null
}

function Restore-FolderPermissions {
    param(
        [string]$DestItemId,
        [object[]]$Permissions
    )
    $uniquePerms = $Permissions | Where-Object { -not $_.inheritedFrom }
    foreach ($perm in $uniquePerms) {
        try {
            # Skip owner/sharing-link type permissions - only restore role assignments
            if ($perm.roles -and $perm.grantedTo) {
                $body = @{
                    roles      = $perm.roles
                    grantedTo  = $perm.grantedTo
                }
                Invoke-GraphRequest `
                    -Uri    "$GraphBaseUrl/drives/$($script:DriveId)/items/$DestItemId/permissions" `
                    -Method POST `
                    -Body   $body | Out-Null
                Write-Log "  Re-applied permission: $($perm.roles -join ',') -> $($perm.grantedTo.user.displayName)" -Level SUCCESS
            }
        } catch {
            Write-Log "  Failed to re-apply permission ($($perm.id)): $_" -Level WARN
        }
    }
}

# -----------------------------------------------
# FOLDER MOVE
# -----------------------------------------------
function Move-SPFolder {
    param(
        [string]$SourceItemId,
        [string]$SourceDisplayPath,
        [string]$DestParentItemId,
        [string]$DestDisplayPath
    )

    Write-Log "Reading permissions for: $SourceDisplayPath"
    $perms = Get-FolderPermissions -ItemId $SourceItemId
    $hasUnique = Test-HasUniquePermissions -Permissions $perms

    if ($DryRun) {
        $permNote = if ($hasUnique) { "has unique permissions - will be re-applied" } else { "inherits from parent" }
        Write-Log "[DRY-RUN] Would move: $SourceDisplayPath -> $DestDisplayPath ($permNote)" -Level "DRY-RUN"
        Add-LogEntry -SourcePath $SourceDisplayPath -DestPath $DestDisplayPath `
                     -Status "DRY-RUN" -Detail $permNote
        return
    }

    Write-Log "Moving: $SourceDisplayPath -> $DestDisplayPath"
    $body = @{
        parentReference = @{ id = $DestParentItemId }
    }

    try {
        Invoke-GraphRequest `
            -Uri    "$GraphBaseUrl/drives/$($script:DriveId)/items/$SourceItemId" `
            -Method PATCH `
            -Body   $body | Out-Null

        Write-Log "Move complete: $SourceDisplayPath" -Level SUCCESS

        if ($hasUnique) {
            Write-Log "Restoring unique permissions on moved folder..." -Level INFO
            # After move, get new item ID at destination to apply perms
            $movedItem = Invoke-GraphRequest -Uri "$GraphBaseUrl/drives/$($script:DriveId)/items/$SourceItemId"
            Restore-FolderPermissions -DestItemId $movedItem.id -Permissions $perms
        }

        Add-LogEntry -SourcePath $SourceDisplayPath -DestPath $DestDisplayPath `
                     -Status "SUCCESS" -Detail $(if ($hasUnique) { "permissions restored" } else { "inherited" })
    } catch {
        Write-Log "Move FAILED: $SourceDisplayPath - $_" -Level ERROR
        Add-LogEntry -SourcePath $SourceDisplayPath -DestPath $DestDisplayPath `
                     -Status "FAILED" -Detail "$_"
    }
}



# -----------------------------------------------
# WINFORMS HELPERS
# -----------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Global sort mode: "date" (oldest first) or "name" (A-Z)
$script:SortMode = "date"

# Sort helper: returns sorted array based on current $script:SortMode
function Sort-FolderItems {
    param([object[]]$Items)
    if ($script:SortMode -eq "name") {
        return @($Items | Sort-Object name)
    } else {
        return @($Items | Sort-Object {
            if ($_.lastModifiedDateTime) { [datetime]::Parse($_.lastModifiedDateTime) }
            else { [datetime]::MinValue }
        })
    }
}

# Collect all checked nodes iteratively
function Get-CheckedTags {
    param([System.Windows.Forms.TreeNodeCollection]$RootNodes)
    $result = [System.Collections.Generic.List[hashtable]]::new()
    $stack  = [System.Collections.Generic.Stack[System.Windows.Forms.TreeNode]]::new()
    foreach ($n in $RootNodes) { $stack.Push($n) }
    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        if ($node.Checked -and $node.Tag -is [hashtable]) {
            $result.Add($node.Tag)
        }
        foreach ($child in $node.Nodes) { $stack.Push($child) }
    }
    return ,$result
}

function New-FolderNode {
    param([object]$Item)
    $dateStr = Format-SPDate -UtcString $Item.lastModifiedDateTime
    $displayName = $Item.name
    if ($displayName.Length -gt 60) {
        $displayName = $displayName.Substring(0, 57) + "..."
    }
    $label = "$displayName  [$dateStr]"
    $node  = New-Object System.Windows.Forms.TreeNode($label)
    $node.Tag = @{
        id           = $Item.id
        name         = $Item.name
        lastModified = $dateStr
        rawDateTime  = $Item.lastModifiedDateTime
    }
    $node.Nodes.Add("__loading__") | Out-Null
    return $node
}

function Expand-TreeNode {
    param([System.Windows.Forms.TreeNode]$Node, [System.Windows.Forms.Label]$StatusLabel)
    if ($Node.Nodes.Count -eq 1 -and $Node.Nodes[0].Text -eq "__loading__") {
        $Node.Nodes.Clear()
        $StatusLabel.Text = "Loading subfolders..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $children = Get-FolderChildren -ItemId $Node.Tag.id
            $sorted   = Sort-FolderItems -Items $children
            foreach ($child in $sorted) {
                $Node.Nodes.Add((New-FolderNode -Item $child)) | Out-Null
            }
            if ($children.Count -eq 0) {
                $StatusLabel.Text = "No subfolders found."
            } else {
                $StatusLabel.Text = "$($children.Count) subfolder(s) loaded."
            }
        } catch {
            $StatusLabel.Text = "Error loading subfolders: $_"
            Write-Log "Error expanding node '$($Node.Tag.name)': $_" -Level ERROR
        }
    }
}

function Load-RootNodes {
    param(
        [System.Windows.Forms.TreeNodeCollection]$NodeCollection,
        [System.Windows.Forms.Label]$StatusLabel
    )
    $NodeCollection.Clear()
    $StatusLabel.Text = "Loading root folders..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $children = Get-FolderChildren -ItemId "root"
        $sorted   = Sort-FolderItems -Items $children
        foreach ($child in $sorted) {
            $NodeCollection.Add((New-FolderNode -Item $child)) | Out-Null
        }
        $sortLabel = if ($script:SortMode -eq "date") { "sorted oldest first" } else { "sorted A-Z" }
        $StatusLabel.Text = "$($children.Count) folder(s) - $sortLabel"
    } catch {
        $StatusLabel.Text = "Error loading folders: $_"
        Write-Log "Error loading root nodes: $_" -Level ERROR
    }
}

# Creates clickable column-header style panel for sorting
function New-SortHeaderPanel {
    param(
        [System.Windows.Forms.TreeView]$Tree,
        [System.Windows.Forms.Label]$StatusLabel
    )
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock      = "Top"
    $headerPanel.Height    = 24
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)

    $nameHeader = New-Object System.Windows.Forms.Label
    $nameHeader.Text      = "  Name"
    $nameHeader.Location  = New-Object System.Drawing.Point(0, 0)
    $nameHeader.Size      = New-Object System.Drawing.Size(300, 24)
    $nameHeader.TextAlign = "MiddleLeft"
    $nameHeader.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $nameHeader.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $headerPanel.Controls.Add($nameHeader)

    $dateHeader = New-Object System.Windows.Forms.Label
    $dateHeader.Text      = "Date Modified  v"
    $dateHeader.Location  = New-Object System.Drawing.Point(300, 0)
    $dateHeader.Size      = New-Object System.Drawing.Size(300, 24)
    $dateHeader.TextAlign = "MiddleLeft"
    $dateHeader.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $dateHeader.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $headerPanel.Controls.Add($dateHeader)

    $nameHeader.Add_Click({
        $script:SortMode = "name"
        $nameHeader.Text = "  Name  v"
        $nameHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $dateHeader.Text = "Date Modified"
        $dateHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
        Load-RootNodes -NodeCollection $Tree.Nodes -StatusLabel $StatusLabel
    })

    $dateHeader.Add_Click({
        $script:SortMode = "date"
        $nameHeader.Text = "  Name"
        $nameHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
        $dateHeader.Text = "Date Modified  v"
        $dateHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        Load-RootNodes -NodeCollection $Tree.Nodes -StatusLabel $StatusLabel
    })

    return $headerPanel
}

# -----------------------------------------------
# STEP 1 - SOURCE PICKER
# -----------------------------------------------
function Show-SourcePicker {
    param([object[]]$Drives)

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "SPFolderMover - Step 1 of 2: Select folders to move"
    $form.Width           = 950
    $form.Height          = 700
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MinimumSize     = New-Object System.Drawing.Size(750, 500)
    $form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

    # -- Bottom panel with FlowLayoutPanel for buttons --
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock   = "Bottom"
    $bottomPanel.Height = 48

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Dock      = "Fill"
    $statusLabel.Padding   = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $statusLabel.TextAlign = "MiddleLeft"
    $statusLabel.ForeColor = [System.Drawing.Color]::Gray
    $statusLabel.Text      = "Loading..."

    # FlowLayoutPanel docked right - buttons auto-position from right edge
    $btnFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnFlow.Dock          = "Right"
    $btnFlow.FlowDirection = "LeftToRight"
    $btnFlow.AutoSize      = $true
    $btnFlow.WrapContents  = $false
    $btnFlow.Padding       = New-Object System.Windows.Forms.Padding(0, 8, 10, 0)

    $btnMoveTo = New-Object System.Windows.Forms.Button
    $btnMoveTo.Text      = "Move To  >"
    $btnMoveTo.Size      = New-Object System.Drawing.Size(115, 30)
    $btnMoveTo.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 184)
    $btnMoveTo.ForeColor = [System.Drawing.Color]::White
    $btnMoveTo.FlatStyle = "Flat"
    $btnMoveTo.Margin    = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
    $btnMoveTo.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton   = $btnMoveTo

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text   = "Cancel"
    $btnCancel.Size   = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel

    $btnFlow.Controls.Add($btnMoveTo)
    $btnFlow.Controls.Add($btnCancel)
    $bottomPanel.Controls.Add($btnFlow)
    $bottomPanel.Controls.Add($statusLabel)

    # -- Instruction bar --
    $instrPanel = New-Object System.Windows.Forms.Panel
    $instrPanel.Dock      = "Top"
    $instrPanel.Height    = 34
    $instrPanel.BackColor = [System.Drawing.Color]::FromArgb(232, 243, 255)

    $instrLabel = New-Object System.Windows.Forms.Label
    $instrLabel.Text      = "  Check the folders you want to move, then click  Move To.   (Shift+Click to check a range)"
    $instrLabel.Dock      = "Fill"
    $instrLabel.TextAlign = "MiddleLeft"
    $instrLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $instrPanel.Controls.Add($instrLabel)

    # -- Drive selector --
    $drivePanel = New-Object System.Windows.Forms.Panel
    $drivePanel.Dock   = "Top"
    $drivePanel.Height = 36

    $driveLabel = New-Object System.Windows.Forms.Label
    $driveLabel.Text     = "Library:"
    $driveLabel.Location = New-Object System.Drawing.Point(10, 9)
    $driveLabel.Size     = New-Object System.Drawing.Size(55, 20)
    $drivePanel.Controls.Add($driveLabel)

    $driveCombo = New-Object System.Windows.Forms.ComboBox
    $driveCombo.Location      = New-Object System.Drawing.Point(68, 7)
    $driveCombo.Size          = New-Object System.Drawing.Size(350, 24)
    $driveCombo.Anchor        = "Left,Right,Top"
    $driveCombo.DropDownStyle = "DropDownList"
    foreach ($d in $Drives) { $driveCombo.Items.Add($d.name) | Out-Null }
    if ($driveCombo.Items.Count -gt 0) { $driveCombo.SelectedIndex = 0 }
    $drivePanel.Controls.Add($driveCombo)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text      = "0 folders checked"
    $countLabel.Location  = New-Object System.Drawing.Point(440, 9)
    $countLabel.Size      = New-Object System.Drawing.Size(200, 20)
    $countLabel.Anchor    = "Right,Top"
    $countLabel.ForeColor = [System.Drawing.Color]::Gray
    $countLabel.TextAlign = "MiddleRight"
    $drivePanel.Controls.Add($countLabel)

    # -- Tree --
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock          = "Fill"
    $tree.CheckBoxes    = $true
    $tree.HideSelection = $false
    $tree.Scrollable    = $true
    $tree.Font          = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $tree.ItemHeight    = 22

    # -- Sort column header --
    $sortHeader = New-SortHeaderPanel -Tree $tree -StatusLabel $statusLabel

    # -- Add to form: Fill first, edges next, Top panels last (last=topmost) --
    $form.Controls.Add($tree)
    $form.Controls.Add($sortHeader)
    $form.Controls.Add($bottomPanel)
    $form.Controls.Add($drivePanel)
    $form.Controls.Add($instrPanel)

    # -- Shift-click range selection --
    $script:lastClickedNode = $null

    $tree.Add_NodeMouseClick({
        param($s, $e)
        $clickedNode = $e.Node
        $shiftHeld   = [System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift

        if ($shiftHeld -and $script:lastClickedNode -ne $null) {
            $anchor  = $script:lastClickedNode
            $target  = $clickedNode
            $anchorParent = $anchor.Parent
            $targetParent = $target.Parent

            if (($anchorParent -eq $null -and $targetParent -eq $null) -or ($anchorParent -eq $targetParent)) {
                $siblings = if ($anchorParent -eq $null) { $tree.Nodes } else { $anchorParent.Nodes }
                $idx1 = $siblings.IndexOf($anchor)
                $idx2 = $siblings.IndexOf($target)
                if ($idx1 -ge 0 -and $idx2 -ge 0) {
                    $lo = [Math]::Min($idx1, $idx2)
                    $hi = [Math]::Max($idx1, $idx2)
                    $newState = $anchor.Checked
                    for ($i = $lo; $i -le $hi; $i++) {
                        $siblings[$i].Checked = $newState
                    }
                }
            }
        }
        $script:lastClickedNode = $clickedNode
    })

    $tree.Add_BeforeExpand({
        param($s, $e)
        Expand-TreeNode -Node $e.Node -StatusLabel $statusLabel
    })

    $tree.Add_AfterCheck({
        param($s, $e)
        $checked = Get-CheckedTags -RootNodes $tree.Nodes
        $n = $checked.Count
        $countLabel.Text = "$n folder(s) checked"
        $countLabel.ForeColor = if ($n -gt 0) {
            [System.Drawing.Color]::FromArgb(0, 103, 184)
        } else {
            [System.Drawing.Color]::Gray
        }
    })

    $driveCombo.Add_SelectedIndexChanged({
        Set-Drive -DriveId $Drives[$driveCombo.SelectedIndex].id
        Load-RootNodes -NodeCollection $tree.Nodes -StatusLabel $statusLabel
        $countLabel.Text      = "0 folders checked"
        $countLabel.ForeColor = [System.Drawing.Color]::Gray
        $script:lastClickedNode = $null
        $form.Refresh()
    })

    Set-Drive -DriveId $Drives[0].id
    Load-RootNodes -NodeCollection $tree.Nodes -StatusLabel $statusLabel
    $form.Refresh()

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        $form.Dispose()
        return $null
    }

    $checked = Get-CheckedTags -RootNodes $tree.Nodes
    $form.Dispose()

    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No folders were checked. Please check at least one folder to move.",
            "Nothing selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $null
    }
    return ,$checked
}

# -----------------------------------------------
# STEP 2 - DESTINATION PICKER + PROGRESS
# -----------------------------------------------
function Show-DestAndExecute {
    param(
        [object[]]$Sources,
        [object[]]$Drives
    )

    if ($Sources -is [hashtable]) {
        $Sources = @($Sources)
    }
    $sourceCount = $Sources.Count

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "SPFolderMover - Step 2 of 2: Choose destination and begin move"
    $form.Width           = 950
    $form.Height          = 750
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MinimumSize     = New-Object System.Drawing.Size(750, 580)
    $form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

    # -- Bottom panel with FlowLayoutPanel for buttons --
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock   = "Bottom"
    $bottomPanel.Height = 60

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Dock    = "Top"
    $progressBar.Height  = 18
    $progressBar.Minimum = 0
    $progressBar.Maximum = [Math]::Max($sourceCount, 1)
    $progressBar.Value   = 0
    $progressBar.Style   = "Continuous"
    $bottomPanel.Controls.Add($progressBar)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Dock      = "Fill"
    $statusLabel.Padding   = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $statusLabel.TextAlign = "MiddleLeft"
    $statusLabel.ForeColor = [System.Drawing.Color]::Gray
    $statusLabel.Text      = "Select a destination folder, then click Begin Move."

    $btnFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnFlow.Dock          = "Right"
    $btnFlow.FlowDirection = "LeftToRight"
    $btnFlow.AutoSize      = $true
    $btnFlow.WrapContents  = $false
    $btnFlow.Padding       = New-Object System.Windows.Forms.Padding(0, 4, 10, 0)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text   = "< Back"
    $btnBack.Size   = New-Object System.Drawing.Size(80, 28)
    $btnBack.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
    $btnBack.DialogResult = [System.Windows.Forms.DialogResult]::Retry

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text   = "Cancel"
    $btnCancel.Size   = New-Object System.Drawing.Size(80, 28)
    $btnCancel.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel

    $btnBegin = New-Object System.Windows.Forms.Button
    $btnBegin.Text      = "Begin Move"
    $btnBegin.Size      = New-Object System.Drawing.Size(115, 28)
    $btnBegin.Margin    = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
    $btnBegin.BackColor = [System.Drawing.Color]::FromArgb(0, 140, 60)
    $btnBegin.ForeColor = [System.Drawing.Color]::White
    $btnBegin.FlatStyle = "Flat"
    $btnBegin.Enabled   = $false

    $btnFlow.Controls.Add($btnBack)
    $btnFlow.Controls.Add($btnCancel)
    $btnFlow.Controls.Add($btnBegin)
    $bottomPanel.Controls.Add($btnFlow)
    $bottomPanel.Controls.Add($statusLabel)

    # -- Instruction bar --
    $instrPanel = New-Object System.Windows.Forms.Panel
    $instrPanel.Dock      = "Top"
    $instrPanel.Height    = 34
    $instrPanel.BackColor = [System.Drawing.Color]::FromArgb(232, 243, 255)

    $instrLabel = New-Object System.Windows.Forms.Label
    $instrLabel.Text      = "  Moving $sourceCount folder(s). Navigate to your destination folder, select it, then click Begin Move."
    $instrLabel.Dock      = "Fill"
    $instrLabel.TextAlign = "MiddleLeft"
    $instrLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $instrPanel.Controls.Add($instrLabel)

    # -- Source summary panel --
    $srcPanel = New-Object System.Windows.Forms.Panel
    $srcPanel.Dock      = "Top"
    $srcPanel.Height    = 28
    $srcPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 230)

    $srcNames = ($Sources | ForEach-Object { $_.name }) -join ", "
    if ($srcNames.Length -gt 120) {
        $srcNames = $srcNames.Substring(0, 117) + "..."
    }
    $srcLabel = New-Object System.Windows.Forms.Label
    $srcLabel.Text      = "  Sources: $srcNames"
    $srcLabel.Dock      = "Fill"
    $srcLabel.TextAlign = "MiddleLeft"
    $srcLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $srcLabel.ForeColor = [System.Drawing.Color]::FromArgb(130, 100, 0)
    $srcPanel.Controls.Add($srcLabel)

    # -- Drive selector --
    $drivePanel = New-Object System.Windows.Forms.Panel
    $drivePanel.Dock   = "Top"
    $drivePanel.Height = 36

    $driveLabel = New-Object System.Windows.Forms.Label
    $driveLabel.Text     = "Library:"
    $driveLabel.Location = New-Object System.Drawing.Point(10, 9)
    $driveLabel.Size     = New-Object System.Drawing.Size(55, 20)
    $drivePanel.Controls.Add($driveLabel)

    $driveCombo = New-Object System.Windows.Forms.ComboBox
    $driveCombo.Location      = New-Object System.Drawing.Point(68, 7)
    $driveCombo.Size          = New-Object System.Drawing.Size(350, 24)
    $driveCombo.Anchor        = "Left,Right,Top"
    $driveCombo.DropDownStyle = "DropDownList"
    foreach ($d in $Drives) { $driveCombo.Items.Add($d.name) | Out-Null }
    if ($driveCombo.Items.Count -gt 0) { $driveCombo.SelectedIndex = 0 }
    $drivePanel.Controls.Add($driveCombo)

    # -- Selected destination label --
    $selPanel = New-Object System.Windows.Forms.Panel
    $selPanel.Dock      = "Top"
    $selPanel.Height    = 26
    $selPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    $selLabel = New-Object System.Windows.Forms.Label
    $selLabel.Text      = "  Destination: (none selected)"
    $selLabel.Dock      = "Fill"
    $selLabel.TextAlign = "MiddleLeft"
    $selLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $selLabel.ForeColor = [System.Drawing.Color]::Gray
    $selPanel.Controls.Add($selLabel)

    # -- SplitContainer (tree + log) --
    $splitter = New-Object System.Windows.Forms.SplitContainer
    $splitter.Dock              = "Fill"
    $splitter.Orientation       = "Horizontal"
    $splitter.SplitterDistance   = 320
    $splitter.SplitterWidth     = 6
    $splitter.Panel1MinSize     = 150
    $splitter.Panel2MinSize     = 120

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock          = "Fill"
    $tree.HideSelection = $false
    $tree.Scrollable    = $true
    $tree.Font          = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $tree.ItemHeight    = 22

    $sortHeader2 = New-SortHeaderPanel -Tree $tree -StatusLabel $statusLabel

    $splitter.Panel1.Controls.Add($tree)
    $splitter.Panel1.Controls.Add($sortHeader2)

    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.Dock       = "Fill"
    $logBox.Font       = New-Object System.Drawing.Font("Consolas", 8.5)
    $logBox.ReadOnly   = $true
    $logBox.BackColor  = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $logBox.WordWrap   = $true
    $logBox.ScrollBars = "ForcedBoth"
    $splitter.Panel2.Controls.Add($logBox)

    # -- Add to form: Fill first, edges next, Top panels last (last=topmost) --
    $form.Controls.Add($splitter)
    $form.Controls.Add($bottomPanel)
    $form.Controls.Add($selPanel)
    $form.Controls.Add($drivePanel)
    $form.Controls.Add($srcPanel)
    $form.Controls.Add($instrPanel)

    # -- Event handlers --
    $tree.Add_AfterSelect({
        param($s, $e)
        if ($e.Node -and $e.Node.Tag -is [hashtable]) {
            $selLabel.Text      = "  Destination: $($e.Node.Tag.name)"
            $selLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 80, 0)
            $selLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $btnBegin.Enabled   = $true
        }
    })

    $tree.Add_BeforeExpand({
        param($s, $e)
        Expand-TreeNode -Node $e.Node -StatusLabel $statusLabel
    })

    $driveCombo.Add_SelectedIndexChanged({
        Set-Drive -DriveId $Drives[$driveCombo.SelectedIndex].id
        Load-RootNodes -NodeCollection $tree.Nodes -StatusLabel $statusLabel
        $selLabel.Text      = "  Destination: (none selected)"
        $selLabel.ForeColor = [System.Drawing.Color]::Gray
        $selLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
        $btnBegin.Enabled   = $false
        $form.Refresh()
    })

    function Append-Log {
        param([string]$Msg, [string]$Color = "Black")
        $logBox.SelectionStart  = $logBox.TextLength
        $logBox.SelectionLength = 0
        $logBox.SelectionColor  = [System.Drawing.Color]::FromName($Color)
        $logBox.AppendText("$Msg`r`n")
        $logBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Begin Move click handler
    $btnBegin.Add_Click({
        try {
            $destTag = if ($tree.SelectedNode -and $tree.SelectedNode.Tag -is [hashtable]) {
                $tree.SelectedNode.Tag
            } else { $null }

            if (-not $destTag) {
                [System.Windows.Forms.MessageBox]::Show("Please select a destination folder first.",
                    "No destination", "OK", "Warning") | Out-Null
                return
            }

            $dryTag = if ($DryRun) { " (DRY-RUN)" } else { "" }
            $confirmMsg = "Move $sourceCount folder(s) to '$($destTag.name)'?$dryTag`n`nThis operation will:`n"
            if ($DryRun) {
                $confirmMsg += "- Preview all moves (no changes will be made)"
            } else {
                $confirmMsg += "- Move the selected folders to the destination`n- Preserve unique permissions where applicable"
            }
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                $confirmMsg,
                "Confirm Move$dryTag",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }

            $btnBegin.Enabled   = $false
            $btnBack.Enabled    = $false
            $btnCancel.Enabled  = $false
            $tree.Enabled       = $false
            $driveCombo.Enabled = $false

            Refresh-GraphToken

            $dryLabel = if ($DryRun) { "[DRY-RUN] " } else { "" }
            Append-Log "$(Get-Date -f 'HH:mm:ss') ${dryLabel}Starting - $sourceCount folder(s) -> '$($destTag.name)'" "DarkBlue"
            Append-Log "------------------------------------------------------------"

            $successCount = 0
            $failCount    = 0
            $skipCount    = 0

            for ($i = 0; $i -lt $sourceCount; $i++) {
                $src = $Sources[$i]

                if ($src.id -eq $destTag.id) {
                    Append-Log "$(Get-Date -f 'HH:mm:ss') [SKIP]  '$($src.name)' - same as destination" "DarkGoldenrod"
                    $skipCount++
                    $progressBar.Value = ($i + 1)
                    [System.Windows.Forms.Application]::DoEvents()
                    continue
                }

                $destIsChild = $false
                try {
                    $destItem = Invoke-GraphRequest -Uri "$GraphBaseUrl/drives/$($script:DriveId)/items/$($destTag.id)?`$select=id,parentReference"
                    if ($destItem.parentReference -and $destItem.parentReference.path -match $src.id) {
                        $destIsChild = $true
                    }
                } catch {}
                if ($destIsChild) {
                    Append-Log "$(Get-Date -f 'HH:mm:ss') [SKIP]  '$($src.name)' - cannot move into own subfolder" "DarkGoldenrod"
                    $skipCount++
                    $progressBar.Value = ($i + 1)
                    [System.Windows.Forms.Application]::DoEvents()
                    continue
                }

                $statusLabel.Text = "Moving ($($i + 1) / $sourceCount): $($src.name)"
                Append-Log "$(Get-Date -f 'HH:mm:ss') [....] Moving '$($src.name)'..."
                [System.Windows.Forms.Application]::DoEvents()

                try {
                    $perms     = Get-FolderPermissions -ItemId $src.id
                    $hasUnique = Test-HasUniquePermissions -Permissions $perms

                    if (-not $DryRun) {
                        $body = @{ parentReference = @{ id = $destTag.id } }
                        Invoke-GraphRequest `
                            -Uri    "$GraphBaseUrl/drives/$($script:DriveId)/items/$($src.id)" `
                            -Method PATCH `
                            -Body   $body | Out-Null

                        if ($hasUnique) {
                            $movedItem = Invoke-GraphRequest -Uri "$GraphBaseUrl/drives/$($script:DriveId)/items/$($src.id)"
                            Restore-FolderPermissions -DestItemId $movedItem.id -Permissions $perms
                            Append-Log "$(Get-Date -f 'HH:mm:ss') [OK]   '$($src.name)' - moved + permissions restored" "DarkGreen"
                        } else {
                            Append-Log "$(Get-Date -f 'HH:mm:ss') [OK]   '$($src.name)' - moved (inherits permissions)" "DarkGreen"
                        }
                        Add-LogEntry -SourcePath $src.name -DestPath $destTag.name `
                                     -Status "SUCCESS" -Detail $(if ($hasUnique) { "perms restored" } else { "inherited" })
                    } else {
                        $permNote = if ($hasUnique) { "has unique perms" } else { "inherits perms" }
                        Append-Log "$(Get-Date -f 'HH:mm:ss') [DRY]  '$($src.name)' would move -> '$($destTag.name)' ($permNote)" "DarkCyan"
                        Add-LogEntry -SourcePath $src.name -DestPath $destTag.name -Status "DRY-RUN" -Detail $permNote
                    }
                    $successCount++

                } catch {
                    Append-Log "$(Get-Date -f 'HH:mm:ss') [FAIL] '$($src.name)' - $_" "Red"
                    Add-LogEntry -SourcePath $src.name -DestPath $destTag.name -Status "FAILED" -Detail "$_"
                    $failCount++
                }

                $progressBar.Value = ($i + 1)
                [System.Windows.Forms.Application]::DoEvents()
            }

            Append-Log "------------------------------------------------------------"
            if ($DryRun) {
                Append-Log "$(Get-Date -f 'HH:mm:ss') DRY-RUN complete - $successCount previewed, $skipCount skipped, $failCount errors." "DarkCyan"
            } else {
                $summaryColor = if ($failCount -gt 0) { "DarkRed" } else { "DarkGreen" }
                Append-Log "$(Get-Date -f 'HH:mm:ss') Done - $successCount succeeded, $skipCount skipped, $failCount failed." $summaryColor
            }

            Save-Log
            Append-Log "$(Get-Date -f 'HH:mm:ss') Log saved: $LogPath"

            $statusLabel.Text = if ($DryRun) {
                "Dry-run complete. No changes made. ($successCount previewed)"
            } else {
                "Done - $successCount succeeded, $skipCount skipped, $failCount failed."
            }
            $statusLabel.ForeColor = if ($failCount -gt 0) {
                [System.Drawing.Color]::DarkRed
            } else {
                [System.Drawing.Color]::DarkGreen
            }

            $progressBar.Value = $progressBar.Maximum
            $btnCancel.Text    = "Close"
            $btnCancel.Enabled = $true

        } catch {
            $errMsg = "Unexpected error during move: $_"
            try { Append-Log "$(Get-Date -f 'HH:mm:ss') [FATAL] $errMsg" "Red" } catch {}
            Write-Log $errMsg -Level ERROR
            Save-Log

            $statusLabel.Text      = "Error occurred. See log for details."
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkRed

            $btnCancel.Text    = "Close"
            $btnCancel.Enabled = $true

            [System.Windows.Forms.MessageBox]::Show(
                "An error occurred:`n$_`n`nPartial results may have been applied. Check the log for details.",
                "Move Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })

    Set-Drive -DriveId $Drives[0].id
    Load-RootNodes -NodeCollection $tree.Nodes -StatusLabel $statusLabel
    $form.Refresh()

    return $form.ShowDialog()
}

# -----------------------------------------------
# MAIN
# -----------------------------------------------
function Main {
    Write-Log "=== SPFolderMover.ps1 starting ===" -Level INFO
    if ($DryRun) { Write-Log "DRY-RUN mode enabled - no changes will be made." -Level "DRY-RUN" }

    Assert-Module -Name "MSAL.PS"
    Get-GraphToken
    Resolve-Site

    Write-Log "Enumerating document libraries..."
    $drives = Get-Drives
    if (-not $drives -or $drives.Count -eq 0) {
        Write-Log "No document libraries found on site. Exiting." -Level ERROR
        return
    }
    Write-Log "Found $($drives.Count) library/libraries." -Level SUCCESS

    while ($true) {
        $sourceFolders = Show-SourcePicker -Drives $drives
        if (-not $sourceFolders) {
            Write-Log "Cancelled at source picker." -Level WARN
            return
        }
        Write-Log "Selected $($sourceFolders.Count) source folder(s)."

        Refresh-GraphToken

        $result = Show-DestAndExecute -Sources $sourceFolders -Drives $drives
        if ($result -eq [System.Windows.Forms.DialogResult]::Retry) {
            continue
        }
        break
    }
}

# -----------------------------------------------
# ENTRY POINT
# -----------------------------------------------
try {
    Main
} catch {
    Write-Log "Fatal error: $_" -Level ERROR
    Save-Log
    [System.Windows.Forms.MessageBox]::Show(
        "A fatal error occurred:`n$_`n`nCheck the console for details.",
        "SPFolderMover - Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}
