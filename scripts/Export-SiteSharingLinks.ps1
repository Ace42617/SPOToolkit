<#
.SYNOPSIS
    Scans a SharePoint site and exports all unique sharing links from every list and library to a CSV.

.DESCRIPTION
    Connects to a site with PnP PowerShell, enumerates all non-hidden, non-system lists and libraries,
    and inspects only items that have unique role assignments (the only items that can carry sharing
    links). Items are pulled in bulk via the SharePoint REST endpoint with
    `$select=HasUniqueRoleAssignments,...` so the unique-permission flag arrives in the same payload
    (no per-item round-trip). For each unique-permission item it pulls sharing links via
    Get-PnPFileSharingLink / Get-PnPFolderSharingLink. Items with sharing links are exported as
    SharingLink rows; items with unique permissions but no sharing link are exported as DirectPermission
    rows with role-assignment details (who has access). Results are deduplicated by link URL where
    applicable and written to a CSV next to this script. Progress is reported against the total item
    count across all lists and libraries.

.PARAMETER SiteUrl
    Full SharePoint site URL, e.g. https://contoso.sharepoint.com/sites/HR.

.PARAMETER ClientId
    Entra (Azure AD) app registration Client ID used by PnP PowerShell for interactive sign-in.

.EXAMPLE
    .\Export-SiteSharingLinks.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/HR" -ClientId "00000000-0000-0000-0000-000000000000"

.NOTES
    Requires: PnP.PowerShell
    Install : Install-Module PnP.PowerShell -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [string] $ClientId
)

$ErrorActionPreference = 'Stop'

# System / template lists that never carry meaningful sharing links - skipped for efficiency.
$ExcludedListTitles = @(
    'Access Requests', 'App Packages', 'appdata', 'appfiles', 'Apps in Testing',
    'Cache Profiles', 'Composed Looks', 'Content and Structure Reports',
    'Content type publishing error log', 'Converted Forms', 'Device Channels',
    'Form Templates', 'fpdatasources', 'Get started with Apps for Office and SharePoint',
    'List Template Gallery', 'Long Running Operation Status', 'Maintenance Log Library',
    'Master Docs', 'Master Page Gallery', 'MicroFeed', 'NintexFormXml',
    'Quick Deploy Items', 'Relationships List', 'Reusable Content',
    'Reporting Metadata', 'Reporting Templates', 'Search Config List',
    'Preservation Hold Library', 'Solution Gallery',
    'Suggested Content Browser Locations', 'Theme Gallery', 'TaxonomyHiddenList',
    'User Information List', 'Web Part Gallery', 'wfpub', 'wfsvc',
    'Workflow History', 'Workflow Tasks', 'Style Library'
)

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser'
}
Import-Module PnP.PowerShell -ErrorAction Stop

# Resolve the script's own directory robustly so the CSV always lands next to it.
$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue
}
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# Enumerate lists/libraries we care about and compute total item count for the progress bar.
$lists = Get-PnPList -Includes BaseType, Hidden, Title, ItemCount, RootFolder |
    Where-Object {
        -not $_.Hidden -and
        $ExcludedListTitles -notcontains $_.Title -and
        ($_.BaseType -eq 'DocumentLibrary' -or $_.BaseType -eq 'GenericList')
    }

$totalItems = ($lists | Measure-Object -Property ItemCount -Sum).Sum
if (-not $totalItems) { $totalItems = 0 }

Write-Host ("Found {0} list(s)/library(ies), {1:N0} total item(s)." -f @($lists).Count, $totalItems) -ForegroundColor Cyan

# Keyed on share link URL so we only keep one row per unique link.
$uniqueLinks          = [System.Collections.Generic.Dictionary[string, object]]::new()
$directPermissionRows = [System.Collections.Generic.List[object]]::new()
$processed            = 0
$uniqueCount          = 0
$linksFound           = 0

function Get-PrincipalTypeName {
    param([int] $PrincipalType)

    switch ($PrincipalType) {
        1 { return 'User' }
        2 { return 'DistributionList' }
        4 { return 'SecurityGroup' }
        8 { return 'SharePointGroup' }
        default { return "Type$PrincipalType" }
    }
}

function Format-AccessEntry {
    param(
        [string] $PrincipalType,
        [string] $Title,
        [string] $Login,
        [string] $Email,
        [string[]] $Roles
    )

    $typeLabel = switch ($PrincipalType) {
        'User'            { 'User' }
        'SharePointGroup' { 'Group' }
        'SecurityGroup'   { 'Security Group' }
        'DistributionList'{ 'Distribution List' }
        default           { $PrincipalType }
    }

    $roleText = if ($Roles.Count -gt 0) { ($Roles -join ', ') } else { '(no roles)' }

    if ($PrincipalType -eq 'User') {
        $identity = if (-not [string]::IsNullOrWhiteSpace($Email)) {
            if (-not [string]::IsNullOrWhiteSpace($Title) -and $Title -ne $Email) {
                "{0} <{1}>" -f $Title, $Email
            }
            else {
                $Email
            }
        }
        elseif (-not [string]::IsNullOrWhiteSpace($Title)) { $Title }
        else { $Login }

        return ("{0}: {1} [{2}]" -f $typeLabel, $identity, $roleText)
    }

    $name = if (-not [string]::IsNullOrWhiteSpace($Title)) { $Title } else { $Login }
    return ("{0}: {1} [{2}]" -f $typeLabel, $name, $roleText)
}

function Get-ItemAccessInfo {
    param(
        [string] $ListId,
        [int]    $ItemId
    )

    $url = "/_api/web/lists(guid'$ListId')/items($ItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings"

    try {
        $response = Invoke-PnPSPRestMethod -Url $url -Method Get
    }
    catch {
        Write-Warning ("Role-assignment API error for item $ItemId : $($_.Exception.Message)")
        return [pscustomobject]@{ Access = '' }
    }

    $entries = [System.Collections.Generic.List[string]]::new()

    foreach ($assignment in @($response.value)) {
        $member = $assignment.Member
        if (-not $member) { continue }

        $roles = @()
        if ($assignment.RoleDefinitionBindings) {
            if ($assignment.RoleDefinitionBindings.results) {
                $roles = @($assignment.RoleDefinitionBindings.results | ForEach-Object { $_.Name })
            }
            else {
                $roles = @($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name })
            }
        }
        $roles = @($roles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

        $entries.Add((Format-AccessEntry `
            -PrincipalType (Get-PrincipalTypeName -PrincipalType ([int]$member.PrincipalType)) `
            -Title ([string]$member.Title) `
            -Login ([string]$member.LoginName) `
            -Email ([string]$member.Email) `
            -Roles $roles)) | Out-Null
    }

    [pscustomobject]@{
        Access = ($entries -join '; ')
    }
}

function Write-ScanProgress {
    param(
        [int]    $Processed,
        [int]    $Total,
        [string] $Status
    )
    $percent = if ($Total -gt 0) { [int](($Processed / $Total) * 100) } else { 0 }
    Write-Progress -Activity 'Scanning items for sharing links' `
        -Status $Status `
        -PercentComplete ([math]::Min($percent, 100))
}

function Add-SharingLinkRow {
    param(
        $Link,
        [string] $LinkUrl,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ObjectType,
        [string] $ItemName,
        [string] $RelativeUrl,
        [int]    $ItemId
    )

    $userEmails = @()
    if ($Link.GrantedToIdentitiesV2) {
        foreach ($id in $Link.GrantedToIdentitiesV2) {
            if ($id.User -and $id.User.Email) { $userEmails += $id.User.Email }
        }
    }

    $blocksDownload = $Link.Link.PreventsDownload
    if ($null -eq $blocksDownload) { $blocksDownload = $Link.Link.PreventsDowload }

    $script:uniqueLinks[$LinkUrl] = [pscustomobject]@{
        RecordType       = 'SharingLink'
        ShareLinkUrl     = $LinkUrl
        ShareLinkType    = $Link.Link.Type
        ShareLinkScope   = $Link.Link.Scope
        LinkRoles        = ($Link.Roles -join '|')
        LinkUsers        = ($userEmails -join '|')
        Expiration       = $Link.ExpirationDateTime
        BlocksDownload   = $blocksDownload
        RequiresPassword = $Link.HasPassword
        ShareId          = $Link.Id
        SiteUrl          = $SiteUrl
        ListTitle        = $ListTitle
        ListUrl          = $ListUrl
        ObjectType       = $ObjectType
        ItemName         = $ItemName
        RelativeUrl      = $RelativeUrl
        ItemId           = $ItemId
        Access           = ''
    }
    $script:linksFound++
}

function Add-DirectPermissionRow {
    param(
        [string] $ListId,
        [int]    $ItemId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ObjectType,
        [string] $ItemName,
        [string] $RelativeUrl
    )

    $access = Get-ItemAccessInfo -ListId $ListId -ItemId $ItemId

    $script:directPermissionRows.Add([pscustomobject]@{
        RecordType       = 'DirectPermission'
        ShareLinkUrl     = ''
        ShareLinkType    = ''
        ShareLinkScope   = ''
        LinkRoles        = ''
        LinkUsers        = ''
        Expiration       = ''
        BlocksDownload   = ''
        RequiresPassword = ''
        ShareId          = ''
        SiteUrl          = $SiteUrl
        ListTitle        = $ListTitle
        ListUrl          = $ListUrl
        ObjectType       = $ObjectType
        ItemName         = $ItemName
        RelativeUrl      = $RelativeUrl
        ItemId           = $ItemId
        Access           = $access.Access
    }) | Out-Null
}

function Get-LinksForPath {
    param(
        [string] $RelativeUrl,
        [bool]   $IsFolder
    )
    try {
        if ($IsFolder) {
            return @(Get-PnPFolderSharingLink -Folder $RelativeUrl -ErrorAction Stop)
        }
        return @(Get-PnPFileSharingLink -Identity $RelativeUrl -ErrorAction Stop)
    }
    catch {
        # Surface the real error instead of hiding it - this is the most common reason
        # for "unique perms but zero links" to be misdiagnosed.
        Write-Warning ("Sharing-link API error for {0} : {1}" -f $RelativeUrl, $_.Exception.Message)
        return @()
    }
}

$scanCompletedOk = $false
$scanError = $null

try {
foreach ($list in $lists) {
    $listUrl = $list.RootFolder.ServerRelativeUrl
    Write-Host ("  -> {0,-50} ({1:N0} items)" -f $list.Title, $list.ItemCount) -ForegroundColor DarkGray

    # Note: we intentionally don't probe the library *root* folder. Get-PnPFolderSharingLink
    # resolves its target via the Graph driveItem API, and the library root isn't a driveItem
    # (it IS the drive), so it always returns "Item not found". Sub-folders inside the library
    # are still checked normally because they appear in /items with FSObjType = 1.

    # Show movement immediately for this list, before the first page returns.
    Write-ScanProgress -Processed $processed -Total $totalItems `
        -Status ("Loading '{0}' ({1:N0} items)..." -f $list.Title, $list.ItemCount)

    # Bulk-pull items via REST with $select including HasUniqueRoleAssignments.
    # This is the key efficiency win - 2000 items per HTTP call, with the unique-perm flag inline,
    # so we don't need a per-item round-trip to discover which items can carry sharing links.
    $listId = $list.Id.ToString().ToLower()
    $url    = "/_api/web/lists(guid'$listId')/items?`$select=Id,HasUniqueRoleAssignments,FileRef,FileLeafRef,Title,FSObjType&`$top=2000"

    do {
        try {
            $response = Invoke-PnPSPRestMethod -Url $url -Method Get
        }
        catch {
            # Fail closed: do not skip a failed list and still write a "complete" CSV.
            $detail = $_.Exception.Message
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = [string]$_ }
            throw ("Sharing-links export failed while scanning list '{0}': {1}. No CSV downloaded. Resolve access or throttling, then retry." -f `
                $list.Title, $detail)
        }

        if ($response.value) {
            foreach ($item in $response.value) {
                $processed++

                if (($processed % 25) -eq 0 -or $processed -eq $totalItems) {
                    Write-ScanProgress -Processed $processed -Total $totalItems `
                        -Status ("{0:N0} / {1:N0}  -  {2}  (unique-perm items checked: {3:N0}, links: {4:N0})" `
                            -f $processed, $totalItems, $list.Title, $uniqueCount, $linksFound)
                }

                # Sharing links always break inheritance - skip everything else immediately.
                if (-not $item.HasUniqueRoleAssignments) { continue }
                $uniqueCount++

                $relativeUrl = [string]$item.FileRef
                if ([string]::IsNullOrWhiteSpace($relativeUrl)) { continue }

                $isFolder   = ($list.BaseType -eq 'DocumentLibrary' -and [int]$item.FSObjType -eq 1)
                $links      = Get-LinksForPath -RelativeUrl $relativeUrl -IsFolder $isFolder
                $objectType = if ($isFolder) { 'Folder' } elseif ($list.BaseType -eq 'DocumentLibrary') { 'File' } else { 'Item' }
                $itemName   = if ($list.BaseType -eq 'DocumentLibrary') { [string]$item.FileLeafRef } else { [string]$item.Title }

                # Always echo each unique-perm item so the user can see what the script is examining.
                # Unique-perm items are rare; this output is small and very useful for diagnosis.
                Write-Host ("       unique-perm [{0}]: {1}  -> {2} link(s)" `
                    -f $objectType, $relativeUrl, $links.Count) -ForegroundColor DarkGray

                if ($links.Count -eq 0) {
                    Add-DirectPermissionRow -ListId $listId -ItemId ([int]$item.Id) `
                        -ListTitle $list.Title -ListUrl $listUrl -ObjectType $objectType `
                        -ItemName $itemName -RelativeUrl $relativeUrl
                    $added = $directPermissionRows[$directPermissionRows.Count - 1]
                    Write-Verbose ("         access: {0}" -f $added.Access)
                    continue
                }

                foreach ($link in $links) {
                    $linkUrl = $link.Link.WebUrl
                    if ([string]::IsNullOrWhiteSpace($linkUrl) -or $uniqueLinks.ContainsKey($linkUrl)) { continue }
                    Add-SharingLinkRow -Link $link -LinkUrl $linkUrl -ListTitle $list.Title -ListUrl $listUrl `
                        -ObjectType $objectType -ItemName $itemName -RelativeUrl $relativeUrl -ItemId ([int]$item.Id)
                }
            }
        }

        # SharePoint REST paginates by returning odata.nextLink (an absolute URL) when more items exist.
        $url = $response.'odata.nextLink'
    } while ($url)
}

    $scanCompletedOk = $true
}
catch {
    # Capture then rethrow after finally so Export-Csv cannot run on a partial scan.
    $scanError = $_
}
finally {
    Write-Progress -Activity 'Scanning items for sharing links' -Completed
}

if ($null -ne $scanError) {
    throw $scanError
}
if (-not $scanCompletedOk) {
    throw 'Sharing-links scan did not complete successfully. No CSV downloaded.'
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath   = Join-Path $ScriptDir "SiteSharingLinks-$timestamp.csv"

$reportRows = @($uniqueLinks.Values) + @($directPermissionRows)
$reportRows | Sort-Object RecordType, ListTitle, RelativeUrl |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ("  Items scanned                       : {0:N0}" -f $processed)
Write-Host ("  Items with unique perms             : {0:N0}" -f $uniqueCount)
Write-Host ("  Sharing link rows exported          : {0:N0}" -f $uniqueLinks.Count)
Write-Host ("  Direct-permission rows exported     : {0:N0}" -f $directPermissionRows.Count)
Write-Host ("  CSV                                 : {0}" -f $csvPath)

if ($VerbosePreference -eq 'Continue' -and $directPermissionRows.Count -gt 0) {
    Write-Verbose "Direct-permission items (unique perms, no sharing link):"
    foreach ($row in $directPermissionRows) {
        Write-Verbose ("       - [{0}] {1}" -f $row.ListTitle, $row.RelativeUrl)
        if ($row.Access) {
            Write-Verbose ("         access: {0}" -f $row.Access)
        }
    }
}
