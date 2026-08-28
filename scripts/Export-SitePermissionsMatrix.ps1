<#
.SYNOPSIS
    Exports a site permissions matrix to a formatted Excel workbook.

.DESCRIPTION
    Scans lists/libraries via bulk REST. Role assignments are fetched only at permission
    boundaries; inherited items reuse cached templates from the nearest ancestor.
    Includes a Sharing Links sheet (same columns as Export-SiteSharingLinks.ps1) built during
    the same scan — role assignments and item metadata are not fetched twice.
    Includes an All Items sheet listing every scanned content item with a Broken permissions
    column for quick filtering/auditing.
    A Group Members sheet lists SharePoint group membership once per group (works with
    -ExpandGroups:$false without exploding the permissions matrix).
    The workbook uses a modern dashboard theme: navy headers, section-colored summary,
    tab colors, and conditional highlights on permissions.

.PARAMETER SiteUrl
    SharePoint site URL, or the tenant SharePoint admin URL (e.g. https://contoso-admin.sharepoint.com).
    When an admin URL is used, the script enumerates site collections you can access and lets you
    run against all or a subset (requires SharePoint admin or equivalent API rights for Get-PnPTenantSite).

.PARAMETER ClientId
    Entra app Client ID for interactive PnP sign-in.

.PARAMETER ExpandGroups
    When true (default), expands SharePoint groups into member rows in the permissions matrix (slower).
    Use -ExpandGroups:$false for faster scans; the Group Members sheet always lists group membership.

.PARAMETER IncludeSubsites
    When true (default), recursively scans all subsites under the connected site URL.

.PARAMETER MaxListItemsForFullMatrix
    Lists with more items than this threshold still scan every item for HasUniqueRoleAssignments,
    but only emit matrix rows for item-level unique permissions (inherited items are covered by
    list/site rows). Applies even when the library itself has unique permissions. Default: 2000.

.PARAMETER ListItemPageSize
    Number of list items requested per REST/CSOM page when scanning libraries. Smaller values
    show progress sooner; larger values (up to 5000) minimize round-trips. Default: 5000.

.PARAMETER ItemScanMode
    How list items are paged when checking HasUniqueRoleAssignments. Default: Rest.
    - Rest: Force REST when supported (falls back to CSOM if the field is missing).
    - Auto: Adaptive fastest — prefer REST on larger lists and switch based on observed throughput.
    - Csom: Use CSOM paging for all lists (falls back to REST only if CSOM fails).

.PARAMETER UniquePermProgressWeight
    How much a unique-permission item counts toward overall progress vs a normal item scan.
    Higher values keep the progress bar moving steadily during slow permission fetches. Default: 12.

.PARAMETER IncludeFolderSharingLinks
    When true (default), sharing-link APIs are called for folders with unique permissions. Files are always checked.

.PARAMETER IncludeAllInheritedItemsInMatrix
    When true, matrix rows are emitted for inherited items too (including site-level inherited paths).
    Default is false: matrix shows broken inheritance only.

.PARAMETER OpenWorkbookOnComplete
    When true (default), opens the exported workbook automatically after save.


.PARAMETER NoPicker
    Skip the pre-export picker dialog and run with command-line parameters only.

.NOTES
    Requires: PnP.PowerShell, ImportExcel
#>
[CmdletBinding()]
param(
    [string] $SiteUrl,
    [string] $ClientId,
    [bool]   $ExpandGroups = $true,
    [bool]   $IncludeSubsites = $true,
    [int]    $MaxListItemsForFullMatrix = 2000,
    [ValidateRange(100, 5000)]
    [int]    $ListItemPageSize = 5000,

    [ValidateSet('Auto', 'Rest', 'Csom')]
    [string] $ItemScanMode = 'Rest',

    [ValidateRange(2, 30)]
    [int]    $UniquePermProgressWeight = 12,

    [bool]   $IncludeFolderSharingLinks = $true,
    [bool]   $IncludeAllInheritedItemsInMatrix = $false,
    [bool]   $OpenWorkbookOnComplete = $true,

    [switch] $NoPicker
)

$ErrorActionPreference = 'Stop'

$ListItemPageSize = [Math]::Max(100, [Math]::Min(5000, $ListItemPageSize))
$LargeListRestThreshold = 250
$script:CurrentListPageSize = 0
$script:RunStartUtc = [datetime]::MinValue
$script:LastRoleDefSiteCollection = ''
$script:ProgressRenderState = $null
$script:ProgressRenderRunspace = $null
$script:ProgressRenderPowerShell = $null
$script:ProgressRenderHandle = $null
$script:ProgressSpinnerFrameMs = 80
$script:ProgressStatusThrottleMs = 150

$ExcludedListTitles = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase
)
@(
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
) | ForEach-Object { [void]$ExcludedListTitles.Add($_) }

$PreferredRoleOrder = @(
    'Full Control', 'Design', 'Edit', 'Contribute', 'Read', 'View Only',
    'Create new subsites', 'Approve', 'Manage Hierarchy', 'Restricted Read',
    'Restricted Interfaces for Translation'
)

$NonInheritableRoles = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase
)
@('Limited Access', 'Restricted View', 'Restricted Read', 'Web-Only Limited Access') |
    ForEach-Object { [void]$NonInheritableRoles.Add($_) }

foreach ($mod in @('PnP.PowerShell', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "$mod is not installed. Run: Install-Module $mod -Scope CurrentUser"
    }
}
Import-Module PnP.PowerShell, ImportExcel -ErrorAction Stop

# --- Console theme (ANSI when available, fallback colors otherwise) ---
$script:MatrixLogReady = $false
$script:MatrixUseAnsi = $false

function Initialize-ConsoleTheme {
    if ($script:MatrixLogReady) { return }
    $script:MatrixLogReady = $true
    try {
        if ($PSVersionTable.PSVersion.Major -ge 7 -and $PSStyle) {
            $PSStyle.OutputRendering = 'Ansi'
            $script:MatrixUseAnsi = $true
        }
    }
    catch { $script:MatrixUseAnsi = $false }

    $script:MC = @{
        R       = "`e[0m"
        B       = "`e[1m"
        Dim     = "`e[2m"
        Banner  = "`e[38;5;51m"
        Phase   = "`e[38;5;213m"
        Site    = "`e[38;5;39m"
        List    = "`e[38;5;252m"
        Lib     = "`e[38;5;180m"
        Success = "`e[38;5;46m"
        Warn    = "`e[38;5;220m"
        Skip    = "`e[38;5;245m"
        Picker  = "`e[38;5;81m"
        Export  = "`e[38;5;45m"
        Pulse   = "`e[38;5;117m"
        Tenant  = "`e[38;5;177m"
        Unique  = "`e[38;5;208m"
        Err     = "`e[38;5;203m"
        Info    = "`e[38;5;255m"
    }
}

function Write-MatrixLog {
    param(
        [Parameter(Mandatory)]
        [string] $Message,
        [ValidateSet('Banner', 'Phase', 'Site', 'List', 'Lib', 'Success', 'Warn', 'Skip', 'Dim', 'Info', 'Export', 'Picker', 'Pulse', 'Tenant', 'Unique', 'Error')]
        [string] $Level = 'Info',
        [switch] $NoNewline
    )

    Initialize-ConsoleTheme
    $icon = if ($Level -eq 'Success') { '✓' } else { '' }
    $prefix = if ($icon) { "$icon " } else { '' }

    if ($script:MatrixUseAnsi) {
        $c = $script:MC[$Level]
        if (-not $c) { $c = $script:MC.Info }
        $line = "$($script:MC.B)$prefix$($script:MC.R)$c$Message$($script:MC.R)"
        if ($NoNewline) { Write-Host $line -NoNewline }
        else { Write-Host $line }
    }
    else {
    $fg = switch ($Level) {
        'Success' { 'Green' }
        'Warn'    { 'Yellow' }
        'Skip'    { 'DarkGray' }
        'Dim'     { 'DarkGray' }
        'Site'    { 'Cyan' }
        'Phase'   { 'Magenta' }
        'Picker'  { 'Cyan' }
        'Export'  { 'Green' }
        'Tenant'  { 'Magenta' }
        'Unique'  { 'DarkYellow' }
        'Pulse'   { 'DarkCyan' }
        'List'    { 'Gray' }
        'Lib'     { 'DarkYellow' }
        'Error'   { 'Red' }
        'Banner'  { 'Cyan' }
        default   { 'White' }
    }
    if ($NoNewline) { Write-Host "$prefix$Message" -ForegroundColor $fg -NoNewline }
    else { Write-Host "$prefix$Message" -ForegroundColor $fg }
    }
}

function Write-MatrixBanner {
    Initialize-ConsoleTheme
    $title = 'SharePoint Permissions Matrix Export'
    $innerWidth = 58
    $pad = [Math]::Max(0, $innerWidth - $title.Length)
    $padLeft = [Math]::Floor($pad / 2)
    $padRight = $pad - $padLeft
    $rule = [string]::new([char]0x2550, $innerWidth)

    if ($script:MatrixUseAnsi) {
        Write-Host ''
        Write-Host ("$($script:MC.Banner)  ╔$rule╗$($script:MC.R)")
        $mid = (' ' * $padLeft) + "$($script:MC.Phase)$title$($script:MC.R)" + (' ' * $padRight)
        Write-Host ("$($script:MC.Banner)  ║$mid║$($script:MC.R)")
        Write-Host ("$($script:MC.Banner)  ╚$rule╝$($script:MC.R)")
        Write-Host ''
        return
    }
    $asciiRule = [string]::new('=', $innerWidth)
    Write-Host ''
    Write-Host "  +$asciiRule+" -ForegroundColor Cyan
    Write-Host ("  |{0}{1}{2}|" -f (' ' * $padLeft), $title, (' ' * $padRight)) -ForegroundColor Cyan
    Write-Host "  +$asciiRule+" -ForegroundColor Cyan
    Write-Host ''
}

$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue
}
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = (Get-Location).Path }
$RunMetricsPath = Join-Path $ScriptDir '_metadata\permissions-matrix-run-metrics.jsonl'

# --- Mutable scan state (reset per site) ---
$S = @{
    RoleColumns          = [System.Collections.Generic.List[string]]::new()
    RoleColumnSet        = [System.Collections.Generic.HashSet[string]]::new()
    RoleDefById          = [System.Collections.Generic.Dictionary[int, string]]::new()
    GroupMembers         = [System.Collections.Generic.Dictionary[int, object[]]]::new()
    AssignmentCache      = @{}
    TemplateCache        = @{}
    UniquePermPaths      = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    ListHasUnique        = [System.Collections.Generic.Dictionary[string, bool]]::new()
    ListUniquePermField  = [System.Collections.Generic.Dictionary[string, bool]]::new()
    PathToItemId         = $null
    EmittedKeys          = $null
    SiteName             = ''
    WebPath              = ''
    WebHasUnique         = $false
    WebItemType          = 'Site'
    SiteHostUrl          = ''
    ConnectedWebUrl      = ''
    PnpClientId          = ''
    IsTenantAdminScope   = $false
    TenantRootHost       = ''
    TenantConnectSkipped = 0
    TenantInventorySkipped = 0
    SiteInheritedTpl     = $null
    ProgressLastUtc      = [datetime]::MinValue
    ProgressConsoleLastUtc = [datetime]::MinValue
    ProgressScanStartUtc = [datetime]::MinValue
    ProgressProcessed    = 0
    ProgressTotal        = 0
    ProgressItemsTotal   = 0
    ProgressSiteIndex    = 0
    ProgressSiteTotal    = 0
    ProgressListTitle    = ''
    ProgressPhase        = ''
    ProgressListItemTotal = 0
    ProgressListItemProcessed = 0
    ProgressListUniqueFound = 0
    ProgressUniqueFetchTotal = 0
    ProgressUniqueFetchDone = 0
    ProgressUniqueFetchCurrentPath = ''
    ProgressUniqueFetchStartUtc = [datetime]::MinValue
    ProgressUniqueConsoleLastUtc = [datetime]::MinValue
    ProgressListPageNum = 0
    ProgressListPageCount = 0
    ProgressScanFinished = $false
    ProgressSharingDone = 0
    ProgressSharingBatchTotal = 0
}

$Script:SharingQueue = [System.Collections.Generic.List[object]]::new()
$script:ScanBackendTiming = @{
    Rest = [pscustomobject]@{ Lists = 0; Items = 0; Ms = 0 }
    Csom = [pscustomobject]@{ Lists = 0; Items = 0; Ms = 0 }
}

function Test-IsInternalRoleCache {
    param($Obj)

    if ($null -eq $Obj) { return $false }
    if ($null -ne $S.AssignmentCache -and $Obj -eq $S.AssignmentCache) { return $true }
    if ($null -ne $S.TemplateCache -and $Obj -eq $S.TemplateCache) { return $true }
    return $false
}

function Test-IsAnyKeyValuePair {
    param($Obj)

    if ($null -eq $Obj) { return $false }
    return ([string]$Obj.GetType().FullName -like 'System.Collections.Generic.KeyValuePair`2*')
}

function Get-RestPropertyValue {
    param(
        $Object,
        [string] $Name
    )

    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
    if (Test-IsInternalRoleCache $Object) { return $null }
    if (Test-IsAnyKeyValuePair $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary] -and $Object -isnot [System.Management.Automation.PSCustomObject]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
    }

    return $null
}

function Test-IsRestPrincipalObject {
    param($Obj)

    if ($null -eq $Obj) { return $false }
    if (Test-IsInternalRoleCache $Obj) { return $false }
    if (Test-IsAnyKeyValuePair $Obj) { return $false }
    if ($Obj -is [System.Collections.IDictionary] -and $Obj -isnot [System.Management.Automation.PSCustomObject]) { return $false }
    if ($Obj -is [System.Array] -and $Obj -isnot [System.Management.Automation.PSCustomObject]) { return $false }

    return ($null -ne (Get-RestPropertyValue $Obj 'Title')) `
        -or ($null -ne (Get-RestPropertyValue $Obj 'LoginName')) `
        -or ($null -ne (Get-RestPropertyValue $Obj 'PrincipalType')) `
        -or ($null -ne (Get-RestPropertyValue $Obj 'Id'))
}

function Test-IsRoleAssignmentObject {
    param($Obj)

    if ($null -eq $Obj) { return $false }
    if (Test-IsInternalRoleCache $Obj) { return $false }
    if ($Obj -is [System.Collections.IDictionary] -and $Obj -isnot [System.Management.Automation.PSCustomObject]) { return $false }
    if (Test-IsAnyKeyValuePair $Obj) { return $false }
    if ($Obj -is [System.Collections.Generic.KeyValuePair[object, object]]) { return $false }
    if ($Obj -is [System.Array] -and $Obj -isnot [System.Management.Automation.PSCustomObject]) { return $false }

    $member = Get-RestPropertyValue $Obj 'Member'
    if (Test-IsRestPrincipalObject $member) { return $true }
    if ($null -ne (Get-RestPropertyValue $Obj 'PrincipalId')) { return $true }
    if ($null -ne (Get-RestPropertyValue $Obj 'Id') -and
        ($null -ne (Get-RestPropertyValue $Obj 'RoleDefinitionBindings') -or $null -ne $member)) {
        return $true
    }
    return $false
}

function Expand-AssignmentCollection {
    param($Raw)

    if ($null -eq $Raw) { return @() }
    if (Test-IsInternalRoleCache $Raw) { return @() }

    if ($Raw -is [System.Array]) {
        return @(Get-NormalizedAssignmentArray $Raw)
    }

    if ($Raw -is [System.Collections.IDictionary] -and $Raw -isnot [System.Management.Automation.PSCustomObject]) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $Raw.Values) {
            if ($entry -is [System.Array] -and $entry -isnot [System.Management.Automation.PSCustomObject]) {
                foreach ($inner in @($entry)) {
                    if (Test-IsRoleAssignmentObject $inner) { $list.Add($inner) | Out-Null }
                }
            }
            elseif (Test-IsRoleAssignmentObject $entry) {
                $list.Add($entry) | Out-Null
            }
        }
        return @($list)
    }

    if (Test-IsRoleAssignmentObject $Raw) { return @($Raw) }
    return @(Get-NormalizedAssignmentArray @($Raw))
}

function Get-NormalizedAssignmentArray {
    param($Assignments)

    if ($null -eq $Assignments) { return @() }
    if (Test-IsInternalRoleCache $Assignments) {
        Write-Warning 'Role assignments parameter was an internal cache dictionary; skipping.'
        return @()
    }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @($Assignments)) {
        if (Test-IsAnyKeyValuePair $item) { $item = $item.Value }
        if (-not (Test-IsRoleAssignmentObject $item)) { continue }
        $list.Add($item) | Out-Null
    }
    return @($list)
}

$Stats = @{
    RootSites            = 0
    Subsites             = 0
    Lists                = 0
    Libraries            = 0
    Files                = 0
    Folders              = 0
    ListItems            = 0
    ItemsScanned         = 0
    SitesWithUniquePerms = 0
    ListsWithUniquePerms = 0
    UniqueItemPaths      = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    SiteDetails          = [System.Collections.Generic.List[object]]::new()
    MatrixTopLevelRows   = 0
    MatrixCustomRows     = 0
    MatrixInheritedRows  = 0
    MatrixSharingLinkRows = 0
    MatrixUsers          = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    MatrixExternalUsers  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    MatrixSpGroups       = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    MatrixPaths          = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    MatrixSiteNames      = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    ExternalUsersBySite  = @{}
    BulkSkippedItems     = 0
}

$SummaryPadRow = 2
$SummaryPadCol = 2
$SummaryOverviewStartRow = 11
$SummaryChartDataMinRow = 80
$SummaryChartDataLabelLogicalCol = 11
$SummaryChartDataValueLogicalCol = 12
$SummaryDetailsLogicalEnd = 3
$SummaryGutterLogicalCol = 4
$SummaryInsightsLogicalStart = 5
$SummaryInsightsLogicalEnd = 10
$SummaryDashboardLastCol = 11
$SummarySheetLastCol = 14

function Get-SummaryRow([int]$LogicalRow) { return $LogicalRow + $SummaryPadRow }
function Get-SummaryCol([int]$LogicalCol) { return $LogicalCol + $SummaryPadCol }
function Get-SummaryDetailsEndCol { return Get-SummaryCol $SummaryDetailsLogicalEnd }
function Get-SummaryInsightsStartCol { return Get-SummaryCol $SummaryInsightsLogicalStart }
function Get-SummaryInsightsEndCol { return Get-SummaryCol $SummaryInsightsLogicalEnd }

$SharingReport = @{
    UniqueLinks          = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    DirectRows           = [System.Collections.Generic.List[object]]::new()
    UniqueItemsChecked   = 0
}

$GroupCatalog = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
$GroupMemberRows = [System.Collections.Generic.List[object]]::new()
$GroupMembersFlushed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$AllItemRows = [System.Collections.Generic.List[object]]::new()

$SharingLinkColumns = @(
    'RecordType', 'ShareLinkUrl', 'ShareLinkType', 'ShareLinkScope', 'LinkRoles', 'LinkUsers',
    'External user', 'Expiration', 'BlocksDownload', 'RequiresPassword', 'ShareId', 'SiteUrl', 'ListTitle',
    'ListUrl', 'ObjectType', 'ItemName', 'RelativeUrl', 'ItemId', 'Access'
)

$GroupMemberColumns = @(
    'Site Name', 'Site URL', 'Site Path', 'Group Name', 'Group Login', 'Group Id',
    'Member Name', 'Member Login', 'Member Email', 'Member Type', 'External user'
)

$AllItemsColumns = @(
    'Site Name', 'List/Library', 'Item path', 'Item Type', 'Item Id', 'Broken permissions',
    'Created', 'Created By', 'Modified', 'Modified By'
)

$SortMatrixRowLimit = 100000
$StyleRowLimit = 50000
$TableHeaderFilterPadding = 2.8
$script:ThemeColorCache = @{}
$script:TableColumnIndexCache = @{}

function Get-ThemeColor {
    param([ValidateSet('Primary', 'PrimaryMid', 'PrimarySoft', 'Accent', 'AccentLight', 'Accent2', 'Surface', 'SurfaceAlt', 'Panel', 'Text', 'Muted', 'Highlight', 'Warning', 'Border', 'BorderLight', 'Custom', 'Inherited', 'TopLevel', 'White', 'ChartBar', 'ChartBar2')]
    [string] $Name)

    if ($script:ThemeColorCache.ContainsKey($Name)) { return $script:ThemeColorCache[$Name] }

    $color = switch ($Name) {
        'Primary'     { [System.Drawing.Color]::FromArgb(15, 39, 68) }
        'PrimaryMid'  { [System.Drawing.Color]::FromArgb(30, 58, 95) }
        'PrimarySoft' { [System.Drawing.Color]::FromArgb(236, 242, 250) }
        'Accent'      { [System.Drawing.Color]::FromArgb(13, 148, 136) }
        'AccentLight' { [System.Drawing.Color]::FromArgb(204, 251, 241) }
        'Accent2'     { [System.Drawing.Color]::FromArgb(59, 130, 246) }
        'Surface'     { [System.Drawing.Color]::FromArgb(248, 250, 252) }
        'SurfaceAlt'  { [System.Drawing.Color]::FromArgb(241, 245, 249) }
        'Panel'       { [System.Drawing.Color]::White }
        'Text'        { [System.Drawing.Color]::FromArgb(15, 23, 42) }
        'Muted'       { [System.Drawing.Color]::FromArgb(100, 116, 139) }
        'Highlight'   { [System.Drawing.Color]::FromArgb(220, 252, 231) }
        'Warning'     { [System.Drawing.Color]::FromArgb(254, 243, 199) }
        'Border'      { [System.Drawing.Color]::FromArgb(203, 213, 225) }
        'BorderLight' { [System.Drawing.Color]::FromArgb(226, 232, 240) }
        'Custom'      { [System.Drawing.Color]::FromArgb(255, 237, 213) }
        'Inherited'   { [System.Drawing.Color]::FromArgb(241, 245, 249) }
        'TopLevel'    { [System.Drawing.Color]::FromArgb(219, 234, 254) }
        'White'       { [System.Drawing.Color]::White }
        'ChartBar'    { [System.Drawing.Color]::FromArgb(13, 148, 136) }
        'ChartBar2'   { [System.Drawing.Color]::FromArgb(59, 130, 246) }
    }
    $script:ThemeColorCache[$Name] = $color
    return $color
}

function Get-TableColumnSpan([OfficeOpenXml.Table.ExcelTable] $Table) {
    $cStart = [int]$Table.Address.Start.Column
    $cEnd = [int]$Table.Address.End.Column
    # EPPlus table addresses can be 0-based for column A; Cells[] is always 1-based.
    if ($cStart -lt 1) {
        $cStart = 1
        $cEnd = [int]$Table.Address.End.Column + 1
    }
    if ($cEnd -lt $cStart) { $cEnd = $cStart }
    return @{ Start = $cStart; End = $cEnd }
}

function Get-ExcelRange {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $FromRow,
        [int] $FromCol,
        [int] $ToRow = 0,
        [int] $ToCol = 0
    )

    if ($ToRow -lt 1) { $ToRow = $FromRow }
    if ($ToCol -lt 1) { $ToCol = $FromCol }
    if ($FromRow -eq $ToRow -and $FromCol -eq $ToCol) {
        # Comma forces a single cell object (ExcelRange is enumerable in PowerShell).
        return , $Worksheet.Cells[$FromRow, $FromCol]
    }
    return , $Worksheet.Cells.Item($FromRow, $FromCol, $ToRow, $ToCol)
}


function Merge-ExcelRange {
    param(
        [OfficeOpenXml.ExcelRange] $Range,
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $TopRow = 0,
        [int] $LeftCol = 0
    )

    if (-not $Range) { return }
    $Range.Merge = $true
    if ($Worksheet -and $TopRow -ge 1 -and $LeftCol -ge 1) {
        return $Worksheet.Cells[$TopRow, $LeftCol]
    }
}

function Set-ExcelCellValue {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $Row,
        [int] $Col,
        $Value
    )
    $Worksheet.Cells[$Row, $Col].Value = $Value
}

function Get-TextDisplayWidth {
    param([string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $len = 0
    foreach ($ch in $Text.ToCharArray()) {
        if ([int]$ch -gt 255) { $len += 2 } else { $len++ }
    }
    return $len
}

function Get-HeaderColumnMinWidth {
    param(
        [string] $HeaderText,
        [double] $ProfileMin = 10
    )

    $headerLen = Get-TextDisplayWidth $HeaderText
    $forFilter = $headerLen + $TableHeaderFilterPadding + 1.5
    return [Math]::Max([double]$ProfileMin, $forFilter)
}

function Set-ColumnWidthForContent {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $Column,
        [int] $RowStart,
        [int] $RowEnd,
        [string] $HeaderText = '',
        [double] $MinWidth = 10,
        [double] $MaxWidth = 60,
        [double] $Padding = 2
    )

    if ($Column -lt 1 -or $RowEnd -lt $RowStart) { return }

    if ($HeaderText) {
        $MinWidth = Get-HeaderColumnMinWidth -HeaderText $HeaderText -ProfileMin $MinWidth
    }

    $maxLen = 0
    for ($r = $RowStart; $r -le $RowEnd; $r++) {
        $len = Get-TextDisplayWidth ([string]$Worksheet.Cells[$r, $Column].Value)
        if ($len -gt $maxLen) { $maxLen = $len }
    }
    $width = [Math]::Min($MaxWidth, [Math]::Max($MinWidth, ($maxLen * 1.05) + $Padding))
    $Worksheet.Column($Column).Width = $width
}

function Format-ExcelTableColumns {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [OfficeOpenXml.Table.ExcelTable] $Table,
        [hashtable] $Profiles = @{},
        [int] $SampleRows = 300,
        [double] $DefaultMin = 12,
        [double] $DefaultMax = 48
    )

    if (-not $Table) { return }

    $hdr = [int]$Table.Address.Start.Row
    $rEnd = [int]$Table.Address.End.Row
    $rSample = if ($SampleRows -gt 0) { [Math]::Min($rEnd, $hdr + $SampleRows) } else { $rEnd }
    $span = Get-TableColumnSpan $Table

    for ($c = $span.Start; $c -le $span.End; $c++) {
        $name = [string]$Worksheet.Cells[$hdr, $c].Value
        $prof = $Profiles[$name]
        $minW = if ($prof) { [double]$prof.Min } else { $DefaultMin }
        $maxW = if ($prof) { [double]$prof.Max } else { $DefaultMax }
        Set-ColumnWidthForContent -Worksheet $Worksheet -Column $c -RowStart $hdr -RowEnd $rSample `
            -HeaderText $name -MinWidth $minW -MaxWidth $maxW

        $hdrCell = $Worksheet.Cells[$hdr, $c]
        $hdrCell.Style.WrapText = $false
        $hdrCell.Style.ShrinkToFit = $false

        if ($prof -and $prof.Wrap -and (($rEnd - $hdr) -le $StyleRowLimit)) {
            $wrapRange = Get-ExcelRange -Worksheet $Worksheet -FromRow ($hdr + 1) -FromCol $c -ToRow $rEnd -ToCol $c
            if ($wrapRange) { $wrapRange.Style.WrapText = $true }
        }
    }
}


function Get-DataTableColumnProfiles {
    param([string] $TableKind)

    switch ($TableKind) {
        'Matrix' {
            return @{
                'Site Name'      = @{ Min = 12; Max = 24 }
                'Name'           = @{ Min = 14; Max = 36 }
                'Item path'      = @{ Min = 34; Max = 72; Wrap = $true }
                'Item Type'      = @{ Min = 12; Max = 16 }
                'Inheritance'    = @{ Min = 13; Max = 16 }
                'Details'        = @{ Min = 28; Max = 68; Wrap = $true }
                'User/group'     = @{ Min = 18; Max = 38 }
                'Principal type' = @{ Min = 16; Max = 24 }
                'Account name'   = @{ Min = 22; Max = 52 }
                'External user'  = @{ Min = 15; Max = 18 }
                'Given through'  = @{ Min = 16; Max = 24 }
            }
        }
        'GroupMembers' {
            return @{
                'Site Name'     = @{ Min = 12; Max = 24 }
                'Site URL'      = @{ Min = 32; Max = 72; Wrap = $true }
                'Site Path'     = @{ Min = 28; Max = 60; Wrap = $true }
                'Group Name'    = @{ Min = 18; Max = 40 }
                'Group Login'   = @{ Min = 18; Max = 40 }
                'Group Id'      = @{ Min = 10; Max = 14 }
                'Member Name'   = @{ Min = 18; Max = 36 }
                'Member Login'  = @{ Min = 22; Max = 52 }
                'Member Email'  = @{ Min = 24; Max = 52 }
                'Member Type'   = @{ Min = 14; Max = 18 }
                'External user' = @{ Min = 15; Max = 18 }
            }
        }
        'SharingLinks' {
            return @{
                'RecordType'       = @{ Min = 14; Max = 20 }
                'ShareLinkUrl'     = @{ Min = 36; Max = 80; Wrap = $true }
                'ShareLinkType'    = @{ Min = 14; Max = 18 }
                'ShareLinkScope'   = @{ Min = 14; Max = 20 }
                'LinkRoles'        = @{ Min = 12; Max = 18 }
                'LinkUsers'        = @{ Min = 28; Max = 72; Wrap = $true }
                'External user'    = @{ Min = 15; Max = 18 }
                'Expiration'       = @{ Min = 12; Max = 20 }
                'BlocksDownload'   = @{ Min = 16; Max = 20 }
                'RequiresPassword' = @{ Min = 18; Max = 22 }
                'ShareId'          = @{ Min = 34; Max = 42 }
                'SiteUrl'          = @{ Min = 32; Max = 72; Wrap = $true }
                'ListTitle'        = @{ Min = 14; Max = 36 }
                'ListUrl'          = @{ Min = 28; Max = 60; Wrap = $true }
                'ObjectType'       = @{ Min = 12; Max = 18 }
                'ItemName'         = @{ Min = 16; Max = 40 }
                'RelativeUrl'      = @{ Min = 28; Max = 68; Wrap = $true }
                'ItemId'           = @{ Min = 10; Max = 14 }
                'Access'           = @{ Min = 20; Max = 60; Wrap = $true }
            }
        }
        'AllItems' {
            return @{
                'Site Name'           = @{ Min = 12; Max = 24 }
                'List/Library'        = @{ Min = 16; Max = 38 }
                'Item path'           = @{ Min = 34; Max = 80; Wrap = $true }
                'Item Type'           = @{ Min = 12; Max = 18 }
                'Item Id'             = @{ Min = 10; Max = 14 }
                'Broken permissions'  = @{ Min = 18; Max = 22 }
                'Created'             = @{ Min = 18; Max = 22 }
                'Created By'          = @{ Min = 18; Max = 32 }
                'Modified'            = @{ Min = 18; Max = 22 }
                'Modified By'         = @{ Min = 18; Max = 32 }
            }
        }
        default { return @{} }
    }
}

function Get-SummarySheetLastTableRow {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $last = 0
    foreach ($tbl in $Worksheet.Tables) {
        if ($tbl.Address.End.Row -gt $last) { $last = [int]$tbl.Address.End.Row }
    }
    return $last
}

function Trim-SummaryWorksheetExtent {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $lastRow = Get-SummarySheetLastTableRow -Worksheet $Worksheet
    if ($lastRow -lt 1) { return }

    $lastRow = [Math]::Max($lastRow + 8, $SummaryChartDataMinRow + 40)
    $dim = $Worksheet.Dimension
    if (-not $dim) { return }

    $endRow = [int]$dim.End.Row
    $endCol = [int]$dim.End.Column

    if ($endCol -gt $SummarySheetLastCol) {
        $fromLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($SummarySheetLastCol + 1)
        $toLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($endCol)
        try {
            $Worksheet.Cells["${fromLtr}1:${toLtr}${endRow}"].Clear()
        }
        catch { }
    }

    if ($endRow -gt $lastRow + 40) {
        $fromLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter(1)
        $toLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($SummarySheetLastCol)
        try {
            $Worksheet.Cells["${fromLtr}$($lastRow + 41):${toLtr}${endRow}"].Clear()
        }
        catch { }
    }
}

function Clear-SummaryStrayContent {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    Remove-BrokenSummaryTables -Worksheet $Worksheet

    $lastRow = Get-SummarySheetLastTableRow -Worksheet $Worksheet
    if ($lastRow -lt 1) { return }

    $dim = $Worksheet.Dimension
    if (-not $dim -or [int]$dim.End.Row -le $lastRow) { return }

    for ($r = [int]$dim.End.Row; $r -gt $lastRow; $r--) {
        $v = [string]$Worksheet.Cells[$r, 1].Value
        if ($v -match '^Column\d+$') {
            $Worksheet.DeleteRow($r, 1)
        }
    }

    if ([int]$dim.End.Column -gt $SummarySheetLastCol) {
        $fromLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($SummarySheetLastCol + 1)
        $toLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter([int]$dim.End.Column)
        $endRow = [Math]::Max($lastRow + 5, [int]$dim.End.Row)
        try {
            $Worksheet.Cells["${fromLtr}1:${toLtr}${endRow}"].Clear()
        }
        catch { }
    }
}

function Set-WorksheetTabColor {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [System.Drawing.Color] $Color
    )
    try {
        $Worksheet.TabColor = $Color
    }
    catch {
        Write-Verbose "Tab color skipped for '$($Worksheet.Name)': $_"
    }
}

function Set-RangeFontStyle {
    param(
        [OfficeOpenXml.ExcelRange] $Range,
        [string] $FontName = 'Segoe UI',
        [float] $FontSize = 10,
        [System.Drawing.Color] $FontColor
    )
    if (-not $Range) { return }
    $Range.Style.Font.Name = $FontName
    $Range.Style.Font.Size = $FontSize
    if ($FontColor) {
        $Range.Style.Font.Color.SetColor($FontColor)
    }
}

function Set-DashboardWorksheetBase {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [System.Drawing.Color] $TabColor
    )

    $Worksheet.View.ShowGridLines = $false
    Set-WorksheetTabColor -Worksheet $Worksheet -Color $TabColor
    $Worksheet.View.ZoomScale = 90
    if ($Worksheet.Dimension) {
        Set-RangeFontStyle -Range $Worksheet.Cells[$Worksheet.Dimension.Address] `
            -FontColor (Get-ThemeColor Text)
    }
}

function Set-RangeFill {
    param(
        [OfficeOpenXml.ExcelRange] $Range,
        [System.Drawing.Color] $Color
    )
    if (-not $Range) { return }
    $Range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $Range.Style.Fill.BackgroundColor.SetColor($Color)
}

function Set-ExcelDrawingFill {
    param(
        $Fill,
        [System.Drawing.Color] $Color,
        [ValidateSet('NoFill', 'SolidFill')]
        [string] $Style = 'SolidFill'
    )

    if (-not $Fill) { return }

    try {
        $Fill.Style = $Style
        if ($Style -eq 'SolidFill') {
            try { $Fill.Color = $Color }
            catch {
                try { $Fill.Color.SetColor($Color) } catch { }
            }
        }
    }
    catch { }
}

function Set-RangeBorder {
    param(
        [OfficeOpenXml.ExcelRange] $Range,
        [OfficeOpenXml.Style.ExcelBorderStyle] $Style = 'Thin',
        [System.Drawing.Color] $Color
    )
    if (-not $Range) { return }
    $border = $Range.Style.Border
    $border.Top.Style = $Style
    $border.Bottom.Style = $Style
    $border.Left.Style = $Style
    $border.Right.Style = $Style
    $colorObj = if ($Color) { $Color } else { Get-ThemeColor Border }
    $border.Top.Color.SetColor($colorObj)
    $border.Bottom.Color.SetColor($colorObj)
    $border.Left.Color.SetColor($colorObj)
    $border.Right.Color.SetColor($colorObj)
}

function Set-WorksheetCanvas {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [switch] $SummaryDashboard
    )

    $Worksheet.View.ShowGridLines = $false
    $Worksheet.View.ZoomScale = 100
    $endRow = 120
    if ($Worksheet.Dimension) {
        $endRow = [Math]::Max(120, [int]$Worksheet.Dimension.End.Row + 15)
    }
    $firstCol = if ($SummaryDashboard) { (Get-SummaryCol 1) } else { 1 }
    $lastCol = if ($SummaryDashboard) { $SummarySheetLastCol } else { 10 }
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow 1 -FromCol $firstCol -ToRow $endRow `
        -ToCol $lastCol) -Color (Get-ThemeColor Surface)
}

function Apply-SummaryPagePadding {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $Worksheet.Column(1).Width = 10
    $Worksheet.Column(2).Width = 10
    $Worksheet.Row(1).Height = 18
    $Worksheet.Row(2).Height = 18
    $surface = Get-ThemeColor Surface
    $padLeft = Get-ExcelRange -Worksheet $Worksheet -FromRow 1 -FromCol 1 -ToRow 220 -ToCol 2
    Set-RangeFill -Range $padLeft -Color $surface
}

function Set-SummaryViewSettings {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $Worksheet.View.ShowGridLines = $false
    try { $Worksheet.View.ShowHeaders = $true } catch { }
    try { $Worksheet.View.ShowRowColHeaders = $false } catch { }
    $Worksheet.View.PageBreakView = $false
    $Worksheet.View.PageLayoutView = $false
    $Worksheet.View.ZoomScale = 100

    $activeRow = Get-SummaryRow 1
    $activeCol = Get-SummaryCol 1
    if ($activeRow -ge 1 -and $activeCol -ge 1) {
        try {
            $activeCell = $Worksheet.Cells[$activeRow, $activeCol]
            if ($null -ne $activeCell -and $activeCell.Start.Row -ge 1 -and $activeCell.Start.Column -ge 1) {
                $Worksheet.View.ActiveCell = $activeCell
            }
        }
        catch {
            try {
                $Worksheet.Select((New-Object OfficeOpenXml.ExcelCellAddress($activeRow, $activeCol)).Address)
            }
            catch { }
        }
    }

    $freezeRow = $SummaryOverviewStartRow
    $freezeCol = Get-SummaryCol 1
    try { $Worksheet.View.FreezePanes($freezeRow, $freezeCol) } catch { }
}

function Apply-SummaryBackdropFill {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $surface = Get-ThemeColor Surface
    $endRow = 200
    if ($Worksheet.Dimension) {
        $endRow = [Math]::Max(80, [int]$Worksheet.Dimension.End.Row + 8)
    }

    # Side margins and hidden chart-data columns only — never paint over data tables.
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow 1 -FromCol 1 -ToRow $endRow -ToCol 2) `
        -Color $surface
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow 1 -FromCol 13 -ToRow $endRow -ToCol 14) `
        -Color $surface

    $overview = $Worksheet.Tables['SummaryOverview']
    if (-not $overview) { return }

    # Paint the full insights band through Site Breakdown so empty cells beside charts are not white/gridlined.
    $panelTop = [int]$overview.Address.Start.Row
    $panelBottom = [int]$overview.Address.End.Row
    $siteBreakdown = $Worksheet.Tables['SiteBreakdown']
    if ($siteBreakdown) {
        $panelBottom = [Math]::Max($panelBottom, [int]$siteBreakdown.Address.End.Row)
    }
    $insightsStart = Get-SummaryInsightsStartCol
    $insightsEnd = Get-SummaryInsightsEndCol
    $gutterCol = Get-SummaryCol $SummaryGutterLogicalCol

    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $panelTop -FromCol $insightsStart `
        -ToRow $panelBottom -ToCol $insightsEnd) -Color $surface
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $panelTop -FromCol $gutterCol `
        -ToRow $panelBottom -ToCol $gutterCol) -Color $surface

    $kpiRow1 = Get-SummaryRow 5
    $kpiRow2 = Get-SummaryRow 6
    $detailsEnd = Get-SummaryDetailsEndCol
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $kpiRow1 -FromCol (Get-SummaryCol 1) `
        -ToRow $kpiRow2 -ToCol $detailsEnd) -Color $surface
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow (Get-SummaryRow 7) -FromCol (Get-SummaryCol 1) `
        -ToRow (Get-SummaryRow 7) -ToCol $insightsEnd) -Color $surface
}

function Restore-SummaryTableHeaders {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    foreach ($tableName in @('SummaryOverview', 'SiteBreakdown')) {
        $tbl = $Worksheet.Tables[$tableName]
        if ($tbl) {
            Format-TableHeaderRow -Worksheet $Worksheet -Table $tbl
        }
    }
}

function Refresh-SummaryChartStyling {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    foreach ($drawing in @($Worksheet.Drawings)) {
        if ($drawing -isnot [OfficeOpenXml.Drawing.Chart.ExcelChart]) { continue }

        $barColor = if ([string]$drawing.Name -eq 'ExternalBySite') {
            Get-ThemeColor ChartBar2
        } else {
            Get-ThemeColor ChartBar
        }
        Set-SummaryChartStyle -Chart $drawing -BarColor $barColor
        Remove-SummaryChartGridlines -Chart $drawing
    }
}

function Remove-SummaryChartGridlines {
    param([OfficeOpenXml.Drawing.Chart.ExcelChart] $Chart)

    foreach ($axisName in @('XAxis', 'YAxis')) {
        try {
            $axis = $Chart.$axisName
            $axis.RemoveGridlines()
            $axis.MajorGridlines.Width = 0
            Set-ExcelDrawingFill -Fill $axis.MajorGridlines.Fill -Color (Get-ThemeColor Surface) -Style 'NoFill'
            $axis.MinorGridlines.Width = 0
            Set-ExcelDrawingFill -Fill $axis.MinorGridlines.Fill -Color (Get-ThemeColor Surface) -Style 'NoFill'
        }
        catch { }
    }
}

function Complete-SummaryDashboardPresentation {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    Apply-SummaryBackdropFill -Worksheet $Worksheet
    Restore-SummaryTableHeaders -Worksheet $Worksheet
    Refresh-SummaryChartStyling -Worksheet $Worksheet
    Set-SummaryViewSettings -Worksheet $Worksheet
}

function Reset-SummaryWorksheetBounds {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $MaxColumn = $SummarySheetLastCol
    )

    $dim = $Worksheet.Dimension
    if (-not $dim -or [int]$dim.End.Column -le $MaxColumn) { return }

    $lastRow = [Math]::Max(60, [int]$dim.End.Row)
    $fromLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($MaxColumn + 1)
    $toLtr = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter([int]$dim.End.Column)
    try {
        $Worksheet.Cells["${fromLtr}1:${toLtr}${lastRow}"].Clear()
    }
    catch {
        # Ignore clear failures on overlapping merged chart/KPI regions.
    }
}

function Remove-ExcelTableByName {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [string] $TableName
    )

    $tbl = $Worksheet.Tables[$TableName]
    if ($tbl) {
        $Worksheet.Tables.Delete($tbl)
    }
}

function Clear-SummaryDrawings {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    while ($Worksheet.Drawings.Count -gt 0) {
        $Worksheet.Drawings.Remove($Worksheet.Drawings.Count)
    }
}

function Remove-BrokenSummaryTables {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    foreach ($tbl in @($Worksheet.Tables)) {
        $hdr = [int]$tbl.Address.Start.Row
        $c1 = [int]$tbl.Address.Start.Column
        if ($c1 -lt 1) { $c1 = 1 }
        $first = [string]$Worksheet.Cells[$hdr, $c1].Value
        if ($first -match '^Column\d+$') {
            $Worksheet.Tables.Delete($tbl)
        }
    }
}

function Write-SiteBreakdownTable {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [object[]] $SiteRows
    )

    if ($SiteRows.Count -lt 1) { return }

    Remove-BrokenSummaryTables -Worksheet $Worksheet
    Remove-ExcelTableByName -Worksheet $Worksheet -TableName 'SiteBreakdown'
    Reset-SummaryWorksheetBounds -Worksheet $Worksheet

    $overview = $Worksheet.Tables['SummaryOverview']
    if (-not $overview) {
        Write-Warning 'SummaryOverview table not found; cannot write site breakdown.'
        return
    }

    $labelRow = [int]$overview.Address.End.Row + 2
    $headerRow = $labelRow + 1
    $headers = @(
        'Site Name', 'Site Path', 'Type', 'Lists/Libraries', 'Items',
        'Files', 'Folders', 'List items', 'Unique items', 'External users', 'Site unique perms'
    )

    $colCount = 11
    $firstCol = Get-SummaryCol 1
    $lastCol = Get-SummaryCol $colCount
    for ($c = 1; $c -le $colCount; $c++) {
        Set-ExcelCellValue -Worksheet $Worksheet -Row $headerRow -Col (Get-SummaryCol $c) -Value $headers[$c - 1]
    }

    $row = $headerRow + 1
    foreach ($site in $SiteRows) {
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 1) -Value ([string]$site.'Site Name')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 2) -Value ([string]$site.'Site Path')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 3) -Value ([string]$site.'Type')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 4) -Value ([int]$site.'Lists/Libraries')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 5) -Value ([int]$site.'Items')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 6) -Value ([int]$site.'Files')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 7) -Value ([int]$site.'Folders')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 8) -Value ([int]$site.'List items')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 9) -Value ([int]$site.'Unique items')
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 10) -Value ([int](Get-ExternalUserCountBySite -SiteName ([string]$site.'Site Name')))
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col (Get-SummaryCol 11) -Value ([string]$site.'Site unique perms')
        $row++
    }

    $lastRow = $row - 1
    $range = Get-ExcelRange -Worksheet $Worksheet -FromRow $headerRow -FromCol $firstCol -ToRow $lastRow -ToCol $lastCol
    if (-not $range) { throw 'Failed to resolve site breakdown table range.' }

    $table = $Worksheet.Tables.Add($range, 'SiteBreakdown')
    $table.ShowHeader = $true
    $table.ShowFilter = $true
    $table.ShowRowStripes = $false
    $table.TableStyle = [OfficeOpenXml.Table.TableStyles]::None
}

function Write-KpiCard {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $Row,
        [int] $ColStart,
        [int] $ColEnd,
        [string] $Label,
        $Value,
        [System.Drawing.Color] $AccentColor
    )

    if ($ColStart -lt 1 -or $ColEnd -lt $ColStart) { return }

    for ($c = $ColStart; $c -le $ColEnd; $c++) {
        if ($Worksheet.Column($c).Width -lt 16) { $Worksheet.Column($c).Width = 16 }
    }

    $card = Get-ExcelRange -Worksheet $Worksheet -FromRow $Row -FromCol $ColStart -ToRow $Row -ToCol $ColEnd
    Merge-ExcelRange -Range $card -Worksheet $Worksheet -TopRow $Row -LeftCol $ColStart | Out-Null
    Set-RangeFill -Range $card -Color (Get-ThemeColor Panel)
    Set-RangeBorder -Range $card -Style ([OfficeOpenXml.Style.ExcelBorderStyle]::Thin) -Color (Get-ThemeColor BorderLight)
    $card.Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
    $card.Style.Border.Left.Color.SetColor($AccentColor)
    $card.Style.Font.Name = 'Segoe UI'
    $card.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $card.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    $card.Style.WrapText = $true
    $card.Style.Indent = 2

    $cell = $Worksheet.Cells[$Row, $ColStart]
    $rt = $cell.RichText
    if ($rt.Count -gt 0) { $rt.Clear() }
    $labelPart = $rt.Add("$Label`r`n")
    $labelPart.Size = 8
    $labelPart.Color = (Get-ThemeColor Muted)
    $valueText = if ($Value -is [string]) { [string]$Value } else { "{0:N0}" -f [double]$Value }
    $valuePart = $rt.Add($valueText)
    $valuePart.Size = 20
    $valuePart.Bold = $true
    $valuePart.Color = (Get-ThemeColor Text)

    $Worksheet.Row($Row).Height = 52
}

function Format-SummaryHero {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [string] $Title,
        [string] $Subtitle
    )

    $r1 = Get-SummaryRow 1
    $r2 = Get-SummaryRow 2
    $r3 = Get-SummaryRow 3
    $c1 = Get-SummaryCol 1
    $cEnd = Get-SummaryInsightsEndCol

    $hero = Get-ExcelRange -Worksheet $Worksheet -FromRow $r1 -FromCol $c1 -ToRow $r2 -ToCol $cEnd
    Merge-ExcelRange -Range $hero -Worksheet $Worksheet -TopRow $r1 -LeftCol $c1 | Out-Null
    Set-RangeFill -Range $hero -Color (Get-ThemeColor Primary)
    $hero.Style.Font.Name = 'Segoe UI'
    $hero.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $hero.Style.WrapText = $true

    $cell = $Worksheet.Cells[$r1, $c1]
    $rt = $cell.RichText
    if ($rt.Count -gt 0) { $rt.Clear() }
    $titlePart = $rt.Add("$Title`r`n")
    $titlePart.Size = 22
    $titlePart.Bold = $true
    $titlePart.Color = (Get-ThemeColor White)
    if ($Subtitle) {
        $subPart = $rt.Add($Subtitle)
        $subPart.Size = 10
        $subPart.Color = [System.Drawing.Color]::FromArgb(186, 204, 224)
    }

    $subLines = if ($Subtitle) { [Math]::Ceiling($Subtitle.Length / 110.0) } else { 0 }
    $Worksheet.Row($r1).Height = [Math]::Max(44, 28 + (12 * $subLines))
    $Worksheet.Row($r2).Height = 6

    $accent = Get-ExcelRange -Worksheet $Worksheet -FromRow $r3 -FromCol $c1 -ToRow $r3 -ToCol $cEnd
    Merge-ExcelRange -Range $accent -Worksheet $Worksheet -TopRow $r3 -LeftCol $c1 | Out-Null
    Set-RangeFill -Range $accent -Color (Get-ThemeColor Accent)
    $Worksheet.Row($r3).Height = 5
}

function Format-SummaryKpiRow {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    # KPIs live in the insights band (cols 5–10) so they never overlap the narrow table/gutter columns.
    $kpiRow1 = Get-SummaryRow 5
    $kpiRow2 = Get-SummaryRow 6
    $totalSites = $Stats.RootSites + $Stats.Subsites
    $uniqueItems = if ($Stats.UniqueItemPaths) { $Stats.UniqueItemPaths.Count } else { 0 }

    Write-KpiCard -Worksheet $Worksheet -Row $kpiRow1 -ColStart (Get-SummaryCol 5) -ColEnd (Get-SummaryCol 7) `
        -Label 'Sites' -Value $totalSites -AccentColor (Get-ThemeColor Accent)
    Write-KpiCard -Worksheet $Worksheet -Row $kpiRow1 -ColStart (Get-SummaryCol 8) -ColEnd (Get-SummaryCol 10) `
        -Label 'Items scanned' -Value $Stats.ItemsScanned -AccentColor (Get-ThemeColor Accent2)
    Write-KpiCard -Worksheet $Worksheet -Row $kpiRow2 -ColStart (Get-SummaryCol 5) -ColEnd (Get-SummaryCol 7) `
        -Label 'External users' -Value (Get-TotalExternalUserCount) `
        -AccentColor ([System.Drawing.Color]::FromArgb(234, 88, 12))
    Write-KpiCard -Worksheet $Worksheet -Row $kpiRow2 -ColStart (Get-SummaryCol 8) -ColEnd (Get-SummaryCol 10) `
        -Label 'Unique permissions' -Value $uniqueItems `
        -AccentColor ([System.Drawing.Color]::FromArgb(217, 119, 6))

    $detailsBand = Get-ExcelRange -Worksheet $Worksheet -FromRow $kpiRow1 -FromCol (Get-SummaryCol 1) `
        -ToRow $kpiRow2 -ToCol (Get-SummaryDetailsEndCol)
    Set-RangeFill -Range $detailsBand -Color (Get-ThemeColor Surface)

    $gutter = Get-ExcelRange -Worksheet $Worksheet -FromRow $kpiRow1 -FromCol (Get-SummaryCol $SummaryGutterLogicalCol) `
        -ToRow $kpiRow2 -ToCol (Get-SummaryCol $SummaryGutterLogicalCol)
    Set-RangeFill -Range $gutter -Color (Get-ThemeColor Surface)

    $spacer = Get-ExcelRange -Worksheet $Worksheet -FromRow (Get-SummaryRow 7) -FromCol (Get-SummaryCol 1) `
        -ToRow (Get-SummaryRow 7) -ToCol (Get-SummaryInsightsEndCol)
    Set-RangeFill -Range $spacer -Color (Get-ThemeColor Surface)
    $Worksheet.Row((Get-SummaryRow 7)).Height = 10
}

function Format-SummaryOverviewTable {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $tbl = $Worksheet.Tables['SummaryOverview']
    if (-not $tbl) { return }

    $tbl.ShowFilter = $false
    $tbl.ShowRowStripes = $false
    $tbl.ShowColumnStripes = $false
    $tbl.TableStyle = [OfficeOpenXml.Table.TableStyles]::None

    $r1 = $tbl.Address.Start.Row + 1
    $r2 = $tbl.Address.End.Row
    $cols = Get-TableColumnSpan $tbl
    $secCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Section'
    $metricCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Metric'
    $valueCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Value'
    if ($secCol -lt 1 -or $metricCol -lt 1 -or $valueCol -lt 1) { return }

    $cardTop = $tbl.Address.Start.Row - 1
    $detailsEnd = Get-SummaryDetailsEndCol
    $insightsStart = Get-SummaryInsightsStartCol
    $insightsEnd = Get-SummaryInsightsEndCol

    if ($cardTop -eq ($SummaryOverviewStartRow - 1)) {
        $leftHdr = Get-ExcelRange -Worksheet $Worksheet -FromRow $cardTop -FromCol $cols.Start -ToRow $cardTop -ToCol $detailsEnd
        Merge-ExcelRange -Range $leftHdr -Worksheet $Worksheet -TopRow $cardTop -LeftCol $cols.Start | Out-Null
        Set-ExcelCellValue -Worksheet $Worksheet -Row $cardTop -Col $cols.Start -Value 'SCAN DETAILS'
        $leftHdr.Style.Font.Size = 8
        $leftHdr.Style.Font.Bold = $true
        $leftHdr.Style.Font.Color.SetColor((Get-ThemeColor Muted))
        $leftHdr.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        $leftHdr.Style.Indent = 1

        $rightHdr = Get-ExcelRange -Worksheet $Worksheet -FromRow $cardTop -FromCol $insightsStart -ToRow $cardTop -ToCol $insightsEnd
        Merge-ExcelRange -Range $rightHdr -Worksheet $Worksheet -TopRow $cardTop -LeftCol $insightsStart | Out-Null
        Set-ExcelCellValue -Worksheet $Worksheet -Row $cardTop -Col $insightsStart -Value 'INSIGHTS'
        $rightHdr.Style.Font.Size = 8
        $rightHdr.Style.Font.Bold = $true
        $rightHdr.Style.Font.Color.SetColor((Get-ThemeColor Muted))
        $rightHdr.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        $rightHdr.Style.Indent = 1
        $Worksheet.Row($cardTop).Height = 20
    }

    $hdr = $tbl.Address.Start.Row
    Format-TableHeaderRow -Worksheet $Worksheet -Table $tbl

    $card = Get-ExcelRange -Worksheet $Worksheet -FromRow $hdr -FromCol $cols.Start -ToRow $r2 -ToCol $cols.End
    Set-RangeBorder -Range $card -Style ([OfficeOpenXml.Style.ExcelBorderStyle]::Thin) -Color (Get-ThemeColor BorderLight)

    $colorPanel = Get-ThemeColor Panel
    $colorPrimarySoft = Get-ThemeColor PrimarySoft
    $colorPrimaryMid = Get-ThemeColor PrimaryMid
    $colorText = Get-ThemeColor Text
    $colorBorder = Get-ThemeColor Border
    $colorBorderLight = Get-ThemeColor BorderLight

    $bodyRange = Get-ExcelRange -Worksheet $Worksheet -FromRow $r1 -FromCol $cols.Start -ToRow $r2 -ToCol $cols.End
    Set-RangeFill -Range $bodyRange -Color $colorPanel
    $bodyRange.Style.Font.Name = 'Segoe UI'

    $metricRange = Get-ExcelRange -Worksheet $Worksheet -FromRow $r1 -FromCol $metricCol -ToRow $r2 -ToCol $metricCol
    $metricRange.Style.Font.Color.SetColor($colorText)
    $metricRange.Style.Font.Size = 10

    $valueRange = Get-ExcelRange -Worksheet $Worksheet -FromRow $r1 -FromCol $valueCol -ToRow $r2 -ToCol $valueCol
    $valueRange.Style.Font.Bold = $true
    $valueRange.Style.Font.Size = 10
    $valueRange.Style.Font.Color.SetColor($colorPrimaryMid)
    $valueRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
    $valueRange.Style.Indent = 1

    $profileMetricCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Metric'
    $profileRows = [System.Collections.Generic.List[int]]::new()
    if ($profileMetricCol -gt 0) {
        for ($pr = $r1; $pr -le $r2; $pr++) {
            $pm = [string]$Worksheet.Cells[$pr, $profileMetricCol].Value
            if ($pm -eq 'Toolkit connection profile') { [void]$profileRows.Add($pr) }
        }
    }

    $prevSec = ''
    for ($r = $r1; $r -le $r2; $r++) {
        if ($profileRows.Contains($r)) {
            $Worksheet.Row($r).Hidden = $true
            continue
        }

        $sec = [string]$Worksheet.Cells[$r, $secCol].Value
        $rowRange = Get-ExcelRange -Worksheet $Worksheet -FromRow $r -FromCol $cols.Start -ToRow $r -ToCol $cols.End

        if ($sec -ne $prevSec) {
            if ($prevSec -and $r -gt $r1) {
                $rowRange.Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $rowRange.Style.Border.Top.Color.SetColor($colorBorder)
            }
            $prevSec = $sec
            Set-RangeFill -Range $rowRange -Color $colorPrimarySoft
            $Worksheet.Cells[$r, $secCol].Style.Font.Bold = $true
            $Worksheet.Cells[$r, $secCol].Style.Font.Size = 8
            $Worksheet.Cells[$r, $secCol].Style.Font.Color.SetColor($colorPrimaryMid)
        }
        else {
            Set-ExcelCellValue -Worksheet $Worksheet -Row $r -Col $secCol -Value ''
            $rowRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
            $rowRange.Style.Border.Bottom.Color.SetColor($colorBorderLight)
        }

        $Worksheet.Row($r).Height = 20
    }
}

function Format-SiteBreakdownTable {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $tbl = $Worksheet.Tables['SiteBreakdown']
    if (-not $tbl) { return }

    $hdr = $tbl.Address.Start.Row
    $tblSpan = Get-TableColumnSpan $tbl
    $titleRow = $hdr - 1
    if ($titleRow -ge 1) {
        $secEnd = [Math]::Min($SummarySheetLastCol, [Math]::Max((Get-SummaryCol 8), $tblSpan.End))
        $titleCol = Get-SummaryCol 1
        $sec = Get-ExcelRange -Worksheet $Worksheet -FromRow $titleRow -FromCol $titleCol -ToRow $titleRow -ToCol $secEnd
        if (-not $sec.Merge) {
            Merge-ExcelRange -Range $sec -Worksheet $Worksheet -TopRow $titleRow -LeftCol $titleCol | Out-Null
        }
        Set-ExcelCellValue -Worksheet $Worksheet -Row $titleRow -Col $titleCol -Value 'SITE BREAKDOWN'
        $sec.Style.Font.Size = 9
        $sec.Style.Font.Bold = $true
        $sec.Style.Font.Color.SetColor((Get-ThemeColor Muted))
        $sec.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        $Worksheet.Row($titleRow).Height = 18
    }

    $tbl.ShowFilter = $true
    $tbl.ShowRowStripes = $false
    $tbl.ShowColumnStripes = $false
    $tbl.TableStyle = [OfficeOpenXml.Table.TableStyles]::None

    $cols = Get-TableColumnSpan $tbl
    Format-TableHeaderRow -Worksheet $Worksheet -Table $tbl

    $r1 = $hdr + 1
    $r2 = $tbl.Address.End.Row
    $dataRange = Get-ExcelRange -Worksheet $Worksheet -FromRow $r1 -FromCol $cols.Start -ToRow $r2 -ToCol $cols.End
    Set-RangeFill -Range $dataRange -Color (Get-ThemeColor Panel)
    Set-RangeBorder -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $hdr -FromCol $cols.Start -ToRow $r2 -ToCol $cols.End) `
        -Style ([OfficeOpenXml.Style.ExcelBorderStyle]::Thin) -Color (Get-ThemeColor BorderLight)

    for ($r = $r1; $r -le $r2; $r++) {
        $Worksheet.Row($r).Height = 18
    }

    Format-ExcelTableColumns -Worksheet $Worksheet -Table $tbl -SampleRows 50 -Profiles @{
        'Site Name'       = @{ Min = 10; Max = 22 }
        'Site Path'       = @{ Min = 26; Max = 58; Wrap = $true }
        'Type'            = @{ Min = 8; Max = 12 }
        'Lists/Libraries' = @{ Min = 26; Max = 42 }
        'Items'           = @{ Min = 8; Max = 12 }
        'Files'           = @{ Min = 8; Max = 12 }
        'Folders'         = @{ Min = 8; Max = 12 }
        'List items'      = @{ Min = 10; Max = 14 }
        'Unique items'      = @{ Min = 12; Max = 14 }
        'External users'    = @{ Min = 12; Max = 14 }
        'Site unique perms' = @{ Min = 14; Max = 18 }
    }

    $listsCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Lists/Libraries'
    if ($listsCol -gt 0 -and $Worksheet.Column($listsCol).Width -lt 26) {
        $Worksheet.Column($listsCol).Width = 26
    }
}

function Format-TableHeaderRow {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [OfficeOpenXml.Table.ExcelTable] $Table
    )

    if (-not $Table) { return }

    $cols = Get-TableColumnSpan $Table
    $hdr = $Table.Address.Start.Row
    $c1 = $cols.Start
    $c2 = $cols.End
    if ($c1 -lt 1 -or $c2 -lt 1 -or $hdr -lt 1) { return }

    $range = Get-ExcelRange -Worksheet $Worksheet -FromRow $hdr -FromCol $c1 -ToRow $hdr -ToCol $c2
    $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $range.Style.Fill.BackgroundColor.SetColor((Get-ThemeColor Primary))
    $range.Style.Font.Bold = $true
    $range.Style.Font.Color.SetColor((Get-ThemeColor White))
    $range.Style.Font.Size = 10
    $range.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    $range.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $range.Style.WrapText = $false
    $range.Style.Indent = 1
    $Worksheet.Row($hdr).Height = 28
}

function Get-TableColumnIndex {
    param(
        [OfficeOpenXml.Table.ExcelTable] $Table,
        [string] $ColumnName
    )

    $cacheKey = "$($Table.Name)|$ColumnName"
    if ($script:TableColumnIndexCache.ContainsKey($cacheKey)) {
        return $script:TableColumnIndexCache[$cacheKey]
    }

    $ws = $Table.Worksheet
    $hdr = $Table.Address.Start.Row
    if ($hdr -lt 1) { return -1 }
    $cols = Get-TableColumnSpan $Table
    for ($c = $cols.Start; $c -le $cols.End; $c++) {
        if ([string]$ws.Cells[$hdr, $c].Value -eq $ColumnName) {
            $script:TableColumnIndexCache[$cacheKey] = $c
            return $c
        }
    }
    $script:TableColumnIndexCache[$cacheKey] = -1
    return -1
}

function Format-DataTableSheet {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [string] $TableName,
        [string[]] $RoleColumns = @(),
        [switch] $ApplyRoleHighlight,
        [switch] $ApplyInheritanceColors
    )

    $tbl = $Worksheet.Tables[$TableName]
    if (-not $tbl) { return }

    $tbl.ShowRowStripes = $false
    $tbl.ShowColumnStripes = $false
    $tbl.TableStyle = [OfficeOpenXml.Table.TableStyles]::None
    Format-TableHeaderRow -Worksheet $Worksheet -Table $tbl

    $cols = Get-TableColumnSpan $tbl
    $hdr = $tbl.Address.Start.Row
    $r1 = $hdr + 1
    $r2 = $tbl.Address.End.Row
    if ($hdr -ge 1) {
        $Worksheet.View.FreezePanes($hdr + 1, $cols.Start)
    }

    if ($r2 -ge $r1 -and ($r2 - $hdr) -le $StyleRowLimit) {
        Set-RangeBorder -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $hdr -FromCol $cols.Start -ToRow $r2 -ToCol $cols.End) `
            -Color (Get-ThemeColor Border)
        for ($r = $r1; $r -le $r2; $r++) {
            $fill = if ($r % 2 -eq 0) { (Get-ThemeColor White) } else { (Get-ThemeColor SurfaceAlt) }
            Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $r -FromCol $cols.Start -ToRow $r -ToCol $cols.End) -Color $fill
            $Worksheet.Row($r).Height = 17
        }
    }

    $rowCount = $r2 - $hdr
    if ($rowCount -gt $StyleRowLimit) { return }

    if ($ApplyInheritanceColors) {
        $inhCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Inheritance'
        if ($inhCol -gt 0) {
            for ($r = $hdr + 1; $r -le $tbl.Address.End.Row; $r++) {
                $val = [string]$Worksheet.Cells[$r, $inhCol].Value
                $color = switch ($val) {
                    'Custom'    { Get-ThemeColor Custom }
                    'Inherited' { Get-ThemeColor Inherited }
                    'Top level' { Get-ThemeColor TopLevel }
                    default     { $null }
                }
                if ($color) {
                    $Worksheet.Cells[$r, $inhCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Worksheet.Cells[$r, $inhCol].Style.Fill.BackgroundColor.SetColor($color)
                }
            }
        }

        $extCol = Get-TableColumnIndex -Table $tbl -ColumnName 'External user'
        if ($extCol -gt 0) {
            for ($r = $hdr + 1; $r -le $tbl.Address.End.Row; $r++) {
                if ([string]$Worksheet.Cells[$r, $extCol].Value -eq 'Yes') {
                    $Worksheet.Cells[$r, $extCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Worksheet.Cells[$r, $extCol].Style.Fill.BackgroundColor.SetColor((Get-ThemeColor Warning))
                    $Worksheet.Cells[$r, $extCol].Style.Font.Bold = $true
                }
            }
        }
    }

    if ($ApplyRoleHighlight -and $RoleColumns.Count -gt 0) {
        foreach ($roleName in $RoleColumns) {
            $col = Get-TableColumnIndex -Table $tbl -ColumnName $roleName
            if ($col -lt 1) { continue }
            if ($hdr + 1 -gt $tbl.Address.End.Row) { continue }
            try {
                $addr = (Get-ExcelRange -Worksheet $Worksheet -FromRow ($hdr + 1) -FromCol $col `
                    -ToRow $tbl.Address.End.Row -ToCol $col).Address
                Add-ConditionalFormatting -Worksheet $Worksheet -Range $addr -RuleType Equal `
                    -ConditionValue 'X' -BackgroundColor (Get-ThemeColor Highlight) `
                    -Bold -HorizontalAlignment Center
            } catch {
                Write-Verbose "Role highlight skipped for column $roleName : $_"
            }
        }
    }
}

function Get-ExternalUserCountBySite {
    param([string] $SiteName)

    if ($Stats.ExternalUsersBySite.ContainsKey($SiteName)) {
        return $Stats.ExternalUsersBySite[$SiteName].Count
    }
    return 0
}

function Get-TotalExternalUserCount {
    $all = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($siteName in @($Stats.ExternalUsersBySite.Keys)) {
        foreach ($login in @($Stats.ExternalUsersBySite[$siteName])) {
            if ($login) { [void]$all.Add([string]$login) }
        }
    }
    return $all.Count
}

function Register-ExternalUserForSite {
    param([string] $SiteName, [string] $Login)

    if ([string]::IsNullOrWhiteSpace($Login)) { return }
    if (-not $Stats.ExternalUsersBySite.ContainsKey($SiteName)) {
        $Stats.ExternalUsersBySite[$SiteName] = [System.Collections.Generic.HashSet[string]]::new(
            [StringComparer]::OrdinalIgnoreCase)
    }
    [void]$Stats.ExternalUsersBySite[$SiteName].Add($Login)
}

function Set-SummaryChartStyle {
    param(
        [OfficeOpenXml.Drawing.Chart.ExcelChart] $Chart,
        [System.Drawing.Color] $BarColor
    )

    try {
        $Chart.Title.Font.Name = 'Segoe UI'
        $Chart.Title.Font.Color.SetColor((Get-ThemeColor Text))
    }
    catch { }
    try { $Chart.Legend.Remove() } catch { }
    try {
        if ($Chart.Legend) {
            $Chart.Legend.Position = [OfficeOpenXml.Drawing.Chart.eLegendPosition]::Top
            $Chart.Legend.Remove()
        }
    }
    catch { }

    $chartBg = Get-ThemeColor Surface
    Set-ExcelDrawingFill -Fill $Chart.Border.Fill -Color $chartBg -Style 'NoFill'
    Set-ExcelDrawingFill -Fill $Chart.Fill -Color $chartBg -Style 'SolidFill'
    Set-ExcelDrawingFill -Fill $Chart.PlotArea.Fill -Color $chartBg -Style 'SolidFill'
    Set-ExcelDrawingFill -Fill $Chart.PlotArea.Border.Fill -Color $chartBg -Style 'NoFill'

    try { $Chart.RoundedCorners = $false } catch { }

    Remove-SummaryChartGridlines -Chart $Chart

    try {
        $Chart.XAxis.Font.Size = 8
        $Chart.XAxis.Font.Color.SetColor((Get-ThemeColor Muted))
        $Chart.YAxis.Font.Size = 9
        $Chart.YAxis.Font.Color.SetColor((Get-ThemeColor Text))
    }
    catch { }

    try {
        if ($Chart.Series.Count -gt 0) {
            Set-ExcelDrawingFill -Fill $Chart.Series[0].Fill -Color $BarColor -Style 'SolidFill'
            $Chart.Series[0].Header = ''
        }
    }
    catch { }

    try { $Chart.ShowHiddenData = $true } catch { }
}

function Set-SummaryChartValueScale {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [OfficeOpenXml.Drawing.Chart.ExcelChart] $Chart,
        [int] $DataStartRow,
        [int] $DataEndRow,
        [int] $ValueCol
    )

    $maxVal = 0.0
    for ($r = $DataStartRow; $r -le $DataEndRow; $r++) {
        $raw = $Worksheet.Cells[$r, $ValueCol].Value
        if ($null -eq $raw) { continue }
        try { $maxVal = [Math]::Max($maxVal, [double]$raw) } catch { }
    }
    $axisMax = [Math]::Max(1.0, [Math]::Ceiling($maxVal * 1.25))

    foreach ($axisName in @('XAxis', 'YAxis')) {
        try {
            $axis = $Chart.$axisName
            $axis.Scaling.MinValue = 0
            $axis.Scaling.MaxValue = $axisMax
        }
        catch { }
    }
}

function Get-SummaryChartType {
    return [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered
}

function Add-SummaryBarChart {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [string] $Name,
        [string] $Title,
        [int] $DataStartRow,
        [int] $DataEndRow,
        [int] $LabelCol,
        [int] $ValueCol,
        [int] $AnchorRow,
        [int] $AnchorCol,
        [System.Drawing.Color] $BarColor,
        [int] $WidthPx = 360,
        [int] $HeightPx = 140
    )

    if ($DataEndRow -lt $DataStartRow) { return $null }

    $values = Get-ExcelRange -Worksheet $Worksheet -FromRow $DataStartRow -FromCol $ValueCol `
        -ToRow $DataEndRow -ToCol $ValueCol
    $labels = Get-ExcelRange -Worksheet $Worksheet -FromRow $DataStartRow -FromCol $LabelCol `
        -ToRow $DataEndRow -ToCol $LabelCol
    if (-not $values -or -not $labels) { return $null }

    $chart = $Worksheet.Drawings.AddChart($Name, (Get-SummaryChartType))
    try { $chart.Direction = [OfficeOpenXml.Drawing.Chart.eDirection]::Bar } catch { }
    $chart.Title.Text = $Title
    $chart.Title.Font.Size = 10
    $chart.Title.Font.Bold = $true
    $null = $chart.Series.Add($values, $labels)
    Set-SummaryChartStyle -Chart $chart -BarColor $BarColor
    Set-SummaryChartValueScale -Worksheet $Worksheet -Chart $chart `
        -DataStartRow $DataStartRow -DataEndRow $DataEndRow -ValueCol $ValueCol
    $chart.SetPosition([Math]::Max(0, $AnchorRow - 1), 4, [Math]::Max(0, $AnchorCol - 1), 4)
    $chart.SetSize($WidthPx, $HeightPx)
    return $chart
}

function Get-ShortChartLabel {
    param([string] $Text)

    $t = [string]$Text
    if ($t.Length -le 22) { return $t }
    return ($t.Substring(0, 20).TrimEnd() + [char]0x2026)
}

function Get-ChartPointsForSites {
    param(
        [object[]] $Sites,
        [string] $ValueProperty,
        [int] $MaxPoints = 12
    )

    $pts = foreach ($s in $Sites) {
        [pscustomobject]@{
            Label = Get-ShortChartLabel -Text ([string]$s.'Site Name')
            Value = [double]$s.$ValueProperty
        }
    }
    return @($pts | Sort-Object -Property Value -Descending | Select-Object -First $MaxPoints)
}

function Write-SummaryChartDataBlock {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $StartRow,
        [int] $LabelCol,
        [int] $ValueCol,
        [object[]] $Points
    )

    $row = $StartRow
    foreach ($pt in @($Points)) {
        if ($null -eq $pt) { continue }
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col $LabelCol -Value ([string]$pt.Label)
        Set-ExcelCellValue -Worksheet $Worksheet -Row $row -Col $ValueCol -Value ([double]$pt.Value)
        $row++
    }
    return @{ Start = $StartRow; End = ($row - 1) }
}

function Clear-SummaryInsightsArea {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [int] $TopRow,
        [int] $BottomRow,
        [int] $ColStart,
        [int] $ColEnd
    )

    if ($BottomRow -lt $TopRow) { return }
    try {
        $area = Get-ExcelRange -Worksheet $Worksheet -FromRow $TopRow -FromCol $ColStart -ToRow $BottomRow -ToCol $ColEnd
        if ($area) {
            $area.Merge = $false
            $area.Clear()
        }
    }
    catch { }
}

function Format-SummaryInsightsPanel {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    Clear-SummaryDrawings -Worksheet $Worksheet

    $sites = @($Stats.SiteDetails)
    $tbl = $Worksheet.Tables['SummaryOverview']
    if ($sites.Count -lt 1 -or -not $tbl) { return }

    $insightsStart = Get-SummaryInsightsStartCol
    $insightsEnd = Get-SummaryInsightsEndCol
    $panelTop = $tbl.Address.Start.Row
    $panelBottom = $tbl.Address.End.Row

    Clear-SummaryInsightsArea -Worksheet $Worksheet -TopRow $panelTop -BottomRow $panelBottom `
        -ColStart $insightsStart -ColEnd $insightsEnd

    $panel = Get-ExcelRange -Worksheet $Worksheet -FromRow $panelTop -FromCol $insightsStart `
        -ToRow $panelBottom -ToCol $insightsEnd
    Set-RangeFill -Range $panel -Color (Get-ThemeColor Surface)

    $gutterCol = Get-SummaryCol $SummaryGutterLogicalCol
    Set-RangeFill -Range (Get-ExcelRange -Worksheet $Worksheet -FromRow $panelTop -FromCol $gutterCol `
        -ToRow $panelBottom -ToCol $gutterCol) -Color (Get-ThemeColor Surface)

    $labelCol     = Get-SummaryCol $SummaryChartDataLabelLogicalCol
    $uniqueValCol = Get-SummaryCol $SummaryChartDataValueLogicalCol
    $dataStartRow = [Math]::Max($SummaryChartDataMinRow, $panelBottom + 5)

    $uniquePoints = Get-ChartPointsForSites -Sites $sites -ValueProperty 'Unique items' -MaxPoints 10
    $uniqueRange  = Write-SummaryChartDataBlock -Worksheet $Worksheet -StartRow $dataStartRow `
        -LabelCol $labelCol -ValueCol $uniqueValCol -Points $uniquePoints

    $itemPoints = Get-ChartPointsForSites -Sites $sites -ValueProperty 'Items' -MaxPoints 10
    $itemRange  = Write-SummaryChartDataBlock -Worksheet $Worksheet -StartRow ($uniqueRange.End + 3) `
        -LabelCol $labelCol -ValueCol $uniqueValCol -Points $itemPoints

    $totalExternal = Get-TotalExternalUserCount
    $extPoints = [System.Collections.Generic.List[object]]::new()
    $hasExternal = $totalExternal -gt 0
    foreach ($s in $sites) {
        $ext = Get-ExternalUserCountBySite -SiteName ([string]$s.'Site Name')
        if ($ext -lt 1) { continue }
        $hasExternal = $true
        [void]$extPoints.Add([pscustomobject]@{
            Label = Get-ShortChartLabel -Text ([string]$s.'Site Name')
            Value = [double]$ext
        })
    }
    $extPointsArr = @($extPoints | Sort-Object -Property Value -Descending | Select-Object -First 10)
    if ($hasExternal -and $extPointsArr.Count -lt 1) {
        $extPointsArr = @([pscustomobject]@{ Label = 'Total (distinct)'; Value = [double]$totalExternal })
    }
    $extRange = $null
    if ($extPointsArr.Count -gt 0) {
        $extRange = Write-SummaryChartDataBlock -Worksheet $Worksheet `
            -StartRow ($itemRange.End + 3) -LabelCol $labelCol -ValueCol $uniqueValCol -Points $extPointsArr
    }

    $chartW = 520

    # Divide the panel into equal segments — one per chart.
    $numCharts = if ($extPointsArr.Count -gt 0) { 3 } else { 2 }
    $panelRows = [Math]::Max($numCharts * 8, $panelBottom - $panelTop + 1)
    $segment   = [int][Math]::Floor($panelRows / $numCharts)

    # Sample an actual data-row height (EPPlus returns points; 1pt ≈ 4/3 px at 96 DPI).
    # This avoids large gaps caused by rows being taller than the assumed 15 px.
    $sampleRow = [Math]::Min($panelTop + 5, $panelBottom)
    $rowHPt    = try { $Worksheet.Row($sampleRow).Height } catch { 15.0 }
    if ($rowHPt -le 0) { $rowHPt = 15.0 }
    $rowPx     = $rowHPt * 4.0 / 3.0
    $chartH    = [Math]::Min(360, [Math]::Max(80, [int]($segment * $rowPx) - 24))

    $chartRow1 = $panelTop + 1
    $chartRow2 = $panelTop + $segment + 1
    $chartRow3 = if ($numCharts -eq 3) { $panelTop + 2 * $segment + 1 } else { 0 }

    if ($uniqueRange.End -ge $uniqueRange.Start) {
        try {
            $null = Add-SummaryBarChart -Worksheet $Worksheet -Name 'UniqueBySite' `
                -Title 'Unique permissions by site' `
                -DataStartRow $uniqueRange.Start -DataEndRow $uniqueRange.End `
                -LabelCol $labelCol -ValueCol $uniqueValCol `
                -AnchorRow $chartRow1 -AnchorCol $insightsStart `
                -BarColor (Get-ThemeColor ChartBar) -WidthPx $chartW -HeightPx $chartH
        }
        catch { Write-Warning "Unique permissions chart failed: $_" }
    }

    if ($itemRange.End -ge $itemRange.Start) {
        try {
            $null = Add-SummaryBarChart -Worksheet $Worksheet -Name 'ItemsBySite' `
                -Title 'Items by site' `
                -DataStartRow $itemRange.Start -DataEndRow $itemRange.End `
                -LabelCol $labelCol -ValueCol $uniqueValCol `
                -AnchorRow $chartRow2 -AnchorCol $insightsStart `
                -BarColor (Get-ThemeColor ChartBar) -WidthPx $chartW -HeightPx $chartH
        }
        catch { Write-Warning "Items by site chart failed: $_" }
    }

    if ($hasExternal -and $extRange -and $extRange.End -ge $extRange.Start) {
        try {
            $null = Add-SummaryBarChart -Worksheet $Worksheet -Name 'ExternalBySite' `
                -Title 'External users by site' `
                -DataStartRow $extRange.Start -DataEndRow $extRange.End `
                -LabelCol $labelCol -ValueCol $uniqueValCol `
                -AnchorRow $chartRow3 -AnchorCol $insightsStart `
                -BarColor (Get-ThemeColor ChartBar2) -WidthPx $chartW -HeightPx $chartH
        }
        catch { Write-Warning "External users chart failed: $_" }
    }
    elseif ($uniqueRange.End -ge $uniqueRange.Start) {
        $noteRow = $panelBottom
        $noteRange = Get-ExcelRange -Worksheet $Worksheet -FromRow $noteRow -FromCol $insightsStart `
            -ToRow $noteRow -ToCol $insightsEnd
        Merge-ExcelRange -Range $noteRange -Worksheet $Worksheet -TopRow $noteRow -LeftCol $insightsStart | Out-Null
        Set-ExcelCellValue -Worksheet $Worksheet -Row $noteRow -Col $insightsStart `
            -Value 'No external users detected in this scan.'
        $noteRange.Style.Font.Name = 'Segoe UI'
        $noteRange.Style.Font.Size = 9
        $noteRange.Style.Font.Italic = $true
        $noteRange.Style.Font.Color.SetColor((Get-ThemeColor Muted))
        $noteRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $noteRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $Worksheet.Row($noteRow).Height = 22
    }
}

function Set-SummaryOverviewValueFormats {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $tbl = $Worksheet.Tables['SummaryOverview']
    if (-not $tbl) { return }

    $valueCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Value'
    if ($valueCol -lt 1) { return }

    $hdr = [int]$tbl.Address.Start.Row
    for ($r = $hdr + 1; $r -le [int]$tbl.Address.End.Row; $r++) {
        $raw = $Worksheet.Cells[$r, $valueCol].Value
        if ($null -eq $raw) { continue }
        if ($raw -is [double] -or $raw -is [int] -or $raw -is [long] -or $raw -is [decimal]) {
            $Worksheet.Cells[$r, $valueCol].Style.Numberformat.Format = '#,##0'
            continue
        }
        $text = [string]$raw
        if ($text -match '^\d{1,3}(,\d{3})*$') {
            $num = 0.0
            if ([double]::TryParse($text.Replace(',', ''), [ref]$num)) {
                $Worksheet.Cells[$r, $valueCol].Value = $num
                $Worksheet.Cells[$r, $valueCol].Style.Numberformat.Format = '#,##0'
            }
        }
    }
}

function Format-SummaryLayout {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $ov = $Worksheet.Tables['SummaryOverview']
    if ($ov) {
        Set-SummaryOverviewValueFormats -Worksheet $Worksheet
        $valueCol = Get-TableColumnIndex -Table $ov -ColumnName 'Value'
        if ($valueCol -gt 0) {
            $hdr = [int]$ov.Address.Start.Row
            $valRange = Get-ExcelRange -Worksheet $Worksheet -FromRow ($hdr + 1) -FromCol $valueCol `
                -ToRow ([int]$ov.Address.End.Row) -ToCol $valueCol
            if ($valRange) { $valRange.Style.WrapText = $false }
        }
    }

    $sb = $Worksheet.Tables['SiteBreakdown']
    if ($sb) {
        Format-ExcelTableColumns -Worksheet $Worksheet -Table $sb -SampleRows 100 -Profiles @{
            'Site Name'         = @{ Min = 14; Max = 24 }
            'Site Path'         = @{ Min = 42; Max = 72; Wrap = $true }
            'Type'              = @{ Min = 10; Max = 14 }
            'Lists/Libraries'   = @{ Min = 30; Max = 46 }
            'Items'             = @{ Min = 8; Max = 12 }
            'Files'             = @{ Min = 8; Max = 12 }
            'Folders'           = @{ Min = 8; Max = 12 }
            'List items'        = @{ Min = 10; Max = 14 }
            'Unique items'      = @{ Min = 12; Max = 14 }
            'External users'    = @{ Min = 12; Max = 14 }
            'Site unique perms' = @{ Min = 14; Max = 18 }
        }
        $sbListsCol = Get-TableColumnIndex -Table $sb -ColumnName 'Lists/Libraries'
        if ($sbListsCol -gt 0 -and $Worksheet.Column($sbListsCol).Width -lt 30) {
            $Worksheet.Column($sbListsCol).Width = 30
        }
    }

    # A–B=margin, C=section, D=metric, E=value, F=gutter, G–L=insights/KPI, M–N=chart data
    $Worksheet.Column(1).Width = 10
    $Worksheet.Column(2).Width = 10
    $Worksheet.Column(3).Width = 0.6
    $Worksheet.Column(4).Width = 40
    $Worksheet.Column(5).Width = 26
    $Worksheet.Column(6).Width = 2
    $Worksheet.Column(7).Width = 16
    $Worksheet.Column(8).Width = 16
    $Worksheet.Column(9).Width = 16
    $Worksheet.Column(10).Width = 16
    $Worksheet.Column(11).Width = 16
    $Worksheet.Column(12).Width = 16
    $Worksheet.Column(13).Width = 12
    $Worksheet.Column(14).Width = 10
    try {
        $Worksheet.Column(13).Hidden = $true
        $Worksheet.Column(14).Hidden = $true
    }
    catch { }

    if ($sb) {
        $sbListsCol = Get-TableColumnIndex -Table $sb -ColumnName 'Lists/Libraries'
        if ($sbListsCol -gt 0 -and $Worksheet.Column($sbListsCol).Width -lt 30) {
            $Worksheet.Column($sbListsCol).Width = 30
        }
    }

    Reset-SummaryWorksheetBounds -Worksheet $Worksheet
    Apply-SummaryPagePadding -Worksheet $Worksheet
    Set-SummaryViewSettings -Worksheet $Worksheet
}

function Format-DataTableColumnLayout {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [string] $TableName,
        [string] $ProfileKind,
        [string[]] $RoleColumns = @()
    )

    $tbl = $Worksheet.Tables[$TableName]
    if (-not $tbl) { return }

    $profiles = Get-DataTableColumnProfiles -TableKind $ProfileKind
    foreach ($role in $RoleColumns) {
        if (-not $profiles.ContainsKey($role)) {
            $profiles[$role] = @{
                Min = (Get-HeaderColumnMinWidth -HeaderText $role -ProfileMin 12)
                Max = 20
            }
        }
    }

    $rowCount = [int]$tbl.Address.End.Row - [int]$tbl.Address.Start.Row
    $sample = if ($rowCount -gt 5000) { 400 } elseif ($rowCount -gt 1000) { 250 } else { 0 }
    Format-ExcelTableColumns -Worksheet $Worksheet -Table $tbl -Profiles $profiles -SampleRows $sample

    $hdr = [int]$tbl.Address.Start.Row
    $Worksheet.Row($hdr).Height = 28
}

function Format-SharingLinksSheet {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $tbl = $Worksheet.Tables['SharingLinks']
    if (-not $tbl) { return }

    Format-DataTableSheet -Worksheet $Worksheet -TableName 'SharingLinks'

    $typeCol = Get-TableColumnIndex -Table $tbl -ColumnName 'RecordType'
    if ($typeCol -lt 0) { return }

    $hdr = $tbl.Address.Start.Row
    for ($r = $hdr + 1; $r -le $tbl.Address.End.Row; $r++) {
        $val = [string]$Worksheet.Cells[$r, $typeCol].Value
        $color = if ($val -eq 'SharingLink') { (Get-ThemeColor AccentLight) } else { (Get-ThemeColor SurfaceAlt) }
        $Worksheet.Cells[$r, $typeCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $Worksheet.Cells[$r, $typeCol].Style.Fill.BackgroundColor.SetColor($color)
        $Worksheet.Cells[$r, $typeCol].Style.Font.Bold = $true
    }

    $extCol = Get-TableColumnIndex -Table $tbl -ColumnName 'External user'
    if ($extCol -gt 0) {
        for ($r = $hdr + 1; $r -le $tbl.Address.End.Row; $r++) {
            if ([string]$Worksheet.Cells[$r, $extCol].Value -eq 'Yes') {
                $Worksheet.Cells[$r, $extCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $Worksheet.Cells[$r, $extCol].Style.Fill.BackgroundColor.SetColor((Get-ThemeColor Warning))
                $Worksheet.Cells[$r, $extCol].Style.Font.Bold = $true
            }
        }
    }
}

function Format-AllItemsSheet {
    param([OfficeOpenXml.ExcelWorksheet] $Worksheet)

    $tbl = $Worksheet.Tables['AllItems']
    if (-not $tbl) { return }

    Format-DataTableSheet -Worksheet $Worksheet -TableName 'AllItems'

    $brokenCol = Get-TableColumnIndex -Table $tbl -ColumnName 'Broken permissions'
    if ($brokenCol -lt 1) { return }

    $hdr = [int]$tbl.Address.Start.Row
    for ($r = $hdr + 1; $r -le $tbl.Address.End.Row; $r++) {
        if ([string]$Worksheet.Cells[$r, $brokenCol].Value -ne 'Yes') { continue }
        $Worksheet.Cells[$r, $brokenCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $Worksheet.Cells[$r, $brokenCol].Style.Fill.BackgroundColor.SetColor((Get-ThemeColor Warning))
        $Worksheet.Cells[$r, $brokenCol].Style.Font.Bold = $true
    }
}

function Protect-ToolkitClientId {
    param([string] $Plain)

    if ([string]::IsNullOrWhiteSpace($Plain)) { return '' }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Plain.Trim())
    $key = [System.Text.Encoding]::UTF8.GetBytes('SPOToolkit-v1')
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $bytes[$i] = $bytes[$i] -bxor $key[$i % $key.Length]
    }
    return 'SPOTK1:' + [Convert]::ToBase64String($bytes)
}

function Get-ToolkitMetaSheetName { return '_ToolkitMeta' }

function Set-ReportEmbeddedClientId {
    param(
        [OfficeOpenXml.ExcelPackage] $Package,
        [string] $ClientId
    )

    if ([string]::IsNullOrWhiteSpace($ClientId)) { return $false }

    $enc = Protect-ToolkitClientId $ClientId
    try {
        $sheetName = Get-ToolkitMetaSheetName
        $ws = $Package.Workbook.Worksheets[$sheetName]
        if (-not $ws) {
            $ws = $Package.Workbook.Worksheets.Add($sheetName)
        }
        $ws.Cells['A1'].Value = $enc
        $ws.Hidden = [OfficeOpenXml.eWorkSheetHidden]::VeryHidden
        return $true
    }
    catch {
        Write-Verbose "Could not embed Client ID in workbook: $_"
        return $false
    }
}

function Apply-ModernWorkbookTheme {
    param(
        [OfficeOpenXml.ExcelPackage] $Package,
        [string[]] $RoleColumns,
        [string] $ReportSubtitle = '',
        [int]    $MatrixRowCount = 0,
        [object[]] $SiteDetails = @()
    )

    $pkg = $Package
    $pkg.Workbook.Properties.Title = 'SharePoint Permissions Report'
    $pkg.Workbook.Properties.Category = 'SharePoint Toolkit'
    $pkg.Workbook.Properties.Comments = 'Generated by Export-SitePermissionsMatrix.ps1'

    $wsSummary = $pkg.Workbook.Worksheets['Summary']
    if ($wsSummary) {
        $script:TableColumnIndexCache = @{}
        $siteRows = [object[]]@($SiteDetails)
        if ($siteRows.Count -lt 1) { $siteRows = [object[]]@($Stats.SiteDetails) }

        $summarySteps = @(
            @{ Name = 'bounds'; Script = { Reset-SummaryWorksheetBounds -Worksheet $wsSummary } }
            @{ Name = 'canvas'; Script = {
                Set-WorksheetCanvas -Worksheet $wsSummary -SummaryDashboard
                Apply-SummaryPagePadding -Worksheet $wsSummary
            } }
            @{ Name = 'tab'; Script = { Set-WorksheetTabColor -Worksheet $wsSummary -Color (Get-ThemeColor Accent) } }
            @{ Name = 'hero'; Script = { Format-SummaryHero -Worksheet $wsSummary -Title 'Permissions scan summary' -Subtitle $ReportSubtitle } }
            @{ Name = 'kpi'; Script = { Format-SummaryKpiRow -Worksheet $wsSummary } }
            @{ Name = 'overview'; Script = { Format-SummaryOverviewTable -Worksheet $wsSummary } }
            @{ Name = 'sites-write'; Script = { Write-SiteBreakdownTable -Worksheet $wsSummary -SiteRows $siteRows } }
            @{ Name = 'sites-format'; Script = { Format-SiteBreakdownTable -Worksheet $wsSummary } }
            @{ Name = 'layout'; Script = { Format-SummaryLayout -Worksheet $wsSummary } }
            @{ Name = 'cleanup'; Script = { Clear-SummaryStrayContent -Worksheet $wsSummary } }
            @{ Name = 'insights'; Script = { Format-SummaryInsightsPanel -Worksheet $wsSummary } }
            @{ Name = 'finish'; Script = {
                Apply-SummaryPagePadding -Worksheet $wsSummary
                Trim-SummaryWorksheetExtent -Worksheet $wsSummary
                Complete-SummaryDashboardPresentation -Worksheet $wsSummary
            } }
        )
        foreach ($step in $summarySteps) {
            try { & $step.Script } catch { Write-Warning "Summary $($step.Name) styling skipped: $_" }
        }
        $wsSummary.Cells.Style.Font.Name = 'Segoe UI'
    }

    $wsMatrix = $pkg.Workbook.Worksheets['Permissions Matrix']
    if ($wsMatrix) {
        if ($MatrixRowCount -gt $StyleRowLimit) {
            Write-MatrixLog ("  Large matrix ({0:N0} rows): applying lightweight styling only." -f $MatrixRowCount) -Level Warn
        }
        foreach ($step in @(
            @{ N = 'canvas'; S = { Set-WorksheetCanvas -Worksheet $wsMatrix } }
            @{ N = 'tab'; S = { Set-WorksheetTabColor -Worksheet $wsMatrix -Color (Get-ThemeColor Primary) } }
            @{ N = 'table'; S = { Format-DataTableSheet -Worksheet $wsMatrix -TableName 'PermissionsMatrix' -RoleColumns $RoleColumns -ApplyRoleHighlight -ApplyInheritanceColors } }
            @{ N = 'fit'; S = { Format-DataTableColumnLayout -Worksheet $wsMatrix -TableName 'PermissionsMatrix' -ProfileKind 'Matrix' -RoleColumns $RoleColumns } }
        )) {
            try { & $step.S } catch { Write-Warning "Permissions Matrix $($step.N) styling skipped: $_" }
        }
        Set-WorksheetDefaultFont -Worksheet $wsMatrix -RowCount $MatrixRowCount
    }

    $wsGroups = $pkg.Workbook.Worksheets['Group Members']
    if ($wsGroups) {
        $groupRows = 0
        $dim = $wsGroups.Dimension
        if ($dim) { $groupRows = [int]$dim.End.Row }
        foreach ($step in @(
            @{ N = 'canvas'; S = { Set-WorksheetCanvas -Worksheet $wsGroups } }
            @{ N = 'tab'; S = { Set-WorksheetTabColor -Worksheet $wsGroups -Color ([System.Drawing.Color]::FromArgb(99, 102, 241)) } }
            @{ N = 'table'; S = { Format-DataTableSheet -Worksheet $wsGroups -TableName 'GroupMembers' } }
            @{ N = 'fit'; S = { Format-DataTableColumnLayout -Worksheet $wsGroups -TableName 'GroupMembers' -ProfileKind 'GroupMembers' } }
        )) {
            try { & $step.S } catch { Write-Warning "Group Members $($step.N) styling skipped: $_" }
        }
        Set-WorksheetDefaultFont -Worksheet $wsGroups -RowCount $groupRows
    }

    $wsSharing = $pkg.Workbook.Worksheets['Sharing Links']
    if ($wsSharing) {
        $sharingRows = 0
        $dim = $wsSharing.Dimension
        if ($dim) { $sharingRows = [int]$dim.End.Row }
        foreach ($step in @(
            @{ N = 'canvas'; S = { Set-WorksheetCanvas -Worksheet $wsSharing } }
            @{ N = 'tab'; S = { Set-WorksheetTabColor -Worksheet $wsSharing -Color ([System.Drawing.Color]::FromArgb(234, 88, 12)) } }
            @{ N = 'table'; S = { Format-SharingLinksSheet -Worksheet $wsSharing } }
            @{ N = 'fit'; S = { Format-DataTableColumnLayout -Worksheet $wsSharing -TableName 'SharingLinks' -ProfileKind 'SharingLinks' } }
        )) {
            try { & $step.S } catch { Write-Warning "Sharing Links $($step.N) styling skipped: $_" }
        }
        Set-WorksheetDefaultFont -Worksheet $wsSharing -RowCount $sharingRows
    }

    $wsAllItems = $pkg.Workbook.Worksheets['All Items']
    if ($wsAllItems) {
        $allItemRows = 0
        $dim = $wsAllItems.Dimension
        if ($dim) { $allItemRows = [int]$dim.End.Row }
        foreach ($step in @(
            @{ N = 'canvas'; S = { Set-WorksheetCanvas -Worksheet $wsAllItems } }
            @{ N = 'tab'; S = { Set-WorksheetTabColor -Worksheet $wsAllItems -Color ([System.Drawing.Color]::FromArgb(37, 99, 235)) } }
            @{ N = 'table'; S = { Format-AllItemsSheet -Worksheet $wsAllItems } }
            @{ N = 'fit'; S = { Format-DataTableColumnLayout -Worksheet $wsAllItems -TableName 'AllItems' -ProfileKind 'AllItems' } }
        )) {
            try { & $step.S } catch { Write-Warning "All Items $($step.N) styling skipped: $_" }
        }
        Set-WorksheetDefaultFont -Worksheet $wsAllItems -RowCount $allItemRows
    }

    $pkg.Workbook.Worksheets['Summary'].Select()
}

function Set-WorksheetDefaultFont {
    param(
        [OfficeOpenXml.ExcelWorksheet] $Worksheet,
        [string] $FontName = 'Segoe UI',
        [int]    $RowCount = 0
    )

    # Setting Cells.Style.Font.Name is a single worksheet-level XML attribute —
    # it is always fast regardless of row count, so no row limit guard needed.
    try { $Worksheet.Cells.Style.Font.Name = $FontName } catch { }
}

function Get-MatrixExportRows {
    param(
        [System.Collections.Generic.List[object]] $MatrixRows,
        [string[]] $Columns,
        [int] $RowCount
    )

    if ($RowCount -gt $SortMatrixRowLimit) {
        Write-MatrixLog '  Large matrix: skipping column re-projection to speed up export.' -Level Warn
        return $MatrixRows
    }
    return @($MatrixRows | Select-Object $Columns)
}

function Write-ExportProgress {
    param([string] $Status, [int] $Percent)
    $pct = [math]::Min($Percent, 100)
    if ($script:ProgressRenderState) {
        $script:ProgressRenderState.Activity   = 'Writing Excel workbook'
        $script:ProgressRenderState.BaseStatus = $Status
        $script:ProgressRenderState.Percent    = $pct
        if ($script:RunStartUtc -ne [datetime]::MinValue) {
            $script:ProgressRenderState.StartTicks = $script:RunStartUtc.Ticks
        }
        return
    }
    Write-Progress -Activity 'Writing Excel workbook' -Status $Status -PercentComplete $pct
}

function Write-ExportPhase {
    param(
        [string] $Message,
        [int]    $Percent = 0
    )

    Write-Host ("  {0}..." -f $Message) -ForegroundColor DarkGray
    Write-ExportProgress $Message $Percent
}

function Register-MatrixRowStats($Row) {
    if (-not (Test-IsMatrixRow $Row)) { return }

    [void]$Stats.MatrixSiteNames.Add((Get-MatrixRowColumnValue $Row 'Site Name'))
    $itemPath = Get-MatrixRowColumnValue $Row 'Item path'
    if ($itemPath) { [void]$Stats.MatrixPaths.Add($itemPath) }

    if ((Get-MatrixRowColumnValue $Row 'Given through') -eq 'Sharing Link') { $Stats.MatrixSharingLinkRows++ }

    if ((Get-MatrixRowColumnValue $Row 'Principal type') -eq 'User') {
        $account = Get-MatrixRowColumnValue $Row 'Account name'
        if ($account) {
            [void]$Stats.MatrixUsers.Add($account)
            if ((Get-MatrixRowColumnValue $Row 'External user') -eq 'Yes') {
                [void]$Stats.MatrixExternalUsers.Add($account)
                $siteName = Get-MatrixRowColumnValue $Row 'Site Name'
                if (-not $Stats.ExternalUsersBySite.ContainsKey($siteName)) {
                    $Stats.ExternalUsersBySite[$siteName] = [System.Collections.Generic.HashSet[string]]::new(
                        [StringComparer]::OrdinalIgnoreCase)
                }
                [void]$Stats.ExternalUsersBySite[$siteName].Add($account)
            }
        }
    }
    elseif ((Get-MatrixRowColumnValue $Row 'Principal type') -eq 'SharePoint group') {
        $group = Get-MatrixRowColumnValue $Row 'User/group'
        if ($group) { [void]$Stats.MatrixSpGroups.Add($group) }
    }
}

function Format-Elapsed([TimeSpan] $Elapsed) {
    $t = [math]::Max(0, [int][math]::Floor($Elapsed.TotalSeconds))
    $h = [int][math]::Floor($t / 3600)
    $m = [int][math]::Floor(($t % 3600) / 60)
    $s = $t % 60
    if ($h -gt 0) { return "{0}h {1}m {2}s" -f $h, $m, $s }
    if ($m -gt 0) { return "{0}m {1}s" -f $m, $s }
    return "{0}s" -f $s
}

function Start-ProgressRenderer {
    param([string] $Activity = 'Building permissions matrix')

    if ($null -ne $script:ProgressRenderHandle -and -not $script:ProgressRenderHandle.IsCompleted) { return }

    $script:ProgressRenderState = [hashtable]::Synchronized(@{
        Activity   = $Activity
        BaseStatus = 'Starting...'
        Percent    = 0
        SpinMs     = $script:ProgressSpinnerFrameMs
        StartTicks = $script:RunStartUtc.Ticks
        Stop       = $false
        Complete   = $false
    })

    $renderScript = {
        param($State)

        $trackWidth = 14
        $blockLen = 3
        $i = 0
        while (-not [bool]$State.Stop) {
            $activity = [string]$State.Activity
            if ([string]::IsNullOrWhiteSpace($activity)) { $activity = 'Building permissions matrix' }

            $baseStatus = [string]$State.BaseStatus
            if ([string]::IsNullOrWhiteSpace($baseStatus)) { $baseStatus = 'Working...' }

            $pct = [int]$State.Percent
            if ($pct -lt 0) { $pct = 0 }
            if ($pct -gt 100) { $pct = 100 }

            $spinMs = [int]$State.SpinMs
            if ($spinMs -lt 20) { $spinMs = 20 }
            if ($spinMs -gt 300) { $spinMs = 300 }

            # Compute elapsed live so it always advances even when main thread is blocked.
            $startTicks = [int64]$State.StartTicks
            $elapsedText = ''
            if ($startTicks -gt 0) {
                $elapsedSec = [int][Math]::Floor(([datetime]::UtcNow.Ticks - $startTicks) / [int64]10000000)
                if ($elapsedSec -lt 0) { $elapsedSec = 0 }
                $h = [int][Math]::Floor($elapsedSec / 3600)
                $m = [int][Math]::Floor(($elapsedSec % 3600) / 60)
                $s = $elapsedSec % 60
                $elapsedText = if ($h -gt 0) { "${h}h ${m}m ${s}s" } elseif ($m -gt 0) { "${m}m ${s}s" } else { "${s}s" }
            }

            # Elapsed goes FIRST so it is never truncated when the terminal is narrow.
            # The list/item title at the end can be clipped without losing critical info.
            $fullStatus = if ($elapsedText) { "elapsed $elapsedText | $baseStatus" } else { $baseStatus }

            $sweep = $trackWidth + $blockLen
            $head = $i % $sweep
            $chars = New-Object char[] $trackWidth
            for ($c = 0; $c -lt $trackWidth; $c++) { $chars[$c] = '.' }
            for ($b = 0; $b -lt $blockLen; $b++) {
                $pos = $head - $b
                if ($pos -lt 0 -or $pos -ge $trackWidth) { continue }
                $chars[$pos] = '='
            }
            $frame = -join $chars
            $i++

            try {
                Write-Progress -Activity $activity -Status ("[{0}] {1}" -f $frame, $fullStatus) -PercentComplete $pct
            }
            catch { }

            Start-Sleep -Milliseconds $spinMs
        }

        if ([bool]$State.Complete) {
            try { Write-Progress -Activity ([string]$State.Activity) -Completed } catch { }
        }
    }

    try {
        $script:ProgressRenderRunspace = [runspacefactory]::CreateRunspace($Host)
        $script:ProgressRenderRunspace.ApartmentState = 'MTA'
        $script:ProgressRenderRunspace.ThreadOptions = 'ReuseThread'
        $script:ProgressRenderRunspace.Open()

        $script:ProgressRenderPowerShell = [powershell]::Create()
        $script:ProgressRenderPowerShell.Runspace = $script:ProgressRenderRunspace
        $null = $script:ProgressRenderPowerShell.AddScript($renderScript.ToString()).AddArgument($script:ProgressRenderState)
        $script:ProgressRenderHandle = $script:ProgressRenderPowerShell.BeginInvoke()
    }
    catch {
        $script:ProgressRenderState = $null
        if ($script:ProgressRenderPowerShell) { $script:ProgressRenderPowerShell.Dispose() }
        if ($script:ProgressRenderRunspace) { $script:ProgressRenderRunspace.Dispose() }
        $script:ProgressRenderPowerShell = $null
        $script:ProgressRenderRunspace = $null
        $script:ProgressRenderHandle = $null
        Write-Verbose ("Progress renderer unavailable: {0}" -f $_.Exception.Message)
    }
}

function Stop-ProgressRenderer {
    param([switch] $Complete)

    $activityToComplete = 'Progress'
    if ($script:ProgressRenderState -and -not [string]::IsNullOrWhiteSpace([string]$script:ProgressRenderState.Activity)) {
        $activityToComplete = [string]$script:ProgressRenderState.Activity
    }

    if ($script:ProgressRenderState) {
        $script:ProgressRenderState.Complete = [bool]$Complete
        $script:ProgressRenderState.Stop = $true
    }

    if ($script:ProgressRenderPowerShell -and $script:ProgressRenderHandle) {
        try { $null = $script:ProgressRenderPowerShell.EndInvoke($script:ProgressRenderHandle) } catch { }
    }

    if ($script:ProgressRenderPowerShell) { try { $script:ProgressRenderPowerShell.Dispose() } catch { } }
    if ($script:ProgressRenderRunspace) { try { $script:ProgressRenderRunspace.Close(); $script:ProgressRenderRunspace.Dispose() } catch { } }

    $script:ProgressRenderState = $null
    $script:ProgressRenderRunspace = $null
    $script:ProgressRenderPowerShell = $null
    $script:ProgressRenderHandle = $null

    if ($Complete) {
        try { Write-Progress -Activity $activityToComplete -Completed } catch { }
    }
}

function Initialize-ListProgress {
    param(
        [string] $ListTitle,
        [int]    $ListItemCount
    )

    $S.ProgressListTitle = $ListTitle
    $S.ProgressListItemTotal = [Math]::Max(0, $ListItemCount)
    $S.ProgressListItemProcessed = 0
    $S.ProgressListUniqueFound = 0
    $S.ProgressPhase = 'Starting'
    $S.ProgressScanStartUtc = [datetime]::UtcNow
    $S.ProgressConsoleLastUtc = [datetime]::MinValue
}

function Get-ProgressPercent {
    param(
        [double] $Done,
        [double] $Total,
        [switch] $AllowComplete
    )

    if ($Total -le 0) { return 0 }
    $pct = [Math]::Floor(($Done / $Total) * 100)
    if (-not $AllowComplete -and -not $S.ProgressScanFinished) {
        return [Math]::Min(99, [int]$pct)
    }
    return [Math]::Min(100, [int]$pct)
}

function Get-OverallProgressPercent {
    if ($S.ProgressScanFinished) { return 100 }
    return Get-ProgressPercent -Done $S.ProgressProcessed -Total $S.ProgressTotal
}

function Complete-ScanProgress {
    $S.ProgressScanFinished = $true
    if ($S.ProgressTotal -gt $S.ProgressProcessed) {
        $S.ProgressProcessed = $S.ProgressTotal
    }
    $itemsTotal = if ($S.ProgressItemsTotal -gt 0) { $S.ProgressItemsTotal } else { $S.ProgressTotal }
    Write-MatrixLog ("Scan complete ({0:N0}/{1:N0} items)" -f $Stats.ItemsScanned, $itemsTotal) -Level Success
    Write-ProgressStatus -Force
}

function Add-ProgressUnits {
    param([int] $Count)

    if ($Count -gt 0) { $S.ProgressProcessed += $Count }
}

function Add-ScanBackendTiming {
    param(
        [ValidateSet('Rest', 'Csom')]
        [string] $Backend,
        [int]    $Items,
        [TimeSpan] $Elapsed
    )

    if ($Items -lt 1) { return }
    if (-not $script:ScanBackendTiming.ContainsKey($Backend)) { return }

    $entry = $script:ScanBackendTiming[$Backend]
    $entry.Lists += 1
    $entry.Items += [int]$Items
    $entry.Ms += [int][Math]::Ceiling($Elapsed.TotalMilliseconds)
}

function Get-ScanBackendItemsPerSecond {
    param([ValidateSet('Rest', 'Csom')] [string] $Backend)

    if (-not $script:ScanBackendTiming.ContainsKey($Backend)) { return 0.0 }
    $entry = $script:ScanBackendTiming[$Backend]
    if ($entry.Items -lt 1 -or $entry.Ms -lt 1) { return 0.0 }
    return ([double]$entry.Items / ([double]$entry.Ms / 1000.0))
}

function Get-MedianValue {
    param([double[]] $Values)

    $arr = @($Values | Where-Object { $_ -gt 0 } | Sort-Object)
    if ($arr.Count -lt 1) { return 0.0 }
    $mid = [int][Math]::Floor($arr.Count / 2)
    if ($arr.Count % 2 -eq 1) { return [double]$arr[$mid] }
    return ([double]$arr[$mid - 1] + [double]$arr[$mid]) / 2.0
}

function Get-BlendedTelemetryValue {
    param(
        [double] $Last,
        [double] $Baseline,
        [int]    $Samples
    )

    if ($Last -le 0 -and $Baseline -le 0) { return 0.0 }
    if ($Last -le 0) { return $Baseline }
    if ($Baseline -le 0) { return $Last }

    $lastWeight = if ($Samples -lt 4) { 0.85 } elseif ($Samples -lt 10) { 0.75 } else { 0.65 }
    return ($Last * $lastWeight) + ($Baseline * (1.0 - $lastWeight))
}

function Get-RunMetricsModel {
    param(
        [string] $Mode = '',
        [bool]   $IncludeAllInheritedItemsInMatrix = $false
    )

    if (-not (Test-Path -LiteralPath $RunMetricsPath)) {
        return [pscustomobject]@{
            RestItemsPerSec = 0.0
            CsomItemsPerSec = 0.0
            ExportRowsPerSec = 0.0
            SiteOverheadSec = 0.0
            LastRestItemsPerSec = 0.0
            LastCsomItemsPerSec = 0.0
            LastExportRowsPerSec = 0.0
            LastSiteOverheadSec = 0.0
            LastTelemetryUtc = ''
            Samples = 0
        }
    }

    $records = @()
    try {
        foreach ($line in Get-Content -LiteralPath $RunMetricsPath -ErrorAction Stop) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $records += (ConvertFrom-Json -InputObject $line -ErrorAction Stop) } catch { }
        }
    }
    catch { }

    $recentAll = @($records | Select-Object -Last 120)
    $recent = @($recentAll)
    if ($Mode) {
        $recent = @($recent | Where-Object { ([string]$_.Mode) -eq $Mode })
    }
    $recent = @($recent | Where-Object {
        $rowInherited = $false
        if ($null -ne $_.PSObject.Properties['IncludeAllInheritedItemsInMatrix']) {
            $rowInherited = [bool]$_.IncludeAllInheritedItemsInMatrix
        }
        $rowInherited -eq $IncludeAllInheritedItemsInMatrix
    })
    if ($recent.Count -lt 1) {
        # Fallback: keep mode sensitivity if possible, but don't fail estimate when history is sparse.
        if ($Mode) {
            $recent = @($recentAll | Where-Object { ([string]$_.Mode) -eq $Mode } | Select-Object -Last 60)
        }
        if ($recent.Count -lt 1) {
            $recent = @($recentAll | Select-Object -Last 60)
        }
    }

    if ($recent.Count -lt 1) {
        return [pscustomobject]@{
            RestItemsPerSec = 0.0
            CsomItemsPerSec = 0.0
            ExportRowsPerSec = 0.0
            SiteOverheadSec = 0.0
            LastRestItemsPerSec = 0.0
            LastCsomItemsPerSec = 0.0
            LastExportRowsPerSec = 0.0
            LastSiteOverheadSec = 0.0
            LastTelemetryUtc = ''
            Samples = 0
        }
    }

    $restRates = @()
    $csomRates = @()
    $exportRates = @()
    $overheads = @()
    foreach ($r in $recent) {
        if ($r.BackendRestItems -gt 0 -and $r.BackendRestMs -gt 0) {
            $restRates += ([double]$r.BackendRestItems / ([double]$r.BackendRestMs / 1000.0))
        }
        if ($r.BackendCsomItems -gt 0 -and $r.BackendCsomMs -gt 0) {
            $csomRates += ([double]$r.BackendCsomItems / ([double]$r.BackendCsomMs / 1000.0))
        }
        if ($r.MatrixRows -gt 0 -and $r.ExportSeconds -gt 0) {
            $exportRates += ([double]$r.MatrixRows / [double]$r.ExportSeconds)
        }
        if ($r.SitesScanned -gt 0 -and $r.ScanSeconds -gt 0) {
            $overheads += ([double]$r.ScanSeconds / [double]$r.SitesScanned)
        }
    }

    $last = $recent[$recent.Count - 1]
    $lastRestRate = if ($last.BackendRestItems -gt 0 -and $last.BackendRestMs -gt 0) {
        ([double]$last.BackendRestItems / ([double]$last.BackendRestMs / 1000.0))
    } else { 0.0 }
    $lastCsomRate = if ($last.BackendCsomItems -gt 0 -and $last.BackendCsomMs -gt 0) {
        ([double]$last.BackendCsomItems / ([double]$last.BackendCsomMs / 1000.0))
    } else { 0.0 }
    $lastExportRate = if ($last.MatrixRows -gt 0 -and $last.ExportSeconds -gt 0) {
        ([double]$last.MatrixRows / [double]$last.ExportSeconds)
    } else { 0.0 }
    $lastOverhead = if ($last.SitesScanned -gt 0 -and $last.ScanSeconds -gt 0) {
        ([double]$last.ScanSeconds / [double]$last.SitesScanned)
    } else { 0.0 }

    $medianRest = Get-MedianValue $restRates
    $medianCsom = Get-MedianValue $csomRates
    $medianExport = Get-MedianValue $exportRates
    $medianOverhead = Get-MedianValue $overheads

    return [pscustomobject]@{
        RestItemsPerSec = (Get-BlendedTelemetryValue -Last $lastRestRate -Baseline $medianRest -Samples $recent.Count)
        CsomItemsPerSec = (Get-BlendedTelemetryValue -Last $lastCsomRate -Baseline $medianCsom -Samples $recent.Count)
        ExportRowsPerSec = (Get-BlendedTelemetryValue -Last $lastExportRate -Baseline $medianExport -Samples $recent.Count)
        SiteOverheadSec = (Get-BlendedTelemetryValue -Last $lastOverhead -Baseline $medianOverhead -Samples $recent.Count)
        LastRestItemsPerSec = $lastRestRate
        LastCsomItemsPerSec = $lastCsomRate
        LastExportRowsPerSec = $lastExportRate
        LastSiteOverheadSec = $lastOverhead
        LastTelemetryUtc = [string]$last.TimestampUtc
        Samples = $recent.Count
    }
}

function Write-RunMetricsRecord {
    param(
        [string] $Mode,
        [TimeSpan] $Elapsed,
        [int] $SitesScanned,
        [int] $ItemsScanned,
        [int] $MatrixRows,
        [double] $ExportSeconds
    )

    try {
        $dir = Split-Path -Parent $RunMetricsPath
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $scanSeconds = [Math]::Max(0.0, $Elapsed.TotalSeconds - $ExportSeconds)
        $rec = [ordered]@{
            TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
            SiteUrl = $SiteUrl
            Mode = $Mode
            SitesScanned = $SitesScanned
            ItemsScanned = $ItemsScanned
            MatrixRows = $MatrixRows
            TotalSeconds = [Math]::Round($Elapsed.TotalSeconds, 3)
            ScanSeconds = [Math]::Round($scanSeconds, 3)
            ExportSeconds = [Math]::Round($ExportSeconds, 3)
            BackendRestItems = [int]$script:ScanBackendTiming.Rest.Items
            BackendRestMs = [int]$script:ScanBackendTiming.Rest.Ms
            BackendCsomItems = [int]$script:ScanBackendTiming.Csom.Items
            BackendCsomMs = [int]$script:ScanBackendTiming.Csom.Ms
            IncludeSubsites = [bool]$IncludeSubsites
            ExpandGroups = [bool]$ExpandGroups
            IncludeFolderSharingLinks = [bool]$IncludeFolderSharingLinks
            IncludeAllInheritedItemsInMatrix = [bool]$IncludeAllInheritedItemsInMatrix
            MaxListItemsForFullMatrix = [int]$MaxListItemsForFullMatrix
            ListItemPageSize = [int]$ListItemPageSize
        }
        $json = ($rec | ConvertTo-Json -Compress -Depth 5)
        Add-Content -LiteralPath $RunMetricsPath -Value $json -Encoding UTF8
    }
    catch {
        Write-Warning ("Failed to write run metrics: {0}" -f $_.Exception.Message)
    }
}

function Get-PreferredBackendByTiming {
    param(
        [int]    $ListItemCount,
        [string] $Default = 'Rest'
    )

    if ($ListItemCount -lt $LargeListRestThreshold) { return $null }

    $rest = $script:ScanBackendTiming['Rest']
    $csom = $script:ScanBackendTiming['Csom']
    if ($rest.Items -lt 1000 -and $csom.Items -lt 1000) { return $Default }
    if ($rest.Items -ge 1000 -and $csom.Items -lt 1000) { return 'Rest' }
    if ($csom.Items -ge 1000 -and $rest.Items -lt 1000) { return 'Csom' }

    $restRate = Get-ScanBackendItemsPerSecond -Backend 'Rest'
    $csomRate = Get-ScanBackendItemsPerSecond -Backend 'Csom'
    if ($restRate -le 0 -and $csomRate -le 0) { return $Default }
    if ($restRate -le 0) { return 'Csom' }
    if ($csomRate -le 0) { return 'Rest' }
    # Require a clear margin before flipping to avoid backend thrash.
    if ($csomRate -ge ($restRate * 1.2)) { return 'Csom' }
    return 'Rest'
}

function Get-ProgressSpinnerFrame {
    $frames = @('|', '/', '-', '\')
    $ticksPerFrame = [int64]2000000
    $idx = [int](([datetime]::UtcNow.Ticks / $ticksPerFrame) % $frames.Count)
    return $frames[$idx]
}

function Get-RunElapsedText {
    if ($script:RunStartUtc -eq [datetime]::MinValue) { return '0s' }
    $elapsed = [datetime]::UtcNow - $script:RunStartUtc
    return (Format-Elapsed $elapsed)
}

function Get-ActiveListPageSize {
    if ($script:CurrentListPageSize -gt 0) { return [int]$script:CurrentListPageSize }
    return [int]$ListItemPageSize
}

function Get-EffectiveListPageSize {
    param(
        [int]    $ListItemCount,
        [ValidateSet('Rest', 'Csom')]
        [string] $Backend
    )

    if ($ListItemCount -lt 1) { return [int]$ListItemPageSize }
    return [Math]::Min(5000, $ListItemCount)
}


function Write-ScanProgressConsole {
    param([switch] $Force)

    $now = [datetime]::UtcNow
    if (-not $Force -and $S.ProgressUniqueConsoleLastUtc -ne [datetime]::MinValue) {
        if (($now - $S.ProgressUniqueConsoleLastUtc).TotalSeconds -lt 1) { return }
    }
    $S.ProgressUniqueConsoleLastUtc = $now

    $pct = Get-OverallProgressPercent
    $pagePart = if ($S.ProgressListPageCount -gt 0) {
        "page $($S.ProgressListPageNum)/$($S.ProgressListPageCount)"
    } else { 'scanning' }

    $detail = if ($S.ProgressPhase -eq 'Sharing links' -and $S.ProgressSharingBatchTotal -gt 0) {
        $sDone = $S.ProgressSharingDone
        $sTotal = $S.ProgressSharingBatchTotal
        $sPct = Get-ProgressPercent -Done $sDone -Total $sTotal
        "sharing links $sDone/$sTotal ($sPct%)"
    }
    elseif ($S.ProgressUniqueFetchTotal -gt 0) {
        $uDone = $S.ProgressUniqueFetchDone
        $uTotal = $S.ProgressUniqueFetchTotal
        $uPct = Get-ProgressPercent -Done $uDone -Total $uTotal
        $uElapsed = if ($S.ProgressUniqueFetchStartUtc -ne [datetime]::MinValue) {
            [math]::Max(0.5, ($now - $S.ProgressUniqueFetchStartUtc).TotalSeconds)
        } else { 1.0 }
        $uRate = if ($uDone -gt 0) { [math]::Max(1, [int]($uDone / $uElapsed)) } else { 0 }
        $eta = if ($uDone -gt 0 -and $uDone -lt $uTotal -and $uRate -gt 0) {
            $sec = [int](($uTotal - $uDone) / $uRate)
            if ($sec -gt 0) { " ~{0} left" -f (Format-Elapsed ([TimeSpan]::FromSeconds($sec))) } else { '' }
        } else { '' }
        "unique perms $uDone/$uTotal ($uPct%, ~$uRate/s$eta)"
    }
    elseif ($S.ProgressPhase -and $S.ProgressPhase -notin @('Scanning items', 'Starting')) {
        $S.ProgressPhase
    }
    else {
        $elapsed = if ($S.ProgressScanStartUtc -ne [datetime]::MinValue) {
            [math]::Max(0.5, ($now - $S.ProgressScanStartUtc).TotalSeconds)
        } else { 1.0 }
        $iDone = $S.ProgressListItemProcessed
        $iTotal = $S.ProgressListItemTotal
        $iRate = if ($iDone -gt 0) { [math]::Max(1, [int]($iDone / $elapsed)) } else { 0 }
        $iPct = Get-ProgressPercent -Done $iDone -Total $iTotal
        "list items $iDone/$iTotal ($iPct%, ~$iRate/s)"
    }

    $runElapsed = Get-RunElapsedText
    Write-Host ("       {0} | {1} | {2} | elapsed {3} | overall {4}% ({5:N0}/{6:N0})" -f `
        $S.ProgressListTitle, $pagePart, $detail, $runElapsed, $pct, $S.ProgressProcessed, $S.ProgressTotal) -ForegroundColor DarkCyan

    Write-ProgressStatus -Force
}

function Process-UniqueListItems {
    param(
        [object[]] $UniqueItems,
        [System.Collections.Generic.List[object]] $AllRows,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $items = @($UniqueItems)
    $total = $items.Count
    if ($total -lt 1) { return }

    $S.ProgressUniqueFetchTotal = $total
    $S.ProgressUniqueFetchDone = 0
    $S.ProgressUniqueFetchCurrentPath = ''
    $S.ProgressUniqueFetchStartUtc = [datetime]::UtcNow
    $S.ProgressUniqueConsoleLastUtc = [datetime]::MinValue
    $S.ProgressPhase = 'Unique perms'
    Write-MatrixLog ("       ... matrix permissions for {0:N0} unique item(s)" -f $total) -Level Pulse
    Write-ScanProgressConsole -Force

    $i = 0
    foreach ($item in $items) {
        $i++
        $S.ProgressUniqueFetchDone = $i

        $path = [string](Get-ScanItemField $item 'FileRef')
        try {
            Process-ListItem -Item $item -AllRows $AllRows -ListId $ListId -ListTitle $ListTitle `
                -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
                -SkipNonUniqueMatrixRows:$true -UniqueOnlyProcessing `
                -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
        }
        catch {
            throw "Unique-perm matrix failed for item '$path' (Id=$($item.Id)): $($_.Exception.Message)"
        }

        Add-ProgressUnits -Count ($UniquePermProgressWeight - 1)
        Write-ScanProgressConsole
    }

    Write-MatrixLog ("       ... matrix done for {0:N0} unique item(s) in {1}" -f $total, `
        (Format-Elapsed (([datetime]::UtcNow - $S.ProgressUniqueFetchStartUtc)))) -Level Pulse

    $S.ProgressUniqueFetchTotal = 0
    $S.ProgressUniqueFetchDone = 0
    $S.ProgressUniqueFetchCurrentPath = ''
    $S.ProgressPhase = 'Scanning items'
    Write-ScanProgressConsole -Force
}

function Write-ScanHeartbeat {
    param([switch] $Force)

    Write-ScanProgressConsole -Force:$Force
}

function Write-ProgressStatus {
    param([switch] $Force)

    if (-not $Force -and $S.ProgressLastUtc -ne [datetime]::MinValue) {
        if (([datetime]::UtcNow - $S.ProgressLastUtc).TotalMilliseconds -lt $script:ProgressStatusThrottleMs) { return }
    }
    $S.ProgressLastUtc = [datetime]::UtcNow

    $pct = Get-OverallProgressPercent
    $statusParts = [System.Collections.Generic.List[string]]::new()
    [void]$statusParts.Add("{0}%" -f $pct)
    if ($S.ProgressSiteTotal -gt 1 -and $S.ProgressSiteIndex -gt 0) {
        [void]$statusParts.Add(("Site {0}/{1}" -f $S.ProgressSiteIndex, $S.ProgressSiteTotal))
    }
    $itemsTotal = if ($S.ProgressItemsTotal -gt 0) { $S.ProgressItemsTotal } else { $S.ProgressTotal }
    [void]$statusParts.Add(("items {0:N0}/{1:N0}" -f $Stats.ItemsScanned, $itemsTotal))
    # Truncate the title so it never pushes the fixed stats off-screen.
    $title = $S.ProgressListTitle
    if ($title -and $title.Length -gt 28) { $title = $title.Substring(0, 26).TrimEnd() + [char]0x2026 }
    $baseStatus = ("{0} | {1}" -f ($statusParts -join ' | '), $title)
    $spinMs = $script:ProgressSpinnerFrameMs

    if ($script:ProgressRenderState) {
        $script:ProgressRenderState.Activity   = 'Building permissions matrix'
        $script:ProgressRenderState.BaseStatus = $baseStatus
        $script:ProgressRenderState.Percent    = $pct
        $script:ProgressRenderState.SpinMs     = $spinMs
        # Keep StartTicks live so renderer always has correct reference.
        if ($script:RunStartUtc -ne [datetime]::MinValue) {
            $script:ProgressRenderState.StartTicks = $script:RunStartUtc.Ticks
        }
        return
    }

    # Fallback (no renderer): elapsed first so it is never truncated.
    $elapsed = Get-RunElapsedText
    Write-Progress -Activity 'Building permissions matrix' `
        -Status ("elapsed $elapsed | {0}" -f ($statusParts -join ' | ') + " | $title") `
        -PercentComplete $pct
}

function Bump-Progress {
    param([string] $ListTitle, [switch] $Force)

    $S.ProgressProcessed++
    $S.ProgressListItemProcessed++
    $S.ProgressListTitle = $ListTitle

    if ($Force -or ($S.ProgressListItemProcessed % 250 -eq 0)) {
        Write-ProgressStatus -Force
        Write-ScanHeartbeat -Force
    }
    else {
        Write-ProgressStatus
        Write-ScanHeartbeat
    }
}


function Add-ListItemScanStats {
    param(
        [string] $ListBaseType,
        [string] $ItemType,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $Stats.ItemsScanned++
    switch ($ItemType) {
        'File'      { $Stats.Files++; $SiteFiles.Value++ }
        'Folder'    { $Stats.Folders++; $SiteFolders.Value++ }
        'List item' { $Stats.ListItems++; $SiteListItems.Value++ }
    }
}

function Add-ItemRowFromObject {
    param(
        $Item,
        [string] $ListTitle,
        [string] $ItemPath,
        [string] $ItemType,
        [bool]   $HasUnique
    )
    $sizeRaw = $null  # file size removed; placeholder if re-added later
    Add-ScannedItemRow -ListTitle $ListTitle -ItemPath $ItemPath -ItemType $ItemType `
        -ItemId ([int](Get-ScanItemId $Item)) -HasUnique $HasUnique `
        -Created    (Get-ScanItemField $Item 'Created') `
        -CreatedBy  (Get-FieldUserDisplayName (Get-ScanItemField $Item 'Author')) `
        -Modified   (Get-ScanItemField $Item 'Modified') `
        -ModifiedBy (Get-FieldUserDisplayName (Get-ScanItemField $Item 'Editor'))
}

function Add-ScannedItemRow {
    param(
        [string]   $ListTitle,
        [string]   $ItemPath,
        [string]   $ItemType,
        [int]      $ItemId,
        [bool]     $HasUnique,
        $Created    = $null,
        [string]   $CreatedBy  = '',
        $Modified   = $null,
        [string]   $ModifiedBy = ''
    )

    if ([string]::IsNullOrWhiteSpace($ItemPath)) { return }

    # Coerce date strings to DateTime so Excel renders them as dates, not text.
    $createdDt  = if ($Created)  { try { [DateTime]$Created }  catch { $Created } }  else { $null }
    $modifiedDt = if ($Modified) { try { [DateTime]$Modified } catch { $Modified } } else { $null }

    $AllItemRows.Add([pscustomobject]@{
        'Site Name'           = $S.SiteName
        'List/Library'        = $ListTitle
        'Item path'           = $ItemPath
        'Item Type'           = $ItemType
        'Item Id'             = $ItemId
        'Broken permissions'  = $(if ($HasUnique) { 'Yes' } else { 'No' })
        'Created'             = $createdDt
        'Created By'          = $CreatedBy
        'Modified'            = $modifiedDt
        'Modified By'         = $ModifiedBy
    }) | Out-Null
}

function Get-ListItemQuickInfo {
    param(
        $Item,
        [string] $ListBaseType,
        [string] $ListUrlNorm
    )

    $fileRef = [string](Get-ScanItemField $Item 'FileRef')
    if (-not $fileRef -or $fileRef.TrimEnd('/') -eq $ListUrlNorm) {
        return $null
    }

    return [pscustomobject]@{
        Id        = (Get-ScanItemId $Item)
        FileRef   = $fileRef
        HasUnique = (Get-ItemHasUniqueRoleAssignments $Item)
        ItemType  = (Get-ContentItemType $ListBaseType (Get-ScanItemFsObjType $Item))
    }
}

function Process-ListItemPage {
    param(
        [object[]] $Batch,
        [System.Collections.Generic.List[object]] $AllRows,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [bool]   $SkipNonUniqueMatrixRows,
        [bool]   $AllItemsInherit = $false,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $items = @($Batch)
    if ($items.Count -lt 1) { return 0 }

    if ($SkipNonUniqueMatrixRows) {
        $uniqueItems = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $items) {
            $info = Get-ListItemQuickInfo -Item $item -ListBaseType $ListBaseType -ListUrlNorm $ListUrlNorm
            if (-not $info) { continue }

            $S.PathToItemId[$info.FileRef] = $info.Id
            if ($info.HasUnique) { [void]$S.UniquePermPaths.Add($info.FileRef) }

            Add-ListItemScanStats -ListBaseType $ListBaseType -ItemType $info.ItemType `
                -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems

            # When the probe confirmed no items have unique permissions, skip the
            # HasUniqueRoleAssignments check entirely and treat every item as inherited.
            $hasUnique = if ($AllItemsInherit) { $false } else { [bool]$info.HasUnique }
            Add-ItemRowFromObject -Item $item -ListTitle $ListTitle -ItemPath $info.FileRef `
                -ItemType $info.ItemType -HasUnique $hasUnique

            if ($hasUnique) {
                [void]$Stats.UniqueItemPaths.Add($info.FileRef)
                $S.ProgressListUniqueFound++
                $uniqueItems.Add($item) | Out-Null
            }
            else {
                $Stats.BulkSkippedItems++
            }
        }

        Add-ProgressUnits -Count $items.Count
        $S.ProgressListItemProcessed += $items.Count
        if ($uniqueItems.Count -gt 0) {
            $S.ProgressTotal += $uniqueItems.Count * ($UniquePermProgressWeight - 1)
            Write-ScanProgressConsole -Force
        }

        Process-UniqueListItems -UniqueItems @($uniqueItems) -AllRows $AllRows -ListId $ListId `
            -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
            -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
        Flush-SharingQueue -ListTitle $ListTitle
        return $items.Count
    }

    foreach ($item in $items) {
        Register-ListItemPath $item
        Process-ListItem -Item $item -AllRows $AllRows -ListId $ListId -ListTitle $ListTitle `
            -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
            -SkipNonUniqueMatrixRows:$false `
            -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
    }

    return $items.Count
}

function Test-MissingUniquePermField {
    param($ErrorRecord)

    $msg = [string]$ErrorRecord
    if ($ErrorRecord.Exception) { $msg = [string]$ErrorRecord.Exception.Message }
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $msg = "$msg $($ErrorRecord.ErrorDetails.Message)"
    }
    return ($msg -match 'HasUniqueRoleAssignments' -and $msg -match 'does not exist')
}

function Test-RestMissingListFieldError {
    param($ErrorRecord)

    $msg = [string]$ErrorRecord
    if ($ErrorRecord.Exception) { $msg = [string]$ErrorRecord.Exception.Message }
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $msg = "$msg $($ErrorRecord.ErrorDetails.Message)"
    }
    return ($msg -match 'does not exist' -and $msg -match 'field or property')
}

function Test-ListSupportsUniquePermField {
    param([string] $ListId)

    if ($S.ListUniquePermField.ContainsKey($ListId)) {
        return $S.ListUniquePermField[$ListId]
    }

    try {
        $null = Invoke-SPRestGet "/_api/web/lists(guid'$ListId')/items?`$select=Id,HasUniqueRoleAssignments&`$top=1"
        $S.ListUniquePermField[$ListId] = $true
        return $true
    }
    catch {
        if (Test-MissingUniquePermField $_) {
            $S.ListUniquePermField[$ListId] = $false
            return $false
        }
        throw
    }
}


function Get-ListItemSelectFields {
    param(
        [bool]   $IncludeUniqueField,
        [switch] $Minimal,
        [ValidateSet('Full', 'Core')]
        [string] $Profile = 'Full'
    )

    $fields = [System.Collections.Generic.List[string]]::new()
    [void]$fields.Add('Id')
    if ($IncludeUniqueField) { [void]$fields.Add('HasUniqueRoleAssignments') }
    [void]$fields.Add('FileRef')
    [void]$fields.Add('FileLeafRef')
    [void]$fields.Add('FSObjType')
    [void]$fields.Add('FileSystemObjectType')

    if ($Profile -eq 'Core') {
        return ($fields -join ',')
    }

    if (-not $Minimal) { [void]$fields.Add('Title') }
    # Metadata for the All Items sheet — Author/Editor require $expand (added at URL level).
    [void]$fields.Add('Created')
    [void]$fields.Add('Modified')
    [void]$fields.Add('Author/Title')
    [void]$fields.Add('Editor/Title')
    return ($fields -join ',')
}

function New-ListItemsRestPageUrl {
    param(
        [string] $ListId,
        [string] $Select,
        [int]    $Top,
        [switch] $ExpandAuthorEditor
    )

    $url = "/_api/web/lists(guid'$ListId')/items?`$select=$Select&`$top=$Top"
    if ($ExpandAuthorEditor.IsPresent) { $url += '&$expand=Author,Editor' }
    return $url
}

# Extracts a display name from a user field value that may be a REST expanded object
# ({Title:...}), a CSOM FieldUserValue ({LookupValue:...}), or a "ID;#Name" string.
function Get-FieldUserDisplayName($FieldValue) {
    if ($null -eq $FieldValue) { return '' }
    if ($FieldValue -is [string]) {
        if ($FieldValue -match ';#(.+)$') { return $Matches[1] }
        return $FieldValue
    }
    $lv = Get-RestPropertyValue $FieldValue 'LookupValue'
    if ($lv) { return [string]$lv }
    $t  = Get-RestPropertyValue $FieldValue 'Title'
    if ($t) { return [string]$t }
    return ''
}

function Get-ScanItemField {
    param(
        $Item,
        [string] $Name
    )

    if ($null -eq $Item) { return $null }

    try {
        $direct = $Item.$Name
        if ($null -ne $direct -and "$direct" -ne '') { return $direct }
    }
    catch { }

    $prop = $Item.PSObject.Properties[$Name]
    if ($prop -and $null -ne $prop.Value -and "$($prop.Value)" -ne '') {
        return $prop.Value
    }

    if ($Item.FieldValues) {
        $fv = $Item.FieldValues[$Name]
        if ($null -ne $fv -and "$fv" -ne '') { return $fv }
    }

    try {
        if ($Item -is [System.Collections.IDictionary] -and $Item -isnot [System.Management.Automation.PSCustomObject]) {
            if ($Item.Contains($Name)) {
                $indexed = $Item[$Name]
                if ($null -ne $indexed -and "$indexed" -ne '') { return $indexed }
            }
        }
        else {
            $indexed = $Item[$Name]
            if ($null -ne $indexed -and "$indexed" -ne '') { return $indexed }
        }
    }
    catch { }

    return $null
}

function Get-ItemHasUniqueRoleAssignments($Item) {
    if ($null -eq $Item) { return $false }

    $prop = $Item.PSObject.Properties['HasUniqueRoleAssignments']
    if ($prop -and ($prop.Value -is [bool])) { return [bool]$prop.Value }

    try {
        if ($Item.HasUniqueRoleAssignments -is [bool]) { return [bool]$Item.HasUniqueRoleAssignments }
    }
    catch { }

    $raw = Get-ScanItemField $Item 'HasUniqueRoleAssignments'
    if ($raw -is [bool]) { return $raw }
    if ($null -ne $raw -and "$raw" -ne '') {
        $s = [string]$raw
        return ($s -eq 'True' -or $s -eq 'true' -or $s -eq '1')
    }

    return $false
}

function Ensure-ListItemsUniquePermLoaded {
    param([object[]] $Items)

    $batch = @($Items | Where-Object { $null -ne $_ })
    if ($batch.Count -lt 1) { return }

    if ($batch.Count -eq 1) {
        Get-PnPProperty -ClientObject $batch[0] -Property HasUniqueRoleAssignments -ErrorAction Stop | Out-Null
        return
    }

    $chunkSize = 4000
    for ($offset = 0; $offset -lt $batch.Count; $offset += $chunkSize) {
        $take = [Math]::Min($chunkSize, $batch.Count - $offset)
        $chunk = @($batch[$offset..($offset + $take - 1)])

        try {
            $ctx = Get-PnPContext
            foreach ($item in $chunk) {
                $ctx.Load($item, [Microsoft.SharePoint.Client.ListItem].GetPropertyByName('HasUniqueRoleAssignments'))
            }
            Invoke-PnPQuery -ErrorAction Stop | Out-Null
        }
        catch {
            foreach ($item in $chunk) {
                Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments -ErrorAction Stop | Out-Null
            }
        }
    }
}

function Register-ListItemPath($Item) {
    $path = [string](Get-ScanItemField $Item 'FileRef')
    if (-not $path) { return }
    $S.PathToItemId[$path] = Get-ScanItemId $Item
    if (Get-ItemHasUniqueRoleAssignments $Item) { [void]$S.UniquePermPaths.Add($path) }
}

function Get-ScanItemFsObjType($Item) {
    $v = Get-ScanItemField $Item 'FSObjType'
    if ($null -eq $v) { $v = Get-ScanItemField $Item 'FileSystemObjectType' }
    return $v
}

function Get-ScanItemId($Item) {
    if ($null -eq $Item) { return 0 }

    $raw = Get-ScanItemField $Item 'ID'
    if ($null -eq $raw -or "$raw" -eq '') { $raw = Get-ScanItemField $Item 'Id' }
    if ($null -ne $raw -and "$raw" -ne '') { return [int]$raw }

    try { return [int]$Item.Id } catch { return 0 }
}

function Get-ODataNextLink($Response) {
    if ($null -eq $Response) { return $null }
    foreach ($name in '@odata.nextLink', 'odata.nextLink') {
        $v = $Response.PSObject.Properties[$name]
        if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) {
            return ([string]$v.Value).Trim()
        }
    }
    return $null
}

function Invoke-SPRestGet([string] $Url) {
    if ([string]::IsNullOrWhiteSpace($Url)) { throw 'REST URL was empty.' }
    $u = $Url.Trim()
    if ($u -notmatch '^https?://') {
        if ($u.StartsWith('/_api/')) {
            $base = if ($S.WebPath -and $S.WebPath -ne '/') { $S.WebPath } else { '' }
            $u = "$($S.SiteHostUrl)$base$u"
        }
    }
    return Invoke-PnPSPRestMethod -Url $u -Method Get
}

function Invoke-SPRestGetWithRetry {
    param(
        [string] $Url,
        [int]    $MaxAttempts = 4
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return Invoke-SPRestGet $Url
        }
        catch {
            $msg = [string]$_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $msg = "$msg $($_.ErrorDetails.Message)"
            }
            $retryable = ($msg -match '(?i)stream|timeout|timed out|429|503|502|504|thrott|busy|temporarily')
            if (-not $retryable -or $attempt -eq $MaxAttempts) { throw }

            $delay = [Math]::Min(15, [Math]::Pow(2, $attempt))
            Write-MatrixLog ("       REST retry {0}/{1} in {2}s ({3})" -f $attempt, ($MaxAttempts - 1), $delay, `
                ($msg.Substring(0, [Math]::Min(80, $msg.Length)))) -Level Warn
            Start-Sleep -Seconds $delay
        }
    }
}

function Add-RoleColumns([string[]] $Names) {
    foreach ($name in $PreferredRoleOrder) {
        if ($Names -contains $name -and $S.RoleColumnSet.Add($name)) {
            $S.RoleColumns.Add($name) | Out-Null
        }
    }
    foreach ($name in ($Names | Sort-Object)) {
        if ($S.RoleColumnSet.Add($name)) { $S.RoleColumns.Add($name) | Out-Null }
    }
}

function Load-RoleDefinitions {
    # Fetches all role definitions from the current site (uses $S.WebPath for URL routing)
    # and populates $S.RoleDefById. Returns the flat list of definition objects.
    $S.RoleDefById.Clear()
    $allDefs = [System.Collections.Generic.List[object]]::new()
    $url = '/_api/web/roledefinitions?$select=Id,Name&$top=500'
    while ($url) {
        $r = Invoke-SPRestGet $url
        $page = Get-RestPropertyValue $r 'value'
        if ($null -eq $page) { $page = Get-RestPropertyValue $r 'results' }
        if ($null -eq $page) { $page = Get-RestPropertyValue (Get-RestPropertyValue $r 'd') 'results' }
        foreach ($rd in @($page)) {
            $id = [int]$rd.Id; $name = [string]$rd.Name
            if ($id -gt 0 -and $name) {
                $S.RoleDefById[$id] = $name
                $allDefs.Add($rd) | Out-Null
            }
        }
        $url = Get-ODataNextLink $r
    }
    return @($allDefs)
}

function Initialize-RoleColumns {
    # $S.WebPath must be seeded from $S.ConnectedWebUrl before calling so that
    # Invoke-SPRestGet targets the correct site collection, not the tenant root.
    $defs = Load-RoleDefinitions
    Add-RoleColumns @($defs | ForEach-Object { [string]$_.Name } | Where-Object { $_ })
    $script:LastRoleDefSiteCollection = $S.ConnectedWebUrl -replace '/[^/]+$'
}

# Refreshes the role-definition ID→name map when switching sites. Skips the
# REST call if the new site is in the same site collection as the last fetch
# (role definition IDs are site-collection-scoped, not per-web).
function Refresh-RoleDefById {
    $siteRoot = $S.ConnectedWebUrl -replace '(/sites/[^/]+).*', '$1'
    if ($siteRoot -and $siteRoot -eq $script:LastRoleDefSiteCollection) { return }
    Load-RoleDefinitions | Out-Null
    $script:LastRoleDefSiteCollection = $siteRoot
}

function Get-RoleDefinitionBindingList {
    param($Bindings)

    if ($null -eq $Bindings) { return @() }
    if (Test-IsInternalRoleCache $Bindings) { return @() }

    # OData v4 / plain array
    if ($Bindings -is [System.Array]) { return @(Get-NormalizedBindingItems $Bindings) }

    if ($Bindings -is [System.Management.Automation.PSCustomObject]) {
        # OData v2 deferred link — caller must follow separately
        if ($null -ne (Get-RestPropertyValue $Bindings '__deferred')) { return @() }

        # OData v2 verbose collection: {"results": [...]}
        $results = Get-RestPropertyValue $Bindings 'results'
        if ($null -ne $results) { return @(Get-NormalizedBindingItems $results) }

        # OData v4 collection: {"value": [...]}
        $value = Get-RestPropertyValue $Bindings 'value'
        if ($null -ne $value) { return @(Get-NormalizedBindingItems $value) }

        # Single binding object — has Name, RoleDefinitionId, RoleDefinition, or a
        # recognisable Id (validated against the known role-definition map)
        $name    = Get-RestPropertyValue $Bindings 'Name'
        $roleId  = Get-RestPropertyValue $Bindings 'RoleDefinitionId'
        $roleDef = Get-RestPropertyValue $Bindings 'RoleDefinition'
        $directId = Get-RestPropertyValue $Bindings 'Id'
        if ($name -or $roleId -or $roleDef -or
            ($null -ne $directId -and [int]$directId -gt 0 -and $S.RoleDefById.ContainsKey([int]$directId))) {
            return @(Get-NormalizedBindingItems @($Bindings))
        }
        return @()
    }

    if ($Bindings -is [System.Collections.IEnumerable] -and $Bindings -isnot [string]) {
        return @(Get-NormalizedBindingItems $Bindings)
    }

    return @()
}

function Get-NormalizedBindingItems {
    param($Items)

    $values = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @($Items)) {
        if ($null -eq $item) { continue }
        if ($item -is [string]) { continue }
        if (Test-IsAnyKeyValuePair $item) { $item = $item.Value }
        if (Test-IsInternalRoleCache $item) { continue }
        if ($item -is [System.Collections.IDictionary] -and $item -isnot [System.Management.Automation.PSCustomObject]) { continue }
        $values.Add($item) | Out-Null
    }
    return @($values)
}

function Get-RoleNameFromBinding {
    param($Binding)

    if (-not $Binding -or $Binding -is [string]) { return '' }
    if (Test-IsInternalRoleCache $Binding) { return '' }
    if ($Binding -is [System.Collections.IDictionary] -and $Binding -isnot [System.Management.Automation.PSCustomObject]) { return '' }

    # Direct Name (full expand returned inline)
    $name = [string](Get-RestPropertyValue $Binding 'Name')
    if ($name) { return $name }

    # Nested RoleDefinition.Name (only when not deferred)
    $roleDef = Get-RestPropertyValue $Binding 'RoleDefinition'
    if ($roleDef -and -not (Get-RestPropertyValue $roleDef '__deferred')) {
        $name = [string](Get-RestPropertyValue $roleDef 'Name')
        if ($name) { return $name }
    }

    # ID-based lookup — try all known property names in priority order:
    #   RoleDefinitionId (OData v2 verbose) → Id (OData v4 binding) → RoleDefinition.Id
    $roleId = 0
    foreach ($prop in @('RoleDefinitionId', 'Id')) {
        $val = Get-RestPropertyValue $Binding $prop
        if ($null -ne $val) { $roleId = [int]$val; break }
    }
    if ($roleId -lt 1 -and $roleDef) {
        $val = Get-RestPropertyValue $roleDef 'Id'
        if ($null -ne $val) { $roleId = [int]$val }
    }

    if ($roleId -gt 0 -and $S.RoleDefById.ContainsKey($roleId)) {
        return $S.RoleDefById[$roleId]
    }
    return ''
}

function Get-RoleNamesFromBindingList {
    param([object[]] $Bindings)

    $roles = [System.Collections.Generic.List[string]]::new()
    foreach ($binding in @($Bindings)) {
        $name = Get-RoleNameFromBinding $binding
        if ($name) { $roles.Add($name) | Out-Null }
    }
    return @($roles | Sort-Object -Unique)
}


function Resolve-RoleAssignmentRoles {
    param(
        $Assignment,
        [string] $Scope,
        [string] $ListId = '',
        [int]    $ItemId = 0
    )

    if (-not (Test-IsRoleAssignmentObject $Assignment)) { return @() }

    try {
        # Fast path: already resolved in a prior call on this assignment object
        if ($Assignment.PSObject.Properties.Match('_ResolvedRoleNames').Count -gt 0) {
            $resolved = Get-ResolvedRoleNamesFromAssignment $Assignment
            if ($null -ne $resolved) { return $resolved }
        }

        # Step 1: bindings already in the response (works when $expand=RoleDefinitionBindings* succeeds)
        $roles = Get-RoleNamesFromBindingList (Get-RoleDefinitionBindingList (Get-RestPropertyValue $Assignment 'RoleDefinitionBindings'))
        if ($roles.Count -gt 0) { return $roles }

        # Step 2: navigate directly to RoleDefinitionBindings — no expand needed.
        # In SharePoint REST, roleassignments are keyed by PrincipalId (OData v4
        # responses omit 'Id'; fall back to Member.Id / Assignment.Id for OData v2).
        $principalId = [int](Get-RestPropertyValue $Assignment 'PrincipalId')
        if ($principalId -lt 1) {
            $mem = Get-RestPropertyValue $Assignment 'Member'
            if ($mem) { $principalId = [int](Get-RestPropertyValue $mem 'Id') }
        }
        if ($principalId -lt 1) { $principalId = [int](Get-RestPropertyValue $Assignment 'Id') }

        if ($principalId -gt 0) {
            $bindUrl = switch -Wildcard ($Scope) {
                'web' { "/_api/web/roleassignments($principalId)/RoleDefinitionBindings" }
                default {
                    if ($ItemId -gt 0) {
                        "/_api/web/lists(guid'$ListId')/items($ItemId)/roleassignments($principalId)/RoleDefinitionBindings"
                    } else {
                        "/_api/web/lists(guid'$ListId')/roleassignments($principalId)/RoleDefinitionBindings"
                    }
                }
            }
            try {
                $bindResp = Invoke-SPRestGetWithRetry $bindUrl
                $bindItems = Get-RestPropertyValue $bindResp 'value'
                if ($null -eq $bindItems) { $bindItems = Get-RestPropertyValue $bindResp 'results' }
                if ($null -eq $bindItems) { $bindItems = Get-RestPropertyValue (Get-RestPropertyValue $bindResp 'd') 'results' }
                $roles = if ($null -ne $bindItems) {
                    Get-RoleNamesFromBindingList (Get-NormalizedBindingItems @($bindItems))
                } else {
                    Get-RoleNamesFromBindingList (Get-RoleDefinitionBindingList $bindResp)
                }
                if ($roles.Count -gt 0) { return $roles }
            }
            catch {
                Write-Verbose "RoleDefinitionBindings navigation failed ($Scope/$principalId): $_"
            }
        }
    }
    catch {
        Write-Warning "Resolve-RoleAssignmentRoles failed ($Scope): $($_.Exception.Message)"
    }

    return @()
}

function Enrich-RoleAssignments {
    param(
        [object[]] $Assignments,
        [string] $Scope,
        [string] $ListId = '',
        [int]    $ItemId = 0
    )

    $enriched = [System.Collections.Generic.List[object]]::new()
    foreach ($assignment in (Get-NormalizedAssignmentArray $Assignments)) {
        try {
            $roles = Resolve-RoleAssignmentRoles -Assignment $assignment -Scope $Scope `
                -ListId $ListId -ItemId $ItemId
            Set-AssignmentResolvedRoles -Assignment $assignment -Roles $roles
            $enriched.Add($assignment) | Out-Null
        }
        catch {
            Write-Warning "Enrich role assignment failed ($Scope): $($_.Exception.Message)"
        }
    }
    return @($enriched)
}

function Test-IsSharePointTenantAdminUrl {
    param([string] $Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    try {
        $u = [Uri]$Url.Trim()
        if ($u.Host -match '-admin\.sharepoint\.com$') { return $true }
        if ($u.Host -match '\.sharepoint\.com$' -and $u.AbsolutePath -match '/_layouts/.*/(online|tenant)/') {
            return $true
        }
    }
    catch { }
    return $false
}

function Get-SharePointTenantHostsFromUrl {
    param([string] $Url)
    $uri = [Uri]$Url.Trim()
    $tenantName = $null
    if ($uri.Host -match '^(.+)-admin\.sharepoint\.com$') {
        $tenantName = $Matches[1]
    }
    elseif ($uri.Host -match '^([^.]+)\.sharepoint\.com$') {
        $tenantName = $Matches[1]
    }
    if (-not $tenantName) {
        throw "Could not resolve SharePoint tenant name from URL: $Url"
    }
    $rootHost = "$tenantName.sharepoint.com"
    $adminHost = "$tenantName-admin.sharepoint.com"
    [pscustomobject]@{
        TenantName = $tenantName
        RootHost   = $rootHost
        AdminHost  = $adminHost
        IsAdminUrl = ($uri.Host -match '-admin\.sharepoint\.com$')
        RootUrl    = "https://$rootHost"
        AdminUrl   = "https://$adminHost"
    }
}

function Set-PnPConnectionScopeFromInputUrl {
    param([string] $Url)
    $info = Get-SharePointTenantHostsFromUrl $Url
    $S.TenantRootHost = $info.RootHost
    $S.IsTenantAdminScope = $info.IsAdminUrl
    if ($info.IsAdminUrl) {
        $S.SiteHostUrl = $info.RootUrl
        $S.ConnectedWebUrl = (Get-NormalizedWebUrl $Url)
    }
    else {
        $S.SiteHostUrl = '{0}://{1}' -f ([Uri]$Url).Scheme, ([Uri]$Url).Host
        $S.ConnectedWebUrl = Get-NormalizedWebUrl $Url
        $S.IsTenantAdminScope = $false
    }
}

function Get-SharePointHostBaseFromUrl {
    param([string] $Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $S.SiteHostUrl }
    try {
        $h = ([Uri]$Url).Host
        if ($h -match '-admin\.sharepoint\.com$') {
            return ('{0}://{1}' -f ([Uri]$Url).Scheme, ($h -replace '-admin\.sharepoint\.com', '.sharepoint.com'))
        }
        return ('{0}://{1}' -f ([Uri]$Url).Scheme, $h)
    }
    catch {
        return $S.SiteHostUrl
    }
}

function Get-NormalizedWebUrl {
    param([string] $UrlOrPath, [switch] $ServerRelative)

    if ($ServerRelative) {
        $p = $UrlOrPath.Trim()
        if (-not $p) { $p = '/' }
        if ($p -ne '/' -and $p.EndsWith('/')) { $p = $p.TrimEnd('/') }
        $hostBase = $S.SiteHostUrl
        if ($hostBase -match '-admin\.sharepoint\.com') {
            if ($S.TenantRootHost) {
                $hostBase = "https://$($S.TenantRootHost)"
            }
            else {
                $hostBase = $hostBase -replace '-admin\.sharepoint\.com', '.sharepoint.com'
            }
        }
        if ($p -eq '/') { return $hostBase.ToLowerInvariant() }
        return ("{0}{1}" -f $hostBase, $p).ToLowerInvariant()
    }
    $uri = [Uri]$UrlOrPath
    $p = $uri.AbsolutePath
    if ($p -ne '/' -and $p.EndsWith('/')) { $p = $p.TrimEnd('/') }
    return ('{0}://{1}{2}' -f $uri.Scheme, $uri.Host, $p).ToLowerInvariant()
}

function Connect-SiteWeb {
    param(
        [object] $Web,
        [switch] $InventoryOnly,
        [switch] $Quiet
    )

    $needProps = @('ServerRelativeUrl', 'Title', 'HasUniqueRoleAssignments')
    $missing = @($needProps | Where-Object {
            -not $Web.PSObject.Properties[$_] -or $null -eq $Web.$_
        })
    if ($missing.Count -gt 0) {
        Get-PnPProperty -ClientObject $Web -Property $needProps | Out-Null
    }

    $target = Get-NormalizedWebUrl ([string]$Web.ServerRelativeUrl) -ServerRelative
    if ($S.ConnectedWebUrl -ne $target) {
        Connect-PnPOnline -Url $target -Interactive -ClientId $S.PnpClientId | Out-Null
        $S.ConnectedWebUrl = $target
        $S.SiteHostUrl = Get-SharePointHostBaseFromUrl $target
    }

    if ($InventoryOnly) {
        $web = $Web
    }
    else {
        $web = Get-PnPWeb
        Get-PnPProperty -ClientObject $web -Property Title, HasUniqueRoleAssignments, ServerRelativeUrl | Out-Null
    }

    $S.SiteName = [string]$web.Title
    $S.WebPath = ([string]$web.ServerRelativeUrl).TrimEnd('/')
    if (-not $S.WebPath) { $S.WebPath = '/' }
    $S.WebHasUnique = [bool]$web.HasUniqueRoleAssignments
    $S.WebItemType = if ($S.WebPath -match '^/sites/[^/]+/.+') { 'Subsite' } else { 'Site' }

    if (-not $InventoryOnly) {
        # Role def IDs are site-collection-scoped; refresh when exporting permissions.
        Refresh-RoleDefById
        try {
            $parent = Get-PnPProperty -ClientObject $web -Property ParentWeb -ErrorAction Stop
            if ($parent -and [string]$parent.ServerRelativeUrl) { $S.WebItemType = 'Subsite' }
        }
        catch { }
    }

    if (-not $Quiet) {
        Write-MatrixLog ("Site: {0} ({1}, {2}) unique permissions: {3}" -f $S.SiteName, $S.WebPath, $S.WebItemType, $S.WebHasUnique) -Level Dim
    }
}

function Reset-SiteState {
    $S.AssignmentCache.Clear()
    $S.TemplateCache.Clear()
    $S.GroupMembers.Clear()
    $S.SiteInheritedTpl = $null
    $S.UniquePermPaths.Clear()
    $S.ListHasUnique.Clear()
    $S.ListUniquePermField.Clear()
    $S.PathToItemId = [System.Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)
    $S.EmittedKeys = [System.Collections.Generic.HashSet[string]]::new()
}

function Get-ScanLists {
    Get-PnPList -Includes BaseType, Hidden, Title, ItemCount, RootFolder, HasUniqueRoleAssignments |
        Where-Object {
            -not $_.Hidden -and
            -not $ExcludedListTitles.Contains($_.Title) -and
            ($_.BaseType -eq 'DocumentLibrary' -or $_.BaseType -eq 'GenericList')
        }
}

function Get-TenantSiteCollectionUrls {
    $excludedTemplates = [System.Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )
    foreach ($t in @(
            'SPSPERS#10', 'SPSPERS#11', 'SPSPERS',
            'REDIRECTSITE#0', 'REDIRECTSITE'
        )) {
        [void]$excludedTemplates.Add($t)
    }

    Write-MatrixLog 'Enumerating site collections from tenant (Get-PnPTenantSite)...' -Level Tenant
    $tenantSites = @()
    try {
        $tenantSites = @(Get-PnPTenantSite)
    }
    catch {
        throw @"
Tenant site enumeration failed. Connect with the SharePoint admin URL (https://<tenant>-admin.sharepoint.com)
and sign in with an account that has SharePoint Administrator rights (or app permissions for SharePoint admin APIs).
Original error: $($_.Exception.Message)
"@
    }

    $picked = [System.Collections.Generic.List[object]]::new()
    $skipped = 0
    foreach ($ts in $tenantSites) {
        $url = [string]$ts.Url
        if ([string]::IsNullOrWhiteSpace($url)) { $skipped++; continue }
        $tpl = [string]$ts.Template
        if ($excludedTemplates.Contains($tpl)) { $skipped++; continue }
        if ($url -match '(?i)-my\.sharepoint\.com|/personal/') { $skipped++; continue }
        $status = [string]$ts.Status
        if ($status -eq 'Deleted') { $skipped++; continue }
        $picked.Add($ts) | Out-Null
    }

    Write-MatrixLog ("Tenant inventory: {0} site collection(s) to scan ({1} skipped: OneDrive/redirect/deleted/empty)." -f `
            $picked.Count, $skipped) -Level Tenant
    return @($picked)
}

function Get-WebsFromConnectedSiteRoot {
    param([bool] $IncludeSubsites)

    $webs = [System.Collections.Generic.List[object]]::new()
    $root = Get-PnPWeb -Includes ServerRelativeUrl, Title, HasUniqueRoleAssignments
    $webs.Add($root) | Out-Null
    if ($IncludeSubsites) {
        try {
            foreach ($sub in @(Get-PnPSubWeb -Recurse -Includes ServerRelativeUrl, Title, HasUniqueRoleAssignments | Sort-Object ServerRelativeUrl)) {
                $webs.Add($sub) | Out-Null
            }
        }
        catch {
            Write-Warning "Some subsites were not listed under '$($root.ServerRelativeUrl)': $($_.Exception.Message)"
        }
    }
    return @($webs)
}

function New-ScanPlanEntryForWeb {
    param($Web)

    Connect-SiteWeb $Web -InventoryOnly -Quiet
    $lists = @(Get-ScanLists)
    $items = 0
    foreach ($l in $lists) { $items += [int]$l.ItemCount }
    return [pscustomobject]@{
        Site  = $Web
        Lists = $lists
        Items = $items
    }
}

function Get-TenantPermissionsMatrixScanPlan {
    param(
        [bool]        $IncludeSubsites,
        [scriptblock] $ProgressCallback = $null
    )

    if ($ProgressCallback) {
        & $ProgressCallback 0 0 'Querying tenant site list (Get-PnPTenantSite)...'
    }

    $collections = @(Get-TenantSiteCollectionUrls)
    if ($collections.Count -lt 1) {
        Write-Warning 'No site collections were returned for this tenant.'
        return @()
    }

    $plan = [System.Collections.Generic.List[object]]::new()
    $S.TenantConnectSkipped = 0
    $S.TenantInventorySkipped = 0
    $collectionsWithAccess = 0
    $idx = 0
    $totalCollections = $collections.Count

    if ($ProgressCallback -and $totalCollections -gt 0) {
        & $ProgressCallback 0 $totalCollections "Found $totalCollections site collection(s) to check..."
    }

    foreach ($ts in $collections) {
        $idx++
        $siteUrl = ([string]$ts.Url).Trim().TrimEnd('/')
        $label = if ($ts.Title) { [string]$ts.Title } else { $siteUrl }
        if ($ProgressCallback) {
            & $ProgressCallback $idx $totalCollections $label
        }
        Write-MatrixLog ("  Site collection {0}/{1}: {2}" -f $idx, $collections.Count, $label) -Level Tenant

        try {
            Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $S.PnpClientId | Out-Null
            $S.SiteHostUrl = Get-SharePointHostBaseFromUrl $siteUrl
            $S.ConnectedWebUrl = Get-NormalizedWebUrl $siteUrl

            $webs = @()
            try {
                $webs = @(Get-WebsFromConnectedSiteRoot -IncludeSubsites $IncludeSubsites)
            }
            catch {
                $S.TenantConnectSkipped++
                Write-Warning "Skipping site collection '$siteUrl' (cannot read site/subsites): $($_.Exception.Message)"
                continue
            }

            if ($webs.Count -lt 1) { continue }
            $collectionsWithAccess++
            $hadInventory = $false

            foreach ($web in $webs) {
                $path = ''
                try { $path = [string]$web.ServerRelativeUrl } catch { }
                try {
                    $plan.Add((New-ScanPlanEntryForWeb -Web $web)) | Out-Null
                    $hadInventory = $true
                }
                catch {
                    $S.TenantInventorySkipped++
                    $pathLabel = if ($path) { $path } else { $siteUrl }
                    Write-Warning "Skipping list inventory for '$pathLabel': $($_.Exception.Message)"
                }
            }

            if (-not $hadInventory) {
                Write-Warning "Site collection '$siteUrl' is reachable but no webs could be inventoried (access denied on lists)."
            }
        }
        catch {
            $S.TenantConnectSkipped++
            Write-Warning "Skipping site collection '$siteUrl' (no access or connection failed): $($_.Exception.Message)"
        }
    }

    $skippedTotal = $S.TenantConnectSkipped + $S.TenantInventorySkipped
    if ($plan.Count -gt 0) {
        Write-MatrixLog ("Tenant inventory ready: {0} web(s) from {1} site collection(s) ({2} site collection(s)/web(s) skipped)." -f `
                $plan.Count, $collectionsWithAccess, $skippedTotal) -Level Success
    }
    elseif ($collectionsWithAccess -gt 0) {
        Write-Warning 'Connected to site collections but could not read list inventory on any web.'
    }

    return @($plan)
}

function Get-AssignmentRoles($Assignment) {
    if (-not (Test-IsRoleAssignmentObject $Assignment)) { return @() }

    try {
        $resolved = Get-ResolvedRoleNamesFromAssignment $Assignment
        if ($null -ne $resolved) { return $resolved }

        $bindings = Get-RestPropertyValue $Assignment 'RoleDefinitionBindings'
        return Get-RoleNamesFromBindingList (Get-RoleDefinitionBindingList $bindings)
    }
    catch {
        Write-Warning "Get-AssignmentRoles failed: $($_.Exception.Message)"
        return @()
    }
}

function Get-GroupMembers([int] $GroupId) {
    if ($S.GroupMembers.ContainsKey($GroupId)) { return $S.GroupMembers[$GroupId] }

    # SharePoint REST defaults to 100 users. SiteGroups/Users often omits
    # @odata.nextLink, so page with $top=5000, follow Get-ODataNextLink when
    # present, and $skip when a full page has no next link. Stop if a page
    # adds no new Ids ($skip ignored) so this cannot loop forever.
    # Keep in sync with lib/groupUsersPaging.mjs and fetchGroupUsers in
    # permissionsMatrixExport.js.
    $members = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    $pageSize = 5000
    $received = 0
    $url = "/_api/web/sitegroups($GroupId)/users?`$select=Id,Title,LoginName,Email,PrincipalType&`$top=$pageSize"
    $pages = 0
    try {
        while ($url -and $pages -lt 50) {
            $pages++
            $r = Invoke-SPRestGet $url
            $page = Get-RestPropertyValue $r 'value'
            if ($null -eq $page) { $page = Get-RestPropertyValue $r 'results' }
            if ($null -eq $page) { $page = Get-RestPropertyValue (Get-RestPropertyValue $r 'd') 'results' }
            $pageItems = @($page | Where-Object { $_ })
            $added = 0
            foreach ($u in $pageItems) {
                $id = [string](Get-RestPropertyValue $u 'Id')
                $login = [string](Get-RestPropertyValue $u 'LoginName')
                $key = if (-not [string]::IsNullOrWhiteSpace($id)) { "id:$id" } elseif (-not [string]::IsNullOrWhiteSpace($login)) { "login:$($login.ToLowerInvariant())" } else { '' }
                if ($key -and $seen.ContainsKey($key)) { continue }
                if ($key) { $seen[$key] = $true }
                $members.Add($u) | Out-Null
                $added++
            }
            if ($added -eq 0) { break }
            $next = Get-ODataNextLink $r
            if ($next) {
                $url = $next
                $received += $pageItems.Count
                continue
            }
            if ($pageItems.Count -ge $pageSize) {
                $received += $pageItems.Count
                $url = "/_api/web/sitegroups($GroupId)/users?`$select=Id,Title,LoginName,Email,PrincipalType&`$top=$pageSize&`$skip=$received"
            }
            else {
                $url = $null
            }
        }
    } catch {
        Write-Warning "Group members failed (group $GroupId): $($_.Exception.Message)"
        if ($members.Count -eq 0) {
            $S.GroupMembers[$GroupId] = @()
            return @()
        }
    }
    $out = @($members)
    $S.GroupMembers[$GroupId] = $out
    return $out
}

function Register-SharePointGroup($Member) {
    if (-not $Member -or [int]$Member.PrincipalType -ne 8) { return }

    $groupId = [int]$Member.Id
    if ($groupId -le 0) { return }

    $key = "{0}|{1}" -f $S.ConnectedWebUrl, $groupId
    if ($GroupCatalog.ContainsKey($key)) { return }

    $GroupCatalog[$key] = [pscustomobject]@{
        SiteName   = $S.SiteName
        SiteUrl    = $S.ConnectedWebUrl
        SitePath   = $S.WebPath
        GroupId    = $groupId
        GroupName  = [string]$Member.Title
        GroupLogin = [string]$Member.LoginName
    }
}

function Add-GroupMemberRowsForSite {
    foreach ($key in @($GroupCatalog.Keys)) {
        if ($GroupMembersFlushed.Contains($key)) { continue }
        if (-not $key.StartsWith("$($S.ConnectedWebUrl)|", [StringComparison]::OrdinalIgnoreCase)) { continue }

        $g = $GroupCatalog[$key]
        $members = @(Get-GroupMembers $g.GroupId)

        if ($members.Count -eq 0) {
            $GroupMemberRows.Add([pscustomobject]@{
                'Site Name'     = $g.SiteName
                'Site URL'      = $g.SiteUrl
                'Site Path'     = $g.SitePath
                'Group Name'    = $g.GroupName
                'Group Login'   = $g.GroupLogin
                'Group Id'      = $g.GroupId
                'Member Name'   = '(no members or unable to read)'
                'Member Login'  = ''
                'Member Email'  = ''
                'Member Type'   = ''
                'External user' = ''
            }) | Out-Null
        }
        else {
            foreach ($u in $members) {
                $login = [string]$u.LoginName
                $email = [string]$u.Email
                $isExternal = Test-ExternalUser $login $email
                if ($isExternal) {
                    Register-ExternalUserForSite -SiteName $g.SiteName -Login $(if ($login) { $login } else { $email })
                }
                $GroupMemberRows.Add([pscustomobject]@{
                    'Site Name'     = $g.SiteName
                    'Site URL'      = $g.SiteUrl
                    'Site Path'     = $g.SitePath
                    'Group Name'    = $g.GroupName
                    'Group Login'   = $g.GroupLogin
                    'Group Id'      = $g.GroupId
                    'Member Name'   = [string]$u.Title
                    'Member Login'  = $login
                    'Member Email'  = $email
                    'Member Type'   = (Get-PrincipalLabel ([int]$u.PrincipalType))
                    'External user' = if ($isExternal) { 'Yes' } else { '' }
                }) | Out-Null
            }
        }

        [void]$GroupMembersFlushed.Add($key)
    }
}

function Get-GroupMemberReportRows {
    return @($GroupMemberRows | Sort-Object 'Site Name', 'Group Name', 'Member Name', 'Member Login')
}

function Set-AssignmentResolvedRoles {
    param(
        $Assignment,
        [string[]] $Roles
    )

    if (-not (Test-IsRoleAssignmentObject $Assignment)) { return }
    if (Test-IsInternalRoleCache $Assignment) { return }
    if (-not $Roles -or $Roles.Count -lt 1) { return }

    try {
        $Assignment | Add-Member -NotePropertyName '_ResolvedRoleNames' -NotePropertyValue $Roles -Force
    }
    catch {
        Write-Verbose "Could not cache resolved roles on assignment: $($_.Exception.Message)"
    }
}

function Get-RestAssignmentItems {
    param($Response)

    if ($null -eq $Response) { return @() }
    if (Test-IsInternalRoleCache $Response) { return @() }

    if ($null -ne (Get-RestPropertyValue $Response 'value')) {
        return Expand-AssignmentCollection (Get-RestPropertyValue $Response 'value')
    }
    $results = Get-RestPropertyValue $Response 'results'
    if ($null -ne $results) {
        return Expand-AssignmentCollection $results
    }
    return Expand-AssignmentCollection $Response
}

function Get-RoleAssignments {
    param([string] $Scope, [string] $ListId = '', [int] $ItemId = 0)

    try {
        if ($S.AssignmentCache.ContainsKey($Scope)) {
            return Get-NormalizedAssignmentArray $S.AssignmentCache[$Scope]
        }

        # For item-level, $expand=.../RoleDefinition silently returns empty bindings in
        # SharePoint Online. Use $expand=RoleDefinitionBindings without the nested path
        # and resolve role names via $S.RoleDefById instead.
        $url = switch -Wildcard ($Scope) {
            'web' { '/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings/RoleDefinition' }
            default {
                if ($ItemId -gt 0) {
                    "/_api/web/lists(guid'$ListId')/items($ItemId)/roleassignments?`$expand=Member,RoleDefinitionBindings"
                } else {
                    "/_api/web/lists(guid'$ListId')/roleassignments?`$expand=Member,RoleDefinitionBindings/RoleDefinition"
                }
            }
        }

        try {
            $assignments = Get-RestAssignmentItems (Invoke-SPRestGetWithRetry $url)
        } catch {
            Write-Warning "Role assignments failed ($Scope): $($_.Exception.Message)"
            $assignments = @()
        }

        $assignments = Enrich-RoleAssignments -Assignments $assignments -Scope $Scope -ListId $ListId -ItemId $ItemId
        $S.AssignmentCache[$Scope] = @($assignments)
        return Get-NormalizedAssignmentArray $assignments
    }
    catch {
        Write-Warning "Get-RoleAssignments failed ($Scope): $($_.Exception.Message)"
        return @()
    }
}

function Test-ListScopedPrincipal([string] $Login, [string] $Title) {
    $b = "$Title|$Login"
    return ($b -match '(?i)Limited Access System Group|SharingLinks\.|Sharing Link')
}

function Test-SharingLink([string] $Login, [string] $Title) {
    return ($Login -like 'SharingLinks.*' -or $Title -like 'SharingLinks.*')
}

function Get-SharingLinkLabel([string] $Login, [string] $Title) {
    $src = if ($Login -like 'SharingLinks.*') { $Login } elseif ($Title -like 'SharingLinks.*') { $Title } else { return 'Sharing link' }
    if ($src -match 'SharingLinks\.[^.]+\.([^.]+)\.') {
        $k = $Matches[1]
        if ($k -match '^Organization(.+)$') { return "Organization ($($Matches[1].ToLower())) sharing link" }
        if ($k -match '^Anonymous(.+)$') { return "Anonymous ($($Matches[1].ToLower())) sharing link" }
        if ($k -match '^Users(.+)$') { return "Specific people ($($Matches[1].ToLower())) sharing link" }
        return "$k sharing link"
    }
    return 'Sharing link'
}

function Get-SharingObjectType([string] $ListBaseType, $FSObjType) {
    if ($ListBaseType -eq 'DocumentLibrary') {
        if ($null -ne $FSObjType -and [int]$FSObjType -eq 1) { return 'Folder' }
        return 'File'
    }
    return 'Item'
}

function Get-SharingItemName($Item, [string] $ListBaseType) {
    if ($ListBaseType -eq 'DocumentLibrary') { return [string](Get-ScanItemField $Item 'FileLeafRef') }
    return [string](Get-ScanItemField $Item 'Title')
}

function Get-SharingPrincipalTypeName([int] $PrincipalType) {
    switch ($PrincipalType) {
        1 { return 'User' }
        2 { return 'DistributionList' }
        4 { return 'SecurityGroup' }
        8 { return 'SharePointGroup' }
        default { return "Type$PrincipalType" }
    }
}

function Format-SharingAccessEntry {
    param(
        [string] $PrincipalType,
        [string] $Title,
        [string] $Login,
        [string] $Email,
        [string[]] $Roles
    )

    $typeLabel = switch ($PrincipalType) {
        'User'             { 'User' }
        'SharePointGroup'  { 'Group' }
        'SecurityGroup'    { 'Security Group' }
        'DistributionList' { 'Distribution List' }
        default            { $PrincipalType }
    }

    $roleText = if ($Roles.Count -gt 0) { ($Roles -join ', ') } else { '(no roles)' }

    if ($PrincipalType -eq 'User') {
        $identity = if (-not [string]::IsNullOrWhiteSpace($Email)) {
            if (-not [string]::IsNullOrWhiteSpace($Title) -and $Title -ne $Email) {
                "{0} <{1}>" -f $Title, $Email
            }
            else { $Email }
        }
        elseif (-not [string]::IsNullOrWhiteSpace($Title)) { $Title }
        else { $Login }

        return ("{0}: {1} [{2}]" -f $typeLabel, $identity, $roleText)
    }

    $name = if (-not [string]::IsNullOrWhiteSpace($Title)) { $Title } else { $Login }
    return ("{0}: {1} [{2}]" -f $typeLabel, $name, $roleText)
}

function Get-ItemAccessText {
    param([string] $ListId, [int] $ItemId)

    $assignments = Get-RoleAssignments "$ListId`:$ItemId" $ListId $ItemId
    $entries = [System.Collections.Generic.List[string]]::new()

    foreach ($assignment in $assignments) {
        $member = $assignment.Member
        if (-not $member) { continue }
        $entries.Add((Format-SharingAccessEntry `
            -PrincipalType (Get-SharingPrincipalTypeName ([int]$member.PrincipalType)) `
            -Title ([string]$member.Title) `
            -Login ([string]$member.LoginName) `
            -Email ([string]$member.Email) `
            -Roles (Get-AssignmentRoles $assignment))) | Out-Null
    }

    return ($entries -join '; ')
}

function Get-LinksForPath {
    param([string] $RelativeUrl, [bool] $IsFolder)

    try {
        if ($IsFolder) {
            return @(Get-PnPFolderSharingLink -Folder $RelativeUrl -ErrorAction Stop)
        }
        return @(Get-PnPFileSharingLink -Identity $RelativeUrl -ErrorAction Stop)
    }
    catch {
        Write-Warning ("Sharing-link API error for {0} : {1}" -f $RelativeUrl, $_.Exception.Message)
        return @()
    }
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

    $userEmails = [System.Collections.Generic.List[string]]::new()
    if ($Link.GrantedToIdentitiesV2) {
        foreach ($id in $Link.GrantedToIdentitiesV2) {
            if ($id.User -and $id.User.Email) { $userEmails.Add([string]$id.User.Email) | Out-Null }
        }
    }

    $hasExternal = Test-SharingLinkHasExternalUser -Link $Link
    if (-not $hasExternal) {
        foreach ($email in @($userEmails)) {
            if (Test-ExternalUser '' $email) { $hasExternal = $true; break }
        }
    }

    $blocksDownload = $Link.Link.PreventsDownload
    if ($null -eq $blocksDownload) { $blocksDownload = $Link.Link.PreventsDowload }

    $SharingReport.UniqueLinks[$LinkUrl] = [pscustomobject]@{
        RecordType       = 'SharingLink'
        ShareLinkUrl     = $LinkUrl
        ShareLinkType    = [string]$Link.Link.Type
        ShareLinkScope   = [string]$Link.Link.Scope
        LinkRoles        = ($Link.Roles -join '|')
        LinkUsers        = ($userEmails -join '|')
        'External user'  = (Get-ExternalUserFlag $hasExternal)
        Expiration       = $Link.ExpirationDateTime
        BlocksDownload   = $blocksDownload
        RequiresPassword = $Link.HasPassword
        ShareId          = $Link.Id
        SiteUrl          = $S.ConnectedWebUrl
        ListTitle        = $ListTitle
        ListUrl          = $ListUrl
        ObjectType       = $ObjectType
        ItemName         = $ItemName
        RelativeUrl      = $RelativeUrl
        ItemId           = $ItemId
        Access           = ''
    }
}

function Add-DirectPermissionSharingRow {
    param(
        [string] $ListId,
        [int]    $ItemId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ObjectType,
        [string] $ItemName,
        [string] $RelativeUrl
    )

    $SharingReport.DirectRows.Add([pscustomobject]@{
        RecordType       = 'DirectPermission'
        ShareLinkUrl     = ''
        ShareLinkType    = ''
        ShareLinkScope   = ''
        LinkRoles        = ''
        LinkUsers        = ''
        'External user'  = (Get-ExternalUserFlag (Test-ItemAssignmentsHaveExternalUser $ListId $ItemId))
        Expiration       = ''
        BlocksDownload   = ''
        RequiresPassword = ''
        ShareId          = ''
        SiteUrl          = $S.ConnectedWebUrl
        ListTitle        = $ListTitle
        ListUrl          = $ListUrl
        ObjectType       = $ObjectType
        ItemName         = $ItemName
        RelativeUrl      = $RelativeUrl
        ItemId           = $ItemId
        Access           = (Get-ItemAccessText -ListId $ListId -ItemId $ItemId)
    }) | Out-Null
}

function Queue-SharingItem {
    param(
        $Item,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListBaseType,
        [string] $ItemType
    )

    if ($ItemType -eq 'Folder' -and -not $IncludeFolderSharingLinks) { return }

    $relativeUrl = [string](Get-ScanItemField $Item 'FileRef')
    if ([string]::IsNullOrWhiteSpace($relativeUrl)) { return }

    $itemId = Get-ScanItemId $Item
    if ($itemId -lt 1) { return }

    $itemName = if ($ListBaseType -eq 'DocumentLibrary') {
        [string](Get-ScanItemField $Item 'FileLeafRef')
    } else {
        [string](Get-ScanItemField $Item 'Title')
    }
    if ([string]::IsNullOrWhiteSpace($itemName)) { $itemName = [System.IO.Path]::GetFileName($relativeUrl.TrimEnd('/')) }

    $S.ProgressTotal++
    $Script:SharingQueue.Add([pscustomobject]@{
        RelativeUrl  = $relativeUrl
        ItemId       = $itemId
        ItemName     = $itemName
        FSObjType    = Get-ScanItemFsObjType $Item
        ListId       = $ListId
        ListTitle    = $ListTitle
        ListUrl      = $ListUrl
        ListBaseType = $ListBaseType
        ItemType     = $ItemType
    }) | Out-Null
}

function Flush-SharingQueue {
    param([string] $ListTitle)

    $queue = @($Script:SharingQueue)
    if ($queue.Count -lt 1) { return }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-MatrixLog ("       ... sharing links for {0:N0} unique item(s)" -f $queue.Count) -Level Pulse

    $S.ProgressPhase = 'Sharing links'
    $S.ProgressSharingBatchTotal = $queue.Count
    $S.ProgressSharingDone = 0
    Write-ScanProgressConsole -Force

    $i = 0
    foreach ($entry in $queue) {
        $i++
        Add-SharingRowsForItem -SharingEntry $entry
        $S.ProgressSharingDone = $i
        Add-ProgressUnits 1
        if ($i -eq $queue.Count -or ($i % 25 -eq 0)) {
            Write-ScanProgressConsole -Force
        }
    }

    $Script:SharingQueue.Clear()
    $S.ProgressSharingBatchTotal = 0
    $S.ProgressSharingDone = 0
    $S.ProgressPhase = 'Scanning items'
    Write-MatrixLog ("       ... sharing links done ({0})" -f (Format-Elapsed $sw.Elapsed)) -Level Pulse
}

function Add-SharingRowsForItem {
    param(
        [Parameter(Mandatory = $true)]
        [object] $SharingEntry
    )

    $relativeUrl = [string]$SharingEntry.RelativeUrl
    if ([string]::IsNullOrWhiteSpace($relativeUrl)) { return }

    $listId = [string]$SharingEntry.ListId
    $listTitle = [string]$SharingEntry.ListTitle
    $listUrl = [string]$SharingEntry.ListUrl
    $listBaseType = [string]$SharingEntry.ListBaseType
    $itemId = [int]$SharingEntry.ItemId
    $itemName = [string]$SharingEntry.ItemName

    $SharingReport.UniqueItemsChecked++
    $fsObjType = $SharingEntry.FSObjType
    $isFolder = ($listBaseType -eq 'DocumentLibrary' -and [int]$fsObjType -eq 1)
    $objectType = Get-SharingObjectType $listBaseType $fsObjType

    $links = Get-LinksForPath -RelativeUrl $relativeUrl -IsFolder $isFolder

    if ($links.Count -eq 0) {
        Add-DirectPermissionSharingRow -ListId $listId -ItemId $itemId `
            -ListTitle $listTitle -ListUrl $listUrl -ObjectType $objectType `
            -ItemName $itemName -RelativeUrl $relativeUrl
        return
    }

    foreach ($link in $links) {
        $linkUrl = [string]$link.Link.WebUrl
        if ([string]::IsNullOrWhiteSpace($linkUrl) -or $SharingReport.UniqueLinks.ContainsKey($linkUrl)) { continue }
        Add-SharingLinkRow -Link $link -LinkUrl $linkUrl -ListTitle $listTitle -ListUrl $listUrl `
            -ObjectType $objectType -ItemName $itemName -RelativeUrl $relativeUrl -ItemId $itemId
    }
}

function Get-SharingReportRows {
    $rows = @($SharingReport.UniqueLinks.Values) + @($SharingReport.DirectRows)
    return @($rows | Sort-Object RecordType, ListTitle, RelativeUrl)
}

function Get-PrincipalLabel([int] $Type) {
    switch ($Type) {
        1 { 'User' }
        2 { 'Distribution list' }
        4 { 'Security group' }
        8 { 'SharePoint group' }
        default { "Principal type $Type" }
    }
}

function Test-ExternalUser([string] $Login, [string] $Email = '') {
    return ($Login -match '#ext#' -or $Email -match '#EXT#@')
}

function Test-SharingLinkHasExternalUser {
    param($Link)

    if ($Link.GrantedToIdentitiesV2) {
        foreach ($id in @($Link.GrantedToIdentitiesV2)) {
            if ($id.User) {
                if (Test-ExternalUser ([string]$id.User.LoginName) ([string]$id.User.Email)) { return $true }
            }
        }
    }

    if ($Link.GrantedToIdentities) {
        foreach ($id in @($Link.GrantedToIdentities)) {
            if ($id.User) {
                if (Test-ExternalUser ([string]$id.User.LoginName) ([string]$id.User.Email)) { return $true }
            }
        }
    }

    return $false
}

function Test-ItemAssignmentsHaveExternalUser {
    param([string] $ListId, [int] $ItemId)

    foreach ($assignment in (Get-RoleAssignments "$ListId`:$ItemId" $ListId $ItemId)) {
        $member = $assignment.Member
        if (-not $member) { continue }
        if ([int]$member.PrincipalType -ne 1) { continue }
        if (Test-ExternalUser ([string]$member.LoginName) ([string]$member.Email)) { return $true }
    }
    return $false
}

function Get-ExternalUserFlag {
    param([bool] $HasExternal)
    if ($HasExternal) { return 'Yes' }
    return ''
}

function Get-NormalizedRoleNames {
    param($Roles)

    if ($null -eq $Roles) { return @() }
    if (Test-IsInternalRoleCache $Roles) { return @() }
    if ($Roles -is [string]) { return @($Roles) }

    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($role in @($Roles)) {
        if ($null -eq $role -or (Test-IsInternalRoleCache $role)) { continue }
        if (Test-IsAnyKeyValuePair $role) { continue }
        if ($role -is [System.Collections.IDictionary] -and $role -isnot [System.Management.Automation.PSCustomObject]) { continue }
        $name = [string]$role
        if ($name) { $out.Add($name) | Out-Null }
    }
    return @($out)
}

function Get-MatrixItemName {
    param(
        [string] $ItemPath,
        [string] $ItemType,
        [string] $SiteName
    )

    if ($ItemType -eq 'Site' -or $ItemType -eq 'Subsite') {
        return [string]$SiteName
    }

    if ([string]::IsNullOrWhiteSpace($ItemPath)) { return '' }

    $path = $ItemPath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path)) { return '' }

    $leaf = $path
    $idx = $path.LastIndexOf('/')
    if ($idx -ge 0) { $leaf = $path.Substring($idx + 1) }

    try {
        return [Uri]::UnescapeDataString($leaf.Replace('+', ' '))
    }
    catch {
        return $leaf
    }
}

function New-MatrixRow {
    param(
        [string] $ItemPath, [string] $ItemType, [string] $Inheritance, [string] $Details,
        [string] $UserGroup, [string] $PrincipalType, [string] $AccountName,
        [string] $GivenThrough, [string[]] $Roles, [bool] $IsExternal = $false
    )
    $row = [ordered]@{
        'Site Name'      = $S.SiteName
        'Name'           = (Get-MatrixItemName $ItemPath $ItemType $S.SiteName)
        'Item path'      = $ItemPath
        'Item Type'      = $ItemType
        'Inheritance'    = $Inheritance
        'Details'        = $Details
        'User/group'     = $UserGroup
        'Principal type' = $PrincipalType
        'Account name'   = $AccountName
        'External user'  = if ($IsExternal) { 'Yes' } else { '' }
        'Given through'  = $GivenThrough
    }
    foreach ($col in $S.RoleColumns) { $row[$col] = '' }
    foreach ($role in (Get-NormalizedRoleNames $Roles)) {
        if ($S.RoleColumnSet.Contains($role)) { $row[$role] = 'X' }
    }
    return [pscustomobject]$row
}

function Get-RowKey($Row) {
    if (-not (Test-IsMatrixRow $Row)) {
        return ([string]$Row.GetType().FullName)
    }

    $sb = [System.Text.StringBuilder]::new(192)
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'Site Name')).Append('|')
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'Item path')).Append('|')
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'Item Type')).Append('|')
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'Inheritance')).Append('|')
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'User/group')).Append('|')
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'Account name')).Append('|')
    [void]$sb.Append((Get-MatrixRowColumnValue $Row 'Given through'))
    foreach ($col in $S.RoleColumns) {
        if ((Get-MatrixRowColumnValue $Row ([string]$col)) -eq 'X') {
            [void]$sb.Append('|').Append($col)
        }
    }
    return $sb.ToString()
}

function Add-Rows {
    param(
        [System.Collections.Generic.List[object]] $Target,
        [object[]] $Rows
    )
    foreach ($row in $Rows) {
        if (-not (Test-IsMatrixRow $row)) {
            Write-Warning ("Skipping invalid matrix row type: {0}" -f $row.GetType().FullName)
            continue
        }
        if ($S.EmittedKeys.Add((Get-RowKey $row))) {
            $Target.Add($row) | Out-Null
            Register-MatrixRowStats $row
        }
    }
}

function New-SummaryLine {
    param([string] $Section, [string] $Metric, $Value)
    [pscustomobject]@{ Section = $Section; Metric = $Metric; Value = $Value }
}

function Build-SummarySheet {
    param(
        [string] $SourceUrl,
        [TimeSpan] $Elapsed,
        [bool] $IncludeSubsites,
        [bool] $ExpandGroups,
        [int] $MatrixRowCount,
        [string] $MatrixSortNote = 'Yes'
    )

    $lines = [System.Collections.Generic.List[object]]::new()
    $add = { param($SectionName, $M, $V) $lines.Add((New-SummaryLine -Section $SectionName -Metric $M -Value $V)) | Out-Null }

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    & $add 'Report' 'Generated at' $generated
    & $add 'Report' 'Source site URL' $SourceUrl
    if ($S.PnpClientId) {
        & $add 'Report' 'Toolkit connection profile' (Protect-ToolkitClientId ([string]$S.PnpClientId))
    }
    & $add 'Report' 'Elapsed time' (Format-Elapsed $Elapsed)
    & $add 'Report' 'Include subsites' $(if ($IncludeSubsites) { 'Yes' } else { 'No' })
    & $add 'Report' 'Expand groups' $(if ($ExpandGroups) { 'Yes' } else { 'No' })
    & $add 'Report' 'Max list items for full matrix' $MaxListItemsForFullMatrix
    & $add 'Report' 'List item page size' $ListItemPageSize
    & $add 'Report' 'Item scan mode' $ItemScanMode
    & $add 'Report' 'Unique perm progress weight' $UniquePermProgressWeight
    & $add 'Report' 'Include folder sharing links' $(if ($IncludeFolderSharingLinks) { 'Yes' } else { 'No' })
    & $add 'Report' 'Include all inherited matrix rows' $(if ($IncludeAllInheritedItemsInMatrix) { 'Yes' } else { 'No' })
    & $add 'Report' 'Open workbook on complete' $(if ($OpenWorkbookOnComplete) { 'Yes' } else { 'No' })
    & $add 'Report' 'Run metrics file' (Split-Path -Leaf $RunMetricsPath)
    & $add 'Performance' 'REST lists scanned' $script:ScanBackendTiming.Rest.Lists
    & $add 'Performance' 'REST items scanned' $script:ScanBackendTiming.Rest.Items
    & $add 'Performance' 'REST throughput (items/sec)' ([int](Get-ScanBackendItemsPerSecond -Backend 'Rest'))
    & $add 'Performance' 'CSOM lists scanned' $script:ScanBackendTiming.Csom.Lists
    & $add 'Performance' 'CSOM items scanned' $script:ScanBackendTiming.Csom.Items
    & $add 'Performance' 'CSOM throughput (items/sec)' ([int](Get-ScanBackendItemsPerSecond -Backend 'Csom'))

    $totalSites = $Stats.RootSites + $Stats.Subsites
    & $add 'Scope' 'Sites and subsites scanned' $totalSites
    & $add 'Scope' 'Root sites' $Stats.RootSites
    & $add 'Scope' 'Subsites' $Stats.Subsites
    & $add 'Scope' 'Lists' $Stats.Lists
    & $add 'Scope' 'Libraries' $Stats.Libraries

    & $add 'Content' 'Total SharePoint items scanned' $Stats.ItemsScanned
    & $add 'Content' 'List items' $Stats.ListItems
    & $add 'Content' 'Files' $Stats.Files
    & $add 'Content' 'Folders' $Stats.Folders
    if ($Stats.BulkSkippedItems -gt 0) {
        & $add 'Content' 'Inherited items (no per-item matrix row)' $Stats.BulkSkippedItems
    }

    & $add 'Permissions' 'Sites with unique permissions' $Stats.SitesWithUniquePerms
    & $add 'Permissions' 'Lists/libraries with unique permissions' $Stats.ListsWithUniquePerms
    & $add 'Permissions' 'Items/folders/files with unique permissions' $Stats.UniqueItemPaths.Count

    & $add 'Principals' 'Distinct users in matrix' $Stats.MatrixUsers.Count
    & $add 'Principals' 'External users (distinct)' (Get-TotalExternalUserCount)
    & $add 'Principals' 'SharePoint groups (distinct)' $Stats.MatrixSpGroups.Count
    & $add 'Principals' 'Sharing link permission rows' $Stats.MatrixSharingLinkRows

    & $add 'Coverage' 'Distinct item paths in matrix' $Stats.MatrixPaths.Count
    & $add 'Coverage' 'Distinct site names in matrix' $Stats.MatrixSiteNames.Count
    & $add 'Coverage' 'All Items sheet rows' $AllItemRows.Count

    & $add 'Sharing links' 'Unique-perm items checked for links' $SharingReport.UniqueItemsChecked
    & $add 'Sharing links' 'Sharing link rows (distinct URLs)' $SharingReport.UniqueLinks.Count
    & $add 'Sharing links' 'Direct-permission rows (no link)' $SharingReport.DirectRows.Count

    & $add 'Groups' 'SharePoint groups found' $GroupCatalog.Count
    & $add 'Groups' 'Group member rows' $GroupMemberRows.Count

    return @{
        Overview    = @($lines)
        SiteDetails = @($Stats.SiteDetails)
    }
}

function Test-KeepInheritedRow($Row) {
    $hasReal = $false
    foreach ($col in $S.RoleColumns) {
        if ((Get-MatrixRowColumnValue $Row ([string]$col)) -ne 'X') { continue }
        if (-not $NonInheritableRoles.Contains($col)) { $hasReal = $true; break }
    }
    return $hasReal
}

function Test-IsMatrixRow {
    param($Row)

    if ($null -eq $Row) { return $false }
    if (Test-IsInternalRoleCache $Row) { return $false }
    if ($Row -is [System.Collections.Generic.KeyValuePair[object, object]]) { return $false }
    if ($Row -is [System.Array] -and $Row -isnot [System.Management.Automation.PSCustomObject]) { return $false }
    if ($Row -is [System.Collections.IDictionary] -and $Row -isnot [System.Management.Automation.PSCustomObject]) {
        return $false
    }
    return $true
}

function Test-ShouldEmitMatrixRow {
    param($Row)

    if (-not (Test-IsMatrixRow $Row)) { return $false }
    $inh = Get-MatrixRowColumnValue $Row 'Inheritance'
    if ($inh -ne 'Inherited') { return $true }
    return (Test-KeepInheritedRow $Row)
}

function Get-PathItemId {
    param([string] $Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return 0 }
    if ($S.PathToItemId.ContainsKey($Path)) { return $S.PathToItemId[$Path] }
    return 0
}

function Get-ResolvedRoleNamesFromAssignment {
    param($Assignment)

    if ($null -eq $Assignment) { return $null }
    if (-not (Test-IsRoleAssignmentObject $Assignment)) { return $null }

    $raw = Get-RestPropertyValue $Assignment '_ResolvedRoleNames'
    if ($null -eq $raw) { return $null }
    if ($raw -is [string]) { return @($raw) }
    if (Test-IsInternalRoleCache $raw) { return @() }

    return @(Get-NormalizedRoleNames $raw)
}

function Get-MatrixRowColumnValue {
    param(
        $Row,
        [string] $ColumnName
    )

    if ($null -eq $Row -or [string]::IsNullOrWhiteSpace($ColumnName)) { return '' }
    if (Test-IsInternalRoleCache $Row) { return '' }
    if ($Row -is [System.Collections.Generic.KeyValuePair[object, object]]) { return '' }

    if ($Row -is [System.Collections.IDictionary] -and $Row -isnot [System.Management.Automation.PSCustomObject]) {
        if ($Row.Contains($ColumnName)) {
            $v = $Row[$ColumnName]
            if ($null -eq $v) { return '' }
            return [string]$v
        }
        return ''
    }

    $p = $Row.PSObject.Properties[$ColumnName]
    if ($null -eq $p -or $null -eq $p.Value) { return '' }
    return [string]$p.Value
}

function Get-NormalizedTemplateRows {
    param($Template)

    if ($null -eq $Template) { return @() }
    if (Test-IsInternalRoleCache $Template) {
        Write-Warning 'Apply-Template received the internal template cache dictionary; skipping inherited rows.'
        return @()
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($Template)) {
        if ($null -eq $entry) { continue }
        if (Test-IsInternalRoleCache $entry) { continue }
        if ($entry -is [System.Collections.Generic.KeyValuePair[object, object]]) { continue }

        if ($entry -is [System.Collections.IDictionary] -and $entry -isnot [System.Management.Automation.PSCustomObject]) {
            continue
        }

        if ($entry -is [System.Array] -and $entry -isnot [System.Management.Automation.PSCustomObject]) {
            foreach ($inner in @($entry)) {
                if ($null -eq $inner) { continue }
                if (Test-IsInternalRoleCache $inner) { continue }
                if (Test-IsAnyKeyValuePair $inner) { continue }
                if ($inner -is [System.Collections.IDictionary] -and $inner -isnot [System.Management.Automation.PSCustomObject]) { continue }
                $rows.Add($inner) | Out-Null
            }
            continue
        }

        $rows.Add($entry) | Out-Null
    }
    return @($rows)
}

function Add-RowsFromAssignments {
    param(
        [System.Collections.Generic.List[object]] $Rows,
        [string] $ItemPath, [string] $ItemType, [string] $Inheritance, [string] $Details,
        [string] $GivenThrough, [object[]] $Assignments,
        [switch] $ExpandMembers, [string] $ExpansionPrefix, [switch] $ForInheritedTemplate
    )

    $normalized = @()
    try {
        $normalized = @(Get-NormalizedAssignmentArray $Assignments)
    }
    catch {
        Write-Warning "Add-RowsFromAssignments could not normalize assignments: $($_.Exception.Message)"
        return
    }

    foreach ($a in $normalized) {
        try {
            $m = Get-RestPropertyValue $a 'Member'
            if (-not (Test-IsRestPrincipalObject $m)) { continue }

            $roles = Get-AssignmentRoles $a
            $title = [string](Get-RestPropertyValue $m 'Title')
            $login = [string](Get-RestPropertyValue $m 'LoginName')
            $email = [string](Get-RestPropertyValue $m 'Email')
            $principalType = [int](Get-RestPropertyValue $m 'PrincipalType')
            $memberId = [int](Get-RestPropertyValue $m 'Id')

            if ($ForInheritedTemplate) {
                if (Test-ListScopedPrincipal $login $title) { continue }
                $roles = @($roles | Where-Object { $_ -and -not $NonInheritableRoles.Contains($_) })
                if ($roles.Count -eq 0) { continue }
            }

            $isLink = Test-SharingLink $login $title
            $userGroup = $title
            $account = $login
            $rowDetails = $Details
            $given = if ($isLink) { 'Sharing Link' } else { $GivenThrough }

            if ($isLink) {
                $label = Get-SharingLinkLabel $login $title
                $userGroup = $label
                $account = $label
                if (-not $rowDetails) { $rowDetails = $label }
            }
            elseif ($principalType -eq 8) {
                Register-SharePointGroup $m
            }

            if ($isLink -and $ExpandGroups) {
                foreach ($u in (Get-GroupMembers $memberId)) {
                    if ([int](Get-RestPropertyValue $u 'PrincipalType') -ne 1) { continue }
                    $Rows.Add((New-MatrixRow -ItemPath $ItemPath -ItemType $ItemType -Inheritance $Inheritance `
                        -Details $rowDetails -UserGroup ([string](Get-RestPropertyValue $u 'Title')) -PrincipalType 'User' `
                        -AccountName ([string](Get-RestPropertyValue $u 'LoginName')) -GivenThrough 'Sharing Link' -Roles $roles `
                        -IsExternal:(Test-ExternalUser ([string](Get-RestPropertyValue $u 'LoginName')) ([string](Get-RestPropertyValue $u 'Email'))))) | Out-Null
                }
                continue
            }

            $Rows.Add((New-MatrixRow -ItemPath $ItemPath -ItemType $ItemType -Inheritance $Inheritance `
                -Details $rowDetails -UserGroup $userGroup -PrincipalType (Get-PrincipalLabel $principalType) `
                -AccountName $account -GivenThrough $given -Roles $roles `
                -IsExternal:(Test-ExternalUser $login $email))) | Out-Null

            if ($ExpandMembers -and ($principalType -eq 8) -and -not $isLink) {
                $pfx = if ($ExpansionPrefix) { "$ExpansionPrefix - $title" } else { $ExpansionPrefix }
                foreach ($u in (Get-GroupMembers $memberId)) {
                    if ([int](Get-RestPropertyValue $u 'PrincipalType') -ne 1) { continue }
                    $ul = [string](Get-RestPropertyValue $u 'LoginName')
                    $ut = [string](Get-RestPropertyValue $u 'Title')
                    if ($ForInheritedTemplate -and (Test-ListScopedPrincipal $ul $ut)) { continue }
                    $Rows.Add((New-MatrixRow -ItemPath $ItemPath -ItemType $ItemType -Inheritance 'Inherited' `
                        -Details $pfx -UserGroup $ut -PrincipalType 'User' -AccountName $ul `
                        -GivenThrough "$title Members" -Roles $roles `
                        -IsExternal:(Test-ExternalUser $ul ([string](Get-RestPropertyValue $u 'Email'))))) | Out-Null
                }
            }
        }
        catch {
            $assignmentId = Get-RestPropertyValue $a 'Id'
            Write-Warning "Add-RowsFromAssignments skipped assignment (Id=$assignmentId): $($_.Exception.Message)"
        }
    }
}

function Apply-Template {
    param([object[]] $Template, [string] $ItemPath, [string] $ItemType)

    $templateRows = Get-NormalizedTemplateRows $Template
    if ($templateRows.Count -eq 0) { return @() }

    $out = [System.Collections.Generic.List[object]]::new($templateRows.Count)
    foreach ($t in $templateRows) {
        if (Test-IsInternalRoleCache $t) { continue }

        $resolvedItemType = if ($ItemType) { $ItemType } else { (Get-MatrixRowColumnValue $t 'Item Type') }
        $row = [ordered]@{
            'Site Name'      = $S.SiteName
            'Name'           = (Get-MatrixItemName $ItemPath $resolvedItemType $S.SiteName)
            'Item path'      = $ItemPath
            'Item Type'      = $resolvedItemType
            'Inheritance'    = (Get-MatrixRowColumnValue $t 'Inheritance')
            'Details'        = (Get-MatrixRowColumnValue $t 'Details')
            'User/group'     = (Get-MatrixRowColumnValue $t 'User/group')
            'Principal type' = (Get-MatrixRowColumnValue $t 'Principal type')
            'Account name'   = (Get-MatrixRowColumnValue $t 'Account name')
            'External user'  = (Get-MatrixRowColumnValue $t 'External user')
            'Given through'  = (Get-MatrixRowColumnValue $t 'Given through')
        }
        foreach ($col in $S.RoleColumns) {
            $row[$col] = (Get-MatrixRowColumnValue $t ([string]$col))
        }
        $out.Add([pscustomobject]$row) | Out-Null
    }
    return @($out)
}

function Build-Template {
    param(
        [string] $SourcePath, [string] $ListId, [int] $SourceItemId,
        [switch] $UseWebAssignments
    )

    $key = if ($UseWebAssignments) { "web:$SourcePath" } else { $SourcePath }
    if ($S.TemplateCache.ContainsKey($key)) { return @($S.TemplateCache[$key]) }

    $assignments = if ($UseWebAssignments) {
        Get-RoleAssignments 'web'
    } elseif ($SourceItemId -gt 0) {
        Get-RoleAssignments "$ListId`:$SourceItemId" $ListId $SourceItemId
    } else {
        Get-RoleAssignments "$ListId`:list" $ListId
    }

    $buf = [System.Collections.Generic.List[object]]::new()
    $pfx = "Inherited from $SourcePath"
    Add-RowsFromAssignments -Rows $buf -ItemPath '' -ItemType '' -Inheritance 'Inherited' `
        -Details $pfx -GivenThrough 'Explicit' -Assignments $assignments `
        -ExpandMembers:$ExpandGroups -ExpansionPrefix $pfx -ForInheritedTemplate:$(-not $UseWebAssignments)

    $S.TemplateCache[$key] = @($buf | Where-Object { Test-ShouldEmitMatrixRow $_ })
    return @($S.TemplateCache[$key])
}

function Get-SiteInheritedTemplate {
    if ($null -ne $S.SiteInheritedTpl) { return @($S.SiteInheritedTpl) }
    $key = "site:$($S.WebPath)"
    if ($S.TemplateCache.ContainsKey($key)) {
        $S.SiteInheritedTpl = @($S.TemplateCache[$key])
        return @($S.SiteInheritedTpl)
    }
    $buf = [System.Collections.Generic.List[object]]::new()
    $pfx = "Inherited from $($S.WebPath)"
    Add-RowsFromAssignments -Rows $buf -ItemPath '' -ItemType '' -Inheritance 'Inherited' `
        -Details $pfx -GivenThrough 'Explicit' -Assignments (Get-RoleAssignments 'web') `
        -ExpandMembers:$ExpandGroups -ExpansionPrefix $pfx
    $S.SiteInheritedTpl = @($buf | Where-Object { Test-ShouldEmitMatrixRow $_ })
    $S.TemplateCache[$key] = @($S.SiteInheritedTpl)
    return @($S.SiteInheritedTpl)
}

function Get-ParentPath([string] $Path) {
    if (-not $Path) { return $null }
    $n = $Path.TrimEnd('/')
    $i = $n.LastIndexOf('/')
    if ($i -le 0) { return $null }
    return $n.Substring(0, $i)
}

function Get-InheritanceSource([string] $ItemPath, [string] $ListRootUrl) {
    $parent = Get-ParentPath $ItemPath
    while ($parent -and $parent.Length -ge $ListRootUrl.Length) {
        if ($S.UniquePermPaths.Contains($parent)) { return $parent }
        if ($parent -eq $ListRootUrl) { break }
        $parent = Get-ParentPath $parent
    }
    return $ListRootUrl
}

function Get-ContentItemType([string] $ListBaseType, $FSObjType) {
    if ($ListBaseType -eq 'DocumentLibrary') {
        $typeVal = $FSObjType
        if ($null -eq $typeVal) { $typeVal = 0 }
        if ([int]$typeVal -eq 1) { return 'Folder' }
        return 'File'
    }
    return 'List item'
}

function Get-ListRootItemType([string] $ListBaseType) {
    if ($ListBaseType -eq 'DocumentLibrary') { return 'Library' }
    return 'List'
}

function Test-IsSubsiteScan {
    return ($S.WebItemType -eq 'Subsite')
}

function Get-SiteRootRows {
    if ((Test-IsSubsiteScan) -and -not $S.WebHasUnique) { return @() }

    $buf = [System.Collections.Generic.List[object]]::new()
    $inh = if ($S.WebHasUnique) { 'Custom' } else { 'Top level' }
    Add-RowsFromAssignments -Rows $buf -ItemPath $S.WebPath -ItemType $S.WebItemType -Inheritance $inh `
        -Details '' -GivenThrough 'Explicit' -Assignments (Get-RoleAssignments 'web') -ExpandMembers:$ExpandGroups
    return @($buf)
}

function Get-ListRootRows([string] $ListId, [string] $ListUrl, [string] $ListBaseType) {
    $listUnique = $S.ListHasUnique[$ListId]
    $itemType = Get-ListRootItemType $ListBaseType
    if (-not $listUnique) {
        if (-not $IncludeAllInheritedItemsInMatrix) { return @() }
        $inheritedRows = [System.Collections.Generic.List[object]]::new()
        foreach ($r in (Apply-Template (Get-SiteInheritedTemplate) $ListUrl $itemType)) { $inheritedRows.Add($r) | Out-Null }
        return @($inheritedRows | Where-Object { Test-ShouldEmitMatrixRow $_ })
    }

    $buf = [System.Collections.Generic.List[object]]::new()

    $listAssignments = @()
    try {
        $listAssignments = @(Get-RoleAssignments "$ListId`:list" $ListId)
    }
    catch {
        Write-Warning "List role assignments unavailable for $ListUrl : $($_.Exception.Message)"
    }

    try {
        Add-RowsFromAssignments -Rows $buf -ItemPath $ListUrl -ItemType $itemType -Inheritance 'Custom' `
            -Details '' -GivenThrough 'Explicit' -Assignments $listAssignments `
            -ExpandMembers:$ExpandGroups
    }
    catch {
        throw "Get-ListRootRows explicit assignments failed (listId=$ListId, url=$ListUrl): $($_.Exception.Message)"
    }

    return @($buf | Where-Object { Test-ShouldEmitMatrixRow $_ })
}

function Get-ItemRows {
    param(
        [string] $ItemPath, [string] $ListId, [int] $ItemId,
        [string] $ListRootUrl, [string] $ItemType, [bool] $HasUnique
    )

    $listUnique = $S.ListHasUnique[$ListId]
    $buf = [System.Collections.Generic.List[object]]::new()

    if ($HasUnique) {
        Add-RowsFromAssignments -Rows $buf -ItemPath $ItemPath -ItemType $ItemType -Inheritance 'Custom' `
            -Details '' -GivenThrough 'Explicit' -Assignments (Get-RoleAssignments "$ListId`:$ItemId" $ListId $ItemId)

        if ($ExpandGroups) {
            $src = Get-InheritanceSource $ItemPath $ListRootUrl
            $inheritsFromSiteLevel = ($src -eq $ListRootUrl -and -not $listUnique)
            if (-not $inheritsFromSiteLevel) {
                $srcId = if ($src -ne $ListRootUrl) { Get-PathItemId $src } else { 0 }
                $parent = if ($srcId -gt 0) {
                    Get-RoleAssignments "$ListId`:$srcId" $ListId $srcId
                } else {
                    Get-RoleAssignments "$ListId`:list" $ListId
                }
                $pfx = "Inherited from $src"
                Add-RowsFromAssignments -Rows $buf -ItemPath $ItemPath -ItemType $ItemType -Inheritance 'Inherited' `
                    -Details $pfx -GivenThrough 'Explicit' -Assignments $parent -ExpandMembers:$ExpandGroups -ForInheritedTemplate
            }
        }
        return @($buf | Where-Object { Test-ShouldEmitMatrixRow $_ })
    }

    if (-not $IncludeAllInheritedItemsInMatrix) { return @() }

    $src = Get-InheritanceSource $ItemPath $ListRootUrl
    $srcId = if ($src -ne $ListRootUrl) { Get-PathItemId $src } else { 0 }
    $useWeb = ($src -eq $ListRootUrl -and -not $listUnique -and $srcId -eq 0)
    $tpl = if ($useWeb) {
        Build-Template $S.WebPath $ListId 0 -UseWebAssignments
    } else {
        Build-Template $src $ListId $srcId
    }
    foreach ($r in (Apply-Template $tpl $ItemPath $ItemType)) { $buf.Add($r) | Out-Null }
    return @($buf)
}

function Get-ListItemScanViewXml {
    param(
        [int]    $RowLimit,
        [switch] $Minimal,
        [switch] $UniquePermsOnly
    )

    $fields = [System.Collections.Generic.List[string]]::new()
    [void]$fields.Add('ID')
    [void]$fields.Add('FileRef')
    [void]$fields.Add('FileLeafRef')
    [void]$fields.Add('FSObjType')
    if (-not $Minimal) { [void]$fields.Add('Title') }
    [void]$fields.Add('Created')
    [void]$fields.Add('Modified')
    [void]$fields.Add('Author')
    [void]$fields.Add('Editor')
    $fieldXml = ($fields | ForEach-Object { "    <FieldRef Name='$_'/>" }) -join [Environment]::NewLine

    $whereXml = if ($UniquePermsOnly) {
        "  <Query><Where><Eq><FieldRef Name='HasUniqueRoleAssignments'/><Value Type='Boolean'>1</Value></Eq></Where></Query>"
    } else {
        "  <Query></Query>"
    }

    return @"
<View Scope='RecursiveAll'>
  <ViewFields>
$fieldXml
  </ViewFields>
$whereXml
  <RowLimit Paged='TRUE'>$RowLimit</RowLimit>
</View>
"@
}

# Returns $true if the list has at least one item with unique permissions,
# $false if confirmed none do, or $null if the probe could not be completed
# (e.g. the list type doesn't support CAML filtering on HasUniqueRoleAssignments).
function Test-ListHasAnyUniquePermItems {
    param([object] $List, [string] $ListId)

    try {
        $probeXml = Get-ListItemScanViewXml -RowLimit 1 -Minimal -UniquePermsOnly
        $probe = @(Get-PnPListItem -List $List -Query $probeXml -PageSize 1 -ErrorAction Stop)
        return ($probe.Count -gt 0)
    }
    catch {
        Write-Verbose "HasUniqueRoleAssignments CAML probe failed for list $ListId : $_"
        return $null
    }
}

function Invoke-CsomListItemScanPage {
    param(
        [object[]] $Batch,
        [System.Collections.Generic.List[object]] $AllRows,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [bool]   $SkipNonUniqueMatrixRows,
        [bool]   $AllItemsInherit = $false,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems,
        [int]    $PageNum,
        [string] $BackendLabel
    )

    if (-not $AllItemsInherit) {
        Ensure-ListItemsUniquePermLoaded -Items $Batch
        $uniqueInBatch = @($Batch | Where-Object { Get-ItemHasUniqueRoleAssignments $_ }).Count
        if ($uniqueInBatch -gt 0) {
            Write-MatrixLog ("       ... {0} page {1}: {2:N0} unique-perm item(s) in batch" -f $BackendLabel, $PageNum, $uniqueInBatch) -Level Pulse
        }
    }
    return (Invoke-ListItemScanPage -Batch $Batch -AllRows $AllRows -ListId $ListId `
        -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
        -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$AllItemsInherit `
        -SiteFiles $SiteFiles -SiteFolders $SiteFolders `
        -SiteListItems $SiteListItems -PageNum $PageNum -BackendLabel $BackendLabel)
}

function Invoke-ListItemScanPage {
    param(
        [object[]] $Batch,
        [System.Collections.Generic.List[object]] $AllRows,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [bool]   $SkipNonUniqueMatrixRows,
        [bool]   $AllItemsInherit = $false,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems,
        [int]    $PageNum,
        [string] $BackendLabel
    )

    if (@($Batch).Count -lt 1) { return 0 }

    $S.ProgressListPageNum = $PageNum
    $S.ProgressPhase = if ($S.ProgressListPageCount -gt 1) { "$BackendLabel page $PageNum/$($S.ProgressListPageCount)" } else { 'Scanning items' }
    Write-ProgressStatus -Force
    Write-ScanHeartbeat -Force
    Write-MatrixLog ("       ... {0} page {1}: received {2:N0} item(s)" -f $BackendLabel, $PageNum, @($Batch).Count) -Level Pulse

    return (Process-ListItemPage -Batch $Batch -AllRows $AllRows -ListId $ListId `
        -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
        -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$AllItemsInherit `
        -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems)
}

function Scan-ListItemsCsomViewXml {
    param(
        [object] $List,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [int]    $ListItemCount,
        [bool]   $SkipNonUniqueMatrixRows,
        [bool]   $AllItemsInherit = $false,
        [System.Collections.Generic.List[object]] $AllRows,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $pageSize = Get-ActiveListPageSize
    if ($AllItemsInherit) {
        Write-MatrixLog ("       scanning {0:N0} items via CSOM ViewXml (~{1:N0} page(s) of {2:N0}) - no unique-perm items (probe confirmed)" -f `
            $ListItemCount, $S.ProgressListPageCount, $pageSize) -Level List
    }
    else {
        Write-MatrixLog ("       scanning {0:N0} items via CSOM ViewXml (~{1:N0} page(s) of {2:N0})..." -f `
            $ListItemCount, $S.ProgressListPageCount, $pageSize) -Level List
    }

    $S.ProgressListTitle = $ListTitle
    $S.ProgressPhase = 'CSOM loading'
    Write-ProgressStatus -Force

    $viewXml = Get-ListItemScanViewXml -RowLimit $pageSize -Minimal:$SkipNonUniqueMatrixRows
    $itemsScanned = 0
    $pageNum = 0
    $buffer = [System.Collections.Generic.List[object]]::new()
    $collectCount = 0
    $heartbeatEvery = 250

    foreach ($item in Get-PnPListItem -List $List -Query $viewXml -PageSize $pageSize -ErrorAction Stop) {
        $buffer.Add($item) | Out-Null
        $collectCount++

        if ($collectCount % $heartbeatEvery -eq 0) {
            $S.ProgressPhase = "CSOM loading ($collectCount/$ListItemCount)"
            Write-ScanProgressConsole -Force
        }

        if ($buffer.Count -ge $pageSize) {
            $pageNum++
            $itemsScanned += Invoke-CsomListItemScanPage -Batch @($buffer) -AllRows $AllRows -ListId $ListId `
                -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
                -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$AllItemsInherit `
                -SiteFiles $SiteFiles -SiteFolders $SiteFolders `
                -SiteListItems $SiteListItems -PageNum $pageNum -BackendLabel 'CSOM'
            $buffer.Clear()
            $S.ProgressPhase = 'CSOM loading'
        }
    }

    if ($buffer.Count -gt 0) {
        $pageNum++
        $itemsScanned += Invoke-CsomListItemScanPage -Batch @($buffer) -AllRows $AllRows -ListId $ListId `
            -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
            -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$AllItemsInherit `
            -SiteFiles $SiteFiles -SiteFolders $SiteFolders `
            -SiteListItems $SiteListItems -PageNum $pageNum -BackendLabel 'CSOM'
    }

    return $itemsScanned
}

function Scan-ListItemsCsomLegacy {
    param(
        [object] $List,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [int]    $ListItemCount,
        [bool]   $SkipNonUniqueMatrixRows,
        [System.Collections.Generic.List[object]] $AllRows,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $pageSize = Get-ActiveListPageSize
    Write-MatrixLog ("       scanning {0:N0} items via CSOM legacy batch (2 round-trips per page)..." -f `
        $ListItemCount) -Level List

    $itemsScanned = 0
    $buffer = [System.Collections.Generic.List[object]]::new()
    $bufferSize = [Math]::Max(100, [Math]::Min(500, $pageSize))
    $scanState = @{ PageNum = 0 }

    function Flush-CsomLegacyBuffer {
        if ($buffer.Count -lt 1) { return 0 }

        $scanState.PageNum++
        $count = Invoke-CsomListItemScanPage -Batch @($buffer) -AllRows $AllRows -ListId $ListId `
            -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
            -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -SiteFiles $SiteFiles -SiteFolders $SiteFolders `
            -SiteListItems $SiteListItems -PageNum $scanState.PageNum -BackendLabel 'CSOM legacy'
        $buffer.Clear()
        return $count
    }

    foreach ($item in Get-PnPListItem -List $List -PageSize $pageSize `
        -Fields 'ID', 'FileRef', 'FileLeafRef', 'Title', 'FSObjType') {
        $buffer.Add($item) | Out-Null
        if ($buffer.Count -ge $bufferSize) {
            $itemsScanned += Flush-CsomLegacyBuffer
        }
    }

    $itemsScanned += Flush-CsomLegacyBuffer
    return $itemsScanned
}

function Scan-ListItemsCsom {
    param(
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [int]    $ListItemCount,
        [bool]   $SkipNonUniqueMatrixRows,
        [bool]   $AllItemsInherit = $false,
        [object] $ListObject = $null,
        [System.Collections.Generic.List[object]] $AllRows,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $backendSw = [System.Diagnostics.Stopwatch]::StartNew()
    $list = if ($ListObject) { $ListObject } else { Get-PnPList -Identity $ListId -ErrorAction Stop }
    $itemsScanned = 0

    try {
        $itemsScanned = Scan-ListItemsCsomViewXml -List $list -ListId $ListId -ListTitle $ListTitle `
            -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType -ListItemCount $ListItemCount `
            -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$AllItemsInherit -AllRows $AllRows `
            -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
    }
    catch {
        Write-Warning "CSOM ViewXml scan failed for '$ListTitle': $($_.Exception.Message)"
        Write-MatrixLog '       falling back to CSOM legacy batch load...' -Level Warn
        $itemsScanned = Scan-ListItemsCsomLegacy -List $list -ListId $ListId -ListTitle $ListTitle `
            -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType -ListItemCount $ListItemCount `
            -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllRows $AllRows `
            -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
    }

    $backendSw.Stop()
    Add-ScanBackendTiming -Backend 'Csom' -Items $itemsScanned -Elapsed $backendSw.Elapsed

    Write-MatrixLog ("       done: {0:N0} item(s) scanned, {1:N0} with unique permissions" -f `
        $itemsScanned, $S.ProgressListUniqueFound) -Level $(if ($S.ProgressListUniqueFound -gt 0) { 'Unique' } else { 'Dim' })
}

function Scan-ListItemsRest {
    param(
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [int]    $ListItemCount,
        [bool]   $SkipNonUniqueMatrixRows,
        [bool]   $AllItemsInherit = $false,
        [System.Collections.Generic.List[object]] $AllRows,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    $backendSw = [System.Diagnostics.Stopwatch]::StartNew()

    $pageSize = Get-ActiveListPageSize

    if ($AllItemsInherit) {
        Write-MatrixLog ("       scanning {0:N0} items (~{1:N0} REST page(s) of {2:N0}) - no unique-perm items (probe confirmed)" -f `
            $ListItemCount, $S.ProgressListPageCount, $pageSize) -Level List
    }
    elseif ($SkipNonUniqueMatrixRows) {
        Write-MatrixLog ("       scanning all {0:N0} items (~{1:N0} REST page(s) of {2:N0}) - matrix rows only where HasUniqueRoleAssignments is true" -f `
            $ListItemCount, $S.ProgressListPageCount, $pageSize) -Level List
    }
    else {
        Write-MatrixLog ("       scanning {0:N0} items (~{1:N0} REST page(s) of {2:N0})..." -f `
            $ListItemCount, $S.ProgressListPageCount, $pageSize) -Level List
    }

    $S.ProgressListTitle = $ListTitle
    Write-ProgressStatus -Force

    if ($ListItemCount -lt 1) {
        Write-MatrixLog '       done: 0 item(s) scanned, 0 with unique permissions' -Level Dim
        return
    }

    $selectFull = Get-ListItemSelectFields -IncludeUniqueField (-not $AllItemsInherit) -Minimal:$SkipNonUniqueMatrixRows -Profile Full
    $selectCore = Get-ListItemSelectFields -IncludeUniqueField (-not $AllItemsInherit) -Minimal:$true -Profile Core
    $useCoreSelect = $false
    $expandAuthorEditor = ($selectFull -match 'Author|Editor')
    $url = New-ListItemsRestPageUrl -ListId $ListId -Select $selectFull -Top $pageSize -ExpandAuthorEditor:$expandAuthorEditor
    $pageNum = 0
    $itemsFromApi = 0

    while ($url) {
        $pageNum++
        $S.ProgressListPageNum = $pageNum
        $S.ProgressPhase = if ($S.ProgressListPageCount -gt 1) { "REST page $pageNum/$($S.ProgressListPageCount)" } else { 'Scanning items' }
        Write-ProgressStatus -Force
        Write-ScanHeartbeat -Force

        try {
            $page = Invoke-SPRestGetWithRetry $url
        }
        catch {
            if (-not $useCoreSelect -and (Test-RestMissingListFieldError $_)) {
                Write-Warning "List '$ListTitle': optional columns missing; retrying item scan with core fields only."
                $useCoreSelect = $true
                $pageNum = 0
                $itemsFromApi = 0
                $url = New-ListItemsRestPageUrl -ListId $ListId -Select $selectCore -Top $pageSize
                continue
            }
            throw
        }

        $batch = @($page.value)
        if ($batch.Count -lt 1) { break }

        if ($batch.Count -gt 0 -and (-not $SkipNonUniqueMatrixRows -or $pageNum -eq 1)) {
            Write-MatrixLog ("       ... REST page {0}: received {1:N0} item(s)" -f $pageNum, $batch.Count) -Level Pulse
        }

        $itemsFromApi += Process-ListItemPage -Batch $batch -AllRows $AllRows -ListId $ListId `
            -ListTitle $ListTitle -ListUrl $ListUrl -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType `
            -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$AllItemsInherit `
            -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems

        $url = Get-ODataNextLink $page
    }

    if ($ListItemCount -gt 0 -and $itemsFromApi -lt 1) {
        Write-Warning "List '$ListTitle' returned 0 items from REST but reports ItemCount=$ListItemCount."
    }
    elseif ($ListItemCount -gt 0 -and [Math]::Abs($ListItemCount - $itemsFromApi) -gt [Math]::Max(50, $ListItemCount * 0.05)) {
        Write-Warning ("List '$ListTitle' ItemCount=$ListItemCount but REST returned $itemsFromApi item(s). Unique-perm detection may be incomplete.")
    }

    $backendSw.Stop()
    Add-ScanBackendTiming -Backend 'Rest' -Items $itemsFromApi -Elapsed $backendSw.Elapsed

    Write-MatrixLog ("       done: {0:N0} item(s) scanned, {1:N0} with unique permissions" -f `
        $itemsFromApi, $S.ProgressListUniqueFound) -Level $(if ($S.ProgressListUniqueFound -gt 0) { 'Unique' } else { 'Dim' })
}

function Resolve-ListItemScanBackend {
    param(
        [string] $ListId,
        [int]    $ListItemCount,
        [string] $ListBaseType = ''
    )

    if ($ItemScanMode -eq 'Csom') { return 'Csom' }

    $restOk = Test-ListSupportsUniquePermField -ListId $ListId

    if ($ItemScanMode -eq 'Rest') {
        if ($restOk) { return 'Rest' }
        return 'Csom'
    }

    if (-not $restOk) {
        return 'Csom'
    }

    # Auto mode: start with REST for large scans (best initial bet for SPO),
    # then allow measured throughput to steer future large lists.
    if ($ItemScanMode -eq 'Auto' -and $ListItemCount -ge $LargeListRestThreshold) {
        $timed = Get-PreferredBackendByTiming -ListItemCount $ListItemCount -Default 'Rest'
        if ($timed) { return $timed }
        return 'Rest'
    }

    if ($ItemScanMode -eq 'Auto' -and $ListBaseType -eq 'DocumentLibrary' -and $ListItemCount -ge 250) {
        return 'Rest'
    }

    return 'Csom'
}

function Scan-ListItems {
    param(
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [int]    $ListItemCount,
        [bool]   $SkipNonUniqueMatrixRows,
        [object] $List = $null,
        [System.Collections.Generic.List[object]] $AllRows,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems
    )

    Initialize-ListProgress -ListTitle $ListTitle -ListItemCount $ListItemCount

    $Script:SharingQueue.Clear()
    $backend = Resolve-ListItemScanBackend -ListId $ListId -ListItemCount $ListItemCount -ListBaseType $ListBaseType
    $script:CurrentListPageSize = Get-EffectiveListPageSize -ListItemCount $ListItemCount -Backend $backend
    $activePageSize = Get-ActiveListPageSize
    $estPages = [Math]::Max(1, [Math]::Ceiling($ListItemCount / [double]$activePageSize))
    $S.ProgressListPageCount = $estPages
    $S.ProgressListPageNum = 0

    if ($activePageSize -ne $ListItemPageSize) {
        Write-MatrixLog ("       auto-tuned page size: {0:N0}" -f $activePageSize) -Level Dim
    }

    if ($backend -eq 'Rest' -and $ItemScanMode -eq 'Auto') {
        Write-MatrixLog ("       auto mode: using REST item scan for '{0}' ({1:N0} items, page size {2:N0})" -f `
            $ListTitle, $ListItemCount, $activePageSize) -Level Dim
    }
    elseif ($backend -eq 'Csom' -and $ItemScanMode -eq 'Auto' -and $ListItemCount -ge $LargeListRestThreshold) {
        Write-MatrixLog ("       auto mode: keeping CSOM for '{0}' based on observed throughput (page size {1:N0})" -f `
            $ListTitle, $activePageSize) -Level Dim
    }

    # Probe: ask SharePoint for 1 item with HasUniqueRoleAssignments=1 via CAML.
    # If none exist we can skip the HasUniqueRoleAssignments field entirely during the
    # full scan — smaller REST payloads and no per-item unique-perm processing overhead.
    $allItemsInherit = $false
    if ($ListItemCount -gt 0 -and $List) {
        $probeResult = Test-ListHasAnyUniquePermItems -List $List -ListId $ListId
        if ($probeResult -eq $false) {
            $allItemsInherit = $true
            Write-MatrixLog '       probe: no unique-perm items in this list — skipping HasUniqueRoleAssignments scan' -Level Dim
        }
    }

    try {
        if ($ListItemCount -lt 1) {
            Write-MatrixLog '       done: 0 item(s) scanned, 0 with unique permissions' -Level Dim
        }
        elseif ($backend -eq 'Rest') {
            Scan-ListItemsRest -ListId $ListId -ListTitle $ListTitle -ListUrl $ListUrl `
                -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType -ListItemCount $ListItemCount `
                -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$allItemsInherit `
                -AllRows $AllRows -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
        }
        else {
            try {
                Scan-ListItemsCsom -ListId $ListId -ListTitle $ListTitle -ListUrl $ListUrl `
                    -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType -ListItemCount $ListItemCount `
                    -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$allItemsInherit `
                    -ListObject $List -AllRows $AllRows `
                    -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
            }
            catch {
                if ((Test-ListSupportsUniquePermField -ListId $ListId)) {
                    Write-Warning "CSOM scan failed for '$ListTitle': $($_.Exception.Message)"
                    Write-MatrixLog '       falling back to REST for this list...' -Level Warn
                    Scan-ListItemsRest -ListId $ListId -ListTitle $ListTitle -ListUrl $ListUrl `
                        -ListUrlNorm $ListUrlNorm -ListBaseType $ListBaseType -ListItemCount $ListItemCount `
                        -SkipNonUniqueMatrixRows:$SkipNonUniqueMatrixRows -AllItemsInherit:$allItemsInherit `
                        -AllRows $AllRows -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems
                }
                else { throw }
            }
        }
    }
    catch {
        Write-Warning "Item scan failed for list '$ListTitle' ($ListUrl): $($_.Exception.Message). Continuing with next list."
    }
    finally {
        $script:CurrentListPageSize = 0
    }

    Flush-SharingQueue -ListTitle $ListTitle
}


function Process-ListItem {
    param(
        $Item,
        [System.Collections.Generic.List[object]] $AllRows,
        [string] $ListId,
        [string] $ListTitle,
        [string] $ListUrl,
        [string] $ListUrlNorm,
        [string] $ListBaseType,
        [bool]   $SkipNonUniqueMatrixRows,
        [ref]    $SiteFiles,
        [ref]    $SiteFolders,
        [ref]    $SiteListItems,
        [switch] $UniqueOnlyProcessing
    )

    if ($UniqueOnlyProcessing) {
        $path = [string](Get-ScanItemField $Item 'FileRef')
        if (-not $path -or $path.TrimEnd('/') -eq $ListUrlNorm) { return }

        $itemType = Get-ContentItemType $ListBaseType (Get-ScanItemFsObjType $Item)
        Add-Rows $AllRows (Get-ItemRows $path $ListId (Get-ScanItemId $Item) $ListUrl $itemType $true)
        Queue-SharingItem -Item $Item -ListId $ListId -ListTitle $ListTitle -ListUrl $ListUrl `
            -ListBaseType $ListBaseType -ItemType $itemType
        return
    }

    Bump-Progress $ListTitle

    $path = [string](Get-ScanItemField $Item 'FileRef')
    if (-not $path -or $path.TrimEnd('/') -eq $ListUrlNorm) { return }

    $itemType = Get-ContentItemType $ListBaseType (Get-ScanItemFsObjType $Item)
    Add-ListItemScanStats -ListBaseType $ListBaseType -ItemType $itemType `
        -SiteFiles $SiteFiles -SiteFolders $SiteFolders -SiteListItems $SiteListItems

    $hasUnique = Get-ItemHasUniqueRoleAssignments $Item
    if ($hasUnique) {
        [void]$Stats.UniqueItemPaths.Add($path)
        $S.ProgressListUniqueFound++
    }
    Add-ItemRowFromObject -Item $Item -ListTitle $ListTitle -ItemPath $path `
        -ItemType $itemType -HasUnique $hasUnique

    if (-not $hasUnique -and (-not $IncludeAllInheritedItemsInMatrix -or $SkipNonUniqueMatrixRows)) {
        $Stats.BulkSkippedItems++
        return
    }

    Add-Rows $AllRows (Get-ItemRows $path $ListId ([int](Get-ScanItemId $Item)) $ListUrl $itemType $hasUnique)

    if ($hasUnique) {
        Queue-SharingItem -Item $Item -ListId $ListId -ListTitle $ListTitle -ListUrl $ListUrl `
            -ListBaseType $ListBaseType -ItemType $itemType
    }
}

function Scan-Site {
    param(
        [System.Collections.Generic.List[object]] $AllRows,
        [object[]] $Lists,
        [ref] $UniqueBoundaryCount
    )

    if ($S.WebHasUnique) { $Stats.SitesWithUniquePerms++ }

    $itemTotal = 0
    foreach ($l in $Lists) { $itemTotal += [int]$l.ItemCount }

    Write-MatrixLog ("Found {0} lists/libraries, {1:N0} items in this site." -f @($Lists).Count, $itemTotal) -Level Site

    $siteRows = Get-SiteRootRows
    Add-Rows $AllRows $siteRows
    Write-MatrixLog ("  Site-level rows: {0:N0}" -f $siteRows.Count) -Level Dim

    $null = Get-SiteInheritedTemplate
    $siteFiles = 0
    $siteFolders = 0
    $siteListItems = 0

    foreach ($list in $Lists) {
        $listUrl = [string]$list.RootFolder.ServerRelativeUrl
        $listId = $list.Id.ToString().ToLower()
        $baseType = [string]$list.BaseType
        $listUrlNorm = $listUrl.TrimEnd('/')
        $S.ListHasUnique[$listId] = [bool]$list.HasUniqueRoleAssignments

        if ($baseType -eq 'DocumentLibrary') { $Stats.Libraries++ } else { $Stats.Lists++ }
        if ($list.HasUniqueRoleAssignments) { $Stats.ListsWithUniquePerms++ }

        Write-MatrixLog ("  {0,-50} ({1:N0} items)" -f $list.Title, $list.ItemCount) -Level Lib
        $S.ProgressListTitle = $list.Title
        Write-ProgressStatus -Force

        try {
            Add-Rows $AllRows (Get-ListRootRows $listId $listUrl $baseType)
        }
        catch {
            throw "List matrix root failed for '$($list.Title)' ($listUrl): $($_.Exception.Message)"
        }

        $listItemCount = [int]$list.ItemCount
        $skipNonUniqueMatrixRows = ($listItemCount -gt $MaxListItemsForFullMatrix -and -not $IncludeAllInheritedItemsInMatrix)

        if ($skipNonUniqueMatrixRows) {
            if ($list.HasUniqueRoleAssignments) {
                Write-MatrixLog ("       large list ({0:N0} items): list-level permissions are on the library row; scanning items for unique-permission rows only" -f `
                    $listItemCount) -Level Dim
            }
            else {
                Write-MatrixLog ("       large inherited list: matrix rows only for unique-permission items ({0:N0} items)" -f `
                    $listItemCount) -Level Dim
            }
        }
        elseif ($IncludeAllInheritedItemsInMatrix -and $listItemCount -gt $MaxListItemsForFullMatrix) {
            Write-MatrixLog ("       inherited-row mode enabled: emitting all inherited matrix rows for {0:N0} items" -f `
                $listItemCount) -Level Warn
        }

        Scan-ListItems -ListId $listId -ListTitle $list.Title -ListUrl $listUrl -ListUrlNorm $listUrlNorm `
            -ListBaseType $baseType -ListItemCount $listItemCount -SkipNonUniqueMatrixRows:$skipNonUniqueMatrixRows `
            -List $list -AllRows $AllRows -SiteFiles ([ref]$siteFiles) -SiteFolders ([ref]$siteFolders) `
            -SiteListItems ([ref]$siteListItems)
    }

    foreach ($p in $S.UniquePermPaths) { [void]$Stats.UniqueItemPaths.Add($p) }
    $UniqueBoundaryCount.Value += $S.UniquePermPaths.Count

    Add-GroupMemberRowsForSite

    $siteUniqueCount = $S.UniquePermPaths.Count
    $siteExternalCount = Get-ExternalUserCountBySite -SiteName $S.SiteName

    $Stats.SiteDetails.Add([pscustomobject]@{
        'Site Name'         = $S.SiteName
        'Site Path'         = $S.WebPath
        'Type'              = $S.WebItemType
        'Lists/Libraries'   = @($Lists).Count
        'Items'             = $itemTotal
        'Files'             = $siteFiles
        'Folders'           = $siteFolders
        'List items'        = $siteListItems
        'Unique items'      = $siteUniqueCount
        'External users'    = $siteExternalCount
        'Site unique perms' = if ($S.WebHasUnique) { 'Yes' } else { 'No' }
    }) | Out-Null

    return $itemTotal
}

function Export-Workbook {
    param(
        [System.Collections.Generic.List[object]] $MatrixRows,
        [System.Collections.Generic.List[object]] $AllItemsRows,
        [object[]] $SharingRows,
        [object[]] $GroupMemberRows,
        [object]   $Summary,
        [string]   $Path,
        [string]   $ReportSubtitle = '',
        [string[]] $RoleColumns = @()
    )

    $cols = @(
        'Site Name', 'Name', 'Item path', 'Item Type', 'Inheritance', 'Details',
        'User/group', 'Principal type', 'Account name', 'External user', 'Given through'
    ) + @($RoleColumns)

    $matrixCount = $MatrixRows.Count
    $exportSw = [System.Diagnostics.Stopwatch]::StartNew()
    $exportCompleted = $false
    Start-ProgressRenderer -Activity 'Writing Excel workbook'
    try {
    Write-ExportPhase 'Writing Summary sheet' 10
        $pkg = $Summary.Overview | Export-Excel -Path $Path -WorksheetName 'Summary' `
            -StartRow $SummaryOverviewStartRow -StartColumn (Get-SummaryCol 1) `
            -TableName 'SummaryOverview' -TableStyle 'None' `
            -BoldTopRow -ClearSheet -PassThru
        Write-MatrixLog ("  Summary sheet written ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export

        $siteDetails = [object[]]@($Summary.SiteDetails)
        $matrixExport = Get-MatrixExportRows -MatrixRows $MatrixRows -Columns $cols -RowCount $matrixCount

        $exportSw.Restart()
        Write-ExportPhase ("Writing Permissions Matrix ({0:N0} rows)" -f $matrixCount) 30
        if ($matrixCount -gt 50000) {
            Write-MatrixLog '  Large matrix export can take several minutes — please wait.' -Level Warn
        }
        $null = $matrixExport | Export-Excel -ExcelPackage $pkg `
            -WorksheetName 'Permissions Matrix' `
            -TableName 'PermissionsMatrix' -TableStyle 'None' `
            -FreezeTopRow -BoldTopRow -PassThru
        Write-MatrixLog ("  Permissions Matrix written ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export

        $exportSw.Restart()
        Write-ExportPhase ("Writing Group Members ({0:N0} rows)" -f $GroupMemberRows.Count) 55
        $null = $GroupMemberRows | Select-Object $GroupMemberColumns | Export-Excel -ExcelPackage $pkg `
            -WorksheetName 'Group Members' `
            -TableName 'GroupMembers' -TableStyle 'None' `
            -FreezeTopRow -BoldTopRow -PassThru
        Write-MatrixLog ("  Group Members written ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export

        $exportSw.Restart()
        Write-ExportPhase ("Writing Sharing Links ({0:N0} rows)" -f $SharingRows.Count) 75
        $null = $SharingRows | Select-Object $SharingLinkColumns | Export-Excel -ExcelPackage $pkg `
            -WorksheetName 'Sharing Links' `
            -TableName 'SharingLinks' -TableStyle 'None' `
            -FreezeTopRow -BoldTopRow -PassThru
        Write-MatrixLog ("  Sharing Links written ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export

        $exportSw.Restart()
        Write-ExportPhase ("Writing All Items ({0:N0} rows)" -f $AllItemsRows.Count) 88
        $null = $AllItemsRows | Select-Object $AllItemsColumns | Export-Excel -ExcelPackage $pkg `
            -WorksheetName 'All Items' `
            -TableName 'AllItems' -TableStyle 'None' `
            -FreezeTopRow -BoldTopRow -PassThru
        Write-MatrixLog ("  All Items written ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export

        $exportSw.Restart()
        Write-ExportPhase 'Applying dashboard theme' 95
        try {
            Apply-ModernWorkbookTheme -Package $pkg -RoleColumns $RoleColumns `
                -ReportSubtitle $ReportSubtitle -MatrixRowCount $matrixCount -SiteDetails $siteDetails
        }
        catch {
            Write-Warning "Dashboard theme partially failed: $_"
        }
        Write-MatrixLog ("  Theme applied ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export

        $wsSummary = $pkg.Workbook.Worksheets['Summary']
        if ($wsSummary) {
            try {
                Complete-SummaryDashboardPresentation -Worksheet $wsSummary
            }
            catch {
                Write-Warning "Summary presentation finalize skipped: $_"
            }
        }

        $embedId = if ($S.PnpClientId) { [string]$S.PnpClientId } else { '' }
        if ($embedId) {
            if (Set-ReportEmbeddedClientId -Package $pkg -ClientId $embedId) {
                Write-MatrixLog '  Stored connection profile for remediation (hidden sheet).' -Level Export
            }
            else {
                Write-MatrixLog '  Could not store connection profile in workbook.' -Level Warn
            }
        }

        $exportSw.Restart()
        Write-ExportPhase 'Saving workbook' 100
        if ($matrixCount -gt 50000) {
            Write-MatrixLog '  Saving a large .xlsx can take 5–15+ minutes — still working.' -Level Warn
        }
        Close-ExcelPackage $pkg
        Write-MatrixLog ("  Workbook saved ({0})" -f (Format-Elapsed $exportSw.Elapsed)) -Level Success
        $exportCompleted = $true
    }
    finally {
        Stop-ProgressRenderer -Complete:$exportCompleted
    }
}

# --- Picker ---

function Get-PermissionsMatrixScanPlan {
    param(
        [bool]        $IncludeSubsites,
        [scriptblock] $ProgressCallback = $null
    )

    if ($S.IsTenantAdminScope) {
        return @(Get-TenantPermissionsMatrixScanPlan -IncludeSubsites $IncludeSubsites -ProgressCallback $ProgressCallback)
    }

    $plan = [System.Collections.Generic.List[object]]::new()
    $S.TenantInventorySkipped = 0
    $webs = @(Get-WebsFromConnectedSiteRoot -IncludeSubsites $IncludeSubsites)
    $totalWebs = $webs.Count
    $idx = 0

    if ($ProgressCallback -and $totalWebs -gt 0) {
        & $ProgressCallback 0 $totalWebs 'Loading site and list inventory...'
    }

    foreach ($web in $webs) {
        $idx++
        $path = ''
        try { $path = [string]$web.ServerRelativeUrl } catch { }
        $title = ''
        try { $title = [string]$web.Title } catch { }
        $label = if ($title) { $title } elseif ($path) { $path } else { "web #$idx" }
        if ($ProgressCallback) {
            & $ProgressCallback $idx $totalWebs $label
        }
        try {
            $plan.Add((New-ScanPlanEntryForWeb -Web $web)) | Out-Null
        }
        catch {
            $S.TenantInventorySkipped++
            Write-Warning "Skipping list inventory for '$label': $($_.Exception.Message)"
        }
    }

    return @($plan)
}

function Get-PermissionsMatrixTimeEstimate {
    param(
        [object[]] $PlanEntries,
        [bool]     $ExpandGroups,
        [bool]     $IncludeFolderSharingLinks,
        [bool]     $IncludeAllInheritedItemsInMatrix = $false,
        [string]   $ItemScanMode = 'Auto'
    )

    $entries = @($PlanEntries)
    if ($entries.Count -lt 1) {
        return [pscustomobject]@{
            TotalSeconds         = 0
            ScanSeconds          = 0
            ExportSeconds        = 0
            EstimatedUniqueItems = 0
            TotalItems           = 0
            SiteCount            = 0
            ListCount            = 0
        }
    }

    $siteCount = $entries.Count
    $totalItems = 0
    $listCount = 0
    $estUnique = 0

    foreach ($entry in $entries) {
        $totalItems += [int]$entry.Items
        $listCount += @($entry.Lists).Count
        foreach ($list in @($entry.Lists)) {
            $ic = [int]$list.ItemCount
            if ([bool]$list.HasUniqueRoleAssignments) {
                $estUnique += [Math]::Min($ic, [Math]::Max(5, [int][Math]::Ceiling($ic * 0.025)))
            }
            else {
                $estUnique += [int][Math]::Ceiling($ic * 0.004)
            }
        }
    }

    if ($estUnique -lt 1 -and $totalItems -gt 0) { $estUnique = 1 }

    # Calibrated from large tenant scans: bulk REST ~8s/1k items; ~1.1s per unique-perm boundary fetch.
    $scanSec = ($siteCount * 26.0) + ($totalItems / 120.0) + ($estUnique * 1.1)
    if ($IncludeFolderSharingLinks) { $scanSec += $estUnique * 0.55 }
    if ($ExpandGroups) { $scanSec *= 1.35 }
    switch ($ItemScanMode) {
        'Rest' { $scanSec *= 0.9 }
        'Auto' { $scanSec *= 0.95 }
        'Csom' { $scanSec *= 1.15 }
        default { }
    }
    if ($IncludeAllInheritedItemsInMatrix) {
        $scanSec *= 1.65
    }

    $historicalModel = Get-RunMetricsModel `
        -Mode ([string]$ItemScanMode) `
        -IncludeAllInheritedItemsInMatrix $IncludeAllInheritedItemsInMatrix

    if ($historicalModel -and $historicalModel.Samples -ge 1) {
            $ips = switch ($ItemScanMode) {
                'Rest' { [double]$historicalModel.RestItemsPerSec }
                'Csom' { [double]$historicalModel.CsomItemsPerSec }
                default {
                    if ($totalItems -ge 250 -and [double]$historicalModel.RestItemsPerSec -gt 0) {
                        [double]$historicalModel.RestItemsPerSec
                    } else {
                        [Math]::Max([double]$historicalModel.CsomItemsPerSec, [double]$historicalModel.RestItemsPerSec)
                    }
                }
            }
            if ($ips -gt 0) {
                $overhead = [Math]::Max(8.0, [double]$historicalModel.SiteOverheadSec)
                $scanSec = ($siteCount * $overhead) + ($totalItems / $ips) + ($estUnique * 0.9)
                if ($IncludeFolderSharingLinks) { $scanSec += $estUnique * 0.45 }
                if ($ExpandGroups) { $scanSec *= 1.3 }
            }
    }

    $estMatrixRows = ($siteCount * 42.0) + ($listCount * 8.0) + ($estUnique * 1.2)
    $exportSec = 50.0 + ($estMatrixRows * 0.00085)
    if ($historicalModel -and $historicalModel.ExportRowsPerSec -gt 0 -and $historicalModel.Samples -ge 1) {
        $exportSec = [Math]::Max(20.0, $estMatrixRows / [double]$historicalModel.ExportRowsPerSec)
    }
    if ($IncludeAllInheritedItemsInMatrix) {
        $exportSec *= 1.8
    }
    if ($totalItems -gt 25000) { $exportSec += ($totalItems - 25000) / 400.0 }
    if ($estMatrixRows -gt 50000) { $exportSec += ($estMatrixRows - 50000) * 0.0025 }

    $totalSec = $scanSec + $exportSec
    if ($historicalModel -and $historicalModel.Samples -ge 1) {
        $low = [int][Math]::Max(30, [Math]::Floor($totalSec * 0.85))
        $high = [int][Math]::Ceiling($totalSec * 1.2)
    }
    else {
        $low = [int][Math]::Max(30, [Math]::Floor($totalSec * 0.75))
        $high = [int][Math]::Ceiling($totalSec * 1.35)
    }

    return [pscustomobject]@{
        TotalSeconds         = [int][Math]::Ceiling($totalSec)
        ScanSeconds          = [int][Math]::Ceiling($scanSec)
        ExportSeconds        = [int][Math]::Ceiling($exportSec)
        EstimatedUniqueItems = [int]$estUnique
        TotalItems           = [int]$totalItems
        SiteCount            = [int]$siteCount
        ListCount            = [int]$listCount
        RangeLowSeconds      = $low
        RangeHighSeconds     = $high
    }
}

function Format-SecondsEstimate {
    param([int] $Seconds)

    if ($Seconds -lt 60) { return "{0}s" -f $Seconds }
    $m = [int][Math]::Floor($Seconds / 60)
    $s = $Seconds % 60
    if ($m -lt 60) {
        if ($s -gt 0) { return "{0}m {1}s" -f $m, $s }
        return "{0}m" -f $m
    }
    $h = [int][Math]::Floor($m / 60)
    $m = $m % 60
    return "{0}h {1}m" -f $h, $m
}

function Show-PermissionsMatrixExportPicker {
    param(
        [string] $InitialSiteUrl = '',
        [string] $InitialClientId = '',
        [bool]   $DefaultExpandGroups = $true,
        [bool]   $DefaultIncludeSubsites = $true,
        [bool]   $DefaultIncludeFolderSharingLinks = $true,
        [bool]   $DefaultIncludeAllInheritedItemsInMatrix = $false,
        [bool]   $DefaultOpenWorkbookOnComplete = $true,
        [string] $DefaultItemScanMode = 'Rest',
        [int]    $DefaultMaxListItemsForFullMatrix = 2000,
        [int]    $DefaultListItemPageSize = 5000,
        [int]    $DefaultUniquePermProgressWeight = 12
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $script:PickerPlan             = @()
    $script:PickerSiteRows         = [System.Collections.Generic.List[object]]::new()
    $script:PickerSuppressCheck    = $false
    $script:PickerBlockCheckUntil  = [datetime]::MinValue

    # ── Rounded-button helper ──────────────────────────────────────────────────
    # Applies owner-drawn rounded corners to any WinForms Button. The button's own
    # rectangular background is erased by Clear() so only the rounded path shows.
    function Set-RoundedButton {
        param($Btn, [int]$Radius = 7, $Bg, $Fg = [System.Drawing.Color]::White,
              [bool]$Border = $false, $BorderClr = $null, $HoverBg = $null)
        $Btn.FlatStyle = 'Flat'
        $Btn.FlatAppearance.BorderSize = 0
        $Btn.BackColor = $Bg
        $Btn.ForeColor = $Fg
        # Derive hover color: darken filled buttons, use brand-faint for ghost
        $resolvedHover = if ($HoverBg) { $HoverBg }
                         elseif ($Border) { [System.Drawing.ColorTranslator]::FromHtml('#E4E6F5') }
                         else { [System.Drawing.Color]::FromArgb(
                                    [Math]::Max(0, $Bg.R - 22),
                                    [Math]::Max(0, $Bg.G - 22),
                                    [Math]::Max(0, $Bg.B - 22)) }
        $Btn.Tag = [PSCustomObject]@{
            R=$Radius; Bg=$Bg; Fg=$Fg; Brd=$Border; BrdClr=$BorderClr
            HoverBg=$resolvedHover; Hover=$false
        }
        $Btn.Add_MouseEnter({ param($s,$e); $s.Tag.Hover=$true;  $s.Invalidate() })
        $Btn.Add_MouseLeave({ param($s,$e); $s.Tag.Hover=$false; $s.Invalidate() })
        $Btn.Add_Paint({
            param($s, $e)
            $t  = $s.Tag
            $g  = $e.Graphics
            $g.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g.PixelOffsetMode  = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $parentBg = if ($s.Parent) { $s.Parent.BackColor } else { [System.Drawing.Color]::White }
            $g.Clear($parentBg)
            $fillClr = if (-not $s.Enabled) {
                [System.Drawing.ColorTranslator]::FromHtml('#C5C9D8')
            }
            elseif ($t.Hover) { $t.HoverBg }
            else { $t.Bg }
            $fgClr = if (-not $s.Enabled) {
                [System.Drawing.ColorTranslator]::FromHtml('#F7F8FC')
            }
            else { $t.Fg }
            $d    = $t.R * 2
            $rc   = New-Object System.Drawing.Rectangle(0, 0, ($s.Width - 1), ($s.Height - 1))
            $path = New-Object System.Drawing.Drawing2D.GraphicsPath
            $path.AddArc($rc.X,          $rc.Y,           $d, $d, 180, 90)
            $path.AddArc($rc.Right - $d, $rc.Y,           $d, $d, 270, 90)
            $path.AddArc($rc.Right - $d, $rc.Bottom - $d, $d, $d,   0, 90)
            $path.AddArc($rc.X,          $rc.Bottom - $d, $d, $d,  90, 90)
            $path.CloseAllFigures()
            $br = New-Object System.Drawing.SolidBrush($fillClr)
            $g.FillPath($br, $path); $br.Dispose()
            if ($t.Brd -and $t.BrdClr -and $s.Enabled) {
                $pen = New-Object System.Drawing.Pen($t.BrdClr, 1)
                $g.DrawPath($pen, $path); $pen.Dispose()
            }
            $sf = New-Object System.Drawing.StringFormat
            $sf.Alignment     = [System.Drawing.StringAlignment]::Center
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $tb = New-Object System.Drawing.SolidBrush($fgClr)
            # Unescape WinForms accelerator encoding (&&→&) for DrawString
            $displayText = $s.Text -replace '&&', '&'
            $g.DrawString($displayText, $s.Font, $tb,
                [System.Drawing.RectangleF]::new(0, 0, $s.Width, $s.Height), $sf)
            $tb.Dispose(); $sf.Dispose(); $path.Dispose()
        })
        $Btn.Add_EnabledChanged({ param($s, $e); $s.Invalidate() })
    }

    # ── Design tokens ─────────────────────────────────────────────────────────
    $clrGreen    = [System.Drawing.ColorTranslator]::FromHtml('#37ae1c')
    $clrSoft     = [System.Drawing.ColorTranslator]::FromHtml('#EEF0FE')
    $clrBg       = [System.Drawing.ColorTranslator]::FromHtml('#F3F4FB')
    $clrBorder   = [System.Drawing.ColorTranslator]::FromHtml('#E4E6F5')
    $clrTextPri  = [System.Drawing.ColorTranslator]::FromHtml('#1C1F4A')
    $clrTextMid  = [System.Drawing.ColorTranslator]::FromHtml('#5A5F8A')
    $clrTextSoft = [System.Drawing.ColorTranslator]::FromHtml('#9EA3C8')
    $clrCoral    = [System.Drawing.ColorTranslator]::FromHtml('#F04E65')

    # ── Form ──────────────────────────────────────────────────────────────────
    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'SharePoint Permissions Matrix Export'
    $form.StartPosition = 'CenterScreen'
    $form.Size          = New-Object System.Drawing.Size(980, 800)
    $form.MinimumSize   = New-Object System.Drawing.Size(860, 700)
    $form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.BackColor     = $clrBg

    # ── Gradient header (54 px, painted) ──────────────────────────────────────
    $pnlHeader        = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock   = 'Top'
    $pnlHeader.Height = 54
    $pnlHeader.Add_Paint({
        param($s, $e)
        $r    = $s.ClientRectangle
        $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $r,
            [System.Drawing.ColorTranslator]::FromHtml('#37ae1c'),
            [System.Drawing.ColorTranslator]::FromHtml('#3D4FD6'),
            [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal
        )
        $e.Graphics.FillRectangle($grad, $r)
        $grad.Dispose()
        $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
        $fTitle = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
        $fSub   = New-Object System.Drawing.Font('Segoe UI', 8)
        $bWhite = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $bSub   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(191, 255, 255, 255))
        $e.Graphics.DrawString('Permissions Matrix Export', $fTitle, $bWhite, 16.0, 8.0)
        $e.Graphics.DrawString('SharePoint Online', $fSub, $bSub, 17.0, 32.0)
        $fTitle.Dispose(); $fSub.Dispose(); $bWhite.Dispose(); $bSub.Dispose()
    })

    # ── Footer (56 px, white, thin top border) ────────────────────────────────
    $pnlFooter          = New-Object System.Windows.Forms.Panel
    $pnlFooter.Dock     = 'Bottom'
    $pnlFooter.Height   = 72
    $pnlFooter.BackColor = [System.Drawing.Color]::White

    $pnlFooterBorder           = New-Object System.Windows.Forms.Panel
    $pnlFooterBorder.Dock      = 'Top'
    $pnlFooterBorder.Height    = 1
    $pnlFooterBorder.BackColor = $clrBorder

    $lblEst          = New-Object System.Windows.Forms.Label
    $lblEst.Location = New-Object System.Drawing.Point(12, 6)
    $lblEst.Size     = New-Object System.Drawing.Size(680, 58)
    $lblEst.Anchor   = 'Left, Top'
    $lblEst.ForeColor = $clrTextMid
    $lblEst.Font     = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblEst.Text     = 'Connect and load sites to see scope and a time estimate.'

    $btnCancel                          = New-Object System.Windows.Forms.Button
    $btnCancel.Text                     = 'Cancel'
    $btnCancel.Size                     = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Location                 = New-Object System.Drawing.Point(754, 20)
    $btnCancel.Anchor                   = 'Top, Right'
    $btnCancel.FlatStyle                = 'Flat'
    $btnCancel.BackColor                = $clrBg
    $btnCancel.ForeColor                = $clrTextMid
    $btnCancel.FlatAppearance.BorderColor = $clrBorder
    $btnCancel.FlatAppearance.BorderSize  = 1
    $btnCancel.DialogResult             = [System.Windows.Forms.DialogResult]::Cancel

    $btnRun                            = New-Object System.Windows.Forms.Button
    $btnRun.Text                       = ("Run export {0}" -f [char]0x2192)
    $btnRun.Size                       = New-Object System.Drawing.Size(108, 32)
    $btnRun.Location                   = New-Object System.Drawing.Point(852, 20)
    $btnRun.Anchor                     = 'Top, Right'
    $btnRun.FlatStyle                  = 'Flat'
    $btnRun.BackColor                  = $clrGreen
    $btnRun.ForeColor                  = [System.Drawing.Color]::White
    $btnRun.FlatAppearance.BorderSize  = 0
    $btnRun.Enabled                    = $false

    $pnlFooter.Controls.AddRange(@($pnlFooterBorder, $lblEst, $btnCancel, $btnRun))

    # ── Scrollable body ───────────────────────────────────────────────────────
    $pnlBody            = New-Object System.Windows.Forms.Panel
    $pnlBody.Dock       = 'Fill'
    $pnlBody.AutoScroll = $false
    $pnlBody.BackColor  = $clrBg
    $pnlBody.Padding    = New-Object System.Windows.Forms.Padding(10, 0, 10, 12)

    # ── Card helper: thin 1-px bottom rule via Paint ──────────────────────────
    $cardBorderPaint = {
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml('#E4E6F5'), 1)
        $w   = $s.ClientSize.Width
        $h   = $s.ClientSize.Height
        $e.Graphics.DrawLine($pen, 0, $h - 1, $w, $h - 1)
        $pen.Dispose()
    }

    # ── CONNECTION card ───────────────────────────────────────────────────────
    # Field labels are drawn via the card's Paint event in the clear zones above
    # each TextBox. Win32 native TextBox HWNDs always paint over managed controls
    # so labels as child controls are never reliable; Paint is the only guaranteed way.
    $cardConn           = New-Object System.Windows.Forms.Panel
    $cardConn.BackColor = [System.Drawing.Color]::White
    $cardConn.Dock      = 'Top'
    $cardConn.Height    = 236

    $cardConn.Add_Paint($cardBorderPaint)

    # Use native RichTextBox controls as labels — they render at the same Win32 z-level
    # as TextBox controls, guaranteeing visibility regardless of WM_PAINT dispatch order.
    function New-NativeLabel {
        param([string]$Text, [System.Drawing.Font]$Font, [System.Drawing.Color]$ForeColor,
              [int]$X, [int]$Y, [int]$W = 220, [int]$H = 18)
        $rtb = New-Object System.Windows.Forms.RichTextBox
        $rtb.Text        = $Text
        $rtb.Font        = $Font
        $rtb.ForeColor   = $ForeColor
        $rtb.BackColor   = [System.Drawing.Color]::White
        $rtb.BorderStyle = 'None'
        $rtb.ReadOnly    = $true
        $rtb.ScrollBars  = 'None'
        $rtb.TabStop     = $false
        $rtb.Location    = New-Object System.Drawing.Point($X, $Y)
        $rtb.Size        = New-Object System.Drawing.Size($W, $H)
        $rtb.Anchor      = 'Top, Left'
        return $rtb
    }
    $fConnSec = New-Object System.Drawing.Font('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)
    $fConnFld = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    # Layout with generous gaps so no label/textbox overlap:
    #   y=10  CONNECTION (16px tall → ends y=26)
    #   y=34  Site URL label (18px → ends y=52)
    #   y=60  URL textbox (26px → ends y=86)   ← 8px clear gap above
    #   y=100 Client ID label (18px → ends y=118)
    #   y=126 Client ID textbox (26px → ends y=152)
    #   y=146 load row: progress (left) + Connect button (right), 12px gap
    #   y=188 status label
    $lblConnSec   = New-NativeLabel 'CONNECTION' $fConnSec $clrTextSoft  14  12  180 16
    $lblUrlHdr    = New-NativeLabel 'Site or admin URL' $fConnFld $clrTextMid 14 30 200 18
    $lblClientHdr = New-NativeLabel 'Client ID'  $fConnFld $clrTextMid   14  92  160 18

    # Textboxes use Dock=Top inside thin wrapper panels so width is always correct.
    # This avoids the "negative right margin" problem that occurs when Anchor=Right is
    # set before the parent panel has a real width (form not yet shown).
    function New-FieldRow {
        param($TopY, $Height = 26)
        $p = New-Object System.Windows.Forms.Panel
        $p.Location  = New-Object System.Drawing.Point(14, $TopY)
        $p.Size      = New-Object System.Drawing.Size(400, $Height)
        $p.Anchor    = 'Top, Left, Right'
        $p.BackColor = [System.Drawing.Color]::White
        return $p
    }

    $rowUrl    = New-FieldRow 50
    $txtUrl    = New-Object System.Windows.Forms.TextBox
    $txtUrl.Dock        = 'Fill'
    $txtUrl.BorderStyle = 'FixedSingle'
    $txtUrl.BackColor   = [System.Drawing.Color]::White
    $txtUrl.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $txtUrl.Text        = $InitialSiteUrl
    $rowUrl.Controls.Add($txtUrl)

    $rowClient = New-FieldRow 112
    $txtClient = New-Object System.Windows.Forms.TextBox
    $txtClient.Dock        = 'Fill'
    $txtClient.BorderStyle = 'FixedSingle'
    $txtClient.BackColor   = [System.Drawing.Color]::White
    $txtClient.Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    $txtClient.Text        = $InitialClientId
    $rowClient.Controls.Add($txtClient)

    $pnlLoadRow           = New-Object System.Windows.Forms.Panel
    $pnlLoadRow.Location  = New-Object System.Drawing.Point(14, 146)
    $pnlLoadRow.Size      = New-Object System.Drawing.Size(400, 36)
    $pnlLoadRow.Anchor    = 'Top, Left, Right'
    $pnlLoadRow.BackColor = [System.Drawing.Color]::White

    $btnLoad                             = New-Object System.Windows.Forms.Button
    $btnLoad.Text                        = 'Connect && Load Sites'
    $btnLoad.Size                        = New-Object System.Drawing.Size(190, 32)
    $btnLoad.Location                    = New-Object System.Drawing.Point(210, 0)
    $btnLoad.Anchor                      = 'Top, Right'
    $btnLoad.FlatStyle                   = 'Flat'
    $btnLoad.BackColor                   = $clrGreen
    $btnLoad.ForeColor                   = [System.Drawing.Color]::White
    $btnLoad.Font                        = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $btnLoad.FlatAppearance.BorderSize   = 0
    $btnLoad.Cursor                      = [System.Windows.Forms.Cursors]::Hand

    $loadProgress                        = New-Object System.Windows.Forms.ProgressBar
    $loadProgress.Location               = New-Object System.Drawing.Point(0, 11)
    $loadProgress.Size                   = New-Object System.Drawing.Size(200, 14)
    $loadProgress.Anchor                 = 'Top, Left, Right'
    $loadProgress.Style                  = 'Continuous'
    $loadProgress.Minimum               = 0
    $loadProgress.Maximum               = 100
    $loadProgress.Value                 = 0
    $loadProgress.Visible                = $false

    $pnlLoadRow.Controls.AddRange(@($loadProgress, $btnLoad))

    $lblLoadStatus          = New-Object System.Windows.Forms.Label
    $lblLoadStatus.Location = New-Object System.Drawing.Point(14, 188)
    $lblLoadStatus.Size     = New-Object System.Drawing.Size(400, 28)
    $lblLoadStatus.Anchor   = 'Top, Left, Right'
    $lblLoadStatus.ForeColor = $clrTextMid
    $lblLoadStatus.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
    $lblLoadStatus.Text     = ''
    $lblLoadStatus.BackColor = [System.Drawing.Color]::White

    $cardConn.Controls.AddRange(@(
        $lblConnSec, $lblUrlHdr, $rowUrl,
        $lblClientHdr, $rowClient,
        $pnlLoadRow, $lblLoadStatus
    ))

    # ── EXPORT OPTIONS card ───────────────────────────────────────────────────
    $cardOpts           = New-Object System.Windows.Forms.Panel
    $cardOpts.BackColor = [System.Drawing.Color]::White
    $cardOpts.Dock      = 'Top'
    $cardOpts.Height    = 186
    $cardOpts.Add_Paint($cardBorderPaint)

    $lblOptsSec          = New-Object System.Windows.Forms.Label
    $lblOptsSec.Text     = 'EXPORT OPTIONS'
    $lblOptsSec.Location = New-Object System.Drawing.Point(14, 10)
    $lblOptsSec.AutoSize = $true
    $lblOptsSec.ForeColor = $clrTextSoft
    $lblOptsSec.Font     = New-Object System.Drawing.Font('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)

    # Row 1 ── checkboxes + scan-mode combo
    $chkExpand          = New-Object System.Windows.Forms.CheckBox
    $chkExpand.Text     = 'Expand groups in matrix'
    $chkExpand.Location = New-Object System.Drawing.Point(14, 34)
    $chkExpand.AutoSize = $true
    $chkExpand.Checked  = $DefaultExpandGroups
    $chkExpand.ForeColor = $clrTextPri

    $chkSubsites          = New-Object System.Windows.Forms.CheckBox
    $chkSubsites.Text     = 'Include subsites'
    $chkSubsites.Location = New-Object System.Drawing.Point(210, 34)
    $chkSubsites.AutoSize = $true
    $chkSubsites.Checked  = $DefaultIncludeSubsites
    $chkSubsites.ForeColor = $clrTextPri

    $lblScanMode          = New-Object System.Windows.Forms.Label
    $lblScanMode.Text     = 'Item scan mode'
    $lblScanMode.Location = New-Object System.Drawing.Point(370, 37)
    $lblScanMode.AutoSize = $true
    $lblScanMode.ForeColor = $clrTextMid

    $cmbScanMode               = New-Object System.Windows.Forms.ComboBox
    $cmbScanMode.Location      = New-Object System.Drawing.Point(468, 33)
    $cmbScanMode.Size          = New-Object System.Drawing.Size(120, 23)
    $cmbScanMode.DropDownStyle = 'DropDownList'
    [void]$cmbScanMode.Items.AddRange(@('Auto', 'Rest', 'Csom'))
    $scanIdx = $cmbScanMode.Items.IndexOf($DefaultItemScanMode)
    $cmbScanMode.SelectedIndex = if ($scanIdx -ge 0) { $scanIdx } else { 0 }

    # Row 2 ── folder links + open workbook + all-inherited
    $chkFolderLinks          = New-Object System.Windows.Forms.CheckBox
    $chkFolderLinks.Text     = 'Include folder sharing links'
    $chkFolderLinks.Location = New-Object System.Drawing.Point(14, 66)
    $chkFolderLinks.AutoSize = $true
    $chkFolderLinks.Checked  = $DefaultIncludeFolderSharingLinks
    $chkFolderLinks.ForeColor = $clrTextPri

    $chkOpenWorkbook          = New-Object System.Windows.Forms.CheckBox
    $chkOpenWorkbook.Text     = 'Open workbook after export'
    $chkOpenWorkbook.Location = New-Object System.Drawing.Point(210, 66)
    $chkOpenWorkbook.AutoSize = $true
    $chkOpenWorkbook.Checked  = $DefaultOpenWorkbookOnComplete
    $chkOpenWorkbook.ForeColor = $clrTextPri

    $chkAllInherited          = New-Object System.Windows.Forms.CheckBox
    $chkAllInherited.Text     = 'Include ALL inherited rows  *** SLOW ***'
    $chkAllInherited.Location = New-Object System.Drawing.Point(410, 66)
    $chkAllInherited.AutoSize = $true
    $chkAllInherited.Checked  = $DefaultIncludeAllInheritedItemsInMatrix
    $chkAllInherited.ForeColor = if ($DefaultIncludeAllInheritedItemsInMatrix) { $clrCoral } else { $clrTextPri }
    $chkAllInherited.Add_CheckedChanged({
        $chkAllInherited.ForeColor = if ($chkAllInherited.Checked) {
            [System.Drawing.ColorTranslator]::FromHtml('#F04E65')
        } else {
            [System.Drawing.ColorTranslator]::FromHtml('#1C1F4A')
        }
    })

    # Row 3 ── numeric spinners
    $lblMaxItems          = New-Object System.Windows.Forms.Label
    $lblMaxItems.Text     = 'Max items'
    $lblMaxItems.Location = New-Object System.Drawing.Point(14, 98)
    $lblMaxItems.AutoSize = $true
    $lblMaxItems.ForeColor = $clrTextMid

    $numMaxItems           = New-Object System.Windows.Forms.NumericUpDown
    $numMaxItems.Location  = New-Object System.Drawing.Point(14, 118)
    $numMaxItems.Size      = New-Object System.Drawing.Size(120, 23)
    $numMaxItems.Minimum   = 500
    $numMaxItems.Maximum   = 500000
    $numMaxItems.Increment = 500
    $numMaxItems.Value     = [Math]::Max($numMaxItems.Minimum, [Math]::Min($numMaxItems.Maximum, $DefaultMaxListItemsForFullMatrix))

    $lblPageSize          = New-Object System.Windows.Forms.Label
    $lblPageSize.Text     = 'Page size'
    $lblPageSize.Location = New-Object System.Drawing.Point(150, 98)
    $lblPageSize.AutoSize = $true
    $lblPageSize.ForeColor = $clrTextMid

    $numPageSize           = New-Object System.Windows.Forms.NumericUpDown
    $numPageSize.Location  = New-Object System.Drawing.Point(150, 118)
    $numPageSize.Size      = New-Object System.Drawing.Size(120, 23)
    $numPageSize.Minimum   = 100
    $numPageSize.Maximum   = 5000
    $numPageSize.Increment = 100
    $numPageSize.Value     = [Math]::Max($numPageSize.Minimum, [Math]::Min($numPageSize.Maximum, $DefaultListItemPageSize))

    $lblWeight          = New-Object System.Windows.Forms.Label
    $lblWeight.Text     = 'Unique perm weight'
    $lblWeight.Location = New-Object System.Drawing.Point(290, 98)
    $lblWeight.AutoSize = $true
    $lblWeight.ForeColor = $clrTextMid

    $numWeight           = New-Object System.Windows.Forms.NumericUpDown
    $numWeight.Location  = New-Object System.Drawing.Point(290, 118)
    $numWeight.Size      = New-Object System.Drawing.Size(100, 23)
    $numWeight.Minimum   = 2
    $numWeight.Maximum   = 30
    $numWeight.Value     = [Math]::Max($numWeight.Minimum, [Math]::Min($numWeight.Maximum, $DefaultUniquePermProgressWeight))

    $cardOpts.Height = 160
    $cardOpts.Controls.AddRange(@(
        $lblOptsSec,
        $chkExpand, $chkSubsites, $lblScanMode, $cmbScanMode,
        $chkFolderLinks, $chkOpenWorkbook, $chkAllInherited,
        $lblMaxItems, $numMaxItems, $lblPageSize, $numPageSize, $lblWeight, $numWeight
    ))

    # Hover help for export options
    $tipOpts = New-Object System.Windows.Forms.ToolTip
    $tipOpts.AutoPopDelay = 25000
    $tipOpts.InitialDelay = 450
    $tipOpts.ReshowDelay = 200
    $tipOpts.ShowAlways = $true
    $tipOpts.SetToolTip($chkExpand, "When checked, SharePoint groups are expanded into individual member rows in the Permissions Matrix (slower).`nThe Group Members sheet always lists membership even when this is off.")
    $tipOpts.SetToolTip($chkSubsites, "When checked, recursively includes subsites under the connected site when loading the scan inventory.`nReload sites after changing this.")
    $tipOpts.SetToolTip($lblScanMode, "How list items are paged when checking for unique permissions.")
    $tipOpts.SetToolTip($cmbScanMode, "Rest: fastest when supported.`nAuto: adaptive — prefers REST on larger lists.`nCsom: use CSOM paging for all lists (fallback if REST unique-perm field is missing).")
    $tipOpts.SetToolTip($chkFolderLinks, "When checked, sharing-link APIs are called for folders with unique permissions.`nFiles are always checked for sharing links.")
    $tipOpts.SetToolTip($chkOpenWorkbook, "When checked, opens the finished Excel workbook when the export completes.")
    $tipOpts.SetToolTip($chkAllInherited, "When checked, emits matrix rows for inherited items too (not only unique-permission boundaries).`nMuch slower and produces far more rows. Prefer leaving this off unless you need every inherited path.")
    $tipOpts.SetToolTip($lblMaxItems, "For lists larger than this, still scan every item for unique permissions, but only emit matrix rows for unique-permission items.`nInherited items are covered by list/site rows.")
    $tipOpts.SetToolTip($numMaxItems, "For lists larger than this, still scan every item for unique permissions, but only emit matrix rows for unique-permission items.`nInherited items are covered by list/site rows.")
    $tipOpts.SetToolTip($lblPageSize, "Items requested per REST/CSOM page when scanning libraries.`nSmaller values show progress sooner; larger values (up to 5000) reduce round-trips.")
    $tipOpts.SetToolTip($numPageSize, "Items requested per REST/CSOM page when scanning libraries.`nSmaller values show progress sooner; larger values (up to 5000) reduce round-trips.")
    $tipOpts.SetToolTip($lblWeight, "How much a unique-permission item counts toward overall progress vs a normal item.`nHigher values keep the progress bar moving during slow permission fetches.")
    $tipOpts.SetToolTip($numWeight, "How much a unique-permission item counts toward overall progress vs a normal item.`nHigher values keep the progress bar moving during slow permission fetches.")
    $tipOpts.SetToolTip($txtUrl, "SharePoint site URL, or the tenant admin URL (https://contoso-admin.sharepoint.com) to enumerate site collections.")
    $tipOpts.SetToolTip($txtClient, "Entra app (client) ID used for interactive PnP sign-in.")
    $tipOpts.SetToolTip($btnLoad, "Sign in and load the site/list inventory for the URL above.")
    $tipOpts.SetToolTip($btnRun, "Start the permissions matrix export for the selected lists and libraries.")
    $tipOpts.SetToolTip($btnCancel, "Close the picker without running an export.")

    # ── SITES TO SCAN card ────────────────────────────────────────────────────
    # ListView is hosted in a Dock=Fill panel so it always uses remaining card height.
    # A fixed pixel height + Top/Bottom anchor often clips the last row on high-DPI displays.
    $cardSites           = New-Object System.Windows.Forms.Panel
    $cardSites.BackColor = [System.Drawing.Color]::White
    $cardSites.Dock      = 'Fill'

    $pnlSitesHdr           = New-Object System.Windows.Forms.Panel
    $pnlSitesHdr.Dock      = 'Top'
    $pnlSitesHdr.Height    = 34
    $pnlSitesHdr.BackColor = [System.Drawing.Color]::White

    $lblSitesSec          = New-Object System.Windows.Forms.Label
    $lblSitesSec.Text     = 'SITES TO SCAN'
    $lblSitesSec.Location = New-Object System.Drawing.Point(14, 10)
    $lblSitesSec.AutoSize = $true
    $lblSitesSec.ForeColor = $clrTextSoft
    $lblSitesSec.Font     = New-Object System.Drawing.Font('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)

    $btnAll                            = New-Object System.Windows.Forms.Button
    $btnAll.Text                       = 'All'
    $btnAll.Size                       = New-Object System.Drawing.Size(52, 22)
    $btnAll.Location                   = New-Object System.Drawing.Point(820, 6)
    $btnAll.Anchor                     = 'Top, Right'
    $btnAll.FlatStyle                  = 'Flat'
    $btnAll.BackColor                  = $clrSoft
    $btnAll.ForeColor                  = $clrGreen
    $btnAll.FlatAppearance.BorderSize  = 0

    $btnNone                           = New-Object System.Windows.Forms.Button
    $btnNone.Text                      = 'None'
    $btnNone.Size                      = New-Object System.Drawing.Size(52, 22)
    $btnNone.Location                  = New-Object System.Drawing.Point(878, 6)
    $btnNone.Anchor                    = 'Top, Right'
    $btnNone.FlatStyle                 = 'Flat'
    $btnNone.BackColor                 = $clrSoft
    $btnNone.ForeColor                 = $clrGreen
    $btnNone.FlatAppearance.BorderSize = 0

    $btnExpandAll                           = New-Object System.Windows.Forms.Button
    $btnExpandAll.Text                      = 'Expand'
    $btnExpandAll.Size                      = New-Object System.Drawing.Size(58, 22)
    $btnExpandAll.Location                  = New-Object System.Drawing.Point(696, 6)
    $btnExpandAll.Anchor                    = 'Top, Right'
    $btnExpandAll.FlatStyle                 = 'Flat'
    $btnExpandAll.BackColor                 = $clrSoft
    $btnExpandAll.ForeColor                 = $clrGreen
    $btnExpandAll.FlatAppearance.BorderSize = 0

    $btnCollapseAll                           = New-Object System.Windows.Forms.Button
    $btnCollapseAll.Text                      = 'Collapse'
    $btnCollapseAll.Size                      = New-Object System.Drawing.Size(66, 22)
    $btnCollapseAll.Location                  = New-Object System.Drawing.Point(758, 6)
    $btnCollapseAll.Anchor                    = 'Top, Right'
    $btnCollapseAll.FlatStyle                 = 'Flat'
    $btnCollapseAll.BackColor                 = $clrSoft
    $btnCollapseAll.ForeColor                 = $clrGreen
    $btnCollapseAll.FlatAppearance.BorderSize = 0

    $pnlSitesHdr.Controls.AddRange(@($lblSitesSec, $btnExpandAll, $btnCollapseAll, $btnAll, $btnNone))
    $tipOpts.SetToolTip($btnExpandAll, "Expand all sites to show their lists and libraries.")
    $tipOpts.SetToolTip($btnCollapseAll, "Collapse all sites to hide list/library rows.")
    $tipOpts.SetToolTip($btnAll, "Select every site and all of its lists/libraries.")
    $tipOpts.SetToolTip($btnNone, "Deselect every site and all of its lists/libraries.")

    $pnlSitesList           = New-Object System.Windows.Forms.Panel
    $pnlSitesList.Dock      = 'Fill'
    $pnlSitesList.BackColor = [System.Drawing.Color]::White
    $pnlSitesList.Padding   = New-Object System.Windows.Forms.Padding(14, 0, 14, 10)

    $lv               = New-Object System.Windows.Forms.ListView
    $lv.Dock          = 'Fill'
    $lv.View          = 'Details'
    $lv.CheckBoxes    = $true
    $lv.FullRowSelect = $true
    $lv.GridLines     = $false
    $lv.BackColor     = [System.Drawing.Color]::White
    $lv.BorderStyle   = 'FixedSingle'
    [void]$lv.Columns.Add('Site / list', 240)
    [void]$lv.Columns.Add('Path', 280)
    [void]$lv.Columns.Add('Type', 70)
    [void]$lv.Columns.Add('Lists', 70, 'Right')
    [void]$lv.Columns.Add('Items', 90, 'Right')
    [void]$lv.Columns.Add('Unique', 70, 'Center')
    $tipOpts.SetToolTip($lv, "Check sites or individual lists/libraries to include in the export.`nClick ▶ to expand a site. Site checkbox selects/deselects all lists under that site.")

    $pnlSitesList.Controls.Add($lv)

    function Sync-SitesListViewHeight {
        if ($pnlSitesList.ClientSize.Height -lt 40) { return }
        $padL = $pnlSitesList.Padding.Left
        $padR = $pnlSitesList.Padding.Right
        $w    = [Math]::Max(100, $pnlSitesList.ClientSize.Width - $padL - $padR)
        $avail = $pnlSitesList.ClientSize.Height - $pnlSitesList.Padding.Bottom
        $rowH  = 22
        if ($lv.Items.Count -gt 0) {
            try {
                $rh = $lv.GetItemRect(0).Height
                if ($rh -gt 8) { $rowH = $rh }
            }
            catch { }
        }
        $rows = [Math]::Max(3, [Math]::Floor($avail / $rowH))
        $lv.Dock = 'None'
        $lv.Location = New-Object System.Drawing.Point($padL, 0)
        $lv.Size     = New-Object System.Drawing.Size($w, ($rows * $rowH))
    }

    $pnlSitesList.Add_Resize({ Sync-SitesListViewHeight })

    # Fill first, then Top header — standard WinForms dock stacking order.
    $cardSites.Controls.AddRange(@($pnlSitesList, $pnlSitesHdr))

    # WinForms Dock stacking: Fill must be added first (it renders behind), then Top items in REVERSE order.
    $spacer1 = New-Object System.Windows.Forms.Panel; $spacer1.Dock='Top'; $spacer1.Height=6; $spacer1.BackColor=$clrBg
    $spacer2 = New-Object System.Windows.Forms.Panel; $spacer2.Dock='Top'; $spacer2.Height=6; $spacer2.BackColor=$clrBg
    $spacer3 = New-Object System.Windows.Forms.Panel; $spacer3.Dock='Top'; $spacer3.Height=8; $spacer3.BackColor=$clrBg
    $pnlBody.Controls.AddRange(@($cardSites, $spacer2, $cardOpts, $spacer1, $cardConn, $spacer3))
    # pnlBody added first (Controls[0]) so layout engine positions it correctly
    # relative to the header/footer; header/footer added after so they're visually
    # on top (lower Controls index = higher z-order in WinForms).
    $form.Controls.Add($pnlBody)
    $form.Controls.Add($pnlFooter)
    $form.Controls.Add($pnlHeader)

    $form.CancelButton = $btnCancel
    $form.AcceptButton = $btnRun

    function Update-PickerLayoutMargins {
        $pad = 14

        # Connection card: field rows, load row (progress + button), status
        $cw = $cardConn.ClientSize.Width
        if ($cw -gt 0) {
            foreach ($r in @($rowUrl, $rowClient)) { $r.Width = $cw - $pad - $pad }
            $pnlLoadRow.Width = $cw - $pad - $pad
            $loadRowGap = 12
            $btnLoad.Left = $pnlLoadRow.ClientSize.Width - $btnLoad.Width
            $loadProgress.Left = 0
            $loadProgress.Width = [Math]::Max(80, $btnLoad.Left - $loadRowGap)
            $lblLoadStatus.Width = $cw - $pad - $pad
        }

        # Sites card: Expand/Collapse/All/None buttons (ListView width follows Dock=Fill host panel)
        $hw = $pnlSitesHdr.ClientSize.Width
        if ($hw -gt 0) {
            $btnNone.Left = $hw - $pad - $btnNone.Width
            $btnAll.Left  = $btnNone.Left - 4 - $btnAll.Width
            $btnCollapseAll.Left = $btnAll.Left - 4 - $btnCollapseAll.Width
            $btnExpandAll.Left   = $btnCollapseAll.Left - 4 - $btnExpandAll.Width
        }

        # Footer: Cancel and Run export buttons
        $fw = $pnlFooter.ClientSize.Width
        if ($fw -gt 0) {
            $btnRun.Left    = $fw - $pad - $btnRun.Width
            $btnCancel.Left = $fw - $pad - $btnRun.Width - 8 - $btnCancel.Width
        }
    }

    # After layout runs, fix right margins for anchored controls whose parent
    # width was 0 at creation time (before the form was shown).
    $form.Add_Shown({
        Update-PickerLayoutMargins
        Sync-SitesListViewHeight
    })
    $form.Add_Resize({
        Update-PickerLayoutMargins
        Sync-SitesListViewHeight
    })

    function Set-PickerLoadProgress {
        param(
            [int]    $Current = 0,
            [int]    $Total = 0,
            [string] $Message = '',
            [ValidateSet('Marquee', 'Determinate', 'Hidden')]
            [string] $Mode = 'Determinate'
        )

        if ($Mode -eq 'Hidden') {
            $loadProgress.Visible = $false
            return
        }

        $loadProgress.Visible = $true
        if ($Mode -eq 'Marquee') {
            $loadProgress.Style = 'Marquee'
            $loadProgress.MarqueeAnimationSpeed = 45
            if ($Message) { $lblLoadStatus.Text = $Message }
        }
        else {
            $loadProgress.Style = 'Continuous'
            $max = [Math]::Max(1, $Total)
            $val = [Math]::Max(0, [Math]::Min($Current, $max))
            $loadProgress.Maximum = $max
            $loadProgress.Value = $val
            $pct = [int][Math]::Round(100.0 * $val / $max)
            if ($Total -gt 0 -and $Current -gt 0) {
                $lblLoadStatus.Text = ('{0} — {1} / {2} ({3}%)' -f $Message, $Current, $Total, $pct)
            }
            elseif ($Message) {
                $lblLoadStatus.Text = $Message
            }
            else {
                $lblLoadStatus.Text = ('Loading... {0}%' -f $pct)
            }
        }

        [System.Windows.Forms.Application]::DoEvents()
    }

    $pickerLoadProgress = {
        param([int] $Current, [int] $Total, [string] $Message)
        Set-PickerLoadProgress -Current $Current -Total $Total -Message $Message -Mode Determinate
    }

    function Update-SiteRowTitle {
        param($SiteRow)

        $glyph = if (@($SiteRow.ListRows).Count -lt 1) {
            '   '
        }
        elseif ($SiteRow.Expanded) {
            "$([char]0x25BC) "  # ▼
        }
        else {
            "$([char]0x25B6) "  # ▶
        }
        $SiteRow.ListItem.Text = $glyph + $SiteRow.Title
    }

    function Test-ListRowChecked {
        param($ListRow)
        if (-not $ListRow) { return $false }
        return [bool]$ListRow.IsChecked
    }

    function Set-ListRowChecked {
        param(
            $ListRow,
            [bool] $Checked
        )
        if (-not $ListRow) { return }
        $ListRow.IsChecked = $Checked
        if ($ListRow.ListItem -and $ListRow.ListItem.Checked -ne $Checked) {
            $ListRow.ListItem.Checked = $Checked
        }
    }

    function Set-SiteRowExpanded {
        param(
            $SiteRow,
            [bool] $Expanded
        )

        if (-not $SiteRow -or $SiteRow.Kind -ne 'Site') { return }
        if (@($SiteRow.ListRows).Count -lt 1) {
            $SiteRow.Expanded = $false
            Update-SiteRowTitle -SiteRow $SiteRow
            return
        }

        if ($SiteRow.Expanded -eq $Expanded) {
            Update-SiteRowTitle -SiteRow $SiteRow
            return
        }

        $SiteRow.Expanded = $Expanded
        Update-SiteRowTitle -SiteRow $SiteRow

        $script:PickerSuppressCheck = $true
        try {
            if ($Expanded) {
                $idx = $SiteRow.ListItem.Index + 1
                $lv.BeginUpdate()
                try {
                    foreach ($lr in @($SiteRow.ListRows)) {
                        if ($lr.ListItem -and -not $lv.Items.Contains($lr.ListItem)) {
                            # Inserting a detached ListViewItem resets Checked — restore from IsChecked after insert.
                            $lv.Items.Insert($idx, $lr.ListItem) | Out-Null
                            $idx++
                        }
                    }
                }
                finally {
                    $lv.EndUpdate()
                }
                foreach ($lr in @($SiteRow.ListRows)) {
                    if ($lr.ListItem) {
                        $lr.ListItem.Checked = [bool]$lr.IsChecked
                    }
                }
                Update-SiteRowSummary -SiteRow $SiteRow
            }
            else {
                foreach ($lr in @($SiteRow.ListRows)) {
                    if ($lr.ListItem -and $lv.Items.Contains($lr.ListItem)) {
                        # Capture current UI state before detach.
                        $lr.IsChecked = [bool]$lr.ListItem.Checked
                        $lv.Items.Remove($lr.ListItem)
                    }
                }
            }
        }
        finally {
            $script:PickerSuppressCheck = $false
        }

        Sync-SitesListViewHeight
    }

    function Get-PickerFilteredPlan {
        $result = [System.Collections.Generic.List[object]]::new()
        foreach ($siteRow in $script:PickerSiteRows) {
            $checkedLists = @(
                $siteRow.ListRows |
                    Where-Object { Test-ListRowChecked $_ } |
                    ForEach-Object { $_.List }
            )
            if ($checkedLists.Count -lt 1) { continue }

            $items = 0
            foreach ($l in $checkedLists) { $items += [int]$l.ItemCount }
            [void]$result.Add([pscustomobject]@{
                Site  = $siteRow.PlanEntry.Site
                Lists = $checkedLists
                Items = $items
            })
        }
        return @($result)
    }

    function Update-SiteRowSummary {
        param($SiteRow)

        $total = @($SiteRow.ListRows).Count
        $sel = 0
        $items = 0
        foreach ($lr in @($SiteRow.ListRows)) {
            if (Test-ListRowChecked $lr) {
                $sel++
                $items += [int]$lr.List.ItemCount
            }
        }

        if ($total -lt 1) {
            $SiteRow.ListItem.SubItems[3].Text = '0'
        }
        elseif ($sel -eq $total) {
            $SiteRow.ListItem.SubItems[3].Text = [string]$total
        }
        else {
            $SiteRow.ListItem.SubItems[3].Text = ('{0}/{1}' -f $sel, $total)
        }
        $SiteRow.ListItem.SubItems[4].Text = ('{0:N0}' -f $items)
        Update-SiteRowTitle -SiteRow $SiteRow

        $wantChecked = ($sel -gt 0)
        if ($SiteRow.ListItem.Checked -ne $wantChecked) {
            $wasSuppressed = $script:PickerSuppressCheck
            $script:PickerSuppressCheck = $true
            try { $SiteRow.ListItem.Checked = $wantChecked }
            finally { $script:PickerSuppressCheck = $wasSuppressed }
        }
    }

    function Set-SiteListsChecked {
        param(
            $SiteRow,
            [bool] $Checked
        )

        $script:PickerSuppressCheck = $true
        try {
            foreach ($lr in @($SiteRow.ListRows)) {
                Set-ListRowChecked -ListRow $lr -Checked $Checked
            }
            $wantSite = $Checked -and (@($SiteRow.ListRows).Count -gt 0)
            if ($SiteRow.ListItem.Checked -ne $wantSite) {
                $SiteRow.ListItem.Checked = $wantSite
            }
            Update-SiteRowSummary -SiteRow $SiteRow
        }
        finally {
            $script:PickerSuppressCheck = $false
        }
    }

    function Update-EstimateLabel {
        if (@($script:PickerSiteRows).Count -lt 1) {
            $btnRun.Enabled = $false
            return
        }

        $planSlice = @(Get-PickerFilteredPlan)
        $btnRun.Enabled = ($planSlice.Count -gt 0)
        if ($planSlice.Count -lt 1) {
            $lblEst.Text = 'Select at least one list or library to export.'
            return
        }

        $est = Get-PermissionsMatrixTimeEstimate -PlanEntries $planSlice `
            -ExpandGroups $chkExpand.Checked `
            -IncludeFolderSharingLinks $chkFolderLinks.Checked `
            -IncludeAllInheritedItemsInMatrix $chkAllInherited.Checked `
            -ItemScanMode ([string]$cmbScanMode.SelectedItem)

        $lblEst.Text = @(
            ("ETA ~{0}–{1}  (scan ~{2}, Excel ~{3})" -f `
                (Format-SecondsEstimate $est.RangeLowSeconds),
                (Format-SecondsEstimate $est.RangeHighSeconds),
                (Format-SecondsEstimate $est.ScanSeconds),
                (Format-SecondsEstimate $est.ExportSeconds))
            ("Scope: {0} site(s) | {1:N0} items | {2:N0} lists/libraries | ~{3:N0} unique-perm items" -f `
                $est.SiteCount, $est.TotalItems, $est.ListCount, $est.EstimatedUniqueItems)
            ("Mode: {0}. All-inherited mode is slower but includes every inherited path. Click ▶ to expand lists." -f [string]$cmbScanMode.SelectedItem)
        ) -join [Environment]::NewLine
    }

    function Populate-SiteList {
        param([object[]] $Plan)

        $script:PickerSuppressCheck = $true
        try {
            $lv.Items.Clear()
            [void]$script:PickerSiteRows.Clear()
            $script:PickerPlan = @($Plan)

            $siteFont = New-Object System.Drawing.Font($lv.Font, [System.Drawing.FontStyle]::Bold)
            $listColor = $clrTextMid

            foreach ($entry in @($Plan)) {
                $site      = $entry.Site
                $path      = [string]$site.ServerRelativeUrl
                $title     = [string]$site.Title
                if ([string]::IsNullOrWhiteSpace($title)) { $title = $path }
                $type      = if ($path -match '^/sites/[^/]+/.+') { 'Subsite' } else { 'Site' }
                $lists     = @($entry.Lists)
                $listCount = $lists.Count
                $items     = [int]$entry.Items
                $unique    = if ([bool]$site.HasUniqueRoleAssignments) { 'Yes' } else { 'No' }

                $item         = New-Object System.Windows.Forms.ListViewItem $title
                $item.Checked = ($listCount -gt 0)
                $item.Font    = $siteFont
                [void]$item.SubItems.Add($path)
                [void]$item.SubItems.Add($type)
                [void]$item.SubItems.Add([string]$listCount)
                [void]$item.SubItems.Add(('{0:N0}' -f $items))
                [void]$item.SubItems.Add($unique)
                [void]$lv.Items.Add($item)

                $siteRow = [pscustomobject]@{
                    Kind      = 'Site'
                    Title     = $title
                    Expanded  = $false
                    ListItem  = $item
                    PlanEntry = $entry
                    Path      = $path
                    ListRows  = [System.Collections.Generic.List[object]]::new()
                }
                $item.Tag = $siteRow
                [void]$script:PickerSiteRows.Add($siteRow)

                foreach ($list in $lists) {
                    $listTitle = [string]$list.Title
                    $listPath  = ''
                    try { $listPath = [string]$list.RootFolder.ServerRelativeUrl } catch { }
                    $baseType  = [string]$list.BaseType
                    $listType  = if ($baseType -eq 'DocumentLibrary') { 'Library' } else { 'List' }
                    $listItems = [int]$list.ItemCount
                    $listUnique = if ([bool]$list.HasUniqueRoleAssignments) { 'Yes' } else { 'No' }

                    # Create list rows but keep them out of the ListView until the site is expanded.
                    $listItem         = New-Object System.Windows.Forms.ListViewItem ('    ' + $listTitle)
                    $listItem.Checked = $true
                    $listItem.ForeColor = $listColor
                    [void]$listItem.SubItems.Add($listPath)
                    [void]$listItem.SubItems.Add($listType)
                    [void]$listItem.SubItems.Add('')
                    [void]$listItem.SubItems.Add(('{0:N0}' -f $listItems))
                    [void]$listItem.SubItems.Add($listUnique)

                    $listRow = [pscustomobject]@{
                        Kind      = 'List'
                        ListItem  = $listItem
                        List      = $list
                        SiteRow   = $siteRow
                        IsChecked = $true
                    }
                    $listItem.Tag = $listRow
                    [void]$siteRow.ListRows.Add($listRow)
                }

                Update-SiteRowSummary -SiteRow $siteRow
            }
        }
        finally {
            $script:PickerSuppressCheck = $false
        }

        Sync-SitesListViewHeight
        Update-EstimateLabel
    }

    $updateHandler = { Update-EstimateLabel }
    $chkExpand.Add_CheckedChanged($updateHandler)
    $chkFolderLinks.Add_CheckedChanged($updateHandler)
    $chkAllInherited.Add_CheckedChanged($updateHandler)
    $cmbScanMode.Add_SelectedIndexChanged($updateHandler)

    $lv.Add_ItemCheck({
        param($sender, $e)
        # Allow programmatic updates (expand restore, All/None, site cascade).
        if ($script:PickerSuppressCheck) { return }
        # Cancel only the accidental toggle from an expand/collapse click.
        if ([datetime]::UtcNow -lt $script:PickerBlockCheckUntil) {
            $e.NewValue = $e.CurrentValue
        }
    })

    $lv.Add_ItemChecked({
        param($sender, $e)
        if ($script:PickerSuppressCheck) { return }

        $row = $e.Item.Tag
        if (-not $row) { return }

        if ($row.Kind -eq 'Site') {
            Set-SiteListsChecked -SiteRow $row -Checked ([bool]$e.Item.Checked)
        }
        else {
            $row.IsChecked = [bool]$e.Item.Checked
            Update-SiteRowSummary -SiteRow $row.SiteRow
        }

        Update-EstimateLabel
    })

    $lv.Add_MouseDown({
        param($sender, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }

        $hit = $lv.HitTest($e.Location)
        if (-not $hit.Item) { return }
        $row = $hit.Item.Tag
        if (-not $row -or $row.Kind -ne 'Site') { return }
        if (@($row.ListRows).Count -lt 1) { return }

        $onLabel = ($hit.Location -band [System.Windows.Forms.ListViewHitTestLocations]::Label) -ne 0
        if (-not $onLabel) { return }

        $labelLeft = $hit.Item.Bounds.Left
        try {
            $labelLeft = $hit.Item.GetBounds([System.Windows.Forms.ItemBoundsPortion]::Label).Left
        }
        catch { }
        $onGlyph = (($e.X - $labelLeft) -ge 0) -and (($e.X - $labelLeft) -lt 22)
        $isDouble = $e.Clicks -ge 2

        if ($onGlyph) {
            $script:PickerBlockCheckUntil = [datetime]::UtcNow.AddMilliseconds(250)
            if ($e.Clicks -gt 1) { return }
            Set-SiteRowExpanded -SiteRow $row -Expanded (-not [bool]$row.Expanded)
            return
        }

        if ($isDouble) {
            $script:PickerBlockCheckUntil = [datetime]::UtcNow.AddMilliseconds(250)
            Set-SiteRowExpanded -SiteRow $row -Expanded (-not [bool]$row.Expanded)
        }
    })

    $lv.Add_KeyDown({
        param($sender, $e)
        if ($null -eq $lv.FocusedItem) { return }
        $row = $lv.FocusedItem.Tag
        if (-not $row -or $row.Kind -ne 'Site') { return }
        if (@($row.ListRows).Count -lt 1) { return }

        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Right -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Add) {
            Set-SiteRowExpanded -SiteRow $row -Expanded $true
            $e.Handled = $true
        }
        elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Left -or $e.KeyCode -eq [System.Windows.Forms.Keys]::Subtract) {
            Set-SiteRowExpanded -SiteRow $row -Expanded $false
            $e.Handled = $true
        }
    })

    $btnExpandAll.Add_Click({
        foreach ($row in $script:PickerSiteRows) {
            Set-SiteRowExpanded -SiteRow $row -Expanded $true
        }
    })

    $btnCollapseAll.Add_Click({
        foreach ($row in $script:PickerSiteRows) {
            Set-SiteRowExpanded -SiteRow $row -Expanded $false
        }
    })

    $btnAll.Add_Click({
        foreach ($row in $script:PickerSiteRows) {
            Set-SiteListsChecked -SiteRow $row -Checked $true
        }
        Update-EstimateLabel
    })

    $btnNone.Add_Click({
        foreach ($row in $script:PickerSiteRows) {
            Set-SiteListsChecked -SiteRow $row -Checked $false
        }
        Update-EstimateLabel
    })

    $btnLoad.Add_Click({
        $url    = $txtUrl.Text.Trim()
        $client = $txtClient.Text.Trim()
        if (-not $url) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Site or tenant admin URL is required.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        if (-not $client) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Client ID is required.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $btnLoad.Enabled    = $false
        $btnLoad.Text       = 'Loading...'
        Set-PickerLoadProgress -Mode Marquee -Message 'Connecting to SharePoint...'
        $form.Cursor        = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-MatrixLog "Picker: connecting to $url ..." -Level Picker
            Connect-PnPOnline -Url $url -Interactive -ClientId $client
            $S.PnpClientId = $client
            Set-PnPConnectionScopeFromInputUrl $url

            if ($S.IsTenantAdminScope) {
                Write-MatrixLog 'Picker: tenant admin URL — loading all accessible site collections...' -Level Tenant
                Set-PickerLoadProgress -Current 0 -Total 1 -Message 'Connected. Enumerating tenant site collections...'
            }
            else {
                Write-MatrixLog 'Picker: loading site and list inventory...' -Level Picker
                Set-PickerLoadProgress -Current 0 -Total 1 -Message 'Connected. Loading site inventory...'
            }
            $plan = @(Get-PermissionsMatrixScanPlan -IncludeSubsites $chkSubsites.Checked -ProgressCallback $pickerLoadProgress)
            if ($plan.Count -lt 1) {
                $msg = if ($S.IsTenantAdminScope) {
                    @'
No site collections were accessible with your account.

Tenant admin lists all site URLs, but you only need member access on each site you want to report on. Try a single site URL you own, or grant yourself access to more site collections and reload.
'@
                }
                else {
                    'No sites or lists were found. Check the URL and your permissions on that site.'
                }
                [System.Windows.Forms.MessageBox]::Show($form, $msg.Trim(), 'No accessible sites', 'OK', 'Warning') | Out-Null
                $lblLoadStatus.Text = 'No accessible sites found for this account.'
            }
            else {
                Populate-SiteList -Plan $plan
                Write-MatrixLog ("Picker: found {0} site(s)." -f $plan.Count) -Level Success
                $itemSum = ($plan | ForEach-Object { $_.Items } | Measure-Object -Sum).Sum
                $lblLoadStatus.Text = ("Loaded {0} site(s), {1:N0} item(s)." -f $plan.Count, $itemSum)
                if ($S.IsTenantAdminScope) {
                    $skipped = $S.TenantConnectSkipped + $S.TenantInventorySkipped
                    if ($skipped -gt 0) {
                        $lblLoadStatus.Text += " ($skipped inaccessible — skipped.)"
                    }
                    $lblLoadStatus.Text += ' Uncheck sites or lists you do not need.'
                }
                else {
                    $lblLoadStatus.Text += ' Uncheck lists/libraries you do not need.'
                }
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Load failed', 'OK', 'Error') | Out-Null
            $lblLoadStatus.Text = 'Load failed. Check credentials/network and try again.'
        }
        finally {
            $form.Cursor        = [System.Windows.Forms.Cursors]::Default
            $btnLoad.Enabled    = $true
            $btnLoad.Text       = 'Connect && Load Sites'
            Set-PickerLoadProgress -Mode Hidden
        }
    })

    $script:PickerResult = $null

    $btnRun.Add_Click({
        $selectedPlan = @(Get-PickerFilteredPlan)
        if ($selectedPlan.Count -lt 1) {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                'Select at least one list or library to export.',
                'Validation',
                'OK',
                'Warning'
            ) | Out-Null
            return
        }

        $script:PickerResult = [pscustomobject]@{
            SiteUrl                          = $txtUrl.Text.Trim()
            ClientId                         = $txtClient.Text.Trim()
            ExpandGroups                     = $chkExpand.Checked
            IncludeSubsites                  = $chkSubsites.Checked
            IncludeFolderSharingLinks        = $chkFolderLinks.Checked
            IncludeAllInheritedItemsInMatrix = $chkAllInherited.Checked
            OpenWorkbookOnComplete           = $chkOpenWorkbook.Checked
            ItemScanMode                     = [string]$cmbScanMode.SelectedItem
            MaxListItemsForFullMatrix        = [int]$numMaxItems.Value
            ListItemPageSize                 = [int]$numPageSize.Value
            UniquePermProgressWeight         = [int]$numWeight.Value
            Plan                             = $selectedPlan
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    # Apply rounded corners to all picker buttons
    $clrGreenDark  = [System.Drawing.ColorTranslator]::FromHtml('#2d9e14')
    $clrBrandDark  = [System.Drawing.ColorTranslator]::FromHtml('#3D4FD6')
    Set-RoundedButton $btnLoad    -Radius 7 -Bg $clrGreen   -Fg ([System.Drawing.Color]::White)
    Set-RoundedButton $btnRun     -Radius 7 -Bg $clrGreen   -Fg ([System.Drawing.Color]::White)
    Set-RoundedButton $btnCancel  -Radius 7 -Bg $clrBg      -Fg $clrTextMid -Border $true -BorderClr $clrBorder
    Set-RoundedButton $btnAll         -Radius 5 -Bg $clrSoft    -Fg $clrGreen
    Set-RoundedButton $btnNone        -Radius 5 -Bg $clrSoft    -Fg $clrGreen
    Set-RoundedButton $btnExpandAll   -Radius 5 -Bg $clrSoft    -Fg $clrGreen
    Set-RoundedButton $btnCollapseAll -Radius 5 -Bg $clrSoft    -Fg $clrGreen

    [void]$form.ShowDialog()
    return $script:PickerResult
}

# --- Main ---
$script:ExportPlanFromPicker = $null
$script:PickerAlreadyConnected = $false

if (-not $NoPicker) {
    $picked = Show-PermissionsMatrixExportPicker `
        -InitialSiteUrl $SiteUrl `
        -InitialClientId $ClientId `
        -DefaultExpandGroups $ExpandGroups `
        -DefaultIncludeSubsites $IncludeSubsites `
        -DefaultIncludeFolderSharingLinks $IncludeFolderSharingLinks `
        -DefaultIncludeAllInheritedItemsInMatrix $IncludeAllInheritedItemsInMatrix `
        -DefaultOpenWorkbookOnComplete $OpenWorkbookOnComplete `
        -DefaultItemScanMode $ItemScanMode `
        -DefaultMaxListItemsForFullMatrix $MaxListItemsForFullMatrix `
        -DefaultListItemPageSize $ListItemPageSize `
        -DefaultUniquePermProgressWeight $UniquePermProgressWeight
    if (-not $picked) {
        Write-MatrixLog 'Export cancelled.' -Level Warn
        return
    }

    $SiteUrl = $picked.SiteUrl
    $ClientId = $picked.ClientId
    $ExpandGroups = $picked.ExpandGroups
    $IncludeSubsites = $picked.IncludeSubsites
    $IncludeFolderSharingLinks = $picked.IncludeFolderSharingLinks
    $IncludeAllInheritedItemsInMatrix = $picked.IncludeAllInheritedItemsInMatrix
    $OpenWorkbookOnComplete = $picked.OpenWorkbookOnComplete
    $ItemScanMode = $picked.ItemScanMode
    $MaxListItemsForFullMatrix = $picked.MaxListItemsForFullMatrix
    $ListItemPageSize = [Math]::Max(100, [Math]::Min(5000, $picked.ListItemPageSize))
    $UniquePermProgressWeight = $picked.UniquePermProgressWeight
    $script:ExportPlanFromPicker = @($picked.Plan)
    $script:PickerAlreadyConnected = $true
}

if ([string]::IsNullOrWhiteSpace($SiteUrl)) { $SiteUrl = Read-Host 'SharePoint site URL' }
if ([string]::IsNullOrWhiteSpace($ClientId)) { $ClientId = Read-Host 'Entra app Client ID' }
$SiteUrl = $SiteUrl.Trim()
$ClientId = $ClientId.Trim()
$S.PnpClientId = $ClientId
if (-not $SiteUrl) { throw 'SiteUrl is required.' }
if (-not $ClientId) { throw 'ClientId is required.' }

Write-MatrixBanner

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$totalSw = [System.Diagnostics.Stopwatch]::StartNew()
$script:RunStartUtc = [datetime]::UtcNow

if (-not $PickerAlreadyConnected) {
    Write-MatrixLog "Connecting to $SiteUrl ..." -Level Picker
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    Set-PnPConnectionScopeFromInputUrl $SiteUrl
    if ($S.IsTenantAdminScope) {
        Write-MatrixLog 'Tenant admin URL detected — export can include all site collections you can access.' -Level Tenant
    }
}

# Seed $S.WebPath from the connected URL so Initialize-RoleColumns fetches role
# definitions from the correct site collection (not the tenant admin endpoint).
if (-not $S.IsTenantAdminScope -and [string]::IsNullOrWhiteSpace($S.WebPath) -and $S.ConnectedWebUrl) {
    try {
        $p = ([Uri]$S.ConnectedWebUrl).AbsolutePath.TrimEnd('/')
        if ($p -and $p -ne '/') { $S.WebPath = $p }
    } catch { }
}

if (-not $S.IsTenantAdminScope) {
    Initialize-RoleColumns
}
Write-Verbose ("Role columns: {0}" -f ($S.RoleColumns -join ', '))

if ($ExportPlanFromPicker) {
    $plan = @($ExportPlanFromPicker)
    $sites = @($plan | ForEach-Object { $_.Site })
    $pickerLists = ($plan | ForEach-Object { @($_.Lists).Count } | Measure-Object -Sum).Sum
    Write-MatrixLog ("Sites to scan: {0} ({1} list(s)/libraries from picker selection)" -f $sites.Count, $pickerLists) -Level Phase
}
else {
    $plan = @(Get-PermissionsMatrixScanPlan -IncludeSubsites $IncludeSubsites)
    Write-MatrixLog ("Sites to scan: {0} (IncludeSubsites={1})" -f @($plan).Count, $IncludeSubsites) -Level Phase
}

if ($S.IsTenantAdminScope) {
    if (@($plan).Count -lt 1) {
        throw @'
No accessible sites to export. Tenant admin URL lists all site collections, but your account must have access to each site you scan.
Connect with a site URL you can open in the browser, or grant access to more site collections and try again.
'@
    }
    Connect-SiteWeb @($plan)[0].Site
    Initialize-RoleColumns
}

$listTotal = ($plan | ForEach-Object { @($_.Lists).Count } | Measure-Object -Sum).Sum
$S.ProgressTotal = ($plan | ForEach-Object { $_.Items } | Measure-Object -Sum).Sum
if (-not $S.ProgressTotal) { $S.ProgressTotal = 0 }
$S.ProgressItemsTotal = $S.ProgressTotal
$S.ProgressSiteTotal = @($plan).Count
$S.ProgressSiteIndex = 0

Write-MatrixLog ("Found {0} lists/libraries, {1:N0} total items across all sites." -f $listTotal, $S.ProgressTotal) -Level Phase
if ($ItemScanMode -eq 'Auto') {
    Write-MatrixLog ("Item scan mode: Auto (adaptive; large-list threshold {0:N0})" -f $LargeListRestThreshold) -Level Dim
}
else {
    Write-MatrixLog ("Item scan mode: {0}" -f $ItemScanMode) -Level Dim
}
if ($S.ProgressTotal -gt 25000 -and $ExpandGroups) {
    Write-Warning "Large scan ($($S.ProgressTotal.ToString('N0')) items) with 'Expand groups in matrix' enabled. Disable for a faster scan. (Refer to 'Group Members' in export for expanded groups)"
}
$S.ProgressListTitle = 'Starting scan'
Start-ProgressRenderer -Activity 'Building permissions matrix'
Write-ProgressStatus -Force

$allRows = [System.Collections.Generic.List[object]]::new()
$AllItemRows.Clear()
$uniqueBoundaries = 0
$itemsScanned = 0
$siteIndex = 0

try {
    foreach ($entry in $plan) {
        $siteIndex++
        $S.ProgressSiteIndex = $siteIndex
        $site = $entry.Site
        $path = [string]$site.ServerRelativeUrl
        $title = [string]$site.Title

        Write-Host ''
        $siteLogMsg = 'Site {0}/{1}: {2} ({3}) - {4}' -f $siteIndex, @($plan).Count, $title, $path, (Format-Elapsed $sw.Elapsed)
        Write-MatrixLog $siteLogMsg -Level Site

        Reset-SiteState
        Connect-SiteWeb $site
        if ($S.WebItemType -eq 'Subsite') { $Stats.Subsites++ } else { $Stats.RootSites++ }
        $itemsScanned += Scan-Site -AllRows $allRows -Lists $entry.Lists -UniqueBoundaryCount ([ref]$uniqueBoundaries)
    }

    Write-MatrixLog ("Unique-permission boundaries: {0:N0}" -f $uniqueBoundaries) -Level Unique
    Complete-ScanProgress
}
finally {
    Stop-ProgressRenderer -Complete:$S.ProgressScanFinished
}

$matrixCount = $allRows.Count
Write-Host ''
Write-MatrixLog ("  Matrix rows: {0:N0}" -f $matrixCount) -Level Dim
Write-MatrixLog ("  All-item rows: {0:N0}" -f $AllItemRows.Count) -Level Dim
if ($Stats.BulkSkippedItems -gt 0) {
    Write-MatrixLog ("  Inherited items without per-item matrix rows: {0:N0}" -f $Stats.BulkSkippedItems) -Level Dim
}

$sw.Stop()
# Derive a clean prefix from the root site name or URL segment.
$rootSiteLabel = ''
if ($plan -and @($plan).Count -gt 0) {
    $rootSite = @($plan)[0].Site
    if ($rootSite -and $rootSite.Title) {
        $rootSiteLabel = [string]$rootSite.Title
    }
}
if (-not $rootSiteLabel -and $S.TenantRootHost) {
    $rootSiteLabel = $S.TenantRootHost -replace '\.sharepoint\.com$', ''
}
if (-not $rootSiteLabel -and $SiteUrl) {
    try {
        if (Test-IsSharePointTenantAdminUrl $SiteUrl) {
            $rootSiteLabel = (Get-SharePointTenantHostsFromUrl $SiteUrl).TenantName
        }
        else {
            $rootSiteLabel = ([Uri]$SiteUrl).AbsolutePath -replace '^.*/', ''
        }
    }
    catch {
        $rootSiteLabel = ([Uri]$SiteUrl).AbsolutePath -replace '^.*/', ''
    }
}
$rootSiteLabel = ($rootSiteLabel -replace '[\\/:*?"<>|]', '_').Trim('_').Trim()
$filePrefix = if ($rootSiteLabel) { "$rootSiteLabel-PermissionsMatrix" } else { 'PermissionsMatrix' }
$xlsxPath = Join-Path $ScriptDir ("{0}-{1}.xlsx" -f $filePrefix, (Get-Date -Format 'yyyyMMdd-HHmmss'))

$matrixSortNote = 'Yes'
$matrixForExport = $allRows
if ($matrixCount -gt $SortMatrixRowLimit) {
    $matrixSortNote = "No (over $($SortMatrixRowLimit.ToString('N0')) rows)"
    Write-MatrixLog ("  Skipping row sort ({0:N0} rows) to speed up export." -f $matrixCount) -Level Warn
}
else {
    Write-MatrixLog ("  Sorting {0:N0} matrix rows..." -f $matrixCount) -Level Export
    $matrixForExport = [System.Collections.Generic.List[object]]::new()
    foreach ($row in @($allRows | Sort-Object 'Site Name', 'Name', 'Item path', 'Item Type', Inheritance, 'User/group')) {
        $matrixForExport.Add($row) | Out-Null
    }
}

$summary = Build-SummarySheet -SourceUrl $SiteUrl -Elapsed $sw.Elapsed `
    -IncludeSubsites $IncludeSubsites -ExpandGroups $ExpandGroups -MatrixRowCount $matrixCount `
    -MatrixSortNote $matrixSortNote
$sharingRows = Get-SharingReportRows
$groupMemberRows = Get-GroupMemberReportRows
$reportSubtitle = "Source: $SiteUrl  |  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')  |  Elapsed: $(Format-Elapsed $sw.Elapsed)"
$exportSw = [System.Diagnostics.Stopwatch]::StartNew()
Export-Workbook -MatrixRows $matrixForExport -AllItemsRows $AllItemRows -SharingRows $sharingRows `
    -GroupMemberRows $groupMemberRows -Summary $summary -Path $xlsxPath `
    -ReportSubtitle $reportSubtitle -RoleColumns @($S.RoleColumns)
$exportSw.Stop()

Write-Progress -Activity 'Writing Excel workbook' -Completed
$totalSw.Stop()
Write-Host ''
Write-MatrixLog 'Done — permissions matrix export complete.' -Level Success
Write-MatrixLog ("  Scan time        : {0}" -f (Format-Elapsed $sw.Elapsed)) -Level Export
Write-MatrixLog ("  Export time      : {0}" -f (Format-Elapsed $exportSw.Elapsed)) -Level Export
Write-MatrixLog ("  Elapsed time     : {0}" -f (Format-Elapsed $totalSw.Elapsed)) -Level Export
Write-MatrixLog ("  Sites scanned    : {0:N0}" -f $sites.Count) -Level Phase
Write-MatrixLog ("  Items scanned    : {0:N0}" -f $itemsScanned) -Level Phase
Write-MatrixLog ("  Matrix rows      : {0:N0}" -f $matrixCount) -Level Phase
Write-MatrixLog ("  All-item rows    : {0:N0}" -f $AllItemRows.Count) -Level Phase
if ($script:ScanBackendTiming.Csom.Items -gt 0 -or $script:ScanBackendTiming.Rest.Items -gt 0) {
    Write-MatrixLog ("  Backend CSOM     : {0:N0} items across {1} list(s), ~{2}/s" -f `
        $script:ScanBackendTiming.Csom.Items, $script:ScanBackendTiming.Csom.Lists, `
        ([int](Get-ScanBackendItemsPerSecond -Backend 'Csom'))) -Level Dim
    Write-MatrixLog ("  Backend REST     : {0:N0} items across {1} list(s), ~{2}/s" -f `
        $script:ScanBackendTiming.Rest.Items, $script:ScanBackendTiming.Rest.Lists, `
        ([int](Get-ScanBackendItemsPerSecond -Backend 'Rest'))) -Level Dim
}
Write-MatrixLog ("  External users   : {0:N0}" -f (Get-TotalExternalUserCount)) -Level Phase
Write-MatrixLog ("  SP groups found  : {0:N0}" -f $GroupCatalog.Count) -Level Phase
Write-MatrixLog ("  Group member rows: {0:N0}" -f $GroupMemberRows.Count) -Level Phase
Write-MatrixLog ("  Sharing link rows: {0:N0}" -f $SharingReport.UniqueLinks.Count) -Level Phase
Write-MatrixLog ("  Direct-perm rows : {0:N0}" -f $SharingReport.DirectRows.Count) -Level Phase
Write-MatrixLog ("  Excel file       : {0}" -f $xlsxPath) -Level Success

if ($OpenWorkbookOnComplete) {
    try {
        Start-Process -FilePath $xlsxPath | Out-Null
        Write-MatrixLog '  Opened workbook  : Yes' -Level Success
    }
    catch {
        Write-Warning ("Failed to open workbook automatically: {0}" -f $_.Exception.Message)
    }
}

# Always record run telemetry so future ETA estimates are automatically calibrated.
Write-RunMetricsRecord -Mode $ItemScanMode -Elapsed $totalSw.Elapsed `
    -SitesScanned $sites.Count -ItemsScanned $itemsScanned -MatrixRows $matrixCount `
    -ExportSeconds $exportSw.Elapsed.TotalSeconds
Write-MatrixLog ("  Metrics captured : {0}" -f $RunMetricsPath) -Level Dim
