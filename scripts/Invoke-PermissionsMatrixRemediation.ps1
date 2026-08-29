<#
.SYNOPSIS
    Remediates SharePoint permissions using a Permissions Matrix Excel report as input.

.DESCRIPTION
    Reads the workbook produced by Export-SitePermissionsMatrix.ps1 and offers an interactive
    menu to:
      - Remove external users (direct assignments and group membership)
      - Remove sharing links
      - Reset unique permissions (blanket inherit from parent)

    Each action shows a preview and requires explicit confirmation before making changes.

.PARAMETER ReportPath
    Path to PermissionsMatrix-*.xlsx from Export-SitePermissionsMatrix.ps1. If omitted, a file
    selection dialog opens.

.PARAMETER ClientId
    Entra app Client ID for interactive PnP sign-in.

.PARAMETER SiteUrl
    Optional override for the root site URL used to connect. If omitted, the script derives it
    from the report Summary "Source site URL", or from Group Members / Sharing Links site URLs.

.PARAMETER WhatIf
    Build and show previews only; never connect or change SharePoint.

.PARAMETER NoPicker
    Skip the GUI and use the interactive console menu (or run immediately when actions are
    implied via -WhatIf with a loaded report).

.NOTES
    Requires: PnP.PowerShell, ImportExcel

    Interactive remediation uses a WinForms picker in this script (no separate GUI file).
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string] $ReportPath,

    [Parameter(Mandatory = $false)]
    [string] $ClientId,

    [string] $SiteUrl,

    [switch] $NoPicker
)

$ErrorActionPreference = 'Stop'

foreach ($mod in @('PnP.PowerShell', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "$mod is not installed. Run: Install-Module $mod -Scope CurrentUser"
    }
}
Import-Module PnP.PowerShell, ImportExcel -ErrorAction Stop

$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue
}
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = (Get-Location).Path }

function Select-ExcelReportPath {
    param([string] $InitialDirectory = '')

    $pickerScript = {
        param([string] $StartDir)

        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        [System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null

        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = 'Select Permissions Matrix report'
        $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*'
        $dialog.FilterIndex = 1
        $dialog.CheckFileExists = $true
        $dialog.Multiselect = $false

        if (-not [string]::IsNullOrWhiteSpace($StartDir) -and (Test-Path -LiteralPath $StartDir)) {
            $dialog.InitialDirectory = (Resolve-Path -LiteralPath $StartDir).Path
        }

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.FileName
        }
        return ''
    }

    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        $ps = [PowerShell]::Create()
        $null = $ps.AddScript($pickerScript).AddArgument($InitialDirectory)
        $ps.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $ps.Runspace.ApartmentState = 'STA'
        $ps.Runspace.Open()
        $picked = @($ps.Invoke()) | Select-Object -First 1
        $ps.Runspace.Close()
        $ps.Dispose()
        return [string]$picked
    }

    return & $pickerScript $InitialDirectory
}

$BaseMatrixColumns = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase
)
@(
    'Site Name', 'Item path', 'Item Type', 'Inheritance', 'Details',
    'User/group', 'Principal type', 'Account name', 'External user', 'Given through'
) | ForEach-Object { [void]$BaseMatrixColumns.Add($_) }

$State = @{
    ConnectedUrl = ''
    SiteHost     = ''
    SiteMap      = @{}   # Site Name -> Site URL
    RoleColumns  = @()
}

function Write-MenuTitle([string] $Text) {
    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('=' * [Math]::Min(72, $Text.Length)) -ForegroundColor DarkGray
}

function Test-ExternalLogin {
    param([string] $Login, [string] $Email = '')

    return ($Login -match '#ext#' -or $Email -match '#EXT#@')
}

function Get-NormalizedSiteUrl {
    param([string] $Url)

    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    $uri = [Uri]$Url.Trim()
    $p = $uri.AbsolutePath
    if ($p -ne '/' -and $p.EndsWith('/')) { $p = $p.TrimEnd('/') }
    return ('{0}://{1}{2}' -f $uri.Scheme, $uri.Host, $p).ToLowerInvariant()
}

function Get-ServerRelativePath {
    param([string] $UrlOrPath)

    if ([string]::IsNullOrWhiteSpace($UrlOrPath)) { return '' }
    $t = $UrlOrPath.Trim()
    if ($t -match '^https?://') {
        $uri = [Uri]$t
        $p = $uri.AbsolutePath
        if ($p -ne '/' -and $p.EndsWith('/')) { $p = $p.TrimEnd('/') }
        return $p
    }
    if (-not $t.StartsWith('/')) { $t = "/$t" }
    if ($t -ne '/' -and $t.EndsWith('/')) { $t = $t.TrimEnd('/') }
    return $t
}

function Normalize-ServerRelativePath {
    param([string] $Path)

    $p = Get-ServerRelativePath $Path
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    return ($p -replace '\\', '/')
}

function Get-ParentServerRelativePath {
    param([string] $Path)

    $p = (Normalize-ServerRelativePath $Path).TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    $idx = $p.LastIndexOf('/')
    if ($idx -le 0) { return $p }
    return $p.Substring(0, $idx)
}

function Get-SiteWebUrlFromPath {
    param(
        [string] $SiteUrl,
        [string] $ItemPath
    )

    $webPath = Normalize-ServerRelativePath $ItemPath
    if ([string]::IsNullOrWhiteSpace($webPath)) { return (Get-NormalizedSiteUrl $SiteUrl) }

    $uri = [Uri](Get-NormalizedSiteUrl $SiteUrl)
    return Get-NormalizedSiteUrl ('{0}://{1}{2}' -f $uri.Scheme, $uri.Host, $webPath)
}

function Resolve-PnPListFromPath {
    param([string] $ListPathOrTitle)

    $path = Normalize-ServerRelativePath $ListPathOrTitle
    if ([string]::IsNullOrWhiteSpace($path)) {
        throw 'List path was empty.'
    }

    foreach ($identity in @($path, ($path.TrimEnd('/') -split '/')[-1])) {
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        try {
            return Get-PnPList -Identity $identity -ErrorAction Stop
        }
        catch {
            Write-Verbose "Get-PnPList failed for identity '$identity': $_"
        }
    }

    throw "Could not resolve list '$path'."
}

function Normalize-RemediationItemType {
    param([string] $ItemType)

    switch ($ItemType) {
        'Item' { return 'List item' }
        default { return $ItemType }
    }
}

function Test-IsSiteCollectionRootWeb {
    param(
        [string] $ItemPath,
        $Web = $null
    )

    if (-not $Web) { $Web = Get-PnPWeb }
    Get-PnPProperty -ClientObject $Web -Property ServerRelativeUrl, ParentWeb | Out-Null

    $targetPath = (Normalize-ServerRelativePath $ItemPath).TrimEnd('/')
    $webPath = ([string]$Web.ServerRelativeUrl).TrimEnd('/')
    if ($targetPath -and $targetPath -ne $webPath) { return $false }

    try {
        if ($Web.ParentWeb.ServerObjectIsNull) { return $true }
        Get-PnPProperty -ClientObject $Web.ParentWeb -Property ServerRelativeUrl | Out-Null
        $parentPath = ([string]$Web.ParentWeb.ServerRelativeUrl).TrimEnd('/')
        if ([string]::IsNullOrWhiteSpace($parentPath) -or $parentPath -eq '/') { return $true }
    }
    catch {
        return $true
    }
    return $false
}

function Get-ListItemContextFromPath-PnP {
    param(
        [string] $ServerRelativePath,
        [string] $ItemType
    )

    $path = Normalize-ServerRelativePath $ServerRelativePath
    $item = $null

    if ($ItemType -eq 'Folder') {
        $folder = Get-PnPFolder -Url $path -ErrorAction Stop
        Get-PnPProperty -ClientObject $folder -Property ListItemAllFields | Out-Null
        $item = $folder.ListItemAllFields
    }
    else {
        foreach ($attempt in @('File', 'Folder')) {
            try {
                if ($attempt -eq 'File') {
                    $item = Get-PnPFile -Url $path -AsListItem -ErrorAction Stop
                }
                else {
                    $folder = Get-PnPFolder -Url $path -ErrorAction Stop
                    Get-PnPProperty -ClientObject $folder -Property ListItemAllFields | Out-Null
                    $item = $folder.ListItemAllFields
                }
                break
            }
            catch {
                Write-Verbose "PnP $attempt lookup failed for '$path': $_"
            }
        }
    }

    if (-not $item) { throw "PnP could not resolve list item for '$path'." }

    Get-PnPProperty -ClientObject $item -Property ParentList | Out-Null
    if (-not $item.ParentList) { throw "Parent list was missing for '$path'." }

    return @{
        ListId    = [string]$item.ParentList.Id
        ListTitle = [string]$item.ParentList.Title
        ItemId    = [int]$item.Id
    }
}

function Get-ListItemContextFromPath-ByFileRef {
    param([string] $ServerRelativePath)

    $path = Normalize-ServerRelativePath $ServerRelativePath
    $current = $path
    while ($current -and $current.Length -gt 1) {
        try {
            $list = Resolve-PnPListFromPath -ListPathOrTitle $current
            $itemId = Resolve-ListItemId -ListTitle $list.Title -ServerRelativePath $path
            if ($itemId -gt 0) {
                return @{
                    ListId    = [string]$list.Id
                    ListTitle = [string]$list.Title
                    ItemId    = $itemId
                }
            }
        }
        catch {
            Write-Verbose "List/item lookup failed at '$current': $_"
        }
        $current = Get-ParentServerRelativePath $current
    }
    throw "Could not resolve list item by FileRef for '$path'."
}

function Get-ListItemContextFromPath {
    param(
        [string] $ServerRelativePath,
        [string] $ItemType
    )

    $path = Normalize-ServerRelativePath $ServerRelativePath
    if ([string]::IsNullOrWhiteSpace($path)) {
        throw 'Item path was empty.'
    }

    $itemType = Normalize-RemediationItemType $ItemType
    $escaped = $path.Replace("'", "''")
    $apiPaths = switch ($itemType) {
        'Folder'    { @('GetFolderByServerRelativeUrl') }
        'File'      { @('GetFileByServerRelativeUrl') }
        'List item' { @('GetFileByServerRelativeUrl', 'GetFolderByServerRelativeUrl') }
        default     { @('GetFileByServerRelativeUrl', 'GetFolderByServerRelativeUrl') }
    }

    $lastError = $null
    foreach ($apiPath in $apiPaths) {
        try {
            $url = "/_api/web/$apiPath('$escaped')/ListItemAllFields?`$select=Id,FileRef&`$expand=ParentList"
            $r = Invoke-PnPSPRestMethod -Url $url -Method Get
            if (-not $r.Id -or -not $r.ParentList) {
                throw "List item fields were incomplete for $path"
            }
            return @{
                ListId    = [string]$r.ParentList.Id
                ListTitle = [string]$r.ParentList.Title
                ItemId    = [int]$r.Id
            }
        }
        catch {
            $lastError = $_
        }
    }

    foreach ($fallback in @(
        { Get-ListItemContextFromPath-PnP -ServerRelativePath $path -ItemType $itemType }
        { Get-ListItemContextFromPath-ByFileRef -ServerRelativePath $path }
    )) {
        try {
            return & $fallback
        }
        catch {
            $lastError = $_
        }
    }

    throw "Could not resolve list item for '$path': $lastError"
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

function Unprotect-ToolkitClientId {
    param([string] $Encoded)

    if ([string]::IsNullOrWhiteSpace($Encoded)) { return '' }
    if (-not $Encoded.StartsWith('SPOTK1:', [StringComparison]::Ordinal)) { return '' }
    try {
        $bytes = [Convert]::FromBase64String($Encoded.Substring(7))
        $key = [System.Text.Encoding]::UTF8.GetBytes('SPOToolkit-v1')
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            $bytes[$i] = $bytes[$i] -bxor $key[$i % $key.Length]
        }
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    catch {
        return ''
    }
}

function Get-ToolkitMetaSheetName { return '_ToolkitMeta' }

function Read-ToolkitMetaSheetClientId {
    param([OfficeOpenXml.ExcelPackage] $Package)

    $sheetName = Get-ToolkitMetaSheetName
    $ws = $Package.Workbook.Worksheets[$sheetName]
    if (-not $ws) { return '' }

    $raw = [string]$ws.Cells['A1'].Value
    if ([string]::IsNullOrWhiteSpace($raw)) { $raw = [string]$ws.Cells['A1'].Text }
    return (Unprotect-ToolkitClientId $raw)
}

function Read-SummaryEmbeddedClientIdFromPackage {
    param([OfficeOpenXml.ExcelPackage] $Package)

    $ws = $Package.Workbook.Worksheets['Summary']
    if (-not $ws) { return '' }

    $metricCol = -1
    $valueCol = -1
    $dataStart = 12
    $dataEnd = 250

    $tbl = $ws.Tables['SummaryOverview']
    if ($tbl) {
        $hdrRow = $tbl.Address.Start.Row
        $dataStart = $hdrRow + 1
        $dataEnd = $tbl.Address.End.Row
        for ($c = $tbl.Address.Start.Column; $c -le $tbl.Address.End.Column; $c++) {
            $h = [string]$ws.Cells[$hdrRow, $c].Value
            if ($h -eq 'Metric') { $metricCol = $c }
            if ($h -eq 'Value') { $valueCol = $c }
        }
    }

    if ($metricCol -lt 1 -or $valueCol -lt 1) {
        $metricCol = 4
        $valueCol = 5
    }

    for ($r = $dataStart; $r -le $dataEnd; $r++) {
        $m = [string]$ws.Cells[$r, $metricCol].Value
        if ([string]::IsNullOrWhiteSpace($m)) { $m = [string]$ws.Cells[$r, $metricCol].Text }
        if ($m -ne 'Toolkit connection profile') { continue }

        $v = [string]$ws.Cells[$r, $valueCol].Value
        if ([string]::IsNullOrWhiteSpace($v)) { $v = [string]$ws.Cells[$r, $valueCol].Text }
        return (Unprotect-ToolkitClientId $v)
    }

    return ''
}

function Read-ReportEmbeddedClientIdFromPackage {
    param([OfficeOpenXml.ExcelPackage] $Package)

    $id = Read-ToolkitMetaSheetClientId -Package $Package
    if ($id) { return $id }
    return (Read-SummaryEmbeddedClientIdFromPackage -Package $Package)
}

function Get-ReportEmbeddedClientId {
    param([string] $Path)

    if (-not (Test-Path -LiteralPath $Path)) { return '' }

    $pkg = $null
    try {
        $pkg = Open-ExcelPackage -Path $Path
        return (Read-ReportEmbeddedClientIdFromPackage -Package $pkg)
    }
    catch {
        return ''
    }
    finally {
        if ($pkg) { Close-ExcelPackage $pkg }
    }
}

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
        return $false
    }
}

function Import-ReportWorkbook {
    param([string] $Path)

    Write-Host "Loading report: $Path" -ForegroundColor Cyan

    $embeddedClientId = ''

    $matrix = @(Import-Excel -Path $Path -WorksheetName 'Permissions Matrix' -ErrorAction Stop)
    $sharing = @()
    $groups = @()
    $summaryUrl = ''

    try {
        $sharing = @(Import-Excel -Path $Path -WorksheetName 'Sharing Links' -ErrorAction Stop)
    }
    catch {
        Write-Warning "Sharing Links sheet not found or could not be read."
    }

    try {
        $groups = @(Import-Excel -Path $Path -WorksheetName 'Group Members' -ErrorAction Stop)
    }
    catch {
        Write-Warning "Group Members sheet not found or could not be read."
    }

    try {
        $summary = @(Import-Excel -Path $Path -WorksheetName 'Summary' -StartRow 11 -StartColumn 3 `
            -ErrorAction Stop)
        $srcRow = @($summary | Where-Object { [string]$_.Metric -eq 'Source site URL' } | Select-Object -First 1)
        if ($srcRow) {
            $summaryUrl = [string]$srcRow.Value
        }
        if (-not $embeddedClientId) {
            $profileRow = @($summary | Where-Object {
                [string]$_.Metric -eq 'Toolkit connection profile'
            } | Select-Object -First 1)
            if ($profileRow) {
                $embeddedClientId = Unprotect-ToolkitClientId ([string]$profileRow.Value)
            }
        }
    }
    catch {
        Write-Verbose "Summary sheet not read: $_"
    }

    if (-not $embeddedClientId) {
        $pkgEmbed = $null
        try {
            $pkgEmbed = Open-ExcelPackage -Path $Path
            $embeddedClientId = Read-ReportEmbeddedClientIdFromPackage -Package $pkgEmbed
        }
        catch {
            Write-Verbose "Embedded Client ID read failed: $_"
        }
        finally {
            if ($pkgEmbed) { Close-ExcelPackage $pkgEmbed }
        }
    }

    if ($matrix.Count -lt 1) { throw 'Permissions Matrix sheet is empty or missing.' }

    $roleCols = @($matrix[0].PSObject.Properties.Name | Where-Object {
        -not $BaseMatrixColumns.Contains([string]$_)
    })

    return @{
        Matrix           = $matrix
        Sharing          = $sharing
        Groups           = $groups
        SummaryUrl       = $summaryUrl
        EmbeddedClientId = $embeddedClientId
        RoleColumns      = $roleCols
    }
}

function Get-ReportRootSiteUrl {
    param(
        [string]   $SummaryUrl,
        [object[]] $Matrix,
        [object[]] $Sharing,
        [object[]] $Groups
    )

    if (-not [string]::IsNullOrWhiteSpace($SummaryUrl)) {
        return (Get-NormalizedSiteUrl $SummaryUrl.Trim())
    }

    $urls = [System.Collections.Generic.List[string]]::new()
    foreach ($g in @($Groups)) {
        $u = Get-NormalizedSiteUrl ([string]$g.'Site URL')
        if ($u) { [void]$urls.Add($u) }
    }
    foreach ($s in @($Sharing)) {
        $u = Get-NormalizedSiteUrl ([string]$s.SiteUrl)
        if ($u) { [void]$urls.Add($u) }
    }

    if ($urls.Count -lt 1) { return '' }

    $hostGroup = @($urls | Group-Object { ([Uri]$_).Host } | Sort-Object Count -Descending)
    $hostUrls = @($hostGroup[0].Group)
    $ordered = @($hostUrls | Sort-Object { ([Uri]$_).AbsolutePath.TrimEnd('/').Length })

    foreach ($candidate in $ordered) {
        $prefix = ([Uri]$candidate).AbsolutePath.TrimEnd('/')
        if ([string]::IsNullOrEmpty($prefix)) { $prefix = '/' }
        $ok = $true
        foreach ($u in $hostUrls) {
            $p = ([Uri]$u).AbsolutePath.TrimEnd('/')
            if ($prefix -ne '/' -and $p -notlike "$prefix*") { $ok = $false; break }
        }
        if ($ok) { return (Get-NormalizedSiteUrl $candidate) }
    }

    return (Get-NormalizedSiteUrl $ordered[0])
}

function Initialize-SiteMap {
    param(
        [object[]] $Matrix,
        [object[]] $Sharing,
        [object[]] $Groups,
        [string]   $SummaryUrl,
        [string]   $DefaultSiteUrl
    )

    $map = @{}

    foreach ($g in $Groups) {
        $name = [string]$g.'Site Name'
        $url = [string]$g.'Site URL'
        if ($name -and $url) { $map[$name] = (Get-NormalizedSiteUrl $url) }
    }

    foreach ($s in $Sharing) {
        $url = [string]$s.SiteUrl
        if (-not $url) { continue }
        $norm = Get-NormalizedSiteUrl $url
        $webPath = Get-ServerRelativePath $norm
        $siteName = ($webPath -split '/')[-1]
        if ($siteName) { $map[$siteName] = $norm }
    }

    foreach ($row in $Matrix) {
        $name = [string]$row.'Site Name'
        if (-not $name -or $map.ContainsKey($name)) { continue }
        $itemPath = Get-ServerRelativePath ([string]$row.'Item path')
        if (-not $itemPath) { continue }
        $itemPath = Normalize-ServerRelativePath $itemPath
        # Best-effort: map site name to the web path prefix from item paths in the matrix.
        $parts = $itemPath -split '/'
        if ($parts.Count -ge 3) {
            $guess = Get-NormalizedSiteUrl ("{0}://{1}" -f ([Uri]$DefaultSiteUrl).Scheme, ([Uri]$DefaultSiteUrl).Host)
            if ($parts[1] -eq 'sites' -and $parts.Count -ge 4) {
                $webRel = ('/' + ($parts[1..3] -join '/'))
                $guess = Get-NormalizedSiteUrl ("{0}://{1}{2}" -f ([Uri]$DefaultSiteUrl).Scheme, ([Uri]$DefaultSiteUrl).Host, $webRel)
            }
            $map[$name] = $guess
        }
    }

    if ($SummaryUrl) {
        $State.SiteHost = '{0}://{1}' -f ([Uri]$SummaryUrl).Scheme, ([Uri]$SummaryUrl).Host
    }
    elseif ($DefaultSiteUrl) {
        $State.SiteHost = '{0}://{1}' -f ([Uri]$DefaultSiteUrl).Scheme, ([Uri]$DefaultSiteUrl).Host
    }

    $State.SiteMap = $map
}

function Get-SiteUrlForRow {
    param(
        [object] $Row,
        [string] $FallbackSiteUrl
    )

    $name = [string]$Row.'Site Name'
    if ($name -and $State.SiteMap.ContainsKey($name)) {
        return $State.SiteMap[$name]
    }
    return (Get-NormalizedSiteUrl $FallbackSiteUrl)
}

function Get-RolesFromMatrixRow {
    param(
        [object] $Row,
        [string[]] $RoleColumns
    )

    $roles = [System.Collections.Generic.List[string]]::new()
    foreach ($col in $RoleColumns) {
        if ([string]$Row.$col -eq 'X') { $roles.Add($col) | Out-Null }
    }
    return @($roles)
}

function Build-ExternalUserPlan {
    param(
        [object[]] $Matrix,
        [object[]] $Groups,
        [string[]] $RoleColumns,
        [string]   $FallbackSiteUrl
    )

    $plan = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($row in $Matrix) {
        if ([string]$row.'External user' -ne 'Yes') { continue }
        if ([string]$row.'Given through' -eq 'Sharing Link') { continue }
        if ([string]$row.Inheritance -eq 'Inherited') { continue }

        $ptype = [string]$row.'Principal type'
        $login = [string]$row.'Account name'
        if (-not $login) { $login = [string]$row.'User/group' }
        if (-not $login) { continue }

        $siteUrl = Get-SiteUrlForRow -Row $row -FallbackSiteUrl $FallbackSiteUrl
        $itemPath = Normalize-ServerRelativePath ([string]$row.'Item path')
        $itemType = [string]$row.'Item Type'
        $roles = Get-RolesFromMatrixRow -Row $row -RoleColumns $RoleColumns

        if ($ptype -eq 'User') {
            $key = "user|$siteUrl|$itemType|$itemPath|$login"
            if ($seen.Add($key)) {
                $plan.Add([pscustomobject]@{
                    Action    = 'RemoveUserPermission'
                    SiteUrl   = $siteUrl
                    ItemPath  = $itemPath
                    ItemType  = $itemType
                    Login     = $login
                    Group     = ''
                    Roles     = ($roles -join ', ')
                    Detail    = "Remove all permission assignments for external user $login at $itemType '$itemPath'"
                }) | Out-Null
            }
        }
        elseif ($ptype -eq 'SharePoint group' -and $roles.Count -gt 0) {
            # External user may appear as expanded row under a group in some exports; handled via Groups sheet.
        }
    }

    foreach ($g in $Groups) {
        if ([string]$g.'External user' -ne 'Yes') { continue }
        $login = [string]$g.'Member Login'
        if (-not $login) { $login = [string]$g.'Member Email' }
        if (-not $login) { continue }

        $siteUrl = Get-NormalizedSiteUrl ([string]$g.'Site URL')
        if (-not $siteUrl) { $siteUrl = Get-NormalizedSiteUrl $FallbackSiteUrl }
        $group = [string]$g.'Group Name'
        $key = "group|$siteUrl|$group|$login"
        if ($seen.Add($key)) {
            $plan.Add([pscustomobject]@{
                Action    = 'RemoveGroupMember'
                SiteUrl   = $siteUrl
                ItemPath  = Get-ServerRelativePath ([string]$g.'Site Path')
                ItemType  = 'GroupMember'
                Login     = $login
                Group     = $group
                Roles     = ''
                Detail    = "Remove external member $login from group '$group'"
            }) | Out-Null
        }
    }

    foreach ($target in @($plan)) {
        if ($target.ItemType -notmatch '^(File|Folder|List item)$') { continue }

        $siteUrl = [string]$target.SiteUrl
        $login = [string]$target.Login
        $webPath = Normalize-ServerRelativePath ([Uri]$siteUrl).AbsolutePath

        $siteKey = "user|$siteUrl|Site|$webPath|$login"
        if ($seen.Add($siteKey)) {
            $plan.Add([pscustomobject]@{
                Action    = 'RemoveUserPermission'
                SiteUrl   = $siteUrl
                ItemPath  = $webPath
                ItemType  = 'Site'
                Login     = $login
                Group     = ''
                Roles     = ''
                Detail    = "Remove site-level grants for external user $login (sharing cleanup)"
            }) | Out-Null
        }

        $listPath = Get-ParentServerRelativePath ([string]$target.ItemPath)
        if ($listPath) {
            $listKey = "user|$siteUrl|Library|$listPath|$login"
            if ($seen.Add($listKey)) {
                $plan.Add([pscustomobject]@{
                    Action    = 'RemoveUserPermission'
                    SiteUrl   = $siteUrl
                    ItemPath  = $listPath
                    ItemType  = 'Library'
                    Login     = $login
                    Group     = ''
                    Roles     = ''
                    Detail    = "Remove library-level grants for external user $login (sharing cleanup)"
                }) | Out-Null
            }
        }
    }

    return @($plan)
}

function Build-SharingLinkPlan {
    param([object[]] $Sharing)

    $plan = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($row in $Sharing) {
        if ([string]$row.RecordType -ne 'SharingLink') { continue }
        $shareId = [string]$row.ShareId
        $rel = Get-ServerRelativePath ([string]$row.RelativeUrl)
        if (-not $shareId -and -not $rel) { continue }

        $siteUrl = Get-NormalizedSiteUrl ([string]$row.SiteUrl)
        $key = "$siteUrl|$shareId|$rel"
        if (-not $seen.Add($key)) { continue }

        $plan.Add([pscustomobject]@{
            Action     = 'RemoveSharingLink'
            SiteUrl    = $siteUrl
            ItemPath   = $rel
            ItemType   = [string]$row.ObjectType
            ListTitle  = [string]$row.ListTitle
            ListUrl    = Get-ServerRelativePath ([string]$row.ListUrl)
            ItemId     = [int]$row.ItemId
            ShareId    = $shareId
            ShareUrl   = [string]$row.ShareLinkUrl
            LinkType   = [string]$row.ShareLinkType
            Detail     = "Remove sharing link ($([string]$row.ShareLinkType)) on $rel"
        }) | Out-Null
    }

    return @($plan)
}

function Build-UniquePermissionResetPlan {
    param(
        [object[]] $Matrix,
        [object[]] $Sharing,
        [string]   $FallbackSiteUrl
    )

    $plan = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($row in $Matrix) {
        if ([string]$row.Inheritance -ne 'Custom') { continue }
        $itemPath = Get-ServerRelativePath ([string]$row.'Item path')
        if (-not $itemPath) { continue }
        $itemPath = Normalize-ServerRelativePath $itemPath
        $siteUrl = Get-SiteUrlForRow -Row $row -FallbackSiteUrl $FallbackSiteUrl
        $itemPath = Normalize-ServerRelativePath ([string]$row.'Item path')
        $itemType = Normalize-RemediationItemType ([string]$row.'Item Type')
        $key = "$siteUrl|$itemPath"
        if (-not $seen.Add($key)) { continue }

        $plan.Add([pscustomobject]@{
            Action   = 'ResetUniquePermissions'
            SiteUrl  = $siteUrl
            ItemPath = $itemPath
            ItemType = $itemType
            ItemId   = 0
            ListUrl  = ''
            Detail   = "Reset unique permissions on $itemType '$itemPath' (inherit from parent)"
        }) | Out-Null
    }

    foreach ($row in $Sharing) {
        if ([string]$row.RecordType -ne 'DirectPermission') { continue }
        $itemPath = Normalize-ServerRelativePath ([string]$row.RelativeUrl)
        if (-not $itemPath) { continue }
        $siteUrl = Get-NormalizedSiteUrl ([string]$row.SiteUrl)
        $itemType = Normalize-RemediationItemType ([string]$row.ObjectType)
        $key = "$siteUrl|$itemPath"
        if (-not $seen.Add($key)) { continue }

        $plan.Add([pscustomobject]@{
            Action   = 'ResetUniquePermissions'
            SiteUrl  = $siteUrl
            ItemPath = $itemPath
            ItemType = $itemType
            ItemId   = [int]$row.ItemId
            ListUrl  = Normalize-ServerRelativePath ([string]$row.ListUrl)
            Detail   = "Reset unique permissions on $itemType '$itemPath' (inherit from parent)"
        }) | Out-Null
    }

    return @($plan | Sort-Object { [string]$_.ItemPath.Length } -Descending)
}

function Build-UniqueItemPrincipalIndex {
    param(
        [object[]] $Matrix,
        [string]   $FallbackSiteUrl,
        [string[]] $RoleColumns,
        [System.Collections.Generic.HashSet[string]] $UniqueItemKeys
    )

    $index = @{}
    $seenPrincipal = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($row in @($Matrix)) {
        $itemPath = Get-ServerRelativePath ([string]$row.'Item path')
        if (-not $itemPath) { continue }
        $itemPath = Normalize-ServerRelativePath $itemPath
        $siteUrl = Get-SiteUrlForRow -Row $row -FallbackSiteUrl $FallbackSiteUrl
        $siteUrl = Get-NormalizedSiteUrl $siteUrl
        $itemKey = "$siteUrl|$itemPath"
        if (-not $UniqueItemKeys.Contains($itemKey)) { continue }

        $login = [string]$row.'Account name'
        if (-not $login) { $login = [string]$row.'User/group' }
        if (-not $login) { continue }

        $principalKey = "$itemKey|$login"
        if (-not $seenPrincipal.Add($principalKey)) { continue }

        if (-not $index.ContainsKey($itemKey)) {
            $index[$itemKey] = [System.Collections.Generic.List[object]]::new()
        }

        $roles = @(Get-RolesFromMatrixRow -Row $row -RoleColumns $RoleColumns)
        $index[$itemKey].Add([pscustomobject]@{
            Login         = $login
            PrincipalType = [string]$row.'Principal type'
            External      = ([string]$row.'External user' -eq 'Yes')
            GivenThrough  = [string]$row.'Given through'
            Roles         = ($roles -join ', ')
            Inheritance   = [string]$row.Inheritance
            SiteUrl       = $siteUrl
            ItemPath      = $itemPath
        }) | Out-Null
    }

    return $index
}

function Get-UniqueItemLookupKey {
    param([object] $UniquePlanItem)

    $siteUrl = Get-NormalizedSiteUrl ([string]$UniquePlanItem.SiteUrl)
    $itemPath = Normalize-ServerRelativePath ([string]$UniquePlanItem.ItemPath)
    return "$siteUrl|$itemPath"
}

function Show-RemediationPreview {
    param(
        [string]   $Title,
        [object[]] $Plan,
        [int]      $SampleCount = 25
    )

    Write-MenuTitle $Title
    Write-Host ("Total targets: {0:N0}" -f @($Plan).Count) -ForegroundColor White

    if (@($Plan).Count -lt 1) {
        Write-Host 'Nothing to do for this action.' -ForegroundColor DarkYellow
        return
    }

    $sample = @($Plan | Select-Object -First $SampleCount)
    $sample | Format-Table -Property @(
        @{ N = 'Site'; E = { ($_.SiteUrl -replace '^https?://[^/]+', '') } ; Width = 28 }
        'ItemType', 'ItemPath', 'Login', 'Group', 'ShareId', 'Detail'
    ) -Wrap -AutoSize | Out-String | Write-Host

    if (@($Plan).Count -gt $SampleCount) {
        Write-Host ("... and {0:N0} more (showing first {1})" -f (@($Plan).Count - $SampleCount), $SampleCount) `
            -ForegroundColor DarkGray
    }
}

function Confirm-RemediationAction {
    param(
        [string] $Prompt,
        [switch] $UseMessageBox
    )

    if ($WhatIfPreference) {
        Write-Host '[WhatIf] No changes will be made.' -ForegroundColor Yellow
        return $false
    }

    if ($UseMessageBox) {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $msg = "$Prompt`n`nThis cannot be undone. Continue?"
        $result = [System.Windows.Forms.MessageBox]::Show(
            $msg, 'Confirm remediation', 'YesNo', 'Warning')
        return ($result -eq 'Yes')
    }

    Write-Host ''
    $answer = Read-Host "$Prompt Type YES to confirm"
    return ($answer.Trim().ToUpperInvariant() -eq 'YES')
}

function Connect-PnPSiteIfNeeded {
    param([string] $TargetSiteUrl)

    $target = Get-NormalizedSiteUrl $TargetSiteUrl
    if (-not $target) { throw 'Site URL was empty.' }
    if ($State.ConnectedUrl -eq $target) { return }

    Connect-PnPOnline -Url $target -Interactive -ClientId $ClientId | Out-Null
    $State.ConnectedUrl = $target
}

function Resolve-ListItemId {
    param(
        [string] $ListTitle,
        [string] $ServerRelativePath
    )

    if ([string]::IsNullOrWhiteSpace($ServerRelativePath)) { return 0 }

    $path = Normalize-ServerRelativePath $ServerRelativePath
    $escaped = $path.Replace("'", "''")
    $filter = [Uri]::EscapeDataString("FileRef eq '$escaped'")
    $list = Resolve-PnPListFromPath -ListPathOrTitle $ListTitle
    $listId = $list.Id.ToString().ToLower()
    $url = "/_api/web/lists(guid'$listId')/items?`$select=Id&`$filter=$filter&`$top=1"
    $r = Invoke-PnPSPRestMethod -Url $url -Method Get
    if ($r.value -and @($r.value).Count -gt 0) {
        return [int]$r.value[0].Id
    }
    return 0
}

function Invoke-ResetUniquePermissionTarget {
    param([object] $Target)

    Connect-PnPSiteIfNeeded -TargetSiteUrl $Target.SiteUrl

    $itemType = Normalize-RemediationItemType ([string]$Target.ItemType)
    $itemPath = Normalize-ServerRelativePath ([string]$Target.ItemPath)
    $listUrl = Normalize-ServerRelativePath ([string]$Target.ListUrl)
    $itemId = [int]$Target.ItemId

    switch -Regex ($itemType) {
        '^(Site|Subsite)$' {
            $webUrl = Get-SiteWebUrlFromPath -SiteUrl $Target.SiteUrl -ItemPath $itemPath
            Connect-PnPSiteIfNeeded -TargetSiteUrl $webUrl
            $web = Get-PnPWeb
            if (Test-IsSiteCollectionRootWeb -ItemPath $itemPath -Web $web) {
                throw 'SKIP: The site collection root web cannot reset inheritance via API. Remove unique permissions manually in Site Settings > Site permissions.'
            }
            Get-PnPProperty -ClientObject $web -Property HasUniqueRoleAssignments | Out-Null
            if ($web.HasUniqueRoleAssignments) {
                $web.ResetRoleInheritance()
                $web.Update()
                Invoke-PnPQuery
            }
        }
        '^(List|Library)$' {
            $list = Resolve-PnPListFromPath -ListPathOrTitle $itemPath
            Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments | Out-Null
            if ($list.HasUniqueRoleAssignments) {
                $list.ResetRoleInheritance()
                $list.Update()
                Invoke-PnPQuery
            }
        }
        '^(File|Folder|List item)$' {
            if ($itemId -lt 1) {
                $ctx = Get-ListItemContextFromPath -ServerRelativePath $itemPath -ItemType $itemType
                $itemId = $ctx.ItemId
                $listTitle = $ctx.ListTitle
            }
            else {
                if (-not $listUrl) {
                    $listUrl = Get-ParentServerRelativePath $itemPath
                }
                $listTitle = (Resolve-PnPListFromPath -ListPathOrTitle $listUrl).Title
            }
            if ($itemId -lt 1) {
                throw "Could not resolve list item id for $itemPath"
            }
            $list = Resolve-PnPListFromPath -ListPathOrTitle $listTitle
            if ($itemType -eq 'Folder') {
                Get-PnPFolder -Url $itemPath -ErrorAction Stop | Set-PnPFolderPermission -List $list `
                    -InheritPermissions -SystemUpdate -ErrorAction Stop
            }
            else {
                Set-PnPListItemPermission -List $list -Identity $itemId -InheritPermissions `
                    -SystemUpdate -ErrorAction Stop
            }
        }
        default {
            throw "Unsupported item type for reset: $itemType"
        }
    }
}

function Remove-ExternalUserAssignmentAtScope {
    param(
        [string] $Login,
        [ValidateSet('Web', 'List', 'ListItem')]
        [string] $ScopeKind,
        $List = $null,
        [int]    $ItemId = 0
    )

    $user = Get-PnPUser -Identity $Login -ErrorAction Stop
    $removed = $false

    switch ($ScopeKind) {
        'Web' {
            $web = Get-PnPWeb
            try {
                $ra = $web.RoleAssignments.GetByPrincipal($user)
                Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings, Member | Out-Null
                $ra.DeleteObject()
                Invoke-PnPQuery
                $removed = $true
            }
            catch {
                if ($_.Exception.Message -notmatch '(?i)not found|does not exist|no role assignment|user cannot be found') {
                    throw
                }
            }
        }
        'List' {
            if (-not $List) { throw 'List object was required.' }
            try {
                Get-PnPProperty -ClientObject $List -Property RoleAssignments | Out-Null
                $ra = $List.RoleAssignments.GetByPrincipal($user)
                Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings | Out-Null
                $ra.DeleteObject()
                $List.Update()
                Invoke-PnPQuery
                $removed = $true
            }
            catch {
                if ($_.Exception.Message -notmatch '(?i)not found|does not exist|no role assignment|user cannot be found') {
                    throw
                }
            }
        }
        'ListItem' {
            if (-not $List -or $ItemId -lt 1) { throw 'List and ItemId were required.' }
            $item = Get-PnPListItem -List $List -Id $ItemId
            try {
                Get-PnPProperty -ClientObject $item -Property RoleAssignments | Out-Null
                $ra = $item.RoleAssignments.GetByPrincipal($user)
                Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings | Out-Null
                $ra.DeleteObject()
                $item.Update()
                Invoke-PnPQuery
                $removed = $true
            }
            catch {
                if ($_.Exception.Message -notmatch '(?i)not found|does not exist|no role assignment|user cannot be found') {
                    throw
                }
            }
        }
    }

    return $removed
}

function Invoke-RemoveExternalUserTarget {
    param(
        [object] $Target,
        [string[]] $RoleColumns
    )

    Connect-PnPSiteIfNeeded -TargetSiteUrl $Target.SiteUrl

    if ($Target.Action -eq 'RemoveGroupMember') {
        Remove-PnPGroupMember -Group $Target.Group -LoginName $Target.Login -ErrorAction Stop
        return
    }

    $itemType = Normalize-RemediationItemType ([string]$Target.ItemType)
    $login = [string]$Target.Login
    $itemPath = Normalize-ServerRelativePath ([string]$Target.ItemPath)
    $removedAny = $false

    switch -Regex ($itemType) {
        '^(Site|Subsite)$' {
            $webUrl = Get-SiteWebUrlFromPath -SiteUrl $Target.SiteUrl -ItemPath $itemPath
            Connect-PnPSiteIfNeeded -TargetSiteUrl $webUrl
            if (Remove-ExternalUserAssignmentAtScope -Login $login -ScopeKind Web) {
                $removedAny = $true
            }
        }
        '^(List|Library)$' {
            $list = Resolve-PnPListFromPath -ListPathOrTitle $itemPath
            if (Remove-ExternalUserAssignmentAtScope -Login $login -ScopeKind List -List $list) {
                $removedAny = $true
            }
        }
        '^(File|Folder|List item)$' {
            $ctx = Get-ListItemContextFromPath -ServerRelativePath $itemPath -ItemType $itemType
            $list = Resolve-PnPListFromPath -ListPathOrTitle $ctx.ListTitle
            if (Remove-ExternalUserAssignmentAtScope -Login $login -ScopeKind ListItem -List $list -ItemId $ctx.ItemId) {
                $removedAny = $true
            }
            if (Remove-ExternalUserAssignmentAtScope -Login $login -ScopeKind List -List $list) {
                $removedAny = $true
            }
            if (Remove-ExternalUserAssignmentAtScope -Login $login -ScopeKind Web) {
                $removedAny = $true
            }
        }
        default {
            Write-Warning "Skipped unsupported external-user target type '$itemType' for $login"
            return
        }
    }

    if (-not $removedAny) {
        Write-Verbose "No role assignment found to remove for $login at $itemType '$itemPath'."
    }
}

function Invoke-RemoveSharingLinkTarget {
    param([object] $Target)

    Connect-PnPSiteIfNeeded -TargetSiteUrl $Target.SiteUrl

    $rel = [string]$Target.ItemPath
    $shareId = [string]$Target.ShareId
    $itemType = [string]$Target.ItemType

    if ($itemType -eq 'Folder') {
        if ($shareId) {
            Remove-PnPFolderSharingLink -Folder $rel -Identity $shareId -Force -ErrorAction Stop
        }
        else {
            Remove-PnPFolderSharingLink -Folder $rel -Force -ErrorAction Stop
        }
    }
    else {
        if ($shareId) {
            Remove-PnPFileSharingLink -FileUrl $rel -Identity $shareId -Force -ErrorAction Stop
        }
        else {
            Remove-PnPFileSharingLink -FileUrl $rel -Force -ErrorAction Stop
        }
    }
}

function Invoke-RemediationPlan {
    param(
        [string]   $ActionName,
        [object[]] $Plan,
        [string[]] $RoleColumns = @(),
        [switch]   $UseMessageBox
    )

    if (@($Plan).Count -lt 1) {
        Write-Host "No targets for $ActionName." -ForegroundColor DarkYellow
        return @{ Ok = 0; Fail = 0; Skip = 0 }
    }

    Show-RemediationPreview -Title "$ActionName - preview" -Plan $Plan

    if (-not (Confirm-RemediationAction -Prompt "Apply $ActionName to $(@($Plan).Count) target(s)?" `
            -UseMessageBox:$UseMessageBox)) {
        Write-Host 'Skipped by user.' -ForegroundColor Yellow
        return @{ Ok = 0; Fail = 0; Skip = @($Plan).Count }
    }

    $ok = 0
    $fail = 0
    $skip = 0
    $i = 0
    $total = @($Plan).Count

    foreach ($target in $Plan) {
        $i++
        Write-Progress -Activity $ActionName -Status ("{0}/{1} {2}" -f $i, $total, $target.Detail) `
            -PercentComplete ([int](($i / $total) * 100))

        try {
            switch ($target.Action) {
                'ResetUniquePermissions' { Invoke-ResetUniquePermissionTarget -Target $target }
                'RemoveUserPermission'   { Invoke-RemoveExternalUserTarget -Target $target -RoleColumns $RoleColumns }
                'RemoveGroupMember'      { Invoke-RemoveExternalUserTarget -Target $target -RoleColumns $RoleColumns }
                'RemoveSharingLink'      { Invoke-RemoveSharingLinkTarget -Target $target }
                default                  { throw "Unknown action $($target.Action)" }
            }
            $ok++
            Write-Host ("  OK  {0}" -f $target.Detail) -ForegroundColor DarkGreen
        }
        catch {
            $msg = [string]$_.Exception.Message
            if ($msg -like 'SKIP:*') {
                $skip++
                Write-Host ("  SKIP {0}" -f ($msg -replace '^SKIP:\s*', '')) -ForegroundColor DarkYellow
            }
            else {
                $fail++
                Write-Warning ("  FAIL {0} - {1}" -f $target.Detail, $msg)
            }
        }
    }

    Write-Progress -Activity $ActionName -Completed
    return @{ Ok = $ok; Fail = $fail; Skip = $skip }
}

function Show-MainMenu {
    param(
        [int] $ExternalCount,
        [int] $LinkCount,
        [int] $UniqueCount,
        [bool] $DoExternal,
        [bool] $DoLinks,
        [bool] $DoUnique
    )

    Write-MenuTitle 'Permissions Matrix Remediation'
    Write-Host "Report : $ReportPath"
    Write-Host ''
    Write-Host 'Select actions (toggle with number, then R to run, P to preview all selected, Q to quit):'
    Write-Host ("  [1] Remove external users      {0,8:N0} targets  [{1}]" -f $ExternalCount, $(if ($DoExternal) { 'X' } else { ' ' }))
    Write-Host ("  [2] Remove sharing links       {0,8:N0} targets  [{1}]" -f $LinkCount, $(if ($DoLinks) { 'X' } else { ' ' }))
    Write-Host ("  [3] Reset unique permissions {0,8:N0} targets  [{1}]" -f $UniqueCount, $(if ($DoUnique) { 'X' } else { ' ' }))
    Write-Host '  [P] Preview selected actions'
    Write-Host '  [R] Run selected actions'
    Write-Host '  [Q] Quit'
}


# --- Remediation picker UI ---
function Get-RemediationSiteScopeRows {
    param(
        [hashtable] $SiteMap,
        [object[]]  $Groups,
        [object[]]  $Sharing
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($kv in $SiteMap.GetEnumerator()) {
        $url = Get-NormalizedSiteUrl ([string]$kv.Value)
        if ($url -and $seen.Add($url)) {
            $rows.Add([pscustomobject]@{ SiteName = [string]$kv.Key; SiteUrl = $url }) | Out-Null
        }
    }

    foreach ($g in @($Groups)) {
        $url = Get-NormalizedSiteUrl ([string]$g.'Site URL')
        if (-not $url) { continue }
        if (-not $seen.Add($url)) { continue }
        $name = [string]$g.'Site Name'
        if (-not $name) {
            $name = (Get-ServerRelativePath $url) -replace '^.*/', ''
        }
        $rows.Add([pscustomobject]@{ SiteName = $name; SiteUrl = $url }) | Out-Null
    }

    foreach ($s in @($Sharing)) {
        $url = Get-NormalizedSiteUrl ([string]$s.SiteUrl)
        if (-not $url) { continue }
        if (-not $seen.Add($url)) { continue }
        $webPath = Get-ServerRelativePath $url
        $name = ($webPath -split '/')[-1]
        if (-not $name) { $name = $url }
        $rows.Add([pscustomobject]@{ SiteName = $name; SiteUrl = $url }) | Out-Null
    }

    return @($rows | Sort-Object { [string]$_.SiteName })
}

function Build-SitePlanCountIndex {
    param(
        [object[]] $ExternalPlan,
        [object[]] $LinkPlan,
        [object[]] $UniquePlan
    )

    $ext = @{}
    $lnk = @{}
    $uni = @{}

    foreach ($e in @($ExternalPlan)) {
        $u = Get-NormalizedSiteUrl ([string]$e.SiteUrl)
        if (-not $u) { continue }
        if (-not $ext.ContainsKey($u)) { $ext[$u] = 0 }
        $ext[$u]++
    }
    foreach ($e in @($LinkPlan)) {
        $u = Get-NormalizedSiteUrl ([string]$e.SiteUrl)
        if (-not $u) { continue }
        if (-not $lnk.ContainsKey($u)) { $lnk[$u] = 0 }
        $lnk[$u]++
    }
    foreach ($e in @($UniquePlan)) {
        $u = Get-NormalizedSiteUrl ([string]$e.SiteUrl)
        if (-not $u) { continue }
        if (-not $uni.ContainsKey($u)) { $uni[$u] = 0 }
        $uni[$u]++
    }

    return @{
        External = $ext
        Links    = $lnk
        Unique   = $uni
    }
}

function Filter-RemediationPlanBySiteScope {
    param(
        [object[]] $Plan,
        [string[]] $SelectedSiteUrls
    )

    if (-not $Plan) { return @() }
    if (-not $SelectedSiteUrls -or @($SelectedSiteUrls).Count -lt 1) { return @() }

    $allowed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($u in $SelectedSiteUrls) {
        $norm = Get-NormalizedSiteUrl $u
        if ($norm) { [void]$allowed.Add($norm) }
    }

    return @($Plan | Where-Object {
        $norm = Get-NormalizedSiteUrl ([string]$_.SiteUrl)
        $allowed.Contains($norm)
    })
}

function Get-RemediationPlanItemKey {
    param([object] $Item)

    switch ([string]$Item.Action) {
        'RemoveGroupMember'      { return "group|$($Item.SiteUrl)|$($Item.Group)|$($Item.Login)" }
        'RemoveUserPermission'   { return "user|$($Item.SiteUrl)|$($Item.ItemType)|$($Item.ItemPath)|$($Item.Login)" }
        'RemoveSharingLink'      { return "link|$($Item.SiteUrl)|$($Item.ShareId)|$($Item.ItemPath)" }
        'ResetUniquePermissions' { return "unique|$($Item.SiteUrl)|$($Item.ItemPath)" }
        'UniquePrincipal'      { return "uniquechild|$($Item.SiteUrl)|$($Item.ItemPath)|$($Item.Login)" }
        default                  { return "other|$($Item.Action)|$($Item.SiteUrl)|$($Item.ItemPath)|$($Item.Login)" }
    }
}

function Get-RemediationPlansForReport {
    param(
        [string] $ReportPath,
        [string] $SiteUrl,
        [string] $ClientId
    )

    $data = Import-ReportWorkbook -Path $ReportPath
    if ($SiteUrl) { $SiteUrl = $SiteUrl.Trim() }
    if (-not $SiteUrl) {
        $SiteUrl = Get-ReportRootSiteUrl -SummaryUrl $data.SummaryUrl -Matrix $data.Matrix `
            -Sharing $data.Sharing -Groups $data.Groups
    }
    if (-not $SiteUrl) {
        throw 'Could not determine root site URL from the report. Ensure the Summary sheet, Group Members, or Sharing Links includes site URLs.'
    }

    $SiteUrl = $SiteUrl.Trim()
    Initialize-SiteMap -Matrix $data.Matrix -Sharing $data.Sharing -Groups $data.Groups `
        -SummaryUrl $data.SummaryUrl -DefaultSiteUrl $SiteUrl

    $externalPlan = Build-ExternalUserPlan -Matrix $data.Matrix -Groups $data.Groups `
        -RoleColumns $data.RoleColumns -FallbackSiteUrl $SiteUrl
    $linkPlan = Build-SharingLinkPlan -Sharing $data.Sharing
    $uniquePlan = Build-UniquePermissionResetPlan -Matrix $data.Matrix -Sharing $data.Sharing `
        -FallbackSiteUrl $SiteUrl

    $uniqueItemKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($u in @($uniquePlan)) {
        [void]$uniqueItemKeys.Add((Get-UniqueItemLookupKey $u))
    }

    $uniquePrincipalIndex = Build-UniqueItemPrincipalIndex -Matrix $data.Matrix `
        -FallbackSiteUrl $SiteUrl -RoleColumns $data.RoleColumns -UniqueItemKeys $uniqueItemKeys

    $externalByKey = @{}
    foreach ($e in @($externalPlan)) {
        $externalByKey[(Get-RemediationPlanItemKey $e)] = $e
    }

    $siteRows = Get-RemediationSiteScopeRows -SiteMap $State.SiteMap -Groups $data.Groups `
        -Sharing $data.Sharing
    $sitePlanCounts = Build-SitePlanCountIndex -ExternalPlan $externalPlan -LinkPlan $linkPlan `
        -UniquePlan $uniquePlan

    return [pscustomobject]@{
        Data                   = $data
        SiteUrl                = $SiteUrl
        ClientId               = $ClientId
        ExternalPlan           = $externalPlan
        LinkPlan               = $linkPlan
        UniquePlan             = $uniquePlan
        UniquePrincipalIndex   = $uniquePrincipalIndex
        ExternalPlanByKey      = $externalByKey
        SiteRows               = $siteRows
        SitePlanCounts         = $sitePlanCounts
    }
}

function Show-RemediationPicker {
    param(
        [string] $InitialReportPath = '',
        [string] $InitialClientId = '',
        [string] $InitialSiteUrl = ''
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $script:RemediationPickerCtx = $null
    $script:RemediationPickerSiteItems = [System.Collections.Generic.List[object]]::new()
    $script:RemediationPickerExternalItems = [System.Collections.Generic.List[object]]::new()
    $script:RemediationPickerLinkItems = [System.Collections.Generic.List[object]]::new()
    $script:RemediationPickerUniqueItems = [System.Collections.Generic.List[object]]::new()
    $script:RemediationPickerUncheckedKeys = @{
        External      = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        Links         = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        Unique        = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        UniqueChild   = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    }
    $script:RemediationPickerUniqueExpanded = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $script:RemediationPickerUniqueGroups = [System.Collections.Generic.List[object]]::new()
    $script:RemediationPickerUniquePopulated = $false
    $script:PickerSiteRowData = [System.Collections.Generic.List[object]]::new()
    $script:PickerExternalRowData = [System.Collections.Generic.List[object]]::new()
    $script:PickerLinksRowData = [System.Collections.Generic.List[object]]::new()
    $script:PickerTabState = @{
        Sites    = @{ SortCol = -1; SortAsc = $true; Filters = @{}; Headers = @('Site', 'Path', 'External', 'Links', 'Unique') }
        External = @{ SortCol = -1; SortAsc = $true; Filters = @{}; Headers = @('User / group', 'Path', 'Type', 'Action') }
        Links    = @{ SortCol = -1; SortAsc = $true; Filters = @{}; Headers = @('Link type', 'Path', 'Site') }
        Unique   = @{ SortCol = -1; SortAsc = $true; Filters = @{}; Headers = @('', 'Item / User', 'Path', 'Site / Type', 'External', 'Roles') }
    }

    function Set-RoundedButton {
        param($Btn, [int]$Radius = 7, $Bg, $Fg = [System.Drawing.Color]::White,
              [bool]$Border = $false, $BorderClr = $null, $HoverBg = $null)
        $Btn.FlatStyle = 'Flat'
        $Btn.FlatAppearance.BorderSize = 0
        $Btn.BackColor = $Bg
        $Btn.ForeColor = $Fg
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
            $g.SmoothingMode   = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $parentBg = if ($s.Parent) { $s.Parent.BackColor } else { [System.Drawing.Color]::White }
            $g.Clear($parentBg)
            $fillClr = if ($t.Hover) { $t.HoverBg } else { $t.Bg }
            $d    = $t.R * 2
            $rc   = New-Object System.Drawing.Rectangle(0, 0, ($s.Width - 1), ($s.Height - 1))
            $path = New-Object System.Drawing.Drawing2D.GraphicsPath
            $path.AddArc($rc.X, $rc.Y, $d, $d, 180, 90)
            $path.AddArc($rc.Right - $d, $rc.Y, $d, $d, 270, 90)
            $path.AddArc($rc.Right - $d, $rc.Bottom - $d, $d, $d, 0, 90)
            $path.AddArc($rc.X, $rc.Bottom - $d, $d, $d, 90, 90)
            $path.CloseAllFigures()
            $br = New-Object System.Drawing.SolidBrush($fillClr)
            $g.FillPath($br, $path); $br.Dispose()
            if ($t.Brd -and $t.BrdClr) {
                $pen = New-Object System.Drawing.Pen($t.BrdClr, 1)
                $g.DrawPath($pen, $path); $pen.Dispose()
            }
            $sf = New-Object System.Drawing.StringFormat
            $sf.Alignment = [System.Drawing.StringAlignment]::Center
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $tb = New-Object System.Drawing.SolidBrush($t.Fg)
            $displayText = $s.Text -replace '&&', '&'
            $g.DrawString($displayText, $s.Font, $tb,
                [System.Drawing.RectangleF]::new(0, 0, $s.Width, $s.Height), $sf)
            $tb.Dispose(); $sf.Dispose(); $path.Dispose()
        })
    }

    $clrGreen    = [System.Drawing.ColorTranslator]::FromHtml('#37ae1c')
    $clrSoft     = [System.Drawing.ColorTranslator]::FromHtml('#EEF0FE')
    $clrBg       = [System.Drawing.ColorTranslator]::FromHtml('#F3F4FB')
    $clrBorder   = [System.Drawing.ColorTranslator]::FromHtml('#E4E6F5')
    $clrTextPri  = [System.Drawing.ColorTranslator]::FromHtml('#1C1F4A')
    $clrTextMid  = [System.Drawing.ColorTranslator]::FromHtml('#5A5F8A')
    $clrTextSoft = [System.Drawing.ColorTranslator]::FromHtml('#9EA3C8')
    $clrCoral    = [System.Drawing.ColorTranslator]::FromHtml('#F04E65')

    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'SharePoint Permissions Matrix Remediation (picker 2025-06-02)'
    $form.StartPosition = 'CenterScreen'
    $form.Size          = New-Object System.Drawing.Size(1020, 820)
    $form.MinimumSize   = New-Object System.Drawing.Size(900, 680)
    $form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.BackColor     = $clrBg

    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = 'Top'
    $pnlHeader.Height = 54
    $pnlHeader.Add_Paint({
        param($s, $e)
        $r = $s.ClientRectangle
        $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $r,
            [System.Drawing.ColorTranslator]::FromHtml('#37ae1c'),
            [System.Drawing.ColorTranslator]::FromHtml('#3D4FD6'),
            [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
        $e.Graphics.FillRectangle($grad, $r)
        $grad.Dispose()
        $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
        $fTitle = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
        $fSub   = New-Object System.Drawing.Font('Segoe UI', 8)
        $bWhite = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $bSub   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(191, 255, 255, 255))
        $e.Graphics.DrawString('Permissions Matrix Remediation', $fTitle, $bWhite, 16.0, 8.0)
        $e.Graphics.DrawString('Select sites and individual targets, then preview or run', $fSub, $bSub, 17.0, 32.0)
        $fTitle.Dispose(); $fSub.Dispose(); $bWhite.Dispose(); $bSub.Dispose()
    })

    $pnlFooter = New-Object System.Windows.Forms.Panel
    $pnlFooter.Dock = 'Bottom'
    $pnlFooter.Height = 72
    $pnlFooter.BackColor = [System.Drawing.Color]::White

    $pnlFooterBorder = New-Object System.Windows.Forms.Panel
    $pnlFooterBorder.Dock = 'Top'
    $pnlFooterBorder.Height = 1
    $pnlFooterBorder.BackColor = $clrBorder

    $lblFooter = New-Object System.Windows.Forms.Label
    $lblFooter.Location = New-Object System.Drawing.Point(12, 8)
    $lblFooter.Size = New-Object System.Drawing.Size(520, 56)
    $lblFooter.ForeColor = $clrTextMid
    $lblFooter.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblFooter.Text = 'Load a report to see targets. Check sites and items to include, then preview or run.'

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'
    $btnPreview.Size = New-Object System.Drawing.Size(88, 32)
    $btnPreview.Location = New-Object System.Drawing.Point(650, 20)
    $btnPreview.Anchor = 'Top, Right'
    $btnPreview.Enabled = $false

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(744, 20)
    $btnCancel.Anchor = 'Top, Right'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = ("Run remediation {0}" -f [char]0x2192)
    $btnRun.Size = New-Object System.Drawing.Size(130, 32)
    $btnRun.Location = New-Object System.Drawing.Point(840, 20)
    $btnRun.Anchor = 'Top, Right'
    $btnRun.Enabled = $false

    $pnlFooter.Controls.AddRange(@($pnlFooterBorder, $lblFooter, $btnPreview, $btnCancel, $btnRun))

    $pnlBody = New-Object System.Windows.Forms.Panel
    $pnlBody.Dock = 'Fill'
    $pnlBody.BackColor = $clrBg
    $pnlBody.Padding = New-Object System.Windows.Forms.Padding(10, 0, 10, 12)

    $cardBorderPaint = {
        param($s, $e)
        $pen = New-Object System.Drawing.Pen($clrBorder, 1)
        $w = $s.ClientSize.Width
        $h = $s.ClientSize.Height
        $e.Graphics.DrawLine($pen, 0, $h - 1, $w, $h - 1)
        $pen.Dispose()
    }

    # --- Report card ---
    $cardReport = New-Object System.Windows.Forms.Panel
    $cardReport.BackColor = [System.Drawing.Color]::White
    $cardReport.Dock = 'Top'
    $cardReport.Height = 188
    $cardReport.Add_Paint($cardBorderPaint)

    $lblRepSec = New-Object System.Windows.Forms.Label
    $lblRepSec.Text = 'REPORT'
    $lblRepSec.Font = New-Object System.Drawing.Font('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)
    $lblRepSec.ForeColor = $clrTextSoft
    $lblRepSec.Location = New-Object System.Drawing.Point(14, 10)
    $lblRepSec.AutoSize = $true

    $lblReport = New-Object System.Windows.Forms.Label
    $lblReport.Text = 'Permissions Matrix workbook'
    $lblReport.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $lblReport.ForeColor = $clrTextMid
    $lblReport.Location = New-Object System.Drawing.Point(14, 28)
    $lblReport.AutoSize = $true

    function New-FieldRow {
        param([int]$TopY, [int]$Height = 26)
        $p = New-Object System.Windows.Forms.Panel
        $p.Location = New-Object System.Drawing.Point(14, $TopY)
        $p.Size = New-Object System.Drawing.Size(400, $Height)
        $p.Anchor = 'Top, Left, Right'
        $p.BackColor = [System.Drawing.Color]::White
        return $p
    }

    $rowReport = New-FieldRow 50
    $txtReport = New-Object System.Windows.Forms.TextBox
    $txtReport.Dock = 'Fill'
    $txtReport.BorderStyle = 'FixedSingle'
    $txtReport.Text = $InitialReportPath

    $pnlReportGap = New-Object System.Windows.Forms.Panel
    $pnlReportGap.Dock = 'Right'
    $pnlReportGap.Width = 12
    $pnlReportGap.BackColor = [System.Drawing.Color]::White

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'
    $btnBrowse.Dock = 'Right'
    $btnBrowse.Width = 96

    $rowReport.Controls.AddRange(@($txtReport, $pnlReportGap, $btnBrowse))

    $lblClient = New-Object System.Windows.Forms.Label
    $lblClient.Text = 'Client ID'
    $lblClient.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $lblClient.ForeColor = $clrTextMid
    $lblClient.Location = New-Object System.Drawing.Point(14, 86)
    $lblClient.AutoSize = $true

    $rowClient = New-FieldRow 106
    $txtClient = New-Object System.Windows.Forms.TextBox
    $txtClient.Dock = 'Fill'
    $txtClient.BorderStyle = 'FixedSingle'
    $txtClient.Text = $InitialClientId
    $rowClient.Controls.Add($txtClient)

    $pnlLoadRow = New-Object System.Windows.Forms.Panel
    $pnlLoadRow.Location = New-Object System.Drawing.Point(14, 142)
    $pnlLoadRow.Size = New-Object System.Drawing.Size(400, 36)
    $pnlLoadRow.Anchor = 'Top, Left, Right'
    $pnlLoadRow.BackColor = [System.Drawing.Color]::White

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load report'
    $btnLoad.Size = New-Object System.Drawing.Size(120, 32)
    $btnLoad.Location = New-Object System.Drawing.Point(0, 0)
    $btnLoad.Anchor = 'Top, Left'

    $lblLoadStatus = New-Object System.Windows.Forms.Label
    $lblLoadStatus.Location = New-Object System.Drawing.Point(132, 8)
    $lblLoadStatus.Size = New-Object System.Drawing.Size(260, 22)
    $lblLoadStatus.Anchor = 'Top, Left, Right'
    $lblLoadStatus.ForeColor = $clrTextMid
    $lblLoadStatus.Text = ''

    $pnlLoadRow.Controls.AddRange(@($btnLoad, $lblLoadStatus))

    $cardReport.Controls.AddRange(@(
        $lblRepSec, $lblReport, $rowReport, $lblClient, $rowClient, $pnlLoadRow))

    function New-ScopeTabPage {
        param(
            [string] $Title,
            [int[]]  $ColumnWidths,
            [string[]] $ColumnHeaders
        )

        $page = New-Object System.Windows.Forms.TabPage
        $page.Text = $Title
        $page.BackColor = [System.Drawing.Color]::White
        $page.Padding = New-Object System.Windows.Forms.Padding(0)

        $hdr = New-Object System.Windows.Forms.Panel
        $hdr.Dock = 'Top'
        $hdr.Height = 36
        $hdr.BackColor = [System.Drawing.Color]::White

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Dock = 'Left'
        $lbl.Padding = New-Object System.Windows.Forms.Padding(14, 10, 0, 0)
        $lbl.AutoSize = $true
        $lbl.ForeColor = $clrTextMid
        $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
        $lbl.Text = '0 selected'

        $btnAllTab = New-Object System.Windows.Forms.Button
        $btnAllTab.Text = 'All'
        $btnAllTab.Size = New-Object System.Drawing.Size(52, 24)
        $btnAllTab.Location = New-Object System.Drawing.Point(780, 6)
        $btnAllTab.Anchor = 'Top, Right'
        $btnAllTab.Enabled = $false

        $btnNoneTab = New-Object System.Windows.Forms.Button
        $btnNoneTab.Text = 'None'
        $btnNoneTab.Size = New-Object System.Drawing.Size(52, 24)
        $btnNoneTab.Location = New-Object System.Drawing.Point(838, 6)
        $btnNoneTab.Anchor = 'Top, Right'
        $btnNoneTab.Enabled = $false

        $hdr.Controls.AddRange(@($lbl, $btnAllTab, $btnNoneTab))

        $listHost = New-Object System.Windows.Forms.Panel
        $listHost.Dock = 'Fill'
        $listHost.Padding = New-Object System.Windows.Forms.Padding(14, 0, 14, 10)
        $listHost.BackColor = [System.Drawing.Color]::White

        $lvTab = New-Object System.Windows.Forms.ListView
        $lvTab.Dock = 'Fill'
        $lvTab.View = 'Details'
        $lvTab.CheckBoxes = $true
        $lvTab.FullRowSelect = $true
        $lvTab.GridLines = $false
        $lvTab.BorderStyle = 'FixedSingle'
        $lvTab.HeaderStyle = 'Clickable'
        $lvTab.Font = New-Object System.Drawing.Font('Segoe UI', 9)

        for ($ci = 0; $ci -lt $ColumnHeaders.Count; $ci++) {
            $w = if ($ci -lt $ColumnWidths.Count) { $ColumnWidths[$ci] } else { 120 }
            [void]$lvTab.Columns.Add($ColumnHeaders[$ci], $w)
        }

        $listHost.Controls.Add($lvTab)
        $page.Controls.Add($listHost)
        $page.Controls.Add($hdr)

        return [pscustomobject]@{
            Page     = $page
            ListView = $lvTab
            Label    = $lbl
            BtnAll   = $btnAllTab
            BtnNone  = $btnNoneTab
            HdrPanel = $hdr
        }
    }

    $cardScope = New-Object System.Windows.Forms.Panel
    $cardScope.BackColor = [System.Drawing.Color]::White
    $cardScope.Dock = 'Fill'
    $cardScope.Add_Paint($cardBorderPaint)

    $tabScope = New-Object System.Windows.Forms.TabControl
    $tabScope.Dock = 'Fill'
    $tabScope.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $tabScope.Enabled = $false

    $tabSites = New-ScopeTabPage -Title 'Sites' -ColumnHeaders @('Site', 'Path', 'External', 'Links', 'Unique') `
        -ColumnWidths @(200, 340, 72, 56, 56)
    $tabExternal = New-ScopeTabPage -Title 'External users' -ColumnHeaders @('User / group', 'Path', 'Type', 'Action') `
        -ColumnWidths @(160, 280, 72, 320)
    $tabLinks = New-ScopeTabPage -Title 'Sharing links' -ColumnHeaders @('Link type', 'Path', 'Site') `
        -ColumnWidths @(100, 380, 200)
    $tabUnique = New-ScopeTabPage -Title 'Unique permissions' `
        -ColumnHeaders @('', 'Item / User', 'Path', 'Site / Type', 'External', 'Roles') `
        -ColumnWidths @(40, 200, 300, 100, 72, 120)

    $lv = $tabSites.ListView
    $btnAll = $tabSites.BtnAll
    $btnNone = $tabSites.BtnNone
    $lvExternal = $tabExternal.ListView
    $lvLinks = $tabLinks.ListView
    $lvUnique = $tabUnique.ListView
    $lvUnique.ShowItemToolTips = $true
    $btnAllExt = $tabExternal.BtnAll
    $btnNoneExt = $tabExternal.BtnNone
    $btnAllLinks = $tabLinks.BtnAll
    $btnNoneLinks = $tabLinks.BtnNone
    $btnAllUnique = $tabUnique.BtnAll
    $btnNoneUnique = $tabUnique.BtnNone
    $lblTabSites = $tabSites.Label
    $lblTabExternal = $tabExternal.Label
    $lblTabLinks = $tabLinks.Label
    $lblTabUnique = $tabUnique.Label

    $tabScope.TabPages.AddRange(@(
        $tabSites.Page, $tabExternal.Page, $tabLinks.Page, $tabUnique.Page))
    $cardScope.Controls.Add($tabScope)

    $sp1 = New-Object System.Windows.Forms.Panel; $sp1.Dock='Top'; $sp1.Height=6; $sp1.BackColor=$clrBg
    $sp2 = New-Object System.Windows.Forms.Panel; $sp2.Dock='Top'; $sp2.Height=6; $sp2.BackColor=$clrBg

    $pnlBody.Controls.Add($cardScope)
    $pnlBody.Controls.Add($sp2)
    $pnlBody.Controls.Add($cardReport)
    $pnlBody.Controls.Add($sp1)

    $form.Controls.Add($pnlBody)
    $form.Controls.Add($pnlFooter)
    $form.Controls.Add($pnlHeader)
    $form.CancelButton = $btnCancel

    function Update-RemediationPickerLayout {
        $pad = 14

        $cw = $cardReport.ClientSize.Width
        if ($cw -gt 0) {
            $inner = $cw - $pad - $pad
            $rowReport.Width = $inner
            $rowClient.Width = $inner
            $pnlLoadRow.Width = $inner
            $lblLoadStatus.Left = $btnLoad.Width + 12
            $lblLoadStatus.Width = [Math]::Max(80, $pnlLoadRow.ClientSize.Width - $lblLoadStatus.Left)
        }

        foreach ($hdrPanel in @($tabSites.HdrPanel, $tabExternal.HdrPanel, $tabLinks.HdrPanel, $tabUnique.HdrPanel)) {
            $hw = $hdrPanel.ClientSize.Width
            if ($hw -le 0) { continue }
            $btns = @($hdrPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] })
            if ($btns.Count -ge 2) {
                $noneBtn = $btns | Where-Object { $_.Text -eq 'None' } | Select-Object -First 1
                $allBtn = $btns | Where-Object { $_.Text -eq 'All' } | Select-Object -First 1
                if ($noneBtn -and $allBtn) {
                    $noneBtn.Left = $hw - $pad - $noneBtn.Width
                    $allBtn.Left = $hw - $pad - $noneBtn.Width - 4 - $allBtn.Width
                }
            }
        }

        $fw = $pnlFooter.ClientSize.Width
        if ($fw -gt 0) {
            $btnRun.Left = $fw - $pad - $btnRun.Width
            $btnCancel.Left = $fw - $pad - $btnRun.Width - 8 - $btnCancel.Width
            $btnPreview.Left = $fw - $pad - $btnRun.Width - 8 - $btnCancel.Width - 8 - $btnPreview.Width
            $lblFooter.Width = [Math]::Max(120, $btnPreview.Left - 16)
        }
    }

    $form.Add_Shown({ Update-RemediationPickerLayout })
    $form.Add_Resize({ Update-RemediationPickerLayout })

    function Get-SelectedSiteUrls {
        $urls = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $script:RemediationPickerSiteItems) {
            if ($item.ListItem -and $item.ListItem.Checked) {
                [void]$urls.Add([string]$item.SiteUrl)
            }
        }
        return @($urls)
    }

    function Get-TrackerChecked {
        param($Tracker)
        if ($Tracker.ListItem) { return $Tracker.ListItem.Checked }
        if ($Tracker.Node) { return $Tracker.Node.Checked }
        return $false
    }

    function Format-PrincipalDisplayName {
        param([string] $Login)

        $name = [string]$Login
        if ($name -match '(?i)\|membership\|(.+)$') { return $Matches[1] }
        if ($name -match '(?i)^i:0#\.f\|membership\|(.+)$') { return $Matches[1] }
        if ($name -match '(?i)^i:0#\.w\|(.+)$') { return $Matches[1] }
        return $name
    }

    function Get-ShortPathLabel {
        param([string] $Path, [int] $MaxLen = 72)

        $path = ($Path -replace '^/+', '/').Trim()
        if ($path.Length -le $MaxLen) { return $path }
        return '...' + $path.Substring($path.Length - ($MaxLen - 3))
    }

    function Save-TargetCheckState {
        param([string] $Kind, [object[]] $Trackers)

        $unchecked = $script:RemediationPickerUncheckedKeys[$Kind]
        foreach ($t in @($Trackers)) {
            if (Get-TrackerChecked $t) {
                [void]$unchecked.Remove([string]$t.PlanKey)
            }
            else {
                [void]$unchecked.Add([string]$t.PlanKey)
            }
        }
    }

    function Save-UniqueTreeCheckState {
        $parentUnchecked = $script:RemediationPickerUncheckedKeys.Unique
        $childUnchecked = $script:RemediationPickerUncheckedKeys.UniqueChild
        foreach ($t in @($script:RemediationPickerUniqueItems)) {
            $unchecked = if ($t.IsParent) { $parentUnchecked } else { $childUnchecked }
            if (Get-TrackerChecked $t) {
                [void]$unchecked.Remove([string]$t.PlanKey)
            }
            else {
                [void]$unchecked.Add([string]$t.PlanKey)
            }
        }
    }

    function Get-SelectedPlanItems {
        param([object[]] $Trackers)

        $selected = [System.Collections.Generic.List[object]]::new()
        foreach ($t in @($Trackers)) {
            if (Get-TrackerChecked $t) {
                $selected.Add($t.PlanItem) | Out-Null
            }
        }
        return @($selected)
    }

    function Find-ExternalPlanForPrincipal {
        param(
            [object] $Principal,
            [object] $Ctx
        )

        $siteUrl = Get-NormalizedSiteUrl ([string]$Principal.SiteUrl)
        $itemPath = Normalize-ServerRelativePath ([string]$Principal.ItemPath)
        $login = [string]$Principal.Login

        foreach ($e in @($Ctx.ExternalPlan)) {
            if ((Get-NormalizedSiteUrl $e.SiteUrl) -ne $siteUrl) { continue }
            if ((Normalize-ServerRelativePath $e.ItemPath) -ne $itemPath) { continue }
            if ([string]$e.Login -ne $login) { continue }
            return $e
        }
        return $null
    }

    function Get-SelectedUniqueRemediation {
        $uniqueResets = [System.Collections.Generic.List[object]]::new()
        $extraExternal = [System.Collections.Generic.List[object]]::new()
        $seenExt = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($t in @($script:RemediationPickerUniqueItems)) {
            if (-not (Get-TrackerChecked $t)) { continue }
            if ($t.IsParent) {
                $uniqueResets.Add($t.PlanItem) | Out-Null
                continue
            }
            if ($t.ExternalPlanItem) {
                $ek = Get-RemediationPlanItemKey $t.ExternalPlanItem
                if ($seenExt.Add($ek)) {
                    $extraExternal.Add($t.ExternalPlanItem) | Out-Null
                }
            }
        }

        return @{
            UniqueResets   = @($uniqueResets)
            ExtraExternal  = @($extraExternal)
        }
    }

    function Set-ListViewChecks {
        param($ListView, [bool] $Checked, [object[]] $Trackers)

        $ListView.BeginUpdate()
        foreach ($t in @($Trackers)) {
            $t.ListItem.Checked = $Checked
        }
        $ListView.EndUpdate()
    }

    function Update-SelectionSummary {
        if (-not $script:RemediationPickerCtx) {
            $lblFooter.Text = 'Load a report to see targets. Check sites and items to include, then preview or run.'
            return
        }

        $siteN = @((Get-SelectedSiteUrls)).Count
        $extSel = @(Get-SelectedPlanItems -Trackers $script:RemediationPickerExternalItems).Count
        $lnkSel = @(Get-SelectedPlanItems -Trackers $script:RemediationPickerLinkItems).Count
        $uniSel = Get-SelectedUniqueRemediation
        $uniResetSel = @($uniSel.UniqueResets).Count
        $uniUserSel = @($uniSel.ExtraExternal).Count
        $extShown = @($script:RemediationPickerExternalItems).Count
        $lnkShown = @($script:RemediationPickerLinkItems).Count
        $uniShown = @($script:RemediationPickerUniqueItems | Where-Object { $_.IsParent }).Count

        $siteShown = @($script:RemediationPickerSiteItems).Count
        $lblTabSites.Text = ('{0:N0} of {1:N0} site(s) selected' -f $siteN, $siteShown)
        $lblTabExternal.Text = ('{0:N0} of {1:N0} selected' -f $extSel, $extShown)
        $lblTabLinks.Text = ('{0:N0} of {1:N0} selected' -f $lnkSel, $lnkShown)
        $lblTabUnique.Text = ('{0:N0} reset(s), {1:N0} user(s) selected' -f $uniResetSel, $uniUserSel)

        $lblFooter.Text = ('{0} site(s) | External {1:N0} | Links {2:N0} | Unique {3:N0} + {4:N0} users' -f `
            $siteN, $extSel, $lnkSel, $uniResetSel, $uniUserSel)
    }

    function Get-SitePlanCountsForUrl {
        param(
            [string] $SiteUrl,
            [hashtable] $CountIndex
        )

        $norm = Get-NormalizedSiteUrl $SiteUrl
        $ext = 0
        $lnk = 0
        $uni = 0
        if ($CountIndex) {
            if ($CountIndex.External.ContainsKey($norm)) { $ext = [int]$CountIndex.External[$norm] }
            if ($CountIndex.Links.ContainsKey($norm)) { $lnk = [int]$CountIndex.Links[$norm] }
            if ($CountIndex.Unique.ContainsKey($norm)) { $uni = [int]$CountIndex.Unique[$norm] }
        }
        return @{ External = $ext; Links = $lnk; Unique = $uni }
    }

    function Get-PickerTabState {
        param([string] $TabKey)
        return $script:PickerTabState[$TabKey]
    }

    function Update-PickerColumnHeaders {
        param(
            [System.Windows.Forms.ListView] $ListView,
            [string] $TabKey
        )

        $tabState = Get-PickerTabState $TabKey
        if (-not $tabState) { return }
        $hdrs = @($tabState['Headers'])
        $activeSort = if ($tabState.ContainsKey('SortCol')) { [int]$tabState['SortCol'] } else { -1 }
        $sortAsc = if ($tabState.ContainsKey('SortAsc')) { [bool]$tabState['SortAsc'] } else { $true }
        for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
            $base = if ($i -lt $hdrs.Count) { [string]$hdrs[$i] } else { '' }
            $suffix = ''
            if ($tabState['Filters'].ContainsKey($i) -and -not [string]::IsNullOrWhiteSpace([string]$tabState['Filters'][$i])) {
                $suffix += ' *'
            }
            if ($activeSort -eq $i) {
                $suffix += if ($sortAsc) { ' ^' } else { ' v' }
            }
            $ListView.Columns[$i].Text = $base + $suffix
        }
    }

    function Test-PickerCellsMatchFilters {
        param(
            [string[]] $Cells,
            [hashtable] $Filters
        )

        if (-not $Filters -or $Filters.Count -lt 1) { return $true }
        $cellArr = @($Cells)
        foreach ($colKey in $Filters.Keys) {
            $filter = [string]$Filters[$colKey]
            if ([string]::IsNullOrWhiteSpace($filter)) { continue }
            $idx = [int]$colKey
            $text = if ($idx -ge 0 -and $idx -lt $cellArr.Count) { [string]$cellArr[$idx] } else { '' }
            if ($text -notlike "*$filter*") { return $false }
        }
        return $true
    }

    function Sort-PickerRows {
        param(
            [object[]] $Rows,
            [int] $SortColumn,
            [bool] $Ascending
        )

        $pickerSortColumnIndex = [int]$SortColumn
        if ($pickerSortColumnIndex -lt 0 -or @($Rows).Count -lt 2) { return @($Rows) }

        $pickerSortAscending = [bool]$Ascending
        $pickerRowList = [System.Collections.Generic.List[object]]::new()
        foreach ($pickerRow in @($Rows)) {
            if ($null -ne $pickerRow) { [void]$pickerRowList.Add($pickerRow) }
        }
        if ($pickerRowList.Count -lt 2) { return @($pickerRowList) }

        # .NET sort with captured locals — avoids PS7 $_ / SortCol name binding in event handlers.
        $pickerSortComparer = {
            param($pickerRowA, $pickerRowB)

            $pickerCellsA = @($pickerRowA.Cells)
            $pickerCellsB = @($pickerRowB.Cells)
            $pickerKeyA = ''
            $pickerKeyB = ''
            if ($pickerSortColumnIndex -ge 0 -and $pickerSortColumnIndex -lt $pickerCellsA.Count) {
                $pickerKeyA = [string]$pickerCellsA[$pickerSortColumnIndex]
            }
            if ($pickerSortColumnIndex -ge 0 -and $pickerSortColumnIndex -lt $pickerCellsB.Count) {
                $pickerKeyB = [string]$pickerCellsB[$pickerSortColumnIndex]
            }
            $pickerCmp = [StringComparer]::OrdinalIgnoreCase.Compare($pickerKeyA, $pickerKeyB)
            if (-not $pickerSortAscending) { $pickerCmp = -$pickerCmp }
            return $pickerCmp
        }

        $pickerRowList.Sort([System.Comparison[object]]$pickerSortComparer)
        return @($pickerRowList)
    }

    function New-ListViewItemFromCells {
        param(
            [string[]] $Cells,
            [bool] $Checked = $true
        )

        $cellArr = @($Cells)
        if ($cellArr.Count -lt 1) { $cellArr = @('') }
        $item = New-Object System.Windows.Forms.ListViewItem ([string]$cellArr[0])
        for ($i = 1; $i -lt $cellArr.Count; $i++) {
            [void]$item.SubItems.Add([string]$cellArr[$i])
        }
        $item.Checked = $Checked
        return $item
    }

    function Get-HeaderColumnFromPoint {
        param(
            [System.Windows.Forms.ListView] $ListView,
            [int] $X
        )

        $scrollX = 0
        if ($ListView.Items.Count -gt 0) {
            $posX = $ListView.Items[0].Position.X
            if ($posX -lt 0) { $scrollX = -$posX }
        }
        return (Get-ListViewColumnIndexAtX -ListView $ListView -X ($X + $scrollX))
    }

    function Sync-PickerSiteRowChecksFromUi {
        $byUrl = @{}
        foreach ($t in @($script:RemediationPickerSiteItems)) {
            $byUrl[[string]$t.SiteUrl] = $t.ListItem.Checked
        }
        foreach ($row in $script:PickerSiteRowData) {
            $url = [string]$row.SiteUrl
            if ($byUrl.ContainsKey($url)) { $row.Checked = [bool]$byUrl[$url] }
        }
    }

    function Build-PickerSiteRowData {
        param($Ctx)

        $savedChecked = @{}
        foreach ($row in $script:PickerSiteRowData) {
            $savedChecked[[string]$row.SiteUrl] = [bool]$row.Checked
        }
        foreach ($t in @($script:RemediationPickerSiteItems)) {
            $savedChecked[[string]$t.SiteUrl] = $t.ListItem.Checked
        }

        $script:PickerSiteRowData.Clear()
        $countIndex = $Ctx.SitePlanCounts
        foreach ($site in @($Ctx.SiteRows)) {
            $url = [string]$site.SiteUrl
            $counts = Get-SitePlanCountsForUrl -SiteUrl $url -CountIndex $countIndex
            $checked = $true
            if ($savedChecked.ContainsKey($url)) { $checked = [bool]$savedChecked[$url] }

            $script:PickerSiteRowData.Add([pscustomobject]@{
                SiteUrl  = $url
                SiteName = [string]$site.SiteName
                Checked  = $checked
                Cells    = @(
                    [string]$site.SiteName
                    ($url -replace '^https?://[^/]+', '')
                    ('{0:N0}' -f $counts.External)
                    ('{0:N0}' -f $counts.Links)
                    ('{0:N0}' -f $counts.Unique)
                )
            }) | Out-Null
        }
    }

    function Render-SitesListView {
        if (-not $script:RemediationPickerCtx) { return }

        $tabState = Get-PickerTabState 'Sites'
        $tabFilters = $tabState['Filters']
        $rows = @($script:PickerSiteRowData | Where-Object {
            Test-PickerCellsMatchFilters -Cells $_.Cells -Filters $tabFilters
        })
        $pickSortIdx = if ($tabState.ContainsKey('SortCol')) { [int]$tabState['SortCol'] } else { -1 }
        $pickSortAsc = if ($tabState.ContainsKey('SortAsc')) { [bool]$tabState['SortAsc'] } else { $true }
        $rows = Sort-PickerRows -Rows $rows -SortColumn $pickSortIdx -Ascending $pickSortAsc

        $script:SuppressScopeRefresh = $true
        $lv.BeginUpdate()
        $lv.Items.Clear()
        $script:RemediationPickerSiteItems.Clear()
        try {
            foreach ($row in $rows) {
                $item = New-ListViewItemFromCells -Cells $row.Cells -Checked $row.Checked
                [void]$lv.Items.Add($item)
                $script:RemediationPickerSiteItems.Add([pscustomobject]@{
                    ListItem = $item
                    SiteName = [string]$row.SiteName
                    SiteUrl  = [string]$row.SiteUrl
                }) | Out-Null
            }
        }
        finally {
            $lv.EndUpdate()
            $script:SuppressScopeRefresh = $false
        }
        Update-PickerColumnHeaders -ListView $lv -TabKey 'Sites'
    }

    function Populate-SiteList {
        param($Ctx)
        Build-PickerSiteRowData -Ctx $Ctx
        Render-SitesListView
    }

    function Rebuild-PickerPlanRowData {
        if (-not $script:RemediationPickerCtx) { return }

        if (-not $script:SuppressTargetSave) {
            Save-TargetCheckState -Kind 'External' -Trackers $script:RemediationPickerExternalItems
            Save-TargetCheckState -Kind 'Links' -Trackers $script:RemediationPickerLinkItems
        }

        $urls = @(Get-SelectedSiteUrls)
        $ctx = $script:RemediationPickerCtx
        $extPlan = @(Filter-RemediationPlanBySiteScope -Plan $ctx.ExternalPlan -SelectedSiteUrls $urls)
        $lnkPlan = @(Filter-RemediationPlanBySiteScope -Plan $ctx.LinkPlan -SelectedSiteUrls $urls)
        $extUnchecked = $script:RemediationPickerUncheckedKeys.External
        $lnkUnchecked = $script:RemediationPickerUncheckedKeys.Links

        $script:PickerExternalRowData.Clear()
        foreach ($p in $extPlan) {
            $key = Get-RemediationPlanItemKey $p
            $name = if ($p.Group) { "$($p.Login) ($($p.Group))" } else { [string]$p.Login }
            $path = ($p.ItemPath -replace '^/+', '/')
            $script:PickerExternalRowData.Add([pscustomobject]@{
                PlanItem = $p
                PlanKey  = $key
                Checked  = -not $extUnchecked.Contains($key)
                Cells    = @($name, $path, [string]$p.ItemType, [string]$p.Detail)
            }) | Out-Null
        }

        $script:PickerLinksRowData.Clear()
        foreach ($p in $lnkPlan) {
            $key = Get-RemediationPlanItemKey $p
            $path = ($p.ItemPath -replace '^/+', '/')
            $script:PickerLinksRowData.Add([pscustomobject]@{
                PlanItem = $p
                PlanKey  = $key
                Checked  = -not $lnkUnchecked.Contains($key)
                Cells    = @([string]$p.LinkType, $path, ($p.SiteUrl -replace '^https?://[^/]+', ''))
            }) | Out-Null
        }
    }

    function Render-PlanListView {
        param(
            [System.Windows.Forms.ListView] $ListView,
            [System.Collections.Generic.List[object]] $Tracker,
            [string] $TabKey,
            [System.Collections.Generic.List[object]] $RowData
        )

        $tabState = Get-PickerTabState $TabKey
        $tabFilters = $tabState['Filters']
        $rows = @($RowData | Where-Object {
            Test-PickerCellsMatchFilters -Cells $_.Cells -Filters $tabFilters
        })
        $pickSortIdx = if ($tabState.ContainsKey('SortCol')) { [int]$tabState['SortCol'] } else { -1 }
        $pickSortAsc = if ($tabState.ContainsKey('SortAsc')) { [bool]$tabState['SortAsc'] } else { $true }
        $rows = Sort-PickerRows -Rows $rows -SortColumn $pickSortIdx -Ascending $pickSortAsc

        $script:SuppressTargetSave = $true
        $ListView.BeginUpdate()
        $ListView.Items.Clear()
        $Tracker.Clear()
        $unchecked = $script:RemediationPickerUncheckedKeys[$TabKey]
        try {
            foreach ($row in $rows) {
                $item = New-ListViewItemFromCells -Cells $row.Cells `
                    -Checked (-not $unchecked.Contains([string]$row.PlanKey))
                [void]$ListView.Items.Add($item)
                $Tracker.Add([pscustomobject]@{
                    ListItem = $item
                    PlanItem = $row.PlanItem
                    PlanKey  = $row.PlanKey
                }) | Out-Null
            }
        }
        finally {
            $ListView.EndUpdate()
            $script:SuppressTargetSave = $false
        }
        Update-PickerColumnHeaders -ListView $ListView -TabKey $TabKey
    }

    function Render-ExternalListView {
        Render-PlanListView -ListView $lvExternal -Tracker $script:RemediationPickerExternalItems `
            -TabKey 'External' -RowData $script:PickerExternalRowData
    }

    function Render-LinksListView {
        Render-PlanListView -ListView $lvLinks -Tracker $script:RemediationPickerLinkItems `
            -TabKey 'Links' -RowData $script:PickerLinksRowData
    }

    function Ensure-UniqueTabPopulated {
        if ($script:RemediationPickerUniquePopulated) { return }
        if (-not $script:RemediationPickerCtx) { return }

        $urls = @(Get-SelectedSiteUrls)
        $uni = @(Filter-RemediationPlanBySiteScope -Plan $script:RemediationPickerCtx.UniquePlan `
            -SelectedSiteUrls $urls)
        Populate-UniqueList -Plan $uni
        $script:RemediationPickerUniquePopulated = $true
        Update-SelectionSummary
    }

    function Populate-AllTargetLists {
        param([switch] $IncludeUnique)

        if (-not $script:RemediationPickerCtx) { return }

        $urls = @(Get-SelectedSiteUrls)
        $ctx = $script:RemediationPickerCtx

        Rebuild-PickerPlanRowData
        Render-ExternalListView
        Render-LinksListView

        if ($IncludeUnique -or $script:RemediationPickerUniquePopulated) {
            $uni = @(Filter-RemediationPlanBySiteScope -Plan $ctx.UniquePlan -SelectedSiteUrls $urls)
            Populate-UniqueList -Plan $uni
            $script:RemediationPickerUniquePopulated = $true
        }
        else {
            $lblTabUnique.Text = 'Open this tab to load unique permissions'
        }

        Update-SelectionSummary
    }

    function Get-ListViewColumnIndexAtX {
        param(
            [System.Windows.Forms.ListView] $ListView,
            [int] $X
        )

        $colX = 0
        for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
            $w = $ListView.Columns[$i].Width
            if ($X -ge $colX -and $X -lt ($colX + $w)) { return $i }
            $colX += $w
        }
        return -1
    }

    function Test-UniqueExpandClick {
        param(
            [System.Windows.Forms.ListView] $ListView,
            [System.Windows.Forms.ListViewItem] $Item,
            [int] $X
        )

        if ($ListView.Columns.Count -lt 1) { return $false }
        if ((Get-ListViewColumnIndexAtX -ListView $ListView -X $X) -ne 0) { return $false }

        $entire = $Item.GetBounds([System.Windows.Forms.ItemBoundsPortion]::Entire)
        # Leave checkbox clicks for ItemChecked only.
        if ($X -lt ($entire.Left + 18)) { return $false }
        return $true
    }

    function Test-UniqueGroupMatchesColumnFilters {
        param([object] $Group)

        $ColumnFilters = (Get-PickerTabState 'Unique').Filters
        if (-not $ColumnFilters -or $ColumnFilters.Count -lt 1) { return $true }

        $p = $Group.ParentRow
        $pathLabel = Get-ShortPathLabel (Normalize-ServerRelativePath ([string]$p.ItemPath))
        $siteLabel = ($p.SiteUrl -replace '^https?://[^/]+', '')

        foreach ($colKey in $ColumnFilters.Keys) {
            $filter = [string]$ColumnFilters[$colKey]
            if ([string]::IsNullOrWhiteSpace($filter)) { continue }
            $cell = switch ([int]$colKey) {
                1 { [string]$p.ItemType }
                2 { $pathLabel }
                3 { $siteLabel }
                4 { [string]$Group.ExternalLabel }
                5 {
                    $roles = @($Group.ChildRows | ForEach-Object {
                        $r = [string]$_.Principal.Roles
                        if (-not $r) { $r = [string]$_.Principal.GivenThrough }
                        $r
                    }) -join ' '
                    $roles
                }
                default { '' }
            }
            if ($cell -notlike "*$filter*") { return $false }
        }
        return $true
    }

    function Show-ColumnFilterDialog {
        param(
            [string] $ColumnName,
            [string] $CurrentFilter
        )

        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = "Filter — $ColumnName"
        $dlg.Size = New-Object System.Drawing.Size(380, 150)
        $dlg.FormBorderStyle = 'FixedDialog'
        $dlg.StartPosition = 'CenterParent'
        $dlg.MaximizeBox = $false
        $dlg.MinimizeBox = $false
        $dlg.ShowInTaskbar = $false
        $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = 'Contains (case-insensitive):'
        $lbl.Location = New-Object System.Drawing.Point(14, 14)
        $lbl.AutoSize = $true

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(14, 36)
        $txt.Size = New-Object System.Drawing.Size(340, 23)
        $txt.Text = $CurrentFilter

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = 'Apply'
        $btnOk.DialogResult = 'OK'
        $btnOk.Location = New-Object System.Drawing.Point(198, 72)
        $btnOk.Size = New-Object System.Drawing.Size(75, 28)

        $btnClear = New-Object System.Windows.Forms.Button
        $btnClear.Text = 'Clear'
        $btnClear.Location = New-Object System.Drawing.Point(118, 72)
        $btnClear.Size = New-Object System.Drawing.Size(75, 28)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = 'Cancel'
        $btnCancel.DialogResult = 'Cancel'
        $btnCancel.Location = New-Object System.Drawing.Point(278, 72)
        $btnCancel.Size = New-Object System.Drawing.Size(75, 28)

        $dlg.Controls.AddRange(@($lbl, $txt, $btnOk, $btnClear, $btnCancel))
        $dlg.AcceptButton = $btnOk
        $dlg.CancelButton = $btnCancel

        $cleared = $false
        $btnClear.Add_Click({
            $txt.Text = ''
            $cleared = $true
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        })

        $result = $dlg.ShowDialog($form)
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        return [pscustomobject]@{ Filter = $txt.Text.Trim(); Cleared = $cleared }
    }

    function Invoke-PickerColumnFilter {
        param(
            [string] $TabKey,
            [int] $ColumnIndex,
            [scriptblock] $OnApply
        )

        $colIdx = [int]$ColumnIndex
        if ($colIdx -lt 0) { return }

        $tabState = Get-PickerTabState $TabKey
        $hdrs = @($tabState['Headers'])
        $base = if ($colIdx -lt $hdrs.Count) { [string]$hdrs[$colIdx] } else { "Column $colIdx" }
        $current = if ($tabState['Filters'].ContainsKey($colIdx)) { [string]$tabState['Filters'][$colIdx] } else { '' }

        $dlgResult = Show-ColumnFilterDialog -ColumnName $base -CurrentFilter $current
        if ($null -eq $dlgResult) { return }

        if ([string]::IsNullOrWhiteSpace($dlgResult.Filter)) {
            if ($tabState['Filters'].ContainsKey($colIdx)) { $tabState['Filters'].Remove($colIdx) | Out-Null }
        }
        else {
            $tabState['Filters'][$colIdx] = $dlgResult.Filter.Trim()
        }

        & $OnApply
    }

    function Refresh-RemediationPickerLists {
        if (-not $script:RemediationPickerCtx) { return }
        Sync-PickerSiteRowChecksFromUi
        Render-SitesListView
        Rebuild-PickerPlanRowData
        Render-ExternalListView
        Render-LinksListView
        if ($script:RemediationPickerUniquePopulated) {
            $urls = @(Get-SelectedSiteUrls)
            $uni = @(Filter-RemediationPlanBySiteScope -Plan $script:RemediationPickerCtx.UniquePlan `
                -SelectedSiteUrls $urls)
            Populate-UniqueList -Plan $uni
        }
    }

    function Register-PickerListView {
        param(
            [System.Windows.Forms.ListView] $ListView,
            [string] $TabKey,
            [int[]] $SkipColumns = @(),
            [scriptblock] $OnRender,
            [scriptblock] $OnFilter
        )

        if (-not $OnFilter) { $OnFilter = $OnRender }

        $lvCtrl = $ListView
        $tabKeyLocal = $TabKey
        $skipLocal = @($SkipColumns)
        $renderLocal = $OnRender
        $filterLocal = $OnFilter

        $lvCtrl.Add_ColumnClick({
            param($sender, $eventArgs)
            if ($null -eq $eventArgs.Column) { return }
            $colIndex = [int]$eventArgs.Column.Index
            if ($skipLocal -contains $colIndex) { return }

            $tabState = $script:PickerTabState[$tabKeyLocal]
            if (-not $tabState) { return }
            $prevSort = if ($tabState.ContainsKey('SortCol')) { [int]$tabState['SortCol'] } else { -1 }
            if ($prevSort -eq $colIndex) {
                $tabState['SortAsc'] = -not [bool]$tabState['SortAsc']
            }
            else {
                $tabState['SortCol'] = $colIndex
                $tabState['SortAsc'] = $true
            }

            $null = $_
            & $renderLocal
        })

        $lvCtrl.Add_MouseUp({
            param($sender, $eventArgs)
            if ($eventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
            $hit = $sender.HitTest($eventArgs.Location)
            if (-not ($hit.Location -band [System.Windows.Forms.ListViewHitTestLocations]::Header)) { return }

            $col = Get-HeaderColumnFromPoint -ListView $sender -X $eventArgs.X
            if ($col -lt 0 -or ($skipLocal -contains $col)) { return }

            Invoke-PickerColumnFilter -TabKey $tabKeyLocal -ColumnIndex $col -OnApply $filterLocal
        })
    }

    function Sort-UniquePermissionGroups {
        param(
            [int] $ColumnIndex,
            [bool] $Ascending
        )

        $pickerUniqueSortCol = [int]$ColumnIndex
        if ($pickerUniqueSortCol -lt 1) { return }

        $tabState = Get-PickerTabState 'Unique'
        $tabState['SortCol'] = $pickerUniqueSortCol
        $tabState['SortAsc'] = [bool]$Ascending

        $col = $pickerUniqueSortCol
        $wrapped = foreach ($g in @($script:RemediationPickerUniqueGroups)) {
            $p = $g.ParentRow
            $sortKey = switch ($col) {
                1 { [string]$p.ItemType }
                2 { [string]$p.ItemPath }
                3 { ($p.SiteUrl -replace '^https?://[^/]+', '') }
                4 { [string]$g.ExternalLabel }
                5 {
                    (@($g.ChildRows | ForEach-Object {
                        $r = [string]$_.Principal.Roles
                        if (-not $r) { $r = [string]$_.Principal.GivenThrough }
                        $r
                    }) -join ' ')
                }
                default { [string]$p.ItemType }
            }
            [pscustomobject]@{ Group = $g; SortKey = $sortKey }
        }
        if ($Ascending) {
            $sorted = @($wrapped | Sort-Object -Property SortKey | ForEach-Object { $_.Group })
        }
        else {
            $sorted = @($wrapped | Sort-Object -Property SortKey -Descending | ForEach-Object { $_.Group })
        }

        $script:RemediationPickerUniqueGroups.Clear()
        foreach ($g in $sorted) {
            $script:RemediationPickerUniqueGroups.Add($g) | Out-Null
        }
        Rebuild-UniqueListFromGroups
        Update-PickerColumnHeaders -ListView $lvUnique -TabKey 'Unique'
    }

    function Refresh-UniqueTabDataAndView {
        if (-not $script:RemediationPickerCtx) { return }
        $urls = @(Get-SelectedSiteUrls)
        $uni = @(Filter-RemediationPlanBySiteScope -Plan $script:RemediationPickerCtx.UniquePlan `
            -SelectedSiteUrls $urls)
        Populate-UniqueList -Plan $uni
    }

    function Apply-UniqueTabSortFromState {
        if (-not $script:RemediationPickerUniquePopulated) { return }
        $tabState = Get-PickerTabState 'Unique'
        $pickSortIdx = if ($tabState.ContainsKey('SortCol')) { [int]$tabState['SortCol'] } else { -1 }
        if ($pickSortIdx -ge 1 -and $script:RemediationPickerUniqueGroups.Count -gt 0) {
            $pickSortAsc = if ($tabState.ContainsKey('SortAsc')) { [bool]$tabState['SortAsc'] } else { $true }
            Sort-UniquePermissionGroups -ColumnIndex $pickSortIdx -Ascending $pickSortAsc
        }
        else {
            Refresh-UniqueTabDataAndView
        }
    }

    function Toggle-UniqueParentExpanded {
        param([string] $ParentKey)

        if ($script:RemediationPickerUniqueExpanded.Contains($ParentKey)) {
            [void]$script:RemediationPickerUniqueExpanded.Remove($ParentKey)
        }
        else {
            [void]$script:RemediationPickerUniqueExpanded.Add($ParentKey)
        }

        Rebuild-UniqueListFromGroups
    }

    function Rebuild-UniqueListFromGroups {
        $parentUnchecked = $script:RemediationPickerUncheckedKeys.Unique
        $childUnchecked = $script:RemediationPickerUncheckedKeys.UniqueChild

        $script:SuppressTargetSave = $true
        $lvUnique.BeginUpdate()
        $lvUnique.Items.Clear()
        $script:RemediationPickerUniqueItems.Clear()

        try {
        foreach ($group in @($script:RemediationPickerUniqueGroups)) {
            $planItem = $group.ParentRow
            $parentKey = Get-RemediationPlanItemKey $planItem
            $principals = @($group.Principals)
            $hasChildren = (@($principals).Count -gt 0)
            $expanded = $script:RemediationPickerUniqueExpanded.Contains($parentKey)
            $glyph = if (-not $hasChildren) { '' } elseif ($expanded) { [char]0x25BC } else { [char]0x25B6 }

            $pathFull = Normalize-ServerRelativePath ([string]$planItem.ItemPath)
            $pathLabel = Get-ShortPathLabel $pathFull
            $siteLabel = ($planItem.SiteUrl -replace '^https?://[^/]+', '')

            $parentItem = New-Object System.Windows.Forms.ListViewItem $glyph
            $parentItem.Checked = -not $parentUnchecked.Contains($parentKey)
            $parentItem.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            [void]$parentItem.SubItems.Add([string]$planItem.ItemType)
            [void]$parentItem.SubItems.Add($pathLabel)
            [void]$parentItem.SubItems.Add($siteLabel)
            [void]$parentItem.SubItems.Add([string]$group.ExternalLabel)
            [void]$parentItem.SubItems.Add('')
            $parentItem.ToolTipText = "Reset unique permissions on this item`n$pathFull"

            [void]$lvUnique.Items.Add($parentItem)
            $script:RemediationPickerUniqueItems.Add([pscustomobject]@{
                ListItem         = $parentItem
                IsParent         = $true
                HasChildren      = $hasChildren
                PlanItem         = $planItem
                PlanKey          = $parentKey
                ExternalPlanItem = $null
                ParentKey        = $null
            }) | Out-Null

            if (-not $expanded) { continue }

            foreach ($child in @($group.ChildRows)) {
                $pr = $child.Principal
                $childKey = $child.PlanKey
                $extPlan = $child.ExternalPlanItem
                $displayName = Format-PrincipalDisplayName ([string]$pr.Login)
                $extChild = if ($pr.External) { 'Yes' } else { '' }
                $roles = [string]$pr.Roles
                if (-not $roles) { $roles = [string]$pr.GivenThrough }

                $childItem = New-Object System.Windows.Forms.ListViewItem ''
                $childItem.Checked = -not $childUnchecked.Contains($childKey)
                $childItem.ForeColor = $clrTextMid
                [void]$childItem.SubItems.Add("    $displayName")
                [void]$childItem.SubItems.Add('')
                [void]$childItem.SubItems.Add([string]$pr.PrincipalType)
                [void]$childItem.SubItems.Add($extChild)
                [void]$childItem.SubItems.Add($roles)
                $childItem.ToolTipText = ([string]$pr.Login)

                [void]$lvUnique.Items.Add($childItem)
                $script:RemediationPickerUniqueItems.Add([pscustomobject]@{
                    ListItem         = $childItem
                    IsParent         = $false
                    PlanItem         = $pr
                    PlanKey          = $childKey
                    ExternalPlanItem = $extPlan
                    ParentKey        = $parentKey
                }) | Out-Null
            }
        }
        }
        finally {
            $lvUnique.EndUpdate()
            $script:SuppressTargetSave = $false
        }
        Update-PickerColumnHeaders -ListView $lvUnique -TabKey 'Unique'
        Update-SelectionSummary
    }

    function Populate-UniqueList {
        param([object[]] $Plan)

        if (-not $script:SuppressTargetSave) {
            Save-UniqueTreeCheckState
        }

        $ctx = $script:RemediationPickerCtx
        $index = $ctx.UniquePrincipalIndex
        $script:RemediationPickerUniqueGroups.Clear()

        foreach ($planItem in @($Plan)) {
            $lookupKey = Get-UniqueItemLookupKey $planItem
            $principals = @()
            if ($index -and $index.ContainsKey($lookupKey)) {
                $principals = @($index[$lookupKey])
            }

            $extCount = @($principals | Where-Object { $_.External }).Count
            $extLabel = if ($extCount -gt 0) { "Yes ($extCount)" } else { '-' }

            $childRows = [System.Collections.Generic.List[object]]::new()
            foreach ($pr in $principals) {
                $childKey = Get-RemediationPlanItemKey ([pscustomobject]@{
                    Action   = 'UniquePrincipal'
                    SiteUrl  = $pr.SiteUrl
                    ItemPath = $pr.ItemPath
                    Login    = $pr.Login
                })
                $extPlan = Find-ExternalPlanForPrincipal -Principal $pr -Ctx $ctx
                if (-not $extPlan -and $ctx.ExternalPlanByKey) {
                    $tryKey = Get-RemediationPlanItemKey ([pscustomobject]@{
                        Action   = 'RemoveUserPermission'
                        SiteUrl  = $pr.SiteUrl
                        ItemType = 'List item'
                        ItemPath = $pr.ItemPath
                        Login    = $pr.Login
                    })
                    if ($ctx.ExternalPlanByKey.ContainsKey($tryKey)) {
                        $extPlan = $ctx.ExternalPlanByKey[$tryKey]
                    }
                }
                $childRows.Add([pscustomobject]@{
                    Principal        = $pr
                    PlanKey          = $childKey
                    ExternalPlanItem = $extPlan
                }) | Out-Null
            }

            $group = [pscustomobject]@{
                ParentRow     = $planItem
                ExternalLabel = $extLabel
                Principals    = $principals
                ChildRows     = @($childRows)
            }
            if (-not (Test-UniqueGroupMatchesColumnFilters -Group $group)) { continue }

            $script:RemediationPickerUniqueGroups.Add($group) | Out-Null
        }

        $tabState = Get-PickerTabState 'Unique'
        $pickSortIdx = if ($tabState.ContainsKey('SortCol')) { [int]$tabState['SortCol'] } else { -1 }
        if ($pickSortIdx -ge 1) {
            $pickSortAsc = if ($tabState.ContainsKey('SortAsc')) { [bool]$tabState['SortAsc'] } else { $true }
            Sort-UniquePermissionGroups -ColumnIndex $pickSortIdx -Ascending $pickSortAsc
        }
        else {
            Rebuild-UniqueListFromGroups
            Update-PickerColumnHeaders -ListView $lvUnique -TabKey 'Unique'
        }
    }

    function Refresh-ScopeAfterSiteChange {
        Sync-PickerSiteRowChecksFromUi
        Save-TargetCheckState -Kind 'External' -Trackers $script:RemediationPickerExternalItems
        Save-TargetCheckState -Kind 'Links' -Trackers $script:RemediationPickerLinkItems
        if ($script:RemediationPickerUniquePopulated) {
            Save-UniqueTreeCheckState
        }
        Populate-AllTargetLists -IncludeUnique:$script:RemediationPickerUniquePopulated
    }

    $tabScope.Add_SelectedIndexChanged({
        if (-not $tabScope.SelectedTab) { return }
        if ($tabScope.SelectedTab -ne $tabUnique.Page) { return }
        if ($script:RemediationPickerUniquePopulated) { return }
        if (-not $script:RemediationPickerCtx) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        try {
            Ensure-UniqueTabPopulated
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $script:SuppressScopeRefresh = $false
    $script:SuppressTargetSave = $false

    $lv.Add_ItemChecked({
        if ($script:SuppressScopeRefresh) { return }
        Refresh-ScopeAfterSiteChange
    })
    $lvExternal.Add_ItemChecked({
        if ($script:SuppressTargetSave) { return }
        Save-TargetCheckState -Kind 'External' -Trackers $script:RemediationPickerExternalItems
        Update-SelectionSummary
    })
    $lvLinks.Add_ItemChecked({
        if ($script:SuppressTargetSave) { return }
        Save-TargetCheckState -Kind 'Links' -Trackers $script:RemediationPickerLinkItems
        Update-SelectionSummary
    })
    $lvUnique.Add_ItemChecked({
        if ($script:SuppressTargetSave) { return }
        Save-UniqueTreeCheckState
        Update-SelectionSummary
    })

    $lvUnique.Add_MouseClick({
        param($s, $e)
        if ($e.Button -ne 'Left') { return }
        $hit = $lvUnique.HitTest($e.Location)
        if (-not $hit.Item) { return }
        if (-not (Test-UniqueExpandClick -ListView $lvUnique -Item $hit.Item -X $e.X)) { return }

        $tracker = @($script:RemediationPickerUniqueItems | Where-Object { $_.ListItem -eq $hit.Item } | Select-Object -First 1)
        if (-not $tracker -or -not $tracker.IsParent -or -not $tracker.HasChildren) { return }

        Toggle-UniqueParentExpanded -ParentKey $tracker.PlanKey
    })

    $lvUnique.Add_MouseDoubleClick({
        param($s, $e)
        if ($e.Button -ne 'Left') { return }
        $hit = $lvUnique.HitTest($e.Location)
        if (-not $hit.Item) { return }

        $tracker = @($script:RemediationPickerUniqueItems | Where-Object { $_.ListItem -eq $hit.Item } | Select-Object -First 1)
        if (-not $tracker -or -not $tracker.IsParent -or -not $tracker.HasChildren) { return }

        Toggle-UniqueParentExpanded -ParentKey $tracker.PlanKey
    })

    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Select Permissions Matrix report'
        $dlg.Filter = 'Excel workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*'
        if ($txtReport.Text -and (Test-Path -LiteralPath (Split-Path $txtReport.Text -Parent))) {
            $dlg.InitialDirectory = (Resolve-Path (Split-Path $txtReport.Text -Parent)).Path
        }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtReport.Text = $dlg.FileName
            if ([string]::IsNullOrWhiteSpace($txtClient.Text)) {
                $embedded = Get-ReportEmbeddedClientId -Path $dlg.FileName
                if ($embedded) { $txtClient.Text = $embedded }
            }
        }
    })

    $btnLoad.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtReport.Text)) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Select a report workbook.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        if (-not (Test-Path -LiteralPath $txtReport.Text)) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Report file not found.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnLoad.Enabled = $false
        $lblLoadStatus.Text = 'Loading report...'
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $overrideSite = if ($InitialSiteUrl) { $InitialSiteUrl.Trim() } else { '' }
            $clientForLoad = $txtClient.Text.Trim()
            $ctx = Get-RemediationPlansForReport -ReportPath $txtReport.Text.Trim() `
                -SiteUrl $overrideSite -ClientId $clientForLoad
            $script:RemediationPickerCtx = $ctx

            if (-not $clientForLoad -and $ctx.Data.EmbeddedClientId) {
                $clientForLoad = $ctx.Data.EmbeddedClientId
                $txtClient.Text = $clientForLoad
            }
            if (-not $clientForLoad) {
                $clientForLoad = Get-ReportEmbeddedClientId -Path $txtReport.Text.Trim()
                if ($clientForLoad) { $txtClient.Text = $clientForLoad }
            }
            $ctx.ClientId = $clientForLoad

            $script:RemediationPickerUncheckedKeys.External.Clear()
            $script:RemediationPickerUncheckedKeys.Links.Clear()
            $script:RemediationPickerUncheckedKeys.Unique.Clear()
            $script:RemediationPickerUncheckedKeys.UniqueChild.Clear()
            $script:RemediationPickerUniqueExpanded.Clear()
            $script:RemediationPickerUniquePopulated = $false
            $script:PickerSiteRowData.Clear()
            $script:PickerExternalRowData.Clear()
            $script:PickerLinksRowData.Clear()
            foreach ($k in @('Sites', 'External', 'Links', 'Unique')) {
                $script:PickerTabState[$k].SortCol = -1
                $script:PickerTabState[$k].SortAsc = $true
                $script:PickerTabState[$k].Filters = @{}
            }
            $lblLoadStatus.Text = 'Populating site list...'
            [System.Windows.Forms.Application]::DoEvents()
            Populate-SiteList -Ctx $ctx

            $lblLoadStatus.Text = 'Populating external users and sharing links...'
            [System.Windows.Forms.Application]::DoEvents()
            Populate-AllTargetLists

            $tabScope.Enabled = $true
            foreach ($b in @($btnAll, $btnNone, $btnAllExt, $btnNoneExt, $btnAllLinks, $btnNoneLinks, $btnAllUnique, $btnNoneUnique)) {
                $b.Enabled = $true
            }
            $btnPreview.Enabled = $true
            $btnRun.Enabled = $true

            $rootHint = ($ctx.SiteUrl -replace '^https?://[^/]+', '')
            $clientHint = if ($clientForLoad) { 'Client ID from report.' } `
                elseif ($ctx.Data.EmbeddedClientId) { 'Client ID found in report — enter above to run.' } `
                else { 'Enter Client ID to run remediation.' }
            $lblLoadStatus.Text = ('Loaded {0:N0} rows, {1} site(s). Root: {2}. {3}' -f `
                @($ctx.Data.Matrix).Count, @($ctx.SiteRows).Count, $rootHint, $clientHint)
            $lblLoadStatus.ForeColor = $clrGreen
        }
        catch {
            $script:RemediationPickerCtx = $null
            [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Load failed', 'OK', 'Error') | Out-Null
            $lblLoadStatus.Text = 'Load failed.'
            $lblLoadStatus.ForeColor = $clrCoral
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnLoad.Enabled = $true
        }
    })

    $btnAll.Add_Click({
        $script:SuppressScopeRefresh = $true
        try {
            foreach ($row in $script:RemediationPickerSiteItems) { $row.ListItem.Checked = $true }
        }
        finally { $script:SuppressScopeRefresh = $false }
        Refresh-ScopeAfterSiteChange
    })

    $btnNone.Add_Click({
        $script:SuppressScopeRefresh = $true
        try {
            foreach ($row in $script:RemediationPickerSiteItems) { $row.ListItem.Checked = $false }
        }
        finally { $script:SuppressScopeRefresh = $false }
        Refresh-ScopeAfterSiteChange
    })

    $btnAllExt.Add_Click({
        Set-ListViewChecks -ListView $lvExternal -Checked $true -Trackers $script:RemediationPickerExternalItems
        Update-SelectionSummary
    })
    $btnNoneExt.Add_Click({
        Set-ListViewChecks -ListView $lvExternal -Checked $false -Trackers $script:RemediationPickerExternalItems
        Save-TargetCheckState -Kind 'External' -Trackers $script:RemediationPickerExternalItems
        Update-SelectionSummary
    })
    $btnAllLinks.Add_Click({
        Set-ListViewChecks -ListView $lvLinks -Checked $true -Trackers $script:RemediationPickerLinkItems
        Update-SelectionSummary
    })
    $btnNoneLinks.Add_Click({
        Set-ListViewChecks -ListView $lvLinks -Checked $false -Trackers $script:RemediationPickerLinkItems
        Save-TargetCheckState -Kind 'Links' -Trackers $script:RemediationPickerLinkItems
        Update-SelectionSummary
    })
    function Set-UniqueTreeChecks {
        param([bool] $Checked)

        $script:SuppressTargetSave = $true
        foreach ($t in @($script:RemediationPickerUniqueItems)) {
            $t.ListItem.Checked = $Checked
        }
        $script:SuppressTargetSave = $false
        Save-UniqueTreeCheckState
    }

    $btnAllUnique.Add_Click({
        Set-UniqueTreeChecks -Checked $true
        Update-SelectionSummary
    })
    $btnNoneUnique.Add_Click({
        Set-UniqueTreeChecks -Checked $false
        Update-SelectionSummary
    })

    function Merge-SelectedExternalPlans {
        param(
            [object[]] $FromExternalTab,
            [object[]] $FromUniqueTree
        )

        $merged = [System.Collections.Generic.List[object]]::new()
        $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($p in @($FromExternalTab) + @($FromUniqueTree)) {
            $k = Get-RemediationPlanItemKey $p
            if ($seen.Add($k)) { $merged.Add($p) | Out-Null }
        }
        return @($merged)
    }

    function Test-HasSelectedTargets {
        $uni = Get-SelectedUniqueRemediation
        return @(
            (Get-SelectedPlanItems -Trackers $script:RemediationPickerExternalItems).Count,
            (Get-SelectedPlanItems -Trackers $script:RemediationPickerLinkItems).Count,
            @($uni.UniqueResets).Count,
            @($uni.ExtraExternal).Count
        ) | Where-Object { $_ -gt 0 }
    }

    $btnPreview.Add_Click({
        if (-not $script:RemediationPickerCtx) { return }
        Ensure-UniqueTabPopulated
        if ((@(Get-SelectedSiteUrls)).Count -lt 1) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Select at least one site.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $ext = Merge-SelectedExternalPlans -FromExternalTab @(Get-SelectedPlanItems -Trackers $script:RemediationPickerExternalItems) `
            -FromUniqueTree @(Get-SelectedUniqueRemediation).ExtraExternal
        $lnk = @(Get-SelectedPlanItems -Trackers $script:RemediationPickerLinkItems)
        $uni = @(Get-SelectedUniqueRemediation).UniqueResets
        $uniUsers = @(Get-SelectedUniqueRemediation).ExtraExternal

        if (-not (Test-HasSelectedTargets)) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Check at least one target in External users, Sharing links, or Unique permissions.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $lines = [System.Collections.Generic.List[string]]::new()
        if ($ext.Count -gt 0) {
            [void]$lines.Add("External users: $($ext.Count) selected")
            @($ext | Select-Object -First 12 | ForEach-Object { "  - $($_.Detail)" }) | ForEach-Object { [void]$lines.Add($_) }
            if ($ext.Count -gt 12) { [void]$lines.Add("  ... and $($ext.Count - 12) more") }
            [void]$lines.Add('')
        }
        if ($lnk.Count -gt 0) {
            [void]$lines.Add("Sharing links: $($lnk.Count) selected")
            @($lnk | Select-Object -First 12 | ForEach-Object { "  - $($_.Detail)" }) | ForEach-Object { [void]$lines.Add($_) }
            if ($lnk.Count -gt 12) { [void]$lines.Add("  ... and $($lnk.Count - 12) more") }
            [void]$lines.Add('')
        }
        if ($uni.Count -gt 0) {
            [void]$lines.Add("Unique resets: $($uni.Count) selected")
            @($uni | Select-Object -First 12 | ForEach-Object { "  - $($_.Detail)" }) | ForEach-Object { [void]$lines.Add($_) }
            if ($uni.Count -gt 12) { [void]$lines.Add("  ... and $($uni.Count - 12) more") }
            [void]$lines.Add('')
        }
        if ($uniUsers.Count -gt 0) {
            [void]$lines.Add("Users under unique items (external removal): $($uniUsers.Count) selected")
            @($uniUsers | Select-Object -First 12 | ForEach-Object { "  - $($_.Detail)" }) | ForEach-Object { [void]$lines.Add($_) }
            if ($uniUsers.Count -gt 12) { [void]$lines.Add("  ... and $($uniUsers.Count - 12) more") }
        }

        [System.Windows.Forms.MessageBox]::Show($form, ($lines -join "`n"), 'Remediation preview', 'OK', 'Information') | Out-Null
    })

    $script:RemediationPickerResult = $null

    $btnRun.Add_Click({
        if (-not $script:RemediationPickerCtx) { return }
        Ensure-UniqueTabPopulated
        if ([string]::IsNullOrWhiteSpace($txtClient.Text)) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Client ID is required (enter it or re-export the report with your Entra app Client ID).', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        $urls = @(Get-SelectedSiteUrls)
        if (@($urls).Count -lt 1) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Select at least one site.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }
        $uniPick = Get-SelectedUniqueRemediation
        $selExt = Merge-SelectedExternalPlans -FromExternalTab @(Get-SelectedPlanItems -Trackers $script:RemediationPickerExternalItems) `
            -FromUniqueTree @($uniPick.ExtraExternal)
        $selLnk = @(Get-SelectedPlanItems -Trackers $script:RemediationPickerLinkItems)
        $selUni = @($uniPick.UniqueResets)

        if (-not ($selExt.Count -or $selLnk.Count -or $selUni.Count)) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Check at least one target in External users, Sharing links, or Unique permissions.', 'Validation', 'OK', 'Warning') | Out-Null
            return
        }

        $script:RemediationPickerResult = [pscustomobject]@{
            ReportPath             = $txtReport.Text.Trim()
            ClientId               = $txtClient.Text.Trim()
            SiteUrl                = $script:RemediationPickerCtx.SiteUrl
            SiteUrls               = $urls
            SelectedExternalPlan   = $selExt
            SelectedLinkPlan       = $selLnk
            SelectedUniquePlan     = $selUni
            Context                = $script:RemediationPickerCtx
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    Set-RoundedButton $btnLoad    -Bg $clrGreen -Fg ([System.Drawing.Color]::White)
    Set-RoundedButton $btnRun     -Bg $clrGreen -Fg ([System.Drawing.Color]::White)
    Set-RoundedButton $btnPreview -Bg $clrSoft  -Fg $clrGreen -Border $true -BorderClr $clrBorder
    Set-RoundedButton $btnCancel  -Bg $clrBg -Fg $clrTextMid -Border $true -BorderClr $clrBorder
    Set-RoundedButton $btnBrowse  -Bg $clrSoft -Fg $clrGreen -Border $true -BorderClr $clrBorder
    Set-RoundedButton $btnAll        -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnNone       -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnAllExt     -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnNoneExt    -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnAllLinks   -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnNoneLinks  -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnAllUnique  -Bg $clrSoft -Fg $clrGreen
    Set-RoundedButton $btnNoneUnique -Bg $clrSoft -Fg $clrGreen

    Register-PickerListView -ListView $lv -TabKey 'Sites' -OnRender { Render-SitesListView }
    Register-PickerListView -ListView $lvExternal -TabKey 'External' -OnRender { Render-ExternalListView }
    Register-PickerListView -ListView $lvLinks -TabKey 'Links' -OnRender { Render-LinksListView }
    Register-PickerListView -ListView $lvUnique -TabKey 'Unique' -SkipColumns @(0) `
        -OnRender { Apply-UniqueTabSortFromState } -OnFilter { Refresh-UniqueTabDataAndView }

    [void]$form.ShowDialog()
    return $script:RemediationPickerResult
}

function Invoke-RemediationFromPickerResult {
    param(
        [object] $PickerResult,
        [switch] $UseMessageBox
    )

    if (-not $PickerResult) { return }

    $ctx = $PickerResult.Context
    $data = $ctx.Data
    $State.RoleColumns = $data.RoleColumns

    $externalPlan = @($PickerResult.SelectedExternalPlan)
    $linkPlan = @($PickerResult.SelectedLinkPlan)
    $uniquePlan = @($PickerResult.SelectedUniquePlan)

    $summary = @()
    if (@($externalPlan).Count -gt 0) {
        $summary += [pscustomobject]@{
            Action = 'External users'
            Result = Invoke-RemediationPlan -ActionName 'Remove external users' `
                -Plan $externalPlan -RoleColumns $data.RoleColumns -UseMessageBox:$UseMessageBox
        }
    }
    if (@($linkPlan).Count -gt 0) {
        $summary += [pscustomobject]@{
            Action = 'Sharing links'
            Result = Invoke-RemediationPlan -ActionName 'Remove sharing links' `
                -Plan $linkPlan -UseMessageBox:$UseMessageBox
        }
    }
    if (@($uniquePlan).Count -gt 0) {
        $summary += [pscustomobject]@{
            Action = 'Unique permissions'
            Result = Invoke-RemediationPlan -ActionName 'Reset unique permissions' `
                -Plan $uniquePlan -UseMessageBox:$UseMessageBox
        }
    }

    Write-MenuTitle 'Remediation summary'
    foreach ($s in $summary) {
        Write-Host ("{0}: {1} succeeded, {2} failed, {3} skipped" -f `
            $s.Action, $s.Result.Ok, $s.Result.Fail, $s.Result.Skip)
    }
}


# --- Main ---
if (-not $NoPicker) {
    $picked = Show-RemediationPicker -InitialReportPath $ReportPath `
        -InitialClientId $ClientId -InitialSiteUrl $SiteUrl
    if (-not $picked) {
        Write-Host 'Remediation cancelled.' -ForegroundColor Yellow
        return
    }

    $ClientId = $picked.ClientId

    Invoke-RemediationFromPickerResult -PickerResult $picked -UseMessageBox
    Write-Host 'Done.' -ForegroundColor Green
    return
}

# Console mode (-NoPicker)
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    Write-Host 'Select the Permissions Matrix Excel report...' -ForegroundColor Cyan
    $ReportPath = Select-ExcelReportPath -InitialDirectory $ScriptDir
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        Write-Host 'No report selected. Exiting.' -ForegroundColor Yellow
        return
    }
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    $ClientId = Get-ReportEmbeddedClientId -Path $ReportPath
}
if ([string]::IsNullOrWhiteSpace($ClientId)) {
    $ClientId = Read-Host 'Entra app Client ID'
}
if ([string]::IsNullOrWhiteSpace($ClientId)) { throw 'ClientId is required.' }
$ClientId = $ClientId.Trim()

$ReportPath = (Resolve-Path -LiteralPath $ReportPath).Path

$data = Import-ReportWorkbook -Path $ReportPath
$State.RoleColumns = $data.RoleColumns

if (-not $SiteUrl) {
    $SiteUrl = Get-ReportRootSiteUrl -SummaryUrl $data.SummaryUrl -Matrix $data.Matrix `
        -Sharing $data.Sharing -Groups $data.Groups
}
if (-not $SiteUrl) {
    $SiteUrl = Read-Host 'Root SharePoint site URL (not found in report; used to connect and resolve paths)'
}
$SiteUrl = $SiteUrl.Trim()
if (-not $SiteUrl) { throw 'SiteUrl is required.' }

Initialize-SiteMap -Matrix $data.Matrix -Sharing $data.Sharing -Groups $data.Groups `
    -SummaryUrl $data.SummaryUrl -DefaultSiteUrl $SiteUrl

$externalPlan = Build-ExternalUserPlan -Matrix $data.Matrix -Groups $data.Groups `
    -RoleColumns $data.RoleColumns -FallbackSiteUrl $SiteUrl
$linkPlan = Build-SharingLinkPlan -Sharing $data.Sharing
$uniquePlan = Build-UniquePermissionResetPlan -Matrix $data.Matrix -Sharing $data.Sharing `
    -FallbackSiteUrl $SiteUrl

$doExternal = $false
$doLinks = $false
$doUnique = $false

while ($true) {
    Show-MainMenu -ExternalCount @($externalPlan).Count -LinkCount @($linkPlan).Count `
        -UniqueCount @($uniquePlan).Count -DoExternal $doExternal -DoLinks $doLinks -DoUnique $doUnique

    $choice = (Read-Host 'Choice').Trim().ToUpperInvariant()
    switch ($choice) {
        '1' { $doExternal = -not $doExternal }
        '2' { $doLinks = -not $doLinks }
        '3' { $doUnique = -not $doUnique }
        'P' {
            if ($doExternal) { Show-RemediationPreview -Title 'External users' -Plan $externalPlan }
            if ($doLinks)    { Show-RemediationPreview -Title 'Sharing links' -Plan $linkPlan }
            if ($doUnique)   { Show-RemediationPreview -Title 'Unique permission resets' -Plan $uniquePlan }
            if (-not ($doExternal -or $doLinks -or $doUnique)) {
                Write-Host 'No actions selected. Toggle 1/2/3 first.' -ForegroundColor Yellow
            }
        }
        'R' {
            if (-not ($doExternal -or $doLinks -or $doUnique)) {
                Write-Host 'No actions selected. Toggle 1/2/3 first.' -ForegroundColor Yellow
                continue
            }

            $summary = @()
            if ($doExternal) {
                $summary += [pscustomobject]@{
                    Action = 'External users'
                    Result = Invoke-RemediationPlan -ActionName 'Remove external users' `
                        -Plan $externalPlan -RoleColumns $data.RoleColumns
                }
            }
            if ($doLinks) {
                $summary += [pscustomobject]@{
                    Action = 'Sharing links'
                    Result = Invoke-RemediationPlan -ActionName 'Remove sharing links' -Plan $linkPlan
                }
            }
            if ($doUnique) {
                $summary += [pscustomobject]@{
                    Action = 'Unique permissions'
                    Result = Invoke-RemediationPlan -ActionName 'Reset unique permissions' -Plan $uniquePlan
                }
            }

            Write-MenuTitle 'Remediation summary'
            foreach ($s in $summary) {
                Write-Host ("{0}: {1} succeeded, {2} failed, {3} skipped" -f `
                    $s.Action, $s.Result.Ok, $s.Result.Fail, $s.Result.Skip)
            }
        }
        'Q' { return }
        'QUIT' { return }
        default {
            Write-Host 'Unknown choice.' -ForegroundColor Yellow
        }
    }
}

Write-Host 'Done.' -ForegroundColor Green
