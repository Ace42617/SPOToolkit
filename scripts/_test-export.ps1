$ErrorActionPreference = 'Stop'
Import-Module ImportExcel
$path = Join-Path $env:TEMP 'sitebd-test.xlsx'
Remove-Item $path -Force -ErrorAction SilentlyContinue
$ov = @([pscustomobject]@{ Section='R'; Metric='M'; Value='V' })
$pkg = $ov | Export-Excel -Path $path -WorksheetName Summary -StartRow 8 -TableName SummaryOverview -PassThru
$ws = $pkg.Workbook.Worksheets['Summary']
Write-Host "Dim after overview: $($ws.Dimension.Address)"

$src = Get-Content (Join-Path $PSScriptRoot 'Export-SitePermissionsMatrix.ps1') -Raw
$fn = $src.IndexOf('function Get-ThemeColor')
$end = $src.IndexOf('# --- Main ---')
Invoke-Expression $src.Substring($fn, $end - $fn)

$sites = @(
    [pscustomobject]@{ 'Site Name'='IT'; 'Site Path'='/a'; 'Type'='Sub'; 'Lists/Libraries'=5; Items=10; Files=1; Folders=0; 'List items'=9; 'Unique items'=2; 'External users'=0; 'Site unique perms'='Yes' }
)
try {
    Write-SiteBreakdownTable -Package $pkg -Worksheet $ws -SiteRows $sites
    Write-Host "Dim after write: $($ws.Dimension.Address)"
    $t = $ws.Tables['SiteBreakdown']
    Write-Host "Table: $($t.Address.Address) H1=$($ws.Cells[$t.Address.Start.Row,1].Value)"
} catch {
    Write-Host "FAIL: $_" -ForegroundColor Red
}
Close-ExcelPackage $pkg
