Import-Module ImportExcel -ErrorAction Stop
$path = Join-Path $env:TEMP 'spotoolkit-proptest.xlsx'
Remove-Item $path -Force -ErrorAction SilentlyContinue

$pkg = Open-ExcelPackage -Path $path -Create
$wb = $pkg.Workbook
$props = $wb.Properties
Write-Host "Properties type: $($props.GetType().FullName)"
Write-Host "CustomDocumentProperties is null: $($null -eq $props.CustomDocumentProperties)"
$props.PSObject.Properties | ForEach-Object { Write-Host "  Prop member: $($_.Name)" }

# Hidden sheet approach
$ws = $pkg.Workbook.Worksheets.Add('_ToolkitMeta')
$ws.Cells['A1'].Value = 'SPOTK1:TESTVALUE'
$ws.Hidden = [OfficeOpenXml.eWorkSheetHidden]::VeryHidden

Export-Excel -ExcelPackage $pkg -WorksheetName 'Sheet1' -InputObject @([pscustomobject]@{X=1}) -TableName T1
Close-ExcelPackage $pkg

$pkg2 = Open-ExcelPackage -Path $path
$meta = $pkg2.Workbook.Worksheets['_ToolkitMeta']
if ($meta) {
    Write-Host "Meta A1: $($meta.Cells['A1'].Text)"
} else {
    Write-Host "Meta sheet missing"
}
Close-ExcelPackage $pkg2
