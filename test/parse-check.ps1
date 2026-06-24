$errs = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot '..\scripts\RemediationPickerGui.ps1'),
    [ref]$null,
    [ref]$errs)
if ($errs) { $errs | ForEach-Object { $_.ToString() }; exit 1 }
Write-Host 'RemediationPickerGui.ps1 OK'
$errs2 = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot '..\scripts\Invoke-PermissionsMatrixRemediation.ps1'),
    [ref]$null,
    [ref]$errs2)
if ($errs2) { $errs2 | ForEach-Object { $_.ToString() }; exit 1 }
Write-Host 'Invoke-PermissionsMatrixRemediation.ps1 OK'
