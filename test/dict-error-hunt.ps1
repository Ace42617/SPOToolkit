$ErrorActionPreference = 'Continue'
$d = [System.Collections.Generic.Dictionary[string,object[]]]::new()
$d['a'] = @([pscustomobject]@{ Name='x' })

function Test-KeepInheritedRow($Row, $RoleColumns) {
    $hasReal = $false
    foreach ($col in $RoleColumns) {
        if ($Row.$col -ne 'X') { continue }
        $hasReal = $true; break
    }
    return $hasReal
}

$RoleColumns = [System.Collections.Generic.List[string]]::new()
'Read','Full Control' | ForEach-Object { $RoleColumns.Add($_) | Out-Null }

$tests = @(
    { Test-KeepInheritedRow $d $RoleColumns },
    { $d.Inheritance -ne 'Inherited' },
    { foreach ($col in $RoleColumns) { $null = $d.$col } },
    { foreach ($col in $RoleColumns) { $row = @{}; $row[$col] = $d.$col } },
    { $sb = ''; foreach ($col in $RoleColumns) { if ($d.$col -eq 'X') { $sb += $col } } },
    { @($d | Where-Object { $_.Inheritance -ne 'Inherited' -or (Test-KeepInheritedRow $_ $RoleColumns) }) }
)

$i = 0
foreach ($t in $tests) {
    $i++
    try {
        $r = & $t
        Write-Host "Test $i OK: $r"
    }
    catch {
        Write-Host "Test $i FAIL: $($_.Exception.Message)"
    }
}

# KeyValuePair as row
foreach ($kvp in $d) {
    try {
        $r = Test-KeepInheritedRow $kvp $RoleColumns
        Write-Host "KVP Test-Keep: $r"
    }
    catch {
        Write-Host "KVP FAIL: $($_.Exception.Message)"
    }
}

# Simulate Get-RowKey on dict
function Get-RowKey($Row, $RoleColumns) {
    foreach ($col in $RoleColumns) {
        if ($Row.$col -eq 'X') { return 'x' }
    }
    return ''
}
try { Get-RowKey $d $RoleColumns } catch { Write-Host "Get-RowKey dict: $($_.Exception.Message)" }

# object[] value indexing
$val = $d['a']
try {
    $x = $val[0]
    Write-Host "val[0] OK: $x"
}
catch {
    Write-Host "val[0]: $($_.Exception.Message)"
}

# Wrong: indexing dict with int for assignment
try {
    $x = $d[0]
    Write-Host "d[0] get: $x"
}
catch {
    Write-Host "d[0] get: $($_.Exception.Message)"
}

# PS 7 style - assignment to index
try {
    $d[0] = @('y')
    Write-Host 'd[0] set OK'
}
catch {
    Write-Host "d[0] set: $($_.Exception.Message)"
}

# foreach ($role in $Roles) when Roles is dict
$Roles = $d
foreach ($role in $Roles) {
    try {
        $row = [ordered]@{}
        $row[$role] = 'X'
    }
    catch {
        Write-Host "row[role] role=$($role.GetType().Name): $($_.Exception.Message)"
    }
}

# [string[]] cast of dict as _ResolvedRoleNames then foreach in New-MatrixRow
$roles = @([string[]]$d)
foreach ($role in $roles) {
    try {
        $row = [ordered]@{}
        $row[$role] = 'X'
    }
    catch {
        Write-Host "row role string: $($_.Exception.Message)"
    }
}
