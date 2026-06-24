$d = [System.Collections.Generic.Dictionary[string,object[]]]::new()
$d['web'] = @([pscustomobject]@{ Id=1 })
$nested = [System.Collections.Generic.Dictionary[string,object[]]]::new()
$d['list'] = @($nested)  # object[] containing dict - valid type in array

try {
    $v = $d['list']
    Write-Host "list value type: $($v.GetType().FullName)"
    $x = $v[0]
    Write-Host "v[0] type: $($x.GetType().FullName)"
}
catch {
    Write-Host "v[0]: $($_.Exception.Message)"
}

try {
    $nested[0]
}
catch {
    Write-Host "nested[0]: $($_.Exception.Message)"
}

# Wrong assignment: store dict directly as object[] element without wrapping - actually we did @($nested)
# Wrong: value slot IS dict not object[]
$d2 = [System.Collections.Generic.Dictionary[string,object[]]]::new()
# Can't assign dict to object[] slot without cast - but PS might allow:
try {
    $d2.Add('bad', $nested)  # Add expects object[] value
}
catch {
    Write-Host "Add dict as value: $($_.Exception.Message)"
}

try {
    $d2['bad'] = $nested  # assign dict where object[] expected
    $d2['bad'][0]
}
catch {
    Write-Host "bad[0]: $($_.Exception.Message)"
}

# KeyValuePair Value access when Value is dict stored wrong
$kvp = [System.Collections.Generic.KeyValuePair[string,object[]]]::new('k', @($nested))
try {
    $kvp.Value[0]
}
catch {
    Write-Host "kvp.Value[0]: $($_.Exception.Message)"
}

# Simulate Get-RoleDefinitionBindingList IEnumerable on Assignments dict when Test-IsRoleCacheDictionary fails
function Test-IsRoleCacheDictionary($Obj) {
    return ($null -ne $Obj -and $Obj -is [System.Collections.Generic.Dictionary[string, object[]]])
}

$Assignments = $d
Write-Host "Test-IsRoleCacheDictionary Assignments: $(Test-IsRoleCacheDictionary $Assignments)"

foreach ($item in $Assignments) {
    Write-Host "foreach item: $($item.GetType().FullName)"
    if ($item -is [System.Collections.Generic.KeyValuePair[string, object[]]]) {
        Write-Host "  KeyValuePair Value type: $($item.Value.GetType().FullName)"
        try {
            foreach ($inner in @($item.Value)) {
                Write-Host "  inner: $($inner.GetType().FullName)"
            }
        }
        catch {
            Write-Host "  inner err: $($_.Exception.Message)"
        }
        try {
            $item.Value[0]
        }
        catch {
            Write-Host "  Value[0]: $($_.Exception.Message)"
        }
    }
}
