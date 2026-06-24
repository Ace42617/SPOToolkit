$ErrorActionPreference = 'Continue'
$d = [System.Collections.Generic.Dictionary[string,object[]]]::new()
$d2 = [System.Collections.Generic.Dictionary[string,[object]]]::new()
$d2['a'] = @([pscustomobject]@{ Id=1; Member=@{ Title='u' } })

foreach ($name in @('d', 'd2')) {
    $dict = (Get-Variable $name).Value
    Write-Host "=== $name type: $($dict.GetType().FullName) ==="
    Write-Host "  is Dictionary[string,object[]]: $($dict -is [System.Collections.Generic.Dictionary[string, object[]]])"
    try {
        $dict._ResolvedRoleNames = @('Read')
        Write-Host '  dot-set OK'
    }
    catch {
        Write-Host ("  dot-set: {0}" -f $_.Exception.Message)
    }
    try {
        $dict | Add-Member -NotePropertyName '_ResolvedRoleNames' -NotePropertyValue @('Read') -Force
        Write-Host '  Add-Member OK'
    }
    catch {
        Write-Host ("  Add-Member: {0}" -f $_.Exception.Message)
    }
    try {
        $dict['RoleDefinitionBindings'] = @()
        Write-Host '  string index set OK'
    }
    catch {
        Write-Host ("  string index set: {0}" -f $_.Exception.Message)
    }
    try {
        $dict[0] = @()
        Write-Host '  int index set OK'
    }
    catch {
        Write-Host ("  int index set: {0}" -f $_.Exception.Message)
    }
}

# Simulate Enrich on dict as assignment
$assignment = $d
try {
    $roles = @('Read')
    if ($assignment.PSObject.Properties.Match('_ResolvedRoleNames').Count -gt 0) {
        $assignment._ResolvedRoleNames = $roles
    }
    else {
        $assignment | Add-Member -NotePropertyName '_ResolvedRoleNames' -NotePropertyValue $roles -Force
    }
    Write-Host 'Enrich simulation OK'
}
catch {
    Write-Host ("Enrich simulation: {0}" -f $_.Exception.Message)
}
