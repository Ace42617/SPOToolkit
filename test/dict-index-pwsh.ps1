$d = [System.Collections.Generic.Dictionary[string,object[]]]::new()
$d['a'] = @('x')

$tests = @(
    { param($x) $null = $d[$x] },
    { param($x) $d[$x] = @('y') },
    { param($x) $d.Item($x) },
    { param($x) $d.($x) }
)

foreach ($key in @(0, '0', $null, [int]0)) {
    foreach ($i in 0..($tests.Count-1)) {
        try {
            & $tests[$i] $key
            Write-Host ("OK key=$key test=$i")
        }
        catch {
            Write-Host ("FAIL key=$key test=$i : $($_.Exception.Message)")
        }
    }
}

# foreach col on RoleColumns simulation - dot access on dict
foreach ($col in @('Read', 'Count', 'Keys', 'Values', '0')) {
    try {
        $v = $d.$col
        Write-Host "dot .$col = $v"
    }
    catch {
        Write-Host "dot .$col : $($_.Exception.Message)"
    }
}

# ordered dict row assign with int key from RoleColumns corruption
$row = [ordered]@{}
foreach ($col in @([int]0, 'Read')) {
    try {
        $row[$col] = 'X'
        Write-Host "row[$col] OK"
    }
    catch {
        Write-Host "row[$col]: $($_.Exception.Message)"
    }
}
