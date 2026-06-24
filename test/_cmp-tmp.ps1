Add-Type -AssemblyName System.IO.Compression.FileSystem
function Get-SheetDims {
    param([string]$Path)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sr = New-Object IO.StreamReader($zip.GetEntry('xl/workbook.xml').Open())
        $wb = [xml]$sr.ReadToEnd()
        $sr.Close()
        $names = @($wb.workbook.sheets.sheet | ForEach-Object { $_.name })
        $result = [ordered]@{ File = [IO.Path]::GetFileName($Path); Sheets = @{} }
        for ($i = 0; $i -lt $names.Count; $i++) {
            $sr2 = New-Object IO.StreamReader($zip.GetEntry("xl/worksheets/sheet$($i+1).xml").Open())
            $xml = $sr2.ReadToEnd()
            $sr2.Close()
            $dim = if ($xml -match 'dimension ref="([^"]+)"') { $matches[1] } else { '?' }
            $rows = ([regex]::Matches($xml, '<row ')).Count
            $result.Sheets[$names[$i]] = @{ Dimension = $dim; XmlRows = $rows }
        }
        return $result
    } finally { $zip.Dispose() }
}
Get-SheetDims 'c:\Users\Alex\Downloads\Calista Brice-PermissionsMatrix-20260604-213453.xlsx' | ConvertTo-Json -Depth 4
Get-SheetDims 'c:\Temp\Calista Brice-PermissionsMatrix-20260603-224121.xlsx' | ConvertTo-Json -Depth 4
