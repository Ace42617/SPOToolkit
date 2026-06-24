Add-Type -AssemblyName System.IO.Compression.FileSystem
function Inspect-Xlsx {
    param([string]$Path)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $result = [ordered]@{ File = [IO.Path]::GetFileName($Path) }
        $wbSr = New-Object IO.StreamReader($zip.GetEntry('xl/workbook.xml').Open())
        $wb = [xml]$wbSr.ReadToEnd(); $wbSr.Close()
        $result.Sheets = @($wb.workbook.sheets.sheet | ForEach-Object {
            $tab = if ($_.tabColor) { $_.tabColor.rgb } else { '(none)' }
            [ordered]@{ Name = $_.name; TabColor = $tab }
        })
        $stSr = New-Object IO.StreamReader($zip.GetEntry('xl/styles.xml').Open())
        $st = $stSr.ReadToEnd(); $stSr.Close()
        $result.Styles = [ordered]@{
            FontCount = ([regex]::Match($st, 'fonts count="(\d+)"')).Groups[1].Value
            FillCount = ([regex]::Match($st, 'fills count="(\d+)"')).Groups[1].Value
            XfCount = ([regex]::Match($st, 'cellXfs count="(\d+)"')).Groups[1].Value
            SegoeUI = ($st -match 'Segoe UI')
            Calibri = ($st -match 'Calibri')
        }
        $s1Sr = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
        $s1 = $s1Sr.ReadToEnd(); $s1Sr.Close()
        $result.SummarySheet = [ordered]@{
            HasDrawings = ($zip.GetEntry('xl/drawings/drawing1.xml') -ne $null)
            FreezePane = ($s1 -match 'pane state="frozen"')
            StyledCells = ([regex]::Matches($s1, ' s="')).Count
            MergeCells = ([regex]::Matches($s1, '<mergeCell')).Count
        }
        # Matrix sheet index
        $names = @($wb.workbook.sheets.sheet | ForEach-Object { $_.name })
        $mxIdx = [array]::IndexOf($names, 'Permissions Matrix') + 1
        if ($mxIdx -gt 0) {
            $mxSr = New-Object IO.StreamReader($zip.GetEntry("xl/worksheets/sheet$mxIdx.xml").Open())
            $mx = $mxSr.ReadToEnd(); $mxSr.Close()
            $result.MatrixSheet = [ordered]@{
                StyledCells = ([regex]::Matches($mx, ' s="')).Count
                AutoFilter = ($mx -match 'autoFilter')
                FreezePane = ($mx -match 'pane state="frozen"')
            }
        }
        return $result
    } finally { $zip.Dispose() }
}
Inspect-Xlsx 'c:\Temp\Calista Brice-PermissionsMatrix-20260603-224121.xlsx' | ConvertTo-Json -Depth 5
Inspect-Xlsx 'c:\Temp\_cmp-213453.xlsx' | ConvertTo-Json -Depth 5
