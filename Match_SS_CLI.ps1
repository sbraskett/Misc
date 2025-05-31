# Usage:
#.\host_matcher.ps1 `
#  -FileA "spreadsheet_a.xlsx" `
#  -FileB "spreadsheet_b.xlsx" `
#  -TriggerPattern "server names.*?encrypting is hosted on:" `
#  -ExtractColumn 7 `
#  -IdColumn 2 `
#  -ReferenceColumn 25 `
#  -OutputPrefix "results"


param(
    [string]$FileA,
    [string]$FileB,
    [string]$TriggerPattern,
    [int]$ExtractColumn = 7,
    [int]$IdColumn = 2,
    [int]$ReferenceColumn = 25,
    [string]$OutputPrefix = "host_check"
)

function Clean-String {
    param ([string]$s)
    if (-not $s) { return "" }

    $s = $s -replace 'â€“', '-'
    $s = $s -replace 'â€”', '-'
    $s = $s -replace 'â€˜', "'"
    $s = $s -replace 'â€™', "'"
    $s = $s -replace 'â€œ', '"'
    $s = $s -replace 'â€', '"'
    $s = $s -replace 'â€¦', '...'
    $s = $s -replace 'Â', ''
    $s = $s -replace 'â', ''
    $s = -join ($s.ToCharArray() | Where-Object { [int]$_ -le 127 })
    $s = $s -replace '\s+', ' '
    return $s.Trim().ToLower()
}

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Load FileB and build reference host set
$wbB = $excel.Workbooks.Open($FileB)
$validHosts = @{}
foreach ($sheet in $wbB.Sheets) {
    $lastRow = $sheet.UsedRange.Rows.Count
    for ($r = 1; $r -le $lastRow; $r++) {
        $val = $sheet.Cells.Item($r, $ReferenceColumn).Text
        $host = (Clean-String $val).Split('.')[0]
        if ($host) { $validHosts[$host] = $true }
    }
}
$wbB.Close($false)

# Load FileA and extract hostnames
$wbA = $excel.Workbooks.Open($FileA)
$results = @()

foreach ($sheet in $wbA.Sheets) {
    $lastRow = $sheet.UsedRange.Rows.Count
    for ($r = 1; $r -le $lastRow; $r++) {
        $colB = Clean-String $sheet.Cells.Item($r, $IdColumn).Text
        $colG = Clean-String $sheet.Cells.Item($r, $ExtractColumn).Text

        if ($colG -match $TriggerPattern) {
            $after = $colG.Substring($matches[0].Length).Trim()
            $firstLine = ($after -split "`r?`n")[0]
            $servers = $firstLine -split ','

            foreach ($server in $servers) {
                $hostname = (Clean-String $server).Split('.')[0]
                if ($hostname) {
                    $status = if ($validHosts.ContainsKey($hostname)) { "Matched" } else { "Unmatched" }
                    $results += [PSCustomObject]@{
                        ColumnB  = $colB
                        Hostname = $hostname
                        Status   = $status
                    }
                }
            }
        }
    }
}
$wbA.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# Output: Details CSV
$detailPath = "$OutputPrefix`_details.csv"
$results | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8
Write-Output "✅ Detailed output saved to: $detailPath"

# Output: Summary CSV
$summary = $results | Group-Object ColumnB | ForEach-Object {
    $group = $_.Group
    $total = $group.Count
    $matched = ($group | Where-Object { $_.Status -eq "Matched" }).Count
    $unmatched = $total - $matched

    [PSCustomObject]@{
        ColumnB   = $_.Name
        Total     = $total
        Matched   = $matched
        Unmatched = $unmatched
    }
}

$summaryPath = "$OutputPrefix`_summary.csv"
$summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
Write-Output "📊 Summary output saved to: $summaryPath"
