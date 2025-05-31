param (
    [string]$FileA,
    [string]$FileB,
    [string]$TriggerPattern,
    [int]$ExtractColumn = 7,
    [int]$IdColumn = 2,
    [int]$ReferenceColumn = 25,
    [string]$OutputPrefix = "host_check",
    [string]$JiraUrl,
    [string]$Jql,
    [string]$JiraFilterId
)

function Clean-String {
    param([string]$Text)
    if (-not $Text) { return "" }

    $Text = $Text -replace '[\u2013\u2014]', '-'         # en/em dashes
    $Text = $Text -replace '[\u2018\u2019\u201A]', "'"  # curly apostrophes
    $Text = $Text -replace '[\u201C\u201D\u201E]', '"'  # curly quotes
    $Text = $Text -replace '[\u00A0\u200B\uFEFF]', ' '  # non-breaking/zero-width space/BOM
    $Text = $Text -replace 'â€“|â€”', '-'
    $Text = $Text -replace 'â€˜|â€™', "'"
    $Text = $Text -replace 'â€œ|â€', '"'
    $Text = $Text -replace 'â€¦', '...'
    $Text = $Text -replace 'Â|â', ''
    $Text = $Text -replace '[^\x00-\x7F]', ''
    $Text = $Text -replace '\s+', ' '
    return $Text.Trim().ToLower()
}

# Load valid hostnames from FileB
$validHosts = @{}
$wbB = Open-ExcelPackage -Path $FileB
foreach ($sheet in $wbB.Workbook.Worksheets) {
    for ($r = 1; $r -le $sheet.Dimension.Rows; $r++) {
        $val = $sheet.Cells[$r, $ReferenceColumn].Text
        if ($val) {
            $hn = (Clean-String $val).Split('.')[0]
            if ($hn -ne '') { $validHosts[$hn] = $true }
        }
    }
}
Close-ExcelPackage $wbB

# Load hostnames from FileA
$results = @()
$wbA = Open-ExcelPackage -Path $FileA
foreach ($sheet in $wbA.Workbook.Worksheets) {
    for ($r = 1; $r -le $sheet.Dimension.Rows; $r++) {
        $rawText = [string]$sheet.Cells[$r, $ExtractColumn].Value
        $id = Clean-String $sheet.Cells[$r, $IdColumn].Text

        if (-not $rawText) { continue }

        $regex = New-Object System.Text.RegularExpressions.Regex(
            $TriggerPattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
        )
        $match = $regex.Match($rawText)

        if ($match.Success) {
            $after = $rawText.Substring($match.Index + $match.Length)
            $after = $after -replace "`r`n|`n|`r", "`n"
            $firstLine = (($after -split "`n" | Where-Object { $_.Trim() -ne "" })[0]).Trim()

            foreach ($server in ($firstLine -split ',')) {
                $hn = (Clean-String ($server.Trim())).Split('.')[0]
                if ($hn) {
                    $results += [PSCustomObject]@{ ID = $id; Hostname = $hn }
                }
            }
        }
    }
}
Close-ExcelPackage $wbA

# Compare to valid hosts
$details = @()
$summary = @{}
$matchedCount = 0
foreach ($rec in $results) {
    $status = if ($validHosts.ContainsKey($rec.Hostname)) { $matchedCount++; 'Matched' } else { 'Unmatched' }
    $details += [PSCustomObject]@{ ID = $rec.ID; Hostname = $rec.Hostname; Status = $status }

    if (-not $summary.ContainsKey($rec.ID)) {
        $summary[$rec.ID] = @{ Total = 0; Matched = 0; Unmatched = 0 }
    }
    $summary[$rec.ID].Total++
    $summary[$rec.ID][$status]++
}

# Export CSVs
$details | Export-Csv -Path "$OutputPrefix`_details.csv" -NoTypeInformation
Write-Host "✅ Saved: $OutputPrefix`_details.csv"
Write-Host "🔍 Total matched hostnames: $matchedCount"

$summary.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        ID = $_.Key
        'Total Hosts' = $_.Value.Total
        Matched = $_.Value.Matched
        Unmatched = $_.Value.Unmatched
    }
} | Export-Csv -Path "$OutputPrefix`_summary.csv" -NoTypeInformation
Write-Host "📊 Saved: $OutputPrefix`_summary.csv"
Write-Host "🔍 $matchedCount hosts matched in $($summary.Keys.Count) unique request IDs"
