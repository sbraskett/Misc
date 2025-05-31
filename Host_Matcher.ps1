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
    $Text = [regex]::Replace($Text, '[\u2013\u2014]', '-')
    $Text = [regex]::Replace($Text, '[\u2018\u2019\u201A]', "'")
    $Text = [regex]::Replace($Text, '[\u201C\u201D\u201E]', '"')
    $Text = $Text -replace '[^\x00-\x7F]', ''
    $Text = $Text -replace '\s+', ' '
    return $Text.Trim().ToLower()
}

if ($Jql -and $JiraFilterId) {
    Write-Error "❌ Use either --Jql or --JiraFilterId, not both."
    exit 1
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

# Load token
$token = [System.IO.File]::ReadAllText("jira_token_cache.pkl")

# Determine JQL
if ($JiraUrl -and $JiraFilterId) {
    $headers = @{ Authorization = "Bearer $token" }
    $filterResp = Invoke-RestMethod "$JiraUrl/rest/api/2/filter/$JiraFilterId" -Headers $headers
    $Jql = $filterResp.jql
    Write-Host "📋 Using JQL from filter $JiraFilterId: $Jql"
}

# Extract hostnames from Jira or FileA
$results = @()

if ($JiraUrl -and $Jql) {
    Write-Host "🔗 Pulling issues from Jira..."
    $startAt = 0
    $total = $null
    $headers = @{ Authorization = "Bearer $token"; Accept = "application/json" }

    do {
        $params = @{
            jql = $Jql
            fields = "key,description"
            startAt = $startAt
            maxResults = 100
        }
        $resp = Invoke-RestMethod -Uri "$JiraUrl/rest/api/2/search" -Headers $headers -Body $params -Method Get
        $issues = $resp.issues
        $total = $resp.total
        Write-Host "🔄 Retrieved $($issues.Count) issues..."

        foreach ($issue in $issues) {
            $id = $issue.key
            $desc = $issue.fields.description
            if (-not $desc) { continue }

            $match = [regex]::Match($desc, $TriggerPattern, 'IgnoreCase')
            if ($match.Success) {
                $after = $desc.Substring($match.Index + $match.Length)
                $line = ($after -split "`n" | Where-Object { $_.Trim() })[0]
                foreach ($server in ($line -split ',')) {
                    $hn = (Clean-String $server).Split('.')[0]
                    if ($hn) { $results += [PSCustomObject]@{ ID = $id; Hostname = $hn } }
                }
            }
        }

        $startAt += 100
    } while ($startAt -lt $total)
}
elseif ($FileA) {
    Write-Host "📂 Reading Excel file A..."
    $wbA = Open-ExcelPackage -Path $FileA
    foreach ($sheet in $wbA.Workbook.Worksheets) {
        for ($r = 1; $r -le $sheet.Dimension.Rows; $r++) {
            $text = $sheet.Cells[$r, $ExtractColumn].Text
            $id = Clean-String $sheet.Cells[$r, $IdColumn].Text
            $match = [regex]::Match($text, $TriggerPattern, 'IgnoreCase')
            if ($match.Success) {
                $after = $text.Substring($match.Index + $match.Length)
                $line = ($after -split "`n" | Where-Object { $_.Trim() })[0]
                foreach ($server in ($line -split ',')) {
                    $hn = (Clean-String $server).Split('.')[0]
                    if ($hn) { $results += [PSCustomObject]@{ ID = $id; Hostname = $hn } }
                }
            }
        }
    }
    Close-ExcelPackage $wbA
}
else {
    Write-Error "❌ Provide either --FileA or Jira parameters."
    exit 1
}

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

# Write CSVs
$details | Export-Csv -Path \"$OutputPrefix`_details.csv\" -NoTypeInformation
Write-Host "✅ Saved: $OutputPrefix`_details.csv"
Write-Host "🔍 Total matched hostnames: $matchedCount"

$summary.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        ID = $_.Key
        'Total Hosts' = $_.Value.Total
        Matched = $_.Value.Matched
        Unmatched = $_.Value.Unmatched
    }
} | Export-Csv -Path \"$OutputPrefix`_summary.csv\" -NoTypeInformation
Write-Host "📊 Saved: $OutputPrefix`_summary.csv"
