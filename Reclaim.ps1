param(
    [string]$Path = ".",
    [switch]$Recurse
)

$gciParams = @{ Path = $Path; File = $true }
if ($Recurse) { $gciParams.Recurse = $true }

function Get-Sha256([string]$FilePath) {
    try { return (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256).Hash }
    catch { return $null }
}

# 1. List all files (skip empty files)
$allFiles = Get-ChildItem @gciParams | Where-Object { $_.Length -gt 0 }

# 2. Size pre-filter: keep only sizes that occur more than once
$sizesWithDupes = $allFiles |
    Group-Object Length |
    Where-Object Count -gt 1 |
    ForEach-Object { [int64]$_.Name }

# 3. Hash only the candidate files
$files = $allFiles |
    Where-Object { $_.Length -in $sizesWithDupes } |
    ForEach-Object {
        $h = Get-Sha256 $_.FullName
        if ($h) {
            [PSCustomObject]@{
                Hash = $h
                Name = $_.Name
                Path = $_.FullName
                Size = $_.Length
            }
        }
    }

# 4. Group by hash, keep only groups with more than one file
$dupeGroups = $files | Group-Object Hash | Where-Object Count -gt 1

# 5. Output duplicate groups
$dupeGroups | ForEach-Object {
    $_.Group | Select-Object @{n='HashPrefix';e={$_.Hash.Substring(0,12)}}, Name, Size, Path
} | Format-Table -AutoSize

# Export-Csv -Path ".\duplicate-report.csv" -NoTypeInformation -Encoding UTF8

# 6. Summary
if ($dupeGroups) {
    $totalBytes       = ($dupeGroups | ForEach-Object { $_.Group } | Measure-Object -Property Size -Sum).Sum
    $reclaimableBytes = ($dupeGroups | ForEach-Object {
        $sizes = $_.Group | Sort-Object Size -Descending
        ($sizes | Select-Object -Skip 1 | Measure-Object -Property Size -Sum).Sum
    } | Measure-Object -Sum).Sum

    $fileCount  = ($dupeGroups | ForEach-Object { $_.Group.Count } | Measure-Object -Sum).Sum
    $groupCount = $dupeGroups.Count

    ""
    "Duplicate groups:             $groupCount"
    "Files in dup groups:          $fileCount"
    "Total size of dupes:          {0:N2} MB" -f ($totalBytes / 1MB)
    "Reclaimable (delete extras):  {0:N2} MB" -f ($reclaimableBytes / 1MB)
} else {
    ""
    "No duplicate content found."
}
