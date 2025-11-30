# MULTI-THREADED CBZ OPTIMIZER (SIMULATION MODE, PARALLEL SAFE)
# This script DOES NOT modify any CBZ files.
# It only simulates size changes and status in cbz_status.csv.
# Requires PowerShell 7+ for ForEach-Object -Parallel.

param(
    [string]$InputList = "unattended_cbz.txt",
    [string]$CsvFile   = "cbz_status.csv",
    [int]   $Throttle  = 4
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7+ (ForEach-Object -Parallel)." -ForegroundColor Red
    exit 1
}

Write-Host "Starting CBZ optimization SIMULATION with throttle limit $Throttle"

if (-not (Test-Path $InputList)) {
    Write-Host "Input list '$InputList' not found. Nothing to do." -ForegroundColor Yellow
    exit 0
}

# Ensure CSV exists with header
if (-not (Test-Path $CsvFile)) {
    "cbz;original_size;compression_date;status;optimized_size" | Out-File $CsvFile -Encoding UTF8
}

# Load existing CSV into a map keyed by cbz path
$existingMap = @{}
$existingRows = @()

if (Test-Path $CsvFile) {
    $existingRows = Import-Csv $CsvFile -Delimiter ";"
    foreach ($row in $existingRows) {
        if ($null -ne $row.cbz -and $row.cbz.Trim() -ne "") {
            $existingMap[$row.cbz] = $row
        }
    }
}

# Build jobs list from unattended_cbz.txt, skipping already-successful/ongoing entries
$rawLines = Get-Content -Path $InputList -Encoding Default | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

$jobs = foreach ($line in $rawLines) {
    $cbzPath = $line.Split(" : ")[0].Trim()
    if (-not $cbzPath) { continue }

    if ($existingMap.ContainsKey($cbzPath)) {
        $s = $existingMap[$cbzPath].status
        if ($s -eq "success" -or $s -eq "ongoing") {
            continue
        }
    }

    [PSCustomObject]@{
        Line = $line
        Cbz  = $cbzPath
    }
}

if (-not $jobs -or $jobs.Count -eq 0) {
    Write-Host "Nothing to simulate (no new CBZ to process)." -ForegroundColor Green
    exit 0
}

# Parallel simulation: no CSV I/O inside the parallel block, we only return results
$results = $jobs | ForEach-Object -Parallel {
    $job = $_

    $cbzPath = $job.Cbz

    $result = [PSCustomObject]@{
        cbz              = $cbzPath
        original_size    = 0
        compression_date = (Get-Date).ToString("s")
        status           = ""
        optimized_size   = 0
    }

    # Defensive check: some entries might be null/empty, skip them cleanly
    if ([string]::IsNullOrWhiteSpace($cbzPath)) {
        $result.status = "fail"
        return $result
    }

    if (-not (Test-Path $cbzPath)) {
        Write-Host "[SIM] Missing file: $cbzPath" -ForegroundColor Yellow
        $result.status = "fail"
        return $result
    }

    try {
        $result.original_size = (Get-Item $cbzPath).Length
    }
    catch {
        $result.original_size = 0
    }

    if ($result.original_size -gt 0) {
        # Simulate an optimization ratio between 0.5 and 1.2
        $rand = Get-Random -Minimum 0.5 -Maximum 1.2
        $result.optimized_size = [int]($result.original_size * $rand)
    }
    else {
        $result.optimized_size = 0
    }

    if ($result.original_size -gt 0 -and $result.optimized_size -lt $result.original_size) {
        $result.status = "success"
    }
    else {
        $result.status = "fail"
    }

    return $result

} -ThrottleLimit $Throttle

# Merge results back into the existing map and write CSV once
foreach ($r in $results) {
    if ($null -ne $r.cbz -and $r.cbz.Trim() -ne "") {
        $existingMap[$r.cbz] = $r
    }
}

$finalRows = $existingMap.GetEnumerator() | ForEach-Object { $_.Value } | Sort-Object cbz
$finalRows | Export-Csv $CsvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8

Write-Host "Simulation complete for $($results.Count) CBZ file(s)." -ForegroundColor Green
