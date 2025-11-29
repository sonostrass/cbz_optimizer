# MULTI-THREADED CBZ OPTIMIZER (SIMULATION MODE, ADAPTED)
# This script DOES NOT modify any CBZ files.
# It only simulates size changes and status in cbz_status.csv.

$inputList = "unattended_cbz.txt"
$csvFile   = "cbz_status.csv"

if (!(Test-Path $csvFile)) {
    "cbz;original_size;compression_date;status;optimized_size" | Out-File $csvFile -Encoding UTF8
}

$csv = Import-Csv $csvFile -Delimiter ";"
$cbzList = Get-Content $inputList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

$cbzList | ForEach-Object -Parallel {

    param($line, $csvFile)

    if ([string]::IsNullOrWhiteSpace($line)) {
        return
    }

    # unattended_cbz.txt lines look like:
    #   C:\path\file.cbz : some description...
    $cbzPath = $line.Split(" : ")[0].Trim()

    # Reload CSV inside this runspace
    $csv = Import-Csv $csvFile -Delimiter ";"

    # Find or create CSV entry
    $entry = $csv | Where-Object { $_.cbz -eq $cbzPath }
    if (-not $entry) {
        $entry = [PSCustomObject]@{
            cbz              = $cbzPath
            original_size    = 0
            compression_date = ""
            status           = ""
            optimized_size   = 0
        }
        $csv += $entry
    }

    # Skip if already processed successfully or currently ongoing
    if ($entry.status -eq "success" -or $entry.status -eq "ongoing") {
        return
    }

    # Fill / refresh base information
    if (Test-Path $cbzPath) {
        $entry.original_size = (Get-Item $cbzPath).Length
    } else {
        # if file is missing, simulate but mark as fail
        $entry.original_size = 0
    }

    $entry.compression_date = (Get-Date).ToString("s")
    $entry.status           = "ongoing"
    $csv | Export-Csv $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8

    # --- Simulation of optimization ---
    # Random factor between 0.5 and 1.2 applied on original size
    if ($entry.original_size -gt 0) {
        $rand = Get-Random -Minimum 0.5 -Maximum 1.2
        $entry.optimized_size = [int]($entry.original_size * $rand)
    } else {
        $entry.optimized_size = 0
    }

    if ($entry.optimized_size -lt $entry.original_size -and $entry.original_size -gt 0) {
        $entry.status = "success"
    } else {
        $entry.status = "fail"
    }

    $csv | Export-Csv $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8

} -ThrottleLimit 4 -ArgumentList $csvFile
