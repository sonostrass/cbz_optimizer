[CmdletBinding()]
param(
    [string]$InputList = "unattended_cbz.txt",
    [string]$CsvFile   = "cbz_status.csv",
    [int]$ThrottleLimit = 4
)

# Create CSV file if it does not exist
if (-not (Test-Path -LiteralPath $CsvFile)) {
    "cbz;original_size;compression_date;status;optimized_size" | Out-File -LiteralPath $CsvFile -Encoding UTF8
}

if (-not (Test-Path -LiteralPath $InputList)) {
    Write-Error "Input list '$InputList' not found."
    exit 1
}

$cbzList = Get-Content -LiteralPath $InputList -ErrorAction Stop

Write-Host "Starting CBZ optimization SIMULATION with throttle limit $ThrottleLimit"

$cbzList | ForEach-Object -Parallel {
    param($line, $CsvFile)

    $cbzPath = $line.Split(" : ", 2)[0].Trim()
    if (-not $cbzPath) { return }

    $csv = Import-Csv -LiteralPath $CsvFile -Delimiter ";"

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

    if ($entry.status -eq "success" -or $entry.status -eq "ongoing") {
        return
    }

    $entry.status           = "ongoing"
    $entry.compression_date = (Get-Date).ToString("s")
    $csv | Export-Csv -LiteralPath $CsvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8

    try {
        if (-not (Test-Path -LiteralPath $cbzPath)) {
            throw "File not found: $cbzPath"
        }

        $fileInfo = Get-Item -LiteralPath $cbzPath
        $originalSize = [int64]$fileInfo.Length
        $entry.original_size = $originalSize

        # Simulate an optimization factor between 0.70 and 1.10
        $rand = Get-Random -Minimum 70 -Maximum 111
        $factor = [double]$rand / 100.0

        $simOptimized = [int64]([math]::Round($originalSize * $factor))
        $entry.optimized_size = $simOptimized

        if ($simOptimized -lt $originalSize) {
            $entry.status = "success"
        }
        else {
            $entry.status = "fail"
        }
    }
    catch {
        $entry.status = "fail"
        Write-Warning "Simulation error for $cbzPath : $($_.Exception.Message)"
    }
    finally {
        $csv | Export-Csv -LiteralPath $CsvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
    }

} -ThrottleLimit $ThrottleLimit -ArgumentList $CsvFile
