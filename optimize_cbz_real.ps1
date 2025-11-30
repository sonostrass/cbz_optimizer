# MULTI-THREADED CBZ OPTIMIZER (REAL MODE, PARALLEL SAFE)
# Requires:
# - PowerShell 7 (for ForEach-Object -Parallel)
# - 7z.exe accessible in PATH (for RAR/CBR extraction of faux CBZ)
# - ImageMagick "magick" in PATH for image format conversion (PNG/GIF/BMP/WEBP/TIF/TIFF -> JPG)
#
# IMPORTANT RULES (as per project specs):
# - Only non-JPG/JPEG images are converted (PNG, GIF, BMP, WEBP, TIF, TIFF).
# - Existing JPG/JPEG files are NEVER recompressed to avoid quality loss.
# - For faux CBZ (RAR/CBR) that already contain only JPG/JPEG, we ONLY change the container
#   to a proper ZIP-based CBZ, without touching JPG/JPEG data.

using namespace System.IO.Compression

param(
    [string]$InputList = "unattended_cbz.txt",
    [string]$CsvFile   = "cbz_status.csv",
    [int]   $Throttle  = 4
)

$ErrorActionPreference = "Stop"

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7+ (ForEach-Object -Parallel)." -ForegroundColor Red
    exit 1
}

Write-Host "Starting CBZ optimization REAL MODE with throttle limit $Throttle"

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

# Detect external tools once (to avoid overhead inside parallel runs)
$magickAvailable   = (Get-Command magick -ErrorAction SilentlyContinue) -ne $null
$sevenZipAvailable = (Get-Command 7z     -ErrorAction SilentlyContinue) -ne $null

if (-not $sevenZipAvailable) {
    Write-Host "WARNING: 7z.exe not found in PATH. RAR/CBR faux CBZ cannot be processed." -ForegroundColor Yellow
}

if (-not $magickAvailable) {
    Write-Host "WARNING: magick (ImageMagick) not found in PATH. Non-JPG images will not be converted." -ForegroundColor Yellow
}

# Extensions for images that may be converted to JPG (JPG/JPEG are intentionally excluded)
$imageExts = ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"

# Build jobs list from unattended_cbz.txt
$rawLines = Get-Content -Path $InputList -Encoding 1252 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

# Build jobs list from unattended_cbz.txt
$jobs = for ($i = 0; $i -lt $rawLines.Count; $i++) {
    $line = $rawLines[$i]
    $cbzPath = $line.Split(" : ")[0].Trim()
    if (-not $cbzPath) { continue }

    [PSCustomObject]@{
        Index = $i + 1
        Total = $rawLines.Count
        Line  = $line
        Cbz   = $cbzPath
    }
}

if (-not $jobs -or $jobs.Count -eq 0) {
    Write-Host "Nothing to process (no new CBZ to optimize)." -ForegroundColor Green
    exit 0
}

# Parallel processing block
$results = $jobs | ForEach-Object -Parallel {
    $job = $_

    # Per-job progress display
    Write-Host "[REAL] $($job.Index)/$($job.Total) - $($job.Cbz)"

    # --- Local helper: detect archive type based on signature ---
    function Get-CbzArchiveType {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )

        if (-not (Test-Path $Path)) {
            return "missing"
        }

        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $bytes = New-Object byte[] 8
            $read  = $fs.Read($bytes, 0, $bytes.Length)

            if ($read -lt 4) {
                return "unknown"
            }

            # ZIP : 50 4B 03 04
            if ($bytes[0] -eq 0x50 -and
                $bytes[1] -eq 0x4B -and
                $bytes[2] -eq 0x03 -and
                $bytes[3] -eq 0x04) {
                return "zip"
            }

            # RAR v4 : 52 61 72 21 1A 07 00
            if ($read -ge 7 -and
                $bytes[0] -eq 0x52 -and
                $bytes[1] -eq 0x61 -and
                $bytes[2] -eq 0x72 -and
                $bytes[3] -eq 0x21 -and
                $bytes[4] -eq 0x1A -and
                $bytes[5] -eq 0x07 -and
                $bytes[6] -eq 0x00) {
                return "rar"
            }

            # RAR v5 : 52 61 72 21 1A 07 01 00
            if ($read -ge 8 -and
                $bytes[0] -eq 0x52 -and
                $bytes[1] -eq 0x61 -and
                $bytes[2] -eq 0x72 -and
                $bytes[3] -eq 0x21 -and
                $bytes[4] -eq 0x1A -and
                $bytes[5] -eq 0x07 -and
                $bytes[6] -eq 0x01 -and
                $bytes[7] -eq 0x00) {
                return "rar"
            }

            return "unknown"
        }
        finally {
            $fs.Dispose()
        }
    }

    $cbzPath = $job.Cbz

    $result = [PSCustomObject]@{
        cbz              = $cbzPath
        original_size    = 0
        compression_date = (Get-Date).ToString("s")
        status           = ""
        optimized_size   = 0
    }

    if (-not (Test-Path $cbzPath)) {
        Write-Host "[REAL] Missing file: $cbzPath" -ForegroundColor Yellow
        $result.status = "fail"
        return $result
    }

    try {
        $result.original_size = (Get-Item $cbzPath).Length
    }
    catch {
        $result.original_size = 0
    }

    $tmp = $null

    try {
        # Temp folder
        $tmp = Join-Path -Path $env:TEMP -ChildPath ("cbzopt_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tmp | Out-Null

        # Detect archive type: ZIP vs RAR (faux CBZ)
        $type = Get-CbzArchiveType -Path $cbzPath

        switch ($type) {
            "zip" {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($cbzPath, $tmp)
            }
            "rar" {
                if (-not $using:sevenZipAvailable) {
                    throw "7z.exe not found in PATH but required to extract RAR/CBR faux CBZ."
                }
                # Faux CBZ -> extract with 7z, then re-compress into a proper CBZ
                & 7z x -y "-o$tmp" -- "$cbzPath" | Out-Null
            }
            "missing" {
                throw "File reported missing while re-checking: $cbzPath"
            }
            default {
                throw "Unknown archive type for $cbzPath : $type"
            }
        }

        # --- Image conversion step ---
        # Only non-JPG/JPEG formats listed in $using:imageExts are converted to JPG.
        if ($using:magickAvailable) {
            Get-ChildItem -Path $tmp -Recurse -File |
                Where-Object { $using:imageExts -contains $_.Extension.ToLowerInvariant() } |
                ForEach-Object {
                    $src = $_.FullName
                    $dst = [System.IO.Path]::ChangeExtension($src, ".jpg")

                    # Avoid overwriting an existing JPG/JPEG
                    if (Test-Path $dst) {
                        $dst = [System.IO.Path]::Combine(
                            [System.IO.Path]::GetDirectoryName($dst),
                            ([System.IO.Path]::GetFileNameWithoutExtension($dst) + "_jpg" + ".jpg")
                        )
                    }

                    & magick "$src" -quality 90 "$dst"
                    if ($LASTEXITCODE -eq 0 -and (Test-Path $dst)) {
                        Remove-Item $src -Force
                    }
                }
        }
        else {
            Write-Host "[REAL] magick not available, non-JPG images left as-is for $cbzPath" -ForegroundColor Yellow
        }

        # --- Recreate optimized CBZ (ZIP, NoCompression) ---
        $optPath = "$cbzPath.opt"
        if (Test-Path $optPath) {
            Remove-Item $optPath -Force
        }

        $zipStream = [System.IO.File]::Open($optPath, [System.IO.FileMode]::Create)
        $zip       = [System.IO.Compression.ZipArchive]::new(
            $zipStream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false
        )

        try {
            Get-ChildItem -Path $tmp -Recurse -File | ForEach-Object {
                $relPath     = $_.FullName.Substring($tmp.Length).TrimStart('\', '/')
                $entryZip    = $zip.CreateEntry($relPath, [System.IO.Compression.CompressionLevel]::NoCompression)
                $entryStream = $entryZip.Open()
                $fileStream  = [System.IO.File]::OpenRead($_.FullName)
                try {
                    $fileStream.CopyTo($entryStream)
                }
                finally {
                    $fileStream.Dispose()
                    $entryStream.Dispose()
                }
            }
        }
        finally {
            $zip.Dispose()
            $zipStream.Dispose()
        }

        if (Test-Path $optPath) {
            $result.optimized_size = (Get-Item $optPath).Length
        }
        else {
            $result.optimized_size = 0
        }

        if ($result.original_size -gt 0 -and
            $result.optimized_size -gt 0 -and
            $result.optimized_size -lt $result.original_size) {

            # Keep optimized file
            Move-Item -Force $optPath $cbzPath
            $result.status = "success"
        }
        else {
            # Optimization not beneficial or failed -> discard .opt
            if (Test-Path $optPath) {
                Remove-Item $optPath -Force
            }
            $result.status = "fail"
        }
    }
    catch {
        Write-Host "[REAL] Error on $cbzPath : $($_.Exception.Message)" -ForegroundColor Red
        $result.status = "fail"
    }
    finally {
        if ($tmp -and (Test-Path $tmp)) {
            Remove-Item $tmp -Recurse -Force
        }
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

Write-Host "REAL optimization complete for $($results.Count) CBZ file(s)." -ForegroundColor Green