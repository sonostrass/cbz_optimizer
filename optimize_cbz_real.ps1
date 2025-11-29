[CmdletBinding()]
param(
    [string]$InputList = "unattended_cbz.txt",
    [string]$CsvFile   = "cbz_status.csv",
    [int]$ThrottleLimit = 4
)

# Ensure required .NET assembly for ZIP operations is loaded
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create CSV file if it does not exist
if (-not (Test-Path -LiteralPath $CsvFile)) {
    "cbz;original_size;compression_date;status;optimized_size" | Out-File -LiteralPath $CsvFile -Encoding UTF8
}

# Load list of CBZ to process (each line: 'full\path\file.cbz : description...')
if (-not (Test-Path -LiteralPath $InputList)) {
    Write-Error "Input list '$InputList' not found."
    exit 1
}

$cbzList = Get-Content -LiteralPath $InputList -ErrorAction Stop

# Helper: detect archive type from first bytes
function Get-CbzArchiveType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return "missing"
    }

    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $bytes = New-Object byte[] 8
        $read = $fs.Read($bytes, 0, $bytes.Length)

        if ($read -lt 4) { return "unknown" }

        # ZIP signature: 50 4B 03 04
        if ($bytes[0] -eq 0x50 -and
            $bytes[1] -eq 0x4B -and
            $bytes[2] -eq 0x03 -and
            $bytes[3] -eq 0x04) {
            return "zip"
        }

        # RAR v4: 52 61 72 21 1A 07 00
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

        # RAR v5: 52 61 72 21 1A 07 01 00
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

# Detect ImageMagick (magick.exe) once in parent scope
$magickCmd = Get-Command magick -ErrorAction SilentlyContinue
$magickAvailable = $null -ne $magickCmd

Write-Host "ImageMagick available: $magickAvailable"
Write-Host "Starting CBZ optimization with throttle limit $ThrottleLimit"

$cbzList | ForEach-Object -Parallel {
    param($line, $CsvFile, $magickAvailable)

    # Each line: 'full\path\file.cbz : description...'
    $cbzPath = $line.Split(" : ", 2)[0].Trim()
    if (-not $cbzPath) { return }

    # Reload CSV in this runspace
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

    $tmp = $null
    try {
        if (-not (Test-Path -LiteralPath $cbzPath)) {
            throw "File not found: $cbzPath"
        }

        $originalSize = (Get-Item -LiteralPath $cbzPath).Length
        $entry.original_size = $originalSize

        # Create temp directory
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("cbzopt_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tmp -Force | Out-Null

        # Detect archive type
        $type = Get-CbzArchiveType -Path $cbzPath
        if ($type -eq "missing") {
            throw "File missing: $cbzPath"
        }

        switch ($type) {
            "zip" {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($cbzPath, $tmp)
            }
            "rar" {
                # Faux CBZ -> extract with 7z (must be in PATH)
                $7z = Get-Command 7z -ErrorAction SilentlyContinue
                if (-not $7z) {
                    throw "7z command not found in PATH. Cannot extract RAR-based CBZ: $cbzPath"
                }
                & $7z.Source x -y "-o$tmp" -- $cbzPath | Out-Null
            }
            default {
                throw "Unknown archive type for $cbzPath: $type"
            }
        }

        # Decide whether to convert images or only change container
        $shouldConvertImages = $true
        if ($type -eq "rar") {
            # For faux CBZ: if only JPG/JPEG images are present, do NOT reconvert them
            $allImageExts = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff")
            $imageFiles = Get-ChildItem -Path $tmp -Recurse -File |
                Where-Object { $allImageExts -contains $_.Extension.ToLowerInvariant() }

            $hasNonJpg = $imageFiles | Where-Object {
                $_.Extension.ToLowerInvariant() -notin @(".jpg", ".jpeg")
            }

            if (-not $hasNonJpg) {
                # Faux CBZ with only JPG/JPEG images
                $shouldConvertImages = $false
            }
        }

        # Convert non-JPG images (PNG/WEBP/BMP/etc.) to JPG if requested
        if ($magickAvailable -and $shouldConvertImages) {
            $imageExtsToConvert = @(".png", ".gif", ".bmp", ".webp", ".tif", ".tiff")

            Get-ChildItem -Path $tmp -Recurse -File |
                Where-Object { $imageExtsToConvert -contains $_.Extension.ToLowerInvariant() } |
                ForEach-Object {
                    $src = $_.FullName
                    $dir = Split-Path $src -Parent
                    $base = [System.IO.Path]::GetFileNameWithoutExtension($src)
                    $dest = Join-Path $dir ($base + ".jpg")

                    # Convert with ImageMagick, modest quality to keep size low
                    & magick $src "-quality" "90" $dest

                    if (Test-Path -LiteralPath $dest) {
                        Remove-Item -LiteralPath $src -Force
                    }
                }
        }

        # Rebuild CBZ as ZIP (STORE / NoCompression)
        $optPath = "$cbzPath.opt"
        if (Test-Path -LiteralPath $optPath) {
            Remove-Item -LiteralPath $optPath -Force
        }

        $zipStream = [System.IO.File]::Open($optPath, [System.IO.FileMode]::Create)
        $zip = [System.IO.Compression.ZipArchive]::new(
            $zipStream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false
        )

        try {
            Get-ChildItem -Path $tmp -Recurse -File | ForEach-Object {
                $full = $_.FullName
                $rel  = $full.Substring($tmp.Length).TrimStart('\','/')

                $entryZip = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::NoCompression)
                $entryStream = $entryZip.Open()
                $fileStream  = [System.IO.File]::OpenRead($full)
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

        $optimizedSize = (Get-Item -LiteralPath $optPath).Length
        $entry.optimized_size = $optimizedSize

        if ($optimizedSize -lt $originalSize) {
            Move-Item -LiteralPath $optPath -Destination $cbzPath -Force
            $entry.status = "success"
        }
        else {
            Remove-Item -LiteralPath $optPath -Force
            $entry.status = "fail"
        }
    }
    catch {
        $entry.status = "fail"
        Write-Warning "Error processing $cbzPath : $($_.Exception.Message)"
    }
    finally {
        if ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Recurse -Force
        }
        $csv | Export-Csv -LiteralPath $CsvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
    }

} -ThrottleLimit $ThrottleLimit -ArgumentList $CsvFile, $magickAvailable
