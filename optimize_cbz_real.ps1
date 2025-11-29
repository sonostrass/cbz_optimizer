# MULTI-THREADED CBZ OPTIMIZER (REAL MODE, ADAPTED)
# Requires:
# - PowerShell 7 (for ForEach-Object -Parallel)
# - 7z.exe accessible in PATH (for RAR/CBR extraction)
# - Optionally ImageMagick "magick" in PATH if you want image format conversion

using namespace System.IO.Compression

$ErrorActionPreference = "Stop"

$inputList = "unattended_cbz.txt"
$csvFile   = "cbz_status.csv"

if (!(Test-Path $csvFile)) {
    "cbz;original_size;compression_date;status;optimized_size" | Out-File $csvFile -Encoding UTF8
}

$csv = Import-Csv $csvFile -Delimiter ";"
$cbzList = Get-Content $inputList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

# Scriptblock executed in parallel
$cbzList | ForEach-Object -Parallel {

    param($line, $csvFile)

    # --- Local helper: detect archive type based on signature ---
    function Get-CbzArchiveType {
        param(
            [Parameter(Mandatory=$true)]
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

    if ([string]::IsNullOrWhiteSpace($line)) {
        return
    }

    # unattended_cbz.txt lines look like:
    #   C:\path\file.cbz : some description...
    $cbzPath = $line.Split(" : ")[0].Trim()

    if (-not (Test-Path $cbzPath)) {
        Write-Host "[REAL] Missing file: $cbzPath" -ForegroundColor Yellow
        return
    }

    # Reload CSV in this runspace
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

    $entry.status           = "ongoing"
    $entry.compression_date = (Get-Date).ToString("s")

    $csv | Export-Csv $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8

    $tmp = $null
    try {
        $originalSize            = (Get-Item $cbzPath).Length
        $entry.original_size     = $originalSize

        # Temp folder
        $tmp = Join-Path -Path $env:TEMP -ChildPath ("cbzopt_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tmp | Out-Null

        # Detect archive type: ZIP vs RAR (faux CBZ)
        $type            = Get-CbzArchiveType -Path $cbzPath
        $forceRepackOnly = $false

        switch ($type) {
            "zip" {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($cbzPath, $tmp)
            }
            "rar" {
                # Faux CBZ -> extract with 7z, then we will ALWAYS recompress into a proper CBZ
                & 7z x -y "-o$tmp" -- "$cbzPath" | Out-Null
                $forceRepackOnly = $true
            }
            default {
                throw "Unknown archive type for $cbzPath : $type"
            }
        }

        # --- OPTIONAL image conversion ---
        # If you want to convert non-JPEG images to JPEG, enable this block.
        # It will run for both real CBZ and faux CBZ (RAR) so that
        # any "inattendus" detected earlier are normalized.
        $imageExts = ".png",".gif",".bmp",".webp",".tif",".tiff"
        $hasMagick = $false
        try {
            $null = & magick -version 2>$null
            if ($LASTEXITCODE -eq 0) { $hasMagick = $true }
        } catch {
            $hasMagick = $false
        }

        if ($hasMagick) {
            Get-ChildItem -Path $tmp -Recurse -File | Where-Object { $imageExts -contains $_.Extension.ToLowerInvariant() } | ForEach-Object {
                $src = $_.FullName
                $dst = [System.IO.Path]::ChangeExtension($src, ".jpg")

                # Avoid overwriting existing jpeg
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

        # --- Recreate optimized CBZ (ZIP, NoCompression) ---
        $optPath = "$cbzPath.opt"
        if (Test-Path $optPath) {
            Remove-Item $optPath -Force
        }

        $zipStream = [System.IO.File]::Open($optPath, [System.IO.FileMode]::Create)
        $zip       = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)

        try {
            Get-ChildItem -Path $tmp -Recurse -File | ForEach-Object {
                $relPath    = $_.FullName.Substring($tmp.Length).TrimStart('\','/')
                $entryZip   = $zip.CreateEntry($relPath, [System.IO.Compression.CompressionLevel]::NoCompression)
                $entryStream= $entryZip.Open()
                $fileStream = [System.IO.File]::OpenRead($_.FullName)
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

        $optimizedSize          = (Get-Item $optPath).Length
        $entry.optimized_size   = $optimizedSize

        if ($optimizedSize -lt $originalSize) {
            Move-Item $optPath $cbzPath -Force
            $entry.status = "success"
        }
        else {
            Remove-Item $optPath -Force
            $entry.status = "fail"
        }
    }
    catch {
        Write-Host "[REAL] Error on $cbzPath : $($_.Exception.Message)" -ForegroundColor Red
        $entry.status = "fail"
    }
    finally {
        if ($tmp -and (Test-Path $tmp)) {
            Remove-Item $tmp -Recurse -Force
        }
        $csv | Export-Csv $csvFile -Delimiter ";" -NoTypeInformation -Encoding UTF8
    }

} -ThrottleLimit 4 -ArgumentList $csvFile
