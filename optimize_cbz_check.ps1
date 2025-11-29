[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Root = ".",

    [Parameter(Mandatory = $false)]
    [string]$LogFile = "unattended_cbz.txt"
)

# Extensions autorisées (en minuscules)
# Tu peux ajuster ici (.json, .nfo, etc. si besoin)
$allowed = @(".jpg", ".jpeg", ".xml", ".css", ".html")

# On vide le fichier de log au démarrage
Set-Content -Path $LogFile -Value ""

# Pour ouvrir les .cbz comme des .zip
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Résolution du chemin racine
$rootFull = (Resolve-Path $Root).Path

# Liste des dossiers : racine + tous les sous-dossiers
$folders = Get-ChildItem -Path $rootFull -Directory -Recurse -ErrorAction SilentlyContinue
$folders = @($rootFull) + $folders.FullName

$folderCount = $folders.Count
$folderIndex = 0

foreach ($folder in $folders) {
    $folderIndex++

    # Ligne de progression sur une seule ligne
    $progressLine = ("[{0}/{1}] Scanning {2}" -f $folderIndex, $folderCount, $folder)
    $width = $Host.UI.RawUI.WindowSize.Width
    if ($progressLine.Length -ge $width) {
        $progressLine = $progressLine.Substring(0, $width - 1)
    }

    Write-Host "`r$progressLine" -NoNewline

    # Fichiers .cbz dans ce dossier
    $cbzFiles = Get-ChildItem -Path $folder -File -Filter *.cbz -ErrorAction SilentlyContinue

    foreach ($cbz in $cbzFiles) {
        try {
            # --- 1) Détection du vrai format via la signature binaire ---

            # On lit les 8 premiers octets
            $bytes = Get-Content -Path $cbz.FullName -Encoding Byte -TotalCount 8

            if ($bytes.Count -lt 4) {
                $msg = "$($cbz.FullName) : fichier trop petit pour être un CBZ valide"
                Write-Host ""
                Write-Host "[WARNING] $msg" -ForegroundColor Yellow
                Add-Content -Path $LogFile -Value $msg
                continue
            }

            # Signature ZIP : 50 4B 03 04
            $isZip = (
                $bytes[0] -eq 0x50 -and
                $bytes[1] -eq 0x4B -and
                $bytes[2] -eq 0x03 -and
                $bytes[3] -eq 0x04
            )

            # Signature RAR v4 : 52 61 72 21 1A 07 00
            $isRarV4 = (
                $bytes.Count -ge 7 -and
                $bytes[0] -eq 0x52 -and
                $bytes[1] -eq 0x61 -and
                $bytes[2] -eq 0x72 -and
                $bytes[3] -eq 0x21 -and
                $bytes[4] -eq 0x1A -and
                $bytes[5] -eq 0x07 -and
                $bytes[6] -eq 0x00
            )

            # Signature RAR v5 : 52 61 72 21 1A 07 01 00
            $isRarV5 = (
                $bytes.Count -ge 8 -and
                $bytes[0] -eq 0x52 -and
                $bytes[1] -eq 0x61 -and
                $bytes[2] -eq 0x72 -and
                $bytes[3] -eq 0x21 -and
                $bytes[4] -eq 0x1A -and
                $bytes[5] -eq 0x07 -and
                $bytes[6] -eq 0x01 -and
                $bytes[7] -eq 0x00
            )

            if (-not $isZip) {
                # Tout ce qui n'est pas ZIP est traité comme "inattendu"
                if ($isRarV4 -or $isRarV5) {
                    $msg = "$($cbz.FullName) : faux CBZ (archive CBR/RAR déguisée)"
                }
                else {
                    $msg = "$($cbz.FullName) : en-tête binaire inattendu (ni ZIP ni RAR)"
                }

                Write-Host ""
                Write-Host "[WARNING] $msg" -ForegroundColor Yellow
                Add-Content -Path $LogFile -Value $msg
                continue    # on ne tente pas l'ouverture ZIP
            }

            # --- 2) Analyse du contenu ZIP pour trouver des fichiers inattendus ---

            $zip = [System.IO.Compression.ZipFile]::OpenRead($cbz.FullName)
            try {
                $unexpected = @()

                foreach ($entry in $zip.Entries) {
                    if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }

                    # Répertoires (dans un ZIP, souvent terminés par /)
                    if ($entry.FullName.EndsWith("/")) { continue }

                    $ext = [System.IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()

                    # Extension non autorisée => contenu inattendu
                    if (-not $allowed.Contains($ext)) {
                        $unexpected += $entry.FullName
                    }
                }

                if ($unexpected.Count -gt 0) {
                    $unexpectedList = $unexpected -join ", "
                    $msg = "$($cbz.FullName) : contenu inattendu -> $unexpectedList"
                    Write-Host ""
                    Write-Host "[WARNING] $msg" -ForegroundColor Yellow
                    Add-Content -Path $LogFile -Value $msg
                }
            }
            finally {
                $zip.Dispose()
            }
        }
        catch {
            $msg = "$($cbz.FullName) : erreur pendant l'analyse - $($_.Exception.Message)"
            Write-Host ""
            Write-Host "[WARNING] $msg" -ForegroundColor Yellow
            Add-Content -Path $LogFile -Value $msg
        }
    }
}

Write-Host ""
Write-Host "`nScan complete. Results saved to $LogFile" -ForegroundColor Green
