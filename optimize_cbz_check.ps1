[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Root = ".",

    [Parameter(Mandatory = $false)]
    [string]$LogFile = "unattended_cbz.txt"
)

# Extensions autorisées (en minuscules)
$allowed = @(".jpg", ".jpeg", ".xml", ".css", ".html")

# On vide le fichier de log au démarrage
Set-Content -Path $LogFile -Value ""

# Pour ouvrir les .cbz comme des .zip
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

        # Signature ZIP : 50 4B 03 04
        if ($bytes[0] -eq 0x50 -and
            $bytes[1] -eq 0x4B -and
            $bytes[2] -eq 0x03 -and
            $bytes[3] -eq 0x04) {
            return "zip"
        }

        # Signature RAR v4 : 52 61 72 21 1A 07 00
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

        # Signature RAR v5 : 52 61 72 21 1A 07 01 00
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

function Show-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Folder
    )

    $line = ("[{0}/{1}] Scanning {2}" -f $Current, $Total, $Folder)
    $width = $Host.UI.RawUI.WindowSize.Width
    if ($line.Length -gt $width) {
        $line = $line.Substring(0, $width - 1)
    }

    Write-Host "`r$line" -NoNewline
}

function Show-ResultLine {
    param(
        [string]$Message,
        [string]$Color = "Yellow"
    )

    # Finaliser la ligne de progression actuelle
    Write-Host ""
    # Afficher la ligne de résultat (KO / info)
    Write-Host $Message -ForegroundColor $Color
}

# Résolution du chemin racine
$rootFull = (Resolve-Path $Root).Path

# Liste des dossiers : racine + tous les sous-dossiers
$folders = Get-ChildItem -Path $rootFull -Directory -Recurse -ErrorAction SilentlyContinue
$folders = @($rootFull) + $folders.FullName

$folderCount = $folders.Count
$folderIndex = 0

foreach ($folder in $folders) {
    $folderIndex++

    # Afficher la progression sur une seule ligne (sera remplacée au fur et à mesure)
    Show-Progress -Current $folderIndex -Total $folderCount -Folder $folder

    # Fichiers .cbz dans ce dossier
    $cbzFiles = Get-ChildItem -Path $folder -File -Filter *.cbz -ErrorAction SilentlyContinue

    foreach ($cbz in $cbzFiles) {
        try {
            $type = Get-CbzArchiveType -Path $cbz.FullName

            switch ($type) {
                "zip" {
                    # Analyse du contenu ZIP pour trouver des extensions inattendues (résumé par extension)
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($cbz.FullName)
                    try {
                        $unexpectedExts = @()

                        foreach ($entry in $zip.Entries) {
                            if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }
                            if ($entry.FullName.EndsWith("/")) { continue }

                            $ext = [System.IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()

                            if (-not $allowed.Contains($ext)) {
                                $unexpectedExts += $ext
                            }
                        }

                        if ($unexpectedExts.Count -gt 0) {
                            $grouped = $unexpectedExts |
                                Group-Object |
                                ForEach-Object { "{0}:{1}" -f $_.Name.TrimStart('.'), $_.Count }

                            $summary = $grouped -join ", "
                            $msg = "$($cbz.FullName) : inattendu -> $summary"

                            # KO → on garde la ligne et on log
                            Show-ResultLine -Message "[WARNING] $msg" -Color "Yellow"
                            Add-Content -Path $LogFile -Value $msg -Encoding 1252
                        }
                        else {
                            # CBZ OK : rien à écrire dans $LogFile
                            # (on ne conserve dans le .txt que les CBZ à modifier)
                        }
                    }
                    finally {
                        $zip.Dispose()
                    }
                }
                "rar" {
                    $msg = "$($cbz.FullName) : faux CBZ (archive RAR)"
                    Show-ResultLine -Message "[WARNING] $msg" -Color "Yellow"
                    Add-Content -Path $LogFile -Value $msg
                }
                "missing" {
                    $msg = "$($cbz.FullName) : fichier introuvable (signalé mais disparu)"
                    Show-ResultLine -Message "[WARNING] $msg" -Color "Yellow"
                    Add-Content -Path $LogFile -Value $msg
                }
                default {
                    $msg = "$($cbz.FullName) : en-tête binaire inattendu (ni ZIP ni RAR)"
                    Show-ResultLine -Message "[WARNING] $msg" -Color "Yellow"
                    Add-Content -Path $LogFile -Value $msg
                }
            }
        }
        catch {
            $msg = "$($cbz.FullName) : erreur pendant l'analyse - $($_.Exception.Message)"
            Show-ResultLine -Message "[WARNING] $msg" -Color "Yellow"
            Add-Content -Path $LogFile -Value $msg
        }

        # Après chaque CBZ, la progression sera réaffichée à la prochaine itération
        # via Show-Progress au début de la boucle de dossier (ou du prochain CBZ)
    }
}

# Finaliser la progression et afficher le message de fin
Write-Host ""
Write-Host "`nScan complete. Results saved to $LogFile" -ForegroundColor Green
