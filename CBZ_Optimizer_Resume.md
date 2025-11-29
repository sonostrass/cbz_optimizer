# 📘 CBZ Optimizer --- Résumé du Projet

Ce projet vise à **optimiser automatiquement des fichiers CBZ**
(archives ZIP utilisées pour les bandes dessinées), en réduisant leur
taille **sans perte significative de qualité** tout en garantissant
**zéro perte de fichiers non-images** (ex : `ComicInfo.xml`).

## 🎯 Objectifs

-   Convertir les images **PNG / BMP / WEBP** en **JPG**.
-   Conserver **tous les fichiers originaux**.
-   Recréer le CBZ en mode **STORE (sans compression ZIP)**.
-   Ne remplacer un CBZ que si l'optimisé est **plus petit**.
-   Utiliser un **CSV de suivi** persistant.
-   Permettre la **reprise après crash**.
-   Proposer une version **réelle** et une version **simulation**.
-   Accélérer grâce au **multithreading PowerShell 7**.

## 🔍 Étape 1 --- Détection

Un script scanne tous les CBZ et détecte ceux contenant des fichiers
non-JPG :\
résultat dans `unattended_cbz.txt`.

## 🔧 Étape 2 --- Optimisation

### ✔ Conversion des images

Les images suivantes sont converties : - `.png`, `.bmp`, `.webp` →
`.jpg` via ImageMagick

### ✔ Zéro perte de données

Tous les fichiers non-images sont conservés : - `ComicInfo.xml` -
`.txt`, `.json`, `.nfo` - dossiers internes - `.jpg` déjà présents

## 🗃️ CSV de suivi

Fichier : **cbz_status.csv**

Colonnes :

  ------------------------------------------------------------------------------
  cbz       original_size     compression_date     status     optimized_size
  --------- ----------------- -------------------- ---------- ------------------
  chemin    octets            datetime             "",        taille optimisée
  complet                                          ongoing,   
                                                   success,   
                                                   fail       

  ------------------------------------------------------------------------------

## ⚙️ Logique du script

1.  Lire `unattended_cbz.txt`
2.  Mettre `status = ongoing`
3.  Extraire le CBZ dans un dossier temporaire
4.  Convertir les images non-JPG en JPG
5.  Recréer le CBZ optimisé en **NoCompression (STORE)**
6.  Comparer les tailles
7.  Remplacer seulement si gain
8.  Mettre à jour le CSV en continu

## ⚡ Multithreading

Avec PowerShell 7 :

``` powershell
ForEach-Object -Parallel { ... } -ThrottleLimit 4
```

→ 4 CBZ traités en parallèle.

## 🧪 Deux versions

### 1. Version réelle

-   conversion, reconstruction, remplacement conditionnel

### 2. Version simulation

-   aucune modification
-   simule un gain ou une perte
-   met à jour le CSV

Scripts fournis : - `optimize_cbz_multithread_real.ps1` -
`optimize_cbz_multithread_simulation.ps1`

## 📦 Résultat

-   CBZ plus légers
-   aucune donnée perdue
-   conversions homogènes en JPG
-   compatibilité parfaite avec lecteurs CBZ
-   suivi CSV + reprise automatique
-   traitement rapide en parallèle
