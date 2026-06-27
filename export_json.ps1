# export_json.ps1 — Export AUTONOME : lit les classements calculés par Excel
# Usage: powershell -ExecutionPolicy Bypass -File export_json.ps1

$concPath = "g:\Mon Drive\CDM2026\CDM2026_Concours.xlsm"
$xlsmPath = "g:\Mon Drive\CDM2026\CDM2026_ClassementsAdditionnels.xlsm"
$jsonPath = "g:\Mon Drive\CDM2026\web\data.json"

Write-Host "=== Export CDM2026 (Classements Directs Excel) -> data.json ==="

$excel = $null
$srcWb = $null
$clWb = $null
$createdNewExcel = $false
$openedConcours = $false
$openedAdditionnels = $false

try {
    # Essayer de se lier à l'instance Excel active (déjà ouverte par l'utilisateur)
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Host "Connexion a l'instance Excel active reussie."
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $createdNewExcel = $true
        Write-Host "Creation d'une nouvelle instance Excel."
    }

    # === 1. Lire le Classement Général ===
    Write-Host "1. Lecture du Classement Général depuis CDM2026_Concours..."
    # Chercher si le classeur est déjà ouvert dans Excel
    foreach ($wb in $excel.Workbooks) {
        if ($wb.Name -eq "CDM2026_Concours.xlsm") {
            $srcWb = $wb
            break
        }
    }
    if ($null -eq $srcWb) {
        $srcWb = $excel.Workbooks.Open($concPath, $false, $true)
        $openedConcours = $true
    }
    $genWs = $srcWb.Worksheets.Item("Classement")
    
    $classementGeneral = @()
    for ($r = 7; $r -le 21; $r++) {
        $name = $genWs.Cells.Item($r, 4).Text
        if ($name -ne $null -and $name -ne "") {
            $rank = $genWs.Cells.Item($r, 3).Text
            $team = $genWs.Cells.Item($r, 5).Text
            $evol = $genWs.Cells.Item($r, 6).Text
            $points = $genWs.Cells.Item($r, 8).Text
            $fiabilite = $genWs.Cells.Item($r, 9).Text
            $detail = $genWs.Cells.Item($r, 10).Text
            $toutPile = $genWs.Cells.Item($r, 11).Text
            $maxPtsMatch = $genWs.Cells.Item($r, 12).Text
            $serieEnCours = $genWs.Cells.Item($r, 13).Text
            $meilleureSerie = $genWs.Cells.Item($r, 14).Text
            $roi = $genWs.Cells.Item($r, 16).Text
            
            $classementGeneral += [PSCustomObject]@{
                rang = if ($rank) { $rank } else { "-" }
                nom = $name
                equipe = if ($team) { $team } else { "" }
                evol = if ($evol) { $evol } else { "" }
                points = if ($points) { [double]$points.Replace(",", ".") } else { 0.0 }
                fiabilite = if ($fiabilite) { $fiabilite } else { "" }
                detail = if ($detail) { $detail } else { "" }
                toutPile = if ($toutPile) { [int]$toutPile } else { 0 }
                maxPtsMatch = if ($maxPtsMatch) { [double]$maxPtsMatch.Replace(",", ".") } else { 0.0 }
                serieEnCours = if ($serieEnCours) { [int]$serieEnCours } else { 0 }
                meilleureSerie = if ($meilleureSerie) { [int]$meilleureSerie } else { 0 }
                roi = if ($roi) { $roi } else { "" }
            }
        }
    }
    $eqWs = $srcWb.Worksheets.Item("ClassementEQ")
    $classementEquipes = @()
    # Lignes 7 à 11 pour les 5 équipes
    for ($r = 7; $r -le 11; $r++) {
        $team = $eqWs.Cells.Item($r, 4).Text
        if ($team -ne $null -and $team -ne "") {
            $rank = $eqWs.Cells.Item($r, 3).Text
            $nbrJoueurs = $eqWs.Cells.Item($r, 5).Text
            $evol = $eqWs.Cells.Item($r, 6).Text
            $moyPoints = $eqWs.Cells.Item($r, 8).Text
            $fiabilite = $eqWs.Cells.Item($r, 9).Text
            $detail = $eqWs.Cells.Item($r, 10).Text
            $moyToutPile = $eqWs.Cells.Item($r, 11).Text
            $ptsMax = $eqWs.Cells.Item($r, 12).Text
            $ptsMin = $eqWs.Cells.Item($r, 13).Text
            $bestClass = $eqWs.Cells.Item($r, 14).Text
            $worstClass = $eqWs.Cells.Item($r, 15).Text
            
            $classementEquipes += [PSCustomObject]@{
                rang = if ($rank) { $rank } else { "-" }
                nom = $team
                nbrJoueurs = if ($nbrJoueurs) { [int]$nbrJoueurs } else { 0 }
                evol = if ($evol) { $evol } else { "" }
                points = if ($moyPoints) { [double]$moyPoints.Replace(",", ".") } else { 0.0 }
                fiabilite = if ($fiabilite) { $fiabilite } else { "" }
                detail = if ($detail) { $detail } else { "" }
                moyToutPile = if ($moyToutPile) { [double]$moyToutPile.Replace(",", ".") } else { 0.0 }
                ptsMax = if ($ptsMax) { [double]$ptsMax.Replace(",", ".") } else { 0.0 }
                ptsMin = if ($ptsMin) { [double]$ptsMin.Replace(",", ".") } else { 0.0 }
                bestClass = if ($bestClass) { [int]$bestClass } else { 0 }
                worstClass = if ($worstClass) { [int]$worstClass } else { 0 }
            }
        }
    }
    Write-Host "   -> Lu $($( $classementEquipes.Count )) lignes de classement par équipes."

    if ($openedConcours) {
        $srcWb.Close($false)
        $srcWb = $null
    }

    # === 2. Lire les classements et détails du classeur additionnel ===
    $maillotVert = @()
    $maillotPois = @()
    $journees = @()

    if (Test-Path $xlsmPath) {
        Write-Host "2. Lecture du classeur additionnel CDM2026_ClassementsAdditionnels..."
        # Chercher si le classeur est déjà ouvert dans Excel
        foreach ($wb in $excel.Workbooks) {
            if ($wb.Name -eq "CDM2026_ClassementsAdditionnels.xlsm") {
                $clWb = $wb
                break
            }
        }
        if ($null -eq $clWb) {
            $clWb = $excel.Workbooks.Open($xlsmPath, $false, $true)
            $openedAdditionnels = $true
        }

        # 2a. Profils de journées
        Write-Host "   Lecture des profils de journées..."
        $profiles = @{}
        $jourWs = $clWb.Worksheets.Item("Journées")
        $row = 2
        while ($jourWs.Cells.Item($row, 1).Value2 -ne $null -and $jourWs.Cells.Item($row, 1).Value2 -ne "" -and $jourWs.Cells.Item($row, 1).Value2 -ne 0) {
            $dVal = $jourWs.Cells.Item($row, 1).Value2
            if ($dVal -is [double] -or $dVal -is [int]) {
                $dateKey = [DateTime]::FromOADate($dVal).ToString("dd/MM/yyyy")
                $prof = [string]$jourWs.Cells.Item($row, 2).Value2
                if ($prof -eq "" -or $prof -eq $null) { $prof = "Plaine" }
                $profiles[$dateKey] = $prof
            }
            $row++
        }

        # 2b. Maillot Vert & Maillot à Pois (onglet Classements, lignes 6 à 20)
        Write-Host "   Lecture des classements Maillot Vert et Maillot à Pois..."
        $classWs = $clWb.Worksheets.Item("Classements")
        for ($r = 6; $r -le 20; $r++) {
            # Maillot Vert
            $mvPlayer = $classWs.Cells.Item($r, 2).Value2
            if ($mvPlayer -ne $null -and $mvPlayer -ne "") {
                $mvRank = $classWs.Cells.Item($r, 1).Value2
                $mvPts = $classWs.Cells.Item($r, 3).Value2
                $maillotVert += [PSCustomObject]@{
                    rang = [int]$mvRank
                    joueur = [string]$mvPlayer
                    points = [double]$mvPts
                }
            }
            
            # Maillot à Pois
            $mapPlayer = $classWs.Cells.Item($r, 5).Value2
            if ($mapPlayer -ne $null -and $mapPlayer -ne "") {
                $mapRank = $classWs.Cells.Item($r, 4).Value2
                $mapPts = $classWs.Cells.Item($r, 6).Value2
                $maillotPois += [PSCustomObject]@{
                    rang = [int]$mapRank
                    joueur = [string]$mapPlayer
                    points = [double]$mapPts
                }
            }
        }

        # 2c. Détails Matchs (onglet Détails Matchs)
        Write-Host "   Lecture de l'onglet Détails Matchs..."
        $detWs = $clWb.Worksheets.Item("Détails Matchs")
        $maxRow = $detWs.UsedRange.Rows.Count
        $r = 1
        while ($r -le $maxRow) {
            $cellVal = $detWs.Cells.Item($r, 1).Value2
            if ($cellVal -ne $null -and ([string]$cellVal).StartsWith("JOURN")) {
                # Entête de journée : "JOURNÉE DU 11/06/2026"
                $dateStr = [string]$cellVal
                if ($dateStr -match '(\d{2}/\d{2}/\d{4})') {
                    $dateStr = $Matches[1]
                } else {
                    $dateStr = $dateStr.Replace("JOURNÉE DU ", "").Trim()
                }

                # Ligne suivante = sous-entête (Rang | Joueur | Match 1 | ... | Total)
                $r++
                
                # Compter le nombre de colonnes actives
                $colCount = 1
                while ($detWs.Cells.Item($r, $colCount).Value2 -ne $null -and $detWs.Cells.Item($r, $colCount).Value2 -ne "") {
                    $colCount++
                }
                $colCount-- # index de la dernière colonne (Total)
                
                $matchCount = $colCount - 3
                $matchLabels = @()
                for ($c = 3; $c -le ($colCount - 1); $c++) {
                    $matchLabels += [string]$detWs.Cells.Item($r, $c).Value2
                }

                # Lire les 15 joueurs
                $classement = @()
                $r++
                $pRowStart = $r
                # On boucle sur les lignes jusqu'à ce qu'on ait lu 15 joueurs ou trouvé une ligne vide
                for ($p = 1; $p -le 15; $p++) {
                    $pName = $detWs.Cells.Item($r, 2).Value2
                    if ($pName -ne $null -and $pName -ne "") {
                        $rank = $detWs.Cells.Item($r, 1).Value2
                        $matchPts = @()
                        for ($c = 3; $c -le ($colCount - 1); $c++) {
                            $matchPts += [double]$detWs.Cells.Item($r, $c).Value2
                        }
                        $total = $detWs.Cells.Item($r, $colCount).Value2

                        $classement += [PSCustomObject]@{
                            rang = [int]$rank
                            joueur = [string]$pName
                            points = $matchPts
                            total = [double]$total
                        }
                    }
                    $r++
                }

                # Profil de la journée
                $profil = "Plaine"
                if ($profiles.ContainsKey($dateStr)) {
                    $profil = $profiles[$dateStr]
                }

                $journees += [PSCustomObject]@{
                    date = $dateStr
                    profil = $profil
                    matchs = $matchLabels
                    classement = $classement
                }
            } else {
                $r++
            }
        }
        
        if ($openedAdditionnels) {
            $clWb.Close($false)
            $clWb = $null
        }
        Write-Host "   -> Lu $($maillotVert.Count) joueurs MV, $($maillotPois.Count) joueurs MAP, $($journees.Count) journées."
    } else {
        Write-Host "   [ATTENTION] Fichier CDM2026_ClassementsAdditionnels.xlsm non trouvé. Les classements Vert, Pois et Journées seront vides dans le fallback."
    }

    # === 3. Écrire le fichier JSON ===
    Write-Host "3. Écriture du fichier data.json..."
    $output = [PSCustomObject]@{
        lastUpdate = (Get-Date).ToString("dd/MM/yyyy HH:mm")
        classementGeneral = $classementGeneral
        classementEquipes = $classementEquipes
        maillotVert = $maillotVert
        maillotPois = $maillotPois
        journees = $journees
    }
    
    $json = $output | ConvertTo-Json -Depth 10 -Compress:$false
    [System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.UTF8Encoding]::new($false))
    
    # Écriture de data.js pour contourner le blocage CORS en local (file://)
    Write-Host "4. Écriture du fichier data.js..."
    $jsPath = Join-Path (Split-Path $jsonPath) "data.js"
    $jsContent = "const DATA_FALLBACK = " + $json + ";"
    [System.IO.File]::WriteAllText($jsPath, $jsContent, [System.Text.UTF8Encoding]::new($false))
    
    Write-Host "=== Done! ==="

} catch {
    $errText = "ERROR at $(Get-Date): $_`r`n" + ($Error[0] | Out-String)
    try {
        [System.IO.File]::WriteAllText("g:\Mon Drive\CDM2026\web\export_log.txt", $errText, [System.Text.UTF8Encoding]::new($false))
    } catch {}
    Write-Host "ERROR: $_"
    $Error[0] | Format-List -Force
} finally {
    if ($openedConcours -and $srcWb) { $srcWb.Close($false) }
    if ($openedAdditionnels -and $clWb) { $clWb.Close($false) }
    if ($createdNewExcel -and $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
