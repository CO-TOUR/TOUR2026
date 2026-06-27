Sub ShowMessage(msg As String, style As Integer, title As String)
    If Application.UserControl Then
        MsgBox msg, style, title
    Else
        Debug.Print title & ": " & msg
    End If
End Sub

Sub RefreshData()
    Dim sourcePath As String
    sourcePath = ThisWorkbook.Path & "\CDM2026_Concours.xlsm"
    
    ' Check if source file exists
    If Dir(sourcePath) = "" Then
        ShowMessage "Le fichier source '" & sourcePath & "' est introuvable. Veuillez vous assurer qu'il se trouve dans le mÃªme dossier.", vbCritical, "Erreur"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim sourceWb As Workbook
    On Error Resume Next
    Set sourceWb = Workbooks.Open(Filename:=sourcePath, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo 0
    
    If sourceWb Is Nothing Then
        ShowMessage "Impossible d'ouvrir le fichier source. Il est peut-Ãªtre verrouillÃ© ou corrompu.", vbCritical, "Erreur"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Dim sourceWs As Worksheet
    On Error Resume Next
    Set sourceWs = sourceWb.Sheets("Points")
    On Error GoTo 0
    
    If sourceWs Is Nothing Then
        ShowMessage "L'onglet 'Points' est introuvable dans le fichier source.", vbCritical, "Erreur"
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' 1. Read pronostiqueurs
    Dim NumPlayers As Integer
    NumPlayers = 0
    Dim col As Integer
    col = 15
    Do While sourceWs.Cells(4, col).Value <> ""
        NumPlayers = NumPlayers + 1
        col = col + 3
    Loop
    
    If NumPlayers = 0 Then
        ShowMessage "Aucun pronostiqueur trouvÃ© dans l'onglet 'Points'.", vbCritical, "Erreur"
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Dim Players() As String
    ReDim Players(1 To NumPlayers)
    col = 15
    Dim P As Integer
    For P = 1 To NumPlayers
        Players(P) = sourceWs.Cells(4, col).Value
        col = col + 3
    Next P
    
    ' 2. Read unique dates for JournÃ©es sheet update
    ' We scan matches 1 to 104 (rows 6 to 109)
    Dim uniqueDates As Object
    Set uniqueDates = CreateObject("Scripting.Dictionary")
    
    Dim r As Long
    Dim dateVal As Variant
    For r = 6 To 109
        dateVal = sourceWs.Cells(r, 1).Value
        If Not IsEmpty(dateVal) And dateVal <> "" Then
            Dim dVal As Date
            If IsDate(dateVal) Then
                dVal = CDate(Int(CDbl(CDate(dateVal))))
                uniqueDates(dVal) = True
            ElseIf IsNumeric(dateVal) And CDbl(dateVal) > 0 Then
                dVal = CDate(Int(CDbl(dateVal)))
                uniqueDates(dVal) = True
            End If
        End If
    Next r
    
    ' Update JournÃ©es sheet
    Dim activeWb As Workbook
    Set activeWb = ThisWorkbook
    Dim jourWs As Worksheet
    Set jourWs = activeWb.Sheets("JournÃ©es")
    
    ' Read existing dates/profiles
    Dim existingProfiles As Object
    Set existingProfiles = CreateObject("Scripting.Dictionary")
    r = 2
    Do While jourWs.Cells(r, 1).Value <> ""
        Dim tempDate As Date
        tempDate = CDate(jourWs.Cells(r, 1).Value)
        existingProfiles(tempDate) = jourWs.Cells(r, 2).Value
        r = r + 1
    Loop
    
    ' Write back combined dates, maintaining profiles
    jourWs.Range("A2:B200").ClearContents
    Dim key As Variant
    Dim idx As Long
    idx = 2
    
    ' Sort dates
    Dim sortedDates() As Date
    Dim numDates As Long
    numDates = uniqueDates.Count
    
    If numDates = 0 Then
        ShowMessage "Aucune date de match trouvÃ©e dans l'onglet 'Points'. VÃ©rifiez que la colonne A contient des dates valides (lignes 6 Ã  109).", vbCritical, "Erreur"
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ReDim sortedDates(1 To numDates)
    Dim dIdx As Long
    dIdx = 1
    For Each key In uniqueDates.Keys
        sortedDates(dIdx) = CDate(key)
        dIdx = dIdx + 1
    Next key
    
    ' Bubble sort dates
    Dim i As Long, j As Long
    Dim tempD As Date
    For i = 1 To numDates - 1
        For j = i + 1 To numDates
            If sortedDates(i) > sortedDates(j) Then
                tempD = sortedDates(i)
                sortedDates(i) = sortedDates(j)
                sortedDates(j) = tempD
            End If
        Next j
    Next i
    
    For idx = 2 To numDates + 1
        Dim currDate As Date
        currDate = sortedDates(idx - 1)
        jourWs.Cells(idx, 1).Value = currDate
        jourWs.Cells(idx, 1).NumberFormat = "dd/mm/yyyy"
        If existingProfiles.Exists(currDate) Then
            jourWs.Cells(idx, 2).Value = existingProfiles(currDate)
        Else
            jourWs.Cells(idx, 2).Value = "Plaine" ' Default
        End If
    Next idx
    
    ' Re-apply Data Validation to JournÃ©es Column B
    With jourWs.Range("B2:B" & (numDates + 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Plaine,VallonnÃ©,Montagne"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Reload updated profiles
    Set existingProfiles = CreateObject("Scripting.Dictionary")
    For idx = 2 To numDates + 1
        existingProfiles(CDate(jourWs.Cells(idx, 1).Value)) = jourWs.Cells(idx, 2).Value
    Next idx
    
    ' 3. Load Configurations
    Dim configWs As Worksheet
    Set configWs = activeWb.Sheets("Configuration")
    Dim ConfigMV As Object, ConfigMAP As Object
    Set ConfigMV = CreateObject("Scripting.Dictionary")
    Set ConfigMAP = CreateObject("Scripting.Dictionary")
    
    Dim profName As String
    For r = 2 To 4
        profName = configWs.Cells(r, 1).Value
        Dim mvPts(1 To 5) As Double
        Dim mapPts(1 To 5) As Double
        Dim k As Integer
        For k = 1 To 5
            mvPts(k) = Val(configWs.Cells(r, 1 + k).Value)
            mapPts(k) = Val(configWs.Cells(r, 6 + k).Value)
        Next k
        ConfigMV(profName) = mvPts
        ConfigMAP(profName) = mapPts
    Next r
    
    ' 4. Sum Match Points per Day and Player
    ' Struct: PlayerDailyPoints(Date_PlayerIndex)
    ' DayHasMatches(Date)
    Dim PlayerDailyPoints As Object
    Set PlayerDailyPoints = CreateObject("Scripting.Dictionary")
    Dim DayHasMatches As Object
    Set DayHasMatches = CreateObject("Scripting.Dictionary")
    
    ' Also store per-match details: MatchInfo(matchIndex) = "date|team1|team2|score1|score2"
    ' PlayerMatchPts(matchIndex, playerIndex) = points
    Dim MatchInfo As Object
    Set MatchInfo = CreateObject("Scripting.Dictionary")
    Dim PlayerMatchPts As Object
    Set PlayerMatchPts = CreateObject("Scripting.Dictionary")
    Dim MatchDates As Object
    Set MatchDates = CreateObject("Scripting.Dictionary")
    Dim matchCount As Long
    matchCount = 0
    
    ' Initialize PlayerDailyPoints keys for all played matches
    For r = 6 To 109
        ' Check if match has scores (played)
        If sourceWs.Cells(r, 5).Value <> "" And sourceWs.Cells(r, 6).Value <> "" Then
            dateVal = sourceWs.Cells(r, 1).Value
            Dim dateOk As Boolean
            dateOk = False
            Dim dMatch As Date
            If IsDate(dateVal) Then
                dMatch = CDate(Int(CDbl(CDate(dateVal))))
                dateOk = True
            ElseIf IsNumeric(dateVal) And CDbl(dateVal) > 0 Then
                dMatch = CDate(Int(CDbl(dateVal)))
                dateOk = True
            End If
            
            If dateOk Then
                DayHasMatches(dMatch) = True
                matchCount = matchCount + 1
                
                ' Store match info
                Dim team1 As String, team2 As String
                Dim score1 As String, score2 As String
                team1 = CStr(sourceWs.Cells(r, 2).Value)
                team2 = CStr(sourceWs.Cells(r, 3).Value)
                score1 = CStr(sourceWs.Cells(r, 5).Value)
                score2 = CStr(sourceWs.Cells(r, 6).Value)
                MatchInfo(matchCount) = Format(dMatch, "dd/mm/yyyy") & "|" & team1 & "|" & team2 & "|" & score1 & "|" & score2
                MatchDates(matchCount) = dMatch
                
                For P = 1 To NumPlayers
                    Dim ptsCol As Integer
                    ptsCol = 15 + (P - 1) * 3 + 2
                    Dim pts As Double
                    pts = Val(sourceWs.Cells(r, ptsCol).Value)
                    
                    Dim dictKey As String
                    dictKey = CStr(dMatch) & "_" & CStr(P)
                    PlayerDailyPoints(dictKey) = PlayerDailyPoints(dictKey) + pts
                    
                    ' Store per-match points
                    Dim matchPtsKey As String
                    matchPtsKey = CStr(matchCount) & "_" & CStr(P)
                    PlayerMatchPts(matchPtsKey) = pts
                Next P
            End If
        End If
    Next r
    
    ' Close source workbook
    sourceWb.Close SaveChanges:=False
    
    ' 5. Process rankings and write details
    Dim detailWs As Worksheet
    Set detailWs = activeWb.Sheets("DÃ©tails Calculs")
    detailWs.Cells.Clear
    
    ' Header
    detailWs.Cells(1, 1).Value = "Date"
    detailWs.Cells(1, 2).Value = "Joueur"
    detailWs.Cells(1, 3).Value = "Points Matchs"
    detailWs.Cells(1, 4).Value = "Rang JournÃ©e"
    detailWs.Cells(1, 5).Value = "Points Maillot Vert"
    detailWs.Cells(1, 6).Value = "Points Maillot Pois"
    
    Dim detailRow As Long
    detailRow = 2
    
    ' Arrays to sum totals
    Dim PlayerTotalMV() As Double
    Dim PlayerTotalMAP() As Double
    ReDim PlayerTotalMV(1 To NumPlayers)
    ReDim PlayerTotalMAP(1 To NumPlayers)
    
    ' For each day that has played matches
    For idx = 1 To numDates
        Dim currD As Date
        currD = sortedDates(idx)
        
        If DayHasMatches.Exists(currD) Then
            Dim prof As String
            If existingProfiles.Exists(currD) Then
                prof = existingProfiles(currD)
            Else
                prof = "Plaine"
            End If
            
            Dim mvPtsArray As Variant
            Dim mapPtsArray As Variant
            mvPtsArray = ConfigMV(prof)
            mapPtsArray = ConfigMAP(prof)
            
            ' Collect points
            Dim dailyPts() As Double
            ReDim dailyPts(1 To NumPlayers)
            For P = 1 To NumPlayers
                Dim kStr As String
                kStr = CStr(currD) & "_" & CStr(P)
                dailyPts(P) = PlayerDailyPoints(kStr)
            Next P
            
            ' Rank players
            Dim dailyRanks() As Integer
            ReDim dailyRanks(1 To NumPlayers)
            For P = 1 To NumPlayers
                Dim rank As Integer
                rank = 1
                For j = 1 To NumPlayers
                    If dailyPts(j) > dailyPts(P) Then
                        rank = rank + 1
                    End If
                Next j
                dailyRanks(P) = rank
            Next P
            
            ' Award points
            For P = 1 To NumPlayers
                Dim rnk As Integer
                rnk = dailyRanks(P)
                
                Dim mvEarned As Double
                Dim mapEarned As Double
                mvEarned = 0
                mapEarned = 0
                
                If rnk >= 1 And rnk <= 5 Then
                    mvEarned = mvPtsArray(rnk)
                    mapEarned = mapPtsArray(rnk)
                End If
                
                PlayerTotalMV(P) = PlayerTotalMV(P) + mvEarned
                PlayerTotalMAP(P) = PlayerTotalMAP(P) + mapEarned
                
                ' Write details
                detailWs.Cells(detailRow, 1).Value = currD
                detailWs.Cells(detailRow, 1).NumberFormat = "dd/mm/yyyy"
                detailWs.Cells(detailRow, 2).Value = Players(P)
                detailWs.Cells(detailRow, 3).Value = dailyPts(P)
                detailWs.Cells(detailRow, 4).Value = rnk
                detailWs.Cells(detailRow, 5).Value = mvEarned
                detailWs.Cells(detailRow, 6).Value = mapEarned
                detailRow = detailRow + 1
            Next P
        End If
    Next idx
    
    ' 6. Write Consolidated Leaderboards
    Dim classWs As Worksheet
    Set classWs = activeWb.Sheets("Classements")
    
    ' Clear existing data in Classements
    classWs.Range("A6:F30").ClearContents
    
    ' Sort Green Jersey
    Dim mvSortedPlayers() As String
    Dim mvSortedPoints() As Double
    ReDim mvSortedPlayers(1 To NumPlayers)
    ReDim mvSortedPoints(1 To NumPlayers)
    For P = 1 To NumPlayers
        mvSortedPlayers(P) = Players(P)
        mvSortedPoints(P) = PlayerTotalMV(P)
    Next P
    
    ' Bubble sort MV
    Dim tempP As String, tempPt As Double
    For i = 1 To NumPlayers - 1
        For j = i + 1 To NumPlayers
            If mvSortedPoints(i) < mvSortedPoints(j) Then
                tempPt = mvSortedPoints(i)
                mvSortedPoints(i) = mvSortedPoints(j)
                mvSortedPoints(j) = tempPt
                tempP = mvSortedPlayers(i)
                mvSortedPlayers(i) = mvSortedPlayers(j)
                mvSortedPlayers(j) = tempP
            End If
        Next j
    Next i
    
    ' Sort Polka Dot Jersey
    Dim mapSortedPlayers() As String
    Dim mapSortedPoints() As Double
    ReDim mapSortedPlayers(1 To NumPlayers)
    ReDim mapSortedPoints(1 To NumPlayers)
    For P = 1 To NumPlayers
        mapSortedPlayers(P) = Players(P)
        mapSortedPoints(P) = PlayerTotalMAP(P)
    Next P
    
    ' Bubble sort MAP
    For i = 1 To NumPlayers - 1
        For j = i + 1 To NumPlayers
            If mapSortedPoints(i) < mapSortedPoints(j) Then
                tempPt = mapSortedPoints(i)
                mapSortedPoints(i) = mapSortedPoints(j)
                mapSortedPoints(j) = tempPt
                tempP = mapSortedPlayers(i)
                mapSortedPlayers(i) = mapSortedPlayers(j)
                mapSortedPlayers(j) = tempP
            End If
        Next j
    Next i
    
    ' Write Green Jersey to Columns B & C
    For P = 1 To NumPlayers
        classWs.Cells(5 + P, 2).Value = mvSortedPlayers(P)
        classWs.Cells(5 + P, 3).Value = mvSortedPoints(P)
    Next P
    
    ' Write Polka Dot Jersey to Columns E & F
    For P = 1 To NumPlayers
        classWs.Cells(5 + P, 5).Value = mapSortedPlayers(P)
        classWs.Cells(5 + P, 6).Value = mapSortedPoints(P)
    Next P
    
    ' Add Ranks
    For P = 1 To NumPlayers
        classWs.Cells(5 + P, 1).Value = P
        classWs.Cells(5 + P, 4).Value = P
    Next P
    
    ' 7. Write DÃ©tails Matchs tab
    Dim matchWs As Worksheet
    Set matchWs = activeWb.Sheets("DÃ©tails Matchs")
    matchWs.Cells.Clear
    
    Dim mRow As Long
    mRow = 1
    
    ' Process each day
    For idx = 1 To numDates
        Dim dayD As Date
        dayD = sortedDates(idx)
        
        If DayHasMatches.Exists(dayD) Then
            ' --- Day Header ---
            Dim dayStr As String
            dayStr = Format(dayD, "dd/mm/yyyy")
            
            ' Calculate daily total per player for sorting
            Dim dayTotals() As Double
            Dim playerOrder() As Integer
            ReDim dayTotals(1 To NumPlayers)
            ReDim playerOrder(1 To NumPlayers)
            For P = 1 To NumPlayers
                Dim kk As String
                kk = CStr(dayD) & "_" & CStr(P)
                If PlayerDailyPoints.Exists(kk) Then
                    dayTotals(P) = PlayerDailyPoints(kk)
                Else
                    dayTotals(P) = 0
                End If
                playerOrder(P) = P
            Next P
            
            ' Sort players by daily total (descending)
            Dim ii As Long, jj As Long
            Dim tmpOrd As Integer, tmpDbl As Double
            For ii = 1 To NumPlayers - 1
                For jj = ii + 1 To NumPlayers
                    If dayTotals(playerOrder(ii)) < dayTotals(playerOrder(jj)) Then
                        tmpOrd = playerOrder(ii)
                        playerOrder(ii) = playerOrder(jj)
                        playerOrder(jj) = tmpOrd
                    End If
                Next jj
            Next ii
            
            ' Collect match indices for this day
            Dim dayMatchIndices() As Long
            Dim dayMatchCount As Long
            dayMatchCount = 0
            Dim m As Long
            For m = 1 To matchCount
                If MatchDates(m) = dayD Then
                    dayMatchCount = dayMatchCount + 1
                End If
            Next m
            ReDim dayMatchIndices(1 To dayMatchCount)
            Dim mi As Long
            mi = 1
            For m = 1 To matchCount
                If MatchDates(m) = dayD Then
                    dayMatchIndices(mi) = m
                    mi = mi + 1
                End If
            Next m
            
            ' --- Write day header row ---
            matchWs.Cells(mRow, 1).Value = "JOURNÃ‰E DU " & UCase(dayStr)
            matchWs.Range(matchWs.Cells(mRow, 1), matchWs.Cells(mRow, 2 + dayMatchCount)).Merge
            matchWs.Cells(mRow, 1).Font.Bold = True
            matchWs.Cells(mRow, 1).Font.Size = 11
            matchWs.Cells(mRow, 1).Font.Name = "Segoe UI"
            matchWs.Cells(mRow, 1).Font.Color = RGB(255, 255, 255)
            matchWs.Cells(mRow, 1).Interior.Color = RGB(32, 80, 62)
            matchWs.Cells(mRow, 1).HorizontalAlignment = xlCenter
            mRow = mRow + 1
            
            ' --- Sub-header: Rang | Joueur | Match1 | Match2 | ... | Total ---
            matchWs.Cells(mRow, 1).Value = "Rang"
            matchWs.Cells(mRow, 2).Value = "Joueur"
            Dim mc As Long
            For mc = 1 To dayMatchCount
                Dim mParts() As String
                mParts = Split(MatchInfo(dayMatchIndices(mc)), "|")
                ' mParts: 0=date, 1=team1, 2=team2, 3=score1, 4=score2
                matchWs.Cells(mRow, 2 + mc).Value = mParts(1) & " " & mParts(3) & "-" & mParts(4) & " " & mParts(2)
            Next mc
            matchWs.Cells(mRow, 2 + dayMatchCount + 1).Value = "Total"
            
            ' Format sub-header
            Dim hdrRange As Range
            Set hdrRange = matchWs.Range(matchWs.Cells(mRow, 1), matchWs.Cells(mRow, 2 + dayMatchCount + 1))
            hdrRange.Font.Bold = True
            hdrRange.Font.Size = 9
            hdrRange.Font.Name = "Segoe UI"
            hdrRange.Interior.Color = RGB(230, 240, 230)
            hdrRange.HorizontalAlignment = xlCenter
            hdrRange.Borders.LineStyle = xlContinuous
            hdrRange.Borders.Weight = xlThin
            mRow = mRow + 1
            
            ' --- Data rows: sorted players ---
            For P = 1 To NumPlayers
                Dim pIdx As Integer
                pIdx = playerOrder(P)
                
                ' Rank
                Dim pRank As Integer
                pRank = 1
                For jj = 1 To NumPlayers
                    If dayTotals(playerOrder(jj)) > dayTotals(pIdx) Then
                        pRank = pRank + 1
                    End If
                Next jj
                
                matchWs.Cells(mRow, 1).Value = pRank
                matchWs.Cells(mRow, 2).Value = Players(pIdx)
                
                Dim rowTotal As Double
                rowTotal = 0
                For mc = 1 To dayMatchCount
                    Dim mpKey As String
                    mpKey = CStr(dayMatchIndices(mc)) & "_" & CStr(pIdx)
                    Dim mpVal As Double
                    mpVal = 0
                    If PlayerMatchPts.Exists(mpKey) Then mpVal = PlayerMatchPts(mpKey)
                    matchWs.Cells(mRow, 2 + mc).Value = mpVal
                    rowTotal = rowTotal + mpVal
                Next mc
                matchWs.Cells(mRow, 2 + dayMatchCount + 1).Value = rowTotal
                
                ' Formatting
                Dim dataRange As Range
                Set dataRange = matchWs.Range(matchWs.Cells(mRow, 1), matchWs.Cells(mRow, 2 + dayMatchCount + 1))
                dataRange.Font.Name = "Segoe UI"
                dataRange.Font.Size = 9
                dataRange.Borders.LineStyle = xlContinuous
                dataRange.Borders.Weight = xlThin
                matchWs.Cells(mRow, 1).HorizontalAlignment = xlCenter
                matchWs.Range(matchWs.Cells(mRow, 3), matchWs.Cells(mRow, 2 + dayMatchCount + 1)).HorizontalAlignment = xlCenter
                
                ' Highlight top 3
                If pRank = 1 Then
                    dataRange.Interior.Color = RGB(212, 175, 55)  ' Gold
                    dataRange.Font.Bold = True
                ElseIf pRank = 2 Then
                    dataRange.Interior.Color = RGB(192, 192, 192) ' Silver
                ElseIf pRank = 3 Then
                    dataRange.Interior.Color = RGB(205, 127, 50)  ' Bronze
                End If
                
                ' Bold the Total column
                matchWs.Cells(mRow, 2 + dayMatchCount + 1).Font.Bold = True
                
                mRow = mRow + 1
            Next P
            
            ' Blank row between days
            mRow = mRow + 1
        End If
    Next idx
    
    ' Auto-fit columns in DÃ©tails Matchs
    matchWs.Columns.AutoFit
    matchWs.Columns.Item(1).ColumnWidth = 6
    matchWs.Columns.Item(2).ColumnWidth = 18
    
    ' Update status
    classWs.Range("B3").Value = "DerniÃ¨re actualisation : " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ThisWorkbook.Save

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Executer le script PowerShell d'export de maniere synchrone (attente de la fin)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim psCmd As String
    psCmd = "powershell -ExecutionPolicy Bypass -File """ & ThisWorkbook.Path & "\web\export_json.ps1"""
    wsh.Run psCmd, 0, True
    On Error GoTo 0

    ShowMessage "Actualisation terminee avec succes !", vbInformation, "Succes"
End Sub