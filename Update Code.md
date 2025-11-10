# 1) Import Excel → Access : passer en set-based (plus de boucle ligne par ligne)

Le plus gros goulot c’est LoadTodayIntoAccess qui pousse 13 paramètres par ligne. Remplace-le par un INSERT … SELECT direct depuis la feuille Excel via le driver ACE. Pour ça, on normalise d’abord les en-têtes de la feuille (pour virer les variantes) puis on pousse tout en un seul SQL.

## 1.1 Normaliser les en-têtes Excel (une fois par run)
```vb
Private Sub NormalizeHeadersInSheet(ByVal ws As Worksheet)
    Dim h As Range, i As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        Set h = ws.Cells(1, i)
        Select Case UCase$(Trim$(h.Value & ""))
            Case "NO NTERVENANT":              h.Value = "NO INTERVENANT"
            Case "NO NTERVENANT GRP":          h.Value = "NO INTERVENANT GRP"
            Case "MONTANT VG (VALEUR DE GAGE-LTV)": h.Value = "Montant VG"
            ' Ajoute ici d'autres alias si besoin
        End Select
    Next i
End Sub


```

## 1.2 Import bulk sans boucle
```vb
Public Function FastLoadTodayIntoAccess(ByVal dbPath As String, ByVal ws As Worksheet) As Long
    Dim cn As Object: Set cn = GetConn(dbPath)
    Dim wbPath As String: wbPath = ws.Parent.FullName
    Dim sh As String: sh = ws.Name & "$"
    
    ' Déterminer le range (A1:... dernière colonne/ligne)
    Dim lastRow As Long, lastCol As Long, lastColLetter As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastColLetter = Split(ws.Cells(1, lastCol).Address(False, False), "$")(0)
    Dim xlRange As String: xlRange = "[" & sh & "A1:" & lastColLetter & lastRow & "]"
    
    ' Vider T_Today
    cn.Execute "DELETE FROM " & T_TODAY, , 128 ' adExecuteNoRecords
    
    ' Push set-based (HDR=YES = utiliser la ligne 1 comme en-têtes)
    Dim sql As String
    sql = "INSERT INTO " & T_TODAY & " (" & _
          "Booking, [Production Date], [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP], " & _
          "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], " & _
          "[Insuffisance VG], [Insuffisance AM], [Insuffisance SL]) " & _
          "SELECT Booking, [Production Date], [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP], " & _
          "       Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], " & _
          "       [Insuffisance VG], [Insuffisance AM], [Insuffisance SL] " & _
          "FROM [Excel 12.0 Xml;HDR=YES;IMEX=0;Database=" & wbPath & "]." & xlRange & ";"
    
    cn.Execute sql, , 128
    FastLoadTodayIntoAccess = cn.RecordsAffected
    cn.Close
End Function
```

**dans  macro principale :**
```vb
NormalizeHeadersInSheet ws
rowsImported = FastLoadTodayIntoAccess(dbPath, ws)
```

# 2) Normalisation : faire 1 seul UPDATE (au lieu d’une rafale)

Tu enchaînes 8 UPDATE. Fais-le en un seul set avec IIf(Is Null(...)) :
```vb
Private Sub NormalizeToday(ByVal dbPath As String)
    Dim cn As Object: Set cn = GetConn(dbPath)
    
    cn.Execute _
    "UPDATE " & T_TODAY & " SET " & _
    "  Booking = UCase(Trim(Booking))," & _
    "  [NO DOSSIER CREDIT] = Trim([NO DOSSIER CREDIT])," & _
    "  [Production Date] = IIf([Production Date] Is Null, Date(), [Production Date])," & _
    "  Limite = IIf(Limite Is Null, 0, Limite)," & _
    "  Consommation = IIf(Consommation Is Null, 0, Consommation)," & _
    "  [Montant VG] = IIf([Montant VG] Is Null, 0, [Montant VG])," & _
    "  [Montant AM] = IIf([Montant AM] Is Null, 0, [Montant AM])," & _
    "  [Montant SL] = IIf([Montant SL] Is Null, 0, [Montant SL])," & _
    "  [Insuffisance VG] = IIf([Insuffisance VG] Is Null, 0, [Insuffisance VG])," & _
    "  [Insuffisance AM] = IIf([Insuffisance AM] Is Null, 0, [Insuffisance AM])," & _
    "  [Insuffisance SL] = IIf([Insuffisance SL] Is Null, 0, [Insuffisance SL]);", , 128

    ' Dé-doublonnage plus sélectif avec jointure sur MAX(Production Date)
    cn.Execute _
    "DELETE a.* FROM " & T_TODAY & " AS a " & _
    "INNER JOIN (SELECT Booking, [NO DOSSIER CREDIT], MAX([Production Date]) AS MaxD " & _
    "           FROM " & T_TODAY & " GROUP BY Booking, [NO DOSSIER CREDIT]) AS m " & _
    "ON a.Booking=m.Booking AND a.[NO DOSSIER CREDIT]=m.[NO DOSSIER CREDIT] " & _
    "WHERE a.[Production Date] < m.MaxD;", , 128

    ' TEST_MODE : set unique (si actif)
    If TEST_MODE Then
        Dim dSim As Date: dSim = GetSimDate(dbPath)
        cn.Execute "UPDATE " & T_TODAY & " SET [Production Date]=" & SqlDate(dSim) & ";", , 128
    End If

    cn.Close
End Sub
```

# 3) Une seule connexion + moins d’allers/retours

Tu ouvres/fermes une connexion dans quasi chaque procédure. Ouvre une seule connexion et passe-la en paramètre (ou module-level). Ça supprime beaucoup de latence COM.

Exemple (squelette) :

```vb
Public Sub MAJ_Retards_Daily_3in1_NoNzIif2()
    Dim cn As Object: Set cn = GetConn(ThisWorkbook.Path & Application.PathSeparator & DB_NAME)
    ' ... passer cn à FastLoadTodayIntoAccess/NormalizeToday/AppendHistory/Process... (surcharges avec cn)
    cn.Close
End Sub


```

# 4) Règles métier : sortir la mise à jour de contexte des 3 updates par nature

Tes updates VG/AM/SL réécrivent à chaque fois les mêmes colonnes de contexte (Limite, Conso, Montants, Intervenants, ProdDate). Fais un unique UPDATE de contexte, puis des updates par nature sans ces colonnes.

## 4.1 Contexte (1 seul passage)

```vb
' A placer au debut de ProcessBusinessRules, apres l’insert des nouvelles cles
cn.Execute _
"UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
"ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
"SET f.Limite=t.Limite, f.Consommation=t.Consommation, " & _
"    f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], " & _
"    f.[Production Date]=t.[Production Date], " & _
"    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP];", , 128

```

## 4.2 Par nature (sans re-écrire le contexte)
```vb
' VG continue
cn.Execute _
"UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
"ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
"SET f.VG_NbJoursRetard = f.VG_NbJoursRetard + DateDiff('d', f.VG_LastUpdateDate, t.[Production Date]), " & _
"    f.VG_LastUpdateDate = t.[Production Date], " & _
"    f.VG_LastInsuffisance = t.[Insuffisance VG] " & _
"WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NOT NULL AND t.[Production Date] > f.VG_LastUpdateDate;", , 128


```
Fais pareil pour AM/SL. Tu supprimes ainsi 2× la réécriture des colonnes de contexte à chaque nature ⇒ moins d’I/O.

## 4.3 DepassementNA : calcule une seule fois (à la fin)
```vb
cn.Execute _
"UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
"ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
"SET f.DepassementNA = IIf(Abs(t.Consommation) > t.Limite, 'YES','NO');", , 128
```

# 5) BuildCloseBuffer : enlever DISTINCT, tout est déjà unique
```vb
INSERT INTO T_CloseBuf (Booking, [NO DOSSIER CREDIT])
SELECT f.Booking, f.[NO DOSSIER CREDIT]
FROM T_Flux AS f
LEFT JOIN T_Today AS t ON ...
WHERE ...


```

# 6) Enrichissement de la feuille : écrire par blocs (arrays), pas cellule par cellule

Remplace la boucle qui fait des .Value ligne par ligne par un chargement de la plage en tableau, puis écriture en une passe pour chaque colonne cible.

```vb
Private Sub EnrichSheetFromFlux_3in1(ByVal dbPath As String, ByVal ws As Worksheet)
    Dim r As Long, lastRow As Long
    Dim cNBJ_VG As Long, cDEB_VG As Long, cFIN_VG As Long
    Dim cNBJ_AM As Long, cDEB_AM As Long, cFIN_AM As Long
    Dim cNBJ_SL As Long, cDEB_SL As Long, cFIN_SL As Long, cNA As Long
    
    cNBJ_VG = EnsureColumn(ws, COL_NBJ_VG): cDEB_VG = EnsureColumn(ws, COL_DEB_VG): cFIN_VG = EnsureColumn(ws, COL_FIN_VG)
    cNBJ_AM = EnsureColumn(ws, COL_NBJ_AM): cDEB_AM = EnsureColumn(ws, COL_DEB_AM): cFIN_AM = EnsureColumn(ws, COL_FIN_AM)
    cNBJ_SL = EnsureColumn(ws, COL_NBJ_SL): cDEB_SL = EnsureColumn(ws, COL_DEB_SL): cFIN_SL = EnsureColumn(ws, COL_FIN_SL)
    cNA = EnsureColumn(ws, COL_FLAG_NA)

    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    MapCol map, ws, COL_BOOKING
    MapCol map, ws, COL_DOSSIER
    MapCol map, ws, COL_CONSO
    MapCol map, ws, COL_LIMITE

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' Charger les clés de la feuille en mémoire (2D -> tableaux cibles)
    Dim rngData As Range, vData As Variant
    Set rngData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
    vData = rngData.Value

    Dim outNBJ_VG(), outDEB_VG(), outFIN_VG()
    Dim outNBJ_AM(), outDEB_AM(), outFIN_AM()
    Dim outNBJ_SL(), outDEB_SL(), outFIN_SL()
    Dim outNA()
    ReDim outNBJ_VG(1 To UBound(vData, 1), 1 To 1)
    ReDim outDEB_VG(1 To UBound(vData, 1), 1 To 1)
    ReDim outFIN_VG(1 To UBound(vData, 1), 1 To 1)
    ReDim outNBJ_AM(1 To UBound(vData, 1), 1 To 1)
    ReDim outDEB_AM(1 To UBound(vData, 1), 1 To 1)
    ReDim outFIN_AM(1 To UBound(vData, 1), 1 To 1)
    ReDim outNBJ_SL(1 To UBound(vData, 1), 1 To 1)
    ReDim outDEB_SL(1 To UBound(vData, 1), 1 To 1)
    ReDim outFIN_SL(1 To UBound(vData, 1), 1 To 1)
    ReDim outNA(1 To UBound(vData, 1), 1 To 1)

    ' Dictionnaires DB
    Dim dVG As Object, dAM As Object, dSL As Object
    Set dVG = LoadFluxDict_3in1(dbPath, "VG")
    Set dAM = LoadFluxDict_3in1(dbPath, "AM")
    Set dSL = LoadFluxDict_3in1(dbPath, "SL")
    Dim aVG As Object, aAM As Object, aSL As Object
    Dim prodDateSheet As Variant: prodDateSheet = DetectSheetProductionDate(ws)
    Set aVG = LoadArchiveEndDictForDate_3in1(dbPath, "VG", prodDateSheet)
    Set aAM = LoadArchiveEndDictForDate_3in1(dbPath, "AM", prodDateSheet)
    Set aSL = LoadArchiveEndDictForDate_3in1(dbPath, "SL", prodDateSheet)

    ' Colonnes des champs dans vData
    Dim cBk&, cDos&, cConso&, cLim&
    cBk = map(COL_BOOKING) - 1 - (rngData.Column - 1)
    cDos = map(COL_DOSSIER) - 1 - (rngData.Column - 1)
    cConso = map(COL_CONSO) - 1 - (rngData.Column - 1)
    cLim = map(COL_LIMITE) - 1 - (rngData.Column - 1)

    Dim k As String, conso As Double, lim As Double
    For r = 1 To UBound(vData, 1)
        k = UCase$(Trim$(CStr(vData(r, cBk)))) & "||" & Trim$(CStr(vData(r, cDos)))
        conso = Val(vData(r, cConso)): lim = Val(vData(r, cLim))
        outNA(r, 1) = IIf(Abs(conso) > lim, "YES", "NO")
        If dVG.Exists(k) Then outNBJ_VG(r, 1) = dVG(k)(1): outDEB_VG(r, 1) = dVG(k)(0)
        If dAM.Exists(k) Then outNBJ_AM(r, 1) = dAM(k)(1): outDEB_AM(r, 1) = dAM(k)(0)
        If dSL.Exists(k) Then outNBJ_SL(r, 1) = dSL(k)(1): outDEB_SL(r, 1) = dSL(k)(0)
        If aVG.Exists(k) Then outFIN_VG(r, 1) = aVG(k)
        If aAM.Exists(k) Then outFIN_AM(r, 1) = aAM(k)
        If aSL.Exists(k) Then outFIN_SL(r, 1) = aSL(k)
    Next r

    ' Ecriture en bloc
    ws.Range(ws.Cells(2, cNBJ_VG), ws.Cells(lastRow, cNBJ_VG)).Value = outNBJ_VG
    ws.Range(ws.Cells(2, cDEB_VG), ws.Cells(lastRow, cDEB_VG)).Value = outDEB_VG
    ws.Range(ws.Cells(2, cFIN_VG), ws.Cells(lastRow, cFIN_VG)).Value = outFIN_VG

    ws.Range(ws.Cells(2, cNBJ_AM), ws.Cells(lastRow, cNBJ_AM)).Value = outNBJ_AM
    ws.Range(ws.Cells(2, cDEB_AM), ws.Cells(lastRow, cDEB_AM)).Value = outDEB_AM
    ws.Range(ws.Cells(2, cFIN_AM), ws.Cells(lastRow, cFIN_AM)).Value = outFIN_AM

    ws.Range(ws.Cells(2, cNBJ_SL), ws.Cells(lastRow, cNBJ_SL)).Value = outNBJ_SL
    ws.Range(ws.Cells(2, cDEB_SL), ws.Cells(lastRow, cDEB_SL)).Value = outDEB_SL
    ws.Range(ws.Cells(2, cFIN_SL), ws.Cells(lastRow, cFIN_SL)).Value = outFIN_SL

    ws.Range(ws.Cells(2, cNA), ws.Cells(lastRow, cNA)).Value = outNA
    
    ws.Columns(cNBJ_VG).NumberFormat = "0"
    ws.Columns(cNBJ_AM).NumberFormat = "0"
    ws.Columns(cNBJ_SL).NumberFormat = "0"
    ws.Columns(cDEB_VG).NumberFormat = "dd/mm/yyyy"
    ws.Columns(cDEB_AM).NumberFormat = "dd/mm/yyyy"
    ws.Columns(cDEB_SL).NumberFormat = "dd/mm/yyyy"
    ws.Columns(cFIN_VG).NumberFormat = "dd/mm/yyyy"
    ws.Columns(cFIN_AM).NumberFormat = "dd/mm/yyyy"
    ws.Columns(cFIN_SL).NumberFormat = "dd/mm/yyyy"
End Sub
```
Tu lis la feuille dans vData, tu calcules, puis tu “affiches” en posant les arrays dans les plages cibles :
```
' 1) Lire la feuille en bloc
Dim rngData As Range, vData As Variant
Set rngData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
vData = rngData.Value    ' <-- pas de recordset ici

' 2) Calculer des arrays de sortie
ReDim outNBJ_VG(1 To UBound(vData, 1), 1 To 1)
' ... idem pour outDEB_VG, outFIN_VG, etc.
' remplir les arrays en boucle, via tes dictionnaires dVG/dAM/dSL

' 3) "Afficher" = écrire par blocs dans les colonnes cibles
ws.Range(ws.Cells(2, cNBJ_VG), ws.Cells(lastRow, cNBJ_VG)).Value = outNBJ_VG
ws.Range(ws.Cells(2, cDEB_VG), ws.Cells(lastRow, cDEB_VG)).Value = outDEB_VG
ws.Range(ws.Cells(2, cFIN_VG), ws.Cells(lastRow, cFIN_VG)).Value = outFIN_VG
' ... idem AM/SL + Flag NA
```


# 7) ADODB.Recordset de lecture : curseur ForwardOnly/ReadOnly

Tes rs.Open sql, cn, 1, 1 utilisent adOpenKeyset. Pour simplement lire, préfère :
```vb
rs.Open sql, cn, 0, 1   ' adOpenForwardOnly=0, adLockReadOnly=1

```
et ajoute , , 128 à tous les cn.Execute (adExecuteNoRecords) pour éviter des curseurs implicites.


# 1) Une seule connexion ADO pour tout le run

Évite d’ouvrir/fermer la connexion à chaque étape. Ouvre une connexion en entrée, passe-la à tous les sous-modules, et ferme à la fin.


```vb
Public Sub MAJ_Retards_Daily_3in1_NoNzIif2()
    On Error GoTo KO
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dbPath As String: dbPath = ThisWorkbook.Path & Application.PathSeparator & DB_NAME
    Dim cn As Object: Set cn = GetConn(dbPath)

    Dim rowsImported As Long, newFlux As Long, contFlux As Long, closedFlux As Long

    ' 2) Staging du jour
    rowsImported = LoadTodayIntoAccessFast(cn, ws)

    ' 2b) Normalisation (une seule UPDATE + transaction)
    NormalizeTodayFast cn

    ' 3) Historisation (1 INSERT avec UNION ALL)
    AppendHistoryFast cn

    ' 4) Règles métier
    ProcessBusinessRules_3in1_NoNzIif_Fast cn, newFlux, contFlux, closedFlux

    ' 5) Enrichissement feuille (arrays)
    EnrichSheetFromFlux_3in1_Fast cn, ws

    ' 6) Log
    WriteRunLogFast cn, rowsImported, newFlux, contFlux, closedFlux

    cn.Close: Set cn = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Mise à jour terminée :" & vbCrLf & _
           "Importés : " & rowsImported & vbCrLf & _
           "Nouveaux flux : " & newFlux & vbCrLf & _
           "Continués : " & contFlux & vbCrLf & _
           "Clôturés : " & closedFlux, vbInformation
    If TEST_MODE Then BumpSimDate dbPath
    Exit Sub
KO:
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub


```

# 2) Insert Excel → Access plus rapide (param types + .Prepared)

Gros gain : types ADO corrects (date/double/texte), un seul Command préparé, transaction englobante.

```vb
Private Const adVarWChar = 202, adDouble = 5, adDate = 7, adCmdText = 1, adExecuteNoRecords = 128

Public Function LoadTodayIntoAccessFast(ByVal cn As Object, ByVal ws As Worksheet) As Long
    cn.Execute "DELETE FROM " & T_TODAY, , adExecuteNoRecords

    Dim sql As String
    sql = "INSERT INTO " & T_TODAY & " (Booking,[Production Date],[NO DOSSIER CREDIT],[NO INTERVENANT],[NO INTERVENANT GRP]," & _
          "Limite,Consommation,[Montant VG],[Montant AM],[Montant SL],[Insuffisance VG],[Insuffisance AM],[Insuffisance SL]) " & _
          "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"

    Dim cmd As Object: Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sql
        .Parameters.Append .CreateParameter(, adVarWChar, 1, 20)   ' Booking
        .Parameters.Append .CreateParameter(, adDate,     1)       ' Production Date
        .Parameters.Append .CreateParameter(, adVarWChar, 1, 50)   ' NDC
        .Parameters.Append .CreateParameter(, adVarWChar, 1, 50)   ' Intervenant
        .Parameters.Append .CreateParameter(, adVarWChar, 1, 50)   ' Intervenant Grp
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' Limite
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' Conso
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' MVG
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' MAM
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' MSL
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' IVG
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' IAM
        .Parameters.Append .CreateParameter(, adDouble,   1)       ' ISL
        .Prepared = True
    End With

    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    MapCol map, ws, COL_BOOKING
    MapCol map, ws, COL_PROD
    MapCol map, ws, COL_DOSSIER
    MapCol map, ws, COL_INTER, COL_INTER_ALT
    MapCol map, ws, COL_INTER_GRP, COL_INTER_GRP_ALT
    MapCol map, ws, COL_LIMITE
    MapCol map, ws, COL_CONSO
    MapCol map, ws, COL_MVG, COL_MVG_ALT
    MapCol map, ws, COL_MAM
    MapCol map, ws, COL_MSL
    MapCol map, ws, COL_IVG
    MapCol map, ws, COL_IAM
    MapCol map, ws, COL_ISL

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    ' Lecture en bloc pour réduire les accès COM
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim vData As Variant
    vData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value

    Dim r As Long, n As Long, rowOffset As Long: rowOffset = 2
    cn.BeginTrans
    For r = 1 To UBound(vData, 1)
        cmd.Parameters(0).Value = vData(r, map(COL_BOOKING))             ' Booking
        cmd.Parameters(1).Value = vData(r, map(COL_PROD))                ' Production Date
        cmd.Parameters(2).Value = vData(r, map(COL_DOSSIER))             ' NDC
        cmd.Parameters(3).Value = vData(r, map(COL_INTER))               ' Interv
        cmd.Parameters(4).Value = vData(r, map(COL_INTER_GRP))           ' IntervGrp
        cmd.Parameters(5).Value = vData(r, map(COL_LIMITE))              ' Limite
        cmd.Parameters(6).Value = vData(r, map(COL_CONSO))               ' Conso
        cmd.Parameters(7).Value = vData(r, map(COL_MVG))                 ' MVG
        cmd.Parameters(8).Value = vData(r, map(COL_MAM))                 ' MAM
        cmd.Parameters(9).Value = vData(r, map(COL_MSL))                 ' MSL
        cmd.Parameters(10).Value = vData(r, map(COL_IVG))                ' IVG
        cmd.Parameters(11).Value = vData(r, map(COL_IAM))                ' IAM
        cmd.Parameters(12).Value = vData(r, map(COL_ISL))                ' ISL
        cmd.Execute , , adExecuteNoRecords
        n = n + 1
    Next
    cn.CommitTrans
    LoadTodayIntoAccessFast = n
End Function
```

# 3) Normalisation : 1 seule UPDATE (+ transaction)

Tu fais actuellement 10+ updates séquentiels. Fusionne en une passe.

```vb
Private Sub NormalizeTodayFast(ByVal cn As Object)
    cn.BeginTrans
    ' Upper/trim + null handling + date simu en une seule passe
    Dim sql As String
    sql = _
    "UPDATE " & T_TODAY & " AS t SET " & _
    "t.Booking = UCase(Trim(t.Booking))," & _
    "t.[NO DOSSIER CREDIT] = Trim(t.[NO DOSSIER CREDIT])," & _
    "t.[Production Date] = IIf(t.[Production Date] Is Null, Date(), t.[Production Date])," & _
    "t.Limite = IIf(t.Limite Is Null, 0, t.Limite)," & _
    "t.Consommation = IIf(t.Consommation Is Null, 0, t.Consommation)," & _
    "t.[Montant VG] = IIf(t.[Montant VG] Is Null, 0, t.[Montant VG])," & _
    "t.[Montant AM] = IIf(t.[Montant AM] Is Null, 0, t.[Montant AM])," & _
    "t.[Montant SL] = IIf(t.[Montant SL] Is Null, 0, t.[Montant SL])," & _
    "t.[Insuffisance VG] = IIf(t.[Insuffisance VG] Is Null, 0, t.[Insuffisance VG])," & _
    "t.[Insuffisance AM] = IIf(t.[Insuffisance AM] Is Null, 0, t.[Insuffisance AM])," & _
    "t.[Insuffisance SL] = IIf(t.[Insuffisance SL] Is Null, 0, t.[Insuffisance SL]);"
    cn.Execute sql, , adExecuteNoRecords

    ' Dédupliquer (garde la date la + récente)
    cn.Execute _
      "DELETE FROM " & T_TODAY & " AS a WHERE EXISTS (" & _
      " SELECT 1 FROM " & T_TODAY & " AS b " & _
      " WHERE b.Booking=a.Booking AND b.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] " & _
      "   AND b.[Production Date] > a.[Production Date]);", , adExecuteNoRecords

    If TEST_MODE Then
        Dim dSim As Date: dSim = GetSimDate(ThisWorkbook.Path & Application.PathSeparator & DB_NAME)
        cn.Execute "UPDATE " & T_TODAY & " SET [Production Date] = " & SqlDate(dSim) & ";", , adExecuteNoRecords
    End If
    cn.CommitTrans
End Sub
```

# 4) Historisation : 1 INSERT avec UNION ALL

Trois scans -> un seul.

```vb
Private Sub AppendHistoryFast(ByVal cn As Object)
    cn.Execute _
    "INSERT INTO " & T_HIST & " (Booking,[NO DOSSIER CREDIT],Nature,[Production Date],Insuffisance,Consommation,Limite,[Montant]) " & _
    "SELECT t.Booking,t.[NO DOSSIER CREDIT],'VG',t.[Production Date],t.[Insuffisance VG],t.Consommation,t.Limite,t.[Montant VG] " & _
    "FROM " & T_TODAY & " t WHERE NOT EXISTS(SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='VG' AND h.[Production Date]=t.[Production Date]) " & _
    "UNION ALL " & _
    "SELECT t.Booking,t.[NO DOSSIER CREDIT],'AM',t.[Production Date],t.[Insuffisance AM],t.Consommation,t.Limite,t.[Montant AM] " & _
    "FROM " & T_TODAY & " t WHERE NOT EXISTS(SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='AM' AND h.[Production Date]=t.[Production Date]) " & _
    "UNION ALL " & _
    "SELECT t.Booking,t.[NO DOSSIER CREDIT],'SL',t.[Production Date],t.[Insuffisance SL],t.Consommation,t.Limite,t.[Montant SL] " & _
    "FROM " & T_TODAY & " t WHERE NOT EXISTS(SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='SL' AND h.[Production Date]=t.[Production Date]);", , adExecuteNoRecords

    ' DepassementNA en 1 passe
    cn.Execute _
    "UPDATE " & T_HIST & " AS h INNER JOIN " & T_TODAY & " AS t " & _
    "ON h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.[Production Date]=t.[Production Date] " & _
    "SET h.DepassementNA = IIf(Abs(t.Consommation) > t.Limite,'YES','NO');", , adExecuteNoRecords
End Sub


```


# 5) Enrichissement feuille 100% arrays (écriture en blocs)

Tu as déjà le principe. Voici une version compacte et très rapide :

```vb
Private Sub EnrichSheetFromFlux_3in1_Fast(ByVal cn As Object, ByVal ws As Worksheet)
    Dim cNBJ_VG&, cDEB_VG&, cFIN_VG&, cNBJ_AM&, cDEB_AM&, cFIN_AM&, cNBJ_SL&, cDEB_SL&, cFIN_SL&, cNA&
    cNBJ_VG = EnsureColumn(ws, COL_NBJ_VG): cDEB_VG = EnsureColumn(ws, COL_DEB_VG): cFIN_VG = EnsureColumn(ws, COL_FIN_VG)
    cNBJ_AM = EnsureColumn(ws, COL_NBJ_AM): cDEB_AM = EnsureColumn(ws, COL_DEB_AM): cFIN_AM = EnsureColumn(ws, COL_FIN_AM)
    cNBJ_SL = EnsureColumn(ws, COL_NBJ_SL): cDEB_SL = EnsureColumn(ws, COL_DEB_SL): cFIN_SL = EnsureColumn(ws, COL_FIN_SL)
    cNA = EnsureColumn(ws, COL_FLAG_NA)

    Dim lastRow&, lastCol&
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim rngData As Range: Set rngData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    Dim vData As Variant: vData = rngData.Value

    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    MapCol map, ws, COL_BOOKING
    MapCol map, ws, COL_DOSSIER
    MapCol map, ws, COL_CONSO
    MapCol map, ws, COL_LIMITE

    ' Dictionnaires Flux/Archive
    Dim dVG As Object, dAM As Object, dSL As Object
    Set dVG = LoadFluxDict_3in1_SQL(cn, "VG")
    Set dAM = LoadFluxDict_3in1_SQL(cn, "AM")
    Set dSL = LoadFluxDict_3in1_SQL(cn, "SL")

    Dim prodDateSheet As Variant
    prodDateSheet = DetectSheetProductionDateFast(ws) ' Max via WorksheetFunction

    Dim aVG As Object, aAM As Object, aSL As Object
    Set aVG = LoadArchiveEndDictForDate_3in1_SQL(cn, "VG", prodDateSheet)
    Set aAM = LoadArchiveEndDictForDate_3in1_SQL(cn, "AM", prodDateSheet)
    Set aSL = LoadArchiveEndDictForDate_3in1_SQL(cn, "SL", prodDateSheet)

    Dim n&: n = UBound(vData, 1)
    Dim outNBJ_VG(), outDEB_VG(), outFIN_VG(), outNBJ_AM(), outDEB_AM(), outFIN_AM(), outNBJ_SL(), outDEB_SL(), outFIN_SL(), outNA()
    ReDim outNBJ_VG(1 To n, 1 To 1): ReDim outDEB_VG(1 To n, 1 To 1): ReDim outFIN_VG(1 To n, 1 To 1)
    ReDim outNBJ_AM(1 To n, 1 To 1): ReDim outDEB_AM(1 To n, 1 To 1): ReDim outFIN_AM(1 To n, 1 To 1)
    ReDim outNBJ_SL(1 To n, 1 To 1): ReDim outDEB_SL(1 To n, 1 To 1): ReDim outFIN_SL(1 To n, 1 To 1)
    ReDim outNA(1 To n, 1 To 1)

    Dim i&, bk$, dos$, k$, conso As Double, lim As Double, item As Variant
    For i = 1 To n
        bk = UCase$(Trim$(CStr(vData(i, map(COL_BOOKING)))))
        dos = Trim$(CStr(vData(i, map(COL_DOSSIER))))
        k = bk & "||" & dos

        conso = CDbl(Val(vData(i, map(COL_CONSO))))
        lim = CDbl(Val(vData(i, map(COL_LIMITE))))
        outNA(i, 1) = IIf(Abs(conso) > lim, "YES", "NO")

        If dVG.Exists(k) Then
            item = dVG(k): outDEB_VG(i, 1) = item(0): outNBJ_VG(i, 1) = item(1)
        End If
        If aVG.Exists(k) Then outFIN_VG(i, 1) = aVG(k)

        If dAM.Exists(k) Then
            item = dAM(k): outDEB_AM(i, 1) = item(0): outNBJ_AM(i, 1) = item(1)
        End If
        If aAM.Exists(k) Then outFIN_AM(i, 1) = aAM(k)

        If dSL.Exists(k) Then
            item = dSL(k): outDEB_SL(i, 1) = item(0): outNBJ_SL(i, 1) = item(1)
        End If
        If aSL.Exists(k) Then outFIN_SL(i, 1) = aSL(k)
    Next i

    ' Ecriture par blocs (ultra-rapide)
    ws.Range(ws.Cells(2, cNBJ_VG), ws.Cells(lastRow, cNBJ_VG)).Value = outNBJ_VG
    ws.Range(ws.Cells(2, cDEB_VG), ws.Cells(lastRow, cDEB_VG)).Value = outDEB_VG
    ws.Range(ws.Cells(2, cFIN_VG), ws.Cells(lastRow, cFIN_VG)).Value = outFIN_VG

    ws.Range(ws.Cells(2, cNBJ_AM), ws.Cells(lastRow, cNBJ_AM)).Value = outNBJ_AM
    ws.Range(ws.Cells(2, cDEB_AM), ws.Cells(lastRow, cDEB_AM)).Value = outDEB_AM
    ws.Range(ws.Cells(2, cFIN_AM), ws.Cells(lastRow, cFIN_AM)).Value = outFIN_AM

    ws.Range(ws.Cells(2, cNBJ_SL), ws.Cells(lastRow, cNBJ_SL)).Value = outNBJ_SL
    ws.Range(ws.Cells(2, cDEB_SL), ws.Cells(lastRow, cDEB_SL)).Value = outDEB_SL
    ws.Range(ws.Cells(2, cFIN_SL), ws.Cells(lastRow, cFIN_SL)).Value = outFIN_SL

    ws.Range(ws.Cells(2, cNA), ws.Cells(lastRow, cNA)).Value = outNA

    ws.Columns(cNBJ_VG).NumberFormat = "0": ws.Columns(cNBJ_AM).NumberFormat = "0": ws.Columns(cNBJ_SL).NumberFormat = "0"
    ws.Columns(cDEB_VG).NumberFormat = "dd/mm/yyyy": ws.Columns(cDEB_AM).NumberFormat = "dd/mm/yyyy": ws.Columns(cDEB_SL).NumberFormat = "dd/mm/yyyy"
    ws.Columns(cFIN_VG).NumberFormat = "dd/mm/yyyy": ws.Columns(cFIN_AM).NumberFormat = "dd/mm/yyyy": ws.Columns(cFIN_SL).NumberFormat = "dd/mm/yyyy"
End Sub

Private Function DetectSheetProductionDateFast(ws As Worksheet) As Variant
    Dim col&: col = FindHeader(ws, COL_PROD)
    If col = 0 Then DetectSheetProductionDateFast = Date: Exit Function
    Dim lastRow&: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Dim rng As Range: Set rng = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
    On Error Resume Next
    DetectSheetProductionDateFast = Application.WorksheetFunction.Max(rng)
    If Err.Number <> 0 Then DetectSheetProductionDateFast = Date
    On Error GoTo 0
End Function

Private Function LoadFluxDict_3in1_SQL(ByVal cn As Object, ByVal nature As String) As Object
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    Dim sql As String
    If nature = "VG" Then
        sql = "SELECT Booking,[NO DOSSIER CREDIT],VG_StartDate,VG_NbJoursRetard FROM " & T_FLUX & " WHERE VG_StartDate Is Not Null"
    ElseIf nature = "AM" Then
        sql = "SELECT Booking,[NO DOSSIER CREDIT],AM_StartDate,AM_NbJoursRetard FROM " & T_FLUX & " WHERE AM_StartDate Is Not Null"
    Else
        sql = "SELECT Booking,[NO DOSSIER CREDIT],SL_StartDate,SL_NbJoursRetard FROM " & T_FLUX & " WHERE SL_StartDate Is Not Null"
    End If
    rs.Open sql, cn, 1, 1
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        d(UCase$(Trim$(rs(0))) & "||" & Trim$(rs(1))) = Array(rs(2).Value, rs(3).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set LoadFluxDict_3in1_SQL = d
End Function

Private Function LoadArchiveEndDictForDate_3in1_SQL(ByVal cn As Object, ByVal nature As String, ByVal endDate As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    If Not IsDate(endDate) Then Set LoadArchiveEndDictForDate_3in1_SQL = d: Exit Function
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    Dim f$, sql$
    f = Format$(CDate(endDate), "mm/dd/yyyy")
    If nature = "VG" Then
        sql = "SELECT Booking,[NO DOSSIER CREDIT],VG_EndDate FROM " & T_ARC & " WHERE FlagVG='YES' AND VG_EndDate=#" & f & "#"
    ElseIf nature = "AM" Then
        sql = "SELECT Booking,[NO DOSSIER CREDIT],AM_EndDate FROM " & T_ARC & " WHERE FlagAM='YES' AND AM_EndDate=#" & f & "#"
    Else
        sql = "SELECT Booking,[NO DOSSIER CREDIT],SL_EndDate FROM " & T_ARC & " WHERE FlagSL='YES' AND SL_EndDate=#" & f & "#"
    End If
    rs.Open sql, cn, 1, 1
    Do While Not rs.EOF
        d(UCase$(Trim$(rs(0))) & "||" & Trim$(rs(1))) = rs(2).Value
        rs.MoveNext
    Loop
    rs.Close
    Set LoadArchiveEndDictForDate_3in1_SQL = d
End Function

```

# Option “turbo” Import


```vb
' Dans MAJ_Retards_Daily_3in1_NoNzIif2 :
rowsImported = BulkLoadTodayFromExcel(cn, ThisWorkbook.FullName, ActiveSheet.Name)

```

```vb
' =====================================================================
'  IMPORT MASSIF EXCEL -> ACCESS (ultra-rapide, 1 seul INSERT SELECT)
'  - wbPath    : chemin complet du classeur Excel source (ex: ThisWorkbook.FullName)
'  - sheetName : nom exact de l’onglet (ex: ActiveSheet.Name)
'  Retourne    : nb de lignes insérées dans T_TODAY
'  Hypothèses  : HDR=YES (première ligne = en-têtes)
'                Les en-têtes acceptent tes variantes (NO NTERVENANT, etc.)
' =====================================================================
Public Function BulkLoadTodayFromExcel(ByVal cn As Object, ByVal wbPath As String, ByVal sheetName As String) As Long
    Dim excelSpec As String, ext As String, rows As Long
    ext = LCase$(Mid$(wbPath, InStrRev(wbPath, ".") + 1))

    Select Case ext
        Case "xlsx", "xlsm": excelSpec = "Excel 12.0 Xml"
        Case "xlsb":         excelSpec = "Excel 12.0"
        Case "xls":          excelSpec = "Excel 8.0"
        Case Else:           excelSpec = "Excel 12.0 Xml"
    End Select

    Dim srcTable As String
    srcTable = "[" & excelSpec & ";HDR=YES;IMEX=0;Database=" & wbPath & "].[" & sheetName & "$]"

    ' NB: On gère tes colonnes alternatives directement dans le SELECT:
    '  - NO INTERVENANT / NO NTERVENANT
    '  - NO INTERVENANT GRP / NO NTERVENANT GRP
    '  - Montant VG (valeur de gage-LTV) / Montant VG
    '  On caste pour fiabiliser les types côté Access (CStr/CDate/CDbl/Nz).
    Dim sqlInsert As String
    sqlInsert = _
      "INSERT INTO " & T_TODAY & " (" & _
      "  Booking, [Production Date], [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP], " & _
      "  Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], " & _
      "  [Insuffisance VG], [Insuffisance AM], [Insuffisance SL]" & _
      ") " & _
      "SELECT " & _
      "  CStr([Booking]) AS Booking," & _
      "  CDate([Production Date]) AS [Production Date]," & _
      "  CStr([NO DOSSIER CREDIT]) AS [NO DOSSIER CREDIT]," & _
      "  CStr(IIf([NO INTERVENANT] Is Null Or [NO INTERVENANT]='',[NO NTERVENANT],[NO INTERVENANT])) AS [NO INTERVENANT]," & _
      "  CStr(IIf([NO INTERVENANT GRP] Is Null Or [NO INTERVENANT GRP]='',[NO NTERVENANT GRP],[NO INTERVENANT GRP])) AS [NO INTERVENANT GRP]," & _
      "  CDbl(Nz([Limite],0)) AS Limite," & _
      "  CDbl(Nz([Consommation],0)) AS Consommation," & _
      "  CDbl(Nz(IIf(Not [Montant VG] Is Null,[Montant VG],[Montant VG (valeur de gage-LTV)]),0)) AS [Montant VG]," & _
      "  CDbl(Nz([Montant AM],0)) AS [Montant AM]," & _
      "  CDbl(Nz([Montant SL],0)) AS [Montant SL]," & _
      "  CDbl(Nz([Insuffisance VG],0)) AS [Insuffisance VG]," & _
      "  CDbl(Nz([Insuffisance AM],0)) AS [Insuffisance AM]," & _
      "  CDbl(Nz([Insuffisance SL],0)) AS [Insuffisance SL] " & _
      "FROM " & srcTable & ";"

    ' Exécution sous transaction (DELETE + INSERT) pour accélérer et sécuriser
    cn.BeginTrans
    cn.Execute "DELETE FROM " & T_TODAY, , 128 ' adExecuteNoRecords
    cn.Execute sqlInsert, rows, 128
    cn.CommitTrans

    BulkLoadTodayFromExcel = rows
End Function
```

Ci-dessous une version ultra-optimisée qui :

évite le DELETE (on drop puis SELECT INTO pour recréer T_TODAY, c’est nettement plus rapide dans ACE) ;

limite la lecture Excel au UsedRange réel (A1:…) pour ne pas scanner 1M de lignes vides ;

fait la normalisation dans le SELECT (trim/upper, Nz→0, fallback de date) pour supprimer plusieurs UPDATE derrière ;

recrée l’index clé de T_TODAY juste après l’import (booste les joins/RULES).

Tu gardes la même logique métier ensuite (NormalizeToday pourra rester mais fera quasi rien).

```vb
Dim rowsImported As Long
rowsImported = BulkLoadTodayFromExcel_Fast cn:=GetConn(dbPath), _
                                           wbPath:=ThisWorkbook.FullName, _
                                           sheetName:=ActiveSheet.Name, _
                                           usedRangeFromExcel:=ActiveSheet.UsedRange.Address(False, False)

```


```vb
' =====================================================================
'  IMPORT MASSIF EXCEL -> ACCESS (turbo)
'  - Drop T_TODAY + SELECT INTO (création table) = plus rapide que DELETE + INSERT
'  - Lecture bornée au UsedRange pour éviter le scan de lignes vides
'  - Normalisation inline (Trim/UCase/Nz/Date fallback) pour éviter des UPDATE
'  - Recréation de l’index (Booking, NO DOSSIER CREDIT)
' =====================================================================
Public Function BulkLoadTodayFromExcel_Fast( _
    ByVal cn As Object, _
    ByVal wbPath As String, _
    ByVal sheetName As String, _
    Optional ByVal usedRangeFromExcel As String = "" _
) As Long

    Dim rows As Long, excelSpec As String, ext As String
    ext = LCase$(Mid$(wbPath, InStrRev(wbPath, ".") + 1))

    Select Case ext
        Case "xlsx", "xlsm": excelSpec = "Excel 12.0 Xml"
        Case "xlsb":         excelSpec = "Excel 12.0"
        Case "xls":          excelSpec = "Excel 8.0"
        Case Else:           excelSpec = "Excel 12.0 Xml"
    End Select

    ' Si on reçoit une plage (ex: A1:N25000), on l’applique: [Feuille$A1:N25000]
    Dim suffixRange As String
    If Len(usedRangeFromExcel) > 0 Then
        ' Forcer l’ancrage à partir de A1 si jamais le UsedRange commence plus à droite/plus bas:
        ' (on présume tes entêtes en ligne 1, colonnes contiguës)
        Dim lastCell As String: lastCell = Split(usedRangeFromExcel, ":")(UBound(Split(usedRangeFromExcel, ":")))
        suffixRange = "A1:" & lastCell
    Else
        suffixRange = "" ' pas de plage => ACE lira tout l’onglet (plus lent)
    End If

    Dim srcTable As String
    srcTable = "[" & excelSpec & ";HDR=YES;IMEX=0;Database=" & wbPath & "].[" & sheetName & "$" & suffixRange & "]"

    ' NOTE colonnes alternatives gérées inline:
    ' - NO INTERVENANT / NO NTERVENANT
    ' - NO INTERVENANT GRP / NO NTERVENANT GRP
    ' - Montant VG (valeur de gage-LTV) / Montant VG
    ' Normalisation inline:
    '  * Booking => UCase(Trim())
    '  * NO DOSSIER CREDIT => Trim()
    '  * Production Date => fallback Date() si vide
    '  * Numériques => CDbl(Nz(...,0))
    Dim sqlInto As String
    sqlInto = _
      "SELECT " & _
      "  UCase(Trim([Booking])) AS Booking," & _
      "  IIf([Production Date] Is Null Or [Production Date]='', Date(), CDate([Production Date])) AS [Production Date]," & _
      "  Trim([NO DOSSIER CREDIT]) AS [NO DOSSIER CREDIT]," & _
      "  CStr(IIf([NO INTERVENANT] Is Null Or [NO INTERVENANT]='',[NO NTERVENANT],[NO INTERVENANT])) AS [NO INTERVENANT]," & _
      "  CStr(IIf([NO INTERVENANT GRP] Is Null Or [NO INTERVENANT GRP]='',[NO NTERVENANT GRP],[NO INTERVENANT GRP])) AS [NO INTERVENANT GRP]," & _
      "  CDbl(Nz([Limite],0)) AS Limite," & _
      "  CDbl(Nz([Consommation],0)) AS Consommation," & _
      "  CDbl(Nz(IIf(Not [Montant VG] Is Null,[Montant VG],[Montant VG (valeur de gage-LTV)]),0)) AS [Montant VG]," & _
      "  CDbl(Nz([Montant AM],0)) AS [Montant AM]," & _
      "  CDbl(Nz([Montant SL],0)) AS [Montant SL]," & _
      "  CDbl(Nz([Insuffisance VG],0)) AS [Insuffisance VG]," & _
      "  CDbl(Nz([Insuffisance AM],0)) AS [Insuffisance AM]," & _
      "  CDbl(Nz([Insuffisance SL],0)) AS [Insuffisance SL] " & _
      "INTO " & T_TODAY & " " & _
      "FROM " & srcTable & ";"

    On Error GoTo FAIL
    cn.BeginTrans

    ' Drop T_TODAY (si existe) pour éviter DELETE massif + conserver des pages propres
    On Error Resume Next
    cn.Execute "DROP TABLE " & T_TODAY, , 128
    On Error GoTo 0

    ' Création directe de T_TODAY via SELECT INTO (types inférés par les casts)
    cn.Execute sqlInto, rows, 128

    ' Index clé pour accélérer les joins/updates suivants
    On Error Resume Next
    cn.Execute "CREATE INDEX IX_Today_Key ON " & T_TODAY & " (Booking, [NO DOSSIER CREDIT])", , 128
    On Error GoTo 0

    cn.CommitTrans
    BulkLoadTodayFromExcel_Fast = rows
    Exit Function

FAIL:
    On Error Resume Next
    cn.RollbackTrans
    Err.Raise Err.Number, , Err.Description
End Function
```
Pourquoi c’est plus rapide

DROP + SELECT INTO : Access/ACE alloue une table neuve et écrit séquentiellement (pas de journalisation ligne/ligne comme un DELETE suivi d’un INSERT).

Plage bornée A1:… : ACE ne parcourt que les cellules utiles ; les classeurs “larges” gagnent énormément.

Normalisation dans le SELECT : tu supprimes 7–8 UPDATE post-import.

Index recréé seulement après import : coût minimal, bénéfice maximal pour les étapes suivantes.


# Une seule connexion ADO pour tout le run

Ouvrir/fermer la connexion coûte cher. Ouvre une connexion en début de macro et passe-la à toutes les fonctions.

```vb
Public Sub MAJ_Retards_Daily_Turbo()
    On Error GoTo KO
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dbPath As String: dbPath = ThisWorkbook.Path & Application.PathSeparator & DB_NAME
    Dim cn As Object: Set cn = GetConn_Tuned(dbPath) ' <— cf. fonction ci-dessous

    cn.BeginTrans

    ' 1) Import massif borné au UsedRange
    Dim rowsImported As Long
    rowsImported = BulkLoadTodayFromExcel_Fast(cn, ThisWorkbook.FullName, ws.Name, ws.UsedRange.Address(False, False))

    ' 2) Historisation (UNION ALL => une seule passe)
    AppendHistory_Fast cn

    ' 3) Règles (regroupements d’UPDATE + refresh contexte en 1 passe)
    Dim nNew As Long, nCont As Long, nClosed As Long
    ProcessBusinessRules_Faster cn, nNew, nCont, nClosed

    cn.CommitTrans

    ' 4) Enrichissement feuille en blocs
    EnrichSheetFromFlux_ByBlocks cn, ws

    ' 5) Log
    WriteRunLog_WithCn cn, rowsImported, nNew, nCont, nClosed

FIN:
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub
KO:
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.RollbackTrans
    Resume FIN
End Sub

```

Connexion “tunée” (tampon ACE plus large, verrouillage optimisé) :
```vb
Private Function GetConn_Tuned(ByVal dbPath As String) As Object
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    Dim cs As String
    cs = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
         "Data Source=" & dbPath & ";" & _
         "Persist Security Info=False;" & _
         "Jet OLEDB:Database Locking Mode=1;" & _          ' page-level, rapide
         "Jet OLEDB:Max Buffer Size=4096;" & _             ' tampon x2
         "OLE DB Services=-2;"                             ' cache minimal
    On Error Resume Next
    cn.Open cs
    If cn.State = 0 Then
        cs = Replace(cs, "12.0", "16.0")
        cn.Open cs
    End If
    On Error GoTo 0
    Set GetConn_Tuned = cn
End Function

```

# Historisation en une seule requête (UNION ALL)

Remplace tes 3 INSERT par un seul INSERT ... SELECT ... UNION ALL ... :

```vb
Private Sub AppendHistory_Fast(ByVal cn As Object)
    cn.Execute _
    "INSERT INTO " & T_HIST & " (Booking,[NO DOSSIER CREDIT],Nature,[Production Date],Insuffisance,Consommation,Limite,[Montant]) " & _
    "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'VG', t.[Production Date], t.[Insuffisance VG], t.Consommation, t.Limite, t.[Montant VG] " & _
    "FROM " & T_TODAY & " t " & _
    "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='VG' AND h.[Production Date]=t.[Production Date]) " & _
    "UNION ALL " & _
    "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'AM', t.[Production Date], t.[Insuffisance AM], t.Consommation, t.Limite, t.[Montant AM] " & _
    "FROM " & T_TODAY & " t " & _
    "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='AM' AND h.[Production Date]=t.[Production Date]) " & _
    "UNION ALL " & _
    "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'SL', t.[Production Date], t.[Insuffisance SL], t.Consommation, t.Limite, t.[Montant SL] " & _
    "FROM " & T_TODAY & " t " & _
    "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='SL' AND h.[Production Date]=t.[Production Date]);", , 128
    ' MAJ DepassementNA en 1 passe
    cn.Execute _
    "UPDATE " & T_HIST & " AS h " & _
    "INNER JOIN " & T_TODAY & " AS t " & _
    "ON h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.[Production Date]=t.[Production Date] " & _
    "SET h.DepassementNA=IIf(Abs(t.Consommation)>t.Limite,'YES','NO');", , 128
End Sub

``` 


# 1) Une seule connexion + une seule transaction “tour complet”

Garde une connexion persistante tout le run, et encapsule tout dans UNE transaction. Ça évite l’overhead ADO/ACE et le journal écrit plusieurs fois.

```vb
' Module-level
Private m_cn As ADODB.Connection

Private Function GetSharedConn(dbPath As String) As ADODB.Connection
    If m_cn Is Nothing Then
        Set m_cn = New ADODB.Connection
        ' ACE 16 puis fallback 12
        On Error Resume Next
        m_cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & dbPath & ";"
        If m_cn.State = 0 Then m_cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
        On Error GoTo 0
    End If
    Set GetSharedConn = m_cn
End Function

Public Sub MAJ_Retards_Daily_Turbo()
    Dim cn As ADODB.Connection
    Set cn = GetSharedConn(ThisWorkbook.Path & "\" & DB_NAME)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ROLLBACK

    cn.BeginTrans

    ' 1. Bulk load staging (voir §3)
    Dim n&: n = BulkLoadTodayFromExcel_NameRange(cn, ThisWorkbook.FullName, ActiveSheet.Name, "tblToday")

    ' 2. Normalisation “1 passe” (voir §2)
    NormalizeToday_OnePass cn

    ' 3. KeyMap + Affected (si tu as activé KeyID)
    Rebuild_KeyMap_And_Affected cn

    ' 4. History “1 INSERT” (voir §4)
    AppendHistory_SingleInsert cn

    ' 5. Règles (réécrites sur KeyID + T_Affected)
    ProcessRules_Turbo cn

    ' 6. Enrichissement feuille par blocs
    EnrichSheet_ByBlocks cn, ActiveSheet

    ' 7. RunLog
    WriteRunLog cn, n, 0, 0, 0

    cn.CommitTrans
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "OK (" & n & " lignes importées)", vbInformation
    Exit Sub
ROLLBACK:
    If cn.State = 1 Then On Error Resume Next: cn.RollbackTrans: On Error GoTo 0
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur: " & Err.Description, vbCritical
End Sub


``` 