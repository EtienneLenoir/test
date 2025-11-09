# 📦 Pack complet — Excel VBA + Access (VG/AM/SL) **sans Nz/IIf dans le SQL**, booléens en **'YES'/'NO'** (string)
> Objectif en une phrase : **historiser le nombre de jours de retard** (Insuffisance < 0) pour chaque clé **(Booking + NO DOSSIER CREDIT)**, **par nature** (VG, AM, SL), avec **flux en cours**, **archives**, **pire insuffisance** & **dates** (début/fin/pire), et **mise à jour quotidienne** via **un bouton** Excel.  
> Les champs booléens dans la base sont stockés en **texte** — `'YES'` ou `'NO'` — (pas de type YES/NO Access).

---

## 🧱 Architecture (propre & maintenable)

### Tables (Access, 1 fichier `.accdb`)
- **T_Today** (staging du jour, miroir de la feuille)
- **T_Flux3** (flux **en cours**, une ligne par clé, avec 3 sous-épisodes VG/AM/SL)
- **T_Archive3** (archives **par clôture** d’un sous-épisode VG/AM/SL ; une ligne par clôture)
- **T_HistoryInsuff** (séries temporelles quotidiennes par clé et par nature — utile pour reporting croisé et flags multi-natures)
- **T_RunLog** (journal d’exécution)

### Clés & relations (logiques)
- Clé **métier unique** = `(Booking, [NO DOSSIER CREDIT])` sur **T_Flux3** (PK) et index identique sur **T_Archive3**/**T_Today**/**T_HistoryInsuff** (via index composites et contraintes d’unicité pertinentes).
- Une **clé** peut avoir **jusqu’à 3 sous-épisodes en parallèle** (VG, AM, SL). Les dates/compteurs/pire sont doublés/triplés dans **T_Flux3**.
- **Archivage** : chaque **clôture** (VG/AM/SL) génère **1 ligne** dans **T_Archive3** et purge les champs de cette nature dans **T_Flux3**.  
- Les **booléens** (dépassement non autorisé / flags de nature) sont des **TEXT(3)** avec `'YES'`/`'NO'` (jamais TRUE/FALSE, jamais YESNO).

### Règles de gestion (par jour **Production Date**)
1) Insuffisance\<0 ⇒ **jour de retard** pour la **nature** concernée.  
2) **Start** : si la nature devient \<0 et n’était pas ouverte ⇒ `StartDate=t.[Production Date]`, `NbJoursRetard=1`, `WorstInsuffisance`=valeur du jour, etc.  
3) **Continue** : si encore \<0 ⇒ on **incrémente** de `DateDiff('d', LastUpdateDate, ProductionDate)` (garde multi-runs) et on met à jour `Worst*` si plus négatif.  
4) **Close** : si la nature n’est plus \<0 (ou la clé a disparu du fichier), on **archive** (`EndDate` = `LastUpdateDate` si clé absente, sinon `t.[Production Date]`) puis **on purge** les champs de la nature dans le flux.  
5) **Dépassement non autorisé** : `'YES'` si `Abs(Consommation) > Limite`, sinon `'NO'` (string).  
6) **Flags d’archives** : `FlagVG/AM/SL` à `'YES'`/`'NO'` selon la nature concernée à la clôture + recalage multi-natures via **T_HistoryInsuff**.

---

## 🗃️ DDL Access (création auto si la base n'existe pas)

```sql
-- FLUX (1 ligne par clé, 3 sous-épisodes VG/AM/SL)
CREATE TABLE T_Flux3 (
  Booking TEXT(20) NOT NULL,
  [NO DOSSIER CREDIT] TEXT(50) NOT NULL,

  VG_StartDate DATE,
  VG_LastUpdateDate DATE,
  VG_NbJoursRetard LONG,
  VG_WorstInsuffisance DOUBLE,
  VG_DateWorstInsuffisance DATE,
  VG_LastInsuffisance DOUBLE,
  VG_EndDate DATE,

  AM_StartDate DATE,
  AM_LastUpdateDate DATE,
  AM_NbJoursRetard LONG,
  AM_WorstInsuffisance DOUBLE,
  AM_DateWorstInsuffisance DATE,
  AM_LastInsuffisance DOUBLE,
  AM_EndDate DATE,

  SL_StartDate DATE,
  SL_LastUpdateDate DATE,
  SL_NbJoursRetard LONG,
  SL_WorstInsuffisance DOUBLE,
  SL_DateWorstInsuffisance DATE,
  SL_LastInsuffisance DOUBLE,
  SL_EndDate DATE,

  Limite DOUBLE,
  Consommation DOUBLE,
  [Montant VG] DOUBLE,
  [Montant AM] DOUBLE,
  [Montant SL] DOUBLE,
  [Production Date] DATE,
  [NO INTERVENANT] TEXT(50),
  [NO INTERVENANT GRP] TEXT(50),

  DepassementNA TEXT(3) DEFAULT 'NO',

  CONSTRAINT PK_Flux3 PRIMARY KEY (Booking, [NO DOSSIER CREDIT])
);
CREATE INDEX IX_Flux3_VG_Start ON T_Flux3 (VG_StartDate);
CREATE INDEX IX_Flux3_AM_Start ON T_Flux3 (AM_StartDate);
CREATE INDEX IX_Flux3_SL_Start ON T_Flux3 (SL_StartDate);

-- ARCHIVES (1 ligne/par clôture d’un sous-épisode)
CREATE TABLE T_Archive3 (
  ID AUTOINCREMENT PRIMARY KEY,
  Booking TEXT(20) NOT NULL,
  [NO DOSSIER CREDIT] TEXT(50) NOT NULL,

  VG_StartDate DATE,
  VG_EndDate DATE,
  VG_Duree LONG,
  VG_WorstInsuffisance DOUBLE,
  VG_DateWorstInsuffisance DATE,
  VG_LastInsuffisance DOUBLE,
  VG_LastUpdateDate DATE,

  AM_StartDate DATE,
  AM_EndDate DATE,
  AM_Duree LONG,
  AM_WorstInsuffisance DOUBLE,
  AM_DateWorstInsuffisance DATE,
  AM_LastInsuffisance DOUBLE,
  AM_LastUpdateDate DATE,

  SL_StartDate DATE,
  SL_EndDate DATE,
  SL_Duree LONG,
  SL_WorstInsuffisance DOUBLE,
  SL_DateWorstInsuffisance DATE,
  SL_LastInsuffisance DOUBLE,
  SL_LastUpdateDate DATE,

  Limite DOUBLE,
  Consommation DOUBLE,
  [Montant VG] DOUBLE,
  [Montant AM] DOUBLE,
  [Montant SL] DOUBLE,
  [Production Date] DATE,
  [NO INTERVENANT] TEXT(50),
  [NO INTERVENANT GRP] TEXT(50),

  DepassementNA TEXT(3) DEFAULT 'NO',
  FlagVG TEXT(3) DEFAULT 'NO',
  FlagAM TEXT(3) DEFAULT 'NO',
  FlagSL TEXT(3) DEFAULT 'NO',

  ArchivedAt DATE DEFAULT Date()
);
CREATE INDEX IX_Arc3_Key ON T_Archive3 (Booking, [NO DOSSIER CREDIT]);
CREATE INDEX IX_Arc3_VG_End ON T_Archive3 (VG_EndDate);
CREATE INDEX IX_Arc3_AM_End ON T_Archive3 (AM_EndDate);
CREATE INDEX IX_Arc3_SL_End ON T_Archive3 (SL_EndDate);

-- STAGING
CREATE TABLE T_Today (
  Booking TEXT(20),
  [Production Date] DATE,
  [NO DOSSIER CREDIT] TEXT(50),
  [NO INTERVENANT] TEXT(50),
  [NO INTERVENANT GRP] TEXT(50),
  Limite DOUBLE,
  Consommation DOUBLE,
  [Montant VG] DOUBLE,
  [Montant AM] DOUBLE,
  [Montant SL] DOUBLE,
  [Insuffisance VG] DOUBLE,
  [Insuffisance AM] DOUBLE,
  [Insuffisance SL] DOUBLE
);
CREATE INDEX IX_Today_Key ON T_Today (Booking, [NO DOSSIER CREDIT]);

-- HISTORY (time-series par jour et par nature)
CREATE TABLE T_HistoryInsuff (
  Booking TEXT(20) NOT NULL,
  [NO DOSSIER CREDIT] TEXT(50) NOT NULL,
  Nature TEXT(2) NOT NULL,           -- 'VG' / 'AM' / 'SL'
  [Production Date] DATE NOT NULL,
  Insuffisance DOUBLE,
  Consommation DOUBLE,
  Limite DOUBLE,
  [Montant] DOUBLE,
  DepassementNA TEXT(3) DEFAULT 'NO',
  CONSTRAINT IX_Hist UNIQUE (Booking, [NO DOSSIER CREDIT], Nature, [Production Date])
);

-- RUN LOG
CREATE TABLE T_RunLog (
  RunAt DATETIME NOT NULL,
  RowsImported LONG,
  NewFlux LONG,
  Continued LONG,
  Closed LONG,
  [User] TEXT(100)
);
```

---

## 💡 Mode opératoire (quotidien)
1. Coller/rafraîchir la feuille du jour (en-têtes exacts, variantes gérées).
2. Cliquer le bouton assigné à **`MAJ_Retards_Daily_3in1_NoNzIif`**.
3. La macro :
   - crée la base si besoin et le schéma,
   - charge la feuille dans **T_Today**,
   - normalise les `NULL` & nettoie,
   - alimente **T_HistoryInsuff**,
   - met à jour **T_Flux3** et **T_Archive3**,
   - enrichit la feuille avec 9 colonnes : NbJ/Début/Fin pour **VG/AM/SL** + flag dépassement `'YES'/'NO'`,
   - journalise dans **T_RunLog**.

---

## 🧰 Module VBA complet

> Colle **tout** ce bloc dans un **Module standard** et assigne **`MAJ_Retards_Daily_3in1_NoNzIif`** à un bouton.

```vb
Option Explicit

' ====== CONFIG ======
Private Const DB_NAME As String = "Lombard_Retards.accdb"
Private Const T_FLUX As String = "T_Flux3"
Private Const T_ARC As String = "T_Archive3"
Private Const T_TODAY As String = "T_Today"
Private Const T_HIST As String = "T_HistoryInsuff"
Private Const T_LOG As String = "T_RunLog"

' En-têtes Excel (ligne 1)
Private Const COL_BOOKING As String = "Booking"
Private Const COL_PROD As String = "Production Date"
Private Const COL_DOSSIER As String = "NO DOSSIER CREDIT"
Private Const COL_INTER As String = "NO INTERVENANT"
Private Const COL_INTER_ALT As String = "NO NTERVENANT"
Private Const COL_INTER_GRP As String = "NO INTERVENANT GRP"
Private Const COL_INTER_GRP_ALT As String = "NO NTERVENANT GRP"

Private Const COL_LIMITE As String = "Limite"
Private Const COL_CONSO As String = "Consommation"

Private Const COL_MVG As String = "Montant VG (valeur de gage-LTV)"
Private Const COL_MVG_ALT As String = "Montant VG"
Private Const COL_MAM As String = "Montant AM"
Private Const COL_MSL As String = "Montant SL"

Private Const COL_IVG As String = "Insuffisance VG"
Private Const COL_IAM As String = "Insuffisance AM"
Private Const COL_ISL As String = "Insuffisance SL"

' Colonnes ajoutées (feuille)
Private Const COL_NBJ_VG As String = "NB Jours retard VG"
Private Const COL_DEB_VG As String = "Date début VG"
Private Const COL_FIN_VG As String = "Date fin VG"

Private Const COL_NBJ_AM As String = "NB Jours retard AM"
Private Const COL_DEB_AM As String = "Date début AM"
Private Const COL_FIN_AM As String = "Date fin AM"

Private Const COL_NBJ_SL As String = "NB Jours retard SL"
Private Const COL_DEB_SL As String = "Date début SL"
Private Const COL_FIN_SL As String = "Date fin SL"

Private Const COL_FLAG_NA As String = "Dépassement non autorisé"

' ====== PUBLIC ======
Public Sub MAJ_Retards_Daily_3in1_NoNzIif()
    On Error GoTo KO
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dbPath As String: dbPath = ThisWorkbook.Path & Application.PathSeparator & DB_NAME

    ' 1) DB + schéma
    EnsureDatabaseAndSchema dbPath

    ' 2) Staging du jour
    Dim rowsImported As Long
    rowsImported = LoadTodayIntoAccess(dbPath, ws)

    ' 2b) Normaliser et neutraliser NULL (=> plus besoin de Nz/IIf dans le SQL)
    NormalizeToday dbPath

    ' 3) Historique quotidien
    AppendHistory dbPath

    ' 4) Règles de gestion 3-en-1
    Dim newFlux As Long, contFlux As Long, closedFlux As Long
    ProcessBusinessRules_3in1_NoNzIif dbPath, newFlux, contFlux, closedFlux

    ' 5) Enrichir la feuille
    EnrichSheetFromFlux_3in1 dbPath, ws

    ' 6) Log
    WriteRunLog dbPath, rowsImported, newFlux, contFlux, closedFlux

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Mise à jour terminée :" & vbCrLf & _
           "Importés : " & rowsImported & vbCrLf & _
           "Nouveaux flux : " & newFlux & vbCrLf & _
           "Continués : " & contFlux & vbCrLf & _
           "Clôturés : " & closedFlux, vbInformation
    Exit Sub
KO:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

' ====== DB bootstrap ======
Private Sub EnsureDatabaseAndSchema(ByVal dbPath As String)
    If Dir(dbPath, vbNormal) = vbNullString Then
        Dim cat As Object: Set cat = CreateObject("ADOX.Catalog")
        cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    End If
    Dim cn As Object: Set cn = GetConn(dbPath)
    ExecIfNotExists cn, T_FLUX, GetDDL_Flux
    ExecIfNotExists cn, T_ARC, GetDDL_Arc
    ExecIfNotExists cn, T_TODAY, GetDDL_Today
    ExecIfNotExists cn, T_HIST, GetDDL_Hist
    ExecIfNotExists cn, T_LOG, GetDDL_Log
    cn.Close: Set cn = Nothing
End Sub

Private Function GetConn(ByVal dbPath As String) As Object
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    On Error Resume Next
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & dbPath & ";Persist Security Info=False;"
    If cn.State = 0 Then cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Persist Security Info=False;"
    On Error GoTo 0
    Set GetConn = cn
End Function

Private Sub ExecIfNotExists(ByVal cn As Object, ByVal tableName As String, ByVal ddl As String)
    Dim rs As Object: Set rs = cn.OpenSchema(20) ' adSchemaTables
    Dim exists As Boolean: exists = False
    Do While Not rs.EOF
        If StrComp(rs.Fields("TABLE_NAME").Value & "", tableName, vbTextCompare) = 0 Then exists = True: Exit Do
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    If Not exists Then cn.Execute ddl
End Sub

Private Function GetDDL_Flux() As String
    GetDDL_Flux = _
      "CREATE TABLE " & T_FLUX & " (" & _
      "Booking TEXT(20) NOT NULL, [NO DOSSIER CREDIT] TEXT(50) NOT NULL, " & _
      "VG_StartDate DATE, VG_LastUpdateDate DATE, VG_NbJoursRetard LONG, VG_WorstInsuffisance DOUBLE, VG_DateWorstInsuffisance DATE, VG_LastInsuffisance DOUBLE, VG_EndDate DATE, " & _
      "AM_StartDate DATE, AM_LastUpdateDate DATE, AM_NbJoursRetard LONG, AM_WorstInsuffisance DOUBLE, AM_DateWorstInsuffisance DATE, AM_LastInsuffisance DOUBLE, AM_EndDate DATE, " & _
      "SL_StartDate DATE, SL_LastUpdateDate DATE, SL_NbJoursRetard LONG, SL_WorstInsuffisance DOUBLE, SL_DateWorstInsuffisance DATE, SL_LastInsuffisance DOUBLE, SL_EndDate DATE, " & _
      "Limite DOUBLE, Consommation DOUBLE, [Montant VG] DOUBLE, [Montant AM] DOUBLE, [Montant SL] DOUBLE, [Production Date] DATE, [NO INTERVENANT] TEXT(50), [NO INTERVENANT GRP] TEXT(50), " & _
      "DepassementNA TEXT(3) DEFAULT 'NO', " & _
      "CONSTRAINT PK_Flux3 PRIMARY KEY (Booking, [NO DOSSIER CREDIT]) );" & _
      "CREATE INDEX IX_Flux3_VG_Start ON " & T_FLUX & " (VG_StartDate);" & _
      "CREATE INDEX IX_Flux3_AM_Start ON " & T_FLUX & " (AM_StartDate);" & _
      "CREATE INDEX IX_Flux3_SL_Start ON " & T_FLUX & " (SL_StartDate);"
End Function

Private Function GetDDL_Arc() As String
    GetDDL_Arc = _
      "CREATE TABLE " & T_ARC & " (" & _
      "ID COUNTER PRIMARY KEY, Booking TEXT(20) NOT NULL, [NO DOSSIER CREDIT] TEXT(50) NOT NULL, " & _
      "VG_StartDate DATE, VG_EndDate DATE, VG_Duree LONG, VG_WorstInsuffisance DOUBLE, VG_DateWorstInsuffisance DATE, VG_LastInsuffisance DOUBLE, VG_LastUpdateDate DATE, " & _
      "AM_StartDate DATE, AM_EndDate DATE, AM_Duree LONG, AM_WorstInsuffisance DOUBLE, AM_DateWorstInsuffisance DATE, AM_LastInsuffisance DOUBLE, AM_LastUpdateDate DATE, " & _
      "SL_StartDate DATE, SL_EndDate DATE, SL_Duree LONG, SL_WorstInsuffisance DOUBLE, SL_DateWorstInsuffisance DATE, SL_LastInsuffisance DOUBLE, SL_LastUpdateDate DATE, " & _
      "Limite DOUBLE, Consommation DOUBLE, [Montant VG] DOUBLE, [Montant AM] DOUBLE, [Montant SL] DOUBLE, [Production Date] DATE, [NO INTERVENANT] TEXT(50), [NO INTERVENANT GRP] TEXT(50), " & _
      "DepassementNA TEXT(3) DEFAULT 'NO', FlagVG TEXT(3) DEFAULT 'NO', FlagAM TEXT(3) DEFAULT 'NO', FlagSL TEXT(3) DEFAULT 'NO', ArchivedAt DATE DEFAULT Date());" & _
      "CREATE INDEX IX_Arc3_Key ON " & T_ARC & " (Booking, [NO DOSSIER CREDIT]);" & _
      "CREATE INDEX IX_Arc3_VG_End ON " & T_ARC & " (VG_EndDate);" & _
      "CREATE INDEX IX_Arc3_AM_End ON " & T_ARC & " (AM_EndDate);" & _
      "CREATE INDEX IX_Arc3_SL_End ON " & T_ARC & " (SL_EndDate);"
End Function

Private Function GetDDL_Today() As String
    GetDDL_Today = _
      "CREATE TABLE " & T_TODAY & " (" & _
      "Booking TEXT(20), [Production Date] DATE, [NO DOSSIER CREDIT] TEXT(50), " & _
      "[NO INTERVENANT] TEXT(50), [NO INTERVENANT GRP] TEXT(50), " & _
      "Limite DOUBLE, Consommation DOUBLE, [Montant VG] DOUBLE, [Montant AM] DOUBLE, [Montant SL] DOUBLE, " & _
      "[Insuffisance VG] DOUBLE, [Insuffisance AM] DOUBLE, [Insuffisance SL] DOUBLE);" & _
      "CREATE INDEX IX_Today_Key ON " & T_TODAY & " (Booking, [NO DOSSIER CREDIT]);"
End Function

Private Function GetDDL_Hist() As String
    GetDDL_Hist = _
      "CREATE TABLE " & T_HIST & " (" & _
      "Booking TEXT(20) NOT NULL, [NO DOSSIER CREDIT] TEXT(50) NOT NULL, Nature TEXT(2) NOT NULL, [Production Date] DATE NOT NULL, " & _
      "Insuffisance DOUBLE, Consommation DOUBLE, Limite DOUBLE, [Montant] DOUBLE, DepassementNA TEXT(3) DEFAULT 'NO', " & _
      "CONSTRAINT IX_Hist UNIQUE (Booking, [NO DOSSIER CREDIT], Nature, [Production Date]) );"
End Function

Private Function GetDDL_Log() As String
    GetDDL_Log = _
      "CREATE TABLE " & T_LOG & " (" & _
      "RunAt DATETIME NOT NULL, RowsImported LONG, NewFlux LONG, Continued LONG, Closed LONG, [User] TEXT(100));"
End Function
```

```vb
' ====== Staging ======
Private Function LoadTodayIntoAccess(ByVal dbPath As String, ByVal ws As Worksheet) As Long
    Dim cn As Object: Set cn = GetConn(dbPath)
    cn.Execute "DELETE FROM " & T_TODAY

    Dim sql As String
    sql = "INSERT INTO " & T_TODAY & " (Booking, [Production Date], [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP], " & _
          "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Insuffisance VG], [Insuffisance AM], [Insuffisance SL]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"

    Dim cmd As Object: Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = sql: cmd.CommandType = 1

    Dim i As Long
    For i = 1 To 13: cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 255): Next i

    Dim hdrRow As Long: hdrRow = 1
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= hdrRow Then LoadTodayIntoAccess = 0: GoTo ExitHere

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

    Dim r As Long, n As Long
    cn.BeginTrans
    For r = hdrRow + 1 To lastRow
        cmd.Parameters(0).Value = ws.Cells(r, map(COL_BOOKING)).Value
        cmd.Parameters(1).Value = ws.Cells(r, map(COL_PROD)).Value
        cmd.Parameters(2).Value = ws.Cells(r, map(COL_DOSSIER)).Value
        cmd.Parameters(3).Value = GetCellMaybe(ws, r, map, COL_INTER, COL_INTER_ALT)
        cmd.Parameters(4).Value = GetCellMaybe(ws, r, map, COL_INTER_GRP, COL_INTER_GRP_ALT)
        cmd.Parameters(5).Value = ws.Cells(r, map(COL_LIMITE)).Value
        cmd.Parameters(6).Value = ws.Cells(r, map(COL_CONSO)).Value
        cmd.Parameters(7).Value = GetCellMaybe(ws, r, map, COL_MVG, COL_MVG_ALT)
        cmd.Parameters(8).Value = ws.Cells(r, map(COL_MAM)).Value
        cmd.Parameters(9).Value = ws.Cells(r, map(COL_MSL)).Value
        cmd.Parameters(10).Value = ws.Cells(r, map(COL_IVG)).Value
        cmd.Parameters(11).Value = ws.Cells(r, map(COL_IAM)).Value
        cmd.Parameters(12).Value = ws.Cells(r, map(COL_ISL)).Value
        cmd.Execute , , 128
        n = n + 1
    Next r
    cn.CommitTrans

    LoadTodayIntoAccess = n
ExitHere:
    cn.Close: Set cn = Nothing
End Function
```

```vb
' ====== Normalisation / zéro NULL ======
Private Sub NormalizeToday(ByVal dbPath As String)
    Dim cn As Object: Set cn = GetConn(dbPath)

    cn.Execute "UPDATE " & T_TODAY & " SET Booking = UCase(Trim(Booking));"
    cn.Execute "UPDATE " & T_TODAY & " SET [NO DOSSIER CREDIT] = Trim([NO DOSSIER CREDIT]);"
    cn.Execute "UPDATE " & T_TODAY & " SET [Production Date] = Date() WHERE [Production Date] IS NULL;"

    cn.Execute "UPDATE " & T_TODAY & " SET Limite=0 WHERE Limite IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET Consommation=0 WHERE Consommation IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET [Montant VG]=0 WHERE [Montant VG] IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET [Montant AM]=0 WHERE [Montant AM] IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET [Montant SL]=0 WHERE [Montant SL] IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET [Insuffisance VG]=0 WHERE [Insuffisance VG] IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET [Insuffisance AM]=0 WHERE [Insuffisance AM] IS NULL;"
    cn.Execute "UPDATE " & T_TODAY & " SET [Insuffisance SL]=0 WHERE [Insuffisance SL] IS NULL;"

    cn.Execute _
      "DELETE FROM " & T_TODAY & " AS a " & _
      "WHERE EXISTS (SELECT 1 FROM " & T_TODAY & " AS b " & _
      "              WHERE b.Booking=a.Booking AND b.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] " & _
      "                AND b.[Production Date] > a.[Production Date]);"

    cn.Close: Set cn = Nothing
End Sub
```

```vb
' ====== Historique quotidien ======
Private Sub AppendHistory(ByVal dbPath As String)
    Dim cn As Object: Set cn = GetConn(dbPath)

    ' INSERT sans DepassementNA (default 'NO')
    cn.Execute _
      "INSERT INTO " & T_HIST & " (Booking, [NO DOSSIER CREDIT], Nature, [Production Date], Insuffisance, Consommation, Limite, [Montant]) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'VG', t.[Production Date], t.[Insuffisance VG], t.Consommation, t.Limite, t.[Montant VG] " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='VG' AND h.[Production Date]=t.[Production Date]);"

    cn.Execute _
      "INSERT INTO " & T_HIST & " (Booking, [NO DOSSIER CREDIT], Nature, [Production Date], Insuffisance, Consommation, Limite, [Montant]) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'AM', t.[Production Date], t.[Insuffisance AM], t.Consommation, t.Limite, t.[Montant AM] " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='AM' AND h.[Production Date]=t.[Production Date]);"

    cn.Execute _
      "INSERT INTO " & T_HIST & " (Booking, [NO DOSSIER CREDIT], Nature, [Production Date], Insuffisance, Consommation, Limite, [Montant]) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'SL', t.[Production Date], t.[Insuffisance SL], t.Consommation, t.Limite, t.[Montant SL] " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='SL' AND h.[Production Date]=t.[Production Date]);"

    ' Flag "Dépassement non autorisé" = 'YES' quand Abs(Conso) > Limite
    cn.Execute _
      "UPDATE " & T_HIST & " AS h INNER JOIN " & T_TODAY & " AS t " & _
      "ON h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.[Production Date]=t.[Production Date] " & _
      "SET h.DepassementNA='YES' " & _
      "WHERE Abs(t.Consommation) > t.Limite;"

    cn.Close: Set cn = Nothing
End Sub
```

```vb
' ====== Règles de gestion (3-en-1) ======
Private Sub ProcessBusinessRules_3in1_NoNzIif(ByVal dbPath As String, ByRef newFlux As Long, ByRef contFlux As Long, ByRef closedFlux As Long)
    Dim cn As Object: Set cn = GetConn(dbPath)
    On Error GoTo ROLLBACK
    cn.BeginTrans

    Dim nNew As Long, nCont As Long, nClosed As Long

    ' 1) INSERT nouvelles clés FLUX (sans DepassementNA)
    cn.Execute _
      "INSERT INTO " & T_FLUX & " (Booking, [NO DOSSIER CREDIT], Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP]) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], t.Limite, t.Consommation, t.[Montant VG], t.[Montant AM], t.[Montant SL], t.[Production Date], t.[NO INTERVENANT], t.[NO INTERVENANT GRP] " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE (t.[Insuffisance VG] < 0 OR t.[Insuffisance AM] < 0 OR t.[Insuffisance SL] < 0) " & _
      "  AND NOT EXISTS (SELECT 1 FROM " & T_FLUX & " AS f WHERE f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT]);"
    nNew = cn.RecordsAffected

    ' Déterminer DepassementNA en 'YES' / 'NO'
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.DepassementNA='YES' WHERE Abs(t.Consommation) > t.Limite;"
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.DepassementNA='NO' WHERE NOT (Abs(t.Consommation) > t.Limite);"

    ' 2) Démarrages par nature
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_StartDate=t.[Production Date], f.VG_LastUpdateDate=t.[Production Date], f.VG_NbJoursRetard=1, f.VG_WorstInsuffisance=t.[Insuffisance VG], f.VG_DateWorstInsuffisance=t.[Production Date], f.VG_LastInsuffisance=t.[Insuffisance VG] " & _
      "WHERE f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0;"
    nNew = nNew + cn.RecordsAffected

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_StartDate=t.[Production Date], f.AM_LastUpdateDate=t.[Production Date], f.AM_NbJoursRetard=1, f.AM_WorstInsuffisance=t.[Insuffisance AM], f.AM_DateWorstInsuffisance=t.[Production Date], f.AM_LastInsuffisance=t.[Insuffisance AM] " & _
      "WHERE f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0;"
    nNew = nNew + cn.RecordsAffected

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_StartDate=t.[Production Date], f.SL_LastUpdateDate=t.[Production Date], f.SL_NbJoursRetard=1, f.SL_WorstInsuffisance=t.[Insuffisance SL], f.SL_DateWorstInsuffisance=t.[Production Date], f.SL_LastInsuffisance=t.[Insuffisance SL] " & _
      "WHERE f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0;"
    nNew = nNew + cn.RecordsAffected

    ' 3) Continuations
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_NbJoursRetard=f.VG_NbJoursRetard + DateDiff('d', f.VG_LastUpdateDate, t.[Production Date]), " & _
      "    f.VG_LastUpdateDate=t.[Production Date], f.VG_LastInsuffisance=t.[Insuffisance VG], " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP] " & _
      "WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NOT NULL AND t.[Production Date] > f.VG_LastUpdateDate;"
    nCont = cn.RecordsAffected

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_WorstInsuffisance=t.[Insuffisance VG], f.VG_DateWorstInsuffisance=t.[Production Date] " & _
      "WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NOT NULL AND (f.VG_WorstInsuffisance IS NULL OR t.[Insuffisance VG] < f.VG_WorstInsuffisance);"

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_NbJoursRetard=f.AM_NbJoursRetard + DateDiff('d', f.AM_LastUpdateDate, t.[Production Date]), " & _
      "    f.AM_LastUpdateDate=t.[Production Date], f.AM_LastInsuffisance=t.[Insuffisance AM], " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP] " & _
      "WHERE t.[Insuffisance AM] < 0 AND f.AM_StartDate IS NOT NULL AND t.[Production Date] > f.AM_LastUpdateDate;"
    nCont = nCont + cn.RecordsAffected

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_WorstInsuffisance=t.[Insuffisance AM], f.AM_DateWorstInsuffisance=t.[Production Date] " & _
      "WHERE t.[Insuffisance AM] < 0 AND f.AM_StartDate IS NOT NULL AND (f.AM_WorstInsuffisance IS NULL OR t.[Insuffisance AM] < f.AM_WorstInsuffisance);"

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_NbJoursRetard=f.SL_NbJoursRetard + DateDiff('d', f.SL_LastUpdateDate, t.[Production Date]), " & _
      "    f.SL_LastUpdateDate=t.[Production Date], f.SL_LastInsuffisance=t.[Insuffisance SL], " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP] " & _
      "WHERE t.[Insuffisance SL] < 0 AND f.SL_StartDate IS NOT NULL AND t.[Production Date] > f.SL_LastUpdateDate;"
    nCont = nCont + cn.RecordsAffected

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_WorstInsuffisance=t.[Insuffisance SL], f.SL_DateWorstInsuffisance=t.[Production Date] " & _
      "WHERE t.[Insuffisance SL] < 0 AND f.SL_StartDate IS NOT NULL AND (f.SL_WorstInsuffisance IS NULL OR t.[Insuffisance SL] < f.SL_WorstInsuffisance);"

    ' Rafraîchir DepassementNA après continuations
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.DepassementNA='YES' WHERE Abs(t.Consommation) > t.Limite;"
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.DepassementNA='NO' WHERE NOT (Abs(t.Consommation) > t.Limite);"

    ' 4) Archivage
    ' VG
    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], VG_StartDate, VG_EndDate, VG_Duree, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance, VG_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], f.VG_StartDate, f.VG_LastUpdateDate, f.VG_NbJoursRetard, f.VG_WorstInsuffisance, f.VG_DateWorstInsuffisance, f.VG_LastInsuffisance, f.VG_LastUpdateDate, " & _
      "f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], f.[Production Date], f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, 'YES','NO','NO' " & _
      "FROM " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.VG_StartDate IS NOT NULL AND t.Booking IS NULL;"
    nClosed = cn.RecordsAffected

    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], VG_StartDate, VG_EndDate, VG_Duree, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance, VG_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], f.VG_StartDate, t.[Production Date], f.VG_NbJoursRetard, f.VG_WorstInsuffisance, f.VG_DateWorstInsuffisance, f.VG_LastInsuffisance, f.VG_LastUpdateDate, " & _
      "f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], t.[Production Date], f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, 'YES','NO','NO' " & _
      "FROM " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.VG_StartDate IS NOT NULL AND t.[Insuffisance VG] >= 0;"
    nClosed = nClosed + cn.RecordsAffected

    ' AM
    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], AM_StartDate, AM_EndDate, AM_Duree, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance, AM_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], f.AM_StartDate, f.AM_LastUpdateDate, f.AM_NbJoursRetard, f.AM_WorstInsuffisance, f.AM_DateWorstInsuffisance, f.AM_LastInsuffisance, f.AM_LastUpdateDate, " & _
      "f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], f.[Production Date], f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, 'NO','YES','NO' " & _
      "FROM " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.AM_StartDate IS NOT NULL AND t.Booking IS NULL;"
    nClosed = nClosed + cn.RecordsAffected

    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], AM_StartDate, AM_EndDate, AM_Duree, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance, AM_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], f.AM_StartDate, t.[Production Date], f.AM_NbJoursRetard, f.AM_WorstInsuffisance, f.AM_DateWorstInsuffisance, f.AM_LastInsuffisance, f.AM_LastUpdateDate, " & _
      "f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], t.[Production Date], f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, 'NO','YES','NO' " & _
      "FROM " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.AM_StartDate IS NOT NULL AND t.[Insuffisance AM] >= 0;"
    nClosed = nClosed + cn.RecordsAffected

    ' SL
    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], SL_StartDate, SL_EndDate, SL_Duree, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance, SL_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], f.SL_StartDate, f.SL_LastUpdateDate, f.SL_NbJoursRetard, f.SL_WorstInsuffisance, f.SL_DateWorstInsuffisance, f.SL_LastInsuffisance, f.SL_LastUpdateDate, " & _
      "f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], f.[Production Date], f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, 'NO','NO','YES' " & _
      "FROM " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.SL_StartDate IS NOT NULL AND t.Booking IS NULL;"
    nClosed = nClosed + cn.RecordsAffected

    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], SL_StartDate, SL_EndDate, SL_Duree, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance, SL_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], f.SL_StartDate, t.[Production Date], f.SL_NbJoursRetard, f.SL_WorstInsuffisance, f.SL_DateWorstInsuffisance, f.SL_LastInsuffisance, f.SL_LastUpdateDate, " & _
      "f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], t.[Production Date], f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, 'NO','NO','YES' " & _
      "FROM " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.SL_StartDate IS NOT NULL AND t.[Insuffisance SL] >= 0;"
    nClosed = nClosed + cn.RecordsAffected

    ' 5) Flags multi-natures via historique (passage à 'YES')
    cn.Execute _
      "UPDATE " & T_ARC & " AS a SET a.FlagAM='YES' " & _
      "WHERE a.FlagVG='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=a.Booking AND h.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] AND h.Nature='AM' " & _
      "    AND h.[Production Date] BETWEEN a.VG_StartDate AND a.VG_EndDate AND h.Insuffisance < 0" & _
      ");"
    cn.Execute _
      "UPDATE " & T_ARC & " AS a SET a.FlagSL='YES' " & _
      "WHERE a.FlagVG='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=a.Booking AND h.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] AND h.Nature='SL' " & _
      "    AND h.[Production Date] BETWEEN a.VG_StartDate AND a.VG_EndDate AND h.Insuffisance < 0" & _
      ");"

    cn.Execute _
      "UPDATE " & T_ARC & " AS a SET a.FlagVG='YES' " & _
      "WHERE a.FlagAM='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=a.Booking AND h.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] AND h.Nature='VG' " & _
      "    AND h.[Production Date] BETWEEN a.AM_StartDate AND a.AM_EndDate AND h.Insuffisance < 0" & _
      ");"
    cn.Execute _
      "UPDATE " & T_ARC & " AS a SET a.FlagSL='YES' " & _
      "WHERE a.FlagAM='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=a.Booking AND h.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] AND h.Nature='SL' " & _
      "    AND h.[Production Date] BETWEEN a.AM_StartDate AND a.AM_EndDate AND h.Insuffisance < 0" & _
      ");"

    cn.Execute _
      "UPDATE " & T_ARC & " AS a SET a.FlagVG='YES' " & _
      "WHERE a.FlagSL='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=a.Booking AND h.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] AND h.Nature='VG' " & _
      "    AND h.[Production Date] BETWEEN a.SL_StartDate AND a.SL_EndDate AND h.Insuffisance < 0" & _
      ");"
    cn.Execute _
      "UPDATE " & T_ARC & " AS a SET a.FlagAM='YES' " & _
      "WHERE a.FlagSL='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=a.Booking AND h.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] AND h.Nature='AM' " & _
      "    AND h.[Production Date] BETWEEN a.SL_StartDate AND a.SL_EndDate AND h.Insuffisance < 0" & _
      ");"

    ' 6) Purge FLUX
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_StartDate=Null, f.VG_LastUpdateDate=Null, f.VG_NbJoursRetard=Null, f.VG_WorstInsuffisance=Null, f.VG_DateWorstInsuffisance=Null, f.VG_LastInsuffisance=Null, f.VG_EndDate=Null " & _
      "WHERE f.VG_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0);"
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_StartDate=Null, f.AM_LastUpdateDate=Null, f.AM_NbJoursRetard=Null, f.AM_WorstInsuffisance=Null, f.AM_DateWorstInsuffisance=Null, f.AM_LastInsuffisance=Null, f.AM_EndDate=Null " & _
      "WHERE f.AM_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0);"
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_StartDate=Null, f.SL_LastUpdateDate=Null, f.SL_NbJoursRetard=Null, f.SL_WorstInsuffisance=Null, f.SL_DateWorstInsuffisance=Null, f.SL_LastInsuffisance=Null, f.SL_EndDate=Null " & _
      "WHERE f.SL_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0);"

    cn.Execute "DELETE FROM " & T_FLUX & " WHERE VG_StartDate IS NULL AND AM_StartDate IS NULL AND SL_StartDate IS NULL;"

    cn.CommitTrans
    newFlux = nNew: contFlux = nCont: closedFlux = nClosed
    cn.Close: Set cn = Nothing
    Exit Sub
ROLLBACK:
    cn.RollbackTrans
    cn.Close: Set cn = Nothing
    Err.Raise Err.Number, , Err.Description
End Sub
```

```vb
' ====== Enrichissement feuille (3 × {NbJ, Début, Fin} + Dépassement 'YES'/'NO') ======
Private Sub EnrichSheetFromFlux_3in1(ByVal dbPath As String, ByVal ws As Worksheet)
    Dim cNBJ_VG As Long, cDEB_VG As Long, cFIN_VG As Long
    Dim cNBJ_AM As Long, cDEB_AM As Long, cFIN_AM As Long
    Dim cNBJ_SL As Long, cDEB_SL As Long, cFIN_SL As Long
    Dim cNA As Long
    cNBJ_VG = EnsureColumn(ws, COL_NBJ_VG): cDEB_VG = EnsureColumn(ws, COL_DEB_VG): cFIN_VG = EnsureColumn(ws, COL_FIN_VG)
    cNBJ_AM = EnsureColumn(ws, COL_NBJ_AM): cDEB_AM = EnsureColumn(ws, COL_DEB_AM): cFIN_AM = EnsureColumn(ws, COL_FIN_AM)
    cNBJ_SL = EnsureColumn(ws, COL_NBJ_SL): cDEB_SL = EnsureColumn(ws, COL_DEB_SL): cFIN_SL = EnsureColumn(ws, COL_FIN_SL)
    cNA = EnsureColumn(ws, COL_FLAG_NA)

    Dim prodDateSheet As Variant
    prodDateSheet = DetectSheetProductionDate(ws)

    Dim dVG As Object, dAM As Object, dSL As Object
    Set dVG = LoadFluxDict_3in1(dbPath, "VG")
    Set dAM = LoadFluxDict_3in1(dbPath, "AM")
    Set dSL = LoadFluxDict_3in1(dbPath, "SL")

    Dim aVG As Object, aAM As Object, aSL As Object
    Set aVG = LoadArchiveEndDictForDate_3in1(dbPath, "VG", prodDateSheet)
    Set aAM = LoadArchiveEndDictForDate_3in1(dbPath, "AM", prodDateSheet)
    Set aSL = LoadArchiveEndDictForDate_3in1(dbPath, "SL", prodDateSheet)

    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    MapCol map, ws, COL_BOOKING
    MapCol map, ws, COL_DOSSIER
    MapCol map, ws, COL_CONSO
    MapCol map, ws, COL_LIMITE

    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        Dim bk As String: bk = Trim$(UCase$(CStr(ws.Cells(r, map(COL_BOOKING)).Value)))
        Dim dos As String: dos = Trim$(CStr(ws.Cells(r, map(COL_DOSSIER)).Value))
        Dim k As String: k = bk & "||" & dos

        Dim conso As Double, lim As Double
        conso = Val(ws.Cells(r, map(COL_CONSO)).Value)
        lim = Val(ws.Cells(r, map(COL_LIMITE)).Value)
        If Abs(conso) > lim Then
            ws.Cells(r, cNA).Value = "YES"
        Else
            ws.Cells(r, cNA).Value = "NO"
        End If

        If dVG.Exists(k) Then
            ws.Cells(r, cNBJ_VG).Value = dVG(k)(1)
            ws.Cells(r, cDEB_VG).Value = dVG(k)(0)
        Else
            ws.Cells(r, cNBJ_VG).ClearContents
            ws.Cells(r, cDEB_VG).ClearContents
        End If
        If aVG.Exists(k) Then ws.Cells(r, cFIN_VG).Value = aVG(k) Else ws.Cells(r, cFIN_VG).ClearContents

        If dAM.Exists(k) Then
            ws.Cells(r, cNBJ_AM).Value = dAM(k)(1)
            ws.Cells(r, cDEB_AM).Value = dAM(k)(0)
        Else
            ws.Cells(r, cNBJ_AM).ClearContents
            ws.Cells(r, cDEB_AM).ClearContents
        End If
        If aAM.Exists(k) Then ws.Cells(r, cFIN_AM).Value = aAM(k) Else ws.Cells(r, cFIN_AM).ClearContents

        If dSL.Exists(k) Then
            ws.Cells(r, cNBJ_SL).Value = dSL(k)(1)
            ws.Cells(r, cDEB_SL).Value = dSL(k)(0)
        Else
            ws.Cells(r, cNBJ_SL).ClearContents
            ws.Cells(r, cDEB_SL).ClearContents
        End If
        If aSL.Exists(k) Then ws.Cells(r, cFIN_SL).Value = aSL(k) Else ws.Cells(r, cFIN_SL).ClearContents
    Next r

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

```vb
' ====== Helpers ======
Private Sub MapCol(ByVal dict As Object, ByVal ws As Worksheet, ByVal name1 As String, Optional ByVal name2 As String = "")
    Dim c As Long: c = FindHeader(ws, name1)
    If c = 0 And Len(name2) > 0 Then c = FindHeader(ws, name2)
    If c = 0 Then Err.Raise vbObjectError + 10, , "Colonne introuvable : " & name1 & IIf(Len(name2) > 0, " (alt : " & name2 & ")", "")
    dict(AddKey(dict, name1)) = c
End Sub

Private Function AddKey(ByVal dict As Object, ByVal k As String) As String
    If Not dict.Exists(k) Then dict.Add k, 0
    AddKey = k
End Function

Private Function FindHeader(ByVal ws As Worksheet, ByVal hdr As String) As Long
    Dim c As Range
    For Each c In ws.Rows(1).Cells
        If Trim$(UCase$(c.Value & "")) = Trim$(UCase$(hdr)) Then FindHeader = c.Column: Exit Function
        If Len(c.Value & "") = 0 Then Exit For
    Next c
End Function

Private Function GetCellMaybe(ws As Worksheet, r As Long, map As Object, name1 As String, name2 As String) As Variant
    If map.Exists(name1) Then GetCellMaybe = ws.Cells(r, map(name1)).Value: Exit Function
    If map.Exists(name2) Then GetCellMaybe = ws.Cells(r, map(name2)).Value: Exit Function
    GetCellMaybe = Null
End Function

Private Function EnsureColumn(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim c As Long: c = FindHeader(ws, header)
    If c > 0 Then EnsureColumn = c: Exit Function
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ws.Cells(1, lastCol + 1).Value = header
    EnsureColumn = lastCol + 1
End Function

' ====== Logging ======
Private Sub WriteRunLog(ByVal dbPath As String, ByVal rowsImp As Long, ByVal newF As Long, ByVal contF As Long, ByVal closedF As Long)
    Dim cn As Object: Set cn = GetConn(dbPath)
    cn.Execute "INSERT INTO " & T_LOG & " (RunAt, RowsImported, NewFlux, Continued, Closed, [User]) VALUES (Now()," & rowsImp & "," & newF & "," & contF & "," & closedF & ",'" & Environ$("Username") & "')"
    cn.Close: Set cn = Nothing
End Sub
```

---

## 📊 Exemples de requêtes de reporting

- **Flux actuel par nature (VG/AM/SL ouverts)**  
```sql
SELECT Booking, [NO DOSSIER CREDIT],
       VG_StartDate, VG_NbJoursRetard, VG_WorstInsuffisance,
       AM_StartDate, AM_NbJoursRetard, AM_WorstInsuffisance,
       SL_StartDate, SL_NbJoursRetard, SL_WorstInsuffisance,
       DepassementNA
FROM T_Flux3
WHERE VG_StartDate IS NOT NULL OR AM_StartDate IS NOT NULL OR SL_StartDate IS NOT NULL
ORDER BY VG_NbJoursRetard DESC, AM_NbJoursRetard DESC, SL_NbJoursRetard DESC;
```

- **Clôtures par mois & nature (avec flags 'YES'/'NO')**  
```sql
SELECT Format(Nz(VG_EndDate, Nz(AM_EndDate, SL_EndDate)),'yyyy-mm') AS Mois,
       SUM(IIF(FlagVG='YES',1,0)) AS Nb_VG,
       SUM(IIF(FlagAM='YES',1,0)) AS Nb_AM,
       SUM(IIF(FlagSL='YES',1,0)) AS Nb_SL,
       SUM(Nz(VG_Duree,0) + Nz(AM_Duree,0) + Nz(SL_Duree,0)) AS JoursCumules
FROM T_Archive3
GROUP BY Format(Nz(VG_EndDate, Nz(AM_EndDate, SL_EndDate)),'yyyy-mm')
ORDER BY Mois DESC;
```

- **Série temps insuffisance VG par clé**  
```sql
SELECT [Production Date], Insuffisance
FROM T_HistoryInsuff
WHERE Booking='CAI' AND [NO DOSSIER CREDIT]='123456' AND Nature='VG'
ORDER BY [Production Date];
```

---

## ✅ Points de robustesse
- Pas de `Nz()`/`IIf()` **dans le SQL** ; valeurs `NULL` neutralisées en amont via `NormalizeToday()`.
- Incrément **idempotent** via `DateDiff('d', LastUpdateDate, ProductionDate)`.
- Booléens **texte** `'YES'/'NO'` — évite les cast implicites Access.
- Index sur dates de start/end pour reporting performant.
- Historique **T_HistoryInsuff** pour reconstituer des **épisodes multi-natures** et stats fines.

---

## 🚀 Mise en place
1. Mets le classeur dans un dossier **écrivable** (la base `.accdb` sera créée à côté).
2. Colle le module VBA ci-dessus dans **VBA → Module standard**.
3. Place un bouton et assigne **`MAJ_Retards_Daily_3in1_NoNzIif`**.
4. Vérifie les en-têtes (les variantes « NO NTERVENANT » et « Montant VG (valeur de gage-LTV) » sont gérées).

Bon run !
