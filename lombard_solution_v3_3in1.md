# Solution *3-en-1* (VG/AM/SL) — Access + VBA
> **Objectif (en une phrase)** : Historiser **par jour**, pour chaque clé **(Booking + NO DOSSIER CREDIT)**, les **jours de retard** pour **VG / AM / SL**, avec **3 compteurs** et **3 couples Start/End** (et *Worst/Last/DateWorst* pour chacun), **flux en cours**, **archives**, **time series journalière**, et **enrichissement** automatique de la feuille Excel (9 colonnes + flag dépassement NA).

---

## 1) Architecture choisie (respect stricte “3 compteurs / 3 dates”)

### Principe
- **1 ligne par clé (Booking + Dossier) dans le flux**, contenant **3 sous-épisodes** simultanés : **VG**, **AM**, **SL**.  
  Chaque sous-épisode a ses champs **StartDate / LastUpdateDate / NbJoursRetard / WorstInsuffisance / DateWorstInsuffisance / LastInsuffisance / EndDate**.
- **Archives** : lorsqu’un sous-épisode se clôture (VG ou AM ou SL), on **insère une ligne** d’archive avec les champs remplis **uniquement** pour la/les natures clôturées et des **flags** `FlagVG/FlagAM/FlagSL` pour qualifier si, **durant l’épisode**, les autres natures ont aussi connu au moins un jour **< 0**.
- **Time series** (T_HistoryInsuff) : 1 point / jour / nature — sert au reporting et au calcul des flags d’archive.

> ⚠️ **EndDate** : en **flux**, une EndDate n’a du sens qu’au **jour de la clôture**. On la garde **NULL** jusqu’à la fermeture. En **archive**, les EndDate des natures clôturées sont **renseignées** (et non NULL).

---

## 2) Schéma Access (DDL) — création automatique

> La base `Lombard_Retards.accdb` est créée et provisionnée automatiquement par le VBA.

```sql
-- =============================================
-- FLUX (1 ligne par clé, 3 sous-épisodes VG/AM/SL)
-- =============================================
CREATE TABLE T_Flux3 (
  Booking TEXT(20) NOT NULL,
  [NO DOSSIER CREDIT] TEXT(50) NOT NULL,

  -- VG
  VG_StartDate DATE,
  VG_LastUpdateDate DATE,
  VG_NbJoursRetard LONG,
  VG_WorstInsuffisance DOUBLE,
  VG_DateWorstInsuffisance DATE,
  VG_LastInsuffisance DOUBLE,
  VG_EndDate DATE,

  -- AM
  AM_StartDate DATE,
  AM_LastUpdateDate DATE,
  AM_NbJoursRetard LONG,
  AM_WorstInsuffisance DOUBLE,
  AM_DateWorstInsuffisance DATE,
  AM_LastInsuffisance DOUBLE,
  AM_EndDate DATE,

  -- SL
  SL_StartDate DATE,
  SL_LastUpdateDate DATE,
  SL_NbJoursRetard LONG,
  SL_WorstInsuffisance DOUBLE,
  SL_DateWorstInsuffisance DATE,
  SL_LastInsuffisance DOUBLE,
  SL_EndDate DATE,

  -- Contexte dernier vu (communs)
  Limite DOUBLE,
  Consommation DOUBLE,
  [Montant VG] DOUBLE,
  [Montant AM] DOUBLE,
  [Montant SL] DOUBLE,
  [Production Date] DATE,
  [NO INTERVENANT] TEXT(50),
  [NO INTERVENANT GRP] TEXT(50),
  DepassementNA YESNO,

  CONSTRAINT PK_Flux3 PRIMARY KEY (Booking, [NO DOSSIER CREDIT])
);
CREATE INDEX IX_Flux3_VG_Start ON T_Flux3 (VG_StartDate);
CREATE INDEX IX_Flux3_AM_Start ON T_Flux3 (AM_StartDate);
CREATE INDEX IX_Flux3_SL_Start ON T_Flux3 (SL_StartDate);

-- =============================================
-- ARCHIVES (1 ligne insérée à chaque clôture d’un sous-épisode)
--   - Les colonnes d’une nature sont remplies si et seulement si FlagNature=True
-- =============================================
CREATE TABLE T_Archive3 (
  ID AUTOINCREMENT PRIMARY KEY,
  Booking TEXT(20) NOT NULL,
  [NO DOSSIER CREDIT] TEXT(50) NOT NULL,

  -- Épisode VG archivé (si FlagVG=True)
  VG_StartDate DATE,
  VG_EndDate DATE,
  VG_Duree LONG,
  VG_WorstInsuffisance DOUBLE,
  VG_DateWorstInsuffisance DATE,
  VG_LastInsuffisance DOUBLE,
  VG_LastUpdateDate DATE,

  -- Épisode AM archivé (si FlagAM=True)
  AM_StartDate DATE,
  AM_EndDate DATE,
  AM_Duree LONG,
  AM_WorstInsuffisance DOUBLE,
  AM_DateWorstInsuffisance DATE,
  AM_LastInsuffisance DOUBLE,
  AM_LastUpdateDate DATE,

  -- Épisode SL archivé (si FlagSL=True)
  SL_StartDate DATE,
  SL_EndDate DATE,
  SL_Duree LONG,
  SL_WorstInsuffisance DOUBLE,
  SL_DateWorstInsuffisance DATE,
  SL_LastInsuffisance DOUBLE,
  SL_LastUpdateDate DATE,

  -- Contexte (copie dernier vu)
  Limite DOUBLE,
  Consommation DOUBLE,
  [Montant VG] DOUBLE,
  [Montant AM] DOUBLE,
  [Montant SL] DOUBLE,
  [Production Date] DATE,
  [NO INTERVENANT] TEXT(50),
  [NO INTERVENANT GRP] TEXT(50),
  DepassementNA YESNO,

  -- Qualification de l’épisode (un / deux / trois)
  FlagVG YESNO,
  FlagAM YESNO,
  FlagSL YESNO,

  ArchivedAt DATE DEFAULT Date()
);
CREATE INDEX IX_Arc3_Key ON T_Archive3 (Booking, [NO DOSSIER CREDIT]);
CREATE INDEX IX_Arc3_VG_End ON T_Archive3 (VG_EndDate);
CREATE INDEX IX_Arc3_AM_End ON T_Archive3 (AM_EndDate);
CREATE INDEX IX_Arc3_SL_End ON T_Archive3 (SL_EndDate);

-- =============================================
-- STAGING (miroir de la feuille Excel)
-- =============================================
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

-- =============================================
-- TIME SERIES JOURNALIÈRE (1 point / jour / nature)
-- =============================================
CREATE TABLE T_HistoryInsuff (
  Booking TEXT(20) NOT NULL,
  [NO DOSSIER CREDIT] TEXT(50) NOT NULL,
  Nature TEXT(2) NOT NULL,            -- 'VG'/'AM'/'SL'
  [Production Date] DATE NOT NULL,
  Insuffisance DOUBLE,
  Consommation DOUBLE,
  Limite DOUBLE,
  [Montant] DOUBLE,
  DepassementNA YESNO,
  CONSTRAINT IX_Hist UNIQUE (Booking, [NO DOSSIER CREDIT], Nature, [Production Date])
);

-- =============================================
-- RUN LOG
-- =============================================
CREATE TABLE T_RunLog (
  RunAt DATETIME NOT NULL,
  RowsImported LONG,
  NewFlux LONG,
  Continued LONG,
  Closed LONG,
  [User] TEXT(100)
);
```

> **Remarque sur “NOT NULL”** : En **archive**, les champs `*_EndDate` de la/les natures clôturées sont **renseignés** et donc **non nuls** *logiquement* (contrôle métier). En **flux**, ils restent **NULL** jusqu’à la clôture.

---

## 3) Macro VBA complète (1 bouton/jour)

> Tout en **late binding** (ADO/ADOX). Colle le module ci-dessous dans *VBA → Module standard* et assigne **`MAJ_Retards_Daily_3in1`** à ton bouton.

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
Public Sub MAJ_Retards_Daily_3in1()
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

    ' 3) Historique quotidien (time series)
    AppendHistory dbPath

    ' 4) Règles métier 3-en-1
    Dim newFlux As Long, contFlux As Long, closedFlux As Long
    ProcessBusinessRules_3in1 dbPath, newFlux, contFlux, closedFlux

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
      "Limite DOUBLE, Consommation DOUBLE, [Montant VG] DOUBLE, [Montant AM] DOUBLE, [Montant SL] DOUBLE, [Production Date] DATE, [NO INTERVENANT] TEXT(50), [NO INTERVENANT GRP] TEXT(50), DepassementNA YESNO, " & _
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
      "Limite DOUBLE, Consommation DOUBLE, [Montant VG] DOUBLE, [Montant AM] DOUBLE, [Montant SL] DOUBLE, [Production Date] DATE, [NO INTERVENANT] TEXT(50), [NO INTERVENANT GRP] TEXT(50), DepassementNA YESNO, " & _
      "FlagVG YESNO, FlagAM YESNO, FlagSL YESNO, ArchivedAt DATE DEFAULT Date());" & _
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
      "Insuffisance DOUBLE, Consommation DOUBLE, Limite DOUBLE, [Montant] DOUBLE, DepassementNA YESNO, " & _
      "CONSTRAINT IX_Hist UNIQUE (Booking, [NO DOSSIER CREDIT], Nature, [Production Date]) );"
End Function

Private Function GetDDL_Log() As String
    GetDDL_Log = _
      "CREATE TABLE " & T_LOG & " (" & _
      "RunAt DATETIME NOT NULL, RowsImported LONG, NewFlux LONG, Continued LONG, Closed LONG, [User] TEXT(100));"
End Function

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
        cmd.Parameters(0).Value = NzS(ws.Cells(r, map(COL_BOOKING)).Value)
        cmd.Parameters(1).Value = NzS(ws.Cells(r, map(COL_PROD)).Value)
        cmd.Parameters(2).Value = NzS(ws.Cells(r, map(COL_DOSSIER)).Value)
        cmd.Parameters(3).Value = NzS(GetCellMaybe(ws, r, map, COL_INTER, COL_INTER_ALT))
        cmd.Parameters(4).Value = NzS(GetCellMaybe(ws, r, map, COL_INTER_GRP, COL_INTER_GRP_ALT))
        cmd.Parameters(5).Value = NzS(ws.Cells(r, map(COL_LIMITE)).Value)
        cmd.Parameters(6).Value = NzS(ws.Cells(r, map(COL_CONSO)).Value)
        cmd.Parameters(7).Value = NzS(GetCellMaybe(ws, r, map, COL_MVG, COL_MVG_ALT))
        cmd.Parameters(8).Value = NzS(ws.Cells(r, map(COL_MAM)).Value)
        cmd.Parameters(9).Value = NzS(ws.Cells(r, map(COL_MSL)).Value)
        cmd.Parameters(10).Value = NzS(ws.Cells(r, map(COL_IVG)).Value)
        cmd.Parameters(11).Value = NzS(ws.Cells(r, map(COL_IAM)).Value)
        cmd.Parameters(12).Value = NzS(ws.Cells(r, map(COL_ISL)).Value)
        cmd.Execute , , 128
        n = n + 1
    Next r
    cn.CommitTrans

    ' Normalisation clés
    cn.Execute "UPDATE " & T_TODAY & " SET Booking = UCase(Trim(Booking)), [NO DOSSIER CREDIT] = Trim([NO DOSSIER CREDIT]);"

    ' Dédoublonnage : ne garder que la ligne la plus récente par clé
    cn.Execute _
      "DELETE FROM " & T_TODAY & " AS a " & _
      "WHERE EXISTS (SELECT 1 FROM " & T_TODAY & " AS b " & _
      "              WHERE b.Booking=a.Booking AND b.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT] " & _
      "                AND Nz(b.[Production Date],#1900-01-01#) > Nz(a.[Production Date],#1900-01-01#));"

    LoadTodayIntoAccess = n
ExitHere:
    cn.Close: Set cn = Nothing
End Function

' ====== Historique quotidien ======
Private Sub AppendHistory(ByVal dbPath As String)
    Dim cn As Object: Set cn = GetConn(dbPath)

    ' VG
    cn.Execute _
      "INSERT INTO " & T_HIST & " (Booking, [NO DOSSIER CREDIT], Nature, [Production Date], Insuffisance, Consommation, Limite, [Montant], DepassementNA) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'VG', Nz(t.[Production Date], Date()), t.[Insuffisance VG], t.Consommation, t.Limite, t.[Montant VG], IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='VG' AND h.[Production Date]=Nz(t.[Production Date], Date()));"

    ' AM
    cn.Execute _
      "INSERT INTO " & T_HIST & " (Booking, [NO DOSSIER CREDIT], Nature, [Production Date], Insuffisance, Consommation, Limite, [Montant], DepassementNA) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'AM', Nz(t.[Production Date], Date()), t.[Insuffisance AM], t.Consommation, t.Limite, t.[Montant AM], IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='AM' AND h.[Production Date]=Nz(t.[Production Date], Date()));"

    ' SL
    cn.Execute _
      "INSERT INTO " & T_HIST & " (Booking, [NO DOSSIER CREDIT], Nature, [Production Date], Insuffisance, Consommation, Limite, [Montant], DepassementNA) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], 'SL', Nz(t.[Production Date], Date()), t.[Insuffisance SL], t.Consommation, t.Limite, t.[Montant SL], IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE NOT EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=t.Booking AND h.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] AND h.Nature='SL' AND h.[Production Date]=Nz(t.[Production Date], Date()));"

    cn.Close: Set cn = Nothing
End Sub

' ====== Règles de gestion (3-en-1) ======
Private Sub ProcessBusinessRules_3in1(ByVal dbPath As String, ByRef newFlux As Long, ByRef contFlux As Long, ByRef closedFlux As Long)
    Dim cn As Object: Set cn = GetConn(dbPath)
    On Error GoTo ROLLBACK
    cn.BeginTrans

    Dim nNew As Long, nCont As Long, nClosed As Long

    ' --- Insérer lignes FLUX nouvelles (clé absente) si au moins une insuff < 0 ---
    cn.Execute _
      "INSERT INTO " & T_FLUX & " (" & _
      "Booking, [NO DOSSIER CREDIT], " & _
      "VG_StartDate, VG_LastUpdateDate, VG_NbJoursRetard, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance, " & _
      "AM_StartDate, AM_LastUpdateDate, AM_NbJoursRetard, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance, " & _
      "SL_StartDate, SL_LastUpdateDate, SL_NbJoursRetard, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA) " & _
      "SELECT t.Booking, t.[NO DOSSIER CREDIT], " & _
      "IIF(t.[Insuffisance VG] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance VG] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance VG] < 0, 1, 0), IIF(t.[Insuffisance VG] < 0, t.[Insuffisance VG], Null), IIF(t.[Insuffisance VG] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance VG] < 0, t.[Insuffisance VG], Null), " & _
      "IIF(t.[Insuffisance AM] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance AM] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance AM] < 0, 1, 0), IIF(t.[Insuffisance AM] < 0, t.[Insuffisance AM], Null), IIF(t.[Insuffisance AM] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance AM] < 0, t.[Insuffisance AM], Null), " & _
      "IIF(t.[Insuffisance SL] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance SL] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance SL] < 0, 1, 0), IIF(t.[Insuffisance SL] < 0, t.[Insuffisance SL], Null), IIF(t.[Insuffisance SL] < 0, Nz(t.[Production Date], Date()), Null), IIF(t.[Insuffisance SL] < 0, t.[Insuffisance SL], Null), " & _
      "t.Limite, t.Consommation, t.[Montant VG], t.[Montant AM], t.[Montant SL], t.[Production Date], t.[NO INTERVENANT], t.[NO INTERVENANT GRP], IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "FROM " & T_TODAY & " AS t LEFT JOIN " & T_FLUX & " AS f " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.Booking IS NULL AND (t.[Insuffisance VG] < 0 OR t.[Insuffisance AM] < 0 OR t.[Insuffisance SL] < 0);"
    nNew = cn.RecordsAffected

    ' --- Démarrer un sous-épisode pour une clé déjà en flux (VG/AM/SL) ---
    ' VG start
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_StartDate = IIF(f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0, Nz(t.[Production Date], Date()), f.VG_StartDate), " & _
      "    f.VG_LastUpdateDate = IIF(f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0, Nz(t.[Production Date], Date()), f.VG_LastUpdateDate), " & _
      "    f.VG_NbJoursRetard = IIF(f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0, 1, f.VG_NbJoursRetard), " & _
      "    f.VG_WorstInsuffisance = IIF(f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0, t.[Insuffisance VG], f.VG_WorstInsuffisance), " & _
      "    f.VG_DateWorstInsuffisance = IIF(f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0, Nz(t.[Production Date], Date()), f.VG_DateWorstInsuffisance) " & _
      "WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NULL;"
    nNew = nNew + cn.RecordsAffected

    ' AM start
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_StartDate = IIF(f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0, Nz(t.[Production Date], Date()), f.AM_StartDate), " & _
      "    f.AM_LastUpdateDate = IIF(f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0, Nz(t.[Production Date], Date()), f.AM_LastUpdateDate), " & _
      "    f.AM_NbJoursRetard = IIF(f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0, 1, f.AM_NbJoursRetard), " & _
      "    f.AM_WorstInsuffisance = IIF(f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0, t.[Insuffisance AM], f.AM_WorstInsuffisance), " & _
      "    f.AM_DateWorstInsuffisance = IIF(f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0, Nz(t.[Production Date], Date()), f.AM_DateWorstInsuffisance) " & _
      "WHERE t.[Insuffisance AM] < 0 AND f.AM_StartDate IS NULL;"
    nNew = nNew + cn.RecordsAffected

    ' SL start
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_StartDate = IIF(f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0, Nz(t.[Production Date], Date()), f.SL_StartDate), " & _
      "    f.SL_LastUpdateDate = IIF(f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0, Nz(t.[Production Date], Date()), f.SL_LastUpdateDate), " & _
      "    f.SL_NbJoursRetard = IIF(f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0, 1, f.SL_NbJoursRetard), " & _
      "    f.SL_WorstInsuffisance = IIF(f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0, t.[Insuffisance SL], f.SL_WorstInsuffisance), " & _
      "    f.SL_DateWorstInsuffisance = IIF(f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0, Nz(t.[Production Date], Date()), f.SL_DateWorstInsuffisance) " & _
      "WHERE t.[Insuffisance SL] < 0 AND f.SL_StartDate IS NULL;"
    nNew = nNew + cn.RecordsAffected

    ' --- Continuer les sous-épisodes ouverts (incrément par écart de jours) ---
    ' VG continue
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_NbJoursRetard = f.VG_NbJoursRetard + IIF(DateDiff('d', Nz(f.VG_LastUpdateDate, Nz(t.[Production Date], Date())), Nz(t.[Production Date], Date()))>0, DateDiff('d', Nz(f.VG_LastUpdateDate, Nz(t.[Production Date], Date())), Nz(t.[Production Date], Date())), 0), " & _
      "    f.VG_LastUpdateDate = IIF(Nz(t.[Production Date], Date())>Nz(f.VG_LastUpdateDate,#1/1/1900#), Nz(t.[Production Date], Date()), f.VG_LastUpdateDate), " & _
      "    f.VG_LastInsuffisance = t.[Insuffisance VG], " & _
      "    f.VG_WorstInsuffisance = IIF(f.VG_WorstInsuffisance IS NULL OR t.[Insuffisance VG] < f.VG_WorstInsuffisance, t.[Insuffisance VG], f.VG_WorstInsuffisance), " & _
      "    f.VG_DateWorstInsuffisance = IIF(f.VG_WorstInsuffisance IS NULL OR t.[Insuffisance VG] < f.VG_WorstInsuffisance, Nz(t.[Production Date], Date()), f.VG_DateWorstInsuffisance), " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP], f.DepassementNA=IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NOT NULL;"
    nCont = cn.RecordsAffected

    ' AM continue
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_NbJoursRetard = f.AM_NbJoursRetard + IIF(DateDiff('d', Nz(f.AM_LastUpdateDate, Nz(t.[Production Date], Date())), Nz(t.[Production Date], Date()))>0, DateDiff('d', Nz(f.AM_LastUpdateDate, Nz(t.[Production Date], Date())), Nz(t.[Production Date], Date())), 0), " & _
      "    f.AM_LastUpdateDate = IIF(Nz(t.[Production Date], Date())>Nz(f.AM_LastUpdateDate,#1/1/1900#), Nz(t.[Production Date], Date()), f.AM_LastUpdateDate), " & _
      "    f.AM_LastInsuffisance = t.[Insuffisance AM], " & _
      "    f.AM_WorstInsuffisance = IIF(f.AM_WorstInsuffisance IS NULL OR t.[Insuffisance AM] < f.AM_WorstInsuffisance, t.[Insuffisance AM], f.AM_WorstInsuffisance), " & _
      "    f.AM_DateWorstInsuffisance = IIF(f.AM_WorstInsuffisance IS NULL OR t.[Insuffisance AM] < f.AM_WorstInsuffisance, Nz(t.[Production Date], Date()), f.AM_DateWorstInsuffisance), " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP], f.DepassementNA=IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "WHERE t.[Insuffisance AM] < 0 AND f.AM_StartDate IS NOT NULL;"
    nCont = nCont + cn.RecordsAffected

    ' SL continue
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_NbJoursRetard = f.SL_NbJoursRetard + IIF(DateDiff('d', Nz(f.SL_LastUpdateDate, Nz(t.[Production Date], Date())), Nz(t.[Production Date], Date()))>0, DateDiff('d', Nz(f.SL_LastUpdateDate, Nz(t.[Production Date], Date())), Nz(t.[Production Date], Date())), 0), " & _
      "    f.SL_LastUpdateDate = IIF(Nz(t.[Production Date], Date())>Nz(f.SL_LastUpdateDate,#1/1/1900#), Nz(t.[Production Date], Date()), f.SL_LastUpdateDate), " & _
      "    f.SL_LastInsuffisance = t.[Insuffisance SL], " & _
      "    f.SL_WorstInsuffisance = IIF(f.SL_WorstInsuffisance IS NULL OR t.[Insuffisance SL] < f.SL_WorstInsuffisance, t.[Insuffisance SL], f.SL_WorstInsuffisance), " & _
      "    f.SL_DateWorstInsuffisance = IIF(f.SL_WorstInsuffisance IS NULL OR t.[Insuffisance SL] < f.SL_WorstInsuffisance, Nz(t.[Production Date], Date()), f.SL_DateWorstInsuffisance), " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP], f.DepassementNA=IIF(Abs(Nz(t.Consommation,0))>Nz(t.Limite,0), True, False) " & _
      "WHERE t.[Insuffisance SL] < 0 AND f.SL_StartDate IS NOT NULL;"
    nCont = nCont + cn.RecordsAffected

    ' --- Clôtures & archivage (VG/AM/SL) ---
    ' VG -> Archive
    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], " & _
      "VG_StartDate, VG_EndDate, VG_Duree, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance, VG_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], " & _
      "       f.VG_StartDate, IIF(t.Booking IS NULL, f.VG_LastUpdateDate, Nz(t.[Production Date], Date())) AS EndDate, " & _
      "       f.VG_NbJoursRetard, f.VG_WorstInsuffisance, f.VG_DateWorstInsuffisance, f.VG_LastInsuffisance, f.VG_LastUpdateDate, " & _
      "       f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], Nz(t.[Production Date], f.VG_LastUpdateDate), f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, " & _
      "       True, " & _
      "       IIF(EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=f.Booking AND h.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT] AND h.Nature='AM' AND h.[Production Date] BETWEEN f.VG_StartDate AND IIF(t.Booking IS NULL, f.VG_LastUpdateDate, Nz(t.[Production Date], Date())) AND Nz(h.Insuffisance,0)<0), True, False), " & _
      "       IIF(EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=f.Booking AND h.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT] AND h.Nature='SL' AND h.[Production Date] BETWEEN f.VG_StartDate AND IIF(t.Booking IS NULL, f.VG_LastUpdateDate, Nz(t.[Production Date], Date())) AND Nz(h.Insuffisance,0)<0), True, False) " & _
      "FROM " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.VG_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0);"
    nClosed = cn.RecordsAffected

    ' AM -> Archive
    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], " & _
      "AM_StartDate, AM_EndDate, AM_Duree, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance, AM_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], " & _
      "       f.AM_StartDate, IIF(t.Booking IS NULL, f.AM_LastUpdateDate, Nz(t.[Production Date], Date())) AS EndDate, " & _
      "       f.AM_NbJoursRetard, f.AM_WorstInsuffisance, f.AM_DateWorstInsuffisance, f.AM_LastInsuffisance, f.AM_LastUpdateDate, " & _
      "       f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], Nz(t.[Production Date], f.AM_LastUpdateDate), f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, " & _
      "       IIF(EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=f.Booking AND h.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT] AND h.Nature='VG' AND h.[Production Date] BETWEEN f.AM_StartDate AND IIF(t.Booking IS NULL, f.AM_LastUpdateDate, Nz(t.[Production Date], Date())) AND Nz(h.Insuffisance,0)<0), True, False), " & _
      "       True, " & _
      "       IIF(EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=f.Booking AND h.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT] AND h.Nature='SL' AND h.[Production Date] BETWEEN f.AM_StartDate AND IIF(t.Booking IS NULL, f.AM_LastUpdateDate, Nz(t.[Production Date], Date())) AND Nz(h.Insuffisance,0)<0), True, False) " & _
      "FROM " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.AM_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0);"
    nClosed = nClosed + cn.RecordsAffected

    ' SL -> Archive
    cn.Execute _
      "INSERT INTO " & T_ARC & " (Booking, [NO DOSSIER CREDIT], " & _
      "SL_StartDate, SL_EndDate, SL_Duree, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance, SL_LastUpdateDate, " & _
      "Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], DepassementNA, FlagVG, FlagAM, FlagSL) " & _
      "SELECT f.Booking, f.[NO DOSSIER CREDIT], " & _
      "       f.SL_StartDate, IIF(t.Booking IS NULL, f.SL_LastUpdateDate, Nz(t.[Production Date], Date())) AS EndDate, " & _
      "       f.SL_NbJoursRetard, f.SL_WorstInsuffisance, f.SL_DateWorstInsuffisance, f.SL_LastInsuffisance, f.SL_LastUpdateDate, " & _
      "       f.Limite, f.Consommation, f.[Montant VG], f.[Montant AM], f.[Montant SL], Nz(t.[Production Date], f.SL_LastUpdateDate), f.[NO INTERVENANT], f.[NO INTERVENANT GRP], f.DepassementNA, " & _
      "       IIF(EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=f.Booking AND h.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT] AND h.Nature='VG' AND h.[Production Date] BETWEEN f.SL_StartDate AND IIF(t.Booking IS NULL, f.SL_LastUpdateDate, Nz(t.[Production Date], Date())) AND Nz(h.Insuffisance,0)<0), True, False), " & _
      "       IIF(EXISTS (SELECT 1 FROM " & T_HIST & " h WHERE h.Booking=f.Booking AND h.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT] AND h.Nature='AM' AND h.[Production Date] BETWEEN f.SL_StartDate AND IIF(t.Booking IS NULL, f.SL_LastUpdateDate, Nz(t.[Production Date], Date())) AND Nz(h.Insuffisance,0)<0), True, False), " & _
      "       True " & _
      "FROM " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE f.SL_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0);"
    nClosed = nClosed + cn.RecordsAffected

    ' --- Nettoyage FLUX après archivage (vider les sous-épisodes clos) ---
    ' VG clear
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_StartDate=Null, f.VG_LastUpdateDate=Null, f.VG_NbJoursRetard=Null, f.VG_WorstInsuffisance=Null, f.VG_DateWorstInsuffisance=Null, f.VG_LastInsuffisance=Null, f.VG_EndDate=Null " & _
      "WHERE (f.VG_StartDate IS NOT NULL) AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0);"

    ' AM clear
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_StartDate=Null, f.AM_LastUpdateDate=Null, f.AM_NbJoursRetard=Null, f.AM_WorstInsuffisance=Null, f.AM_DateWorstInsuffisance=Null, f.AM_LastInsuffisance=Null, f.AM_EndDate=Null " & _
      "WHERE (f.AM_StartDate IS NOT NULL) AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0);"

    ' SL clear
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_StartDate=Null, f.SL_LastUpdateDate=Null, f.SL_NbJoursRetard=Null, f.SL_WorstInsuffisance=Null, f.SL_DateWorstInsuffisance=Null, f.SL_LastInsuffisance=Null, f.SL_EndDate=Null " & _
      "WHERE (f.SL_StartDate IS NOT NULL) AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0);"

    ' Supprimer ligne FLUX si plus aucun sous-épisode actif
    cn.Execute _
      "DELETE FROM " & T_FLUX & " WHERE VG_StartDate IS NULL AND AM_StartDate IS NULL AND SL_StartDate IS NULL;"

    cn.CommitTrans

    newFlux = nNew: contFlux = nCont: closedFlux = nClosed
    cn.Close: Set cn = Nothing
    Exit Sub
ROLLBACK:
    cn.RollbackTrans
    cn.Close: Set cn = Nothing
    Err.Raise Err.Number, , Err.Description
End Sub

' ====== Enrichissement feuille (3 × {NbJ, Début, Fin}) ======
Private Sub EnrichSheetFromFlux_3in1(ByVal dbPath As String, ByVal ws As Worksheet)
    ' Colonnes
    Dim cNBJ_VG&, cDEB_VG&, cFIN_VG&
    Dim cNBJ_AM&, cDEB_AM&, cFIN_AM&
    Dim cNBJ_SL&, cDEB_SL&, cFIN_SL&
    Dim cNA&
    cNBJ_VG = EnsureColumn(ws, COL_NBJ_VG): cDEB_VG = EnsureColumn(ws, COL_DEB_VG): cFIN_VG = EnsureColumn(ws, COL_FIN_VG)
    cNBJ_AM = EnsureColumn(ws, COL_NBJ_AM): cDEB_AM = EnsureColumn(ws, COL_DEB_AM): cFIN_AM = EnsureColumn(ws, COL_FIN_AM)
    cNBJ_SL = EnsureColumn(ws, COL_NBJ_SL): cDEB_SL = EnsureColumn(ws, COL_DEB_SL): cFIN_SL = EnsureColumn(ws, COL_FIN_SL)
    cNA = EnsureColumn(ws, COL_FLAG_NA)

    Dim prodDateSheet As Variant
    prodDateSheet = DetectSheetProductionDate(ws)

    ' Dicos Flux par Nature
    Dim dVG As Object, dAM As Object, dSL As Object
    Set dVG = LoadFluxDict_3in1(dbPath, "VG")
    Set dAM = LoadFluxDict_3in1(dbPath, "AM")
    Set dSL = LoadFluxDict_3in1(dbPath, "SL")

    ' Dicos Archives du jour (EndDate de la nature clôturée == prodDateSheet)
    Dim aVG As Object, aAM As Object, aSL As Object
    Set aVG = LoadArchiveEndDictForDate_3in1(dbPath, "VG", prodDateSheet)
    Set aAM = LoadArchiveEndDictForDate_3in1(dbPath, "AM", prodDateSheet)
    Set aSL = LoadArchiveEndDictForDate_3in1(dbPath, "SL", prodDateSheet)

    ' Mapping colonnes
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

        ' Dépassement NA
        Dim conso As Double, lim As Double
        conso = Val(ws.Cells(r, map(COL_CONSO)).Value)
        lim = Val(ws.Cells(r, map(COL_LIMITE)).Value)
        ws.Cells(r, cNA).Value = (Abs(conso) > lim)

        ' VG
        If dVG.Exists(k) Then
            ws.Cells(r, cNBJ_VG).Value = dVG(k)(1)
            ws.Cells(r, cDEB_VG).Value = dVG(k)(0)
        Else
            ws.Cells(r, cNBJ_VG).ClearContents
            ws.Cells(r, cDEB_VG).ClearContents
        End If
        If aVG.Exists(k) Then ws.Cells(r, cFIN_VG).Value = aVG(k) Else ws.Cells(r, cFIN_VG).ClearContents

        ' AM
        If dAM.Exists(k) Then
            ws.Cells(r, cNBJ_AM).Value = dAM(k)(1)
            ws.Cells(r, cDEB_AM).Value = dAM(k)(0)
        Else
            ws.Cells(r, cNBJ_AM).ClearContents
            ws.Cells(r, cDEB_AM).ClearContents
        End If
        If aAM.Exists(k) Then ws.Cells(r, cFIN_AM).Value = aAM(k) Else ws.Cells(r, cFIN_AM).ClearContents

        ' SL
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

Private Function DetectSheetProductionDate(ws As Worksheet) As Variant
    Dim col As Long: col = FindHeader(ws, COL_PROD)
    If col = 0 Then DetectSheetProductionDate = Date: Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Dim r As Long, maxDt As Date: maxDt = #1/1/1900#
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, col).Value) Then
            If CDate(ws.Cells(r, col).Value) > maxDt Then maxDt = CDate(ws.Cells(r, col).Value)
        End If
    Next r
    If maxDt = #1/1/1900# Then DetectSheetProductionDate = Date Else DetectSheetProductionDate = maxDt
End Function

Private Function LoadFluxDict_3in1(ByVal dbPath As String, ByVal nature As String) As Object
    Dim cn As Object: Set cn = GetConn(dbPath)
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    Dim sql As String
    If nature = "VG" Then
        sql = "SELECT Booking, [NO DOSSIER CREDIT], VG_StartDate, VG_NbJoursRetard FROM " & T_FLUX & " WHERE VG_StartDate IS NOT NULL"
    ElseIf nature = "AM" Then
        sql = "SELECT Booking, [NO DOSSIER CREDIT], AM_StartDate, AM_NbJoursRetard FROM " & T_FLUX & " WHERE AM_StartDate IS NOT NULL"
    Else
        sql = "SELECT Booking, [NO DOSSIER CREDIT], SL_StartDate, SL_NbJoursRetard FROM " & T_FLUX & " WHERE SL_StartDate IS NOT NULL"
    End If
    rs.Open sql, cn, 1, 1
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        d(Trim$(UCase$(rs(0).Value)) & "||" & Trim$(rs(1).Value)) = Array(rs(2).Value, rs(3).Value)
        rs.MoveNext
    Loop
    rs.Close: cn.Close
    Set LoadFluxDict_3in1 = d
End Function

Private Function LoadArchiveEndDictForDate_3in1(ByVal dbPath As String, ByVal nature As String, ByVal endDate As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    If IsDate(endDate) = False Then Set LoadArchiveEndDictForDate_3in1 = d: Exit Function
    Dim cn As Object: Set cn = GetConn(dbPath)
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    Dim sql As String
    If nature = "VG" Then
        sql = "SELECT Booking, [NO DOSSIER CREDIT], VG_EndDate FROM " & T_ARC & " WHERE FlagVG=True AND VG_EndDate = #" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf nature = "AM" Then
        sql = "SELECT Booking, [NO DOSSIER CREDIT], AM_EndDate FROM " & T_ARC & " WHERE FlagAM=True AND AM_EndDate = #" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    Else
        sql = "SELECT Booking, [NO DOSSIER CREDIT], SL_EndDate FROM " & T_ARC & " WHERE FlagSL=True AND SL_EndDate = #" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    End If
    rs.Open sql, cn, 1, 1
    Do While Not rs.EOF
        d(Trim$(UCase$(rs(0).Value)) & "||" & Trim$(rs(1).Value)) = rs(2).Value
        rs.MoveNext
    Loop
    rs.Close: cn.Close
    Set LoadArchiveEndDictForDate_3in1 = d
End Function

' ====== Helpers communs ======
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

Private Function NzS(v) As Variant
    If IsError(v) Then NzS = Null: Exit Function
    If Len(v & "") = 0 Then NzS = Null Else NzS = v
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

## 4) Comment ça marche (pas-à-pas)

1) **Tu colles la feuille du jour** → clic bouton **`MAJ_Retards_Daily_3in1`**.  
2) Le module :  
   - crée la **base** & **tables** si besoin ;  
   - charge **T_Today** (gère les en-têtes y/c coquilles), **normalise les clés** (`Booking=UCase(Trim)`, `Dossier=Trim`), **dédoublonne** par clé (garde la plus récente) ;  
   - pousse la **time series** `T_HistoryInsuff` (VG/AM/SL) ;  
   - met à jour **T_Flux3** :  
     - crée une ligne **si la clé n’existe pas** et qu’au moins une insuff < 0 (initialise les champs de la nature concernée) ;  
     - démarre un **sous-épisode** si la clé existe mais la nature était **inactive** et aujourd’hui `< 0` ;  
     - **continue** un sous-épisode actif (incrément par **DateDiff('d', LastUpdateDate, ProductionDate)**, mise à jour Worst/Last, contexte, flag NA) ;  
     - **archive** toute nature qui **n’est pas < 0** aujourd’hui (ou clé absente) : insère dans **T_Archive3** la partie nature (avec **FlagVG/AM/SL** calculés via l’historique), puis **vide** cette nature dans le flux ;  
     - **supprime** la ligne Flux si **les 3 natures** sont devenues inactives.  
   - enrichit la **feuille Excel** avec :  
     - `NB Jours retard VG/AM/SL` + `Date début VG/AM/SL` (depuis `T_Flux3`) ;  
     - `Date fin VG/AM/SL` du **jour** (depuis `T_Archive3`) ;  
     - le **flag** `Dépassement non autorisé` = `Abs(Consommation)>Limite`.  
   - journalise dans `T_RunLog`.

---

## 5) Requêtes de contrôle / reporting

- **Flux courant (3 sous-épisodes sur une ligne)**
```sql
SELECT Booking, [NO DOSSIER CREDIT],
       VG_StartDate, VG_NbJoursRetard, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance,
       AM_StartDate, AM_NbJoursRetard, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance,
       SL_StartDate, SL_NbJoursRetard, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance,
       Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], DepassementNA
FROM T_Flux3
ORDER BY Booking, [NO DOSSIER CREDIT];
```

- **Clôtures (derniers jours)**
```sql
SELECT Booking, [NO DOSSIER CREDIT], ArchivedAt,
       VG_EndDate, AM_EndDate, SL_EndDate, FlagVG, FlagAM, FlagSL
FROM T_Archive3
WHERE Nz(VG_EndDate, #1900-01-01#) >= DateAdd('d', -7, Date())
   OR Nz(AM_EndDate, #1900-01-01#) >= DateAdd('d', -7, Date())
   OR Nz(SL_EndDate, #1900-01-01#) >= DateAdd('d', -7, Date())
ORDER BY ArchivedAt DESC;
```

- **Trajectoire quotidienne (ex VG)**
```sql
SELECT [Production Date], Insuffisance, Limite, Consommation, [Montant], DepassementNA
FROM T_HistoryInsuff
WHERE Booking='IND' AND [NO DOSSIER CREDIT]='12345' AND Nature='VG'
ORDER BY [Production Date];
```

---

## 6) Remarques importantes

- **Tu as bien 3 compteurs** (`VG_NbJoursRetard`, `AM_NbJoursRetard`, `SL_NbJoursRetard`) et **3 couples `StartDate/EndDate`** (End en archive — ou NULL en flux tant que non clos).  
- **WorstInsuffisance** = le **plus négatif** observé pendant l’épisode ; **DateWorstInsuffisance** = la date de ce plus bas.  
- **Booking + NO DOSSIER CREDIT** = **clé unique** ; les jointures/updates se font **toujours** sur ces 2 champs.  
- Les **flags** en archive (`FlagVG/AM/SL`) qualifient si **un / deux / trois** types étaient en retard pendant l’épisode clôturé.  
- Les **index** fournis accélèrent **fortement** les opérations ; ne les retire pas.

---

## 7) À faire une fois

1. Placer un bouton sur ta feuille et lui affecter `MAJ_Retards_Daily_3in1`.  
2. Laisser le classeur dans un dossier **avec droits d’écriture** (la base `.accdb` est créée **à côté**).  
3. Vérifier les **en-têtes** (coquilles gérées : `NO NTERVENANT`, `NO NTERVENANT GRP`).  
4. Lancer **quotidiennement** après avoir collé/rafraîchi la feuille.

Bon run 🚀
