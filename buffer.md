# Archivage consolidé via `T_CloseBuf` — Blocs à ajouter et à modifier

Parfait — voici tous les blocs à **AJOUTER** et ceux à **MODIFIER** pour passer à l’archivage consolidé via le buffer `T_CloseBuf`, avec les optimisations (calcul des flags et de `DepassementNA` avant l’INSERT dans l’archive).

---

## 🔧 AJOUTER

### 1) Constante (en haut, avec les autres `Private Const`)
```vb
Private Const T_CLOSE As String = "T_CloseBuf"
```

### 2) DDL de la table buffer
```vb
Private Function GetDDL_Close() As String
    GetDDL_Close = _
      "CREATE TABLE " & T_CLOSE & " (" & _
      "  Booking TEXT(20) NOT NULL, [NO DOSSIER CREDIT] TEXT(50) NOT NULL, " & _
      "  VG_StartDate DATE, VG_EndDate DATE, VG_Duree LONG, VG_WorstInsuffisance DOUBLE, VG_DateWorstInsuffisance DATE, VG_LastInsuffisance DOUBLE, VG_LastUpdateDate DATE, " & _
      "  AM_StartDate DATE, AM_EndDate DATE, AM_Duree LONG, AM_WorstInsuffisance DOUBLE, AM_DateWorstInsuffisance DATE, AM_LastInsuffisance DOUBLE, AM_LastUpdateDate DATE, " & _
      "  SL_StartDate DATE, SL_EndDate DATE, SL_Duree LONG, SL_WorstInsuffisance DOUBLE, SL_DateWorstInsuffisance DATE, SL_LastInsuffisance DOUBLE, SL_LastUpdateDate DATE, " & _
      "  Limite DOUBLE, Consommation DOUBLE, [Montant VG] DOUBLE, [Montant AM] DOUBLE, [Montant SL] DOUBLE, " & _
      "  [Production Date] DATE, [NO INTERVENANT] TEXT(50), [NO INTERVENANT GRP] TEXT(50), " & _
      "  DepassementNA TEXT(3) DEFAULT 'NO', " & _
      "  FlagVG TEXT(3) DEFAULT 'NO', FlagAM TEXT(3) DEFAULT 'NO', FlagSL TEXT(3) DEFAULT 'NO', " & _
      "  CONSTRAINT PK_CloseBuf PRIMARY KEY (Booking, [NO DOSSIER CREDIT])" & _
      ");"
End Function
```

### 3) Builder du buffer (consolide VG/AM/SL + calcule flags + `DepassementNA`)
```vb
Private Sub BuildCloseBuffer(ByVal cn As Object)
    cn.Execute "DELETE FROM " & T_CLOSE

    ' Clés à clôturer si AU MOINS une nature se ferme
    cn.Execute _
      "INSERT INTO " & T_CLOSE & " (Booking, [NO DOSSIER CREDIT]) " & _
      "SELECT DISTINCT f.Booking, f.[NO DOSSIER CREDIT] " & _
      "FROM " & T_FLUX & " AS f " & _
      "LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE (f.VG_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0)) " & _
      "   OR (f.AM_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0)) " & _
      "   OR (f.SL_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0));"

    ' --- VG ---
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.VG_StartDate=f.VG_StartDate, b.VG_Duree=f.VG_NbJoursRetard, " & _
      "    b.VG_WorstInsuffisance=f.VG_WorstInsuffisance, b.VG_DateWorstInsuffisance=f.VG_DateWorstInsuffisance, " & _
      "    b.VG_LastInsuffisance=f.VG_LastInsuffisance, b.VG_LastUpdateDate=f.VG_LastUpdateDate, b.FlagVG='YES', " & _
      "    b.VG_EndDate=IIf(t.Booking Is Null, f.VG_LastUpdateDate, t.[Production Date]) " & _
      "WHERE f.VG_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0);"

    ' --- AM ---
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.AM_StartDate=f.AM_StartDate, b.AM_Duree=f.AM_NbJoursRetard, " & _
      "    b.AM_WorstInsuffisance=f.AM_WorstInsuffisance, b.AM_DateWorstInsuffisance=f.AM_DateWorstInsuffisance, " & _
      "    b.AM_LastInsuffisance=f.AM_LastInsuffisance, b.AM_LastUpdateDate=f.AM_LastUpdateDate, b.FlagAM='YES', " & _
      "    b.AM_EndDate=IIf(t.Booking Is Null, f.AM_LastUpdateDate, t.[Production Date]) " & _
      "WHERE f.AM_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0);"

    ' --- SL ---
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.SL_StartDate=f.SL_StartDate, b.SL_Duree=f.SL_NbJoursRetard, " & _
      "    b.SL_WorstInsuffisance=f.SL_WorstInsuffisance, b.SL_DateWorstInsuffisance=f.SL_DateWorstInsuffisance, " & _
      "    b.SL_LastInsuffisance=f.SL_LastInsuffisance, b.SL_LastUpdateDate=f.SL_LastUpdateDate, b.FlagSL='YES', " & _
      "    b.SL_EndDate=IIf(t.Booking Is Null, f.SL_LastUpdateDate, t.[Production Date]) " & _
      "WHERE f.SL_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0);"

    ' Champs communs (Today si présent, sinon Flux)
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.Limite=Nz(t.Limite,f.Limite), b.Consommation=Nz(t.Consommation,f.Consommation), " & _
      "    b.[Montant VG]=Nz(t.[Montant VG],f.[Montant VG]), b.[Montant AM]=Nz(t.[Montant AM],f.[Montant AM]), b.[Montant SL]=Nz(t.[Montant SL],f.[Montant SL]), " & _
      "    b.[Production Date]=Nz(t.[Production Date],f.[Production Date]), b.[NO INTERVENANT]=Nz(t.[NO INTERVENANT],f.[NO INTERVENANT]), " & _
      "    b.[NO INTERVENANT GRP]=Nz(t.[NO INTERVENANT GRP],f.[NO INTERVENANT GRP]);"

    ' DepassementNA en une passe (Today dispo)
    cn.Execute _
      "UPDATE " & T_CLOSE & " b INNER JOIN " & T_TODAY & " t " & _
      "ON b.Booking=t.Booking AND b.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.DepassementNA = IIf(Abs(t.Consommation) > t.Limite, 'YES','NO') " & _
      "WHERE Nz(b.DepassementNA,'') <> IIf(Abs(t.Consommation) > t.Limite, 'YES','NO');"

    ' DepassementNA hérité de Flux quand Today absent
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking ET b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.DepassementNA = IIf(Abs(f.Consommation) > f.Limite, 'YES','NO') " & _
      "WHERE t.Booking IS NULL AND Nz(b.DepassementNA,'') <> IIf(Abs(f.Consommation) > f.Limite, 'YES','NO');"

    ' ===== Chevauchements via T_HIST (calcul des Flags dans le buffer) =====
    ' VG -> AM/SL
    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagAM='YES' " & _
      "WHERE b.FlagVG='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='AM' " & _
      "    AND h.[Production Date] BETWEEN b.VG_StartDate AND b.VG_EndDate AND h.Insuffisance < 0);"

    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagSL='YES' " & _
      "WHERE b.FlagVG='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='SL' " & _
      "    AND h.[Production Date] BETWEEN b.VG_StartDate AND b.VG_EndDate AND h.Insuffisance < 0);"

    ' AM -> VG/SL
    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagVG='YES' " & _
      "WHERE b.FlagAM='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='VG' " & _
      "    AND h.[Production Date] BETWEEN b.AM_StartDate AND b.AM_EndDate AND h.Insuffisance < 0);"

    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagSL='YES' " & _
      "WHERE b.FlagAM='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='SL' " & _
      "    AND h.[Production Date] BETWEEN b.AM_StartDate AND b.AM_EndDate AND h.Insuffisance < 0);"

    ' SL -> VG/AM
    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagVG='YES' " & _
      "WHERE b.FlagSL='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='VG' " & _
      "    AND h.[Production Date] BETWEEN b.SL_StartDate AND b.SL_EndDate AND h.Insuffisance < 0);"

    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagAM='YES' " & _
      "WHERE b.FlagSL='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='AM' " & _
      "    AND h.[Production Date] BETWEEN b.SL_StartDate AND b.SL_EndDate AND h.Insuffisance < 0);"
End Sub
```

### 4) Insertion du buffer vers l’archive (une seule passe)
```vb
Private Function InsertCloseBufferToArchive(ByVal cn As Object) As Long
    cn.Execute _
      "INSERT INTO " & T_ARC & " (" & _
      "  Booking, [NO DOSSIER CREDIT], " & _
      "  VG_StartDate, VG_EndDate, VG_Duree, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance, VG_LastUpdateDate, " & _
      "  AM_StartDate, AM_EndDate, AM_Duree, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance, AM_LastUpdateDate, " & _
      "  SL_StartDate, SL_EndDate, SL_Duree, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance, SL_LastUpdateDate, " & _
      "  Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], " & _
      "  DepassementNA, FlagVG, FlagAM, FlagSL" & _
      ") " & _
      "SELECT " & _
      "  b.Booking, b.[NO DOSSIER CREDIT], " & _
      "  b.VG_StartDate, b.VG_EndDate, b.VG_Duree, b.VG_WorstInsuffisance, b.VG_DateWorstInsuffisance, b.VG_LastInsuffisance, b.VG_LastUpdateDate, " & _
      "  b.AM_StartDate, b.AM_EndDate, b.AM_Duree, b.AM_WorstInsuffisance, b.AM_DateWorstInsuffisance, b.AM_LastInsuffisance, b.AM_LastUpdateDate, " & _
      "  b.SL_StartDate, b.SL_EndDate, b.SL_Duree, b.SL_WorstInsuffisance, b.SL_DateWorstInsuffisance, b.SL_LastInsuffisance, b.SL_LastUpdateDate, " & _
      "  b.Limite, b.Consommation, b.[Montant VG], b.[Montant AM], b.[Montant SL], b.[Production Date], b.[NO INTERVENANT], b.[NO INTERVENANT GRP], " & _
      "  b.DepassementNA, b.FlagVG, b.FlagAM, b.FlagSL " & _
      "FROM " & T_CLOSE & " b;"
    InsertCloseBufferToArchive = cn.RecordsAffected
End Function
```

---

## ✏️ MODIFIER

### 5) `EnsureDatabaseAndSchema` (ajoute la ligne pour créer `T_CloseBuf`)
```vb
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
    ' >>> AJOUT <<<
    ExecIfNotExists cn, T_CLOSE, GetDDL_Close
    ' ------------
    cn.Close: Set cn = Nothing
End Sub
```

### 6) `ProcessBusinessRules_3in1_NoNzIif` (changements principaux résumés)
- `DepassementNA` sur `T_FLUX` en **une seule passe** (`IIf`) avant l’archivage.
- **Bloc (6)** remplacé par **`BuildCloseBuffer` + `InsertCloseBufferToArchive`** (archivage consolidé).
- **Bloc (7)** supprimé (les `Flag*` sont calculés dans le buffer).
- **Bloc (8)** (purge `T_FLUX`) **inchangé**.

```vb
Private Sub ProcessBusinessRules_3in1_NoNzIif(ByVal dbPath As String, ByRef newFlux As Long, ByRef contFlux As Long, ByRef closedFlux As Long)
    Dim cn As Object: Set cn = GetConn(dbPath)
    On Error GoTo ROLLBACK
    cn.BeginTrans

    Dim nNew As Long, nCont As Long, nClosed As Long

    ' 1) INSERT clés en FLUX si au moins une insuff < 0
    cn.Execute _
      "INSERT INTO " & T_FLUX & " (" & _
      "  Booking, [NO DOSSIER CREDIT], Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP]" & _
      ") " & _
      "SELECT " & _
      "  CStr(t.Booking), CStr(t.[NO DOSSIER CREDIT]), " & _
      "  CDbl(t.Limite), CDbl(t.Consommation), CDbl(t.[Montant VG]), CDbl(t.[Montant AM]), CDbl(t.[Montant SL]), " & _
      "  CDate(t.[Production Date]), CStr(t.[NO INTERVENANT]), CStr(t.[NO INTERVENANT GRP]) " & _
      "FROM " & T_TODAY & " AS t " & _
      "WHERE (t.[Insuffisance VG] < 0 OR t.[Insuffisance AM] < 0 OR t.[Insuffisance SL] < 0) " & _
      "  AND NOT EXISTS (SELECT 1 FROM " & T_FLUX & " AS f WHERE f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT]);"
    nNew = RecordsAffected(cn)

    ' 2) DepassementNA (une seule passe, only dirty rows)
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.DepassementNA = IIf(Abs(t.Consommation) > t.Limite,'YES','NO') " & _
      "WHERE Nz(f.DepassementNA,'') <> IIf(Abs(t.Consommation) > t.Limite,'YES','NO');"

    ' 3) Démarrages VG/AM/SL (inchangé)
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_StartDate=t.[Production Date], f.VG_LastUpdateDate=t.[Production Date], f.VG_NbJoursRetard=1, f.VG_WorstInsuffisance=t.[Insuffisance VG], f.VG_DateWorstInsuffisance=t.[Production Date], f.VG_LastInsuffisance=t.[Insuffisance VG] " & _
      "WHERE f.VG_StartDate IS NULL AND t.[Insuffisance VG] < 0;"
    nNew = nNew + RecordsAffected(cn)

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_StartDate=t.[Production Date], f.AM_LastUpdateDate=t.[Production Date], f.AM_NbJoursRetard=1, f.AM_WorstInsuffisance=t.[Insuffisance AM], f.AM_DateWorstInsuffisance=t.[Production Date], f.AM_LastInsuffisance=t.[Insuffisance AM] " & _
      "WHERE f.AM_StartDate IS NULL AND t.[Insuffisance AM] < 0;"
    nNew = nNew + RecordsAffected(cn)

    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_StartDate=t.[Production Date], f.SL_LastUpdateDate=t.[Production Date], f.SL_NbJoursRetard=1, f.SL_WorstInsuffisance=t.[Insuffisance SL], f.SL_DateWorstInsuffisance=t.[Production Date], f.SL_LastInsuffisance=t.[Insuffisance SL] " & _
      "WHERE f.SL_StartDate IS NULL AND t.[Insuffisance SL] < 0;"
    nNew = nNew + RecordsAffected(cn)

    ' 4) Continuations + "pire" (inchangé)
    ' VG
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_NbJoursRetard=f.VG_NbJoursRetard + DateDiff('d', f.VG_LastUpdateDate, t.[Production Date]), " & _
      "    f.VG_LastUpdateDate=t.[Production Date], f.VG_LastInsuffisance=t.[Insuffisance VG], " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP] " & _
      "WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NOT NULL AND t.[Production Date] > f.VG_LastUpdateDate;"
    nCont = RecordsAffected(cn)
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.VG_WorstInsuffisance=t.[Insuffisance VG], f.VG_DateWorstInsuffisance=t.[Production Date] " & _
      "WHERE t.[Insuffisance VG] < 0 AND f.VG_StartDate IS NOT NULL AND (f.VG_WorstInsuffisance IS NULL OR t.[Insuffisance VG] < f.VG_WorstInsuffisance);"

    ' AM
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_NbJoursRetard=f.AM_NbJoursRetard + DateDiff('d', f.AM_LastUpdateDate, t.[Production Date]), " & _
      "    f.AM_LastUpdateDate=t.[Production Date], f.AM_LastInsuffisance=t.[Insuffisance AM], " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP] " & _
      "WHERE t.[Insuffisance AM] < 0 AND f.AM_StartDate IS NOT NULL AND t.[Production Date] > f.AM_LastUpdateDate;"
    nCont = nCont + RecordsAffected(cn)
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.AM_WorstInsuffisance=t.[Insuffisance AM], f.AM_DateWorstInsuffisance=t.[Production Date] " & _
      "WHERE t.[Insuffisance AM] < 0 AND f.AM_StartDate IS NOT NULL AND (f.AM_WorstInsuffisance IS NULL OR t.[Insuffisance AM] < f.AM_WorstInsuffisance);"

    ' SL
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_NbJoursRetard=f.SL_NbJoursRetard + DateDiff('d', f.SL_LastUpdateDate, t.[Production Date]), " & _
      "    f.SL_LastUpdateDate=t.[Production Date], f.SL_LastInsuffisance=t.[Insuffisance SL], " & _
      "    f.Limite=t.Limite, f.Consommation=t.Consommation, f.[Montant VG]=t.[Montant VG], f.[Montant AM]=t.[Montant AM], f.[Montant SL]=t.[Montant SL], f.[Production Date]=t.[Production Date], " & _
      "    f.[NO INTERVENANT]=t.[NO INTERVENANT], f.[NO INTERVENANT GRP]=t.[NO INTERVENANT GRP] " & _
      "WHERE t.[Insuffisance SL] < 0 AND f.SL_StartDate IS NOT NULL AND t.[Production Date] > f.SL_LastUpdateDate;"
    nCont = nCont + RecordsAffected(cn)
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.SL_WorstInsuffisance=t.[Insuffisance SL], f.SL_DateWorstInsuffisance=t.[Production Date] " & _
      "WHERE t.[Insuffisance SL] < 0 AND f.SL_StartDate IS NOT NULL AND (f.SL_WorstInsuffisance IS NULL OR t.[Insuffisance SL] < f.SL_WorstInsuffisance);"

    ' 5) Rafraîchir DepassementNA après continuations (une passe)
    cn.Execute _
      "UPDATE " & T_FLUX & " AS f INNER JOIN " & T_TODAY & " AS t " & _
      "ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET f.DepassementNA = IIf(Abs(t.Consommation) > t.Limite,'YES','NO') " & _
      "WHERE Nz(f.DepassementNA,'') <> IIf(Abs(t.Consommation) > t.Limite,'YES','NO');"

    ' 6) Archivage consolidé (buffer -> archive)
    BuildCloseBuffer cn
    nClosed = InsertCloseBufferToArchive(cn)

    ' 7) (supprimé — flags déjà calculés dans le buffer)

    ' 8) Purge FLUX des sous-épisodes clos (inchangé)
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

    ' Supprime la ligne si plus rien d'ouvert
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

*(Le code complet a été donné dans les échanges précédents. Seule la section d’archivage est remplacée par l’utilisation du buffer.)*

---

## 🔒 (Optionnel) Garde-fou anti-doublon “une ligne par jour”
Ajoute un `NOT EXISTS` basé sur `(Booking, [NO DOSSIER CREDIT], [Production Date])` :
```vb
Private Function InsertCloseBufferToArchive(ByVal cn As Object) As Long
    cn.Execute _
      "INSERT INTO " & T_ARC & " (" & _
      "  Booking, [NO DOSSIER CREDIT], " & _
      "  VG_StartDate, VG_EndDate, VG_Duree, VG_WorstInsuffisance, VG_DateWorstInsuffisance, VG_LastInsuffisance, VG_LastUpdateDate, " & _
      "  AM_StartDate, AM_EndDate, AM_Duree, AM_WorstInsuffisance, AM_DateWorstInsuffisance, AM_LastInsuffisance, AM_LastUpdateDate, " & _
      "  SL_StartDate, SL_EndDate, SL_Duree, SL_WorstInsuffisance, SL_DateWorstInsuffisance, SL_LastInsuffisance, SL_LastUpdateDate, " & _
      "  Limite, Consommation, [Montant VG], [Montant AM], [Montant SL], [Production Date], [NO INTERVENANT], [NO INTERVENANT GRP], " & _
      "  DepassementNA, FlagVG, FlagAM, FlagSL" & _
      ") " & _
      "SELECT " & _
      "  b.Booking, b.[NO DOSSIER CREDIT], " & _
      "  b.VG_StartDate, b.VG_EndDate, b.VG_Duree, b.VG_WorstInsuffisance, b.VG_DateWorstInsuffisance, b.VG_LastInsuffisance, b.VG_LastUpdateDate, " & _
      "  b.AM_StartDate, b.AM_EndDate, b.AM_Duree, b.AM_WorstInsuffisance, b.AM_DateWorstInsuffisance, b.AM_LastInsuffisance, b.AM_LastUpdateDate, " & _
      "  b.SL_StartDate, b.SL_EndDate, b.SL_Duree, b.SL_WorstInsuffisance, b.SL_DateWorstInsuffisance, b.SL_LastInsuffisance, b.SL_LastUpdateDate, " & _
      "  b.Limite, b.Consommation, b.[Montant VG], b.[Montant AM], b.[Montant SL], b.[Production Date], b.[NO INTERVENANT], b.[NO INTERVENANT GRP], " & _
      "  b.DepassementNA, b.FlagVG, b.FlagAM, b.FlagSL " & _
      "FROM " & T_CLOSE & " b " & _
      "WHERE NOT EXISTS ( " & _
      "  SELECT 1 FROM " & T_ARC & " a " & _
      "  WHERE a.Booking=b.Booking " & _
      "    AND a.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] " & _
      "    AND a.[Production Date]=b.[Production Date] );"
    InsertCloseBufferToArchive = cn.RecordsAffected
End Function
```

---

## ✅ Résultat
- **Une seule ligne par (Booking, NO DOSSIER CREDIT, jour)** dans `T_Archive` pour les natures qui se ferment le même jour.
- Les colonnes `VG_*`, `AM_*`, `SL_*` et `Flag*` sont **remplies ensemble**.
- Pas de doublon même si le traitement est relancé (avec le garde-fou).


Private Sub BuildCloseBuffer(ByVal cn As Object)
    cn.Execute "DELETE FROM " & T_CLOSE

    ' Clés à clôturer si AU MOINS une nature se ferme
    cn.Execute _
      "INSERT INTO " & T_CLOSE & " (Booking, [NO DOSSIER CREDIT]) " & _
      "SELECT DISTINCT f.Booking, f.[NO DOSSIER CREDIT] " & _
      "FROM " & T_FLUX & " AS f " & _
      "LEFT JOIN " & T_TODAY & " AS t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "WHERE (f.VG_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0)) " & _
      "   OR (f.AM_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0)) " & _
      "   OR (f.SL_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0));"

    ' --- VG ---
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.VG_StartDate=f.VG_StartDate, b.VG_Duree=f.VG_NbJoursRetard, " & _
      "    b.VG_WorstInsuffisance=f.VG_WorstInsuffisance, b.VG_DateWorstInsuffisance=f.VG_DateWorstInsuffisance, " & _
      "    b.VG_LastInsuffisance=f.VG_LastInsuffisance, b.VG_LastUpdateDate=f.VG_LastUpdateDate, b.FlagVG='YES', " & _
      "    b.VG_EndDate=IIf(t.Booking Is Null, f.VG_LastUpdateDate, t.[Production Date]) " & _
      "WHERE f.VG_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance VG] >= 0);"

    ' --- AM ---
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.AM_StartDate=f.AM_StartDate, b.AM_Duree=f.AM_NbJoursRetard, " & _
      "    b.AM_WorstInsuffisance=f.AM_WorstInsuffisance, b.AM_DateWorstInsuffisance=f.AM_DateWorstInsuffisance, " & _
      "    b.AM_LastInsuffisance=f.AM_LastInsuffisance, b.AM_LastUpdateDate=f.AM_LastUpdateDate, b.FlagAM='YES', " & _
      "    b.AM_EndDate=IIf(t.Booking Is Null, f.AM_LastUpdateDate, t.[Production Date]) " & _
      "WHERE f.AM_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance AM] >= 0);"

    ' --- SL ---
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.SL_StartDate=f.SL_StartDate, b.SL_Duree=f.SL_NbJoursRetard, " & _
      "    b.SL_WorstInsuffisance=f.SL_WorstInsuffisance, b.SL_DateWorstInsuffisance=f.SL_DateWorstInsuffisance, " & _
      "    b.SL_LastInsuffisance=f.SL_LastInsuffisance, b.SL_LastUpdateDate=f.SL_LastUpdateDate, b.FlagSL='YES', " & _
      "    b.SL_EndDate=IIf(t.Booking Is Null, f.SL_LastUpdateDate, t.[Production Date]) " & _
      "WHERE f.SL_StartDate IS NOT NULL AND (t.Booking IS NULL OR t.[Insuffisance SL] >= 0);"

    ' Champs communs (Today si présent, sinon Flux) — sans Nz
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.Limite = IIf(t.Limite Is Null, f.Limite, t.Limite), " & _
      "    b.Consommation = IIf(t.Consommation Is Null, f.Consommation, t.Consommation), " & _
      "    b.[Montant VG] = IIf(t.[Montant VG] Is Null, f.[Montant VG], t.[Montant VG]), " & _
      "    b.[Montant AM] = IIf(t.[Montant AM] Is Null, f.[Montant AM], t.[Montant AM]), " & _
      "    b.[Montant SL] = IIf(t.[Montant SL] Is Null, f.[Montant SL], t.[Montant SL]), " & _
      "    b.[Production Date] = IIf(t.[Production Date] Is Null, f.[Production Date], t.[Production Date]), " & _
      "    b.[NO INTERVENANT] = IIf(t.[NO INTERVENANT] Is Null, f.[NO INTERVENANT], t.[NO INTERVENANT]), " & _
      "    b.[NO INTERVENANT GRP] = IIf(t.[NO INTERVENANT GRP] Is Null, f.[NO INTERVENANT GRP], t.[NO INTERVENANT GRP]);"

    ' DepassementNA (Today dispo) — on met à jour sans Nz
    cn.Execute _
      "UPDATE " & T_CLOSE & " b INNER JOIN " & T_TODAY & " t " & _
      "ON b.Booking=t.Booking AND b.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.DepassementNA = IIf(Abs(t.Consommation) > t.Limite, 'YES','NO');"

    ' DepassementNA hérité de Flux quand Today absent (corrigé AND)
    cn.Execute _
      "UPDATE (" & T_CLOSE & " b INNER JOIN " & T_FLUX & " f ON b.Booking=f.Booking AND b.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT]) " & _
      "LEFT JOIN " & T_TODAY & " t ON f.Booking=t.Booking AND f.[NO DOSSIER CREDIT]=t.[NO DOSSIER CREDIT] " & _
      "SET b.DepassementNA = IIf(Abs(f.Consommation) > f.Limite, 'YES','NO') " & _
      "WHERE t.Booking IS NULL;"

    ' ===== Chevauchements via T_HIST (calcul des Flags dans le buffer) =====
    ' VG -> AM/SL
    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagAM='YES' " & _
      "WHERE b.FlagVG='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='AM' " & _
      "    AND h.[Production Date] BETWEEN b.VG_StartDate AND b.VG_EndDate AND h.Insuffisance < 0);"

    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagSL='YES' " & _
      "WHERE b.FlagVG='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='SL' " & _
      "    AND h.[Production Date] BETWEEN b.VG_StartDate AND b.VG_EndDate AND h.Insuffisance < 0);"

    ' AM -> VG/SL
    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagVG='YES' " & _
      "WHERE b.FlagAM='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='VG' " & _
      "    AND h.[Production Date] BETWEEN b.AM_StartDate AND b.AM_EndDate AND h.Insuffisance < 0);"

    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagSL='YES' " & _
      "WHERE b.FlagAM='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='SL' " & _
      "    AND h.[Production Date] BETWEEN b.AM_StartDate AND b.AM_EndDate AND h.Insuffisance < 0);"

    ' SL -> VG/AM
    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagVG='YES' " & _
      "WHERE b.FlagSL='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='VG' " & _
      "    AND h.[Production Date] BETWEEN b.SL_StartDate AND b.SL_EndDate AND h.Insuffisance < 0);"

    cn.Execute _
      "UPDATE " & T_CLOSE & " AS b SET b.FlagAM='YES' " & _
      "WHERE b.FlagSL='YES' AND EXISTS (" & _
      "  SELECT 1 FROM " & T_HIST & " h " & _
      "  WHERE h.Booking=b.Booking AND h.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT] AND h.Nature='AM' " & _
      "    AND h.[Production Date] BETWEEN b.SL_StartDate AND b.SL_EndDate AND h.Insuffisance < 0);"
End Sub
