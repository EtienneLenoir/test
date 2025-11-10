[TOC]

# Pack complet de requêtes Access — Reporting & Statistiques (Markdown)

> **Mode d’emploi :** copiez chaque bloc en **Affichage SQL** dans Access et enregistrez la requête.
> Variables de paramètres utilisées dans plusieurs requêtes : **[StartDate]**, **[EndDate]**, **[AsOfDate]**, **[WindowDays]**, **[SLA_Days]**, **[J]**.

---

# 0) Pré‑requis (vues de base)

## 0.1) `Q_Arc_UnpivotNature` — *Normalisation par nature (UNION VG/AM/SL)*  
Produit une vue “longue” : une ligne par épisode et par nature, à partir de `T_Archive`.
```sql
SELECT
    ID AS ArcID, Booking, [NO DOSSIER CREDIT],
    'VG' AS Nature,
    VG_StartDate AS StartDate, VG_EndDate AS EndDate, VG_Duree AS Duree,
    VG_WorstInsuffisance AS WorstInsuffisance, VG_LastInsuffisance AS LastInsuffisance,
    VG_LastUpdateDate AS LastUpdateDate,
    [Montant VG] AS Montant, FlagVG AS Flag,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive
WHERE FlagVG='YES'
UNION ALL
SELECT
    ID AS ArcID, Booking, [NO DOSSIER CREDIT],
    'AM' AS Nature,
    AM_StartDate, AM_EndDate, AM_Duree,
    AM_WorstInsuffisance, AM_LastInsuffisance,
    AM_LastUpdateDate,
    [Montant AM], FlagAM,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive
WHERE FlagAM='YES'
UNION ALL
SELECT
    ID AS ArcID, Booking, [NO DOSSIER CREDIT],
    'SL' AS Nature,
    SL_StartDate, SL_EndDate, SL_Duree,
    SL_WorstInsuffisance, SL_LastInsuffisance,
    SL_LastUpdateDate,
    [Montant SL], FlagSL,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive
WHERE FlagSL='YES';
```

## 0.2) `Q_Arc_UnpivotNature_W` — *Ajoute la date du pire (`WorstDate`)*  
Même principe, avec la date de l’insuffisance la plus négative par nature.
```sql
SELECT
    ID AS ArcID, Booking, [NO DOSSIER CREDIT], 'VG' AS Nature,
    VG_StartDate AS StartDate, VG_EndDate AS EndDate, VG_Duree AS Duree,
    VG_WorstInsuffisance AS WorstInsuffisance, VG_DateWorstInsuffisance AS WorstDate,
    VG_LastInsuffisance AS LastInsuffisance, VG_LastUpdateDate AS LastUpdateDate,
    [Montant VG] AS Montant, FlagVG AS Flag,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive WHERE FlagVG='YES'
UNION ALL
SELECT ID, Booking, [NO DOSSIER CREDIT], 'AM',
       AM_StartDate, AM_EndDate, AM_Duree,
       AM_WorstInsuffisance, AM_DateWorstInsuffisance,
       AM_LastInsuffisance, AM_LastUpdateDate,
       [Montant AM], FlagAM,
       [Production Date], Limite, Consommation, DepassementNA,
       [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive WHERE FlagAM='YES'
UNION ALL
SELECT ID, Booking, [NO DOSSIER CREDIT], 'SL',
       SL_StartDate, SL_EndDate, SL_Duree,
       SL_WorstInsuffisance, SL_DateWorstInsuffisance,
       SL_LastInsuffisance, SL_LastUpdateDate,
       [Montant SL], FlagSL,
       [Production Date], Limite, Consommation, DepassementNA,
       [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive WHERE FlagSL='YES';
```


## 0.3) Q_KeyNature_Interv — *Mapping clé+nature → intervenant (fallback “dernier connu”)*
> Sert pour les séries historiques sur `T_HistoryInsuff` (qui ne contient pas l’intervenant).
```sql
SELECT Nature, Booking, [NO DOSSIER CREDIT],
       MAX([NO INTERVENANT]) AS Intervenant,
       MAX([NO INTERVENANT GRP]) AS IntervGrp
FROM Q_Arc_UnpivotNature
GROUP BY Nature, Booking, [NO DOSSIER CREDIT];
```

# 1) Intervenants en retard — par nature — Archive + Flux

> “En retard” = **Flux** : sous-épisode ouvert (`*_StartDate` non null).  
> “Archive” : épisodes **clos** sur la période (EndDate dans [StartDate]…[EndDate]).  
> `Booking` filtrable via `[BookingLike]` (laisser vide pour tout).

## 1.1) VG — Intervenants + infos VG (Archive + Flux, filtre Booking)
```sql
SELECT 'FLUX' AS Source, 'VG' AS Nature,
       f.Booking, f.[NO DOSSIER CREDIT], f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
       f.[Production Date],
       f.VG_StartDate AS StartDate, Null AS EndDate, f.VG_NbJoursRetard AS Duree,
       f.VG_WorstInsuffisance AS WorstInsuffisance, f.VG_LastInsuffisance AS LastInsuffisance,
       f.Limite, f.Consommation, f.[Montant VG] AS Montant, f.DepassementNA
FROM T_Flux AS f
WHERE f.VG_StartDate IS NOT NULL
  AND ( [BookingLike] IS NULL OR f.Booking LIKE [BookingLike] )
UNION ALL
SELECT 'ARCHIVE' AS Source, 'VG',
       a.Booking, a.[NO DOSSIER CREDIT], a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
       a.[Production Date],
       a.VG_StartDate, a.VG_EndDate, a.VG_Duree,
       a.VG_WorstInsuffisance, a.VG_LastInsuffisance,
       a.Limite, a.Consommation, a.[Montant VG], a.DepassementNA
FROM T_Archive AS a
WHERE a.FlagVG='YES'
  AND a.VG_EndDate BETWEEN [StartDate] AND [EndDate]
  AND ( [BookingLike] IS NULL OR a.Booking LIKE [BookingLike] );
```

## 1.2) AM — Intervenants + infos AM (Archive + Flux, filtre Booking)
```sql
SELECT 'FLUX' AS Source, 'AM' AS Nature,
       f.Booking, f.[NO DOSSIER CREDIT], f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
       f.[Production Date],
       f.AM_StartDate AS StartDate, Null AS EndDate, f.AM_NbJoursRetard AS Duree,
       f.AM_WorstInsuffisance AS WorstInsuffisance, f.AM_LastInsuffisance AS LastInsuffisance,
       f.Limite, f.Consommation, f.[Montant AM] AS Montant, f.DepassementNA
FROM T_Flux AS f
WHERE f.AM_StartDate IS NOT NULL
  AND ( [BookingLike] IS NULL OR f.Booking LIKE [BookingLike] )
UNION ALL
SELECT 'ARCHIVE' AS Source, 'AM',
       a.Booking, a.[NO DOSSIER CREDIT], a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
       a.[Production Date],
       a.AM_StartDate, a.AM_EndDate, a.AM_Duree,
       a.AM_WorstInsuffisance, a.AM_LastInsuffisance,
       a.Limite, a.Consommation, a.[Montant AM], a.DepassementNA
FROM T_Archive AS a
WHERE a.FlagAM='YES'
  AND a.AM_EndDate BETWEEN [StartDate] AND [EndDate]
  AND ( [BookingLike] IS NULL OR a.Booking LIKE [BookingLike] );
```

## 1.3) SL — Intervenants + infos SL (Archive + Flux, filtre Booking)
```sql
SELECT 'FLUX' AS Source, 'SL' AS Nature,
       f.Booking, f.[NO DOSSIER CREDIT], f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
       f.[Production Date],
       f.SL_StartDate AS StartDate, Null AS EndDate, f.SL_NbJoursRetard AS Duree,
       f.SL_WorstInsuffisance AS WorstInsuffisance, f.SL_LastInsuffisance AS LastInsuffisance,
       f.Limite, f.Consommation, f.[Montant SL] AS Montant, f.DepassementNA
FROM T_Flux AS f
WHERE f.SL_StartDate IS NOT NULL
  AND ( [BookingLike] IS NULL OR f.Booking LIKE [BookingLike] )
UNION ALL
SELECT 'ARCHIVE' AS Source, 'SL',
       a.Booking, a.[NO DOSSIER CREDIT], a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
       a.[Production Date],
       a.SL_StartDate, a.SL_EndDate, a.SL_Duree,
       a.SL_WorstInsuffisance, a.SL_LastInsuffisance,
       a.Limite, a.Consommation, a.[Montant SL], a.DepassementNA
FROM T_Archive AS a
WHERE a.FlagSL='YES'
  AND a.SL_EndDate BETWEEN [StartDate] AND [EndDate]
  AND ( [BookingLike] IS NULL OR a.Booking LIKE [BookingLike] );
```

## 1.4) ALL — Intervenants en retard (toutes natures) — Archive + Flux (toutes bookings)
```sql
SELECT 'FLUX' AS Source, 'VG' AS Nature,
       f.Booking, f.[NO DOSSIER CREDIT], f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
       f.[Production Date], f.VG_StartDate AS StartDate, Null AS EndDate,
       f.VG_NbJoursRetard AS Duree, f.VG_WorstInsuffisance AS WorstInsuffisance, f.VG_LastInsuffisance AS LastInsuffisance,
       f.Limite, f.Consommation, f.[Montant VG] AS Montant, f.DepassementNA
FROM T_Flux AS f
WHERE f.VG_StartDate IS NOT NULL
UNION ALL
SELECT 'FLUX','AM', f.Booking, f.[NO DOSSIER CREDIT], f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
       f.[Production Date], f.AM_StartDate, Null,
       f.AM_NbJoursRetard, f.AM_WorstInsuffisance, f.AM_LastInsuffisance,
       f.Limite, f.Consommation, f.[Montant AM], f.DepassementNA
FROM T_Flux AS f
WHERE f.AM_StartDate IS NOT NULL
UNION ALL
SELECT 'FLUX','SL', f.Booking, f.[NO DOSSIER CREDIT], f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
       f.[Production Date], f.SL_StartDate, Null,
       f.SL_NbJoursRetard, f.SL_WorstInsuffisance, f.SL_LastInsuffisance,
       f.Limite, f.Consommation, f.[Montant SL], f.DepassementNA
FROM T_Flux AS f
WHERE f.SL_StartDate IS NOT NULL
UNION ALL
SELECT 'ARCHIVE','VG', a.Booking, a.[NO DOSSIER CREDIT], a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
       a.[Production Date], a.VG_StartDate, a.VG_EndDate,
       a.VG_Duree, a.VG_WorstInsuffisance, a.VG_LastInsuffisance,
       a.Limite, a.Consommation, a.[Montant VG], a.DepassementNA
FROM T_Archive AS a
WHERE a.FlagVG='YES' AND a.VG_EndDate BETWEEN [StartDate] AND [EndDate]
UNION ALL
SELECT 'ARCHIVE','AM', a.Booking, a.[NO DOSSIER CREDIT], a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
       a.[Production Date], a.AM_StartDate, a.AM_EndDate,
       a.AM_Duree, a.AM_WorstInsuffisance, a.AM_LastInsuffisance,
       a.Limite, a.Consommation, a.[Montant AM], a.DepassementNA
FROM T_Archive AS a
WHERE a.FlagAM='YES' AND a.AM_EndDate BETWEEN [StartDate] AND [EndDate]
UNION ALL
SELECT 'ARCHIVE','SL', a.Booking, a.[NO DOSSIER CREDIT], a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
       a.[Production Date], a.SL_StartDate, a.SL_EndDate,
       a.SL_Duree, a.SL_WorstInsuffisance, a.SL_LastInsuffisance,
       a.Limite, a.Consommation, a.[Montant SL], a.DepassementNA
FROM T_Archive AS a
WHERE a.FlagSL='YES' AND a.SL_EndDate BETWEEN [StartDate] AND [EndDate];
```

---

# 2) TOP 30 Intervenants “en retard” (Archive + Flux)

> Classements par **nombre d’épisodes** : `Flux` (ouverts) + `Archive` (clos sur période).

## 2.1) TOP 30 VG (Archive + Flux)
```sql
SELECT TOP 30 Intervenant, COUNT(*) AS NbEpisodes
FROM (
  SELECT f.[NO INTERVENANT] AS Intervenant
  FROM T_Flux AS f
  WHERE f.VG_StartDate IS NOT NULL
  UNION ALL
  SELECT a.[NO INTERVENANT]
  FROM T_Archive AS a
  WHERE a.FlagVG='YES' AND a.VG_EndDate BETWEEN [StartDate] AND [EndDate]
) AS U
GROUP BY Intervenant
ORDER BY COUNT(*) DESC;
```

## 2.2) TOP 30 AM (Archive + Flux)
```sql
SELECT TOP 30 Intervenant, COUNT(*) AS NbEpisodes
FROM (
  SELECT f.[NO INTERVENANT] AS Intervenant
  FROM T_Flux AS f
  WHERE f.AM_StartDate IS NOT NULL
  UNION ALL
  SELECT a.[NO INTERVENANT]
  FROM T_Archive AS a
  WHERE a.FlagAM='YES' AND a.AM_EndDate BETWEEN [StartDate] AND [EndDate]
) AS U
GROUP BY Intervenant
ORDER BY COUNT(*) DESC;
```

## 2.3) TOP 30 SL (Archive + Flux)
```sql
SELECT TOP 30 Intervenant, COUNT(*) AS NbEpisodes
FROM (
  SELECT f.[NO INTERVENANT] AS Intervenant
  FROM T_Flux AS f
  WHERE f.SL_StartDate IS NOT NULL
  UNION ALL
  SELECT a.[NO INTERVENANT]
  FROM T_Archive AS a
  WHERE a.FlagSL='YES' AND a.SL_EndDate BETWEEN [StartDate] AND [EndDate]
) AS U
GROUP BY Intervenant
ORDER BY COUNT(*) DESC;
```

## 2.4) TOP 30 AM & SL combinés (Archive + Flux)
```sql
SELECT TOP 30 Intervenant, COUNT(*) AS NbEpisodes
FROM (
  SELECT f.[NO INTERVENANT] AS Intervenant
  FROM T_Flux AS f
  WHERE f.AM_StartDate IS NOT NULL
  UNION ALL
  SELECT f.[NO INTERVENANT]
  FROM T_Flux AS f
  WHERE f.SL_StartDate IS NOT NULL
  UNION ALL
  SELECT a.[NO INTERVENANT]
  FROM T_Archive AS a
  WHERE a.FlagAM='YES' AND a.AM_EndDate BETWEEN [StartDate] AND [EndDate]
  UNION ALL
  SELECT a.[NO INTERVENANT]
  FROM T_Archive AS a
  WHERE a.FlagSL='YES' AND a.SL_EndDate BETWEEN [StartDate] AND [EndDate]
) AS U
GROUP BY Intervenant
ORDER BY COUNT(*) DESC;
```

> Variante : ajoutez `, Nature` dans les sous-requêtes et faites un **crosstab** pour détailler AM vs SL.


# Évolutions temporelles par intervenant

> Utilise `T_HistoryInsuff` pour compter les **jours où la nature est négative** (`Insuffisance < 0`).  
> Mapping de l’intervenant via `Q_KeyNature_Interv`.

## 1) VG — Nombre de retards (mensuel) pour un intervenant
```sql
SELECT Format(o.[Production Date],'yyyy-mm') AS YearMonth,
       COUNT(*) AS NbRetards
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='VG' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY Format(o.[Production Date],'yyyy-mm')
ORDER BY YearMonth;
```

## 2) AM — Nombre de retards (mensuel) pour un intervenant
```sql
SELECT Format(o.[Production Date],'yyyy-mm') AS YearMonth,
       COUNT(*) AS NbRetards
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='AM' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY Format(o.[Production Date],'yyyy-mm')
ORDER BY YearMonth;
```

## 3) SL — Nombre de retards (mensuel) pour un intervenant
```sql
SELECT Format(o.[Production Date],'yyyy-mm') AS YearMonth,
       COUNT(*) AS NbRetards
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='SL' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY Format(o.[Production Date],'yyyy-mm')
ORDER BY YearMonth;
```


# Évolutions “montant + retard” par intervenant

> Même logique que §3, avec agrégats sur **Insuffisance** (magnitude) et **Montant**.

## 1) VG — Montant en retard (somme & moyenne) + nombre de retards (mensuel)
```sql
SELECT Format(o.[Production Date],'yyyy-mm') AS YearMonth,
       COUNT(*) AS NbRetards,
       ROUND(SUM(ABS(Nz(o.Insuffisance,0))),2) AS Somme_Insuff,
       ROUND(AVG(ABS(Nz(o.Insuffisance,0))),2) AS Moy_Insuff,
       ROUND(SUM(Nz(o.[Montant],0)),2) AS Somme_Montant
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='VG' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY Format(o.[Production Date],'yyyy-mm')
ORDER BY YearMonth;
```

## 2) AM — Montant en retard (somme & moyenne) + nombre de retards (mensuel)
```sql
SELECT Format(o.[Production Date],'yyyy-mm') AS YearMonth,
       COUNT(*) AS NbRetards,
       ROUND(SUM(ABS(Nz(o.Insuffisance,0))),2) AS Somme_Insuff,
       ROUND(AVG(ABS(Nz(o.Insuffisance,0))),2) AS Moy_Insuff,
       ROUND(SUM(Nz(o.[Montant],0)),2) AS Somme_Montant
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='AM' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY Format(o.[Production Date],'yyyy-mm')
ORDER BY YearMonth;
```

## 3) SL — Montant en retard (somme & moyenne) + nombre de retards (mensuel)
```sql
SELECT Format(o.[Production Date],'yyyy-mm') AS YearMonth,
       COUNT(*) AS NbRetards,
       ROUND(SUM(ABS(Nz(o.Insuffisance,0))),2) AS Somme_Insuff,
       ROUND(AVG(ABS(Nz(o.Insuffisance,0))),2) AS Moy_Insuff,
       ROUND(SUM(Nz(o.[Montant],0)),2) AS Somme_Montant
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='SL' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY Format(o.[Production Date],'yyyy-mm')
ORDER BY YearMonth;
```

---

# Variantes utiles

## .1) Évolutions **journalières** (remplacer le mois par la date)
> Exemple pour VG (adapter AM/SL) :
```sql
SELECT o.[Production Date] AS J,
       COUNT(*) AS NbRetards,
       ROUND(SUM(ABS(Nz(o.Insuffisance,0))),2) AS Somme_Insuff
FROM T_HistoryInsuff AS o
INNER JOIN Q_KeyNature_Interv AS k
  ON o.Booking=k.Booking AND o.[NO DOSSIER CREDIT]=k.[NO DOSSIER CREDIT] AND o.Nature=k.Nature
WHERE o.Nature='VG' AND o.Insuffisance<0 AND k.Intervenant=[IntervParam]
GROUP BY o.[Production Date]
ORDER BY o.[Production Date];
```

## 2) TOP 30 par **jours de retard cumulés** (Archive + Flux)
> Somme des durées pour Archive (`Duree`) + **NbJoursRetard actuel** pour Flux.
```sql
SELECT TOP 30 Intervenant, SUM(TotalJours) AS JoursCumules
FROM (
  SELECT a.[NO INTERVENANT] AS Intervenant, Nz(a.VG_Duree,0)+Nz(a.AM_Duree,0)+Nz(a.SL_Duree,0) AS TotalJours
  FROM T_Archive AS a
  WHERE a.[Production Date] BETWEEN [StartDate] AND [EndDate]
  UNION ALL
  SELECT f.[NO INTERVENANT],
         Nz(f.VG_NbJoursRetard,0)+Nz(f.AM_NbJoursRetard,0)+Nz(f.SL_NbJoursRetard,0) AS TotalJours
  FROM T_Flux AS f
  WHERE f.VG_StartDate IS NOT NULL OR f.AM_StartDate IS NOT NULL OR f.SL_StartDate IS NOT NULL
) AS U
GROUP BY Intervenant
ORDER BY SUM(TotalJours) DESC;
```

# 1) Volumétrie & Séries temporelles

## 1.1) Clôtures par **jour** et **nature**  
Comptage quotidien des épisodes clôturés par nature.
```sql
SELECT EndDate, Nature, COUNT(*) AS NbClosures
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY EndDate, Nature
ORDER BY EndDate, Nature;
```

## 1.2) Clôtures par **mois (AAAA‑MM)** et **nature**  
Synthèse mensuelle des clôtures par nature.
```sql
SELECT Format([EndDate],'yyyy-mm') AS YearMonth, Nature, COUNT(*) AS NbClosures
FROM Q_Arc_UnpivotNature
GROUP BY Format([EndDate],'yyyy-mm'), Nature
ORDER BY YearMonth, Nature;
```

## 1.3) Clôtures par **jour de semaine**  
Répartition des clôtures par jour de semaine (1=Lundi avec `Weekday(...,2)`).
```sql
SELECT Weekday([EndDate],2) AS DowIndex,
       WeekdayName(Weekday([EndDate],2), True, 2) AS Dow,
       Nature, COUNT(*) AS NbClosures
FROM Q_Arc_UnpivotNature
GROUP BY Weekday([EndDate],2), WeekdayName(Weekday([EndDate],2), True, 2), Nature
ORDER BY DowIndex, Nature;
```

## 1.4) **Clés distinctes** fermées par **jour** (toutes natures)  
Nombre d’entrées d’archive par date de production.
```sql
SELECT [Production Date] AS ProdDate, COUNT(*) AS NbKeys
FROM T_Archive
WHERE [Production Date] BETWEEN [StartDate] AND [EndDate]
GROUP BY [Production Date]
ORDER BY [Production Date];
```

---

# 2) Durée d’épisode

## 2.1) Durée moyenne et extrêmes par nature  
Statistiques de base sur `Duree` (jours).
```sql
SELECT Nature,
       COUNT(*) AS N,
       ROUND(AVG(Duree),2) AS Duree_Moy,
       MIN(Duree) AS Duree_Min,
       MAX(Duree) AS Duree_Max
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY Nature;
```

## 2.2) Buckets de durée (1–7 / 8–30 / 31–90 / >90 jours)  
Distribution simple des durées par classe.
```sql
SELECT Nature,
       SWITCH(
           Duree BETWEEN 1 AND 7, '01-07',
           Duree BETWEEN 8 AND 30, '08-30',
           Duree BETWEEN 31 AND 90, '31-90',
           Duree > 90, '>90',
           TRUE, '0'
       ) AS Bucket,
       COUNT(*) AS Nb
FROM Q_Arc_UnpivotNature
GROUP BY Nature,
         SWITCH(
           Duree BETWEEN 1 AND 7, '01-07',
           Duree BETWEEN 8 AND 30, '08-30',
           Duree BETWEEN 31 AND 90, '31-90',
           Duree > 90, '>90',
           TRUE, '0'
         )
ORDER BY Nature, Bucket;
```

## 2.3) SLA (exemple : **≤ 5** jours)  
Part des épisodes clôturés dans un délai cible.
```sql
SELECT Nature,
       SUM(IIf(Duree<=5,1,0)) AS Nb_SLA_OK,
       COUNT(*) AS Nb_Total,
       ROUND(100.0*SUM(IIf(Duree<=5,1,0))/COUNT(*),2) AS Pct_SLA_OK
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY Nature;
```

---

# 3) Sévérité & Ratios

## 3.1) Sévérité moyenne / max (|WorstInsuffisance|)  
Analyse de la magnitude d’insuffisance la plus forte.
```sql
SELECT Nature,
       ROUND(AVG(ABS(Nz(WorstInsuffisance,0))),2) AS AvgSeverity,
       MAX(ABS(Nz(WorstInsuffisance,0))) AS MaxSeverity
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY Nature;
```

## 3.2) Ratio d’utilisation (|Consommation| / Limite) à la clôture  
Mesure d’utilisation moyenne du plafond à la fin de l’épisode.
```sql
SELECT Nature,
       ROUND(AVG(IIf(Limite=0,Null,ABS(Consommation)/Limite)),3) AS AvgUsageRate
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY Nature;
```

## 3.3) Montants moyens & totaux à la clôture  
Lecture agrégée des montants par nature.
```sql
SELECT Nature,
       ROUND(AVG(Nz(Montant,0)),2) AS Montant_Moyen,
       ROUND(SUM(Nz(Montant,0)),2) AS Montant_Total
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY Nature;
```

---

# 4) Multi‑natures (même clé, même jour)

## 4.1) Pattern de clôture (VG / AM / SL / combinaisons)  
Typologie des combinaisons de flags par date de production.
```sql
SELECT
    [Production Date],
    IIf(FlagVG='YES',1,0)+IIf(FlagAM='YES',1,0)+IIf(FlagSL='YES',1,0) AS FlagsCount,
    IIf(FlagVG='YES' AND FlagAM='YES' AND FlagSL='YES','VG+AM+SL',
       IIf(FlagVG='YES' AND FlagAM='YES','VG+AM',
       IIf(FlagVG='YES' AND FlagSL='YES','VG+SL',
       IIf(FlagAM='YES' AND FlagSL='YES','AM+SL',
       IIf(FlagVG='YES','VG',
       IIf(FlagAM='YES','AM',
       IIf(FlagSL='YES','SL','NONE'))))))) AS Pattern,
    COUNT(*) AS Nb
FROM T_Archive
WHERE [Production Date] BETWEEN [StartDate] AND [EndDate]
GROUP BY [Production Date],
    IIf(FlagVG='YES',1,0)+IIf(FlagAM='YES',1,0)+IIf(FlagSL='YES',1,0),
    IIf(FlagVG='YES' AND FlagAM='YES' AND FlagSL='YES','VG+AM+SL',
       IIf(FlagVG='YES' AND FlagAM='YES','VG+AM',
       IIf(FlagVG='YES' AND FlagSL='YES','VG+SL',
       IIf(FlagAM='YES' AND FlagSL='YES','AM+SL',
       IIf(FlagVG='YES','VG',
       IIf(FlagAM='YES','AM',
       IIf(FlagSL='YES','SL','NONE')))))))
ORDER BY [Production Date], Pattern;
```

## 4.2) Part de multi‑natures (≥ 2 flags) par **mois**  
Poids des co‑clôtures dans le flux mensuel.
```sql
SELECT
    Format([Production Date],'yyyy-mm') AS YearMonth,
    SUM(IIf( (IIf(FlagVG='YES',1,0)+IIf(FlagAM='YES',1,0)+IIf(FlagSL='YES',1,0))>=2 ,1,0)) AS Nb_Multi,
    COUNT(*) AS Nb_Total,
    ROUND(100.0*SUM(IIf( (IIf(FlagVG='YES',1,0)+IIf(FlagAM='YES',1,0)+IIf(FlagSL='YES',1,0))>=2 ,1,0))/COUNT(*),2) AS Pct_Multi
FROM T_Archive
GROUP BY Format([Production Date],'yyyy-mm')
ORDER BY YearMonth;
```

---

# 5) Intervenants & Groupes

## 5.1) Clôtures par intervenant (Top 20)  
Volume de clôtures par nature pour les 20 intervenants les plus actifs.
```sql
SELECT [NO INTERVENANT] AS Intervenant, Nature, COUNT(*) AS NbClosures
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO INTERVENANT], Nature
ORDER BY COUNT(*) DESC;
```

## 5.2) Durée moyenne par intervenant (Top 20)  
Focus sur les intervenants dont les épisodes durent le plus.
```sql
SELECT TOP 20 [NO INTERVENANT] AS Intervenant, Nature,
       COUNT(*) AS N, ROUND(AVG(Duree),2) AS Duree_Moy
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO INTERVENANT], Nature
ORDER BY AVG(Duree) DESC;
```

## 5.3) Sévérité max par intervenant (Top 20)  
Valeur d’insuffisance maximale observée par intervenant et nature.
```sql
SELECT TOP 20 [NO INTERVENANT] AS Intervenant, Nature,
       MAX(ABS(Nz(WorstInsuffisance,0))) AS MaxSeverity
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO INTERVENANT], Nature
ORDER BY MAX(ABS(Nz(WorstInsuffisance,0))) DESC;
```

## 5.4) Clôtures par **groupe d’intervenants**  
Vue agrégée par groupe.
```sql
SELECT [NO INTERVENANT GRP] AS IntervGrp, Nature, COUNT(*) AS NbClosures
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO INTERVENANT GRP], Nature
ORDER BY COUNT(*) DESC;
```

---

# 6) Snapshot du portefeuille (ouvert **à date**)

> Un épisode est “ouvert” le jour **D** si dans `T_HistoryInsuff` la nature a **Insuffisance < 0** à la date **D**.

## 6.1) `Q_Open_OnDate` — Liste des clés/natures ouvertes à la date  
```sql
SELECT Booking, [NO DOSSIER CREDIT], Nature
FROM T_HistoryInsuff
WHERE [Production Date]=[AsOfDate] AND Insuffisance<0;
```

## 6.2) Volume d’ouvertures par nature (à la date)  
```sql
SELECT Nature, COUNT(*) AS NbOpen
FROM Q_Open_OnDate;
```

## 6.3) Top intervenants sur l’**ouvert** (à la date)  
```sql
SELECT h.[NO INTERVENANT], h.[NO INTERVENANT GRP], h.Nature, COUNT(*) AS NbOpen
FROM (
    SELECT th.Booking, th.[NO DOSSIER CREDIT], th.Nature
    FROM T_HistoryInsuff AS th
    WHERE th.[Production Date]=[AsOfDate] AND th.Insuffisance<0
) AS o
LEFT JOIN Q_Arc_UnpivotNature AS h
  ON o.Booking=h.Booking AND o.[NO DOSSIER CREDIT]=h.[NO DOSSIER CREDIT] AND o.Nature=h.Nature
GROUP BY h.[NO INTERVENANT], h.[NO INTERVENANT GRP], h.Nature
ORDER BY COUNT(*) DESC;
```

---

# 7) Tendances & Récurrence

## 7.1) Clôtures **cumulées** par nature  
```sql
SELECT A.EndDate, A.Nature,
       (SELECT COUNT(*)
        FROM Q_Arc_UnpivotNature B
        WHERE B.Nature=A.Nature AND B.EndDate<=A.EndDate) AS CumClosures
FROM (SELECT DISTINCT EndDate, Nature FROM Q_Arc_UnpivotNature) AS A
ORDER BY A.EndDate, A.Nature;
```

## 7.2) Clôtures **glissantes 7 jours** par nature  
```sql
SELECT A.EndDate, A.Nature,
       (SELECT COUNT(*)
        FROM Q_Arc_UnpivotNature B
        WHERE B.Nature=A.Nature
          AND B.EndDate BETWEEN DateAdd('d',-6,A.EndDate) AND A.EndDate) AS Closures_7d
FROM (SELECT DISTINCT EndDate, Nature FROM Q_Arc_UnpivotNature) AS A
ORDER BY A.EndDate, A.Nature;
```

## 7.3) Nombre d’épisodes par **clé** (proxy de ré‑ouverture)  
```sql
SELECT Nature, Booking, [NO DOSSIER CREDIT], COUNT(*) AS NbEpisodes
FROM Q_Arc_UnpivotNature
GROUP BY Nature, Booking, [NO DOSSIER CREDIT]
ORDER BY Nature, NbEpisodes DESC;
```

## 7.4) Nombre **moyen** d’épisodes par clé (par nature)  
```sql
SELECT Nature, ROUND(AVG(NbEpisodes),2) AS AvgEpisodesPerKey
FROM (
  SELECT Nature, Booking, [NO DOSSIER CREDIT], COUNT(*) AS NbEpisodes
  FROM Q_Arc_UnpivotNature
  GROUP BY Nature, Booking, [NO DOSSIER CREDIT]
) AS X
GROUP BY Nature;
```

---

# 8) Qualité des données

## 8.1) Doublons potentiels (même clé, même Production Date)  
```sql
SELECT [Production Date], Booking, [NO DOSSIER CREDIT], COUNT(*) AS N
FROM T_Archive
GROUP BY [Production Date], Booking, [NO DOSSIER CREDIT]
HAVING COUNT(*)>1
ORDER BY [Production Date], Booking, [NO DOSSIER CREDIT];
```

## 8.2) Lignes d’archive **sans aucun flag** (anomalie)  
```sql
SELECT *
FROM T_Archive
WHERE Nz(FlagVG,'NO')='NO' AND Nz(FlagAM,'NO')='NO' AND Nz(FlagSL,'NO')='NO';
```

## 8.3) Dates inversées (Start > End)  
```sql
SELECT *
FROM Q_Arc_UnpivotNature
WHERE StartDate>EndDate;
```

## 8.4) Durée manquante ou négative (alors que Flag='YES')  
```sql
SELECT *
FROM Q_Arc_UnpivotNature
WHERE Flag='YES' AND (Duree IS NULL OR Duree<=0);
```

---

# 9) Tableaux de bord “one‑page”

## 9.1) KPI **quotidiens** (jour **[J]**)  
```sql
SELECT
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE EndDate=[J]) AS NbClosures_J,
  (SELECT COUNT(*) FROM T_Archive WHERE [Production Date]=[J]) AS NbKeys_J,
  (SELECT ROUND(AVG(Duree),2) FROM Q_Arc_UnpivotNature WHERE EndDate=[J]) AS DureeMoy_J,
  (SELECT SUM(IIf(Duree<=5,1,0)) FROM Q_Arc_UnpivotNature WHERE EndDate=[J]) AS SLA_OK_J,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE EndDate=[J] AND Nature='VG') AS VG_J,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE EndDate=[J] AND Nature='AM') AS AM_J,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE EndDate=[J] AND Nature='SL') AS SL_J;
```

## 9.2) KPI **mensuels** (AAAA‑MM)  
```sql
SELECT
  [YYYYMM] AS YearMonth,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE Format([EndDate],'yyyy-mm')=[YYYYMM]) AS NbClosures,
  (SELECT ROUND(AVG(Duree),2) FROM Q_Arc_UnpivotNature WHERE Format([EndDate],'yyyy-mm')=[YYYYMM]) AS DureeMoy,
  (SELECT ROUND(100.0*SUM(IIf(Duree<=5,1,0))/COUNT(*),2) FROM Q_Arc_UnpivotNature WHERE Format([EndDate],'yyyy-mm')=[YYYYMM]) AS PctSLA5j,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE Format([EndDate],'yyyy-mm')=[YYYYMM] AND Nature='VG') AS VG,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE Format([EndDate],'yyyy-mm')=[YYYYMM] AND Nature='AM') AS AM,
  (SELECT COUNT(*) FROM Q_Arc_UnpivotNature WHERE Format([EndDate],'yyyy-mm')=[YYYYMM] AND Nature='SL') AS SL
FROM (
  SELECT DISTINCT Format([EndDate],'yyyy-mm') AS [YYYYMM]
  FROM Q_Arc_UnpivotNature
) AS M
ORDER BY YearMonth;
```

---

# 10) Démarrages vs Clôtures (flux mensuel)

## 10.1) Démarrages par mois (via **StartDate**)  
```sql
SELECT Format([StartDate],'yyyy-mm') AS YearMonth, Nature, COUNT(*) AS NbStarts
FROM Q_Arc_UnpivotNature
GROUP BY Format([StartDate],'yyyy-mm'), Nature
ORDER BY YearMonth, Nature;
```

## 10.2) **Balance** mensuelle (Starts − Closures) par nature  
```sql
SELECT M.YM AS YearMonth, M.Nature,
       Nz(S.NbStarts,0) AS NbStarts,
       Nz(C.NbClosures,0) AS NbClosures,
       Nz(S.NbStarts,0)-Nz(C.NbClosures,0) AS Delta
FROM (
  SELECT DISTINCT Format([EndDate],'yyyy-mm') AS YM, Nature FROM Q_Arc_UnpivotNature
  UNION
  SELECT DISTINCT Format([StartDate],'yyyy-mm') AS YM, Nature FROM Q_Arc_UnpivotNature
) AS M
LEFT JOIN (
  SELECT Format([StartDate],'yyyy-mm') AS YM, Nature, COUNT(*) AS NbStarts
  FROM Q_Arc_UnpivotNature
  GROUP BY Format([StartDate],'yyyy-mm'), Nature
) AS S ON S.YM=M.YM AND S.Nature=M.Nature
LEFT JOIN (
  SELECT Format([EndDate],'yyyy-mm') AS YM, Nature, COUNT(*) AS NbClosures
  FROM Q_Arc_UnpivotNature
  GROUP BY Format([EndDate],'yyyy-mm'), Nature
) AS C ON C.YM=M.YM AND C.Nature=M.Nature
ORDER BY YearMonth, Nature;
```

---

# 11) Réouvertures / Récidive

## 11.1) % Réouverture ≤ **[WindowDays]** jours (même nature, même clé)  
```sql
SELECT a.Nature,
       COUNT(*) AS Nb_Closures,
       SUM(IIf(EXISTS
         (SELECT 1 FROM Q_Arc_UnpivotNature b
           WHERE b.Booking=a.Booking
             AND b.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT]
             AND b.Nature=a.Nature
             AND b.StartDate > a.EndDate
             AND b.StartDate <= DateAdd('d',[WindowDays],a.EndDate)),1,0)) AS Nb_Reopen,
       ROUND(100.0*SUM(IIf(EXISTS
         (SELECT 1 FROM Q_Arc_UnpivotNature b
           WHERE b.Booking=a.Booking
             AND b.[NO DOSSIER CREDIT]=a.[NO DOSSIER CREDIT]
             AND b.Nature=a.Nature
             AND b.StartDate > a.EndDate
             AND b.StartDate <= DateAdd('d',[WindowDays],a.EndDate)),1,0))/COUNT(*),2) AS Pct_Reopen
FROM Q_Arc_UnpivotNature AS a
WHERE a.EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY a.Nature;
```

## 11.2) Réouverture **≤ 30 jours** (même nature, même clé) — volume  
```sql
SELECT a.Nature, COUNT(*) AS Nb_Reouvertures_30j
FROM Q_Arc_UnpivotNature AS a
INNER JOIN Q_Arc_UnpivotNature AS b
   ON a.Booking=b.Booking
  AND a.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT]
  AND a.Nature=b.Nature
  AND b.StartDate > a.EndDate
  AND b.StartDate <= DateAdd('d',30,a.EndDate)
WHERE a.EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY a.Nature;
```

## 11.3) Transitions **inter‑natures** ≤ 7 jours (matrice VG→AM, etc.)  
```sql
SELECT a.Nature AS Nature_From, b.Nature AS Nature_To, COUNT(*) AS Nb
FROM Q_Arc_UnpivotNature a
INNER JOIN Q_Arc_UnpivotNature b
  ON a.Booking=b.Booking
 AND a.[NO DOSSIER CREDIT]=b.[NO DOSSIER CREDIT]
 AND b.StartDate > a.EndDate
 AND b.StartDate <= DateAdd('d',7,a.EndDate)
WHERE a.EndDate BETWEEN [StartDate] AND [EndDate]
  AND a.Nature<>b.Nature
GROUP BY a.Nature, b.Nature
ORDER BY a.Nature, b.Nature;
```

---

# 12) “Time to worst” (position du pire)

## 12.1) Jours jusqu’au pire & position du pire dans l’épisode  
Calcule (WorstDate − StartDate) et sa position relative dans la durée.
```sql
SELECT Nature,
       ROUND(AVG(Nz(DateDiff('d',StartDate,WorstDate),Null)),2) AS Jours_Jusqu_Au_Pire,
       ROUND(AVG(IIf(Duree Is Null Or Duree=0, Null, Nz(DateDiff('d',StartDate,WorstDate),0)/Duree)),3) AS Pct_Position_Pire
FROM Q_Arc_UnpivotNature_W
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
  AND WorstDate IS NOT NULL
GROUP BY Nature;
```

---

# 13) “Open” en temps réel (depuis `T_Flux`)

## 13.1) Épisodes **ouverts** actuellement (par nature)  
```sql
SELECT
    SUM(IIf(VG_StartDate IS NOT NULL,1,0)) AS Open_VG,
    SUM(IIf(AM_StartDate IS NOT NULL,1,0)) AS Open_AM,
    SUM(IIf(SL_StartDate IS NOT NULL,1,0)) AS Open_SL
FROM T_Flux;
```

## 13.2) **Âge moyen** des épisodes ouverts (par nature)  
```sql
SELECT
    ROUND(AVG(Nz(VG_NbJoursRetard,0)),2) AS AgeMoy_VG,
    ROUND(AVG(Nz(AM_NbJoursRetard,0)),2) AS AgeMoy_AM,
    ROUND(AVG(Nz(SL_NbJoursRetard,0)),2) AS AgeMoy_SL
FROM T_Flux
WHERE VG_StartDate IS NOT NULL OR AM_StartDate IS NOT NULL OR SL_StartDate IS NOT NULL;
```

---

# 14) Cohortes & Classements “business”

## 14.1) Répartition des durées par **mois de StartDate** (cohortes)  
```sql
SELECT Format([StartDate],'yyyy-mm') AS CohortMonth, Nature,
       COUNT(*) AS N,
       ROUND(AVG(Duree),2) AS Duree_Moy,
       SUM(IIf(Duree<=7,1,0)) AS D_01_07,
       SUM(IIf(Duree BETWEEN 8 AND 30,1,0)) AS D_08_30,
       SUM(IIf(Duree BETWEEN 31 AND 90,1,0)) AS D_31_90,
       SUM(IIf(Duree>90,1,0)) AS D_GT_90
FROM Q_Arc_UnpivotNature
GROUP BY Format([StartDate],'yyyy-mm'), Nature
ORDER BY CohortMonth, Nature;
```

## 14.2) Top 20 **épisodes les plus longs** (par nature)  
```sql
SELECT TOP 20 Nature, Booking, [NO DOSSIER CREDIT], StartDate, EndDate, Duree
FROM Q_Arc_UnpivotNature
ORDER BY Duree DESC, EndDate DESC;
```

## 14.3) Top 20 **pires insuffisances** (magnitude)  
```sql
SELECT TOP 20 Nature, Booking, [NO DOSSIER CREDIT], EndDate,
       ABS(Nz(WorstInsuffisance,0)) AS Severity
FROM Q_Arc_UnpivotNature
ORDER BY ABS(Nz(WorstInsuffisance,0)) DESC, EndDate DESC;
```

## 14.4) Clients (NDC) les plus “récidivistes” (nb de fermetures)  
```sql
SELECT [NO DOSSIER CREDIT], COUNT(*) AS NbClosures
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO DOSSIER CREDIT]
ORDER BY COUNT(*) DESC;
```

## 14.5) Ratio d’utilisation moyen à la clôture **par intervenant**  
```sql
SELECT [NO INTERVENANT],
       ROUND(AVG(IIf(Limite=0,Null,ABS(Consommation)/Limite)),3) AS AvgUsageRate,
       COUNT(*) AS N
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO INTERVENANT]
HAVING COUNT(*) >= 10
ORDER BY AvgUsageRate DESC;
```

---

# 15) Tableaux **croisés** (crosstab)

## 15.1) Clôtures par **mois × nature**  
```sql
TRANSFORM COUNT(*) AS Nb
SELECT Format([EndDate],'yyyy-mm') AS YearMonth
FROM Q_Arc_UnpivotNature
GROUP BY Format([EndDate],'yyyy-mm')
PIVOT Nature IN ('VG','AM','SL');
```

## 15.2) SLA ≤ **[SLA_Days]** jours par **mois × nature**  
```sql
TRANSFORM SUM(IIf(Duree<=[SLA_Days],1,0)) AS SLA_OK
SELECT Format([EndDate],'yyyy-mm') AS YearMonth
FROM Q_Arc_UnpivotNature
GROUP BY Format([EndDate],'yyyy-mm')
PIVOT Nature IN ('VG','AM','SL');
```

## 15.3) Part DepassementNA='YES' à la clôture **par mois × nature**  
```sql
TRANSFORM ROUND(100.0*AVG(IIf(DepassementNA='YES',1,0)),2) AS Pct_NA
SELECT Format([EndDate],'yyyy-mm') AS YearMonth
FROM Q_Arc_UnpivotNature
GROUP BY Format([EndDate],'yyyy-mm')
PIVOT Nature IN ('VG','AM','SL');
```

## 15.4) Clôtures **Top 15 intervenants × nature** (crosstab en 2 temps)  
Créer d’abord la table temporaire `Q_TopIntervenants15`, puis le crosstab.
```sql
SELECT TOP 15 [NO INTERVENANT] AS Intervenant, COUNT(*) AS N
INTO Q_TopIntervenants15
FROM Q_Arc_UnpivotNature
WHERE EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY [NO INTERVENANT]
ORDER BY COUNT(*) DESC;

TRANSFORM COUNT(*) AS Nb
SELECT a.Intervenant
FROM Q_Arc_UnpivotNature u
INNER JOIN Q_TopIntervenants15 a
  ON u.[NO INTERVENANT]=a.Intervenant
WHERE u.EndDate BETWEEN [StartDate] AND [EndDate]
GROUP BY a.Intervenant
PIVOT u.Nature IN ('VG','AM','SL');
```

---

## Notes & Conseils
- Commencez par **créer et enregistrer** `Q_Arc_UnpivotNature` (et `Q_Arc_UnpivotNature_W` si besoin).
- Pour de bonnes perfs, conservez les index proposés (sur `T_Archive` : `(Booking, [NO DOSSIER CREDIT])` et `VG/AM/SL_EndDate`; sur `T_Today` : `(Booking, [NO DOSSIER CREDIT])`).
- Paramétrage simple via des formulaires Access avec des zones `[StartDate]`, `[EndDate]`, `[AsOfDate]`, `[WindowDays]`, `[SLA_Days]`, `[J]`.
- Pour exporter vers Excel : enregistrez les requêtes puis **Données externes → Exporter → Excel**.

Bonnes analyses !
