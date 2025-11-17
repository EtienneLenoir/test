Parfait — j’aligne tes requêtes avec ta nouvelle logique (flux “résidualisé” vs archive) et MS-Access (parenthèses partout dans les JOIN).
L’idée : créer une vue pivot unique qui mixe Archive (épisodes clos) + Flux résidualisé (part du flux non déjà archivée), puis réécrire les requêtes qui mélangeaient Archive/Flux pour éviter tout double-comptage.


0) Pré-requis (vues de base)
	Tes 0.1 et 0.2 restent valides (je les garde identiques). On ajoute 2 vues clefs : Q_Flux_Residualized puis Q_Episodes_All.
0.1) Q_Arc_UnpivotNature (inchangé)

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
SELECT ID, Booking, [NO DOSSIER CREDIT], 'AM',
       AM_StartDate, AM_EndDate, AM_Duree,
       AM_WorstInsuffisance, AM_LastInsuffisance,
       AM_LastUpdateDate,
       [Montant AM], FlagAM,
       [Production Date], Limite, Consommation, DepassementNA,
       [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive
WHERE FlagAM='YES'
UNION ALL
SELECT ID, Booking, [NO DOSSIER CREDIT], 'SL',
       SL_StartDate, SL_EndDate, SL_Duree,
       SL_WorstInsuffisance, SL_LastInsuffisance,
       SL_LastUpdateDate,
       [Montant SL], FlagSL,
       [Production Date], Limite, Consommation, DepassementNA,
       [NO INTERVENANT], [NO INTERVENANT GRP]
FROM T_Archive
WHERE FlagSL='YES';

0.2) Q_Arc_UnpivotNature_W (inchangé – ajoute WorstDate)

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

0.3) (optionnel) Q_KeyNature_Interv (inchangé)

SELECT Nature, Booking, [NO DOSSIER CREDIT],
       MAX([NO INTERVENANT]) AS Intervenant,
       MAX([NO INTERVENANT GRP]) AS IntervGrp
FROM Q_Arc_UnpivotNature
GROUP BY Nature, Booking, [NO DOSSIER CREDIT];

0.4) NOUVEAU — Q_Flux_Residualized (Flux “netté” de l’archive)
	Une ligne par nature en cours, avec démarrage résiduel et durée résiduelle (jamais négative).
	Parenthésage Access strict.

/* ===================== VG ===================== */
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'VG' AS Nature,
    IIf((av.MaxVgEnd Is Null) OR (f.VG_StartDate > av.MaxVgEnd),
        f.VG_StartDate,
        DateAdd('d',1,av.MaxVgEnd)) AS StartDate,
    NULL AS EndDate,
    IIf((av.MaxVgEnd Is Null) OR (av.MaxVgEnd < f.VG_StartDate),
        f.VG_NbJoursRetard,
        IIf( (f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1)) < 0,
             0,
             f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1)
        )
    ) AS Duree,
    f.VG_WorstInsuffisance AS WorstInsuffisance,
    f.VG_DateWorstInsuffisance AS WorstDate,
    f.VG_LastInsuffisance AS LastInsuffisance,
    f.VG_LastUpdateDate AS LastUpdateDate,
    f.[Montant VG] AS Montant,
    f.DepassementNA,
    f.[Production Date],
    f.Limite,
    f.Consommation,
    f.[NO INTERVENANT],
    f.[NO INTERVENANT GRP],
    'FLUX' AS Source
FROM
(
  (
    T_Flux AS f
    LEFT JOIN
    ( SELECT Booking, [NO DOSSIER CREDIT], MAX(VG_EndDate) AS MaxVgEnd
      FROM T_Archive
      WHERE FlagVG='YES'
      GROUP BY Booking, [NO DOSSIER CREDIT]
    ) AS av
    ON (f.[NO DOSSIER CREDIT]=av.[NO DOSSIER CREDIT]) AND (f.Booking=av.Booking)
  )
)
WHERE (f.VG_StartDate IS NOT NULL)
  AND (f.VG_LastInsuffisance < 0)
  AND (
       IIf((av.MaxVgEnd Is Null) OR (av.MaxVgEnd < f.VG_StartDate),
           f.VG_NbJoursRetard,
           IIf( (f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1)) < 0,
                0,
                f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1)
           )
       )
      ) > 0

UNION ALL

/* ===================== AM ===================== */
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'AM' AS Nature,
    IIf((aa.MaxAmEnd Is Null) OR (f.AM_StartDate > aa.MaxAmEnd),
        f.AM_StartDate,
        DateAdd('d',1,aa.MaxAmEnd)) AS StartDate,
    NULL AS EndDate,
    IIf((aa.MaxAmEnd Is Null) OR (aa.MaxAmEnd < f.AM_StartDate),
        f.AM_NbJoursRetard,
        IIf( (f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1)) < 0,
             0,
             f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1)
        )
    ) AS Duree,
    f.AM_WorstInsuffisance AS WorstInsuffisance,
    f.AM_DateWorstInsuffisance AS WorstDate,
    f.AM_LastInsuffisance AS LastInsuffisance,
    f.AM_LastUpdateDate AS LastUpdateDate,
    f.[Montant AM] AS Montant,
    f.DepassementNA,
    f.[Production Date],
    f.Limite,
    f.Consommation,
    f.[NO INTERVENANT],
    f.[NO INTERVENANT GRP],
    'FLUX' AS Source
FROM
(
  (
    T_Flux AS f
    LEFT JOIN
    ( SELECT Booking, [NO DOSSIER CREDIT], MAX(AM_EndDate) AS MaxAmEnd
      FROM T_Archive
      WHERE FlagAM='YES'
      GROUP BY Booking, [NO DOSSIER CREDIT]
    ) AS aa
    ON (f.[NO DOSSIER CREDIT]=aa.[NO DOSSIER CREDIT]) AND (f.Booking=aa.Booking)
  )
)
WHERE (f.AM_StartDate IS NOT NULL)
  AND (f.AM_LastInsuffisance < 0)
  AND (
       IIf((aa.MaxAmEnd Is Null) OR (aa.MaxAmEnd < f.AM_StartDate),
           f.AM_NbJoursRetard,
           IIf( (f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1)) < 0,
                0,
                f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1)
           )
       )
      ) > 0

UNION ALL

/* ===================== SL ===================== */
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'SL' AS Nature,
    IIf((asl.MaxSlEnd Is Null) OR (f.SL_StartDate > asl.MaxSlEnd),
        f.SL_StartDate,
        DateAdd('d',1,asl.MaxSlEnd)) AS StartDate,
    NULL AS EndDate,
    IIf((asl.MaxSlEnd Is Null) OR (asl.MaxSlEnd < f.SL_StartDate),
        f.SL_NbJoursRetard,
        IIf( (f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1)) < 0,
             0,
             f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1)
        )
    ) AS Duree,
    f.SL_WorstInsuffisance AS WorstInsuffisance,
    f.SL_DateWorstInsuffisance AS WorstDate,
    f.SL_LastInsuffisance AS LastInsuffisance,
    f.SL_LastUpdateDate AS LastUpdateDate,
    f.[Montant SL] AS Montant,
    f.DepassementNA,
    f.[Production Date],
    f.Limite,
    f.Consommation,
    f.[NO INTERVENANT],
    f.[NO INTERVENANT GRP],
    'FLUX' AS Source
FROM
(
  (
    T_Flux AS f
    LEFT JOIN
    ( SELECT Booking, [NO DOSSIER CREDIT], MAX(SL_EndDate) AS MaxSlEnd
      FROM T_Archive
      WHERE FlagSL='YES'
      GROUP BY Booking, [NO DOSSIER CREDIT]
    ) AS asl
    ON (f.[NO DOSSIER CREDIT]=asl.[NO DOSSIER CREDIT]) AND (f.Booking=asl.Booking)
  )
)
WHERE (f.SL_StartDate IS NOT NULL)
  AND (f.SL_LastInsuffisance < 0)
  AND (
       IIf((asl.MaxSlEnd Is Null) OR (asl.MaxSlEnd < f.SL_StartDate),
           f.SL_NbJoursRetard,
           IIf( (f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1)) < 0,
                0,
                f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1)
           )
       )
      ) > 0;

0.5) NOUVEAU — Q_Episodes_All (Archive + Flux résidualisé)
	Vue unifiée pour toutes les requêtes aval.

SELECT
    a.Booking, a.[NO DOSSIER CREDIT], a.Nature,
    a.StartDate, a.EndDate, a.Duree,
    a.WorstInsuffisance, a.WorstDate,
    a.LastInsuffisance, a.LastUpdateDate,
    a.Montant, a.DepassementNA,
    a.[Production Date], a.Limite, a.Consommation,
    a.[NO INTERVENANT], a.[NO INTERVENANT GRP],
    'ARCHIVE' AS Source
FROM Q_Arc_UnpivotNature_W AS a

UNION ALL

SELECT
    f.Booking, f.[NO DOSSIER CREDIT], f.Nature,
    f.StartDate, f.EndDate, f.Duree,
    f.WorstInsuffisance, f.WorstDate,
    f.LastInsuffisance, f.LastUpdateDate,
    f.Montant, f.DepassementNA,
    f.[Production Date], f.Limite, f.Consommation,
    f.[NO INTERVENANT], f.[NO INTERVENANT GRP],
    'FLUX' AS Source
FROM Q_Flux_Residualized AS f;

0.6) (option) Q_Episodes_All_Today — restreint aux clés du jour

SELECT e.*
FROM
(
  (SELECT DISTINCT Booking, [NO DOSSIER CREDIT] FROM T_Today) AS k
  INNER JOIN Q_Episodes_All AS e
    ON (k.[NO DOSSIER CREDIT]=e.[NO DOSSIER CREDIT]) AND (k.Booking=e.Booking)
);



1) Intervenants en retard — refactor vers Q_Episodes_All
	On ne fait plus d’UNION manuel : on filtre Source.
	Rappel : pour la période, seules les lignes ARCHIVE sont bornées par EndDate ; les FLUX restent ouvertes.
1.1) VG — Intervenants + infos VG (Archive + Flux, filtre Booking)

SELECT Source, 'VG' AS Nature,
       Booking, [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP],
       [Production Date], StartDate, EndDate, Duree,
       WorstInsuffisance, LastInsuffisance,
       Limite, Consommation, Montant, DepassementNA
FROM Q_Episodes_All
WHERE Nature='VG'
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
  AND ( [BookingLike] IS NULL OR Booking LIKE [BookingLike] );

1.2) AM — Intervenants + infos AM (Archive + Flux)

SELECT Source, 'AM' AS Nature,
       Booking, [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP],
       [Production Date], StartDate, EndDate, Duree,
       WorstInsuffisance, LastInsuffisance,
       Limite, Consommation, Montant, DepassementNA
FROM Q_Episodes_All
WHERE Nature='AM'
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
  AND ( [BookingLike] IS NULL OR Booking LIKE [BookingLike] );

1.3) SL — Intervenants + infos SL (Archive + Flux)

SELECT Source, 'SL' AS Nature,
       Booking, [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP],
       [Production Date], StartDate, EndDate, Duree,
       WorstInsuffisance, LastInsuffisance,
       Limite, Consommation, Montant, DepassementNA
FROM Q_Episodes_All
WHERE Nature='SL'
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
  AND ( [BookingLike] IS NULL OR Booking LIKE [BookingLike] );

1.4) ALL — Intervenants en retard (toutes natures)

SELECT Source, Nature,
       Booking, [NO DOSSIER CREDIT], [NO INTERVENANT], [NO INTERVENANT GRP],
       [Production Date], StartDate, EndDate, Duree,
       WorstInsuffisance, LastInsuffisance,
       Limite, Consommation, Montant, DepassementNA
FROM Q_Episodes_All
WHERE ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) );



2) TOP Intervenants
2.1) TOP 30 VG (Archive + Flux)

SELECT TOP 30 [NO INTERVENANT] AS Intervenant, COUNT(*) AS NbEpisodes
FROM Q_Episodes_All
WHERE Nature='VG'
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
GROUP BY [NO INTERVENANT]
ORDER BY COUNT(*) DESC;

2.2) TOP 30 AM (Archive + Flux)

SELECT TOP 30 [NO INTERVENANT] AS Intervenant, COUNT(*) AS NbEpisodes
FROM Q_Episodes_All
WHERE Nature='AM'
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
GROUP BY [NO INTERVENANT]
ORDER BY COUNT(*) DESC;

2.3) TOP 30 SL (Archive + Flux)

SELECT TOP 30 [NO INTERVENANT] AS Intervenant, COUNT(*) AS NbEpisodes
FROM Q_Episodes_All
WHERE Nature='SL'
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
GROUP BY [NO INTERVENANT]
ORDER BY COUNT(*) DESC;

2.4) TOP 30 AM & SL combinés (Archive + Flux)

SELECT TOP 30 [NO INTERVENANT] AS Intervenant, COUNT(*) AS NbEpisodes
FROM Q_Episodes_All
WHERE Nature IN ('AM','SL')
  AND ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
GROUP BY [NO INTERVENANT]
ORDER BY COUNT(*) DESC;



2 bis) TOP 30 par jours cumulés (Archive + Flux résidualisé)
	Remplace l’ancienne version (plus besoin de sommer “à la main” archive + NbJ flux).

SELECT TOP 30 [NO INTERVENANT] AS Intervenant,
       SUM(Duree) AS JoursCumules
FROM Q_Episodes_All
WHERE ( Source='FLUX'
        OR (Source='ARCHIVE' AND EndDate BETWEEN [StartDate] AND [EndDate]) )
GROUP BY [NO INTERVENANT]
ORDER BY SUM(Duree) DESC;



Remarques
	• Toutes tes requêtes Archive-only (clôtures, SLA, stats de durée à la clôture, etc.) peuvent rester sur Q_Arc_UnpivotNature ou Q_Arc_UnpivotNature_W (elles ne sont pas concernées par le double-comptage).
	• Dès que tu veux Archive + Flux, pars désormais de Q_Episodes_All (qui garantit la non-duplication intra-nature grâce au résiduel).
	• Si tu veux limiter l’univers aux clés présentes dans le fichier du jour, utilise Q_Episodes_All_Today.
Dis-moi si tu veux aussi que je refasse les autres blocs (KPI, crosstabs, transitions, etc.) pour les faire pointer sur Q_Episodes_All quand c’est pertinent — je te les bascule vite fait.

# Comm

A) Unpivot Archive avec commentaires
A.1) Q_Arc_UnpivotNature_WC

(= Archive déroulée par nature + WorstDate + commentaires)

SELECT
    ID AS ArcID, Booking, [NO DOSSIER CREDIT],
    'VG' AS Nature,
    VG_StartDate AS StartDate, VG_EndDate AS EndDate, VG_Duree AS Duree,
    VG_WorstInsuffisance AS WorstInsuffisance, VG_DateWorstInsuffisance AS WorstDate,
    VG_LastInsuffisance AS LastInsuffisance, VG_LastUpdateDate AS LastUpdateDate,
    [Montant VG] AS Montant, FlagVG AS Flag,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP],
    risk_cmt, front_cmt, close_cmt
FROM T_Archive
WHERE FlagVG='YES'
UNION ALL
SELECT
    ID, Booking, [NO DOSSIER CREDIT],
    'AM',
    AM_StartDate, AM_EndDate, AM_Duree,
    AM_WorstInsuffisance, AM_DateWorstInsuffisance,
    AM_LastInsuffisance, AM_LastUpdateDate,
    [Montant AM], FlagAM,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP],
    risk_cmt, front_cmt, close_cmt
FROM T_Archive
WHERE FlagAM='YES'
UNION ALL
SELECT
    ID, Booking, [NO DOSSIER CREDIT],
    'SL',
    SL_StartDate, SL_EndDate, SL_Duree,
    SL_WorstInsuffisance, SL_DateWorstInsuffisance,
    SL_LastInsuffisance, SL_LastUpdateDate,
    [Montant SL], FlagSL,
    [Production Date], Limite, Consommation, DepassementNA,
    [NO INTERVENANT], [NO INTERVENANT GRP],
    risk_cmt, front_cmt, close_cmt
FROM T_Archive
WHERE FlagSL='YES';

B) Derniers commentaires archivés par clé
B.1) Q_Arc_LastCommentsPerKey

(= pour chaque (Booking, NDC), on prend la ligne unpivot à EndDate max et on récupère risk_cmt, front_cmt, close_cmt)

SELECT u.Booking,
       u.[NO DOSSIER CREDIT] AS NDC,
       FIRST(u.risk_cmt)  AS risk_cmt_arc,
       FIRST(u.front_cmt) AS front_cmt_arc,
       FIRST(u.close_cmt) AS close_cmt_arc
FROM Q_Arc_UnpivotNature_WC AS u
WHERE u.EndDate = (
    SELECT MAX(z.EndDate)
    FROM Q_Arc_UnpivotNature_WC AS z
    WHERE z.Booking=u.Booking
      AND z.[NO DOSSIER CREDIT]=u.[NO DOSSIER CREDIT]
)
GROUP BY u.Booking, u.[NO DOSSIER CREDIT];

C) Commentaires “courants” (Flux prioritaire, sinon dernier Archive)
C.1) Q_Comments_Current
SELECT
    k.Booking,
    k.[NO DOSSIER CREDIT] AS NDC,
    IIf(Len(Trim(Nz(f.risk_cmt,'')))>0,  f.risk_cmt,  lc.risk_cmt_arc)  AS risk_cmt_current,
    IIf(Len(Trim(Nz(f.front_cmt,'')))>0, f.front_cmt, lc.front_cmt_arc) AS front_cmt_current,
    lc.close_cmt_arc AS close_cmt_last
FROM
    (
      SELECT DISTINCT Booking, [NO DOSSIER CREDIT]
      FROM T_Flux
      UNION
      SELECT DISTINCT Booking, [NO DOSSIER CREDIT]
      FROM T_Archive
    ) AS k
LEFT JOIN T_Flux AS f
  ON (k.Booking=f.Booking) AND (k.[NO DOSSIER CREDIT]=f.[NO DOSSIER CREDIT])
LEFT JOIN Q_Arc_LastCommentsPerKey AS lc
  ON (k.Booking=lc.Booking) AND (k.[NO DOSSIER CREDIT]=lc.NDC);


risk_cmt_current et front_cmt_current : Flux si présent, sinon dernier Archive.
close_cmt_last : dernier commentaire de clôture (Archive), utile pour rappel d’historique.

D) Episodes à date = Archive UNION Flux résidualisé (avec commentaires)

C’est la source unique à utiliser pour tous les cumuls jours/montants sans double-compte.

D.1) Q_Episodes_All

(= Archive unpivot + Flux résiduel VG/AM/SL, avec risk_cmt, front_cmt, close_cmt)

/* ===== ARCHIVE (déjà clos) ===== */
SELECT
    a.Booking,
    a.[NO DOSSIER CREDIT]       AS NDC,
    a.Nature,
    a.StartDate,
    a.EndDate,
    a.Duree,
    a.WorstInsuffisance,
    a.WorstDate,
    a.LastInsuffisance,
    a.LastUpdateDate,
    a.DepassementNA,
    a.Montant,
    a.[NO INTERVENANT],
    a.[NO INTERVENANT GRP],
    a.risk_cmt,
    a.front_cmt,
    a.close_cmt,
    'ARCHIVE'                   AS Source
FROM Q_Arc_UnpivotNature_WC AS a

UNION ALL

/* ===== FLUX résidualisé — VG ===== */
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT]                                   AS NDC,
    'VG'                                                     AS Nature,
    IIf(av.MaxVgEnd Is Null Or f.VG_StartDate > av.MaxVgEnd,
        f.VG_StartDate,
        DateAdd('d',1,av.MaxVgEnd))                         AS StartDate,
    NULL                                                     AS EndDate,
    IIf(av.MaxVgEnd Is Null Or av.MaxVgEnd < f.VG_StartDate,
        f.VG_NbJoursRetard,
        IIf(f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1) < 0,
            0,
            f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1)
        )
    )                                                        AS Duree,
    f.VG_WorstInsuffisance                                   AS WorstInsuffisance,
    f.VG_DateWorstInsuffisance                               AS WorstDate,
    f.VG_LastInsuffisance                                    AS LastInsuffisance,
    f.VG_LastUpdateDate                                      AS LastUpdateDate,
    f.DepassementNA,
    f.[Montant VG]                                           AS Montant,
    f.[NO INTERVENANT],
    f.[NO INTERVENANT GRP],
    f.risk_cmt,
    f.front_cmt,
    NULL                                                     AS close_cmt,
    'FLUX'                                                   AS Source
FROM
(
  (
    ( T_Flux AS f )
    LEFT JOIN
    (
      SELECT Booking, [NO DOSSIER CREDIT], MAX(VG_EndDate) AS MaxVgEnd
      FROM T_Archive
      WHERE FlagVG='YES'
      GROUP BY Booking, [NO DOSSIER CREDIT]
    ) AS av
    ON (f.Booking=av.Booking) AND (f.[NO DOSSIER CREDIT]=av.[NO DOSSIER CREDIT])
  )
)
WHERE f.VG_StartDate IS NOT NULL
  AND f.VG_LastInsuffisance < 0

UNION ALL

/* ===== FLUX résidualisé — AM ===== */
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'AM',
    IIf(aa.MaxAmEnd Is Null Or f.AM_StartDate > aa.MaxAmEnd,
        f.AM_StartDate,
        DateAdd('d',1,aa.MaxAmEnd)),
    NULL,
    IIf(aa.MaxAmEnd Is Null Or aa.MaxAmEnd < f.AM_StartDate,
        f.AM_NbJoursRetard,
        IIf(f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1) < 0,
            0,
            f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1)
        )
    ),
    f.AM_WorstInsuffisance,
    f.AM_DateWorstInsuffisance,
    f.AM_LastInsuffisance,
    f.AM_LastUpdateDate,
    f.DepassementNA,
    f.[Montant AM],
    f.[NO INTERVENANT],
    f.[NO INTERVENANT GRP],
    f.risk_cmt,
    f.front_cmt,
    NULL,
    'FLUX'
FROM
(
  (
    ( T_Flux AS f )
    LEFT JOIN
    (
      SELECT Booking, [NO DOSSIER CREDIT], MAX(AM_EndDate) AS MaxAmEnd
      FROM T_Archive
      WHERE FlagAM='YES'
      GROUP BY Booking, [NO DOSSIER CREDIT]
    ) AS aa
    ON (f.Booking=aa.Booking) AND (f.[NO DOSSIER CREDIT]=aa.[NO DOSSIER CREDIT])
  )
)
WHERE f.AM_StartDate IS NOT NULL
  AND f.AM_LastInsuffisance < 0

UNION ALL

/* ===== FLUX résidualisé — SL ===== */
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'SL',
    IIf(asl.MaxSlEnd Is Null Or f.SL_StartDate > asl.MaxSlEnd,
        f.SL_StartDate,
        DateAdd('d',1,asl.MaxSlEnd)),
    NULL,
    IIf(asl.MaxSlEnd Is Null Or asl.MaxSlEnd < f.SL_StartDate,
        f.SL_NbJoursRetard,
        IIf(f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1) < 0,
            0,
            f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1)
        )
    ),
    f.SL_WorstInsuffisance,
    f.SL_DateWorstInsuffisance,
    f.SL_LastInsuffisance,
    f.SL_LastUpdateDate,
    f.DepassementNA,
    f.[Montant SL],
    f.[NO INTERVENANT],
    f.[NO INTERVENANT GRP],
    f.risk_cmt,
    f.front_cmt,
    NULL,
    'FLUX'
FROM
(
  (
    ( T_Flux AS f )
    LEFT JOIN
    (
      SELECT Booking, [NO DOSSIER CREDIT], MAX(SL_EndDate) AS MaxSlEnd
      FROM T_Archive
      WHERE FlagSL='YES'
      GROUP BY Booking, [NO DOSSIER CREDIT]
    ) AS asl
    ON (f.Booking=asl.Booking) AND (f.[NO DOSSIER CREDIT]=asl.[NO DOSSIER CREDIT])
  )
)
WHERE f.SL_StartDate IS NOT NULL
  AND f.SL_LastInsuffisance < 0

ORDER BY
    Booking, NDC, Nature,
    IIf(StartDate Is Null, #1/1/1900#, StartDate),
    IIf(EndDate   Is Null, #12/31/9999#, EndDate);

E) Episodes + commentaires “résolus” (Flux prioritaire)

Si tu veux afficher les commentaires “courants” même pour les lignes d’Archive, fais simplement :

E.1) Q_Episodes_All_WithResolvedComments
SELECT
    e.*,
    IIf(Len(Trim(Nz(e.risk_cmt,'')))>0,  e.risk_cmt,  c.risk_cmt_current)  AS risk_cmt_resolved,
    IIf(Len(Trim(Nz(e.front_cmt,'')))>0, e.front_cmt, c.front_cmt_current) AS front_cmt_resolved,
    IIf(e.Source='ARCHIVE', Nz(e.close_cmt, c.close_cmt_last), Null)        AS close_cmt_resolved
FROM
  ( Q_Episodes_All AS e
    LEFT JOIN Q_Comments_Current AS c
      ON (e.Booking=c.Booking) AND (e.NDC=c.NDC)
  )