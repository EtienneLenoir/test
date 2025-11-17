1) Détail par épisode (Archive + Flux)
SELECT
    a.Booking,
    a.[NO DOSSIER CREDIT]       AS NDC,
    'VG'                         AS Nature,
    a.VG_StartDate               AS StartDate,
    a.VG_EndDate                 AS EndDate,
    a.VG_Duree                   AS Duree,
    a.VG_WorstInsuffisance       AS WorstInsuffisance,
    a.VG_DateWorstInsuffisance   AS DateWorstInsuffisance,
    a.VG_LastInsuffisance        AS LastInsuffisance,
    a.VG_LastUpdateDate          AS LastUpdateDate,
    a.DepassementNA              AS DepassementNA,
    'ARCHIVE'                    AS Source
FROM T_Archive AS a
WHERE a.FlagVG='YES'

UNION ALL
SELECT
    a.Booking,
    a.[NO DOSSIER CREDIT],
    'AM'                         AS Nature,
    a.AM_StartDate,
    a.AM_EndDate,
    a.AM_Duree,
    a.AM_WorstInsuffisance,
    a.AM_DateWorstInsuffisance,
    a.AM_LastInsuffisance,
    a.AM_LastUpdateDate,
    a.DepassementNA,
    'ARCHIVE'                    AS Source
FROM T_Archive AS a
WHERE a.FlagAM='YES'

UNION ALL
SELECT
    a.Booking,
    a.[NO DOSSIER CREDIT],
    'SL'                         AS Nature,
    a.SL_StartDate,
    a.SL_EndDate,
    a.SL_Duree,
    a.SL_WorstInsuffisance,
    a.SL_DateWorstInsuffisance,
    a.SL_LastInsuffisance,
    a.SL_LastUpdateDate,
    a.DepassementNA,
    'ARCHIVE'                    AS Source
FROM T_Archive AS a
WHERE a.FlagSL='YES'

UNION ALL
-- FLUX : on ne sort SL que si SL est EN COURS
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'SL'                         AS Nature,
    f.SL_StartDate               AS StartDate,
    NULL                         AS EndDate,
    f.SL_NbJoursRetard           AS Duree,
    f.SL_WorstInsuffisance       AS WorstInsuffisance,
    f.SL_DateWorstInsuffisance   AS DateWorstInsuffisance,
    f.SL_LastInsuffisance        AS LastInsuffisance,
    f.SL_LastUpdateDate          AS LastUpdateDate,
    f.DepassementNA,
    'FLUX'                       AS Source
FROM T_Flux AS f
WHERE f.SL_StartDate IS NOT NULL AND f.SL_LastInsuffisance < 0

UNION ALL
-- FLUX : AM en cours UNIQUEMENT s'il n'y a PAS SL en cours
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'AM',
    f.AM_StartDate,
    NULL,
    f.AM_NbJoursRetard,
    f.AM_WorstInsuffisance,
    f.AM_DateWorstInsuffisance,
    f.AM_LastInsuffisance,
    f.AM_LastUpdateDate,
    f.DepassementNA,
    'FLUX'
FROM T_Flux AS f
WHERE f.AM_StartDate IS NOT NULL
  AND f.AM_LastInsuffisance < 0
  AND NOT (f.SL_StartDate IS NOT NULL AND f.SL_LastInsuffisance < 0)

UNION ALL
-- FLUX : VG en cours UNIQUEMENT s'il n'y a NI SL NI AM en cours
SELECT
    f.Booking,
    f.[NO DOSSIER CREDIT],
    'VG',
    f.VG_StartDate,
    NULL,
    f.VG_NbJoursRetard,
    f.VG_WorstInsuffisance,
    f.VG_DateWorstInsuffisance,
    f.VG_LastInsuffisance,
    f.VG_LastUpdateDate,
    f.DepassementNA,
    'FLUX'
FROM T_Flux AS f
WHERE f.VG_StartDate IS NOT NULL
  AND f.VG_LastInsuffisance < 0
  AND NOT (
        (f.SL_StartDate IS NOT NULL AND f.SL_LastInsuffisance < 0)
     OR (f.AM_StartDate IS NOT NULL AND f.AM_LastInsuffisance < 0)
  )
ORDER BY
    Booking, [NO DOSSIER CREDIT], Nature,
    IIF(StartDate IS NULL, #1/1/1900#, StartDate),
    IIF(EndDate IS NULL, #12/31/9999#, EndDate);


Archive : 1 ligne par épisode clos et par nature (VG/AM/SL) → plusieurs lignes possibles pour le même (Booking, NDC), c’est normal.

Flux : au plus 1 ligne par (Booking, NDC) (la plus haute nature en cours : SL > AM > VG) → évite les doublons quand SL est en cours (et entraîne forcément AM & VG).

Durée en archive = *_Duree. En flux = *_NbJoursRetard.

Worst-of = *_WorstInsuffisance + *_DateWorstInsuffisance déjà calculés par ton moteur.

Variante “périmètre du jour” : si tu veux limiter aux bookings importés aujourd’hui, ajoute à chaque SELECT AND ... Booking IN (SELECT DISTINCT Booking FROM T_Today).



2) Vue “cumul” par clé et nature (Archive + Flux en cours)

Si tu veux un total de jours par (Booking, NDC, Nature) sans double-compte (hiérarchie respectée côté flux) :

SELECT
    Z.Booking,
    Z.[NO DOSSIER CREDIT] AS NDC,
    Z.Nature,
    SUM(Z.Duree)          AS CumulJours,
    MIN(Z.StartDate)      AS FirstStart,
    MAX(Z.EndDate)        AS LastEnd
FROM (
    /* Reprend exactement la requête détaillée ci-dessus, SANS le ORDER BY */
    /* ... colle ici le bloc UNION ALL précédent et remplace simplement le ORDER BY par rien ... */
) AS Z
GROUP BY Z.Booking, Z.[NO DOSSIER CREDIT], Z.Nature
ORDER BY Z.Booking, Z.[NO DOSSIER CREDIT], Z.Nature;


Cette agrégation additionne :
somme(archives) + jours en cours (flux) de la nature la plus haute uniquement.

Tu peux ainsi afficher “Cumul jours VG/AM/SL” sans surcompter quand SL est en cours.

Notes rapides

Index utiles (tu les as déjà) :
T_Flux(Booking,[NO DOSSIER CREDIT]), T_Archive(Booking,[NO DOSSIER CREDIT]), plus *_StartDate/*_EndDate si tu filtres/date-ordres.

Si tu veux voir toutes les natures en cours (et pas la hiérarchie), supprime les conditions NOT (...) dans les 2 SELECT FLUX (AM & VG).

Si tu veux que je te prépare une vue enregistrée (ou deux : Episodes_Detail et Durations_Summary) à coller tel quel dans Access, dis-moi si tu veux la version “périmètre du jour” ou “globale”.

# top

Yes, je vois exactement ce que tu veux éviter : ne pas recompter des jours “en cours (flux)” qui ont déjà été historisés dans l’archive pour la même nature (VG/AM/SL) et la même clé (Booking + N° crédit) — notamment quand l’épisode en cours a le même début que l’épisode archivé.
👉 La bonne approche est de résidualiser le flux par nature : on soustrait la partie qui chevauche déjà l’archive (jusqu’à la dernière fin archivée), et on ne garde que les jours supplémentaires.

Ci-dessous, une requête générale en 2 parties :

- Détail des épisodes = toutes les lignes d’archive (épisodes clos) + les épisodes en cours du flux résidualisés par nature.

- Cumul par nature = somme (Archive + Flux résiduel) par (Booking, NDC, Nature).

⚠️ Je ne filtre pas AM ou VG quand SL est en cours (tu verras donc bien VG + AM + SL si présents). Si tu veux “écraser” par hiérarchie (SL > AM > VG) on peut ajouter un filtre, mais pour ton besoin (ne pas doubler et bien sommer) ce n’est pas nécessaire

```sql
SELECT *
FROM
(
    /* ========== ARCHIVE (3 natures) ========== */
    SELECT
        a.Booking,
        a.[NO DOSSIER CREDIT]           AS NDC,
        'VG'                            AS Nature,
        a.VG_StartDate                  AS StartDate,
        a.VG_EndDate                    AS EndDate,
        a.VG_Duree                      AS Duree,
        a.VG_WorstInsuffisance          AS WorstInsuffisance,
        a.VG_DateWorstInsuffisance      AS DateWorstInsuffisance,
        a.VG_LastInsuffisance           AS LastInsuffisance,
        a.VG_LastUpdateDate             AS LastUpdateDate,
        a.DepassementNA,
        'ARCHIVE'                       AS Source
    FROM T_Archive AS a
    WHERE a.FlagVG='YES'

    UNION ALL
    SELECT
        a.Booking, a.[NO DOSSIER CREDIT], 'AM',
        a.AM_StartDate, a.AM_EndDate, a.AM_Duree,
        a.AM_WorstInsuffisance, a.AM_DateWorstInsuffisance,
        a.AM_LastInsuffisance, a.AM_LastUpdateDate,
        a.DepassementNA, 'ARCHIVE'
    FROM T_Archive AS a
    WHERE a.FlagAM='YES'

    UNION ALL
    SELECT
        a.Booking, a.[NO DOSSIER CREDIT], 'SL',
        a.SL_StartDate, a.SL_EndDate, a.SL_Duree,
        a.SL_WorstInsuffisance, a.SL_DateWorstInsuffisance,
        a.SL_LastInsuffisance, a.SL_LastUpdateDate,
        a.DepassementNA, 'ARCHIVE'
    FROM T_Archive AS a
    WHERE a.FlagSL='YES'

    UNION ALL
    /* ========== FLUX RÉSIDUALISÉ — VG ========== */
    SELECT
        f.Booking,
        f.[NO DOSSIER CREDIT]                           AS NDC,
        'VG'                                            AS Nature,
        IIf(av.MaxVgEnd Is Null Or f.VG_StartDate > av.MaxVgEnd,
            f.VG_StartDate,
            DateAdd('d', 1, av.MaxVgEnd))              AS StartDate,
        NULL                                            AS EndDate,
        IIf(av.MaxVgEnd Is Null Or av.MaxVgEnd < f.VG_StartDate,
            f.VG_NbJoursRetard,
            IIf(f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1) < 0,
                0,
                f.VG_NbJoursRetard - (DateDiff('d', f.VG_StartDate, av.MaxVgEnd) + 1)
            )
        )                                               AS Duree,
        f.VG_WorstInsuffisance                          AS WorstInsuffisance,
        f.VG_DateWorstInsuffisance                      AS DateWorstInsuffisance,
        f.VG_LastInsuffisance                           AS LastInsuffisance,
        f.VG_LastUpdateDate                             AS LastUpdateDate,
        f.DepassementNA,
        'FLUX'                                          AS Source
    FROM
    (
        ( T_Flux AS f
          LEFT JOIN
          ( SELECT Booking, [NO DOSSIER CREDIT], MAX(VG_EndDate) AS MaxVgEnd
            FROM T_Archive
            WHERE FlagVG='YES'
            GROUP BY Booking, [NO DOSSIER CREDIT]
          ) AS av
          ON (f.Booking = av.Booking) AND (f.[NO DOSSIER CREDIT] = av.[NO DOSSIER CREDIT])
        )
    )
    WHERE f.VG_StartDate IS NOT NULL
      AND f.VG_LastInsuffisance < 0

    UNION ALL
    /* ========== FLUX RÉSIDUALISÉ — AM ========== */
    SELECT
        f.Booking,
        f.[NO DOSSIER CREDIT],
        'AM',
        IIf(aa.MaxAmEnd Is Null Or f.AM_StartDate > aa.MaxAmEnd,
            f.AM_StartDate,
            DateAdd('d', 1, aa.MaxAmEnd))              AS StartDate,
        NULL                                            AS EndDate,
        IIf(aa.MaxAmEnd Is Null Or aa.MaxAmEnd < f.AM_StartDate,
            f.AM_NbJoursRetard,
            IIf(f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1) < 0,
                0,
                f.AM_NbJoursRetard - (DateDiff('d', f.AM_StartDate, aa.MaxAmEnd) + 1)
            )
        )                                               AS Duree,
        f.AM_WorstInsuffisance,
        f.AM_DateWorstInsuffisance,
        f.AM_LastInsuffisance,
        f.AM_LastUpdateDate,
        f.DepassementNA,
        'FLUX'
    FROM
    (
        ( T_Flux AS f
          LEFT JOIN
          ( SELECT Booking, [NO DOSSIER CREDIT], MAX(AM_EndDate) AS MaxAmEnd
            FROM T_Archive
            WHERE FlagAM='YES'
            GROUP BY Booking, [NO DOSSIER CREDIT]
          ) AS aa
          ON (f.Booking = aa.Booking) AND (f.[NO DOSSIER CREDIT] = aa.[NO DOSSIER CREDIT])
        )
    )
    WHERE f.AM_StartDate IS NOT NULL
      AND f.AM_LastInsuffisance < 0

    UNION ALL
    /* ========== FLUX RÉSIDUALISÉ — SL ========== */
    SELECT
        f.Booking,
        f.[NO DOSSIER CREDIT],
        'SL',
        IIf(asl.MaxSlEnd Is Null Or f.SL_StartDate > asl.MaxSlEnd,
            f.SL_StartDate,
            DateAdd('d', 1, asl.MaxSlEnd))             AS StartDate,
        NULL                                            AS EndDate,
        IIf(asl.MaxSlEnd Is Null Or asl.MaxSlEnd < f.SL_StartDate,
            f.SL_NbJoursRetard,
            IIf(f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1) < 0,
                0,
                f.SL_NbJoursRetard - (DateDiff('d', f.SL_StartDate, asl.MaxSlEnd) + 1)
            )
        )                                               AS Duree,
        f.SL_WorstInsuffisance,
        f.SL_DateWorstInsuffisance,
        f.SL_LastInsuffisance,
        f.SL_LastUpdateDate,
        f.DepassementNA,
        'FLUX'
    FROM
    (
        ( T_Flux AS f
          LEFT JOIN
          ( SELECT Booking, [NO DOSSIER CREDIT], MAX(SL_EndDate) AS MaxSlEnd
            FROM T_Archive
            WHERE FlagSL='YES'
            GROUP BY Booking, [NO DOSSIER CREDIT]
          ) AS asl
          ON (f.Booking = asl.Booking) AND (f.[NO DOSSIER CREDIT] = asl.[NO DOSSIER CREDIT])
        )
    )
    WHERE f.SL_StartDate IS NOT NULL
      AND f.SL_LastInsuffisance < 0
)  AS Z
WHERE Z.Duree > 0
ORDER BY
    Z.Booking, Z.NDC, Z.Nature,
    IIf(IsNull(Z.StartDate), #1/1/1900#, Z.StartDate),
    IIf(IsNull(Z.EndDate),   #12/31/9999#, Z.EndDate);


```

Ce que fait la partie “résidualisée” (les 3 blocs FLUX) :
	• récupère, pour chaque nature, la dernière date de fin archivée (Max*_End).
	• Décale le début affiché sur flux à MaxEnd + 1 jour si l’archive couvre déjà ce segment.
	• Réduit la durée affichée sur flux de la partie qui chevauche l’archive (formule inclusive : DateDiff('d', Start, End) + 1).
	• La durée ne peut pas devenir négative (on la coupe à 0).
	• Résultat : pas de double-compte si même début et, plus généralement, si l’archive couvre déjà une partie du flux.
Astuce “périmètre du jour” : si tu veux limiter aux clés importées aujourd’hui, ajoute à chaque SELECT un AND ... IN (SELECT DISTINCT Booking FROM T_Today) ou un JOIN avec T_Today.


2) Cumul par nature (Archive + Flux résiduel)

Enregistre la requête ci-dessus comme Q_Episodes_All.
Puis agrège

```
SELECT
    Z.Booking,
    Z.NDC,
    Z.Nature,
    SUM(Z.Duree)                         AS CumulJours,
    MIN(Z.StartDate)                     AS FirstStart,
    MAX(Z.EndDate)                       AS LastEnd
FROM Q_Episodes_All AS Z
WHERE Z.Duree > 0
GROUP BY Z.Booking, Z.NDC, Z.Nature
ORDER BY Z.Booking, Z.NDC, Z.Nature;

```

# Précision

Oui, exactement — mais précisons la règle pour qu’il n’y ait aucun double-comptage :

Règle par nature (VG, AM, SL), pour une clé (Booking + N° crédit)

Total archive = somme des Duree des épisodes clos en archive pour cette nature.

Flux résiduel = jours en cours non déjà archivés :

Si pas d’épisode ouvert → 0.

Sinon, prends NbJoursRetard du flux moins le chevauchement avec le dernier épisode archivé :
overlap = DateDiff('d', StartFlux, MaxEndArchive) + 1 (0 si MaxEndArchive < StartFlux).
residuel = Max(0, NbJoursRetard - overlap).

Durée de retard à date = Total archive + Flux résiduel (calculé indépendamment pour VG, AM, SL).

NB : la hiérarchie métier (SL ⇒ AM & VG, AM+VG ⇒ VG) n’implique pas de fusion de durées : on calcule par nature.

Formule Access (même logique que dans tes requêtes)

Pour chaque nature, on compare le flux au MAX(EndDate) de l’archive, et on soustrait le chevauchement inclusif (+1) :

IIf(MaxEnd Is Null Or MaxEnd < StartFlux,
    NbJoursRetard,
    IIf(NbJoursRetard - (DateDiff('d', StartFlux, MaxEnd) + 1) < 0,
        0,
        NbJoursRetard - (DateDiff('d', StartFlux, MaxEnd) + 1)
    )
)

Le plus simple dans tes rapports

Comme on a déjà construit un flux résidualisé :

Q_Flux_Residualized = flux “nettoyé” du déjà archivé

Q_Episodes_All = Archive UNION Flux résidualisé

…alors ta durée à date par nature se résume à :

SELECT Booking, [NO DOSSIER CREDIT], Nature, SUM(Duree) AS Duree_Retard_A_Date
FROM Q_Episodes_All
GROUP BY Booking, [NO DOSSIER CREDIT], Nature;


Ça additionne directement “tout archive” + “flux résiduel” sans doublon, pour VG / AM / SL.

## precision 2

toutes les requetes qui me données les montant ou retard , archive + flux en VG et AM et SL prennent bien en compte ça

Oui — à condition d’utiliser la vue “flux résidualisé” dans les requêtes qui additionnent durées ou montants.
C’est déjà ce qu’on a posé : Archive (épisodes clos) UNION ALL Flux résidualisé (jours restants après soustraction du chevauchement avec le dernier épisode archivé, avec le +1 inclusif dans DateDiff).
Concrètement
	• Pas de double-compte : pour chaque nature (VG, AM, SL) et pour chaque clé (Booking+N° crédit), on additionne
Σ Duree_archive + Duree_flux_residuelle.
	• Les requêtes “listing” (1.1 / 1.2 / 1.3 / 1.4) ne somment rien → pas concernées.
	• Les requêtes qui agrègent (cumuls, tops, KPI) doivent s’appuyer sur la vue Q_Episodes_All (= Archive + Flux résidualisé).
C’est ce que j’ai proposé/ajusté.
Mini check-list
	• Dans ta vue UNION (Archive + Flux résidualisé), on garde par nature les champs :
Booking, [NO DOSSIER CREDIT], Nature, StartDate, EndDate, Duree, WorstInsuffisance, DateWorstInsuffisance, LastInsuffisance, LastUpdateDate, DepassementNA.
Si tu veux sommer les montants à date, ajoute un alias Montant:
		○ Archive : [Montant VG] / [Montant AM] / [Montant SL] AS Montant
		○ Flux : f.[Montant VG] / f.[Montant AM] / f.[Montant SL] AS Montant
Deux requêtes “prêtes à coller” pour tes tableaux
	• Durée (jours) à date par clé et nature (Archive + Flux résiduel)
(zéro doublon) :
