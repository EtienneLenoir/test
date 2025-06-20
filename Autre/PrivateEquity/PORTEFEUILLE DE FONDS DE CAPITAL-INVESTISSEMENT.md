Altenative Investments CAIA Level 2- CAIA Association

# PORTEFEUILLE DE FONDS DE CAPITAL-INVESTISSEMENT

## Pour comparer un portefeuille de fonds de capital-investissement

Le portefeuille doit être comparé à un portefeuille similaire de fonds de capital-investissement. Cette comparaison pose toutefois deux problèmes :
1. les fournisseurs de bases de données accessibles au public déclarent qu’il y a trop peu de fonds de fonds pour qu’une comparaison soit significative, et
2. ces fonds de fonds mettent en œuvre diverses stratégies d’investissement, ont des compositions de portefeuille différentes et, ce qui est le plus problématique, ont généralement une structure différente selon le millésime.

Pour contourner ces problèmes, les portefeuilles synthétiques peuvent être générés avec la même allocation aux différentes sous-classes d’actifs (par exemple, millésime, stade et zones géographiques) que celle à référencer. Une telle analyse comparative permet d’évaluer les compétences de sélection du gestionnaire de portefeuille (c’est-à-dire sa capacité à sélectionner les meilleurs gestionnaires de fonds dans les allocations définies). Si le portefeuille est composé de 40 % de rachats et de 60 % de capital-risque, le portefeuille synthétique devrait avoir cette même répartition 40/60.

Comme la performance d’un portefeuille est l’agrégation de la performance d’un fonds individuel, l’analyse comparative d’un portefeuille est simplement le prolongement de l’analyse comparative d’un fonds individuel. Lors de l’analyse comparative, il est important d’utiliser la même méthode d’agrégation pour le portefeuille et l’indice de référence.

---

## Mesures du rendement

Étant donné qu’un portefeuille de capital-investissement est une agrégation de fonds, ses mesures de performance sont simplement l’agrégation des mesures utilisées pour les fonds individuels (IIRR, TVPI, DPI ou RVPI) et peuvent être calculées sur la base de l’une des méthodes suivantes :

### Moyenne simple

La moyenne arithmétique des mesures de performance des fonds de capital-investissement :

$$
\text{IIRR}_{P,T} = \frac{1}{N} \sum_{i=1}^{N} \text{IIRR}_{i,T}
$$

Ici, l’IIRR est l’IIRR du portefeuille \(P\) à la fin de la période \(T\), l’IIRR est l’IIRR du fonds \(i\) à la fin de la période \(T\) et \(N\) est le nombre de fonds dans le portefeuille. Notez que ce taux de rendement moyen peut donner des signaux trompeurs sur la performance du portefeuille. Par exemple, la moyenne simple serait différente de l’IIRR calculée à l’aide des flux de trésorerie moyens. Par conséquent, la moyenne simple de l’IIRR doit être utilisée avec prudence.

### Médiane

La valeur apparaissant à mi-chemin d’un tableau classant la performance de chaque fonds détenu dans le portefeuille.

### Pondérée en fonction de l’engagement

La moyenne pondérée en fonction de l’engagement des mesures de performance des fonds :

$$
\text{IIRR}_{P,T} = \frac{1}{\sum_{i=1}^N CC_i} \sum_{i=1}^N CC_i \times \text{IIRR}_{i,T}
$$

Ici, \(CC_i\) est l’engagement pris de financer le fonds \(i\). Dans ce cas, le rapport d’impact de chaque investissement est pondéré par le montant relatif de l’engagement de chaque investissement.

Notez que ce calcul fonctionne pour l’IIRR ou l’IRR.

Le TRI pondéré en fonction de l’engagement est une moyenne calculée en pondérant les taux de rendement par engagement. Le TRI pondéré en fonction de l’engagement ne tient pas compte de l’investissement réel des fonds. Il est utile pour mesurer l’habileté du LP à sélectionner et à allouer le bon montant aux fonds.

### Pooled

Performance du portefeuille obtenue en combinant les flux de trésorerie et les valeurs résiduelles de tous les fonds individuels, comme s’ils provenaient d’un seul fonds, et en résolvant l’équation pour IIR :

$$
\sum_{t=0}^T \sum_{i=1}^N \frac{CF_{i,t}}{(1 + \text{IIRR}_{P,T})^t} + \sum_{i=1}^N \frac{NAV_{i,T}}{(1 + \text{IIRR}_{P,T})^T} = 0
$$

Ici, \(CF_{i,t}\) est le flux de trésorerie net au cours de la période \(t\) entre le fonds \(i\) et l’investisseur, \(T\) est le nombre de périodes, la VNI est la dernière VNI du fonds \(i\), l’IIRR est l’IIRR du portefeuille \(P\) à la fin de la période \(T\), et \(N\) est le nombre de fonds dans le portefeuille.

Le TRI commun est une mesure qui tente de saisir le moment et l’ampleur des investissements, et est calculé en traitant tous les fonds comme s’il s’agissait d’un seul fonds composite. La série de flux de trésorerie de ce fonds composite est ensuite utilisée pour calculer le TRI commun. L’avantage de cette mesure est qu’elle tient compte de l’ampleur et du calendrier des flux de trésorerie et reflète donc également les compétences du LP à anticiper le marché. L’inconvénient est que les flux de trésorerie plus importants auront plus de poids, de sorte que dans un portefeuille composé de petits fonds en phase de démarrage et de grands fonds à un stade ultérieur ou de rachat, les grands fonds fausseront le tableau.

On peut soutenir que la mesure groupée donne le véritable rendement financier du portefeuille. Cependant, il peut également être judicieux d’utiliser les autres, en fonction de ce que l’on souhaite mesurer. Par exemple, la moyenne simple peut être un bon indicateur des compétences de sélection, tandis que la moyenne pondérée en fonction de l’engagement peut être utile pour évaluer la valeur ajoutée résultant de la décision de la taille de l’engagement à prendre dans chaque fonds spécifique.

Enfin, dans certains cas, le TRI peut ne pas évaluer correctement la performance à long terme. Cela s’explique par le fait que lorsque le TRI est positif, les flux importants des dernières années n’augmentent pas le rendement, mais lorsque le TRI est négatif, ils le feront en termes absolus. Cela explique encore pourquoi il est important d’évaluer également la performance avec, par exemple, le multiple. Une alternative à ce problème consiste à utiliser le TRI zéro, qui est un TRI regroupé calculé en supposant que tous les investissements du portefeuille commencent à la même date. Ceci est utilisé pour éviter que l’ordre des investissements n’affecte le TRI du portefeuille (voir la pièce 8.10). La méthodologie du TRI peut conduire à des réponses multiples ou incorrectes lorsque les flux de trésorerie changent de signe à différents moments. Le problème peut être aggravé lors de l’utilisation de l’approche commune, car plusieurs fonds sur plusieurs périodes sont plus susceptibles d’entraîner des calculs trompeurs du TRI.

La performance des portefeuilles de fonds est mesurée, par exemple, sous forme de TRI commun ou de TRI pondéré en fonction des engagements. La figure 8.11 illustre toutes ces mesures de performance agrégées pour un portefeuille de trois fonds.

---

## Détermination des indices de référence du portefeuille

Outre la question de savoir comment calculer l’IIRR du portefeuille de fonds de capital-investissement d’un investisseur, il y a la question de savoir comment les IIRR d’un groupe de pairs sont agrégés pour déterminer l’IIRR à utiliser comme référence. Prenons le cas des IIRR pondérés en fonction des engagements. Pour comparer des pommes avec des pommes, il faut comparer le rendement du portefeuille pondéré en fonction de l’engagement à celui de l’indice de référence pondéré en fonction de l’engagement.

L’indice de référence du portefeuille peut être construit à l’aide de la moyenne pondérée en fonction de l’engagement de l’indice de référence pour chaque fonds individuel constituant les cohortes du groupe de pairs (par exemple, le même millésime, la même zone géographique et le même stade d’étape) :

$$
BM_{P,T} = \frac{1}{\sum_{i=1}^N CC_i} \sum_{i=1}^N CC_i \times BM_{i,T}
$$

où \(BM_{P,T}\) est l’indice de référence du portefeuille à la fin de la période \(T\), \(CC_i\) est l’engagement dans le fonds \(i\), \(N\) est le nombre de fonds dans le portefeuille et \(BM_{i,T}\) est l’indice de référence du fonds \(i\) à la fin de la période \(T\).

Si tous les fonds investissaient tout leur capital lors du premier démantèlement, le rendement commun et le rendement pondéré en fonction de l’engagement seraient identiques. La question de savoir s’il faut tenir compte de l’impact des engagements non utilisés des fonds crée une confusion supplémentaire. Ni le TRI mis en commun ni le TRI pondéré en fonction des engagements ne reflètent la manière dont le LP gère le capital dédié aux investissements en capital-investissement mais non appelé par les fonds.

## Monte Carlo Simulation

La simulation de Monte Carlo est une technique qui peut être utilisée pour générer des portefeuilles similaires à celui qui fait l’objet d’un benchmarking en tirant, à chaque simulation, le même nombre de fonds qui sont dans le portefeuille parmi toutes les cohortes de groupes de pairs pertinents, puis en pondérant les performances en fonction de la taille de l’engagement des fonds du portefeuille.

Par exemple (en ignorant pour des raisons de simplification le millésime et les dimensions de risque géographiques), pour un portefeuille composé de huit fonds en phase de démarrage et de cinq fonds en phase ultérieure, la simulation tirera pour chaque série huit fonds du groupe de référence des pairs en phase de démarrage et cinq fonds du groupe de référence des pairs en phase ultérieure. Après pondération de la performance de chaque fonds tiré en fonction d’une taille d’engagement correspondante, une performance de portefeuille est obtenue. Cette opération est répétée plusieurs fois afin de créer une distribution, qui est ensuite utilisée pour comparer le portefeuille. Le taux d’indexation pondéré en fonction des engagements du portefeuille est ensuite comparé à cet indice de référence synthétique (voir figure 8.12).

Les résultats obtenus grâce à cette approche doivent être analysés avec soin. En effet, par construction (c’est-à-dire les choix aléatoires), il est implicitement supposé que le gestionnaire de fonds connaît et a accès à l’ensemble de la population des groupes de référence de pairs. En outre, il est également implicitement supposé que le gestionnaire ne prend aucune décision sur la répartition entre les marchés de capital-investissement ou sur le niveau de diversification, ce qui en réalité, ce n’est souvent pas le cas. Ces limitations peuvent être résolues en exécutant une simulation qui reflète mieux la flexibilité accordée au gestionnaire ou les contraintes qui lui sont imposées. Par exemple, alors qu’un portefeuille peut être composé de 60 % de fonds de rachat et de 40 % de fonds de capital-risque, la composition des portefeuilles de référence peut se situer entre 50 % et 75 % de fonds de rachat et entre 25 % et 50 % de fonds de capital-risque, si c’est ce qui est prescrit dans la politique d’investissement imposée au gestionnaire évalué.

## Conclusion

La présence d’un indice de référence est essentielle pour tout programme d’investissement de grande envergure (par exemple, le programme d’allocation d’actifs des investisseurs institutionnels). Les indices de référence sont utilisés à diverses étapes du processus d’investissement et sont utilisés pour élaborer des modèles stratégiques d’allocation d’actifs. Une fois qu’une allocation stratégique a été approuvée, l’étape suivante est la sélection du gestionnaire. Encore une fois, les points de repère jouent un rôle essentiel dans le processus d’évaluation. Une fois les allocations effectuées, les gestionnaires doivent faire l’objet d’un suivi régulier. Pour comprendre si le gestionnaire fonctionne aussi bien que prévu, un point de référence ainsi que d’autres mesures de performance sont utilisés.

Pour construire des points de référence significatifs, il faut disposer de mesures de performance bien définies et ancrées dans les principes fondamentaux de la finance. Ce chapitre a examiné diverses méthodes de mesure de la performance des investissements en capital-investissement. Ces mesures peuvent également être utilisées pour comparer la performance d’un portefeuille à deux types d’indices de référence. L’un d’eux est construit à l’aide de groupes de pairs, ce qui signifie que l’indice de référence est un portefeuille de gestionnaires ayant des caractéristiques similaires à celles du gestionnaire évalué. Alternativement, on peut utiliser un indice de titres cotés en bourse. Dans ce cas, diverses mesures équivalentes au marché public ont été développées. Deux de ces méthodes ont été examinées dans le présent chapitre.