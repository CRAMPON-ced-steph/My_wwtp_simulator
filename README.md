# OCEAN — port React de la filière eau

Port des modules VBA du classeur `140822_OCEAN_CCR.xlsm` (Veolia Water Tech) en nœuds React
autonomes, assemblables par glisser-déposer pour simuler une station.

```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # dist/
```

## Architecture

```
src/core/stream.js       flux d'eau (Q m³/j, charges kg/j), flux de boues, données de site (Valeurs_générales)
src/core/hypotheses.js   onglet Hypothèses + constantes pd_*, ratio() (AA_collection), fonctions communes :
                         facteur K, saturation O2, Patm, rendement moteur, répartition pompes, précipitation P
src/core/engine.js       defineNode(), runChain() : enchaîne les nœuds en double passe
                         (dimensionnement sur l'eau nominale, fonctionnement réel sur l'eau × NC_*)
src/nodes/*.js           un fichier par procédé (voir tableau)
src/nodes/_stub.js       fabrique de nœud "à porter" (traversée sans modification)
src/components/          Palette (glisser-déposer), Canvas (chaîne + tuyaux), Inspector (Vp / Vr / Ve, résultats, élec)
vba-source/              classes VBA découpées par procédé + modules communs + Hypotheses.csv
```

### Contrat d'un nœud

```js
defineNode({
  id, label, family, vba, description,
  choices: [{ key, label, options:[{value,label}], default }],          // boutons btn_* / *_choix_*
  params:  [{ key, label, unit, group, default: (ctx) => number, hint }], // Vp (défaut) — l'utilisateur force Vr
  compute(ctx) => { outNominal, outReel, sludge?, eauxSales?, reactifs?, results, electricity, warnings }
})
// ctx = { site, inNominal, inReel, choices, p (=Ve), forced (=Vr), defaults, upstream }
```

`p[key]` est la valeur effective (forcée si présente, sinon défaut) ; `forced[key]` permet de
reproduire les branches `If xxx_force = True`. `upstream.primaire` reproduit `choix_primaire`.

## État du port

| Nœud | Classe VBA | État |
|---|---|---|
| Dessablage – déshuilage | C2_Dessablage_Deshuilage | porté |
| Décantation simple / lamellaire | D1_decanteur_simple | porté |
| Décantation avec réactifs | D2_Decantation_reactif | porté |
| BA forte charge | E1_BA_forte_charge | porté |
| BA moyenne charge | E2_BA_moyenne_charge | porté |
| BA faible charge | E3_BA_faible_charge | porté (fabrique `atvFaibleCharge.js`) |
| Aération prolongée | E4_BA_aeration_prolongee | porté (= E3 avec G_ref 20 j, NH4 1, NO3 6) |
| HybAS | E5_HybAS | porté (6 configurations de cuves aérées) |
| MBBR | E6_MBBR | porté (10 cuves, 14 flux, mode Mox) |
| MBR | — (absent du classeur) | extension dérivée de E3 : membranes au lieu du clarificateur |
| Biostyr | E8_Biostyr | porté (C, N, NDNc, NDNs) |
| Biostyr PDN | E9_Biostyr_PDN | porté (fabrique `biostyrPdn.js`) |
| Biostyr nitrifiant III | F1_Biostyr_N_III | porté |
| Biostyr PDN III | F2_Biostyr_PDN_III | porté (= E9, MES sur by-pass) |
| Décantation tertiaire | F3_Decantation_III | porté |
| Discfilter | F4_Discfilter_III | porté |
| Filtration sable | F5_Filtration_sable | porté |
| Désinfection UV | G1_Desinfection_UV | porté |
| Chloration | G2_Desinfection_Cl | porté |
| Décantation eaux sales | H1_Decantation_eaux_sales | porté (rebouclage en tête) |

Les vingt classes de la file eau du classeur sont désormais portées, plus l'extension MBR.
Pour reprendre un module : ouvrir `vba-source/<classe>.cls`, transcrire
`hypotheses` / `attribution_valeur_par_defaut` dans `params`, puis `dimensionnement`,
`fonctionnement_reel` et `calcul_consommation_electrique` dans `compute()`.

## Écarts assumés par rapport au VBA

- **C2** : `graisse_DCO` utilise `ratio("codigestion_graisses","NK_MV")` (0,002) dans le VBA, ce qui
  est manifestement une erreur ; le port utilise `DCO_MV` (2,8).
- **E1** : les `If T<12 … If T<18 …` sans `ElseIf` écrasent la valeur T<12. Le port reproduit ce
  comportement (`VBA_BUG_COMPAT = true` dans `baForteCharge.js`) ; passer à `false` pour 3 / 2,5 / 1,5 g/L.
- **E1** : dans le VBA, `boues_Q` est soustrait de Q avant d'être calculé (valeur de l'appel
  précédent). Le port calcule les boues avant de mettre à jour Q.
- **D1** : le module CCR référence des variables non déclarées (`boues_concentration`,
  `NombreOuvrages`…) ; le port les mappe sur les paramètres du bloc Calculs.
- Les consommations spécifiques surpresseurs sont celles codées dans E1 (roots 4,5) et non celles de
  l'onglet Hypothèses (roots 5).
- **E3/E4** : le VBA déclare `Pt_residuel = 0,1 mg/L` sans jamais l'utiliser et peut produire un Pt
  de sortie négatif (bio-P + assimilation supérieurs au Pt entrant). Le port plafonne le Pt de
  sortie à ce résiduel et l'affiche en avertissement.
- **E3/E4** : `boues_Q`, `boues_chimiques` et `boues_methanol` sont des états de classe partagés
  entre `dimensionnement` et `fonctionnement_reel` (utilisés avant d'être calculés) ; le port
  reproduit cet enchaînement (première passe avec 0).
- **Biofiltres (E8/E9/F1/F2)** : le VBA calcule `rdt_DBO = rdt_DCO + 0,1` sans plafond, ce qui
  peut produire une DBO de sortie négative en tertiaire ; le port plafonne `rdt_DBO` à 1.
- **E8/F1** : le Pt traverse le biofiltre inchangé (comportement VBA : la part assimilée est
  réglée dans la file boues, non modélisée ici).
- **F2** : les MES éliminées y sont calculées sur la fraction by-passée `(1 − ratio)` là où E9
  utilise `ratio` ; l'écart du VBA est reproduit (drapeau `mesElimineSurBypass`).
- **Rebouclage des eaux sales** : le moteur reproduit la boucle `iteration_ES` d'OCEAN. Les flux
  `eauxSales` émis par les nœuds (biofiltres, Discfilter, filtre à sable) retournent en tête de
  filière ; si un nœud « Décantation des eaux sales » (H1) est présent, ils y sont d'abord
  traités et seul le surnageant retourne en tête. Itération à point fixe (max 4 passes,
  convergence sur le débit de retour à 0,5 %).
- **MBR** : n'existe pas dans OCEAN. Le nœud dérive du cœur biologique E3 (fabrique
  `atvFaibleCharge.js`) : membranes dimensionnées sur le flux de pointe, boues extraites à la
  concentration du bassin, électricité de décolmatage SADm × surface × cyclage et pompes de
  perméat. Les ratios de boues `II_MBR` sont marqués « A REVOIR » dans le classeur.

### E6 — MBBR

- Le choix du média, de l'aérateur et du mode d'agitation est proposé **par étape** et non cuve par
  cuve comme dans la feuille de calcul ; le calcul, lui, reste bien par cuve (le média de chaque
  cuve est une variable distincte).
- `reox_DBOremoved` est déclarée dans la classe et utilisée dans le calcul de la charge à nitrifier,
  mais **jamais affectée** : elle vaut donc toujours 0. Reproduit tel quel.
- Dans `fonctionnement_reel`, la production de boues de la cuve de dé-ox est calculée avec un
  `deox_nit` **local** resté nul, alors que la variable de classe du même nom porte la valeur utile.
  La nitrification de la dé-ox ne contribue donc pas aux boues au réel. Reproduit tel quel.
- Plusieurs affectations d'O2 dissous emploient l'indice `I` hérité d'une boucle `For` précédente
  (`cuve_O2_dissous(I)` dans les blocs dé-ox et ré-ox). Le port conserve cette variable et son
  indice, y compris quand il pointe une cuve inattendue.
- Les variables de classe `boues_Q`, `boues_postdenit` et le flux post-séparation valent 0 au
  premier passage du dimensionnement : reproduit.
- Les boucles de point fixe et de Newton-Raphson sont bornées (200 à 300 itérations) et lèvent un
  avertissement en cas de non-convergence ; le VBA les laisse non bornées.

### E5 — HybAS

- `dimensionnement()` est **relancé jusqu'à stabilisation** (25 passes au plus). Dans OCEAN, les
  variables que cette routine corrige d'un appel à l'autre (`reox_nit`, `boues_methanol`,
  `DBO_sortie`, `recirculation_taux` incrémenté de 0,1 quand les boues décantent mal,
  `NO3_last_aerated_tank` réduit de 10 % quand la recirculation calculée est négative) ne prennent
  effet qu'à l'itération suivante du programme principal, que le nœud isolé ne voit pas.
- `dimensionne_config_4_6` calcule sa borne haute de dichotomie par
  `1000 * nit / 24 * MES * rate`, c'est-à-dire une **multiplication** là où une division est
  attendue : la borne obtenue est inférieure au volume cherché et bloque la dichotomie. Le port
  retient la formule correcte, avec un facteur de sécurité de 5.
- Dans `fonctionnement_reel`, les besoins en O2 dus aux sulfures sont ajoutés à
  `cuve_aeree_O2_besoins(I)` alors que la variable de boucle est `cuve` : `I` vaut 0 et la
  contribution est perdue. Reproduit, avec un avertissement quand de l'H2S est présent.
- L'azote assimilé (par la biomasse des cuves aérées, par la dé-ox, par la biomasse sur méthanol en
  post-dénitrification) peut dépasser l'ammonium puis le nitrate disponibles et rendre les
  concentrations de sortie négatives. Le port plafonne NH4 et NO3 à zéro et le signale.
- Le choix du média et de l'aérateur est proposé globalement pour les cuves H, non cuve par cuve.
- Toutes les boucles imbriquées (boues → liqueur mixte → dénitrification → post-dénit) sont bornées.

## File boues, retours, utilités, transverse

Les sources VBA des quatre blocs restants ont été extraites du classeur et versées dans
`vba-source/_file_boues/`, `_retours/`, `_utilites/` et `_transverse/`. Le classeur compte
106 000 lignes de VBA au total ; la file eau en représentait 20 000, il en reste 53 000 à porter.

| Bloc | Classes | Lignes VBA | État |
|---|---|---|---|
| 1 — File boues | 11 | 22 300 | **11 portés sur 11 — bloc terminé** |
| 2 — Retours et deammonification | 6 | 5 400 | **6 portés sur 6 — bloc terminé** |
| 3 — Utilités | 17 | 22 400 | **portés — le PINCH est réécrit, non transposé** |
| 4 — Transverse | 4 | 2 900 | **4 portés sur 4 — bloc terminé** |

### Infrastructure de la file boues

`src/core/sludge.js` reproduit la double indexation d'OCEAN : les flux de boues sont portés par
`TableauRecapitulatifFluxBoues(étape, type_boue, paramètre)`, 23 étapes × 6 types × 10 paramètres,
doublé de `boues_pollution_soluble` pour la pollution dissoute véhiculée par l'eau interstitielle.
C'est cette indexation par origine de boue qui permet à l'épaississement de retrouver la
concentration de référence propre à chaque procédé amont, et au bilan matière de se vérifier étape
par étape (colonnes `verif_flux` / `flux_in`).

`src/core/sludgeEngine.js` enchaîne les procédés, alimente le vecteur de retour en tête et calcule
le flux évacué comme le reliquat non consommé de chaque étape. `apportsDepuisFileEau()` fait la
jonction avec la file eau : les nœuds qui ne renseignent pas la composition de leurs boues la
voient reconstituée depuis la table de ratios d'`AA_collection` (DCO, DBO et NK rapportés aux
matières volatiles, Pt aux MES), exactement comme le programme principal d'OCEAN.

### Procédés portés

- **Épaississement** (`z_Epaississement.cls`) : gravitaire, centrifugeuse, table d'égouttage et
  flottation. Taux de capture par technologie, siccité de sortie calculée comme la moyenne des
  concentrations de référence par origine de boue pondérée par les MES, dosage et électricité de
  préparation du polymère, pompes d'alimentation et d'extraction avec leur arbre de répartition
  débit / nombre / durée, et pour la flottation le circuit d'eau blanche (pompe + compresseur,
  corrélation Predimboo). Jusqu'à cinq épaississeurs en parallèle, chacun alimenté par une ou deux
  entrées avec une part de flux admise réglable.

Écart signalé : le VBA calcule le débit de boues épaissies avec une siccité d'amorçage propre à la
technologie, puis la recalcule par origine sans reprendre les débits. Le port reprend les débits
avec la siccité définitive, faute de quoi le bilan matière ne boucle pas.

- **Digestion anaérobie** (`z_Digestion_simple.cls`) : mésophile ou thermophile. La réduction des
  matières volatiles part d'un taux de référence par origine de boue (0,638 pour une boue primaire,
  0,30 pour une boue d'aération prolongée), corrigé de la température par une courbe à deux
  optimums et du temps de séjour par une loi en puissance. Les MV détruites donnent le biogaz, dont
  la densité suit la teneur en méthane ; l'azote et le phosphore libérés passent en pollution
  soluble et repartiront en tête aux étapes aval. Bilan thermique complet (pertes de l'enceinte
  selon climat et isolation, chauffage des boues), agitation mécanique, par recirculation ou par
  injection de biogaz, gazomètre dimensionné sur le mode de valorisation.

  Écarts signalés : la codigestion de coferments (BMP, facteur de sécurité) n'est pas portée, elle
  dépend d'une collection saisie dans un formulaire du classeur ; le détail par eaux sales
  (`boues_I_detail`…) n'est pas reproduit, le coefficient de réduction étant pris sur l'origine
  principale de chaque type de boue, ce qui revient au même hors filière biofiltration séparée ;
  l'énergie thermique est calculée et exposée mais aucun consommateur ne la récupère tant que la
  gestion d'énergie n'est pas portée.

- **Déshydratation** (`z_Deshydratation.cls`) : centrifugeuse, filtre à plateaux, filtre à bandes.
  Siccité de sortie par moyenne des siccités de référence par origine, majorée d'environ 20 g/L au
  prorata de la part de boues digérées dans l'alimentation. Le dosage de polymère n'est pas une
  constante par technologie mais dépend de l'origine de chaque boue *et* de son passage éventuel en
  digestion : deux barèmes distincts, appliqués flux par flux. Conditionnement chaux + FeCl3 pour le
  filtre à plateaux, avec conversion CaO / Ca(OH)2. Les filtrats et l'azote solubilisé repartent en
  tête ; la concentration en NK des filtrats est exposée, c'est elle qui dimensionnera un futur
  traitement des retours.

  Écarts signalés : le circuit technosable (filtre à plateaux en aval d'un Athos, siccité 600 g/L,
  retour dédié) n'est pas porté tant qu'Athos ne l'est pas ; le VBA tient un vecteur
  `retour_digestion` distinct du retour général pour isoler les retours chargés en azote, le port
  n'en tient qu'un seul et expose la part digérée en résultat.

- **Chaulage** (`z_Chaulage_boues.cls`) : hygiénisation à 30 % de CaO rapporté aux MS, ou
  relèvement de siccité à 300 g/L par un bilan de masse tenant compte de l'eau consommée par
  l'extinction de la chaux vive, ou le plus élevé des deux taux. Chaux vive ou éteinte. Le chaulage
  n'extrait pas d'eau : pas de retour en tête, mais un gain de masse sèche qui figure au bilan.

  Écart signalé : le VBA ajoute au flux sortant une contribution `stockage_boues_pdtes` issue d'un
  stockage amont non renseigné dans le port ; elle est ignorée.

- **Séchages** (`Sechage_thermique.cls`, `Sechage_bioco.cls`, `sechage_inos.cls`) : les trois
  classes partagent le même cœur — évaporation d'eau, besoins thermiques proportionnels à l'eau
  évaporée, condensation des buées, retour des condensats en tête — et sont donc produites par une
  **factory**, comme E3/E4 et E9/F2 sur la file eau. Elles ne diffèrent que sur trois points :

  | | Électricité | Chaleur | Particularité |
  |---|---|---|---|
  | Thermique | 50 à 150 kWh/TEE | 0 à 1 070 kWh/TEE | 4 technologies ; la CMV ne consomme pas de chaleur |
  | Bioco | 43 kWh/TEE | 880 kWh/TEE | technologie unique |
  | Inos | corrélation sur le tonnage | 1 100 kWh/TEE | déshydratation intégrée au sécheur |

  La condensation des buées existe en trois montages : indirecte par échangeur (la chaleur latente
  est récupérable et exposée en résultat), directe par lavage à l'eau froide (le débit d'eau se
  déduit d'un bilan enthalpique entre la chaleur latente à évacuer et l'échauffement admis sur
  l'eau de lavage), ou indirecte puis directe. Le Inos calcule d'abord une siccité intermédiaire
  par origine de boue, comme un filtre à plateaux, et n'évapore que l'écart jusqu'à 90 %.

  Écarts signalés : l'ajustement automatique de la siccité pour rendre les boues autocombustibles
  en incinération (`incineration_siccite_ajustable`) n'est pas porté tant que l'incinération ne
  l'est pas ; le VBA cumule la pollution soluble à chaque itération de la boucle sur les types de
  boue, ce qui la compte autant de fois qu'il y a de types présents, le port ne la compte qu'une
  fois ; **le Inos ne verse pas au retour en tête l'eau retirée par sa déshydratation intégrée** —
  une centaine de m³/j sur une station de 100 000 EH — seule la condensation des buées y figure. Ce
  trou de bilan hydrique du classeur est reproduit, mais le volume manquant est calculé, exposé en
  résultat et signalé par un avertissement.

- **Incinération** (`Incineration_boues.cls`) : four à lit fluidisé. Tout le calcul tourne autour de
  l'autocombustibilité — le PCI des boues humides est comparé à l'énergie qu'il faut pour porter
  fumées, boues et air de combustion à la température de four visée, pertes comprises. Si le PCI
  est insuffisant, l'écart donne l'appoint de combustible (biogaz, gaz naturel ou fioul, l'air
  préchauffé venant en déduction du PCI au dénominateur) ; s'il est excédentaire, l'écart devient
  une énergie de refroidissement récupérable. La **siccité d'autocombustibilité** est calculée et
  exposée : c'est la valeur que le séchage amont devrait viser pour se passer d'appoint.

  Suit la chaîne fumées : le soufre et le chlore des boues, dont les teneurs dépendent du type de
  boue et de la présence d'une digestion, donnent le SO2 et le HCl à neutraliser en voie sèche
  (bicarbonate + charbon actif, avec dilution à l'air froid avant les manches) ou humide (soude).
  Les sels formés et les réactifs en excès constituent les REFIB. Traitement SNCR des NOx à
  l'ammoniaque ou à l'urée, diamètre du four dimensionné sur la vitesse des fumées aux conditions
  réelles, et barèmes de consommation thermique de maintien et de démarrage.

  **Correction apportée au classeur.** Le VBA emploie deux ratios d'air comburant différents dans
  la même routine : 10,9 kg d'air par kg de MV au calcul d'autocombustibilité et 11,4 au débit
  d'air, pour des boues digérées. L'origine est un nommage inversé des deux constantes par rapport
  à leurs propres commentaires :

  ```
  air_inlet_kgair_kgMV_boues_digerees     = 10.9  'Débit d'air boues mixtes ou biologiques
  air_inlet_kgair_kgMV_boues_non_digerees = 11.4  'Débit d'air boues digérées
  ```

  Ce sont donc les noms qui sont faux, pas l'un des deux usages : les commentaires font foi, et le
  sens qu'ils donnent est celui qu'on attend physiquement — les MV résiduelles d'une digestion,
  plus réfractaires, demandent davantage d'air comburant. Le port réassocie les valeurs (digérées
  11,4, fraîches 10,9) et les applique aux deux endroits. Le débit d'air était déjà correct et ne
  bouge pas ; c'est le PCI requis pour l'autocombustibilité qui est corrigé, de quelques pour cent :
  la siccité d'autocombustibilité passe de 33,7 à 34,4 % sur une boue digérée, et l'appoint de
  combustible baisse de 29 % sur une boue fraîche non autocombustible.

  Écarts signalés : la siccité d'autocombustibilité n'est pas réinjectée dans les sécheurs amont,
  le moteur de la file boues exécutant les nœuds dans l'ordre sans seconde passe ; le calcul des
  réactifs de neutralisation ajoute la part « excès » à la somme déjà dosée, ce qui revient à
  compter deux fois le stœchiométrique — reproduit tel quel, faute de savoir si c'est un choix de
  sécurité ou une erreur.

- **Athos** (`z_CLS_BouesAthos.cls`) : oxydation par voie humide, 250 °C sous 55 bars, à l'oxygène
  liquide. Le procédé est une boucle thermique fermée — entrée, épaississement, mélange avec la
  recirculation de technosable, échangeur de préchauffage, réacteur, deux échangeurs de
  récupération, décanteur, dont la moitié du volume repart en tête de boucle. Deux calculs imbriqués
  gouvernent tout le reste :

  1. un **bilan d'énergie de cinq équations à cinq inconnues** (températures d'entrée et de sortie
     du premier échangeur, température du réacteur, énergie échangée entre les deux échangeurs,
     température de sortie du second), résolu par élimination de Gauss avec pivot partiel ;
  2. une **dichotomie sur la siccité des boues épaissies**, entre 4 et 10 %, pour trouver la valeur
     à laquelle le réacteur atteint exactement 250 °C sans apport extérieur. Si les boues brutes
     suffisent, aucun épaississement n'a lieu ; sinon on épaissit juste ce qu'il faut.

  Le bilan de masse est lui-même itératif, la recirculation de technosable bouclant sur le mélange
  d'entrée. L'humidité des fumées saturées emploie le polynôme de degré 6 de
  `Pv_sat_H2O_bar` (MOD_FonctionsPubliques), calé sur les données thermodynamiques du classeur :
  une corrélation générique de type Rankine surestime la pression de vapeur saturante de 25 % à
  250 °C, ce qui suffit à faire diverger le bilan d'énergie.

  Sur les cas testés, la dichotomie converge à 6,6 % de siccité pour des boues fraîches et 8,3 %
  pour des boues digérées — l'écart est logique, les boues digérées ayant perdu une partie de leurs
  matières volatiles, il faut les concentrer davantage pour que la réaction reste autotherme.

  Écarts signalés : le VBA tient trois vecteurs de retour distincts (général, digestion, Athos) pour
  permettre un traitement séparé des retours azotés, le port n'en tient qu'un et expose la part
  Athos en résultat ; le circuit « technosable » de la déshydratation par filtre à plateaux n'est pas
  encore branché ; **le VBA n'affecte jamais les MES du surnageant du décanteur Athos** — les MS qui
  échappent à la décantation, environ 500 kg/j sur une station de 100 000 EH, ne figurent dans aucun
  bilan. Reproduit, mais le flux manquant est calculé, exposé en résultat et signalé.

- **Biothelys** (`z_Biothelys.cls`) : lyse thermique en batch à 165 °C par injection de vapeur
  saturée, suivie d'une digestion anaérobie. Ce n'est pas un simple prétraitement : la classe porte
  sa propre digestion, et surtout **deux** tables de réduction des matières volatiles.

  | Origine | Non lysée | Lysée | Gain |
  |---|---|---|---|
  | Boue primaire | 0,638 | 0,673 | +5 % |
  | Boue faible charge (eau brute) | 0,378 | 0,472 | +25 % |
  | Boue MBR | 0,300 | 0,394 | +31 % |
  | Graisses | 0,629 | 0,765 | +22 % |

  Le gain est d'autant plus marqué que la boue est réfractaire : négligeable sur une boue primaire
  déjà très dégradable, il dépasse 30 % sur une boue de MBR. La table des non lysées est identique
  à celle de `digestion.js`, ce qui permet d'alimenter le digesteur avec un mélange de boues lysées
  et non lysées — le port gère cette entrée by-pass — en conservant la cinétique propre à chacune.

  La consommation de vapeur suit une corrélation linéaire décroissante sur la température des boues
  entrantes, et non un bilan enthalpique : le procédé est un batch dont une partie de la vapeur est
  récupérée par détente flash entre réacteurs. La corrélation cesse d'être valable au-delà de 35 °C
  en entrée, la récupération flash diminuant ; le port le signale. La température de vapeur saturée
  est calculée par la corrélation du classeur — 189,9 °C à 12,5 bars, valeur exacte.

  Sur une filière primaire + faible charge, la lyse fait passer la réduction des MV de 48 à 55 %,
  soit +14 % en relatif après pondération par les MV de chaque origine, pour 220 kg de vapeur par m³
  de boues et 17 000 kWh/j de besoin thermique — dont 5 700 récupérables au refroidissement des
  boues lysées.

  Écarts signalés : la codigestion de coferments (BMP, part de graisses internes) n'est pas portée,
  elle dépend d'une collection saisie dans un formulaire du classeur ; le raccordement au module
  PINCH d'intégration énergétique n'est pas fait, les besoins et disponibilités thermiques étant
  calculés et exposés sans qu'aucun consommateur s'en saisisse.

- **Exelys DLD** (`z_Exelys_DLD.cls`) : configuration Digestion-Lyse-Digestion. Contrairement à ce
  que la parenté de nommage laissait supposer, ce n'est pas une variante de Biothelys et **la
  factory envisagée n'avait pas lieu d'être** : la chaîne comporte deux digesteurs séparés par une
  déshydratation et une lyse continue.

  ```
  entrée → digesteur 1 (35 °C) → centrifugeuse (230 g/L) → [+ boues fraîches en by-pass]
         → lyse continue 165 °C → échangeurs → dilution 100 g/L → digesteur 2 (38 °C)
  ```

  Trois particularités par rapport à Biothelys :

  1. La **centrifugeuse intermédiaire** ne capture pas toutes les MV : la première digestion en a
     solubilisé une part, qui repart au centrat. Cette part dépend de l'origine de la boue *et* du
     temps de séjour du premier digesteur, via deux jeux de coefficients directeurs selon que le
     temps de séjour réduit est inférieur ou supérieur à la référence. C'est le calcul le plus
     spécifique du module, et celui qui gouverne le rendement d'ensemble.
  2. Le **second digesteur** applique un taux unique de 0,3656 aux boues digérées puis lysées :
     après ce double traitement, l'origine de la boue n'a plus d'influence. Les boues fraîches
     admises en by-pass gardent en revanche leur taux « lysées » propre.
  3. La consommation de vapeur vient d'un **bilan enthalpique** (procédé continu) et non d'une
     corrélation comme dans le Biothelys en batch.

  Sur une filière primaire + faible charge, le DLD porte la réduction des MV à 62 % contre 44 % en
  digestion simple, soit +40 %, pour 1 925 Nm³/j de biogaz dont 28 % produits par le second
  digesteur — c'est exactement ce que la lyse va chercher : la fraction que la digestion seule ne
  peut plus attaquer.

  Écarts signalés : la codigestion de coferments n'est pas portée ; le raccordement au module PINCH
  n'est pas fait ; le centrat de la centrifugeuse part au retour général alors que le VBA le range
  dans un vecteur `RetoursDigestion` distinct, que le port ne tient pas séparément.

## Bloc 1 terminé

Les onze classes de la file boues sont portées. Le test `test10.mjs` couvre 21 filières et vérifie
le bilan matière sur chacune, en tenant compte des MV parties en biogaz, de la chaux ajoutée au
chaulage, de l'eau évaporée au séchage, de la masse incinérée et des MS oxydées par Athos. L'écart
est nul sur les 21.

## Bloc 2 — retours et deammonification

### Retours séparés par origine

Le bloc 2 a imposé une évolution d'infrastructure : OCEAN tient **cinq vecteurs de retour
distincts** là où le port n'en avait qu'un. La raison est fonctionnelle — le traitement des retours
ne s'applique qu'aux jus les plus chargés en azote, ceux de la digestion et de l'Athos, et non aux
surnageants d'épaississement ou aux condensats de séchage.

`sludge.js` porte désormais `makeRetours()` et `ajouterRetour()` : chaque procédé écrit dans le
vecteur `total` et dans celui de son origine (`digestion`, `athos`, `autres`), avec la part soluble
en parallèle. Les sept procédés qui produisent des retours ont été adaptés. `retour` reste l'alias
du total, si bien que le code existant continue de fonctionner.

### Procédés portés

- **Aiguillage des retours** (`z_CLS_RetoursAdmisTraitement.cls`) : sélectionne les jus dirigés vers
  le traitement — digestion seule, Athos seul, ou les deux — et calcule leurs concentrations. Après
  traitement, `remplacerRetoursAdmis()` substitue les flux traités aux flux admis dans le total.

- **ANITA Mox** (`z_ANITA_Mox.cls`) : déammonification sur biofilm des jus de digestion et d'Athos.
  Trois chiffres résument l'intérêt du procédé, tous portés : 1,94 kg d'O2 par kg d'azote éliminé
  contre 4,57 en nitrification classique ; 0,11 kg de nitrate formé par kg de N traité contre 1 ;
  et aucune source de carbone à ajouter.

  La chaîne comporte au plus deux cuves. La **cuve carbone n'existe que si** le rapport DCO soluble
  dégradable sur N-NH4 dépasse 0,4 : au-delà, les hétérotrophes concurrenceraient les bactéries
  anammox. Elle est alors dimensionnée pour ramener ce rapport à 0,2. La **cuve Mox** est
  dimensionnée sur la charge azotée, avec un temps de séjour minimal de 12 h qui peut la
  surdimensionner.

  Les vitesses de déammonification dépendent du type de digestion amont : 2,6 g N/(m²·j) à 30 °C sur
  des jus de digestion simple, mais 1,35 seulement sur des jus d'hydrolyse thermique, dont la DCO
  résiduelle est plus réfractaire. Sur les cas testés, cela donne 83 % d'élimination de l'azote
  derrière une digestion simple contre 71 % derrière un Exelys, et une cuve carbone quatre fois plus
  volumineuse dans le second cas.

  Écarts signalés : le VBA fait deux passes (dimensionnement puis fonctionnement réel), la file
  boues n'ayant qu'un jeu de flux le port n'en fait qu'une et expose la charge nominale en
  paramètre forçable ; le contrôle du temps de séjour minimal existe dans le VBA mais son message
  d'alerte y est vide, le port émet un avertissement ; les trois gammes de débit du surpresseur
  portent la même consommation spécifique dans le classeur, le port ne retient que le type.

- **ANITA Shunt** (`z_ANITA_Shunt.cls`) : l'alternative à ANITA Mox, par nitritation-dénitritation.
  En arrêtant l'oxydation au nitrite au lieu d'aller jusqu'au nitrate, le procédé économise 25 % de
  l'oxygène (3,43 kg O2/kg N au lieu de 4,57) et 40 % du carbone (2,4 kg DCO/kg N au lieu de 4).

  Différence de nature avec ANITA Mox : le Shunt travaille en biomasse libre dans un réacteur
  séquentiel, précédé d'une bâche tampon dimensionnée sur le décalage entre le rythme
  d'alimentation en jus et le fonctionnement continu du réacteur. Il **produit des boues**,
  extraites à l'âge de boues visé de 15 jours, et consomme du méthanol dès que la DBO soluble des
  jus ne suffit pas à la dénitritation.

  Comparaison sur les mêmes jus de digestion, à charge azotée identique (109 kg N/j) :

  | | ANITA Mox | ANITA Shunt |
  |---|---|---|
  | Élimination de l'azote | 83 % | 83 % |
  | Besoin en O2 | 229 kg/j | 499 kg/j |
  | Méthanol | 0 | 146 kg/j |
  | Boues produites | 0 | 83 kg MES/j |
  | Électricité | 237 kWh/j | 325 kWh/j |

  À performance égale sur l'azote, la voie anammox est nettement plus économe — c'est bien ce qui
  justifie son surcoût d'investissement.

  Écarts signalés : le VBA fait deux passes, le port n'en fait qu'une ; `boues_Q` est employé dans
  le calcul de `boues_MES` avant d'être calculé (il vaut 0 au premier passage et converge sur les
  itérations du programme principal), le port résout directement le point fixe ; **le bilan solide
  ne boucle pas** — les boues extraites sont calculées sur l'âge de boues visé et non par
  différence entre l'entrée et la sortie, ce qui laisse ici 152 kg MES/j non affectés. Reproduit,
  calculé et signalé.

- **Biolix graisses** (`z_Biolix_graisses.cls`) : traitement biologique aérobie des graisses
  extraites au dégraisseur. Les graisses sont diluées à 75 g/L de DCO, homogénéisées un jour, puis
  dégradées en aérobie dans un bassin à **20 jours de temps de passage**. Cette durée est la
  signature du procédé — les graisses sont lentement biodégradables, ce qui est précisément la
  raison pour laquelle on ne les envoie pas telles quelles au traitement secondaire.

  Le procédé consomme des nutriments : riches en carbone mais pauvres en azote et en phosphore, les
  graisses en réclament 25 g de N et 5 g de P par kg de DCO entrante. Sur le cas testé — 227 kg de
  graisses par jour, soit 573 kg de DCO — cela représente 14 kg d'azote et 2,9 kg de phosphore à
  apporter, pour 158 kg de boues produites et 427 kWh/j.

  Écart signalé : `md_PartBouesBiolixVersTraitementSecondaireEau` vaut toujours 1 dans le classeur,
  et le test « si pas de traitement secondaire » y a un corps vide ; les boues partent donc
  systématiquement au secondaire.

### Correction d'infrastructure

Le raccordement du Biolix a révélé un défaut du port : `lireEntrees()` ne vérifiait pas qu'un flux
était encore disponible. Un épaississeur configuré sur « toutes origines » en aval d'un Biolix
reprenait donc les graisses que celui-ci avait déjà consommées. Le contrôle sur `verif_flux` est
ajouté — c'est exactement l'usage auquel cette colonne est destinée dans le classeur.

Le dégraisseur route par ailleurs désormais ses graisses vers le type de boue qui leur est propre,
sans quoi le Biolix ne les retrouvait pas dans la file boues.

- **Cristallisation MAP** (`z_CLS_RetoursMAP.cls`) : troisième voie de traitement des jus, d'une
  nature différente des deux ANITA. Au lieu d'éliminer l'azote, elle le **récupère** avec le
  phosphore sous forme de struvite, un engrais commercialisable :

  ```
  Mg²⁺ + NH₄⁺ + PO₄³⁻ + 6 H₂O → MgNH₄PO₄·6H₂O
  ```

  La stœchiométrie est équimolaire, et c'est ce qui limite le procédé sur l'azote : chaque mole de
  phosphate précipitée n'emporte qu'une mole d'ammonium. Or les jus de digestion ont un rapport
  molaire N/P de l'ordre de 20. Sur le cas testé, la MAP élimine donc 80 % du phosphate mais
  seulement **4 % de l'azote** — là où ANITA Mox en élimine 83 % sans toucher au phosphore. Les deux
  procédés sont complémentaires plutôt que concurrents.

  Le magnésium est le réactif limitant : absent des jus en quantité suffisante, il est apporté sous
  forme de MgCl2 à raison de 1,2 mole par mole de phosphate. Quatre conditions d'applicabilité sont
  vérifiées (P-PO4 ≥ 70 mg/L, MES ≤ 2 000 mg/L, rapport N/P ≥ 2, rapport Ca/P ≤ 1), plus le pH.

  Le port dimensionne aussi la gamme de réacteurs Ostara PEARL. Sur la station de 100 000 EH du
  test, la production de struvite (66 kg/j) reste sous le seuil d'un PEARL 500, ce que le port
  signale — la MAP demande une charge en phosphore que cette filière n'atteint pas.

  Écarts signalés : les quatre contrôles d'applicabilité et celui du pH existent dans le classeur
  mais leurs messages y sont vides, le port émet de vrais avertissements ; la struvite est exposée
  en résultat comme produit valorisé et ne rejoint pas la table des boues.

## Bloc 2 terminé

Les six classes du bloc des retours sont portées. Trois voies de traitement des jus coexistent
désormais, comparables sur la même filière :

| | ANITA Mox | ANITA Shunt | Cristallisation MAP |
|---|---|---|---|
| Azote éliminé | 83 % | 83 % | 4 % |
| Phosphate éliminé | — | — | 80 % |
| Réactif | aucun | méthanol 146 kg/j | MgCl2 38 kg/j |
| Boues produites | aucune | 83 kg MES/j | aucune |
| Produit valorisé | aucun | aucun | struvite 24 t/an |
| Électricité | 237 kWh/j | 325 kWh/j | 46 kWh/j |

## Bloc 3 — utilités

### Infrastructure

`src/core/utilityEngine.js` introduit un troisième moteur. Les utilités ne transforment ni l'eau ni
les boues : elles consomment ou produisent de l'électricité, des réactifs et de la chaleur à partir
de grandeurs déjà calculées par les deux files. `contexteDepuisFilieres()` rassemble ces grandeurs
— débit nominal et réel, sulfures strippés, biogaz produit, besoins thermiques, réactifs consommés
— sans que chaque nœud ait à connaître la structure interne des deux autres moteurs.

Une production d'énergie est portée comme une consommation négative, ce qui permet au bilan
électrique de sommer sans cas particulier tout en exposant séparément l'énergie produite.

### Procédés portés

- **Désodorisation biologique et chimique** (`z_Desodorisation_biologique.cls`,
  `z_Desodorisation_chimique.cls`) : les deux classes partagent le calcul du débit d'air vicié, qui
  se déduit du nombre de renouvellements horaires appliqués au débit nominal — 13 pour une
  couverture partielle des ouvrages, 42 pour une couverture totale. Ce seul facteur triple la
  consommation de ventilation.

  La chimique enchaîne jusqu'à quatre tours de lavage, chacune ciblant des composés distincts :
  acide sulfurique pour l'ammoniac et l'azote organique, oxydo-basique pH 9 puis pH 11 à la soude
  et au chlore pour l'hydrogène sulfuré et les mercaptans, neutre au bisulfite pour neutraliser le
  chlore résiduel.

  Le point notable du port : **la teneur en H2S de l'air vicié n'est pas une donnée**, elle se
  déduit des sulfures strippés par la file eau rapportés au débit d'air. La qualité de l'eau brute
  gouverne donc directement les consommations de réactifs. Sur une eau à 5 mg/L de sulfures,
  l'air vicié titre 95 mg/Nm³ de H2S et la consommation de chlore passe de 9 à 769 kg/j.

  Écart signalé : la tour pH 11 remet à zéro H2S et mercaptans comme la pH 9 le fait déjà ;
  enchaîner les deux ne consomme donc de réactif soufré qu'à la première. Reproduit, et signalé
  quand les deux tours sont demandées.

- **Turbine hydraulique** (`Turbine_hydraulique.cls`) : récupération sur la chute en sortie de
  station. Le rendement global retenu est la moyenne entre le rendement nominal (0,90) et celui
  constaté à 70 % de charge (0,67), ce qui donne 0,79.

- **Gestion des réactifs** (`Gestion_reactifs.cls`) : consolide les consommations des deux files et
  des utilités, et les convertit en produit commercial selon la pureté de chaque produit. La soude
  à 30 % pèse ainsi 3,3 fois son équivalent pur. Seul l'hypochlorite fait exception : son titre
  s'exprime en degrés chlorométriques et non en fraction massique.

### HVAC — modèle climatique simplifié

Le classeur s'appuie sur une base de 10 896 stations météo pour calculer des degrés-jours et des
grammes-jours jour par jour. Embarquer cette base représenterait plusieurs mégaoctets pour un module
qui pèse quelques pour cent de la consommation d'une station.

Le port retient un **modèle climatique sinusoïdal** : la température mensuelle se reconstitue à
partir de trois grandeurs — moyenne annuelle, amplitude, hémisphère — issues d'un préréglage (sept
climats types) ou saisies directement. Les degrés-jours en découlent par intégration mensuelle,
méthode usuelle en thermique du bâtiment. **Les sorties sont identiques à celles du classeur** :
besoins de chauffage et de climatisation au m², consommations d'électricité, de gaz et de fioul,
COP, débits d'eau usée en entrée de pompe à chaleur, coefficient de déperdition, taux de
renouvellement.

Les trois bâtiments (administratif, exploitation, local électrique) ne diffèrent que par leurs
consignes, leur hauteur sous plafond et leur ventilation : ils sont produits par une factory. Le
bâtiment de bureaux calcule son débit d'air neuf depuis l'occupation, les deux autres depuis un taux
de renouvellement.

**Un piège numérique évité.** Le polynôme de pression de vapeur saturante du classeur, repris pour
l'Athos, est calé sur 0–373 °C : il est juste au-dessus de 50 °C mais **devient négatif aux
températures ambiantes** — −0,025 bar à 20 °C contre 0,023 en réalité. Employé tel quel, il donnait
un besoin de déshumidification vingt fois trop élevé. Le HVAC emploie donc la formule de Magnus,
précise à mieux que 0,4 % entre −45 et +60 °C.

Sur un bureau de 400 m² en climat tempéré, isolation RT 2000 : 181 kWh/(m²·an) de chauffage et
14 de climatisation — des ordres de grandeur cohérents avec la pratique.

`outils/extraire_meteo.py` reste fourni pour qui voudrait alimenter le modèle depuis la base du
classeur plutôt que depuis un préréglage.

### Photovoltaïque

Bonne surprise à la lecture : le modèle du classeur est **astronomique**, pas tabulé. Déclinaison
solaire, angle horaire au lever, distance Terre-Soleil — tout se calcule à partir de la seule
latitude et du jour de l'année. Cette partie est reprise telle quelle, sans base de données.

La seule grandeur puisée dans la base de stations est l'indice de clarté, fraction du rayonnement
extraterrestre atteignant le sol. Le port en fait un paramètre avec des préréglages par climat, de
0,40 sous un ciel océanique à 0,65 en désertique — une donnée de site que l'utilisateur connaît
mieux qu'un identifiant de station.

Deux calages ont été nécessaires. Le rapport des angles d'incidence à midi solaire donne le gain
**maximal** de la journée, pas le gain intégré : l'écart au facteur unité est atténué de moitié, ce
qui restitue les 12 à 15 % attendus pour une inclinaison égale à la latitude. Les indices de clarté
ont ensuite été ajustés sur les productibles observés. Le résultat : **1 135 kWh/(kWc·an) à Paris et
1 456 à Marseille**, soit les valeurs de référence pour ces sites.

Le rendement des modules est corrigé de leur échauffement, estimé depuis la NOCT et l'irradiation.

Écart signalé : la répartition horaire de l'irradiation, calculée au pas de 0,1 h dans le classeur,
n'est pas portée — seules les productions mensuelles et annuelle le sont, qui sont les grandeurs
utilisées par le bilan électrique.

## Bloc 4 — transverse

Ces modules ne calculent aucun procédé : ils agrègent ce que les trois files ont produit.
`src/core/simulation.js` orchestre désormais les quatre moteurs dans l'ordre du programme principal
d'OCEAN — la file eau alimente la file boues, les deux alimentent les utilités, et le transverse
agrège le tout. L'ordre compte : le bilan électrique a besoin des trois files, l'empreinte CO2 du
bilan électrique.

- **Bilan électrique** (`z_Bilan_electrique.cls`) : répartit la consommation sur dix-sept postes et
  distingue la **part fixe** (agitation, brassage, racleurs — ce qui tourne quelle que soit la
  charge) de la **part variable** (aération, pompages). Les nœuds portés exposaient déjà
  `electricity.fixed` ; la part variable en est le complément. Sur la filière de référence, la part
  fixe représente 43 % de la consommation, ce qui est le genre de chiffre qui oriente une stratégie
  d'exploitation.

  L'électricité verte peut être autoconsommée ou vendue, et le port suit le classeur sur ce point :
  autoconsommée elle **réduit** les émissions du site, vendue elle en **évite** ailleurs. Les deux
  ne s'additionnent pas dans le même total.

- **Empreinte CO2** (`z_Empreinte_CO2.cls`) : convertit électricité, réactifs et combustibles en
  émissions, transport compris, avec les facteurs d'émission du classeur (ADEME, ASTEE, Eco-Invent,
  Carbone 4) et leurs incertitudes propagées. Chaque réactif porte son facteur et sa source, toutes
  deux exposées dans le tableau de bord. Le méthanol est le seul à cumuler deux facteurs :
  521 kg CO2/t pour sa production et 1 375 pour sa combustion.

  Écarts signalés : le poste « autres » lu depuis une feuille du classeur devient un paramètre
  forçable ; le protoxyde d'azote et le méthane fugitif ne figurent pas dans la classe d'origine,
  l'empreinte ne couvre donc que le CO2 des consommations et pas les émissions directes du procédé ;
  le classeur écrase plusieurs fois `FE_reactifs_incertitude(H2O2)` en lieu et place de CuSO4 et
  MgCl2 — reproduit, sans effet puisque leurs facteurs sont nuls.

### Tableau de bord

`src/components/Dashboard.jsx` ajoute une vue « Bilan » à l'application, accessible depuis la barre
supérieure. Elle présente, dans l'ordre où on les lit en conception :

1. les indicateurs de synthèse — consommation, kWh/m³, kWh/(EH·an), autosuffisance, émissions ;
2. la répartition électrique par poste, en ruban empilé puis en barres horizontales, avec un code
   couleur par file (eau, boues, utilités) et le vert pour la production ;
3. la part fixe et la part variable ;
4. le **détail procédé par procédé** en tableau trié, avec fixe, variable, total, part et jauge ;
5. l'empreinte carbone par compartiment, puis réactif par réactif avec la source de chaque facteur ;
6. le contexte de la simulation — débit, pollution éliminée, biogaz, boues évacuées.

Tout est en SVG et CSS, sans dépendance graphique ajoutée.

### Correction trouvée par le contrôle de bouclage

Le test vérifie que la somme des postes égale le total des moteurs. Il a révélé que la production de
la turbine était comptée deux fois dans `runUtilities` : une fois comme total négatif, une fois via
le champ `produite`. Corrigé — l'écart est désormais nul.

### Volet économique

- **Coûts d'exploitation** (`Gestion_OPEX.cls`) : valorise électricité, réactifs, combustibles et
  évacuation des sous-produits, moins les recettes de vente d'électricité, de biogaz et de struvite.
  Le prix de l'électricité ne s'applique qu'à la part **achetée** : l'électricité verte
  autoconsommée ne coûte rien, et la part fixe comme la part variable sont valorisées au prorata.

  Sur la filière de référence, l'évacuation des boues pèse 538 €/j contre 467 pour l'électricité —
  le poste le plus lourd n'est pas celui qu'on attend, et c'est précisément ce que l'outil doit
  faire ressortir.

- **Retour sur investissement** (`Retour_investissement.cls`) : compare la filière simulée à une
  référence sur trois indicateurs — temps de retour simple, coût complet sur la durée du contrat, et
  valeur actuelle nette année par année. La convention du classeur mérite d'être explicitée : une
  variante ne « rembourse » que si elle coûte **moins cher à exploiter** que la référence ; si son
  OPEX est supérieur ou égal, aucun temps de retour n'est calculé quel que soit l'écart
  d'investissement.

  Écarts signalés : le classeur gère deux devises en parallèle avec un taux de change, le port n'en
  tient qu'une ; les prix par défaut sont lus depuis une feuille de saisie et non depuis le VBA, les
  valeurs retenues ici sont des ordres de grandeur européens, tous forçables ; la répartition des
  réactifs entre files est établie par procédé consommateur plutôt que par réactif, ce qui est plus
  juste quand un même produit sert aux deux files.

Le tableau de bord s'enrichit de deux sections : la répartition des coûts en ruban avec le détail
par réactif, et une **courbe SVG de valeur actuelle nette** année par année, dont le franchissement
de zéro marque le temps de retour actualisé.

## Bloc 4 terminé

Les quatre classes transverses sont portées, et les quatre moteurs sont chaînés dans l'ordre du
programme principal : file eau → file boues → utilités → transverse, chaque agrégat alimentant le
contexte du suivant (bilan électrique → empreinte CO2 et OPEX → retour sur investissement).

### Traitement des sulfures

`A1_Traitement_prev_sulfure.cls` et `A2_Traitement_curatif_sulfure.cls`, réunis en un seul nœud.
Les sulfures se forment dans le réseau en anaérobiose ; deux stratégies s'opposent :

- **Préventif** au nitrate de calcium, qui empêche leur formation en donnant aux bactéries
  dénitrifiantes un accepteur d'électrons plus favorable que le sulfate. Le dosage se compte en
  millilitres de solution par m³ et par heure de temps de séjour dans le réseau.
- **Curatif**, qui détruit les sulfures déjà formés. Les trois dosages viennent de la
  stœchiométrie : 2 g d'O2, 1,5 g de H2O2 ou 3,42 g de FeCl3 par gramme de soufre. Le rendement de
  transfert de l'oxygène dépend du mode d'injection — 99 % en oxygénateur dédié, mais 50 à 90 % en
  injection directe selon la longueur de conduite disponible pour la dissolution.

Le module calcule aussi le **taux de stripping** des sulfures résiduels, qui charge l'air vicié en
H2S. C'est la liaison quantitative entre les sulfures et la désodorisation : sur une eau à 8 mg/L,
traiter en curatif à l'oxygène fait tomber le H2S de l'air vicié de 7,0 à 4,4 mg/Nm³, la
consommation de chlore de 65 à 44 kg/j, et le tonnage total de réactifs de 99 à 85 t/an.

Écarts signalés : le calcul du stripping et le contrôle « la consigne de sortie dépasse la charge
entrante » sont **commentés** dans le classeur ; le premier est repris parce qu'il est la seule
liaison quantitative avec la désodorisation, le second est implémenté comme garde-fou. Par ailleurs
le VBA multiplie le produit commercial par la pureté pour obtenir la masse « kgj » puis la divise
par la même pureté pour le produit pur, ce qui inverse les deux grandeurs ; le port rétablit la
convention du reste du projet, où le pur est toujours inférieur au commercial.

### Analyse pincement — réécriture

Le module PINCH du classeur fait 8 400 lignes, dont l'essentiel est de l'affichage Excel : tracé de
graphiques, représentation du réseau d'échangeurs, gestion des couleurs. Le cœur algorithmique y est
mince. Plutôt que de le transposer, `src/core/pinch.js` implémente la méthode de Linnhoff et
Hindmarsh dans son ordre canonique — segmentation, décalage, table de problème, cascade, composites,
grande courbe composite, placement des utilités, cibles de conception.

Quatre points le distinguent du module d'origine :

1. **Contributions ΔT individuelles.** Un ΔT_min global suppose que tous les flux offrent le même
   coefficient d'échange, ce qui est faux : un gaz de combustion et une boue liquide n'ont pas le
   même film. Chaque flux porte ici sa contribution propre, et l'écart exigé entre deux flux est la
   somme des deux. Le cas test le vérifie : sur un même couple de flux, passer de 5 K + 5 K à
   20 K + 2 K fait apparaître 4 kW d'utilité chaude là où il n'en fallait aucune.
2. **Changements de phase segmentés** plutôt que linéarisés, ce qui évite de sous-estimer la
   récupération sur un condenseur.
3. **Pincements multiples détectés**, avec les quasi-pincements signalés : ne retenir que le premier
   conduit à des réseaux infaisables.
4. **Utilités multiples placées sur la grande courbe composite**, du moins cher au plus cher — le
   biogaz avant le gaz naturel, dans la limite de ce que la courbe autorise.

**Deux erreurs corrigées en cours d'écriture, l'une de fond.** D'abord, j'avais décalé la composite
chaude de l'utilité froide ; c'est la **froide** qu'il faut décaler, puisque l'extrémité gauche de
l'axe enthalpique porte la partie la plus froide de la composite chaude, celle que l'utilité froide
évacue. Ensuite, et c'est le point important : dès que les contributions ΔT diffèrent d'un flux à
l'autre, les composites tracées en températures **réelles** perdent leur sens, deux flux de
contributions 2 K et 20 K ne pouvant être agrégés sur une courbe dont on lirait un écart minimal
unique. Les composites sont donc construites dans l'échelle décalée, où l'écart minimal vaut zéro
par construction ; une version en températures réelles est fournie pour l'affichage, avec un drapeau
`exactes` qui indique si sa lecture usuelle reste valide.

Le test `test13.mjs` valide sur trois cas analytiques calculables à la main, puis sur 154 jeux de
flux aléatoires par trois invariants : bilan énergétique exact, cascade positive partout, et
composites jamais croisées — cette dernière étant la preuve que le ΔT est respecté sur toute la
plage. Les 154 passent.

Sur une filière Exelys + séchage + incinération, l'analyse identifie sept flux thermiques et cible
338 kW de récupération pour 504 kW de besoins, soit **67 % d'économie** sur l'utilité chaude, avec un
pincement à 105/95 °C et dix échangeurs au minimum.

### Reste à porter

Le classeur est couvert, à l'exception de l'interface Excel du module PINCH — construction assistée
du réseau d'échangeurs et rendu graphique — dont l'équivalent serait une fonctionnalité d'interface
plutôt qu'un portage.


## Tests

Des scénarios de non-régression headless sont fournis dans `tests/` (aucune dépendance,
exécutables avec Node depuis la racine du projet) :

```bash
node tests/test.mjs    # filière classique complète (C2→D1→E1→F3→F4→F5→G1→G2)
node tests/test2.mjs   # E3 / E4 / MBR
node tests/test3.mjs   # D2 → E3 (co-précipitation, plancher Pt)
node tests/test4.mjs   # D2 → E2 → F3 → G1 + contrôle NaN/négatifs
node tests/test5.mjs   # Biostyr C / N / NDNc / NDNs
node tests/test6.mjs   # Biostyr N + PDN (garantie NGL)
node tests/test7.mjs   # filière biofiltration 3 étages
node tests/test8.mjs   # filière biofiltration + décanteur d'eaux sales et rebouclage
node tests/test9.mjs   # MBBR (4 configurations) et HybAS (4 configurations)
node tests/test10.mjs  # file boues et retours : 26 filières, bilan matière
node tests/test11.mjs  # utilités : désodorisation, turbine, réactifs, bilan électrique global
node tests/test12.mjs  # transverse : bilan, CO2, OPEX, retour sur investissement
node tests/test13.mjs  # analyse pincement : cas analytiques et invariants sur 154 jeux
```
