// ---------------------------------------------------------------------------
// Port de A1_Traitement_prev_sulfure.cls et A2_Traitement_curatif_sulfure.cls.
//
// Les sulfures se forment dans le réseau d'assainissement, en anaérobiose, dès
// que le temps de séjour dépasse quelques heures. Ils corrodent les ouvrages,
// incommodent le voisinage et chargent l'air vicié en H2S — c'est cette
// dernière conséquence qui relie ces deux modules à la désodorisation.
//
// Deux stratégies, portées ici dans un même nœud :
//
//   Préventif  on empêche la formation des sulfures en injectant du nitrate de
//              calcium dans le réseau : les bactéries sulfato-réductrices sont
//              concurrencées par les dénitrifiantes, qui disposent alors d'un
//              accepteur d'électrons plus favorable. Le dosage se compte en
//              millilitres de solution commerciale par m³ et par heure de temps
//              de séjour dans le réseau.
//
//   Curatif    on détruit les sulfures déjà formés, par oxygène pur, peroxyde
//              d'hydrogène ou chlorure ferrique. Les trois dosages viennent
//              directement de la stœchiométrie des réactions, ce qui les rend
//              faciles à vérifier :
//                  O2    2 g/g de S     oxydation en soufre élémentaire
//                  H2O2  1,5 g/g        HS⁻ + H2O2 + H⁺ → ⅛ S8 + 2 H2O
//                  FeCl3 3,42 g/g       2 FeCl3 + 3 HS⁻ → 2 FeS + S + 6 Cl⁻ + 3 H⁺
//
// Le rendement de transfert de l'oxygène dépend du mode d'injection : 99 % dans
// un oxygénateur dédié, mais seulement 50 à 90 % en injection directe selon la
// longueur de conduite disponible pour la dissolution.
//
// Le module calcule aussi le **taux de stripping** des sulfures résiduels, qui
// détermine la charge en H2S de l'air vicié et donc les réactifs de la
// désodorisation chimique. Cette corrélation figure en commentaire dans le
// classeur, désactivée mais complète ; elle est portée et signalée.
//
// Écart au VBA, volontaire et signalé (voir README) : le calcul du stripping et
// le contrôle « la consigne de sortie dépasse la charge entrante » sont
// commentés dans le classeur. Le premier est repris parce qu'il est la seule
// liaison quantitative entre les sulfures et la désodorisation ; le second est
// implémenté comme un simple garde-fou.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'

const H = {
  // préventif : nitrate de calcium
  dose_CaNO3_mL_m3h_defaut: 15,
  reseau_tps_retention_h_defaut: 10,
  masse_volumique_CaNO3_commercial: 1420, // kg/m³, source Yara (Nutriox)
  purete_CaNO3: 0.45,
  // curatif : dosages stœchiométriques, g de réactif par g de S-HS⁻
  dose_O2: 2,
  dose_H2O2: 1.5,
  dose_FeCl3: 3.42,
  curatif_outlet_defaut: 5, // mg/L
  // rendement de transfert de l'oxygène
  rdt_O2_oxygenateur: 0.99,
  longueur_conduite_defaut: 900, // m
  longueur_min: 100,
  longueur_max: 900,
  rdt_O2_min: 0.5,
  rdt_O2_max: 0.9,
  // taux de stripping des sulfures résiduels, corrélation linéaire en T
  a_rate_stripping: [-0.046, 0.0069],
}

export const traitementSulfures = defineUtilityNode({
  id: 'traitement-sulfures',
  label: 'Traitement des sulfures',
  short: 'Sulfures',
  family: 'reactifs',
  vba: 'A1_Traitement_prev_sulfure.cls, A2_Traitement_curatif_sulfure.cls',
  description:
    "Traitement des sulfures formés dans le réseau : préventif au nitrate de calcium, curatif à l'oxygène, au peroxyde d'hydrogène ou au chlorure ferrique. Le taux de stripping des sulfures résiduels détermine la charge en H2S de l'air vicié.",
  choices: [
    { key: 'preventif', label: 'Traitement préventif', default: 'non', options: [
      { value: 'non', label: 'non' },
      { value: 'oui', label: 'oui, au nitrate de calcium' },
    ] },
    { key: 'curatif', label: 'Traitement curatif', default: 'non', options: [
      { value: 'non', label: 'non' },
      { value: 'O2', label: "injection d'oxygène pur" },
      { value: 'H2O2', label: "peroxyde d'hydrogène" },
      { value: 'FeCl3', label: 'chlorure ferrique' },
    ] },
    { key: 'oxygenateur', label: "Mode d'injection de l'oxygène", default: 'oxygenateur', options: [
      { value: 'oxygenateur', label: 'oxygénateur dédié' },
      { value: 'reseau', label: 'injection directe dans le réseau' },
    ] },
  ],
  params: [
    { key: 'HS_nominal_mgL', label: 'Sulfures en entrée de station', unit: 'mg/L', group: 'Charge entrante', default: undefined, hint: 'valeur du site si non forcée' },
    { key: 'Q_nominal', label: 'Débit nominal', unit: 'm³/j', group: 'Charge entrante', default: undefined, hint: 'débit du site si non forcé' },
    { key: 'temperature', label: "Température de l'eau", unit: '°C', group: 'Charge entrante', default: undefined, hint: 'température de dimensionnement du site' },
    { key: 'dose_CaNO3', label: 'Dosage de nitrate de calcium', unit: 'mL/(m³·h)', group: 'Préventif', default: 15 },
    { key: 'reseau_tps_retention', label: 'Temps de rétention dans le réseau', unit: 'h', group: 'Préventif', default: 10 },
    { key: 'curatif_outlet_HS', label: 'Sulfures visés en sortie de traitement', unit: 'mg/L', group: 'Curatif', default: 5 },
    { key: 'longueur_conduite', label: 'Longueur de conduite après injection', unit: 'm', group: 'Curatif', default: 900, hint: 'dissolution de l\'oxygène en injection directe' },
    { key: 'rdt_transfert_O2', label: "Rendement de transfert de l'oxygène", unit: '-', group: 'Curatif', default: undefined, hint: "99 % en oxygénateur, 50 à 90 % en réseau" },
    { key: 'taux_stripping', label: 'Taux de stripping des sulfures résiduels', unit: '-', group: 'Stripping', default: undefined, hint: '−0,046 + 0,0069 × T' },
  ],

  compute(ctx) {
    const { site, contexte, p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)

    const Q = f('Q_nominal') ?? contexte.Q_nominal ?? site.Q_nominal
    const HS_mgL = f('HS_nominal_mgL') ?? site.HS_nominal_mgL ?? 0
    const T = f('temperature') ?? site.T_eau_design
    const HS_nominal = (HS_mgL * Q) / 1000 // kg/j

    if (!(HS_nominal > 0)) {
      return {
        results: [
          { key: 'HS_in', label: 'Sulfures en entrée', unit: 'kg S/j', value: 0 },
        ],
        electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Aucun sulfure déclaré en entrée de station : renseigner la teneur de l'eau brute pour dimensionner un traitement."],
      }
    }

    // ---- traitement préventif au nitrate de calcium
    // Le dosage s'exprime en millilitres de solution commerciale par m³ d'eau
    // et par heure de temps de séjour dans le réseau : plus le réseau est long,
    // plus il faut de nitrate pour tenir les sulfato-réductrices en échec.
    let CaNO3_commercial = 0
    let CaNO3_pur = 0
    let dose_CaNO3 = 0
    let tps_retention = 0
    if (choices.preventif === 'oui') {
      dose_CaNO3 = p.dose_CaNO3 ?? H.dose_CaNO3_mL_m3h_defaut
      tps_retention = p.reseau_tps_retention ?? H.reseau_tps_retention_h_defaut
      const volume_m3j = (dose_CaNO3 / 1e6) * tps_retention * Q
      CaNO3_commercial = volume_m3j * H.masse_volumique_CaNO3_commercial
      // Le VBA multiplie le produit commercial par la pureté pour obtenir la
      // masse « kgj », puis la divise par la même pureté pour le produit pur :
      // les deux grandeurs sont donc dans le rapport inverse de l'usage
      // habituel. Le port rétablit la convention du reste du projet, où le pur
      // est toujours inférieur au commercial.
      CaNO3_pur = CaNO3_commercial * H.purete_CaNO3
    }

    // ---- traitement curatif
    let HS_removed = 0
    let outlet_HS_mgL = HS_mgL
    let dose_reactif = 0
    let reactif_pur = 0
    let rdt_O2 = 0
    let longueur = 0
    const reactifs = {}
    if (choices.curatif !== 'non') {
      outlet_HS_mgL = p.curatif_outlet_HS ?? H.curatif_outlet_defaut
      if (outlet_HS_mgL >= HS_mgL) {
        // la consigne est déjà tenue sans traitement (contrôle commenté dans le VBA)
        outlet_HS_mgL = HS_mgL
        warnings.push(`La teneur visée en sortie (${(p.curatif_outlet_HS ?? H.curatif_outlet_defaut).toFixed(1)} mg/L) est déjà atteinte en entrée : aucun traitement curatif nécessaire.`)
      } else {
        HS_removed = HS_nominal - (outlet_HS_mgL * Q) / 1000
        if (choices.curatif === 'O2') {
          if (choices.oxygenateur === 'oxygenateur') {
            rdt_O2 = f('rdt_transfert_O2') ?? H.rdt_O2_oxygenateur
          } else {
            longueur = p.longueur_conduite ?? H.longueur_conduite_defaut
            // la dissolution progresse avec la longueur de conduite disponible
            let rdt
            if (longueur < H.longueur_min) rdt = H.rdt_O2_min
            else if (longueur > H.longueur_max) rdt = H.rdt_O2_max
            else {
              rdt = H.rdt_O2_min
                + ((longueur - H.longueur_min) / (H.longueur_max - H.longueur_min)) * (H.rdt_O2_max - H.rdt_O2_min)
            }
            rdt_O2 = f('rdt_transfert_O2') ?? rdt
            if (longueur < H.longueur_min) {
              warnings.push(`Conduite courte (${longueur} m) : le rendement de dissolution de l'oxygène est plafonné à ${(H.rdt_O2_min * 100).toFixed(0)} %.`)
            }
          }
          dose_reactif = H.dose_O2 / rdt_O2
          reactif_pur = dose_reactif * HS_removed
          reactifs.oxygene_liquide = reactif_pur
        } else if (choices.curatif === 'H2O2') {
          dose_reactif = H.dose_H2O2
          reactif_pur = dose_reactif * HS_removed
          reactifs.H2O2 = reactif_pur
        } else {
          dose_reactif = H.dose_FeCl3
          reactif_pur = dose_reactif * HS_removed
          reactifs.FeCl3 = reactif_pur
        }
      }
    }
    if (CaNO3_pur > 0) reactifs.Ca_2NO3 = CaNO3_pur

    // ---- stripping des sulfures résiduels
    // C'est cette grandeur qui relie le module à la désodorisation : la part
    // des sulfures qui passe en phase gazeuse charge l'air vicié en H2S.
    let taux_stripping = f('taux_stripping')
    if (taux_stripping == null) {
      taux_stripping = H.a_rate_stripping[0] + H.a_rate_stripping[1] * T
      if (taux_stripping < 0) taux_stripping = 0
      if (taux_stripping > 1) taux_stripping = 1
    }
    const HS_residuel = Math.max(0, HS_nominal - HS_removed)
    const HS_strippe = taux_stripping * HS_residuel

    if (choices.preventif === 'non' && choices.curatif === 'non') {
      warnings.push(`Aucun traitement retenu : les ${HS_nominal.toFixed(1)} kg S/j entrants restent à traiter en aval, dont ${HS_strippe.toFixed(1)} kg/j passeront en phase gazeuse.`)
    }

    return {
      results: [
        { key: 'HS_in_mgL', label: 'Sulfures en entrée', unit: 'mg/L', value: HS_mgL },
        { key: 'HS_in', label: 'Charge en sulfures entrante', unit: 'kg S/j', value: HS_nominal },
        ...(choices.preventif === 'oui' ? [
          { key: 'dose_CaNO3', label: 'Dosage de nitrate de calcium', unit: 'mL/(m³·h)', value: dose_CaNO3 },
          { key: 'tps_reseau', label: 'Temps de rétention dans le réseau', unit: 'h', value: tps_retention },
          { key: 'CaNO3_com', label: 'Nitrate de calcium commercial', unit: 'kg/j', value: CaNO3_commercial },
          { key: 'CaNO3_pur', label: 'Nitrate de calcium pur', unit: 'kg/j', value: CaNO3_pur },
        ] : []),
        ...(choices.curatif !== 'non' ? [
          { key: 'HS_out_mgL', label: 'Sulfures en sortie de traitement', unit: 'mg/L', value: outlet_HS_mgL },
          { key: 'HS_removed', label: 'Sulfures éliminés', unit: 'kg S/j', value: HS_removed },
          { key: 'rdt', label: "Rendement d'élimination", unit: '-', value: HS_nominal > 0 ? HS_removed / HS_nominal : 0 },
          ...(choices.curatif === 'O2' ? [
            ...(choices.oxygenateur === 'reseau' ? [{ key: 'longueur', label: 'Longueur de conduite après injection', unit: 'm', value: longueur }] : []),
            { key: 'rdt_O2', label: "Rendement de transfert de l'oxygène", unit: '-', value: rdt_O2 },
          ] : []),
          { key: 'dose', label: 'Dosage du réactif', unit: 'g/g de S', value: dose_reactif },
          { key: 'reactif', label: 'Réactif pur consommé', unit: 'kg/j', value: reactif_pur },
        ] : []),
        { key: 'taux_stripping', label: 'Taux de stripping des sulfures résiduels', unit: '-', value: taux_stripping },
        { key: 'HS_residuel', label: 'Sulfures résiduels', unit: 'kg S/j', value: HS_residuel },
        { key: 'HS_strippe', label: "Sulfures strippés vers l'air vicié", unit: 'kg S/j', value: HS_strippe },
      ],
      reactifs,
      // la charge strippée alimente la désodorisation chimique placée en aval
      sulfures: { HS_nominal, HS_removed, HS_residuel, HS_strippe, taux_stripping },
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})
