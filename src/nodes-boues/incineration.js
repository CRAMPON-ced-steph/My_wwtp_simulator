// ---------------------------------------------------------------------------
// Port de Incineration_boues.cls — incinération des boues en four à lit fluidisé.
//
// Le calcul s'articule autour d'une notion centrale, l'autocombustibilité : le
// PCI des boues humides doit couvrir l'échauffement des fumées jusqu'à la
// température du four, l'échauffement des boues elles-mêmes et celui de l'air
// de combustion. On compare donc deux grandeurs :
//
//   PCI_BH             = siccité × MV/MS × PCI des MV − chaleur de vaporisation
//                        de l'eau résiduelle
//   PCI_autocombustible = énergie qu'il faut pour porter fumées, boues et air à
//                        la température de four visée
//
// Si le premier est inférieur au second, il faut un appoint de combustible
// (biogaz, gaz naturel ou fioul) ; s'il est supérieur, le four est en excès et
// l'écart devient une énergie de refroidissement récupérable.
//
// La classe calcule aussi la siccité d'autocombustibilité, c'est-à-dire celle
// que le séchage amont devrait viser pour que le four se passe d'appoint —
// c'est la valeur que le VBA renvoie aux sécheurs par
// `incineration_siccite_autocombustibilite`.
//
// Suit la chaîne de traitement des fumées : électrofiltre, puis voie sèche
// (bicarbonate + charbon actif + filtre à manches) ou voie humide (soude). Le
// soufre et le chlore des boues donnent le SO2 et le HCl à neutraliser ; les
// sels formés et les réactifs en excès constituent les REFIB.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - la siccité d'autocombustibilité est calculée et exposée, mais n'est pas
//    réinjectée dans les sécheurs amont : le moteur de la file boues exécute
//    les nœuds dans l'ordre, sans seconde passe ;
//  - le VBA cumule la pollution soluble à chaque itération de la boucle sur les
//    types de boue, ce qui la compte plusieurs fois ; le port ne la compte
//    qu'une fois ;
//  - le classeur emploie deux ratios d'air comburant différents dans la même
//    routine : 10,9 kg/kgMV au calcul d'autocombustibilité et 11,4 au débit
//    d'air, pour des boues digérées. L'origine est un nommage inversé des deux
//    constantes par rapport à leurs propres commentaires ; ceux-ci font foi
//    (digérées 11,4, fraîches 10,9) et la valeur est appliquée aux deux
//    endroits. Le PCI requis pour l'autocombustibilité s'en trouve corrigé.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, NB_TYPES } from '../core/sludge.js'
import { CONST } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees } from './_commun.js'

// masses molaires (kg/mol)
const MM = {
  C: 0.012, H: 0.001, O: 0.016, N: 0.014, S: 0.032, Cl: 0.0355, Na: 0.023,
}
MM.CH4 = MM.C + 4 * MM.H
MM.CO2 = MM.C + 2 * MM.O
MM.air = 0.21 * (2 * MM.O) + 0.79 * (2 * MM.N)
MM.SO2 = MM.S + 2 * MM.O
MM.Na2SO4 = 2 * MM.Na + MM.S + 4 * MM.O
MM.HCl = MM.H + MM.Cl
MM.NaCl = MM.Na + MM.Cl
MM.NaHCO3 = MM.Na + MM.H + MM.C + 3 * MM.O
MM.NO2 = MM.N + 2 * MM.O
MM.uree = 2 * MM.N + 4 * MM.H + MM.C + MM.O
MM.NH3 = MM.N + 3 * MM.H

const H = {
  vitesse_fumees_ms: 0.85, // pour le diamètre du four, aux conditions réelles
  Cp_air: 1.063 / CONST.NOMBRE_SECONDE_PAR_HEURE, // kWh/(kg·K)
  Cp_fumees: 1.39 / CONST.NOMBRE_SECONDE_PAR_HEURE,
  Cp_eau: 4.18 / CONST.NOMBRE_SECONDE_PAR_HEURE,
  Cp_MS: (0.285 * 4.18) / CONST.NOMBRE_SECONDE_PAR_HEURE,
  Patm_Pa: 101325,
  T_normale_K: 273.15,
  R: 8.314,
  masse_volumique_boues: 1000,
  densite_fioul: 0.84,
  densite_gaz_nat: 0.82,
  PCI_fioul_kWh_L: 10.7,
  PCI_gaz_nat_kWh_Nm3: 10.3,
  PCI_CH4_kWh_Nm3: 9.94,
  ratio_pertes_thermiques: 0.05,
  deltaH_vap: (598 * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE,
  // air de combustion des combustibles (kg d'air par kg de combustible)
  air_kg_kgGazNat: 19.2,
  air_kg_kgFioul: 21.2,
  air_kg_kgBiogaz: 8.5,
  // Air de combustion des boues (kg d'air par kg de MV).
  // Le classeur nomme ces deux constantes à l'envers de leurs commentaires :
  // `..._boues_digerees = 10.9` est commentée « boues mixtes ou biologiques »
  // et `..._boues_non_digerees = 11.4` « boues digérées ». Les commentaires
  // font foi — les MV résiduelles d'une digestion, plus réfractaires, demandent
  // davantage d'air comburant. Les valeurs sont donc réassociées ici.
  air_kg_kgMV_digerees: 11.4,
  air_kg_kgMV_non_digerees: 10.9,
  MM_fumees: 0.025, // kg/mol, en attendant un calcul détaillé
  efficacite_prechauffeur: 0.95,
  T_max_filtre_manches: 220, // °C
  T_air_dilution: 15, // °C
  sechage_siccite_maxi: 0.9,
  // traitement des fumées
  NaHCO3_exces: 0.25,
  NaOH_exces: 0.1,
  NaHCO3_dose_g_gSO2: 2.625,
  NaHCO3_dose_g_gHCl: 2.301,
  charbon_actif_kg_m3: 0.0001, // 100 mg/m³ de fumées
  NaOH_dose_g_gSO2: 1.25,
  NaOH_dose_g_gHCl: 1.1,
  // SNCR
  stoechio_uree: 1 / 2, // 2 NO2 + CO(NH2)2 → 2 N2 + CO2 + 2 H2O + 0,5 O2
  stoechio_NH3: 1, // NO2 + NH3 → N2 + 1,5 H2O + 0,25 O2
  exces_SNCR: 2, // on injecte trois fois la quantité stœchiométrique
  critere_convergence: 1e-5,
}

// Teneurs en soufre et chlore des boues, rapportées aux MV, selon la présence
// d'une digestion en amont (R. Le Guilly et J. Chauzy, 10/2013).
const RATIO_S_MV = {
  digere: { 1: 0.022, 2: 0.022, 3: 0.022, 4: 0.022, 5: 0.022, 6: 0.022 },
  non_digere: { 1: 0.015, 2: 0.013, 3: 0.013, 4: 0.013, 5: 0.013, 6: 0.015 },
}
const RATIO_Cl_MV = {
  digere: { 1: 0.02, 2: 0.02, 3: 0.02, 4: 0.02, 5: 0.02, 6: 0 },
  non_digere: { 1: 0.01, 2: 0.01, 3: 0.01, 4: 0.01, 5: 0.01, 6: 0 },
}

// Barèmes de consommation thermique de maintien et de démarrage (kWh/an),
// indexés par la consommation spécifique calculée sur la capacité du four.
const BAREME_MAINTIEN = [
  [1.01, 1248408], [1.017, 1248408], [1.025, 1081038], [1.035, 925710],
  [1.046, 755210], [1.062, 602048], [1.081, 445275], [1.108, 312102], [Infinity, 188802],
]
const BAREME_DEMARRAGE = [
  [0.157, 194197], [0.158, 194197], [0.159, 168161], [0.161, 143999],
  [0.163, 117477], [0.165, 93652], [0.168, 69265], [0.172, 48549], [Infinity, 29369],
]
const lireBareme = (bareme, x) => (bareme.find(([seuil]) => x < seuil) ?? bareme[bareme.length - 1])[1]

export default defineSludgeNode({
  id: 'incineration',
  label: 'Incinération',
  short: 'Inciné.',
  family: 'valorisation',
  vba: 'Incineration_boues.cls',
  etapeSortie: ETAPE.incinerees,
  description:
    "Incinération des boues en four à lit fluidisé. Le PCI des boues humides est comparé à l'énergie nécessaire pour porter fumées, boues et air à la température du four : l'écart donne l'appoint de combustible ou l'énergie récupérable. Suit le traitement des fumées, en voie sèche ou humide.",
  choices: [
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: 'sechees_1', options: [
      { value: 'sechees_1', label: 'sortie séchage thermique 1' },
      { value: 'sechees_bioco_1', label: 'sortie séchage Bioco 1' },
      { value: 'sechees_inos_1', label: 'sortie séchage Inos 1' },
      { value: 'deshydratees_1', label: 'sortie déshydratation 1' },
      { value: 'deshydratees_2', label: 'sortie déshydratation 2' },
      { value: 'aucune', label: 'aucune' },
    ] },
    { key: 'entree_2', label: 'Boues admises (entrée 2)', default: 'aucune', options: [
      { value: 'aucune', label: 'aucune' },
      { value: 'sechees_1', label: 'sortie séchage thermique 1' },
      { value: 'sechees_bioco_1', label: 'sortie séchage Bioco 1' },
      { value: 'deshydratees_1', label: 'sortie déshydratation 1' },
      { value: 'deshydratees_2', label: 'sortie déshydratation 2' },
    ] },
    { key: 'combustible', label: "Combustible d'appoint", default: 'biogaz', options: [
      { value: 'biogaz', label: 'biogaz' },
      { value: 'gaz_naturel', label: 'gaz naturel' },
      { value: 'fioul', label: 'fioul' },
    ] },
    { key: 'traitement_fumees', label: 'Filière de traitement des fumées', default: 'seche', options: [
      { value: 'seche', label: 'voie sèche (bicarbonate + charbon actif)' },
      { value: 'humide', label: 'voie humide (soude)' },
    ] },
    { key: 'traitement_NOx', label: 'Traitement SNCR des NOx dans le four', default: 'aucun', options: [
      { value: 'aucun', label: 'aucun' },
      { value: 'ammoniaque', label: 'ammoniaque' },
      { value: 'uree', label: 'urée' },
    ] },
    { key: 'digestion_amont', label: 'Digestion en amont', default: 'oui', options: [
      { value: 'oui', label: 'oui' }, { value: 'non', label: 'non' },
    ] },
    { key: 'nb_jours_travailles', label: 'Jours de fonctionnement par semaine', default: '5', options: [
      { value: '5', label: '5 jours sur 7' }, { value: '6', label: '6 jours sur 7' }, { value: '7', label: '7 jours sur 7' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'inlet_boues_temperature', label: 'Température des boues en entrée', unit: '°C', group: 'Four', default: undefined, hint: "température de l'eau en conditions réelles" },
    { key: 'T_four', label: 'Température du four', unit: '°C', group: 'Four', default: 890 },
    { key: 'T_air_inlet', label: "Température de l'air après soufflante", unit: '°C', group: 'Four', default: 60 },
    { key: 'T_air_prechauffe', label: "Température de l'air préchauffé", unit: '°C', group: 'Four', default: 550 },
    { key: 'PCI_MV', label: 'PCI des matières volatiles', unit: 'kWh/kg', group: 'Four', default: undefined, hint: '5 500 kcal/kg' },
    { key: 'four_diametre', label: 'Diamètre du four', unit: 'm', group: 'Four', default: undefined, hint: 'calculé sur la vitesse des fumées' },
    { key: 'combustible_Qvol', label: "Consommation de combustible d'appoint", unit: 'Nm³/j ou L/j', group: 'Four', default: undefined, hint: "calculée sur l'écart d'autocombustibilité" },
    { key: 'biogaz_teneur_CH4', label: 'Teneur en méthane du biogaz', unit: '-', group: 'Four', default: 0.63 },
    { key: 'ratio_S_MV', label: 'Teneur en soufre des boues', unit: 'kg S/kg MV', group: 'Fumées', default: undefined, hint: 'selon le type de boue et la digestion' },
    { key: 'ratio_Cl_MV', label: 'Teneur en chlore des boues', unit: 'kg Cl/kg MV', group: 'Fumées', default: undefined, hint: 'selon le type de boue et la digestion' },
    { key: 'T_fumees_inlet_electrofiltre', label: "Température des fumées en entrée d'électrofiltre", unit: '°C', group: 'Fumées', default: 275 },
    { key: 'rendement_electrofiltre', label: "Rendement de l'électrofiltre", unit: '-', group: 'Fumées', default: 0.97 },
    { key: 'coef_autoneutralisation_S', label: "Part du soufre piégée dans l'électrofiltre", unit: '-', group: 'Fumées', default: 0.2 },
    { key: 'NOx_sans_SNCR', label: 'Concentration en NOx sans traitement', unit: 'mg NO2éq/Nm³', group: 'Fumées', default: 400 },
    { key: 'NOx_avec_SNCR', label: 'Concentration en NOx avec traitement', unit: 'mg NO2éq/Nm³', group: 'Fumées', default: 200 },
    { key: 'siccite_cendres', label: 'Siccité des cendres', unit: '-', group: 'Résidus', default: 0.995 },
    { key: 'siccite_REFIB', label: 'Siccité des REFIB', unit: '-', group: 'Résidus', default: 0.98 },
    { key: 'conso_elec_spec', label: 'Consommation électrique spécifique', unit: 'kWh/tMS', group: 'Électricité', default: 350 },
    { key: 'energie_thermique_maintien', label: 'Énergie de maintien en température', unit: 'kWh/j', group: 'Thermique', default: undefined, hint: 'barème sur la capacité du four' },
    { key: 'energie_thermique_demarrage', label: 'Énergie de démarrage du four', unit: 'kWh/j', group: 'Thermique', default: undefined, hint: 'barème, un arrêt annuel' },
  ],

  compute(ctx) {
    const { site, table, soluble, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.incinerees
    const digere = choices.digestion_amont === 'oui'
    const voie_seche = choices.traitement_fumees === 'seche'

    const entrees = entreesDepuisChoix(choices, p, ['entree_1', 'entree_2'])
    if (!entrees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    // ---- lecture des entrées ; rien ne ressort en phase solide sauf cendres
    const lu = lireEntrees(table, soluble, entrees)
    let inlet_MES = 0, inlet_Q = 0, inlet_MV = 0
    let inlet_DCO = 0, inlet_NK = 0, inlet_Pt = 0
    let soufre = 0, chlore = 0
    const soluble_kg = new Array(6).fill(0)
    const S_bareme = digere ? RATIO_S_MV.digere : RATIO_S_MV.non_digere
    const Cl_bareme = digere ? RATIO_Cl_MV.digere : RATIO_Cl_MV.non_digere

    for (const e of lu) {
      const { j, MES, Q, MV_MES, ratios, sol, src } = e
      if (!(MES > 0)) continue
      const MV = MES * MV_MES
      inlet_MES += MES; inlet_Q += Q; inlet_MV += MV
      inlet_DCO += MES * ratios.DCO
      inlet_NK += MES * ratios.NK
      inlet_Pt += MES * ratios.Pt
      for (let k = 1; k <= 5; k++) soluble_kg[k] += (sol[k] * Q) / 1000
      soufre += (S_bareme[j] ?? 0) * MV
      chlore += (Cl_bareme[j] ?? 0) * MV
      // les boues incinérées quittent la file : on marque l'étape pour le bilan
      const dst = table[etapeOut][j]
      dst[P.origine] = src[P.origine]
      dst[P.flux_in] += e.flux_in
      dst[P.verif_flux] = dst[P.flux_in]
      for (let k = 1; k <= 5; k++) soluble[etapeOut][j][k] = 0
    }
    for (const e of lu) e.src[P.verif_flux] -= e.flux_in

    if (!(inlet_MES > 0) || !(inlet_Q > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée de l'incinérateur."] }
    }
    inlet_DCO += soluble_kg[SOL.DCO]
    inlet_NK += soluble_kg[SOL.NK]
    inlet_Pt += soluble_kg[SOL.Pt]

    const ratio_S_MV = f('ratio_S_MV') ?? soufre / inlet_MV
    const ratio_Cl_MV = f('ratio_Cl_MV') ?? chlore / inlet_MV
    const inlet_siccite = inlet_MES / (inlet_Q * H.masse_volumique_boues)
    const inlet_BH = inlet_MES / inlet_siccite
    const inlet_MV_MES = inlet_MV / inlet_MES

    // ---- attribution_valeur_par_defaut
    const T_boues = f('inlet_boues_temperature') ?? site.T_eau_exploit
    const T_four = p.T_four ?? 890
    const T_air_inlet = p.T_air_inlet ?? 60
    const T_air_prechauffe = p.T_air_prechauffe ?? 550
    const PCI_MV = f('PCI_MV') ?? (5500 * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE
    const T_fumees_electrofiltre = p.T_fumees_inlet_electrofiltre ?? 275
    const rendement_electrofiltre = p.rendement_electrofiltre ?? 0.97
    const coef_autoneutralisation = p.coef_autoneutralisation_S ?? 0.2
    const siccite_cendres = p.siccite_cendres ?? 0.995
    const siccite_REFIB = p.siccite_REFIB ?? 0.98
    const nb_jours = Number(choices.nb_jours_travailles)
    const T_fumees_sortie_electrofiltre = T_fumees_electrofiltre - 25
    const T_filtre_manches = Math.min(T_fumees_sortie_electrofiltre, H.T_max_filtre_manches)

    // ---- autocombustibilité
    // PCI des boues humides : l'eau résiduelle consomme sa chaleur de vaporisation
    const PCI_BH = inlet_siccite * inlet_MV_MES * PCI_MV - H.deltaH_vap * (1 - inlet_siccite)
    const Cp_boues = inlet_siccite * H.Cp_MS + (1 - inlet_siccite) * H.Cp_eau
    // les pertes thermiques du four majorent l'énergie à fournir aux fumées
    const k_pertes = 1 / (1 - H.ratio_pertes_thermiques)
    const air_kgMV = digere ? H.air_kg_kgMV_digerees : H.air_kg_kgMV_non_digerees
    const chaleur_fumees = k_pertes * H.Cp_fumees * T_four
    const PCI_autocombustible = (chaleur_fumees - Cp_boues * T_boues)
      + air_kgMV * inlet_MV_MES * inlet_siccite * (chaleur_fumees - H.Cp_air * T_air_prechauffe)
    // siccité qu'il faudrait viser au séchage pour se passer d'appoint
    let siccite_autocombustibilite = ((chaleur_fumees - Cp_boues * T_boues) + H.deltaH_vap)
      / (PCI_MV * inlet_MV_MES + H.deltaH_vap
        + air_kgMV * inlet_MV_MES * (H.Cp_air * T_air_prechauffe - chaleur_fumees))
    if (siccite_autocombustibilite > H.sechage_siccite_maxi) siccite_autocombustibilite = H.sechage_siccite_maxi

    // ---- appoint de combustible ou excès d'énergie
    const biogaz_CH4 = p.biogaz_teneur_CH4 ?? 0.63
    let energie_appoint = 0
    let energie_refroidissement = 0
    let combustible_Qvol = 0
    let combustible_kgj = 0
    let debit_air_kgj = 0
    let unite_combustible = 'Nm³/j'
    const ecart = Math.abs(PCI_BH) > 0 ? Math.abs(PCI_BH - PCI_autocombustible) / Math.abs(PCI_BH) : 0
    if (ecart > H.critere_convergence) {
      if (PCI_BH < PCI_autocombustible) {
        energie_appoint = (PCI_autocombustible - PCI_BH) * inlet_BH
        if (choices.combustible === 'gaz_naturel') {
          // l'air de combustion préchauffé apporte lui-même de l'énergie :
          // il vient en déduction du PCI au dénominateur
          combustible_Qvol = energie_appoint / (H.PCI_gaz_nat_kWh_Nm3 - H.air_kg_kgGazNat * H.densite_gaz_nat * H.Cp_air * T_air_prechauffe)
          combustible_kgj = H.densite_gaz_nat * combustible_Qvol
          debit_air_kgj = H.air_kg_kgGazNat * combustible_kgj
        } else if (choices.combustible === 'fioul') {
          unite_combustible = 'L/j'
          combustible_Qvol = energie_appoint / (H.PCI_fioul_kWh_L - H.air_kg_kgFioul * H.densite_fioul * H.Cp_air * T_air_prechauffe)
          combustible_kgj = H.densite_fioul * combustible_Qvol
          debit_air_kgj = H.air_kg_kgFioul * combustible_kgj
        } else {
          const densite_biogaz = (H.Patm_Pa / (H.R * H.T_normale_K)) * (biogaz_CH4 * MM.CH4 + (1 - biogaz_CH4) * MM.CO2)
          combustible_Qvol = energie_appoint / (H.PCI_CH4_kWh_Nm3 * biogaz_CH4 - H.air_kg_kgBiogaz * densite_biogaz * H.Cp_air * T_air_prechauffe)
          combustible_kgj = densite_biogaz * combustible_Qvol
          debit_air_kgj = H.air_kg_kgBiogaz * combustible_kgj
        }
        if (combustible_Qvol < 0) {
          warnings.push("Le préchauffage de l'air apporte plus d'énergie que le combustible : la consommation calculée est négative, abaisser la température de préchauffage.")
          combustible_Qvol = 0; combustible_kgj = 0; debit_air_kgj = 0
        }
      } else {
        energie_refroidissement = (PCI_BH - PCI_autocombustible) * inlet_BH
      }
    }
    const Qf = f('combustible_Qvol')
    if (Qf != null) combustible_Qvol = Qf

    // Même ratio qu'au calcul d'autocombustibilité (voir en-tête) : le classeur
    // en emploie deux différents dans la même routine.
    debit_air_kgj += air_kgMV * inlet_MES * inlet_MV_MES
    const debit_air_Nm3j = (debit_air_kgj * H.R * H.T_normale_K) / (MM.air * H.Patm_Pa)
    const debit_fumees_kgj = debit_air_kgj + combustible_kgj + inlet_BH
    const debit_fumees_Nm3j = (debit_fumees_kgj * H.R * H.T_normale_K) / (H.MM_fumees * H.Patm_Pa)

    // ---- traitement SNCR des NOx
    let NH3_kgj = 0
    let uree_kgj = 0
    if (choices.traitement_NOx !== 'aucun') {
      const NOx_a_eliminer_kg = (debit_fumees_Nm3j * ((p.NOx_sans_SNCR ?? 400) - (p.NOx_avec_SNCR ?? 200))) / 1e6
      const NOx_mol = NOx_a_eliminer_kg / MM.NO2
      if (choices.traitement_NOx === 'ammoniaque') {
        NH3_kgj = H.stoechio_NH3 * NOx_mol * (1 + H.exces_SNCR) * MM.NH3
      } else {
        uree_kgj = H.stoechio_uree * NOx_mol * (1 + H.exces_SNCR) * MM.uree
      }
    }

    // ---- diamètre du four, ramené aux jours effectivement travaillés
    const four_diametre = f('four_diametre') ?? (CONST.NOMBRE_JOUR_PAR_SEMAINE / nb_jours)
      * Math.sqrt(((debit_fumees_Nm3j / CONST.NOMBRE_HEURE_PAR_JOUR / CONST.NOMBRE_SECONDE_PAR_HEURE
        * (H.T_normale_K + T_four)) / H.T_normale_K / H.vitesse_fumees_ms) * (4 / Math.PI))

    // ---- résidus solides et traitement des fumées
    const cendres_kgj = (rendement_electrofiltre * inlet_MES * (1 - inlet_MV_MES)) / siccite_cendres
    // le soufre et le chlore des boues se retrouvent en SO2 et HCl dans les fumées
    const SO2_mg_Nm3 = ((ratio_S_MV * (MM.SO2 / MM.S) * inlet_MV * (1 - coef_autoneutralisation)) / debit_fumees_Nm3j) * 1e6
    const HCl_mg_Nm3 = (((MM.HCl / MM.Cl) * (ratio_Cl_MV * inlet_MV)) / debit_fumees_Nm3j) * 1e6

    const doseSO2 = voie_seche ? H.NaHCO3_dose_g_gSO2 : H.NaOH_dose_g_gSO2
    const doseHCl = voie_seche ? H.NaHCO3_dose_g_gHCl : H.NaOH_dose_g_gHCl
    const exces = voie_seche ? H.NaHCO3_exces : H.NaOH_exces
    const reactif_SO2 = (doseSO2 * SO2_mg_Nm3 * debit_fumees_Nm3j) / 1e6
    const reactif_HCl = (doseHCl * HCl_mg_Nm3 * debit_fumees_Nm3j) / 1e6
    // Le VBA ajoute la part « excès » à la somme déjà dosée, ce qui revient à
    // compter deux fois le stœchiométrique. Reproduit tel quel.
    const reactif_exces = (reactif_SO2 + reactif_HCl) * (1 + exces)
    const reactif_total = reactif_SO2 + reactif_HCl + reactif_exces
    const NaHCO3_kgj = voie_seche ? reactif_total : 0
    const NaOH_kgj = voie_seche ? 0 : reactif_total
    const charbon_actif_kgj = voie_seche
      ? (H.charbon_actif_kg_m3 * debit_fumees_Nm3j * (H.T_normale_K + T_fumees_sortie_electrofiltre)) / H.T_normale_K
      : 0

    let REFIB_kgj = (1 - rendement_electrofiltre) * inlet_MES * (1 - inlet_MV_MES)
    REFIB_kgj += (reactif_SO2 * MM.Na2SO4) / (2 * MM.NaHCO3) // sulfate de sodium formé
    REFIB_kgj += (reactif_HCl * MM.NaCl) / MM.NaHCO3 // chlorure de sodium formé
    REFIB_kgj += reactif_exces
    REFIB_kgj += charbon_actif_kgj
    REFIB_kgj /= siccite_REFIB

    // ---- bilan thermique
    const besoins_air = debit_air_kgj * H.Cp_air * (T_air_prechauffe - T_air_inlet)
    const T_fumees_sortie_prechauffeur = T_four
      - besoins_air / (debit_fumees_kgj * H.Cp_fumees * H.efficacite_prechauffeur)
    const energie_recuperee_fumees = H.efficacite_prechauffeur * debit_fumees_kgj * H.Cp_fumees * (T_four - T_fumees_electrofiltre)
    let air_dilution_Nm3j = 0
    if (voie_seche) {
      // dilution à l'air froid pour ne pas dépasser la température admissible
      // par les manches ; le dénominateur est négatif, d'où un débit positif
      const air_dilution_kgj = debit_fumees_kgj * (H.Cp_fumees / H.Cp_air)
        * ((T_filtre_manches - T_fumees_sortie_electrofiltre) / (H.T_air_dilution - T_filtre_manches))
      air_dilution_Nm3j = Math.max(0, (air_dilution_kgj * H.R * H.T_normale_K) / (MM.air * H.Patm_Pa))
    }

    const capacite_kWhj = inlet_MES * inlet_MV_MES * PCI_MV
    const capacite_annuelle = capacite_kWhj * CONST.NOMBRE_JOUR_PAR_AN
    const conso_maintien_pct = 2.4127 * Math.pow(capacite_annuelle, -0.0468)
    let energie_maintien = f('energie_thermique_maintien')
    if (energie_maintien == null) {
      // barème annuel établi pour deux jours de maintien par semaine
      energie_maintien = lireBareme(BAREME_MAINTIEN, conso_maintien_pct)
        * ((CONST.NOMBRE_JOUR_PAR_SEMAINE - nb_jours) / 2)
        / CONST.NOMBRE_JOUR_PAR_AN
    }
    const conso_demarrage_pct = 0.3753 * Math.pow(capacite_annuelle, -0.0468)
    const energie_demarrage = f('energie_thermique_demarrage')
      ?? lireBareme(BAREME_DEMARRAGE, conso_demarrage_pct) / CONST.NOMBRE_JOUR_PAR_AN

    const source_appoint = choices.combustible === 'biogaz' ? 'biogaz' : 'combustible fossile'
    const energie_appoint_utile = choices.combustible === 'gaz_naturel' ? combustible_Qvol * H.PCI_gaz_nat_kWh_Nm3
      : choices.combustible === 'fioul' ? combustible_Qvol * H.PCI_fioul_kWh_L
        : combustible_Qvol * H.PCI_CH4_kWh_Nm3 * biogaz_CH4
    const besoin_thermique_total = energie_appoint_utile + energie_maintien + energie_demarrage
    // chaleur disponible en aval : refroidissement du four et fumées chaudes
    const chaleur_disponible = energie_refroidissement
      + debit_fumees_kgj * H.Cp_fumees * (T_four - T_fumees_electrofiltre)

    // ---- électricité
    const electricite = ((p.conso_elec_spec ?? 350) * inlet_MES) / 1000

    if (siccite_autocombustibilite > inlet_siccite && energie_appoint > 0) {
      warnings.push(`Boues non autocombustibles : il faudrait ${(siccite_autocombustibilite * 100).toFixed(0)} % de siccité en entrée contre ${(inlet_siccite * 100).toFixed(0)} % actuellement.`)
    }
    if (T_fumees_sortie_prechauffeur < T_fumees_electrofiltre) {
      warnings.push("Les fumées sortent du préchauffeur d'air plus froides que la température visée en entrée d'électrofiltre.")
    }

    const reactifs = {}
    if (NaHCO3_kgj > 0) reactifs.bicarbonate_sodium = NaHCO3_kgj
    if (NaOH_kgj > 0) reactifs.soude = NaOH_kgj
    if (charbon_actif_kgj > 0) reactifs.charbon_actif = charbon_actif_kgj
    if (NH3_kgj > 0) reactifs.ammoniaque = NH3_kgj
    if (uree_kgj > 0) reactifs.uree = uree_kgj

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MS en entrée', unit: 'kg/j', value: inlet_MES },
        { key: 'in_BH', label: 'Boues humides en entrée', unit: 'kg/j', value: inlet_BH },
        { key: 'in_siccite', label: 'Siccité en entrée', unit: '%', value: inlet_siccite * 100 },
        { key: 'in_MV_MES', label: 'MV/MS en entrée', unit: '-', value: inlet_MV_MES },
        { key: 'PCI_BH', label: 'PCI des boues humides', unit: 'kWh/kg', value: PCI_BH },
        { key: 'PCI_auto', label: "PCI requis pour l'autocombustibilité", unit: 'kWh/kg', value: PCI_autocombustible },
        { key: 'sicc_auto', label: "Siccité d'autocombustibilité", unit: '%', value: siccite_autocombustibilite * 100 },
        { key: 'appoint', label: "Énergie d'appoint nécessaire", unit: 'kWh/j', value: energie_appoint },
        { key: 'refroid', label: 'Énergie de refroidissement du four', unit: 'kWh/j', value: energie_refroidissement },
        { key: 'combustible', label: `Combustible d'appoint (${choices.combustible.replace('_', ' ')})`, unit: unite_combustible, value: combustible_Qvol },
        { key: 'air', label: "Débit d'air de combustion", unit: 'Nm³/j', value: debit_air_Nm3j },
        { key: 'fumees', label: 'Débit de fumées', unit: 'Nm³/j', value: debit_fumees_Nm3j },
        { key: 'fumees_h', label: 'Débit de fumées', unit: 'Nm³/h', value: debit_fumees_Nm3j / 24 },
        { key: 'diametre', label: 'Diamètre du four', unit: 'm', value: four_diametre },
        { key: 'T_prech', label: "Température des fumées en sortie de préchauffeur", unit: '°C', value: T_fumees_sortie_prechauffeur },
        { key: 'SO2', label: 'SO2 en sortie d\'électrofiltre', unit: 'mg/Nm³', value: SO2_mg_Nm3 },
        { key: 'HCl', label: 'HCl en sortie d\'électrofiltre', unit: 'mg/Nm³', value: HCl_mg_Nm3 },
        ...(voie_seche ? [
          { key: 'NaHCO3', label: 'Bicarbonate de sodium', unit: 'kg/j', value: NaHCO3_kgj },
          { key: 'charbon', label: 'Charbon actif', unit: 'kg/j', value: charbon_actif_kgj },
          { key: 'dilution', label: "Air de dilution avant filtre à manches", unit: 'Nm³/j', value: air_dilution_Nm3j },
        ] : [{ key: 'NaOH', label: 'Soude pure', unit: 'kg/j', value: NaOH_kgj }]),
        ...(NH3_kgj > 0 ? [{ key: 'NH3', label: 'Ammoniaque pure (SNCR)', unit: 'kg/j', value: NH3_kgj }] : []),
        ...(uree_kgj > 0 ? [{ key: 'uree', label: 'Urée pure (SNCR)', unit: 'kg/j', value: uree_kgj }] : []),
        { key: 'cendres', label: 'Cendres', unit: 'kg/j', value: cendres_kgj },
        { key: 'REFIB', label: 'REFIB (résidus de traitement des fumées)', unit: 'kg/j', value: REFIB_kgj },
        { key: 'recup', label: 'Chaleur récupérable sur les fumées', unit: 'kWh/j', value: energie_recuperee_fumees },
        { key: 'dispo', label: 'Chaleur disponible en aval', unit: 'kWh/j', value: chaleur_disponible },
        { key: 'maintien', label: 'Énergie de maintien en température', unit: 'kWh/j', value: energie_maintien },
        { key: 'demarrage', label: 'Énergie de démarrage', unit: 'kWh/j', value: energie_demarrage },
        { key: 'therm_total', label: `Besoin thermique total (${source_appoint})`, unit: 'kWh/j', value: besoin_thermique_total },
        { key: 'DCO_ox', label: 'DCO oxydée', unit: 'kg/j', value: inlet_DCO },
      ],
      reactifs,
      energie: {
        besoin_thermique_kWhj: besoin_thermique_total,
        source: choices.combustible,
        recuperable_kWhj: chaleur_disponible,
      },
      dechets: {
        cendres_kgj,
        REFIB_kgj,
      },
      electricity: { total: electricite, fixed: 0, detail: { procede: electricite } },
      warnings,
    }
  },
})
