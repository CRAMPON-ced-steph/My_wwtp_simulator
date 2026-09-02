// ---------------------------------------------------------------------------
// Port de z_CLS_BouesAthos.cls — Athos, oxydation par voie humide.
//
// Les boues sont portées à 250 °C sous 55 bars et oxydées à l'oxygène liquide.
// Le procédé se lit comme une boucle thermique fermée :
//
//   entrée → épaississement → mélange avec la recirculation de technosable
//          → échangeur 1 (préchauffage) → réacteur (250 °C, O2 pur)
//          → échangeur 2 (récupération) → échangeur 3 → décanteur
//          → technosable, dont la moitié du volume est recirculée en tête
//
// Deux calculs imbriqués gouvernent tout le reste :
//
//  1. Un **bilan d'énergie** de cinq équations à cinq inconnues (températures
//     d'entrée et de sortie de l'échangeur 1, température du réacteur, énergie
//     échangée entre les deux échangeurs, température de sortie de l'échangeur
//     2), résolu par élimination de Gauss. C'est lui qui donne la température
//     réellement atteinte par le réacteur.
//
//  2. Une **dichotomie sur la siccité des boues épaissies** : on cherche la
//     siccité, entre 4 et 10 %, pour laquelle le réacteur atteint exactement
//     250 °C sans apport extérieur. Si les boues brutes suffisent, aucun
//     épaississement n'est nécessaire ; sinon on épaissit juste ce qu'il faut.
//
// Le bilan de masse est lui-même itératif, la recirculation de technosable
// bouclant sur le mélange d'entrée.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - le VBA tient trois vecteurs de retour distincts (général, digestion,
//    Athos) pour permettre un traitement séparé des retours azotés ; le port
//    n'en tient qu'un et expose la part Athos en résultat ;
//  - la boucle du bilan de masse et celle de la dichotomie ne sont pas bornées
//    dans le VBA : des gardes sont ajoutées ;
//  - la sortie technosable alimente l'étape `boues_athos`, mais le circuit
//    « technosable » de la déshydratation par filtre à plateaux (siccité
//    600 g/L, retour dédié) n'est pas encore branché ;
//  - le VBA n'affecte jamais les MES du surnageant du décanteur Athos : les MS
//    qui échappent à la décantation ne figurent dans aucun bilan. Reproduit,
//    mais le flux manquant est calculé, exposé et signalé.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, RET, NB_TYPES, RET_ORIGINE, ajouterRetour } from '../core/sludge.js'
import { CONST } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees } from './_commun.js'

// masses molaires (g/mol)
const MM_CO2 = 1000 * (0.012 + 2 * 0.016)
const MM_O2 = 1000 * (2 * 0.016)
const MM_H2O = 1000 * (2 * 0.001 + 0.016)

/** capacité calorifique en kWh/(kg·°C) à partir d'une valeur en kJ/(mol·K) */
const CpDe = (kJ_mol_K, MM_g_mol) => (kJ_mol_K / MM_g_mol) * 1000 / CONST.NOMBRE_SECONDE_PAR_HEURE

const H = {
  masse_volumique_boues_froides: 1000, // kg/m³
  masse_volumique_boues_250: 800, // kg/m³ — les boues se dilatent à 250 °C
  masse_volumique_technosable: 1000,
  concentration_polymere_gL: 1.5,
  capture_MES_epaississement: 0.9,
  // capacités calorifiques, d'après Perry pour l'eau liquide à 5 MPa
  Cp_boues_froides: CpDe(0.07507, MM_H2O), // 300 K
  Cp_boues_110: CpDe(0.0762, MM_H2O), // 400 K
  Cp_boues_150: CpDe(0.078, MM_H2O),
  Cp_boues_250: CpDe(0.0872, MM_H2O), // 500 K
  Cp_O2_15: CpDe(0.032003, MM_O2),
  Cp_O2_250: CpDe(0.031815, MM_O2),
  Cp_CO2_250: CpDe(0.0486, MM_CO2),
  // enthalpie totale de la vapeur présente dans les fumées, référence eau à 0 °C
  enthalpie_vapeur_250: (669 * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE,
  PCI_MV: (5500 * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE,
  siccite_min_epaissies: 0.04,
  siccite_max_epaissies: 0.1,
  taux_recirculation_technosable: 0.5, // en volume, rapporté à l'entrée
  ratio_MV_particulaire_oxydees: 0.5,
  reduction_MV: 0.9,
  reduction_DCO: 0.8,
  T_reacteur_cible: 250, // °C
  ecart_T_echangeur2: 140, // °C entre réacteur et sortie d'échangeur froid
  T_technosable_sortie: 40, // °C
  T_oxygene_entree: 15, // °C
  efficacite_echangeurs: 0.9,
  ratio_O2_DCO: 1, // kg d'O2 par kg de DCO
  // 1 h de temps de séjour, majoré car Athos ne fonctionne que 7 500 h/an
  temps_sejour_h: 1 / (7500 / 8760),
  conso_elec_kWh_m3: 20, // rapportée aux boues épaissies, avant recirculation
  ratio_NH4_NK_soluble_digestion: 0.95,
  // retours du décanteur et du filtre à plateaux, en part de l'entrée
  retours_DCO_soluble_sur_totale: 0.95,
  retours_NK_sur_entree: 0.92,
  retours_NK_soluble_sur_total: 0.92,
  retours_NH4_sur_NK_avec_digestion: 0.85,
  retours_NH4_sur_NK_sans_digestion: 0.6,
  retours_Pt_sur_entree_avec_digestion: 0.08,
  retours_Pt_sur_entree_sans_digestion: 0.12,
  retours_Pt_soluble_sur_total: 0.06,
  // RTO : 64 Nm³/h de gaz naturel à 10 kWh/Nm³ pour 0,8 tMV/h (Bruxelles)
  conso_RTO_kWhPCI_tMV: (64 * 10) / 0.8,
  pression_reacteur_bars: 55,
  critere_siccite: 1e-8,
  critere_bilan_masse: 1e-3,
}
// capacité calorifique des fumées, moyenne pondérée O2 / CO2 selon le rendement
H.Cp_fumees = (H.reduction_DCO * H.Cp_O2_250
  + (1 - H.reduction_DCO) * (MM_CO2 / MM_O2) * H.Cp_CO2_250)
  / (H.reduction_DCO + (1 - H.reduction_DCO) * (MM_CO2 / MM_O2))

/**
 * Résout un système linéaire n×n par élimination de Gauss avec pivot partiel.
 * `a` est la matrice augmentée [n][n+1] ; renvoie le vecteur solution.
 */
function resoudreGauss(a, n) {
  for (let i = 0; i < n; i++) {
    // pivot partiel : on prend la plus grande valeur absolue de la colonne
    let pivot = i
    for (let k = i + 1; k < n; k++) if (Math.abs(a[k][i]) > Math.abs(a[pivot][i])) pivot = k
    if (pivot !== i) { const t = a[i]; a[i] = a[pivot]; a[pivot] = t }
    if (Math.abs(a[i][i]) < 1e-300) return null
    for (let k = i + 1; k < n; k++) {
      const c = a[k][i] / a[i][i]
      for (let j = i; j <= n; j++) a[k][j] -= c * a[i][j]
    }
  }
  const x = new Array(n).fill(0)
  for (let i = n - 1; i >= 0; i--) {
    let s = a[i][n]
    for (let j = i + 1; j < n; j++) s -= a[i][j] * x[j]
    x[i] = s / a[i][i]
  }
  return x
}

/**
 * Pression de vapeur saturante de l'eau (bars), polynôme de degré 6 calé sur
 * les données thermodynamiques du classeur, valable de 0 à 373 °C
 * (Pv_sat_H2O_bar de MOD_FonctionsPubliques).
 */
function PvSatH2O_bar(T_C) {
  const t = T_C < 0 ? 0 : T_C > 373 ? 373 : T_C
  return 9.481e-14 * Math.pow(t, 6)
    - 9.203e-11 * Math.pow(t, 5)
    + 4.801e-8 * Math.pow(t, 4)
    - 7.31e-6 * Math.pow(t, 3)
    + 5.988e-4 * t * t
    - 0.01745 * t
    + 0.136
}

/**
 * Humidité absolue d'un gaz (kg d'eau par kg de gaz sec), à la pression et à la
 * température données (humidite_air_gH2O_gAS de MOD_FonctionsPubliques).
 */
function humiditeSaturation(MM_gaz_g_mol, P_bars, T_C, humidite_relative = 1) {
  const Pv = humidite_relative * PvSatH2O_bar(T_C)
  const h = (MM_H2O / MM_gaz_g_mol) * (Pv / (P_bars - Pv))
  return h > 0 ? h : 0
}

export default defineSludgeNode({
  id: 'athos',
  label: 'Athos',
  short: 'Athos',
  family: 'valorisation',
  vba: 'z_CLS_BouesAthos.cls',
  etapeSortie: ETAPE.athos,
  description:
    "Athos : oxydation des boues par voie humide, à 250 °C sous 55 bars, à l'oxygène liquide. La siccité d'alimentation est ajustée par dichotomie pour que la réaction soit autotherme ; le résidu minéral (technosable) est décanté et partiellement recirculé.",
  choices: [
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: 'epaississeur_1', options: [
      { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
      { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
      { value: 'digerees', label: 'boues digérées' },
      { value: 'toutes', label: 'boues extraites, toutes origines' },
      { value: 'I', label: 'boues primaires brutes' },
      { value: 'II', label: 'boues secondaires brutes' },
      { value: 'aucune', label: 'aucune' },
    ] },
    { key: 'entree_2', label: 'Boues admises (entrée 2)', default: 'aucune', options: [
      { value: 'aucune', label: 'aucune' },
      { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
      { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
      { value: 'digerees', label: 'boues digérées' },
      { value: 'I', label: 'boues primaires brutes' },
      { value: 'II', label: 'boues secondaires brutes' },
    ] },
    { key: 'combustible_RTO', label: 'Combustible du traitement des fumées RTO', default: 'biogaz', options: [
      { value: 'biogaz', label: 'biogaz' },
      { value: 'externe', label: 'combustible externe' },
    ] },
    { key: 'digestion_amont', label: 'Type de digestion en amont', default: 'aucune', options: [
      { value: 'aucune', label: 'pas de digestion' },
      { value: 'simple', label: 'digestion simple' },
      { value: 'avancee', label: 'digestion avancée' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'temperature_entree', label: 'Température des boues en entrée', unit: '°C', group: 'Alimentation', default: undefined, hint: "température de l'eau de dimensionnement" },
    { key: 'capture_MES_epaississement', label: "Taux de capture de l'épaississement", unit: '-', group: 'Épaississement', default: 0.9 },
    { key: 'dose_polymere', label: 'Dosage de polymère', unit: 'kg/tMS', group: 'Épaississement', default: 6 },
    { key: 'siccite_boues_epaissies', label: 'Siccité des boues épaissies', unit: 'g/L', group: 'Épaississement', default: undefined, hint: 'ajustée par dichotomie pour être autotherme' },
    { key: 'reacteur_volume', label: 'Volume du réacteur', unit: 'm³', group: 'Réacteur', default: undefined, hint: "temps de séjour d'une heure" },
    { key: 'reduction_MV', label: 'Taux de réduction des MV', unit: '-', group: 'Réacteur', default: 0.9 },
    { key: 'reduction_DCO', label: 'Taux de réduction de la DCO', unit: '-', group: 'Réacteur', default: 0.8 },
    { key: 'T_reacteur_cible', label: 'Température visée dans le réacteur', unit: '°C', group: 'Réacteur', default: 250 },
    { key: 'consommation_O2', label: "Consommation d'oxygène liquide pur", unit: 'kg/j', group: 'Réacteur', default: undefined, hint: '1 kg par kg de DCO' },
    { key: 'consommation_CuSO4', label: 'Consommation de sulfate de cuivre', unit: 'kg/j', group: 'Réacteur', default: 0 },
    { key: 'decanteur_concentration', label: 'Concentration du technosable décanté', unit: 'g/L', group: 'Décanteur', default: 100 },
    { key: 'decanteur_capture', label: 'Taux de capture du décanteur', unit: '-', group: 'Décanteur', default: 0.95 },
    { key: 'taux_recirculation', label: 'Taux de recirculation du technosable', unit: '-', group: 'Décanteur', default: 0.5 },
    { key: 'conso_RTO', label: 'Consommation de combustible du RTO', unit: 'kWh PCI/j', group: 'Fumées', default: undefined, hint: '800 kWh PCI par tonne de MV' },
    { key: 'conso_elec_spec', label: 'Consommation électrique spécifique', unit: 'kWh/m³ de boues épaissies', group: 'Électricité', default: 20 },
  ],

  compute(ctx) {
    const { site, table, soluble, retour, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.athos
    const digestion = choices.digestion_amont !== 'aucune'
    const ratio_DBO_DCO_soluble = choices.digestion_amont === 'avancee' ? 0.2 * 0.8 : 0.3 * 0.8

    const entrees = entreesDepuisChoix(choices, p, ['entree_1', 'entree_2'])
    if (!entrees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    // =====================================================================
    // calcul_cocktail_boues_entrees
    // =====================================================================
    const lu = lireEntrees(table, soluble, entrees)
    let in_MES = 0, in_Q = 0, in_MV = 0
    let in_DCO = 0, in_DBO = 0, in_NK = 0, in_Pt = 0
    let MES_digerees = 0
    const soluble_kg = new Array(6).fill(0)
    const ratioMineral = {}
    for (const e of lu) {
      const { j, MES, Q, MV_MES, ratios, sol, src } = e
      if (!(MES > 0)) continue
      in_MES += MES; in_Q += Q; in_MV += MES * MV_MES
      in_DCO += MES * ratios.DCO
      in_DBO += MES * ratios.DBO
      in_NK += MES * ratios.NK
      in_Pt += MES * ratios.Pt
      for (let k = 1; k <= 5; k++) soluble_kg[k] += (sol[k] * Q) / 1000
      if (e.etape === ETAPE.digerees) MES_digerees += MES
      // la part minérale de chaque type de boue détermine la répartition du
      // technosable en sortie
      if (!ratioMineral[j]) ratioMineral[j] = { MES: 0, origine: src[P.origine], flux_in: 0 }
      ratioMineral[j].MES += MES * (1 - MV_MES)
      ratioMineral[j].flux_in += e.flux_in
    }
    for (const e of lu) e.src[P.verif_flux] -= e.flux_in

    if (!(in_MES > 0) || !(in_Q > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée d'Athos."] }
    }
    const in_MV_MES = in_MV / in_MES
    const in_siccite = in_MES / (in_Q * H.masse_volumique_boues_froides)
    const in_concentration = in_MES / in_Q
    const part_digeree = MES_digerees / in_MES
    const T_entree = f('temperature_entree') ?? site.T_eau_design
    const reduction_MV = p.reduction_MV ?? H.reduction_MV
    const reduction_DCO = p.reduction_DCO ?? H.reduction_DCO
    const T_cible = p.T_reacteur_cible ?? H.T_reacteur_cible
    const decanteur_conc = p.decanteur_concentration ?? 100
    const decanteur_capture = p.decanteur_capture ?? 0.95
    const taux_recirc = p.taux_recirculation ?? H.taux_recirculation_technosable

    // =====================================================================
    // Bilans de masse et d'énergie, pour une siccité d'épaississement donnée
    // =====================================================================
    /**
     * Enchaîne épaississement, bilan de masse itératif (boucle de recirculation
     * du technosable) et bilan d'énergie. Renvoie l'état complet du procédé.
     */
    function simuler(siccite_epaissies, capture) {
      // ---- calcul_epaississement
      const ep_MS = capture * in_MES
      const ep_BH = ep_MS / siccite_epaissies
      const ep_MV_MES = in_MV_MES
      const ep_MV = ep_MV_MES * ep_MS
      const ep_Q = ep_BH / H.masse_volumique_boues_froides
      // la pollution soluble suit l'eau conservée dans les boues épaissies
      const part_soluble = in_Q > 0 ? ep_Q / in_Q : 0
      const sol_epaissies = soluble_kg.map((v) => v * part_soluble)
      const DCO_particulaire = in_DCO * capture
      const DCO_epaissies = DCO_particulaire + sol_epaissies[SOL.DCO]

      // ---- consommation d'oxygène, proportionnelle à la DCO à oxyder
      const O2_kgj = f('consommation_O2') ?? H.ratio_O2_DCO * DCO_epaissies

      // ---- calcul_bilan_masse : point fixe sur la recirculation de technosable
      const techno_siccite = decanteur_conc / H.masse_volumique_technosable
      let rec_BH = taux_recirc * ep_BH
      let rec_MS = techno_siccite * rec_BH
      let rec_MV_MES = decanteur_capture * H.ratio_MV_particulaire_oxydees * (1 - reduction_MV) * ep_MV_MES
      let dec = null
      let eau_evaporee = 0
      let fumees_MS = 0, fumees_BH = 0
      let melange = null, reacteur_sortie = null
      let garde = 0
      let ecart = Infinity
      while (ecart > H.critere_bilan_masse && garde++ < 300) {
        const prev_MS = rec_MS, prev_BH = rec_BH, prev_MV = rec_MV_MES * rec_MS

        // mélange boues épaissies + technosable recirculé
        const mel_MS = ep_MS + rec_MS
        const mel_BH = ep_BH + rec_BH
        const mel_MV = ep_MV + rec_MV_MES * rec_MS
        melange = { MS: mel_MS, BH: mel_BH, MV_MES: mel_MS > 0 ? mel_MV / mel_MS : 0 }

        // fumées : MV oxydées + oxygène consommé, saturées en humidité
        fumees_MS = reduction_MV * ep_MV + O2_kgj
        const MM_fumees = reduction_DCO * MM_CO2 + (1 - reduction_DCO) * MM_O2
        fumees_BH = fumees_MS * (1 + humiditeSaturation(MM_fumees, H.pression_reacteur_bars, T_cible, 1))
        eau_evaporee = fumees_BH - fumees_MS

        // sortie de réacteur : ce qui entre, plus l'oxygène, moins les fumées
        const r_BH = mel_BH + O2_kgj - fumees_BH
        const r_MS = mel_MS - reduction_MV * ep_MV
        const r_MV = mel_MV - reduction_MV * ep_MV
        reacteur_sortie = { BH: r_BH, MS: r_MS, MV_MES: r_MS > 0 ? r_MV / r_MS : 0, siccite: r_BH > 0 ? r_MS / r_BH : 0 }

        // décanteur : le minéral est intégralement capté, la MV résiduelle
        // seulement pour sa fraction particulaire
        const d_MS = decanteur_capture * r_MS
          * (H.ratio_MV_particulaire_oxydees * reacteur_sortie.MV_MES + (1 - reacteur_sortie.MV_MES))
        const d_BH = d_MS / techno_siccite
        const d_MV = decanteur_capture * H.ratio_MV_particulaire_oxydees * reacteur_sortie.MV_MES * r_MS
        dec = { MS: d_MS, BH: d_BH, siccite: techno_siccite, MV_MES: d_MS > 0 ? d_MV / d_MS : 0 }

        rec_BH = taux_recirc * ep_BH
        rec_MS = techno_siccite * rec_BH
        rec_MV_MES = dec.MV_MES

        ecart = (rec_MS > 0 ? Math.pow(prev_MS - rec_MS, 2) / rec_MS : 0)
          + (rec_BH > 0 ? Math.pow(prev_BH - rec_BH, 2) / rec_BH : 0)
          + (rec_MV_MES * rec_MS > 0 ? Math.pow(prev_MV - rec_MV_MES * rec_MS, 2) / (rec_MV_MES * rec_MS) : 0)
      }
      if (garde >= 300) warnings.push("Le bilan de masse n'a pas convergé en 300 itérations.")

      // ---- calcul_bilan_energie : cinq équations, cinq inconnues
      // x0 T entrée échangeur 1, x1 T sortie échangeur 1, x2 T réacteur,
      // x3 énergie échangée entre les deux échangeurs, x4 T sortie échangeur 2
      const C_mel = melange.BH * H.Cp_boues_froides
      const C_ech1 = melange.BH * H.Cp_boues_150
      const C_reac = reacteur_sortie.BH * H.Cp_boues_250
      const C_ech2 = reacteur_sortie.BH * H.Cp_boues_110
      const A = [
        // mélange des boues épaissies et du technosable recirculé
        [C_mel, 0, 0, 0, 0, ep_BH * H.Cp_boues_froides * T_entree + rec_BH * H.Cp_boues_froides * H.T_technosable_sortie],
        // échangeur 1 : préchauffage des boues par l'énergie récupérée
        [C_mel, -C_ech1, 0, H.efficacite_echangeurs, 0, 0],
        // réacteur : combustion des MV, moins l'eau évaporée, plus l'oxygène
        [0, -C_ech1, fumees_MS * H.Cp_fumees + C_reac, 0, 0,
          H.PCI_MV * reduction_MV * ep_MV
          - H.enthalpie_vapeur_250 * eau_evaporee
          + O2_kgj * H.Cp_O2_15 * H.T_oxygene_entree],
        // échangeur 2 : cède l'énergie récupérée
        [0, 0, C_reac, -1, -C_ech2, 0],
        // écart de température imposé aux bornes de l'échangeur 2
        [0, 0, 1, 0, -1, H.ecart_T_echangeur2],
      ]
      const x = resoudreGauss(A, 5)
      if (!x) return null
      return {
        ep_MS, ep_BH, ep_Q, ep_MV, ep_MV_MES, siccite_epaissies, capture,
        sol_epaissies, DCO_particulaire, DCO_epaissies, O2_kgj,
        melange, reacteur_sortie, dec, rec_BH, rec_MS, rec_MV_MES,
        fumees_MS, fumees_BH, eau_evaporee,
        T_melange: x[0], T_ech1: x[1], T_reacteur: x[2], energie_echangee: x[3], T_ech2: x[4],
      }
    }

    // =====================================================================
    // calcul_siccite_boues_epaissies — dichotomie sur l'autothermie
    // =====================================================================
    let etat = null
    let epaississement_necessaire = false
    const siccite_forcee = f('siccite_boues_epaissies')
    if (siccite_forcee != null) {
      epaississement_necessaire = true
      const capture = p.capture_MES_epaississement ?? H.capture_MES_epaississement
      etat = simuler(siccite_forcee / H.masse_volumique_boues_froides, capture)
    } else if (in_siccite > H.siccite_max_epaissies) {
      // boues déjà trop concentrées : on les prend telles quelles
      etat = simuler(in_siccite, 1)
      warnings.push(`Siccité d'entrée (${(in_siccite * 100).toFixed(1)} %) supérieure au maximum admissible : les boues sont admises sans épaississement.`)
    } else {
      // première tentative sans épaississement
      etat = simuler(in_siccite, 1)
      if (etat && etat.T_reacteur < T_cible) {
        // pas autotherme : on épaissit juste ce qu'il faut
        epaississement_necessaire = true
        const capture = p.capture_MES_epaississement ?? H.capture_MES_epaississement
        let min = Math.max(in_siccite, H.siccite_min_epaissies)
        let max = H.siccite_max_epaissies
        let s = (min + max) / 2
        let critere = H.critere_siccite + 1
        let garde = 0
        while (critere > H.critere_siccite && garde++ < 200) {
          const essai = simuler(s, capture)
          if (!essai) break
          etat = essai
          if (essai.T_reacteur < T_cible) min = s
          else if (essai.T_reacteur > T_cible) max = s
          else { min = s; max = s }
          s = (min + max) / 2
          critere = s > 0 ? (max - min) / s : 0
        }
        etat = simuler(s, capture)
        if (etat && etat.T_reacteur < T_cible - 1) {
          warnings.push(`Réacteur non autotherme même à ${(H.siccite_max_epaissies * 100).toFixed(0)} % de siccité : température atteinte ${etat.T_reacteur.toFixed(0)} °C contre ${T_cible} °C visés.`)
        }
      }
    }
    if (!etat) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Le bilan d'énergie du réacteur n'a pas pu être résolu."] }
    }

    // =====================================================================
    // Sortie technosable, volume du réacteur, réactifs
    // =====================================================================
    const techno_MS = etat.dec.MS - etat.rec_MS
    const techno_BH = etat.dec.BH - etat.rec_BH
    const techno_Q = techno_BH / H.masse_volumique_boues_froides
    const techno_MV_MES = etat.dec.MV_MES
    const reacteur_volume = f('reacteur_volume')
      ?? (H.temps_sejour_h / CONST.NOMBRE_HEURE_PAR_JOUR) * etat.melange.BH / H.masse_volumique_boues_250
    const polymere_kgj = epaississement_necessaire ? ((p.dose_polymere ?? 6) * in_MES) / 1000 : 0
    const polymere_Q = polymere_kgj / H.concentration_polymere_gL
    const CuSO4_kgj = p.consommation_CuSO4 ?? 0
    const conso_RTO = f('conso_RTO') ?? (H.conso_RTO_kWhPCI_tMV * etat.ep_MV) / 1000
    const DCO_oxydee = reduction_DCO * etat.DCO_epaissies

    // =====================================================================
    // calcul_composition_retours
    // =====================================================================
    // pollution portée par les boues épaissies, une fois la capture appliquée
    const ep_DCO = in_DCO * etat.capture + etat.sol_epaissies[SOL.DCO]
    const ep_DBO = in_DBO * etat.capture + ratio_DBO_DCO_soluble * etat.sol_epaissies[SOL.DCO]
    const ep_NK = in_NK * etat.capture + etat.sol_epaissies[SOL.NK]
    const ep_Pt = in_Pt * etat.capture + etat.sol_epaissies[SOL.Pt]

    let ret_Q = 0, ret_MES = 0, ret_DCO = 0, ret_DBO = 0, ret_NK = 0, ret_NH4 = 0, ret_Pt = 0
    // surnageant de l'épaississement, s'il a lieu
    if (epaississement_necessaire) {
      ret_Q += in_Q + polymere_Q - etat.ep_Q
      ret_MES += in_MES * (1 - etat.capture)
      ret_DCO += in_DCO * (1 - etat.capture)
      ret_DBO += in_DBO * (1 - etat.capture)
      ret_NK += in_NK * (1 - etat.capture)
      ret_Pt += in_Pt * (1 - etat.capture)
      const dDCO = soluble_kg[SOL.DCO] - etat.sol_epaissies[SOL.DCO]
      const dNK = soluble_kg[SOL.NK] - etat.sol_epaissies[SOL.NK]
      const dPt = soluble_kg[SOL.Pt] - etat.sol_epaissies[SOL.Pt]
      ret_DCO += dDCO
      ret_DBO += ratio_DBO_DCO_soluble * dDCO
      ret_NK += dNK
      ret_NH4 += H.ratio_NH4_NK_soluble_digestion * dNK
      ret_Pt += dPt
    }
    // surnageant du décanteur et du filtre à plateaux, regroupés
    const athos_Q = (etat.reacteur_sortie.BH - etat.dec.BH) / H.masse_volumique_boues_froides
    const athos_DCO = (1 - reduction_DCO) * ep_DCO
    const athos_NK = H.retours_NK_sur_entree * ep_NK
    const athos_NK_soluble = athos_NK * H.retours_NK_soluble_sur_total
    const athos_NH4 = athos_NK_soluble * (digestion
      ? H.retours_NH4_sur_NK_avec_digestion
      : H.retours_NH4_sur_NK_sans_digestion)
    const athos_Pt = (digestion ? H.retours_Pt_sur_entree_avec_digestion : H.retours_Pt_sur_entree_sans_digestion) * ep_Pt
    ret_Q += athos_Q
    ret_DCO += athos_DCO
    ret_DBO += athos_DCO // le VBA reprend la DCO telle quelle pour la DBO
    ret_NK += athos_NK
    ret_NH4 += athos_NH4
    ret_Pt += athos_Pt

    // Le VBA n'affecte jamais `FluxRetoursTotauxAthos(repere_ret_MES)` : les MS
    // qui échappent au décanteur ne sont comptées nulle part. Le comportement
    // est reproduit, mais le flux manquant est calculé et signalé.
    const MS_non_captees = Math.max(0, etat.reacteur_sortie.MS - etat.dec.MS)
    if (MS_non_captees > 0) {
      warnings.push(`Les MS non captées par le décanteur (${MS_non_captees.toFixed(0)} kg/j) ne sont pas comptées au retour en tête, conformément au classeur d'origine.`)
    }

    // Le surnageant du décanteur Athos est tenu à part : c'est l'un des deux
    // jus que le traitement des retours peut prendre en charge.
    const vecteur = []
    vecteur[RET.Q] = ret_Q
    vecteur[RET.MES] = ret_MES
    vecteur[RET.DCO] = ret_DCO
    vecteur[RET.DBO] = ret_DBO
    vecteur[RET.NK] = ret_NK
    vecteur[RET.NH4] = ret_NH4
    vecteur[RET.Pt] = ret_Pt
    const vecteur_soluble = []
    vecteur_soluble[RET.Q] = athos_Q
    vecteur_soluble[RET.DCO] = athos_DCO
    vecteur_soluble[RET.NK] = athos_NK_soluble
    vecteur_soluble[RET.NH4] = athos_NH4
    vecteur_soluble[RET.Pt] = athos_Pt * H.retours_Pt_soluble_sur_total
    if (ctx.retours) ajouterRetour(ctx.retours, RET_ORIGINE.athos, vecteur, vecteur_soluble)
    else for (let i = 1; i <= 8; i++) retour[i] += vecteur[i] || 0

    // =====================================================================
    // calcul_boues_sortie_procede — le technosable rejoint l'étape aval
    // =====================================================================
    let mineral_total = 0
    for (const t of Object.values(ratioMineral)) mineral_total += t.MES
    const Pt_technosable = ep_Pt - athos_Pt
    for (const [jStr, t] of Object.entries(ratioMineral)) {
      const j = Number(jStr)
      const part = mineral_total > 0 ? t.MES / mineral_total : 0
      const dst = table[etapeOut][j]
      dst[P.origine] = t.origine
      dst[P.MES] = part * techno_MS
      dst[P.Q] = part * techno_Q
      dst[P.MV_MES] = techno_MV_MES
      // le technosable est minéral : plus de matière organique ni d'azote
      dst[P.ratio_DCO_MES] = 0
      dst[P.ratio_DBO_MES] = 0
      dst[P.ratio_NK_MES] = 0
      dst[P.ratio_Pt_MES] = techno_MS > 0 ? Math.max(0, Pt_technosable) / techno_MS : 0
      dst[P.flux_in] = t.flux_in
      dst[P.verif_flux] = t.flux_in
      for (let k = 1; k <= 5; k++) soluble[etapeOut][j][k] = 0
    }

    // =====================================================================
    const electricite = (p.conso_elec_spec ?? H.conso_elec_kWh_m3) * etat.ep_Q

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MS en entrée', unit: 'kg/j', value: in_MES },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: in_Q },
        { key: 'in_conc', label: 'Siccité en entrée', unit: 'g/L', value: in_concentration },
        { key: 'in_MV_MES', label: 'MV/MS en entrée', unit: '-', value: in_MV_MES },
        { key: 'epaiss', label: 'Épaississement nécessaire', unit: '', value: epaississement_necessaire ? 1 : 0 },
        { key: 'ep_conc', label: 'Siccité des boues épaissies', unit: 'g/L', value: etat.siccite_epaissies * H.masse_volumique_boues_froides },
        { key: 'ep_MS', label: 'MS des boues épaissies', unit: 'kg/j', value: etat.ep_MS },
        { key: 'ep_Q', label: 'Débit de boues épaissies', unit: 'm³/j', value: etat.ep_Q },
        ...(polymere_kgj > 0 ? [{ key: 'poly', label: 'Polymère pur', unit: 'kg/j', value: polymere_kgj }] : []),
        { key: 'T_reac', label: 'Température du réacteur', unit: '°C', value: etat.T_reacteur },
        { key: 'T_mel', label: "Température en entrée d'échangeur 1", unit: '°C', value: etat.T_melange },
        { key: 'T_ech1', label: "Température en sortie d'échangeur 1", unit: '°C', value: etat.T_ech1 },
        { key: 'T_ech2', label: "Température en sortie d'échangeur 2", unit: '°C', value: etat.T_ech2 },
        { key: 'E_ech', label: 'Énergie échangée entre les échangeurs', unit: 'kWh/j', value: etat.energie_echangee },
        { key: 'V_reac', label: 'Volume du réacteur', unit: 'm³', value: reacteur_volume },
        { key: 'P_reac', label: 'Pression absolue du réacteur', unit: 'bars', value: H.pression_reacteur_bars },
        { key: 'O2', label: "Oxygène liquide pur", unit: 'kg/j', value: etat.O2_kgj },
        ...(CuSO4_kgj > 0 ? [{ key: 'CuSO4', label: 'Sulfate de cuivre pur', unit: 'kg/j', value: CuSO4_kgj }] : []),
        { key: 'fumees', label: 'Fumées humides', unit: 'kg/j', value: etat.fumees_BH },
        { key: 'evap', label: 'Eau évaporée avec les fumées', unit: 'kg/j', value: etat.eau_evaporee },
        { key: 'RTO', label: `Combustible du RTO (${choices.combustible_RTO})`, unit: 'kWh PCI/j', value: conso_RTO },
        { key: 'DCO_ox', label: 'DCO oxydée', unit: 'kg/j', value: DCO_oxydee },
        { key: 'red_MV', label: 'Réduction des MV', unit: '-', value: reduction_MV },
        { key: 'techno_MS', label: 'Technosable produit', unit: 'kg MS/j', value: techno_MS },
        { key: 'techno_Q', label: 'Débit de technosable', unit: 'm³/j', value: techno_Q },
        { key: 'techno_conc', label: 'Concentration du technosable', unit: 'g/L', value: decanteur_conc },
        { key: 'techno_MV', label: 'MV/MS du technosable', unit: '-', value: techno_MV_MES },
        { key: 'rec_Q', label: 'Technosable recirculé', unit: 'm³/j', value: etat.rec_BH / H.masse_volumique_boues_froides },
        { key: 'MV_oxydees', label: 'MV oxydées dans le réacteur', unit: 'kg/j', value: reduction_MV * etat.ep_MV },
        { key: 'MS_non_captees', label: 'MS échappant au décanteur (non comptées au retour)', unit: 'kg/j', value: MS_non_captees },
        { key: 'ret_Q', label: 'Retours en tête', unit: 'm³/j', value: ret_Q },
        { key: 'ret_athos_Q', label: 'dont surnageant du décanteur Athos', unit: 'm³/j', value: athos_Q },
        { key: 'ret_NK', label: 'NK au retour en tête', unit: 'kg/j', value: ret_NK },
        { key: 'ret_NH4', label: 'dont N-NH4', unit: 'kg/j', value: ret_NH4 },
        { key: 'ret_NK_mgL', label: 'Concentration en NK des retours', unit: 'mg/L', value: ret_Q > 0 ? (ret_NK / ret_Q) * 1000 : 0 },
        { key: 'ret_DCO', label: 'DCO au retour en tête', unit: 'kg/j', value: ret_DCO },
      ],
      reactifs: {
        oxygene_liquide: etat.O2_kgj,
        ...(polymere_kgj > 0 ? { polymere: polymere_kgj } : {}),
        ...(CuSO4_kgj > 0 ? { CuSO4: CuSO4_kgj } : {}),
      },
      energie: {
        besoin_thermique_kWhj: conso_RTO,
        source: choices.combustible_RTO === 'biogaz' ? 'biogaz' : 'combustible',
        recuperable_kWhj: 0,
      },
      electricity: { total: electricite, fixed: 0, detail: { procede: electricite } },
      warnings,
    }
  },
})
