// ---------------------------------------------------------------------------
// Port de z_Exelys_DLD.cls — Exelys en configuration Digestion-Lyse-Digestion.
//
// Contrairement à ce que la parenté de nommage laisse croire, Exelys DLD n'est
// pas une variante de Biothelys : la chaîne comporte **deux digesteurs** séparés
// par une déshydratation et une lyse continue.
//
//   entrée → digesteur 1 (35 °C, SRT 12 j)
//          → centrifugeuse intermédiaire (230 g/L, 97 % de capture)
//          → mélange avec les boues fraîches non digérées éventuelles
//          → lyse continue à 165 °C par injection directe de vapeur (15 bars)
//          → échangeurs 90 °C puis 55 °C
//          → dilution à 100 g/L → digesteur 2 (38 °C, SRT 15 j)
//
// Trois particularités par rapport à Biothelys :
//
//  1. La **centrifugeuse intermédiaire** ne capture pas la totalité des MV : la
//     digestion 1 a solubilisé une partie des matières volatiles, qui repart au
//     centrat. Cette part solubilisée dépend de l'origine de la boue et du temps
//     de séjour du premier digesteur, via deux jeux de coefficients directeurs
//     (temps de séjour court ou long). C'est le calcul le plus spécifique du
//     module.
//  2. Le **second digesteur** applique un taux de réduction unique de 0,3656
//     aux boues déjà digérées puis lysées — la nature de la boue d'origine n'a
//     plus d'importance après ce double traitement. Les boues fraîches admises
//     en by-pass, elles, gardent leur taux « lysées » propre.
//  3. La consommation de vapeur est calculée par **bilan enthalpique** (procédé
//     continu), et non par corrélation comme dans le Biothelys en batch.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - la codigestion de coferments n'est pas portée ;
//  - le raccordement au module PINCH n'est pas fait, les besoins et
//    disponibilités thermiques étant calculés et exposés sans consommateur ;
//  - le centrat de la centrifugeuse intermédiaire part au retour en tête, comme
//    dans le VBA, mais celui-ci le range dans un vecteur `RetoursDigestion`
//    distinct que le port ne tient pas séparément.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, RET, NB_TYPES, RET_ORIGINE, ajouterRetour } from '../core/sludge.js'
import { CONST, rendementMoteur } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees, repartitionPompage, rdtPompeBoues, elecPompage } from './_commun.js'
import { REDUCTION_MV_LYSEES, REDUCTION_MV_NON_LYSEES, correctionTemperatureDigesteur, correctionSRT } from './biothelys.js'

// Part des MV solubilisées lors de la digestion 1, au temps de séjour de
// référence (20 j en mésophile). Ces MV échappent à la centrifugeuse.
const MV_SOLUBLES_REFERENCE = {
  I_simple: 0.0649, I_reactif: 0.0599,
  II_forte: 0.0409, II_moyenne: 0.0351, II_faible_EB: 0.015, II_faible_ED: 0.015,
  II_prolongee_EB: 0.015, II_prolongee_ED: 0.015, II_MBR: 0.015, II_MBBR: 0.015, II_HybAS: 0.015,
  II_biostyr_C: 0.0249, II_biostyr_N: 0.0249, II_biostyr_NDN: 0.0249, II_biostyr_PDN: 0.0249,
  III_decantation: 0.0064, III_biostyr_N: 0.0064, III_biostyr_PDN: 0.0064,
  codigestion_graisses: 0.0726,
}
// Coefficients directeurs de la correction sur le temps de séjour réduit,
// selon qu'il est inférieur (court) ou supérieur (long) à la référence.
const CORRECTION_MV_SOLUBLE_COURT = {
  I_simple: -3.3846, I_reactif: -3.5276,
  II_forte: -3.9355, II_moyenne: -4.0652, II_faible_EB: -4.62, II_faible_ED: -4.62,
  II_prolongee_EB: -4.62, II_prolongee_ED: -4.62, II_MBR: -4.62, II_MBBR: -4.62, II_HybAS: -4.62,
  II_biostyr_C: -4.3336, II_biostyr_N: -4.3336, II_biostyr_NDN: -4.3336, II_biostyr_PDN: -4.3336,
  III_decantation: -5.1835, III_biostyr_N: -5.1835, III_biostyr_PDN: -5.1835,
  codigestion_graisses: -2.5583,
}
const CORRECTION_MV_SOLUBLE_LONG = {
  I_simple: -1.0293, I_reactif: -1.0551,
  II_forte: -1.0905, II_moyenne: -1.0997, II_faible_EB: -1.14, II_faible_ED: -1.14,
  II_prolongee_EB: -1.14, II_prolongee_ED: -1.14, II_MBR: -1.14, II_MBBR: -1.14, II_HybAS: -1.14,
  II_biostyr_C: -1.1165, II_biostyr_N: -1.1165, II_biostyr_NDN: -1.1165, II_biostyr_PDN: -1.1165,
  III_decantation: -1.2188, III_biostyr_N: -1.2188, III_biostyr_PDN: -1.2188,
  codigestion_graisses: -0.6997,
}

const H = {
  // le taux de réduction du second digesteur ne dépend plus de l'origine :
  // 35 % à 15 jours, ramenés à 20 jours par la correction de temps de séjour
  reduction_MV_digerees_lysees: 0.3656,
  T_limite_meso_thermo: 50, // °C
  SRT_reference_dig1_meso: 20, // j
  SRT_reference_dig1_thermo: 12,
  SRT_reference_dig2: 15,
  T_reference_dig1: 35, // °C
  T_reference_dig2: 38,
  a_correction_SRT: 1.25,
  b_correction_SRT: 0.7,
  a_densite_biogaz: [44 / 22.4, (16 - 44) / 22.4],
  ratio_CH4_DCO: 0.35,
  ratio_NK_soluble_MVdeg: 0.9,
  ratio_P_soluble_P_entree: 0.16,
  ratio_DCO_soluble_dig1: 0.05,
  ratio_DCO_soluble_dig2: 0.2,
  concentration_polymere_gL: 1.5,
  perte_T_avant_lyse: 10, // °C entre digesteur 1 et entrée d'hydrolyse
  ratio_pertes_vapeur: 0.05,
  pertes_thermiques: {
    froid: { oui: 0.6, non: 3.72 },
    tempere: { oui: 0.36, non: 2.2 },
    chaud: { oui: 0.24, non: 1.44 },
  },
  efficacite_echangeur: 0.95,
  Cp_boues_kWh_m3C: 1.163,
  Cp_eau_kcal_kgC: 1,
  ratio_NH4_NK_soluble: 0.95,
  ratio_DBO_DCO_soluble: 0.2 * 0.8,
  masse_volumique: 1000,
  coef_T_saturation: [3816.44, 18.3, 750, 46.13, 273.15],
  coef_enthalpie_vap: [616.87, -0.7542],
  concentration_dilution_dig2: 100, // g/L
  agitation_W_m3: 7,
  agitation_hj: 24,
  brassage_biogaz_Nm_h: 1.25,
  conso_surpresseur: 4, // Wh/(Nm³·mCE)
  T_lyse: 165, // °C
  pression_vapeur_bars: 15,
  T_prechauffage_eau: 95, // °C
  T_initiale_eau: 15, // °C
  centrifugeuse_capture: 0.97,
  centrifugeuse_concentration: 230, // g/L
  centrifugeuse_polymere_kg_tMS: 10,
  conso_elec_centrifugeuse_kWh_m3: 1.5,
}

const kcalVersKWh = (kcal) => (kcal * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE
function temperatureSaturation(P_bars) {
  const c = H.coef_T_saturation
  return c[0] / (c[1] - Math.log(c[2] * P_bars)) + c[3] - c[4]
}

const ENTREES = [
  { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
  { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
  { value: 'epaississeur_3', label: 'sortie épaississeur 3' },
  { value: 'toutes', label: 'boues extraites, toutes origines' },
  { value: 'I', label: 'boues primaires brutes' },
  { value: 'II', label: 'boues secondaires brutes' },
  { value: 'graisses', label: 'graisses' },
]

export default defineSludgeNode({
  id: 'exelys',
  label: 'Exelys DLD',
  short: 'Exelys',
  family: 'stabilisation',
  vba: 'z_Exelys_DLD.cls',
  etapeSortie: ETAPE.digerees,
  description:
    "Exelys en configuration Digestion-Lyse-Digestion : un premier digesteur, une centrifugeuse intermédiaire, une lyse thermique continue à 165 °C, puis un second digesteur. La lyse relance la dégradation de boues déjà digérées, que la digestion seule ne peut plus attaquer.",
  choices: [
    { key: 'entree_1', label: 'Boues admises au digesteur 1 (entrée 1)', default: 'epaississeur_1', options: [...ENTREES, { value: 'aucune', label: 'aucune' }] },
    { key: 'entree_2', label: 'Boues admises au digesteur 1 (entrée 2)', default: 'aucune', options: [{ value: 'aucune', label: 'aucune' }, ...ENTREES] },
    { key: 'entree_non_digeree', label: 'Boues fraîches admises en lyse (by-pass du digesteur 1)', default: 'aucune', options: [{ value: 'aucune', label: 'aucune' }, ...ENTREES] },
    { key: 'agitation_dig1', label: 'Agitation du digesteur 1', default: 'biogaz', options: [
      { value: 'mecanique', label: 'mécanique' },
      { value: 'recirculation', label: 'par recirculation des boues' },
      { value: 'biogaz', label: 'par injection de biogaz' },
    ] },
    { key: 'agitation_dig2', label: 'Agitation du digesteur 2', default: 'biogaz', options: [
      { value: 'mecanique', label: 'mécanique' },
      { value: 'recirculation', label: 'par recirculation des boues' },
      { value: 'biogaz', label: 'par injection de biogaz' },
    ] },
    { key: 'isolation', label: 'Digesteurs isolés', default: 'oui', options: [
      { value: 'oui', label: 'oui' }, { value: 'non', label: 'non' },
    ] },
    { key: 'climat', label: 'Climat du site', default: 'tempere', options: [
      { value: 'froid', label: 'froid' }, { value: 'tempere', label: 'tempéré' }, { value: 'chaud', label: 'chaud' },
    ] },
    { key: 'biogaz_stock', label: 'Gazomètre de stockage', default: 'oui', options: [
      { value: 'oui', label: 'oui' }, { value: 'non', label: 'non' },
    ] },
    { key: 'valorisation_biogaz', label: 'Valorisation du biogaz', default: 'cogeneration', options: [
      { value: 'chaudiere', label: 'chaudière' },
      { value: 'cogeneration', label: 'cogénération' },
      { value: 'microturbine', label: 'microturbine' },
      { value: 'torchere', label: 'torchère seule' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_3', label: 'Part du flux amont admise (by-pass)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'temperature_entree', label: 'Température des boues en entrée', unit: '°C', group: 'Alimentation', default: undefined, hint: "température de l'eau en conditions réelles" },
    { key: 'T_dig1', label: 'Température du digesteur 1', unit: '°C', group: 'Digesteur 1', default: 35 },
    { key: 'SRT_dig1', label: 'Temps de séjour du digesteur 1', unit: 'j', group: 'Digesteur 1', default: 12 },
    { key: 'volume_dig1', label: 'Volume du digesteur 1', unit: 'm³', group: 'Digesteur 1', default: undefined, hint: 'débit × temps de séjour' },
    { key: 'hauteur_dig1', label: 'Hauteur du digesteur 1', unit: 'm', group: 'Digesteur 1', default: 10 },
    { key: 'tx_recirculation_dig1', label: 'Taux de recirculation pour agitation (digesteur 1)', unit: 'vol/j', group: 'Digesteur 1', default: 5 },
    { key: 'centrifugeuse_capture', label: 'Taux de capture de la centrifugeuse', unit: '-', group: 'Centrifugeuse', default: 0.97 },
    { key: 'centrifugeuse_concentration', label: 'Siccité en sortie de centrifugeuse', unit: 'g/L', group: 'Centrifugeuse', default: 230 },
    { key: 'centrifugeuse_polymere', label: 'Dosage de polymère', unit: 'kg m.a./tMS', group: 'Centrifugeuse', default: 10 },
    { key: 'pression_vapeur', label: 'Pression absolue de la vapeur saturante', unit: 'bars', group: 'Lyse', default: 15 },
    { key: 'temperature_lyse', label: 'Température de lyse', unit: '°C', group: 'Lyse', default: 165 },
    { key: 'debit_vapeur', label: 'Consommation de vapeur', unit: 'kg/j', group: 'Lyse', default: undefined, hint: 'bilan enthalpique sur la lyse' },
    { key: 'temperature_initiale_eau_vapeur', label: "Température initiale de l'eau alimentant la chaudière", unit: '°C', group: 'Lyse', default: 15 },
    { key: 'temperature_prechauffage_eau_vapeur', label: "Température de préchauffage de l'eau", unit: '°C', group: 'Lyse', default: 95 },
    { key: 'T_dig2', label: 'Température du digesteur 2', unit: '°C', group: 'Digesteur 2', default: 38, hint: 'mésophile uniquement' },
    { key: 'SRT_dig2', label: 'Temps de séjour du digesteur 2', unit: 'j', group: 'Digesteur 2', default: 15 },
    { key: 'volume_dig2', label: 'Volume du digesteur 2', unit: 'm³', group: 'Digesteur 2', default: undefined, hint: 'débit × temps de séjour' },
    { key: 'hauteur_dig2', label: 'Hauteur du digesteur 2', unit: 'm', group: 'Digesteur 2', default: 10 },
    { key: 'tx_recirculation_dig2', label: 'Taux de recirculation pour agitation (digesteur 2)', unit: 'vol/j', group: 'Digesteur 2', default: 5 },
    { key: 'concentration_dilution', label: 'Siccité visée en entrée du digesteur 2', unit: 'g/L', group: 'Digesteur 2', default: 100 },
    { key: 'debit_eau_dilution', label: "Débit d'eau de dilution", unit: 'm³/j', group: 'Digesteur 2', default: undefined, hint: 'calculé sur la siccité visée' },
    { key: 'biogaz_teneur_CH4', label: 'Teneur en méthane du biogaz', unit: '-', group: 'Biogaz', default: 0.63 },
    { key: 'ratio_biogaz_valorise', label: 'Part du biogaz valorisée', unit: '-', group: 'Biogaz', default: 1 },
    { key: 'pompage_dig1_nb', label: 'Nombre de pompes digesteur 1 → centrifugeuse', unit: 'u', group: 'Pompages', default: 2 },
    { key: 'pompage_dig1_P_refoulement', label: 'Pression de refoulement digesteur 1 → centrifugeuse', unit: 'mCE', group: 'Pompages', default: 50 },
    { key: 'pompage_dig1_tps_fonctionnement', label: 'Durée de fonctionnement digesteur 1 → centrifugeuse', unit: 'h/j', group: 'Pompages', default: 24 },
    { key: 'pompage_lyse_nb', label: 'Nombre de pompes centrifugeuse → lyse', unit: 'u', group: 'Pompages', default: 1 },
    { key: 'pompage_lyse_P_refoulement', label: 'Pression de refoulement centrifugeuse → lyse', unit: 'mCE', group: 'Pompages', default: 120 },
    { key: 'pompage_lyse_tps_fonctionnement', label: 'Durée de fonctionnement centrifugeuse → lyse', unit: 'h/j', group: 'Pompages', default: 24 },
    { key: 'pompage_sortie_nb', label: 'Nombre de pompes de sortie du digesteur 2', unit: 'u', group: 'Pompages', default: 2 },
    { key: 'pompage_sortie_P_refoulement', label: 'Pression de refoulement de sortie', unit: 'mCE', group: 'Pompages', default: 50 },
    { key: 'pompage_sortie_tps_fonctionnement', label: 'Durée de fonctionnement de sortie', unit: 'h/j', group: 'Pompages', default: 24 },
  ],

  compute(ctx) {
    const { site, table, soluble, retour, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.digerees

    const entreesDig1 = entreesDepuisChoix(choices, p, ['entree_1', 'entree_2'])
    const entreesFraiches = entreesDepuisChoix(
      { entree_1: choices.entree_non_digeree }, { ratio_admis_1: p.ratio_admis_3 }, ['entree_1'])
    if (!entreesDig1.length && !entreesFraiches.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    // ---- attribution_valeur_par_defaut
    const T_dig1 = p.T_dig1 ?? H.T_reference_dig1
    const T_dig2 = p.T_dig2 ?? H.T_reference_dig2
    if (T_dig2 > H.T_limite_meso_thermo) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: [`Le second digesteur ne peut fonctionner qu'en mésophile : température demandée ${T_dig2} °C, maximum ${H.T_limite_meso_thermo} °C.`],
      }
    }
    // la référence de temps de séjour du digesteur 1 dépend du régime retenu
    const SRT_ref_dig1 = T_dig1 <= H.T_limite_meso_thermo ? H.SRT_reference_dig1_meso : H.SRT_reference_dig1_thermo
    const SRT_dig1 = p.SRT_dig1 ?? 12
    const SRT_dig2 = p.SRT_dig2 ?? H.SRT_reference_dig2
    const T_entree = f('temperature_entree') ?? site.T_eau_exploit
    const P_vapeur = p.pression_vapeur ?? H.pression_vapeur_bars
    const T_lyse = p.temperature_lyse ?? H.T_lyse
    const capture = p.centrifugeuse_capture ?? H.centrifugeuse_capture
    const conc_centri = p.centrifugeuse_concentration ?? H.centrifugeuse_concentration
    const biogaz_CH4 = p.biogaz_teneur_CH4 ?? 0.63
    const T_initiale_eau = p.temperature_initiale_eau_vapeur ?? H.T_initiale_eau
    const T_prechauffage = p.temperature_prechauffage_eau_vapeur ?? H.T_prechauffage_eau

    const T_saturation = temperatureSaturation(P_vapeur)
    let enthalpie_vap = 0
    for (let i = 0; i <= 1; i++) enthalpie_vap += H.coef_enthalpie_vap[i] * Math.pow(T_saturation, i)
    const enthalpie_totale_kcal = H.Cp_eau_kcal_kgC * T_saturation + enthalpie_vap

    // ---- lecture des flux
    function lireFlux(entrees) {
      const lu = lireEntrees(table, soluble, entrees)
      const parType = {}
      let MS = 0, Q = 0, MV = 0
      const sol = new Array(6).fill(0)
      for (const e of lu) {
        const { j, MES, Q: q, MV_MES, ratios, sol: s, src } = e
        if (!(MES > 0)) continue
        const origine = src[P.origine]
        MS += MES; Q += q; MV += MES * MV_MES
        for (let k = 1; k <= 5; k++) sol[k] += (s[k] * q) / 1000
        if (!parType[j]) parType[j] = { MS: 0, Q: 0, MV: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, origine }
        const t = parType[j]
        t.MS += MES; t.Q += q; t.MV += MES * MV_MES
        t.DCO += MES * ratios.DCO
        t.DBO += MES * ratios.DBO
        t.NK += MES * ratios.NK
        t.Pt += MES * ratios.Pt
        t.flux_in += e.flux_in
      }
      for (const e of lu) e.src[P.verif_flux] -= e.flux_in
      return { parType, MS, Q, MV, sol }
    }
    const amont = lireFlux(entreesDig1)
    const fraiches = lireFlux(entreesFraiches)
    if (!(amont.MS + fraiches.MS > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée de l'Exelys."] }
    }

    // =====================================================================
    // Digesteur 1 — réduction par origine, boues non lysées
    // =====================================================================
    const volume_dig1 = f('volume_dig1') ?? amont.Q * SRT_dig1
    const SRT_dig1_reel = amont.Q > 0 ? volume_dig1 / amont.Q : SRT_dig1
    const corr_SRT_dig1 = correctionSRT(SRT_dig1_reel, SRT_ref_dig1)
    const corr_T_dig1 = correctionTemperatureDigesteur(T_dig1)

    let dig1_MV_reduites = 0
    const apresDig1 = {}
    for (const [jStr, t] of Object.entries(amont.parType)) {
      const j = Number(jStr)
      const red_ref = REDUCTION_MV_NON_LYSEES[t.origine]
      if (red_ref == null && t.origine) {
        warnings.push(`Origine de boue « ${t.origine} » sans taux de réduction des MV : boue considérée non dégradable.`)
      }
      const reduit = (red_ref ?? 0) * t.MV * corr_SRT_dig1 * corr_T_dig1
      dig1_MV_reduites += reduit
      apresDig1[j] = {
        MS: t.MS - reduit,
        MV: t.MV - reduit,
        Q: t.Q - reduit / H.masse_volumique,
        DCO: t.DCO, DBO: t.DBO, NK: t.NK, Pt: t.Pt,
        flux_in: t.flux_in, origine: t.origine, reduit,
      }
    }
    const dig1_MS = Object.values(apresDig1).reduce((s, t) => s + t.MS, 0)
    const dig1_MV = Object.values(apresDig1).reduce((s, t) => s + t.MV, 0)
    const dig1_Q = Object.values(apresDig1).reduce((s, t) => s + t.Q, 0)

    // =====================================================================
    // calcul_MV_soluble_boues_digerees — part des MV restée particulaire
    //
    // La digestion 1 solubilise une partie des MV, qui échappera à la
    // centrifugeuse. La part solubilisée dépend de l'origine et du temps de
    // séjour réduit, avec deux pentes selon qu'il est court ou long.
    // =====================================================================
    const ratio_SRT = SRT_ref_dig1 > 0 ? SRT_dig1_reel / SRT_ref_dig1 : 1
    let MV_solubles_ratio = 0
    for (const [jStr, t] of Object.entries(apresDig1)) {
      const origine = t.origine
      const pente = ratio_SRT < 1
        ? (CORRECTION_MV_SOLUBLE_COURT[origine] ?? 0)
        : (CORRECTION_MV_SOLUBLE_LONG[origine] ?? 0)
      const correction = pente * (ratio_SRT - 1) + 1
      // le VBA pondère par la part de réduction de ce type dans le total
      const part = dig1_MV_reduites > 0 ? t.reduit / dig1_MV_reduites : 0
      MV_solubles_ratio += part * (MV_SOLUBLES_REFERENCE[origine] ?? 0) * correction
    }
    const MV_particulaires_ratio = Math.max(0, Math.min(1, 1 - MV_solubles_ratio))

    // =====================================================================
    // calcul_bilan_masse_centrifugeuse
    // =====================================================================
    const polymere_kgj = ((p.centrifugeuse_polymere ?? H.centrifugeuse_polymere_kg_tMS) * dig1_MS) / 1000
    const polymere_Q = polymere_kgj / H.concentration_polymere_gL
    const apresCentri = {}
    let centri_MS = 0, centri_MV = 0
    for (const [jStr, t] of Object.entries(apresDig1)) {
      const j = Number(jStr)
      // le minéral est capté en totalité par la centrifugeuse, les MV
      // seulement pour leur fraction restée particulaire
      const mineral = (t.MS - t.MV) * capture
      const MV = t.MV * capture * MV_particulaires_ratio
      const MS = mineral + MV
      centri_MS += MS
      centri_MV += MV
      apresCentri[j] = {
        MS, MV,
        DCO: t.DCO * capture, DBO: t.DBO * capture,
        NK: t.NK * capture, Pt: t.Pt * capture,
        flux_in: t.flux_in, origine: t.origine, reduit: t.reduit,
      }
    }
    const centri_Q = conc_centri > 0 ? centri_MS / conc_centri : 0
    // centrat : ce que la centrifugeuse n'a pas capté, plus l'eau libérée
    const centrat_Q = dig1_Q + polymere_Q - centri_Q
    const centrat_MS = dig1_MS - centri_MS
    const centrat_MV = dig1_MV - centri_MV

    // =====================================================================
    // Mélange avec les boues fraîches, puis lyse continue
    // =====================================================================
    const melange_MS = centri_MS + fraiches.MS
    const melange_Q = centri_Q + fraiches.Q
    const T_entree_lyse = Math.max(0, T_dig1 - H.perte_T_avant_lyse)
    const T_melange = melange_Q > 0
      ? (centri_Q * T_entree_lyse + fraiches.Q * T_entree) / melange_Q
      : T_entree_lyse

    // bilan enthalpique de la lyse continue : la vapeur injectée porte les
    // boues de T_melange à T_lyse, en tenant compte des pertes
    const denominateur = (1 - H.ratio_pertes_vapeur) * enthalpie_totale_kcal - H.Cp_eau_kcal_kgC * T_lyse
    let debit_vapeur = f('debit_vapeur')
    if (debit_vapeur == null) {
      debit_vapeur = denominateur > 0
        ? (melange_Q * H.masse_volumique * H.Cp_eau_kcal_kgC * (T_lyse - T_melange)) / denominateur
        : 0
      if (denominateur <= 0) {
        warnings.push("La température de lyse dépasse ce que la vapeur saturante peut apporter : consommation de vapeur non calculable.")
      }
    }
    const apresLyse_Q = melange_Q + debit_vapeur / H.masse_volumique
    const apresLyse_conc = apresLyse_Q > 0 ? melange_MS / apresLyse_Q : 0

    // ---- dilution avant le second digesteur
    const conc_visee = p.concentration_dilution ?? H.concentration_dilution_dig2
    let eau_dilution = f('debit_eau_dilution')
    if (eau_dilution == null) {
      eau_dilution = conc_visee > apresLyse_conc ? 0
        : melange_MS / conc_visee - melange_MS / apresLyse_conc
    }
    const dig2_Q_entree = apresLyse_Q + eau_dilution
    const dig2_conc_entree = dig2_Q_entree > 0 ? melange_MS / dig2_Q_entree : 0

    // =====================================================================
    // Digesteur 2 — taux unique pour les boues digérées puis lysées,
    // taux « lysées » propre pour les boues fraîches admises en by-pass
    // =====================================================================
    const volume_dig2 = f('volume_dig2') ?? dig2_Q_entree * SRT_dig2
    const SRT_dig2_reel = dig2_Q_entree > 0 ? volume_dig2 / dig2_Q_entree : SRT_dig2
    const corr_SRT_dig2 = correctionSRT(SRT_dig2_reel, H.SRT_reference_dig2)
    const corr_T_dig2 = correctionTemperatureDigesteur(T_dig2)

    const sortie = {}
    let dig2_MV_reduites = 0
    for (const [jStr, t] of Object.entries(apresCentri)) {
      const j = Number(jStr)
      const reduit = H.reduction_MV_digerees_lysees * t.MV * corr_SRT_dig2 * corr_T_dig2
      dig2_MV_reduites += reduit
      sortie[j] = {
        MS: t.MS - reduit, MV: t.MV - reduit,
        DCO: t.DCO, DBO: t.DBO, NK: t.NK, Pt: t.Pt,
        flux_in: t.flux_in, origine: t.origine, reduit,
      }
    }
    for (const [jStr, t] of Object.entries(fraiches.parType)) {
      const j = Number(jStr)
      const red_ref = REDUCTION_MV_LYSEES[t.origine] ?? 0
      const reduit = red_ref * t.MV * corr_SRT_dig2 * corr_T_dig2
      dig2_MV_reduites += reduit
      if (!sortie[j]) {
        sortie[j] = { MS: 0, MV: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, origine: t.origine, reduit: 0 }
      }
      const d = sortie[j]
      d.MS += t.MS - reduit
      d.MV += t.MV - reduit
      d.DCO += t.DCO; d.DBO += t.DBO; d.NK += t.NK; d.Pt += t.Pt
      d.flux_in += t.flux_in
      d.reduit += reduit
    }

    const out_MS = Object.values(sortie).reduce((s, t) => s + t.MS, 0)
    const out_MV = Object.values(sortie).reduce((s, t) => s + t.MV, 0)
    const out_Q = dig2_Q_entree - dig2_MV_reduites / H.masse_volumique
    const out_conc = out_Q > 0 ? out_MS / out_Q : 0
    const in_MV_total = amont.MV + fraiches.MV
    const MV_reduites_total = dig1_MV_reduites + dig2_MV_reduites
    const reduction_globale = in_MV_total > 0 ? MV_reduites_total / in_MV_total : 0

    // =====================================================================
    // calcul_biogaz — les deux digesteurs contribuent
    // =====================================================================
    let densite_biogaz = 0
    for (let i = 0; i <= 1; i++) densite_biogaz += H.a_densite_biogaz[i] * Math.pow(biogaz_CH4, i)
    const production_specifique = densite_biogaz > 0 ? 1 / densite_biogaz : 0
    const biogaz_dig1 = production_specifique * dig1_MV_reduites
    const biogaz_dig2 = production_specifique * dig2_MV_reduites
    const biogaz_Nm3j = biogaz_dig1 + biogaz_dig2
    const DCO_biogaz = H.ratio_CH4_DCO > 0 ? (biogaz_CH4 * biogaz_Nm3j) / H.ratio_CH4_DCO : 0

    // =====================================================================
    // Retours du centrat et écriture de l'étape aval
    // =====================================================================
    const soluble_amont = amont.sol.map((v, i) => v + fraiches.sol[i])
    // la pollution soluble du digesteur 1 part majoritairement au centrat
    const part_centrat = dig1_Q + polymere_Q > 0 ? centrat_Q / (dig1_Q + polymere_Q) : 0
    const sol_centrat = soluble_amont.map((v) => v * part_centrat)
    const centrat_DCO = Object.values(apresDig1).reduce((s, t) => s + t.DCO, 0) * (1 - capture) + sol_centrat[SOL.DCO]
    const centrat_NK = Object.values(apresDig1).reduce((s, t) => s + t.NK, 0) * (1 - capture) + sol_centrat[SOL.NK]
    const centrat_Pt = Object.values(apresDig1).reduce((s, t) => s + t.Pt, 0) * (1 - capture) + sol_centrat[SOL.Pt]
    const vecteur = []
    vecteur[RET.Q] = centrat_Q
    vecteur[RET.MES] = centrat_MS
    vecteur[RET.DCO] = centrat_DCO
    vecteur[RET.DBO] = H.ratio_DBO_DCO_soluble * sol_centrat[SOL.DCO]
    vecteur[RET.NK] = centrat_NK
    vecteur[RET.NH4] = H.ratio_NH4_NK_soluble * sol_centrat[SOL.NK]
    vecteur[RET.Pt] = centrat_Pt
    const vecteur_soluble = []
    vecteur_soluble[RET.Q] = centrat_Q
    vecteur_soluble[RET.DCO] = sol_centrat[SOL.DCO]
    vecteur_soluble[RET.NK] = sol_centrat[SOL.NK]
    vecteur_soluble[RET.NH4] = H.ratio_NH4_NK_soluble * sol_centrat[SOL.NK]
    vecteur_soluble[RET.Pt] = sol_centrat[SOL.Pt]
    if (ctx.retours) ajouterRetour(ctx.retours, RET_ORIGINE.digestion, vecteur, vecteur_soluble)
    else for (let i = 1; i <= 8; i++) retour[i] += vecteur[i] || 0

    const sol_restant = soluble_amont.map((v) => v * (1 - part_centrat))
    for (const [jStr, d] of Object.entries(sortie)) {
      const j = Number(jStr)
      if (!(d.MS > 0)) continue
      const Qj = out_MS > 0 ? (d.MS / out_MS) * out_Q : 0
      if (!(Qj > 0)) continue
      const dst = table[etapeOut][j]
      const s = soluble[etapeOut][j]
      const part_MV = MV_reduites_total > 0 ? d.reduit / MV_reduites_total : 0

      let DCO = d.DCO - part_MV * DCO_biogaz
      let dissous = H.ratio_DCO_soluble_dig2 * DCO * corr_T_dig2
      s[SOL.DCO] = (dissous / Qj) * 1000 + (sol_restant[SOL.DCO] / out_Q) * 1000
      dst[P.ratio_DCO_MES] = Math.max(0, DCO - dissous) / d.MS
      dst[P.ratio_DBO_MES] = 0

      const ratio_NK_MV = d.MV + d.reduit > 0 ? d.NK / (d.MV + d.reduit) : 0
      dissous = H.ratio_NK_soluble_MVdeg * d.reduit * ratio_NK_MV
      s[SOL.NK] = (dissous / Qj) * 1000 + (sol_restant[SOL.NK] / out_Q) * 1000
      dst[P.ratio_NK_MES] = Math.max(0, d.NK - dissous) / d.MS

      dissous = d.Pt * H.ratio_P_soluble_P_entree * corr_T_dig2
      s[SOL.Pt] = (dissous / Qj) * 1000 + (sol_restant[SOL.Pt] / out_Q) * 1000
      dst[P.ratio_Pt_MES] = Math.max(0, d.Pt - dissous) / d.MS

      s[SOL.MS_soluble] = 0
      s[SOL.MV_soluble] = 0
      dst[P.origine] = d.origine
      dst[P.MES] = d.MS
      dst[P.Q] = Qj
      dst[P.MV_MES] = d.MV / d.MS
      dst[P.flux_in] = d.flux_in
      dst[P.verif_flux] = d.flux_in
    }

    // =====================================================================
    // calcul_bilan_thermique
    // =====================================================================
    const besoin_prechauffage = (debit_vapeur * kcalVersKWh(H.Cp_eau_kcal_kgC) * (T_prechauffage - T_initiale_eau)) / H.efficacite_echangeur
    const besoin_vaporisation = (debit_vapeur * kcalVersKWh(enthalpie_totale_kcal - H.Cp_eau_kcal_kgC * T_prechauffage)) / H.efficacite_echangeur
    const besoin_vapeur_total = besoin_prechauffage + besoin_vaporisation
    const pertesDe = (V) => H.pertes_thermiques[choices.climat][choices.isolation] * V
    const chauffage_dig1 = Math.max(0, (amont.Q * H.Cp_boues_kWh_m3C * (T_dig1 - T_entree)) / H.efficacite_echangeur)
    // le second digesteur reçoit des boues chaudes venues de la lyse
    const T_apres_lyse_refroidie = 55 // °C après les deux échangeurs
    const T_entree_dig2 = dig2_Q_entree > 0
      ? (apresLyse_Q * T_apres_lyse_refroidie + eau_dilution * T_entree) / dig2_Q_entree
      : T_apres_lyse_refroidie
    const chauffage_dig2 = Math.max(0, (dig2_Q_entree * H.Cp_boues_kWh_m3C * (T_dig2 - T_entree_dig2)) / H.efficacite_echangeur)
    const besoin_digestion = pertesDe(volume_dig1) + chauffage_dig1 + pertesDe(volume_dig2) + chauffage_dig2
    const besoin_thermique_total = besoin_vapeur_total + besoin_digestion
    // récupération au refroidissement des boues lysées, de 165 °C à 55 °C
    const energie_refroidissement = apresLyse_Q * H.masse_volumique * kcalVersKWh(H.Cp_eau_kcal_kgC) * (T_lyse - T_apres_lyse_refroidie)

    // =====================================================================
    // calcul_consommation_electrique
    // =====================================================================
    function elecAgitation(mode, volume, hauteur, taux) {
      if (mode === 'mecanique') return (H.agitation_W_m3 * H.agitation_hj * volume) / 1000
      if (mode === 'recirculation') {
        const Q = (taux * volume) / H.agitation_hj
        const puissance = (Q / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * 10 / 0.4
        return (puissance / rendementMoteur(puissance)) * H.agitation_hj
      }
      const Q = hauteur > 0 ? (H.brassage_biogaz_Nm_h * H.agitation_hj * volume) / hauteur : 0
      return (H.conso_surpresseur * Q * (hauteur + 2)) / 1000
    }
    const electricite_agitation =
      elecAgitation(choices.agitation_dig1, volume_dig1, p.hauteur_dig1 ?? 10, p.tx_recirculation_dig1 ?? 5)
      + elecAgitation(choices.agitation_dig2, volume_dig2, p.hauteur_dig2 ?? 10, p.tx_recirculation_dig2 ?? 5)

    const electricite_centrifugeuse = H.conso_elec_centrifugeuse_kWh_m3 * (dig1_Q + polymere_Q)

    let electricite_pompages = 0
    for (const [prefixe, debit] of [['pompage_dig1', dig1_Q], ['pompage_lyse', melange_Q], ['pompage_sortie', out_Q]]) {
      const r = repartitionPompage(debit, p[`${prefixe}_nb`], p[`${prefixe}_tps_fonctionnement`], forced, prefixe, 24)
      if (r.incoherence) warnings.push(`Incohérence sur le pompage « ${prefixe.replace('pompage_', '')} ».`)
      const rdt = rdtPompeBoues(r.Qu, p[`${prefixe}_P_refoulement`])
      electricite_pompages += elecPompage(r.Qu, r.nb, r.tps, p[`${prefixe}_P_refoulement`], rdt)
    }

    let electricite_stockage = 0
    let volume_gazometre = 0
    if (choices.biogaz_stock === 'oui') {
      const SRT_gaz = choices.valorisation_biogaz === 'chaudiere' ? 3 : 8
      volume_gazometre = (biogaz_Nm3j / CONST.NOMBRE_HEURE_PAR_JOUR) * SRT_gaz
      const soufflage = volume_gazometre < 1225 ? 300 : 500
      electricite_stockage = (H.conso_surpresseur * soufflage * 0.25 * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
    }
    const soufflage_valo = (biogaz_Nm3j * (p.ratio_biogaz_valorise ?? 1)) / CONST.NOMBRE_HEURE_PAR_JOUR
    const HMT_valo = choices.valorisation_biogaz === 'chaudiere' ? 1 : 3
    const electricite_surpresseur = (H.conso_surpresseur * soufflage_valo * HMT_valo * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000

    const total = electricite_agitation + electricite_centrifugeuse + electricite_pompages
      + electricite_stockage + electricite_surpresseur

    // comparaison avec une digestion simple sur le même flux
    // Le comparatif porte sur l'ensemble du flux admis, by-pass compris : sans
    // Exelys, toutes ces boues passeraient par une digestion simple.
    let MV_reduites_simple = 0
    for (const t of Object.values(amont.parType)) {
      MV_reduites_simple += (REDUCTION_MV_NON_LYSEES[t.origine] ?? 0) * t.MV * corr_SRT_dig1 * corr_T_dig1
    }
    for (const t of Object.values(fraiches.parType)) {
      MV_reduites_simple += (REDUCTION_MV_NON_LYSEES[t.origine] ?? 0) * t.MV * corr_SRT_dig1 * corr_T_dig1
    }
    const gain_DLD = MV_reduites_simple > 0 ? MV_reduites_total / MV_reduites_simple - 1 : 0

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MS en entrée', unit: 'kg/j', value: amont.MS + fraiches.MS },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: amont.Q + fraiches.Q },
        { key: 'in_MV', label: 'MV en entrée', unit: 'kg/j', value: in_MV_total },
        { key: 'V_dig1', label: 'Volume du digesteur 1', unit: 'm³', value: volume_dig1 },
        { key: 'SRT_dig1', label: 'Temps de séjour du digesteur 1', unit: 'j', value: SRT_dig1_reel },
        { key: 'red_dig1', label: 'MV réduites au digesteur 1', unit: 'kg/j', value: dig1_MV_reduites },
        { key: 'MV_sol', label: 'Part des MV solubilisées au digesteur 1', unit: '-', value: MV_solubles_ratio },
        { key: 'centri_MS', label: 'MS captées par la centrifugeuse', unit: 'kg/j', value: centri_MS },
        { key: 'centri_Q', label: 'Débit en sortie de centrifugeuse', unit: 'm³/j', value: centri_Q },
        { key: 'centri_conc', label: 'Siccité en sortie de centrifugeuse', unit: 'g/L', value: conc_centri },
        { key: 'poly', label: 'Polymère (matière active)', unit: 'kg/j', value: polymere_kgj },
        { key: 'centrat_Q', label: 'Centrat retourné en tête', unit: 'm³/j', value: centrat_Q },
        { key: 'centrat_MS', label: 'MS au centrat', unit: 'kg/j', value: centrat_MS },
        { key: 'T_sat', label: 'Température de la vapeur saturée', unit: '°C', value: T_saturation },
        { key: 'T_lyse', label: 'Température de lyse', unit: '°C', value: T_lyse },
        { key: 'vapeur', label: 'Consommation de vapeur', unit: 'kg/j', value: debit_vapeur },
        { key: 'vapeur_spec', label: 'Vapeur par m³ de boues lysées', unit: 'kg/m³', value: melange_Q > 0 ? debit_vapeur / melange_Q : 0 },
        { key: 'conc_lyse', label: 'Siccité après lyse', unit: 'g/L', value: apresLyse_conc },
        { key: 'dilution', label: "Eau de dilution", unit: 'm³/j', value: eau_dilution },
        { key: 'conc_dig2', label: 'Siccité en entrée du digesteur 2', unit: 'g/L', value: dig2_conc_entree },
        { key: 'V_dig2', label: 'Volume du digesteur 2', unit: 'm³', value: volume_dig2 },
        { key: 'SRT_dig2', label: 'Temps de séjour du digesteur 2', unit: 'j', value: SRT_dig2_reel },
        { key: 'red_dig2', label: 'MV réduites au digesteur 2', unit: 'kg/j', value: dig2_MV_reduites },
        { key: 'MV_reduites_total', label: 'MV réduites au total (parties en biogaz)', unit: 'kg/j', value: MV_reduites_total },
        { key: 'red_MV', label: 'Réduction globale des matières volatiles', unit: '-', value: reduction_globale },
        { key: 'gain_DLD', label: 'Gain du schéma DLD sur une digestion simple', unit: '-', value: gain_DLD },
        { key: 'out_MES', label: 'MS des boues digérées', unit: 'kg/j', value: out_MS },
        { key: 'out_Q', label: 'Débit de boues digérées', unit: 'm³/j', value: out_Q },
        { key: 'out_conc', label: 'Siccité des boues digérées', unit: 'g/L', value: out_conc },
        { key: 'out_MV_MES', label: 'MV/MS des boues digérées', unit: '-', value: out_MS > 0 ? out_MV / out_MS : 0 },
        { key: 'biogaz', label: 'Production totale de biogaz', unit: 'Nm³/j', value: biogaz_Nm3j },
        { key: 'biogaz_1', label: 'dont digesteur 1', unit: 'Nm³/j', value: biogaz_dig1 },
        { key: 'biogaz_2', label: 'dont digesteur 2', unit: 'Nm³/j', value: biogaz_dig2 },
        { key: 'CH4_Q', label: 'Méthane produit', unit: 'Nm³/j', value: biogaz_Nm3j * biogaz_CH4 },
        { key: 'DCO_biogaz', label: 'DCO méthanisée', unit: 'kg/j', value: DCO_biogaz },
        ...(volume_gazometre > 0 ? [{ key: 'V_gaz', label: 'Volume du gazomètre', unit: 'm³', value: volume_gazometre }] : []),
        { key: 'therm_vap', label: 'Besoin thermique pour la vapeur', unit: 'kWh/j', value: besoin_vapeur_total },
        { key: 'therm_dig', label: 'Besoin thermique des digesteurs', unit: 'kWh/j', value: besoin_digestion },
        { key: 'therm_total', label: 'Besoin thermique total', unit: 'kWh/j', value: besoin_thermique_total },
        { key: 'refroid', label: 'Chaleur récupérable au refroidissement des boues lysées', unit: 'kWh/j', value: energie_refroidissement },
      ],
      reactifs: { polymere: polymere_kgj },
      energie: {
        biogaz_Nm3j,
        biogaz_CH4,
        besoin_thermique_kWhj: besoin_thermique_total,
        recuperable_kWhj: energie_refroidissement,
        pression_vapeur_bars: P_vapeur,
        niveau: 'HT',
      },
      electricity: {
        total,
        fixed: electricite_agitation,
        detail: {
          agitation: electricite_agitation,
          centrifugeuse: electricite_centrifugeuse,
          pompages: electricite_pompages,
          stockage_biogaz: electricite_stockage,
          surpresseur_biogaz: electricite_surpresseur,
        },
      },
      warnings,
    }
  },
})
