// ---------------------------------------------------------------------------
// Port de z_Biothelys.cls — lyse thermique en batch, suivie de digestion.
//
// Biothelys n'est pas un simple prétraitement : la classe porte sa propre
// digestion, avec sa propre table de réduction des matières volatiles. Elle en
// tient d'ailleurs **deux**, et c'est là tout l'intérêt du procédé :
//
//   boue primaire        non lysée 0,638 → lysée 0,673
//   boue faible charge   non lysée 0,378 → lysée 0,472
//   boue MBR             non lysée 0,300 → lysée 0,394
//   graisses             non lysée 0,629 → lysée 0,765
//
// Le gain est d'autant plus marqué que la boue est réfractaire : négligeable
// sur une boue primaire déjà très dégradable, il dépasse 25 % en relatif sur
// une boue biologique d'aération prolongée. La table des non lysées est
// identique à celle de `digestion.js`, ce qui permet d'alimenter le digesteur
// avec un mélange de boues lysées et non lysées et de conserver la bonne
// cinétique pour chacune.
//
// La chaîne est la suivante :
//
//   entrée → lyse à 165 °C par injection de vapeur saturée (12,5 bars)
//          → échangeur de refroidissement à 120 °C
//          → mélange avec les boues non lysées éventuelles
//          → dilution à 100 g/L → digesteur → boues digérées
//
// La consommation de vapeur suit une corrélation linéaire sur la température
// des boues entrantes, et non un bilan enthalpique : c'est un procédé batch,
// dont une partie de la vapeur est récupérée en détente flash entre réacteurs.
// La corrélation n'est valable que jusqu'à 35 °C en entrée, au-delà de quoi la
// récupération flash diminue — le port le signale.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - la codigestion de coferments (BMP, part de graisses internes) n'est pas
//    portée, elle dépend d'une collection saisie dans un formulaire ;
//  - le raccordement au module PINCH d'intégration énergétique n'est pas fait :
//    les besoins et disponibilités thermiques sont calculés et exposés, mais
//    aucun consommateur ne s'en saisit ;
//  - le VBA range dans `RATIO_DCO_MES` et voisins des flux en kg/j avant de les
//    diviser par les MS en fin de routine ; le port tient des flux explicites.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, NB_TYPES } from '../core/sludge.js'
import { CONST, rendementMoteur } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees, repartitionPompage, rdtPompeBoues, elecPompage } from './_commun.js'

// Réduction des MV de référence, boues préalablement lysées (DCA, 31/05/11).
export const REDUCTION_MV_LYSEES = {
  I_simple: 0.673, I_reactif: 0.678,
  II_forte: 0.553, II_moyenne: 0.533, II_faible_EB: 0.472, II_faible_ED: 0.419,
  II_prolongee_EB: 0.451, II_prolongee_ED: 0.394, II_MBR: 0.394, II_MBBR: 0.451, II_HybAS: 0.451,
  II_biostyr_C: 0.493, II_biostyr_N: 0.493, II_biostyr_NDN: 0.493, II_biostyr_PDN: 0.493,
  III_decantation: 0.161, III_biostyr_N: 0.161, III_biostyr_PDN: 0.161,
  codigestion_graisses: 0.765,
}
// Réduction des MV de référence, boues non lysées — identique à digestion.js.
export const REDUCTION_MV_NON_LYSEES = {
  I_simple: 0.638, I_reactif: 0.614,
  II_forte: 0.546, II_moyenne: 0.501, II_faible_EB: 0.378, II_faible_ED: 0.32,
  II_prolongee_EB: 0.358, II_prolongee_ED: 0.3, II_MBR: 0.3, II_MBBR: 0.358, II_HybAS: 0.358,
  II_biostyr_C: 0.48, II_biostyr_N: 0.48, II_biostyr_NDN: 0.48, II_biostyr_PDN: 0.48,
  III_decantation: 0.157, III_biostyr_N: 0.157, III_biostyr_PDN: 0.157,
  codigestion_graisses: 0.629,
}

const H = {
  T_digestion_reference: 38, // °C
  T_digestion_min: 35,
  T_digestion_max: 42,
  SRT_reference: 15, // j
  a_correction_SRT: 1.25,
  b_correction_SRT: 0.7,
  a_densite_biogaz: [44 / 22.4, (16 - 44) / 22.4],
  ratio_CH4_DCO: 0.35, // Nm³ CH4 par kg de DCO méthanisée
  ratio_NK_soluble_MVdeg: 0.9,
  ratio_P_soluble_P_entree: 0.16,
  ratio_DCO_soluble_totale: 0.08,
  T_lyse: 165, // °C dans le réacteur, sous la température de vapeur saturée
  T_sortie_lyse_biothelys: 120, // °C après échangeur
  ratio_pertes_vapeur: 0.05,
  concentration_entree_digestion: 100, // g/L visés après dilution
  T_eau_dilution_max: 90, // °C
  // consommation spécifique de vapeur : a·T + b, en kWh par m³ de boues
  a_conso_vapeur: -1.6202,
  b_conso_vapeur: 193.88,
  T_max_validite_bilan: 35, // °C en entrée de lyse
  pertes_thermiques: {
    froid: { oui: 0.6, non: 3.72 },
    tempere: { oui: 0.36, non: 2.2 },
    chaud: { oui: 0.24, non: 1.44 },
  },
  efficacite_echangeur: 0.9,
  Cp_boues_kWh_m3C: 1.163,
  Cp_eau_kcal_kgC: 1,
  masse_volumique: 1000,
  // température de vapeur saturée : a1/(a2 − ln(a3·P)) + a4 − a5
  coef_T_saturation: [3816.44, 18.3, 750, 46.13, 273.15],
  // enthalpie de vaporisation à T : a0 + a1·T, en kcal/kg
  coef_enthalpie_vap: [616.87, -0.7542],
  agitation_W_m3: 7,
  agitation_hj: 24,
  brassage_biogaz_Nm_h: 1.25,
  conso_surpresseur: 4, // Wh/(Nm³·mCE)
  T_prechauffage_eau_vapeur: 85, // °C
  T_initiale_eau_vapeur: 15, // °C
  pression_vapeur_bars: 12.5,
}

/** facteur de correction lié à la température du digesteur (F_facteur_correction_temperature_digesteur) */
export function correctionTemperatureDigesteur(T) {
  if (T < 0) return 0
  if (T < 35) return T / 35
  if (T < 38) return 1
  if (T < 45) return -T / 14 + (3 + 5 / 7)
  if (T < 55) return (5 * T) / 91 - 1.975
  if (T < 58) return 1.05
  if (T < 80) return -T / 21 + (3 + 9 / 11)
  return 0
}
/** facteur de correction lié au temps de séjour */
export function correctionSRT(SRT, SRT_ref = H.SRT_reference) {
  const r = SRT_ref > 0 ? SRT / SRT_ref : 0
  const x = Math.pow(r, H.b_correction_SRT)
  return (H.a_correction_SRT * x) / (x + H.a_correction_SRT - 1)
}
/** température de la vapeur saturée à la pression absolue donnée (°C) */
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
const kcalVersKWh = (kcal) => (kcal * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE

export default defineSludgeNode({
  id: 'biothelys',
  label: 'Biothelys',
  short: 'Biothelys',
  family: 'stabilisation',
  vba: 'z_Biothelys.cls',
  etapeSortie: ETAPE.digerees,
  description:
    "Biothelys : lyse thermique en batch à 165 °C par injection de vapeur, suivie d'une digestion anaérobie. La lyse rend les matières volatiles nettement plus dégradables — le gain dépasse 25 % en relatif sur une boue biologique — au prix d'une consommation de vapeur importante.",
  choices: [
    { key: 'entree_lysee_1', label: 'Boues envoyées en lyse (entrée 1)', default: 'epaississeur_1', options: [...ENTREES, { value: 'aucune', label: 'aucune' }] },
    { key: 'entree_lysee_2', label: 'Boues envoyées en lyse (entrée 2)', default: 'aucune', options: [{ value: 'aucune', label: 'aucune' }, ...ENTREES] },
    { key: 'entree_non_lysee', label: 'Boues admises sans lyse (by-pass)', default: 'aucune', options: [{ value: 'aucune', label: 'aucune' }, ...ENTREES] },
    { key: 'agitation', label: 'Agitation du digesteur', default: 'biogaz', options: [
      { value: 'mecanique', label: 'mécanique' },
      { value: 'recirculation', label: 'par recirculation des boues' },
      { value: 'biogaz', label: 'par injection de biogaz' },
    ] },
    { key: 'isolation', label: 'Digesteur isolé', default: 'oui', options: [
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
    { key: 'temperature_entree_lyse', label: 'Température des boues en entrée de lyse', unit: '°C', group: 'Lyse', default: undefined, hint: "température de l'eau en conditions réelles" },
    { key: 'pression_vapeur', label: 'Pression absolue de la vapeur saturante', unit: 'bars', group: 'Lyse', default: 12.5 },
    { key: 'temperature_sortie_lyse', label: 'Température des boues en sortie de lyse', unit: '°C', group: 'Lyse', default: 120 },
    { key: 'debit_vapeur', label: 'Consommation de vapeur', unit: 'kg/j', group: 'Lyse', default: undefined, hint: 'corrélation sur la température des boues' },
    { key: 'temperature_initiale_eau_vapeur', label: "Température initiale de l'eau alimentant la chaudière", unit: '°C', group: 'Lyse', default: 15 },
    { key: 'temperature_prechauffage_eau_vapeur', label: "Température de préchauffage de l'eau", unit: '°C', group: 'Lyse', default: 85 },
    { key: 'concentration_entree_digestion', label: 'Siccité visée en entrée de digesteur', unit: 'g/L', group: 'Dilution', default: 100 },
    { key: 'debit_eau_dilution', label: "Débit d'eau de dilution", unit: 'm³/j', group: 'Dilution', default: undefined, hint: 'calculé sur la siccité visée' },
    { key: 'temperature_eau_dilution', label: "Température de l'eau de dilution", unit: '°C', group: 'Dilution', default: 15 },
    { key: 'SRT_design', label: 'Temps de séjour de dimensionnement', unit: 'j', group: 'Digestion', default: 15 },
    { key: 'volume_digesteur', label: 'Volume total de digestion', unit: 'm³', group: 'Digestion', default: undefined, hint: 'débit × temps de séjour' },
    { key: 'T_digesteur', label: 'Température du digesteur', unit: '°C', group: 'Digestion', default: 38, hint: 'entre 35 et 42 °C' },
    { key: 'hauteur_digesteur', label: 'Hauteur du digesteur', unit: 'm', group: 'Digestion', default: 10 },
    { key: 'tx_recirculation_agitation', label: 'Taux de recirculation pour agitation', unit: 'vol/j', group: 'Digestion', default: 5 },
    { key: 'reduction_MV', label: 'Réduction globale des MV', unit: '-', group: 'Digestion', default: undefined, hint: 'calculée par origine et par état de lyse' },
    { key: 'biogaz_teneur_CH4', label: 'Teneur en méthane du biogaz', unit: '-', group: 'Biogaz', default: 0.63 },
    { key: 'ratio_biogaz_valorise', label: 'Part du biogaz valorisée', unit: '-', group: 'Biogaz', default: 1 },
    { key: 'pompage_lyse_nb', label: 'Nombre de pompes lyse → digesteur', unit: 'u', group: 'Pompages', default: 2 },
    { key: 'pompage_lyse_P_refoulement', label: 'Pression de refoulement lyse → digesteur', unit: 'mCE', group: 'Pompages', default: 50 },
    { key: 'pompage_lyse_tps_fonctionnement', label: 'Durée de fonctionnement lyse → digesteur', unit: 'h/j', group: 'Pompages', default: 24 },
    { key: 'pompage_lyse_Q_unitaire', label: 'Débit unitaire lyse → digesteur', unit: 'm³/h', group: 'Pompages', default: undefined, hint: 'calculé si non forcé' },
    { key: 'pompage_digerees_nb', label: 'Nombre de pompes de boues digérées', unit: 'u', group: 'Pompages', default: 2 },
    { key: 'pompage_digerees_P_refoulement', label: 'Pression de refoulement des boues digérées', unit: 'mCE', group: 'Pompages', default: 50 },
    { key: 'pompage_digerees_tps_fonctionnement', label: 'Durée de fonctionnement des boues digérées', unit: 'h/j', group: 'Pompages', default: 24 },
    { key: 'pompage_digerees_Q_unitaire', label: 'Débit unitaire des boues digérées', unit: 'm³/h', group: 'Pompages', default: undefined, hint: 'calculé si non forcé' },
  ],

  compute(ctx) {
    const { site, table, soluble, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.digerees

    // ---- lecture_choix : deux entrées lysées, une entrée by-pass
    const entreesLysees = entreesDepuisChoix(choices, p, ['entree_lysee_1', 'entree_lysee_2'])
    const entreesNonLysees = entreesDepuisChoix(
      { entree_1: choices.entree_non_lysee }, { ratio_admis_1: p.ratio_admis_3 }, ['entree_1'])
    if (!entreesLysees.length && !entreesNonLysees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    // ---- attribution_valeur_par_defaut
    const T_digesteur = p.T_digesteur ?? H.T_digestion_reference
    if (T_digesteur < H.T_digestion_min || T_digesteur > H.T_digestion_max) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: [`Température de digestion (${T_digesteur} °C) hors de la plage admise par le procédé : ${H.T_digestion_min} à ${H.T_digestion_max} °C.`],
      }
    }
    const T_entree_lyse = f('temperature_entree_lyse') ?? site.T_eau_exploit
    const P_vapeur = p.pression_vapeur ?? H.pression_vapeur_bars
    const T_sortie_lyse = p.temperature_sortie_lyse ?? H.T_sortie_lyse_biothelys
    const SRT_design = p.SRT_design ?? H.SRT_reference
    const biogaz_CH4 = p.biogaz_teneur_CH4 ?? 0.63
    const conc_visee = p.concentration_entree_digestion ?? H.concentration_entree_digestion
    const T_initiale_eau = p.temperature_initiale_eau_vapeur ?? H.T_initiale_eau_vapeur
    const T_prechauffage = p.temperature_prechauffage_eau_vapeur ?? H.T_prechauffage_eau_vapeur

    // caractéristiques de la vapeur saturée à la pression retenue
    const T_saturation = temperatureSaturation(P_vapeur)
    let enthalpie_vaporisation = 0
    for (let i = 0; i <= 1; i++) enthalpie_vaporisation += H.coef_enthalpie_vap[i] * Math.pow(T_saturation, i)
    const enthalpie_totale_vapeur_kcal = H.Cp_eau_kcal_kgC * T_saturation + enthalpie_vaporisation

    if (T_entree_lyse > H.T_max_validite_bilan) {
      warnings.push(`Température des boues en entrée de lyse (${T_entree_lyse.toFixed(0)} °C) supérieure à ${H.T_max_validite_bilan} °C : la corrélation de consommation de vapeur sous-estime le besoin, la récupération par détente flash étant moindre.`)
    }

    // =====================================================================
    // Lecture des flux, séparés selon qu'ils passent ou non par la lyse
    // =====================================================================
    /** cumule un jeu d'entrées et renseigne la table de réduction applicable */
    function lireFlux(entrees, lysees) {
      const lu = lireEntrees(table, soluble, entrees)
      const parType = {}
      let MS = 0, Q = 0, MV = 0
      const sol = new Array(6).fill(0)
      for (const e of lu) {
        const { j, MES, Q: q, MV_MES, ratios, sol: s, src } = e
        if (!(MES > 0)) continue
        const origine = src[P.origine]
        const bareme = lysees ? REDUCTION_MV_LYSEES : REDUCTION_MV_NON_LYSEES
        const red = bareme[origine]
        if (red == null && origine) {
          warnings.push(`Origine de boue « ${origine} » sans taux de réduction des MV de référence : boue considérée non dégradable.`)
        }
        MS += MES; Q += q; MV += MES * MV_MES
        for (let k = 1; k <= 5; k++) sol[k] += (s[k] * q) / 1000
        if (!parType[j]) parType[j] = { MS: 0, Q: 0, MV: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, origine, reduction: red ?? 0 }
        const t = parType[j]
        t.MS += MES; t.Q += q; t.MV += MES * MV_MES
        t.DCO += MES * ratios.DCO
        t.DBO += MES * ratios.DBO
        t.NK += MES * ratios.NK
        t.Pt += MES * ratios.Pt
        t.flux_in += e.flux_in
      }
      for (const e of lu) e.src[P.verif_flux] -= e.flux_in
      return { lu, parType, MS, Q, MV, sol }
    }

    const lysees = lireFlux(entreesLysees, true)
    const nonLysees = lireFlux(entreesNonLysees, false)
    const in_MS = lysees.MS + nonLysees.MS
    const in_Q = lysees.Q + nonLysees.Q
    const in_MV = lysees.MV + nonLysees.MV
    if (!(in_MS > 0) || !(in_Q > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée du Biothelys."] }
    }

    // =====================================================================
    // calcul_bilan_masse_lyse — consommation de vapeur
    // =====================================================================
    // La corrélation donne une consommation spécifique en kWh par m³ de boues,
    // décroissante avec leur température ; la vapeur condensée dilue les boues.
    let debit_vapeur = 0
    if (lysees.Q > 0) {
      const conso_spec_kWh_m3 = H.a_conso_vapeur * T_entree_lyse + H.b_conso_vapeur
      debit_vapeur = f('debit_vapeur')
        ?? (conso_spec_kWh_m3 * lysees.Q) / kcalVersKWh(enthalpie_totale_vapeur_kcal)
    }
    const Q_apres_lyse = lysees.Q + debit_vapeur / H.masse_volumique
    const conc_apres_lyse = Q_apres_lyse > 0 ? lysees.MS / Q_apres_lyse : 0

    // =====================================================================
    // Mélange lysées + non lysées, puis dilution
    // =====================================================================
    const melange_MS = in_MS
    const melange_Q = Q_apres_lyse + nonLysees.Q
    const melange_conc = melange_Q > 0 ? melange_MS / melange_Q : 0
    // température du mélange, pondérée par les débits massiques
    const T_melange = melange_Q > 0
      ? (Q_apres_lyse * T_sortie_lyse + nonLysees.Q * T_entree_lyse) / melange_Q
      : T_sortie_lyse

    let eau_dilution = f('debit_eau_dilution')
    if (eau_dilution == null) {
      eau_dilution = conc_visee > melange_conc ? 0
        : melange_MS / conc_visee - melange_MS / melange_conc
    }
    const digesteur_Q_entree = melange_Q + eau_dilution
    const digesteur_conc_entree = digesteur_Q_entree > 0 ? melange_MS / digesteur_Q_entree : 0
    const T_eau_dilution = p.temperature_eau_dilution ?? 15
    const T_entree_digesteur = digesteur_Q_entree > 0
      ? (melange_Q * T_melange + eau_dilution * T_eau_dilution) / digesteur_Q_entree
      : T_melange

    // =====================================================================
    // calcul_bilan_masse_digestion
    // =====================================================================
    const volume_digesteur = f('volume_digesteur') ?? digesteur_Q_entree * SRT_design
    const SRT_reel = digesteur_Q_entree > 0 ? volume_digesteur / digesteur_Q_entree : SRT_design
    const corr_SRT = correctionSRT(SRT_reel)
    const corr_T = correctionTemperatureDigesteur(T_digesteur)

    // fusion des deux jeux de flux : chaque type conserve son taux propre
    const digestion = {}
    const ajouter = (source) => {
      for (const [jStr, t] of Object.entries(source.parType)) {
        const j = Number(jStr)
        if (!digestion[j]) digestion[j] = { MS: 0, Q: 0, MV: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, reduit: 0, origine: t.origine }
        const d = digestion[j]
        d.MS += t.MS; d.Q += t.Q; d.MV += t.MV
        d.DCO += t.DCO; d.DBO += t.DBO; d.NK += t.NK; d.Pt += t.Pt
        d.flux_in += t.flux_in
        d.reduit += t.reduction * t.MV * corr_SRT * corr_T
        d.origine = d.origine || t.origine
      }
    }
    ajouter(lysees)
    ajouter(nonLysees)

    let MV_reduites = 0
    let out_MS = 0, out_Q = 0, out_MV = 0
    for (const d of Object.values(digestion)) {
      MV_reduites += d.reduit
      out_MS += d.MS - d.reduit
      out_MV += d.MV - d.reduit
      // les MV détruites quittent la phase solide : le volume baisse d'autant
      out_Q += d.Q - d.reduit / H.masse_volumique
    }
    // la dilution s'ajoute au débit sortant
    out_Q += eau_dilution + (lysees.Q > 0 ? debit_vapeur / H.masse_volumique : 0)
    const reduction_globale = in_MV > 0 ? MV_reduites / in_MV : 0
    const out_conc = out_Q > 0 ? out_MS / out_Q : 0
    const out_MV_MES = out_MS > 0 ? out_MV / out_MS : 0

    // =====================================================================
    // calcul_biogaz
    // =====================================================================
    let densite_biogaz = 0
    for (let i = 0; i <= 1; i++) densite_biogaz += H.a_densite_biogaz[i] * Math.pow(biogaz_CH4, i)
    const production_specifique = densite_biogaz > 0 ? 1 / densite_biogaz : 0
    const biogaz_Nm3j = production_specifique * MV_reduites
    const DCO_biogaz = H.ratio_CH4_DCO > 0 ? (biogaz_CH4 * biogaz_Nm3j) / H.ratio_CH4_DCO : 0

    // =====================================================================
    // calcul_solubilisation_pendant_digestion + écriture de l'étape aval
    // =====================================================================
    const soluble_amont = lysees.sol.map((v, i) => v + nonLysees.sol[i])
    for (const [jStr, d] of Object.entries(digestion)) {
      const j = Number(jStr)
      const MS = d.MS - d.reduit
      const Q = d.Q - d.reduit / H.masse_volumique + (out_Q > 0 ? 0 : 0)
      // le débit de chaque type est proratisé sur le débit total sortant
      const Qj = out_MS > 0 ? (MS / out_MS) * out_Q : 0
      const dst = table[etapeOut][j]
      const s = soluble[etapeOut][j]
      if (!(MS > 0) || !(Qj > 0)) continue
      const part_MV_reduites = MV_reduites > 0 ? d.reduit / MV_reduites : 0

      // DCO : on retranche celle partie en biogaz, puis on solubilise
      let DCO = d.DCO - part_MV_reduites * DCO_biogaz
      let dissous = H.ratio_DCO_soluble_totale * DCO * corr_T
      s[SOL.DCO] = (dissous / Qj) * 1000 + (soluble_amont[SOL.DCO] / out_Q) * 1000
      dst[P.ratio_DCO_MES] = Math.max(0, DCO - dissous) / MS
      dst[P.ratio_DBO_MES] = 0 // hypothèse du classeur VBA : boues digérées non biodégradables

      // azote libéré au prorata des MV dégradées de ce type de boue
      const ratio_NK_MV = d.MV > 0 ? d.NK / d.MV : 0
      dissous = H.ratio_NK_soluble_MVdeg * d.reduit * ratio_NK_MV
      s[SOL.NK] = (dissous / Qj) * 1000 + (soluble_amont[SOL.NK] / out_Q) * 1000
      dst[P.ratio_NK_MES] = Math.max(0, d.NK - dissous) / MS

      dissous = d.Pt * H.ratio_P_soluble_P_entree * corr_T
      s[SOL.Pt] = (dissous / Qj) * 1000 + (soluble_amont[SOL.Pt] / out_Q) * 1000
      dst[P.ratio_Pt_MES] = Math.max(0, d.Pt - dissous) / MS

      s[SOL.MS_soluble] = 0
      s[SOL.MV_soluble] = 0
      dst[P.origine] = d.origine
      dst[P.MES] = MS
      dst[P.Q] = Qj
      dst[P.MV_MES] = (d.MV - d.reduit) / MS
      dst[P.flux_in] = d.flux_in
      dst[P.verif_flux] = d.flux_in
    }

    // =====================================================================
    // calcul_bilan_thermique
    // =====================================================================
    // production de vapeur : préchauffage de l'eau puis vaporisation
    const besoin_prechauffage = (debit_vapeur * kcalVersKWh(H.Cp_eau_kcal_kgC) * (T_prechauffage - T_initiale_eau)) / H.efficacite_echangeur
    const besoin_vaporisation = (debit_vapeur * kcalVersKWh(enthalpie_totale_vapeur_kcal - H.Cp_eau_kcal_kgC * T_prechauffage)) / H.efficacite_echangeur
    const besoin_vapeur_total = besoin_prechauffage + besoin_vaporisation
    // refroidissement des boues lysées, de 165 °C à la température de sortie
    const energie_refroidissement = lysees.Q > 0
      ? Q_apres_lyse * H.masse_volumique * kcalVersKWh(H.Cp_eau_kcal_kgC) * (H.T_lyse - T_sortie_lyse)
      : 0
    // maintien du digesteur : pertes de l'enceinte + chauffage des boues
    const pertes = H.pertes_thermiques[choices.climat][choices.isolation] * volume_digesteur
    const chauffage_boues = Math.max(0,
      (digesteur_Q_entree * H.Cp_boues_kWh_m3C * (T_digesteur - T_entree_digesteur)) / H.efficacite_echangeur)
    const besoin_digestion = pertes + chauffage_boues
    const besoin_thermique_total = besoin_vapeur_total + besoin_digestion

    // =====================================================================
    // calcul_consommation_electrique
    // =====================================================================
    let electricite_agitation = 0
    if (choices.agitation === 'mecanique') {
      electricite_agitation = (H.agitation_W_m3 * H.agitation_hj * volume_digesteur) / 1000
    } else if (choices.agitation === 'recirculation') {
      const Q = ((p.tx_recirculation_agitation ?? 5) * volume_digesteur) / H.agitation_hj
      const puissance = (Q / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * 10 / 0.4
      electricite_agitation = (puissance / rendementMoteur(puissance)) * H.agitation_hj
    } else {
      const hauteur = p.hauteur_digesteur ?? 10
      const Q = hauteur > 0 ? (H.brassage_biogaz_Nm_h * H.agitation_hj * volume_digesteur) / hauteur : 0
      electricite_agitation = (H.conso_surpresseur * Q * (hauteur + 2)) / 1000
    }

    let electricite_pompages = 0
    for (const [prefixe, debit] of [['pompage_lyse', melange_Q], ['pompage_digerees', out_Q]]) {
      const r = repartitionPompage(debit, p[`${prefixe}_nb`], p[`${prefixe}_tps_fonctionnement`], forced, prefixe, 24)
      if (r.incoherence) warnings.push(`Incohérence sur le pompage « ${prefixe.replace('pompage_', '')} ».`)
      const rdt = rdtPompeBoues(r.Qu, p[`${prefixe}_P_refoulement`])
      electricite_pompages += elecPompage(r.Qu, r.nb, r.tps, p[`${prefixe}_P_refoulement`], rdt)
    }

    let electricite_stockage = 0
    let volume_gazometre = 0
    if (choices.biogaz_stock === 'oui') {
      const SRT_gaz = choices.valorisation_biogaz === 'chaudiere' ? 3 : 8 // h
      volume_gazometre = (biogaz_Nm3j / CONST.NOMBRE_HEURE_PAR_JOUR) * SRT_gaz
      const soufflage = volume_gazometre < 1225 ? 300 : 500
      electricite_stockage = (H.conso_surpresseur * soufflage * 0.25 * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
    }
    const soufflage_valo = (biogaz_Nm3j * (p.ratio_biogaz_valorise ?? 1)) / CONST.NOMBRE_HEURE_PAR_JOUR
    const HMT_valo = choices.valorisation_biogaz === 'chaudiere' ? 1 : 3
    const electricite_surpresseur = (H.conso_surpresseur * soufflage_valo * HMT_valo * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000

    const total = electricite_agitation + electricite_pompages + electricite_stockage + electricite_surpresseur

    // comparaison avec ce qu'aurait donné une digestion sans lyse
    let MV_reduites_sans_lyse = 0
    for (const [jStr, t] of Object.entries(lysees.parType)) {
      const red = REDUCTION_MV_NON_LYSEES[t.origine] ?? 0
      MV_reduites_sans_lyse += red * t.MV * corr_SRT * corr_T
    }
    for (const t of Object.values(nonLysees.parType)) MV_reduites_sans_lyse += t.reduction * t.MV * corr_SRT * corr_T
    const gain_lyse = MV_reduites_sans_lyse > 0 ? MV_reduites / MV_reduites_sans_lyse - 1 : 0

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MS en entrée', unit: 'kg/j', value: in_MS },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: in_Q },
        { key: 'in_conc', label: 'Siccité en entrée', unit: 'g/L', value: in_MS / in_Q },
        { key: 'in_MV', label: 'MV en entrée', unit: 'kg/j', value: in_MV },
        { key: 'lys_MS', label: 'MS envoyées en lyse', unit: 'kg/j', value: lysees.MS },
        { key: 'lys_part', label: 'Part de boues lysées', unit: '-', value: in_MS > 0 ? lysees.MS / in_MS : 0 },
        { key: 'T_sat', label: 'Température de la vapeur saturée', unit: '°C', value: T_saturation },
        { key: 'H_vap', label: 'Enthalpie totale de la vapeur saturée', unit: 'kcal/kg', value: enthalpie_totale_vapeur_kcal },
        { key: 'vapeur', label: 'Consommation de vapeur', unit: 'kg/j', value: debit_vapeur },
        { key: 'vapeur_spec', label: 'Vapeur par m³ de boues lysées', unit: 'kg/m³', value: lysees.Q > 0 ? debit_vapeur / lysees.Q : 0 },
        { key: 'conc_lyse', label: 'Siccité après lyse', unit: 'g/L', value: conc_apres_lyse },
        { key: 'dilution', label: "Eau de dilution", unit: 'm³/j', value: eau_dilution },
        { key: 'conc_dig', label: 'Siccité en entrée de digesteur', unit: 'g/L', value: digesteur_conc_entree },
        { key: 'T_dig_in', label: 'Température en entrée de digesteur', unit: '°C', value: T_entree_digesteur },
        { key: 'V_dig', label: 'Volume de digestion', unit: 'm³', value: volume_digesteur },
        { key: 'SRT', label: 'Temps de séjour réel', unit: 'j', value: SRT_reel },
        { key: 'corr_T', label: 'Correctif de température', unit: '-', value: corr_T },
        { key: 'corr_SRT', label: 'Correctif de temps de séjour', unit: '-', value: corr_SRT },
        { key: 'red_MV', label: 'Réduction des matières volatiles', unit: '-', value: reduction_globale },
        { key: 'gain_lyse', label: 'Gain de la lyse sur la réduction des MV', unit: '-', value: gain_lyse },
        { key: 'out_MES', label: 'MS des boues digérées', unit: 'kg/j', value: out_MS },
        { key: 'out_Q', label: 'Débit de boues digérées', unit: 'm³/j', value: out_Q },
        { key: 'out_conc', label: 'Siccité des boues digérées', unit: 'g/L', value: out_conc },
        { key: 'out_MV_MES', label: 'MV/MS des boues digérées', unit: '-', value: out_MV_MES },
        { key: 'biogaz', label: 'Production de biogaz', unit: 'Nm³/j', value: biogaz_Nm3j },
        { key: 'CH4_Q', label: 'Méthane produit', unit: 'Nm³/j', value: biogaz_Nm3j * biogaz_CH4 },
        { key: 'DCO_biogaz', label: 'DCO méthanisée', unit: 'kg/j', value: DCO_biogaz },
        ...(volume_gazometre > 0 ? [{ key: 'V_gaz', label: 'Volume du gazomètre', unit: 'm³', value: volume_gazometre }] : []),
        { key: 'therm_vap', label: 'Besoin thermique pour la vapeur', unit: 'kWh/j', value: besoin_vapeur_total },
        { key: 'therm_prech', label: "dont préchauffage de l'eau", unit: 'kWh/j', value: besoin_prechauffage },
        { key: 'therm_vapo', label: 'dont vaporisation', unit: 'kWh/j', value: besoin_vaporisation },
        { key: 'therm_dig', label: 'Besoin thermique du digesteur', unit: 'kWh/j', value: besoin_digestion },
        { key: 'therm_total', label: 'Besoin thermique total', unit: 'kWh/j', value: besoin_thermique_total },
        { key: 'refroid', label: 'Chaleur récupérable au refroidissement des boues lysées', unit: 'kWh/j', value: energie_refroidissement },
      ],
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
          pompages: electricite_pompages,
          stockage_biogaz: electricite_stockage,
          surpresseur_biogaz: electricite_surpresseur,
        },
      },
      warnings,
    }
  },
})
