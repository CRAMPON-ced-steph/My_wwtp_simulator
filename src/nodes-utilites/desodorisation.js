// ---------------------------------------------------------------------------
// Port de z_Desodorisation_biologique.cls et z_Desodorisation_chimique.cls.
//
// Les deux classes partagent le calcul du débit d'air vicié et la ventilation ;
// elles ne diffèrent que sur le traitement lui-même.
//
//   Biologique : l'air traverse un garnissage colonisé, arrosé en continu. Pas
//                de réactif, seulement une pompe de lavage.
//   Chimique   : jusqu'à quatre tours de lavage en série, chacune ciblant des
//                composés différents :
//                  acide (H2SO4)         → NH3 et azote organique
//                  oxydo-basique pH 9    → H2S et mercaptans, à la soude + Cl2
//                  oxydo-basique pH 11   → le reliquat, plus agressif sur le CO2
//                  neutre (NaHSO3)       → neutralise le chlore résiduel
//
// Le point commun aux deux : le débit d'air vicié se déduit du nombre de
// renouvellements horaires appliqués au débit nominal — 13 fois pour une
// couverture partielle des ouvrages, 42 fois pour une couverture totale. C'est
// le facteur qui gouverne tout le dimensionnement.
//
// Une subtilité du VBA, portée telle quelle : la teneur en H2S de l'air vicié
// n'est pas une donnée mais se déduit des sulfures strippés par la file eau,
// rapportés au débit d'air. Le calcul relie donc directement la qualité de
// l'eau brute au dimensionnement de la désodorisation.
//
// Écart au VBA, volontaire et signalé (voir README) : la tour oxydo-basique
// pH 11 remet à zéro H2S et mercaptans comme la pH 9 le fait déjà ; enchaîner
// les deux ne consomme donc de réactif qu'à la première. Reproduit, et signalé
// quand les deux tours sont demandées.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { CONST } from '../core/hypotheses.js'

const H = {
  renouvellement_couverture_partielle: 13, // volumes de Q nominal par jour
  renouvellement_couverture_totale: 42,
  ventilateur_rdt_machine: 0.7,
  ventilateur_rdt_moteur: 0.8,
  pompe_lavage_rdt_machine: 0.8,
  // désodorisation chimique : dosages en g de réactif par g de polluant
  acide_H2SO4_NH3: 2.9,
  acide_H2SO4_Norga: 3.5,
  pH9_NaOH_H2S: 2.5,
  pH9_NaOH_RSH: 4.1,
  pH9_NaOH_CO2: 30, // g de NaOH par m³ d'air, pour le CO2 atmosphérique
  pH11_NaOH_H2S: 3,
  pH11_NaOH_RSH: 3,
  pH11_NaOH_CO2: 100,
  Cl2_H2S: 9.5,
  Cl2_RSH: 11,
  NaHSO3_Cl2: 1.85,
  residuel_Cl2_mg_Nm3: 1,
  T_reference_K: 273.15,
}

const COUVERTURE = [
  { value: 'partielle', label: 'couverture partielle des ouvrages' },
  { value: 'totale', label: 'couverture totale des ouvrages' },
]
const OUI_NON = [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }]

/** paramètres communs de ventilation et de soufflage */
const PARAMS_VENTILATION = [
  { key: 'ventilation_tps_fonctionnement', label: 'Durée de fonctionnement', unit: 'h/j', group: 'Ventilation', default: 24 },
  { key: 'air_vicie_Q_Nm3h', label: "Débit d'air vicié", unit: 'Nm³/h', group: 'Ventilation', default: undefined, hint: 'renouvellements × débit nominal' },
  { key: 'air_vicie_T', label: "Température de l'air à traiter", unit: '°C', group: 'Ventilation', default: 15 },
  { key: 'ventilation_P_reseau', label: 'Pression du réseau de ventilation', unit: 'Pa', group: 'Ventilation', default: 2500 },
  { key: 'ventilation_rdt_global', label: 'Rendement global du ventilateur', unit: '-', group: 'Ventilation', default: 0.7 * 0.8 },
  { key: 'soufflage_P_reseau', label: 'Pression du réseau de soufflage', unit: 'Pa', group: 'Soufflage', default: 300 },
  { key: 'soufflage_Tair', label: "Température de l'air avant soufflage", unit: '°C', group: 'Soufflage', default: undefined, hint: "température de l'air vicié" },
  { key: 'soufflage_rdt_global', label: 'Rendement global du ventilateur de soufflage', unit: '-', group: 'Soufflage', default: 0.7 * 0.8 },
]

/**
 * Débit d'air vicié et électricité de ventilation, communs aux deux procédés.
 * La puissance d'un ventilateur est le produit du débit volumique réel par la
 * pression du réseau : le débit normalisé est donc corrigé de la température.
 */
function ventilation(ctx, p, f, choices) {
  const tps = p.ventilation_tps_fonctionnement ?? 24
  const renouvellement = choices.couverture === 'totale'
    ? H.renouvellement_couverture_totale
    : H.renouvellement_couverture_partielle
  const Q_Nm3h = f('air_vicie_Q_Nm3h') ?? (tps > 0 ? (renouvellement * ctx.contexte.Q_nominal) / tps : 0)
  const T = p.air_vicie_T ?? 15
  const T_soufflage = f('soufflage_Tair') ?? T
  const corrT = (t) => (t + H.T_reference_K) / H.T_reference_K

  const elec_ventilation = tps * (((Q_Nm3h * corrT(T)) / CONST.NOMBRE_SECONDE_PAR_HEURE) * (p.ventilation_P_reseau ?? 2500) / (p.ventilation_rdt_global ?? 0.56)) / 1000
  const elec_soufflage = choices.soufflage_air === 'oui'
    ? tps * (((Q_Nm3h * corrT(T_soufflage)) / CONST.NOMBRE_SECONDE_PAR_HEURE) * (p.soufflage_P_reseau ?? 300) / (p.soufflage_rdt_global ?? 0.56)) / 1000
    : 0
  return { tps, renouvellement, Q_Nm3h, T, elec_ventilation, elec_soufflage }
}

/** rendement global d'une pompe de lavage : machine 0,8 × moteur */
function rdtPompeLavage(Q_m3h, HMT) {
  const puissance = ((Q_m3h / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * HMT) / H.pompe_lavage_rdt_machine
  const rdt_moteur = puissance < 60 ? (5.7195 * Math.log(puissance) + 72.682) / 100 : 0.961
  return H.pompe_lavage_rdt_machine * rdt_moteur
}

// ---------------------------------------------------------------------------
export const desodorisationBiologique = defineUtilityNode({
  id: 'desodorisation-bio',
  label: 'Désodorisation biologique',
  short: 'Désodo bio',
  family: 'desodorisation',
  vba: 'z_Desodorisation_biologique.cls',
  description:
    "Traitement de l'air vicié sur garnissage colonisé, arrosé en continu. Sans réactif : seules la ventilation et la pompe de lavage consomment.",
  choices: [
    { key: 'couverture', label: 'Couverture des ouvrages', default: 'partielle', options: COUVERTURE },
    { key: 'soufflage_air', label: "Soufflage de l'air traité", default: 'non', options: OUI_NON },
  ],
  params: [
    ...PARAMS_VENTILATION,
    { key: 'pompe_lavage_rate_Leau_m3air', label: "Ratio d'eau de lavage", unit: 'L/m³ d\'air', group: 'Lavage', default: 0.1 },
    { key: 'pompe_lavage_HMT', label: 'HMT des pompes de lavage', unit: 'mCE', group: 'Lavage', default: 20 },
    { key: 'pompe_lavage_rdt', label: 'Rendement des pompes de lavage', unit: '-', group: 'Lavage', default: undefined, hint: 'machine 0,8 × moteur' },
  ],

  compute(ctx) {
    const { p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const v = ventilation(ctx, p, f, choices)
    if (!(v.Q_Nm3h > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Débit d'air vicié nul : vérifier le débit nominal de la station."] }
    }

    const Q_lavage = ((p.pompe_lavage_rate_Leau_m3air ?? 0.1) / 1000) * v.Q_Nm3h
    const HMT = p.pompe_lavage_HMT ?? 20
    const rdt = f('pompe_lavage_rdt') ?? rdtPompeLavage(Q_lavage, HMT)
    const elec_lavage = rdt > 0
      ? ((Q_lavage / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * HMT / rdt) * v.tps
      : 0
    const total = v.elec_ventilation + v.elec_soufflage + elec_lavage

    return {
      results: [
        { key: 'renouv', label: 'Renouvellements appliqués au débit nominal', unit: 'vol/j', value: v.renouvellement },
        { key: 'air', label: "Débit d'air vicié", unit: 'Nm³/h', value: v.Q_Nm3h },
        { key: 'air_j', label: "Débit d'air vicié", unit: 'Nm³/j', value: v.Q_Nm3h * v.tps },
        { key: 'T', label: "Température de l'air", unit: '°C', value: v.T },
        { key: 'Q_lavage', label: "Débit d'eau de lavage", unit: 'm³/h', value: Q_lavage },
        { key: 'elec_spec', label: 'Consommation par Nm³ traité', unit: 'kWh/1000 Nm³', value: v.Q_Nm3h * v.tps > 0 ? (total / (v.Q_Nm3h * v.tps)) * 1000 : 0 },
      ],
      electricity: {
        total,
        fixed: total,
        detail: { ventilation: v.elec_ventilation, soufflage: v.elec_soufflage, lavage: elec_lavage },
      },
      warnings,
    }
  },
})

// ---------------------------------------------------------------------------
export const desodorisationChimique = defineUtilityNode({
  id: 'desodorisation-chimique',
  label: 'Désodorisation chimique',
  short: 'Désodo chim',
  family: 'desodorisation',
  vba: 'z_Desodorisation_chimique.cls',
  description:
    "Traitement de l'air vicié par lavage chimique en tours successives : acide pour l'ammoniac, oxydo-basique pour les composés soufrés, neutre pour le chlore résiduel.",
  choices: [
    { key: 'couverture', label: 'Couverture des ouvrages', default: 'partielle', options: COUVERTURE },
    { key: 'soufflage_air', label: "Soufflage de l'air traité", default: 'non', options: OUI_NON },
    { key: 'tour_acide', label: 'Tour acide (H2SO4)', default: 'oui', options: OUI_NON },
    { key: 'tour_pH9', label: 'Tour oxydo-basique pH 9', default: 'oui', options: OUI_NON },
    { key: 'tour_pH11', label: 'Tour oxydo-basique pH 11', default: 'non', options: OUI_NON },
    { key: 'tour_neutre', label: 'Tour neutre (bisulfite)', default: 'oui', options: OUI_NON },
  ],
  params: [
    ...PARAMS_VENTILATION,
    { key: 'air_vicie_NH3', label: "NH3 dans l'air vicié", unit: 'mg/Nm³', group: 'Charge polluante', default: 10 },
    { key: 'air_vicie_Norga', label: "Azote organique dans l'air vicié", unit: 'mg/Nm³', group: 'Charge polluante', default: 2 },
    { key: 'air_vicie_RSH', label: "Mercaptans dans l'air vicié", unit: 'mg/Nm³', group: 'Charge polluante', default: 1 },
    { key: 'air_vicie_H2S', label: "H2S dans l'air vicié", unit: 'mg/Nm³', group: 'Charge polluante', default: undefined, hint: 'déduit des sulfures strippés par la file eau' },
    { key: 'H2SO4_pur', label: 'Consommation de H2SO4 pur', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculée sur NH3 et azote organique' },
    { key: 'NaOH_pur_pH9', label: 'Consommation de NaOH pur (tour pH 9)', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculée sur H2S, mercaptans et CO2' },
    { key: 'NaOH_pur_pH11', label: 'Consommation de NaOH pur (tour pH 11)', unit: 'kg/j', group: 'Réactifs', default: undefined },
    { key: 'Cl2_pur', label: 'Consommation de Cl2 pur', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculée sur H2S et mercaptans' },
    { key: 'NaHSO3_pur', label: 'Consommation de NaHSO3 pur', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'neutralisation du chlore résiduel' },
    { key: 'pompe_lavage_rate_Leau_m3air', label: "Ratio d'eau de lavage par tour", unit: 'L/m³ d\'air', group: 'Lavage', default: 0.1 },
    { key: 'pompe_lavage_HMT', label: 'HMT des pompes de lavage', unit: 'mCE', group: 'Lavage', default: 20 },
    { key: 'pompe_lavage_rdt', label: 'Rendement des pompes de lavage', unit: '-', group: 'Lavage', default: undefined, hint: 'machine 0,8 × moteur' },
  ],

  compute(ctx) {
    const { contexte, p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const on = (v) => v === 'oui'
    const v = ventilation(ctx, p, f, choices)
    if (!(v.Q_Nm3h > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Débit d'air vicié nul : vérifier le débit nominal de la station."] }
    }

    // ---- charge polluante de l'air vicié
    // Le H2S n'est pas une donnée : il vient des sulfures strippés par la file
    // eau, rapportés au débit d'air. La qualité de l'eau brute gouverne donc
    // directement le dimensionnement de la désodorisation.
    const H2S_mgNm3 = f('air_vicie_H2S')
      ?? (v.Q_Nm3h > 0 ? ((contexte.HS_strippe_kgj / CONST.NOMBRE_HEURE_PAR_JOUR) / v.Q_Nm3h) * 1e6 : 0)
    let NH3_gh = (v.Q_Nm3h * (p.air_vicie_NH3 ?? 10)) / 1000
    let Norga_gh = (v.Q_Nm3h * (p.air_vicie_Norga ?? 2)) / 1000
    let RSH_gh = (v.Q_Nm3h * (p.air_vicie_RSH ?? 1)) / 1000
    let H2S_gh = (v.Q_Nm3h * H2S_mgNm3) / 1000
    const kgj = (gh) => (gh / 1000) * CONST.NOMBRE_HEURE_PAR_JOUR

    let H2SO4 = 0, NaOH_pH9 = 0, NaOH_pH11 = 0, Cl2 = 0, NaHSO3 = 0
    let nb_tours = 0

    // ---- tour acide : ammoniac et azote organique
    if (on(choices.tour_acide)) {
      nb_tours += 1
      H2SO4 = f('H2SO4_pur')
        ?? H.acide_H2SO4_NH3 * kgj(NH3_gh) + H.acide_H2SO4_Norga * kgj(Norga_gh)
      NH3_gh = 0
      Norga_gh = 0
    }
    // ---- tour oxydo-basique pH 9 : composés soufrés
    if (on(choices.tour_pH9)) {
      nb_tours += 1
      NaOH_pH9 = f('NaOH_pur_pH9')
        ?? H.pH9_NaOH_H2S * kgj(H2S_gh) + H.pH9_NaOH_RSH * kgj(RSH_gh)
          + H.pH9_NaOH_CO2 * kgj(v.Q_Nm3h / 1000)
      Cl2 = f('Cl2_pur') ?? H.Cl2_H2S * kgj(H2S_gh) + H.Cl2_RSH * kgj(RSH_gh)
      H2S_gh = 0
      RSH_gh = 0
    }
    // ---- tour oxydo-basique pH 11 : le reliquat
    if (on(choices.tour_pH11)) {
      nb_tours += 1
      NaOH_pH11 = f('NaOH_pur_pH11')
        ?? H.pH11_NaOH_H2S * kgj(H2S_gh) + H.pH11_NaOH_RSH * kgj(RSH_gh)
          + H.pH11_NaOH_CO2 * kgj(v.Q_Nm3h / 1000)
      if (f('Cl2_pur') == null) Cl2 += H.Cl2_H2S * kgj(H2S_gh) + H.Cl2_RSH * kgj(RSH_gh)
      if (on(choices.tour_pH9)) {
        warnings.push("La tour pH 9 a déjà abattu la totalité des composés soufrés : la tour pH 11 ne consomme de la soude que pour le CO2, conformément au classeur d'origine.")
      }
      H2S_gh = 0
      RSH_gh = 0
    }
    // ---- tour neutre : neutralisation du chlore résiduel
    if (on(choices.tour_neutre)) {
      nb_tours += 1
      NaHSO3 = f('NaHSO3_pur')
        ?? H.NaHSO3_Cl2 * H.residuel_Cl2_mg_Nm3 * kgj(v.Q_Nm3h / 1000)
    }
    if (nb_tours === 0) warnings.push("Aucune tour de lavage retenue : l'air vicié n'est pas traité.")

    // ---- électricité : une pompe de lavage par tour
    const Q_lavage = ((p.pompe_lavage_rate_Leau_m3air ?? 0.1) / 1000) * v.Q_Nm3h
    const HMT = p.pompe_lavage_HMT ?? 20
    const rdt = f('pompe_lavage_rdt') ?? rdtPompeLavage(Q_lavage, HMT)
    const elec_lavage = rdt > 0
      ? ((Q_lavage / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * HMT / rdt) * v.tps * nb_tours
      : 0
    const total = v.elec_ventilation + v.elec_soufflage + elec_lavage

    const reactifs = {}
    if (H2SO4 > 0) reactifs.H2SO4 = H2SO4
    if (NaOH_pH9 + NaOH_pH11 > 0) reactifs.NaOH = NaOH_pH9 + NaOH_pH11
    if (Cl2 > 0) reactifs.Cl2 = Cl2
    if (NaHSO3 > 0) reactifs.NaHSO3 = NaHSO3

    return {
      results: [
        { key: 'renouv', label: 'Renouvellements appliqués au débit nominal', unit: 'vol/j', value: v.renouvellement },
        { key: 'air', label: "Débit d'air vicié", unit: 'Nm³/h', value: v.Q_Nm3h },
        { key: 'H2S', label: "H2S dans l'air vicié", unit: 'mg/Nm³', value: H2S_mgNm3 },
        { key: 'nb_tours', label: 'Nombre de tours de lavage', unit: 'u', value: nb_tours },
        ...(H2SO4 > 0 ? [{ key: 'H2SO4', label: 'Acide sulfurique pur', unit: 'kg/j', value: H2SO4 }] : []),
        ...(NaOH_pH9 > 0 ? [{ key: 'NaOH9', label: 'Soude pure (tour pH 9)', unit: 'kg/j', value: NaOH_pH9 }] : []),
        ...(NaOH_pH11 > 0 ? [{ key: 'NaOH11', label: 'Soude pure (tour pH 11)', unit: 'kg/j', value: NaOH_pH11 }] : []),
        ...(Cl2 > 0 ? [{ key: 'Cl2', label: 'Chlore pur', unit: 'kg/j', value: Cl2 }] : []),
        ...(NaHSO3 > 0 ? [{ key: 'NaHSO3', label: 'Bisulfite de sodium pur', unit: 'kg/j', value: NaHSO3 }] : []),
        { key: 'elec_spec', label: 'Consommation par Nm³ traité', unit: 'kWh/1000 Nm³', value: v.Q_Nm3h * v.tps > 0 ? (total / (v.Q_Nm3h * v.tps)) * 1000 : 0 },
      ],
      reactifs,
      electricity: {
        total,
        fixed: total,
        detail: { ventilation: v.elec_ventilation, soufflage: v.elec_soufflage, lavage: elec_lavage },
      },
      warnings,
    }
  },
})
