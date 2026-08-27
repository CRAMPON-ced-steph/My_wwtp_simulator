// Port de E1_BA_forte_charge.cls (Boue_Activee_Forte_Charge, attribution_valeur_par_defaut,
// dimensionnement, fonctionnement_reel, calcul_consommation_electrique)
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, HYP, besoinsO2HS, facteurK } from '../core/hypotheses.js'

// Reproduit le comportement effectif du VBA : les "If T<12 / If T<18" successifs
// (sans ElseIf) font que la valeur T<12 est écrasée par la valeur T<18.
// Passer à false pour obtenir le comportement vraisemblablement voulu (3 / 2,5 / 1,5).
export const VBA_BUG_COMPAT = true

const H = {
  T_reference: 12,
  G_reference: 1, // j
  correctif_T: 1.072,
  a1_Cm: 0.8778,
  a0_Cm: -0.0631,
  FtoM_lim: HYP.BA_rdtDBO_FtoM_max, // 1.9 kgDBO/(kgMES·j)
  rdt_DCO: 0.6,
  O2resp: HYP.BA_besoinO2_DBO_resp, // a0 0.56, a1 0.15, a2 0.17, corrT 1.072, Tref 15
  alpha0_autres: 0.75,
  alpha_fb: { a0: HYP.BA_alpha_finebulle_a0, aMVS: HYP.BA_alpha_finebulle_aMVS, aG: HYP.BA_alpha_finebulle_aGeq, max: 1, min: 0.05 },
  alpha_corr: { a0: HYP.BA_alpha_correction_MES_a0, a1: HYP.BA_alpha_correction_MES_a1, Cref: HYP.BA_alpha_correction_MES_Cref },
  rdt_dissolution: { fines: HYP.BA_rdt_dissolution_O2_eau_claire.fines_bulles, moyennes: HYP.BA_rdt_dissolution_O2_eau_claire.moyennes_bulles },
  ratio_kgO2_Nm3air: HYP.ratio_kgO2_Nm3air,
  Pmax_roots: HYP.surpresseur_roots_Pmax_conseillee,
  ratio_P_synthese: 0.01, // kgP/kgDBO éliminée
  clarif_diametre_limite: 32,
}

const AERATEURS = [
  { value: 'fines', label: 'diffuseur fines bulles', hauteur: 6, insufflation: true },
  { value: 'moyennes', label: 'diffuseur moyennes bulles', hauteur: 6, insufflation: true },
  { value: 'brosses', label: "brosses d'aération", hauteur: 3, ASB: HYP.BA_ASB.brosses },
  { value: 'turbines_lentes', label: 'turbines lentes', hauteur: 3, ASB: HYP.BA_ASB.turbines_lentes },
  { value: 'turbines_rapides', label: 'turbines rapides', hauteur: 3, ASB: HYP.BA_ASB.turbines_rapides },
]
const aer = (c) => AERATEURS.find((a) => a.value === c.choices.aerateur)

function O2DissousDefaut(c) {
  const T = c.site.T_eau_exploit
  switch (c.choices.regulation) {
    case 'horloge': return 4
    case 'sans_variateur': return T < 14 ? 2 : 1.5
    default: return T < 14 ? 1.5 : 1
  }
}
const mesBassinDefaut = (T) => (VBA_BUG_COMPAT ? (T < 18 ? 2.5 : 1.5) : T < 12 ? 3 : T < 18 ? 2.5 : 1.5)
const mvMesDefaut = (T) => (VBA_BUG_COMPAT ? (T < 18 ? 0.72 : 0.75) : T < 12 ? 0.7 : T < 18 ? 0.72 : 0.75)

const polyRdtDBO = (coef, x) => coef.reduce((acc, a, i) => acc + a * Math.pow(x, i), 0)

export default defineNode({
  id: 'ba-forte-charge',
  label: 'Boue activée forte charge',
  short: 'BA forte charge',
  family: 'secondaire',
  vba: 'E1_BA_forte_charge.cls',
  description: 'Boue activée forte charge (âge de boues ≈ 1 j à 12 °C). Traitement du carbone seul ; nitrification non modélisée. Clarificateur dimensionné sur la pointe temps de pluie.',
  choices: [
    { key: 'aerateur', label: "Type d'aération", default: 'fines', options: AERATEURS.map(({ value, label }) => ({ value, label })) },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [{ value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' }] },
    { key: 'regulation', label: "Régulation de l'aération", default: 'variateur', options: [{ value: 'horloge', label: 'régulation sur horloge' }, { value: 'sans_variateur', label: 'sans variateur de fréquence' }, { value: 'variateur', label: 'classique – avec variateur' }, { value: 'avance', label: 'avancé' }] },
    { key: 'racleur', label: 'Type de râcleur du clarificateur', default: 'racle', options: [{ value: 'racle', label: 'râclé' }, { value: 'racle_suce', label: 'râclé-sucé' }, { value: 'kruger', label: 'Kruger' }] },
  ],
  params: [
    // nominal
    { key: 'nominal_G', label: 'Âge de boues design', unit: 'j', group: 'Nominal', default: (c) => H.G_reference * Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_design) },
    { key: 'nominal_MES_bassin', label: 'MES design dans les bassins', unit: 'g/L', group: 'Nominal', default: (c) => mesBassinDefaut(c.site.T_eau_design) },
    { key: 'nominal_MV_MES', label: 'MV/MES des boues design', unit: '-', group: 'Nominal', default: (c) => mvMesDefaut(c.site.T_eau_design) },
    { key: 'volume_bassins', label: 'Volume des bassins', unit: 'm³', group: 'Nominal', default: undefined, hint: 'calculé si non forcé' },
    // réel
    { key: 'reel_G', label: 'Âge de boues réel', unit: 'j', group: 'Réel', default: (c) => (c.p.nominal_G / Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_design)) * Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_exploit) },
    { key: 'reel_MV_MES', label: 'MV/MES des boues réel', unit: '-', group: 'Réel', default: (c) => mvMesDefaut(c.site.T_eau_design) },
    { key: 'reel_MES_bassin', label: 'MES réel dans les bassins', unit: 'g/L', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_dissous', label: 'O2 dissous moyen', unit: 'mg/L', group: 'Réel', default: O2DissousDefaut },
    { key: 'sortie_DBO', label: 'DBO5 en sortie (réel)', unit: 'mg/L', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
    // aération
    { key: 'hauteur_bassin', label: "Hauteur d'eau du bassin", unit: 'm', group: 'Aération', default: (c) => aer(c).hauteur },
    { key: 'O2_besoin', label: 'Besoin en O2', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_facteur_alpha', label: 'Facteur alpha', unit: '-', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_rdt_transfert', label: 'Rendement de dissolution eau claire', unit: '%/m', group: 'Aération', default: (c) => (c.choices.aerateur === 'moyennes' ? H.rdt_dissolution.moyennes : H.rdt_dissolution.fines) },
    { key: 'air_Q_Nm3j', label: "Débit d'air", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'diffuseur_encrassement', label: 'Durée depuis dernier nettoyage des diffuseurs', unit: 'an', group: 'Aération', default: 0 },
    { key: 'air_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Aération', default: (c) => c.p.hauteur_bassin + 2 + 0.25 * c.p.diffuseur_encrassement },
    { key: 'surpresseur_conso_spec', label: 'Conso spécifique surpresseur', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'ASB_eau_claire', label: 'ASB eau claire (aérateurs de surface)', unit: 'kg O2/kWh', group: 'Aération', default: (c) => aer(c).ASB ?? 0 },
    // recirculation
    { key: 'recirculation_taux', label: 'Taux de recirculation', unit: '-', group: 'Recirculation', default: 1 },
    { key: 'recirculation_P_refoulement', label: 'Pression de refoulement recirculation', unit: 'mCE', group: 'Recirculation', default: 5 },
    { key: 'recirculation_pompe_rdt', label: 'Rendement global pompes recirculation', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
    // clarificateur
    { key: 'nb_clarificateurs', label: 'Nombre de clarificateurs', unit: 'u', group: 'Clarificateur', default: 1 },
    { key: 'indice_Mohlman', label: 'Indice de Mohlman', unit: 'mL/g', group: 'Clarificateur', default: (c) => (c.upstream.primaire ? 250 : 80) },
    { key: 'sortie_MES', label: 'MES eau traitée', unit: 'mg/L', group: 'Clarificateur', default: 50 },
    { key: 'clarif_hauteur', label: 'Hauteur du clarificateur', unit: 'm', group: 'Clarificateur', default: 4 },
    { key: 'clarif_vitesse_max', label: 'Vitesse hydraulique maximale', unit: 'm/h', group: 'Clarificateur', default: (c) => (100 * c.p.clarif_hauteur * Math.sqrt(c.p.sortie_MES / 3.15)) / ((1 + c.p.recirculation_taux) * c.p.indice_Mohlman * c.p.nominal_MES_bassin) },
    { key: 'clarif_surface', label: 'Surface de radier du clarificateur', unit: 'm²', group: 'Clarificateur', default: undefined, hint: 'calculée si non forcée' },
    // boues
    { key: 'boues_concentration', label: 'Concentration des boues extraites', unit: 'g/L', group: 'Boues', default: undefined, hint: 'calculée si non forcée' },
    { key: 'boues_MES', label: 'Boues extraites', unit: 'kg MES/j', group: 'Boues', default: undefined, hint: 'calculé si non forcé' },
    { key: 'extraction_P_refoulement', label: "Pression de refoulement de l'extraction", unit: 'mCE', group: 'Boues', default: 5 },
    { key: 'extraction_pompe_rdt', label: "Rendement global pompes d'extraction", unit: '-', group: 'Boues', default: 0.7 * 0.88 },
  ],
  compute(ctx) {
    const { p, site, forced, choices } = ctx
    const warnings = []
    const a = aer(ctx)
    const Td = site.T_eau_design
    const Tr = site.T_eau_exploit
    const coefRdt = ctx.upstream.primaire ? HYP.BA_rdtDBO_coef_ED : HYP.BA_rdtDBO_coef_EB

    // ---------------- DIMENSIONNEMENT (eau nominale)
    const N = ctx.inNominal
    let G_eq = p.nominal_G * Math.pow(H.correctif_T, Td - H.T_reference)
    let Cm_eq = 1 / (H.a1_Cm * G_eq + H.a0_Cm)
    let Cm = Cm_eq * Math.pow(H.correctif_T, Td - H.T_reference)
    let MVS = N.DBO / Cm
    let volume_bassins = MVS / (p.nominal_MV_MES * p.nominal_MES_bassin)
    if (forced.volume_bassins != null) {
      volume_bassins = forced.volume_bassins
      MVS = volume_bassins * p.nominal_MV_MES * p.nominal_MES_bassin
      Cm = N.DBO / MVS
      Cm_eq = Cm / Math.pow(H.correctif_T, Td - H.T_reference)
    }
    let clarif_surface = (site.Q_nominal * site.pointe_TP + (N.Q - site.Q_nominal)) / (CONST.NOMBRE_HEURE_PAR_JOUR * p.clarif_vitesse_max)
    if (forced.clarif_surface != null) clarif_surface = forced.clarif_surface

    const rdtDBO_nom = polyRdtDBO(coefRdt, Math.min(Cm_eq / p.nominal_MV_MES, H.FtoM_lim))
    const boues_produites_nom = MVS / (p.nominal_MV_MES * p.nominal_G)
    const fT_O2 = Math.pow(H.O2resp.correctionT, Tr - H.O2resp.Tref)
    const ratioO2 = (G) => H.O2resp.a0 + (H.O2resp.a1 * G * fT_O2) / (1 + H.O2resp.a2 * G * fT_O2)
    const O2_besoin_nom = ratioO2(p.reel_G) * rdtDBO_nom * N.DBO + besoinsO2HS(N.Sh)
    const clarif_vmax_recalc = (site.Q_nominal * site.pointe_TP + (N.Q - site.Q_nominal)) / (CONST.NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    const charge_radier = (p.nominal_MES_bassin * N.Q * (1 + p.recirculation_taux)) / (CONST.NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    if (charge_radier > HYP.BA_charge_radier_max[choices.racleur]) warnings.push(`Charge au radier du clarificateur (${charge_radier.toFixed(1)} kg/m²/h) supérieure au maximum admissible (${HYP.BA_charge_radier_max[choices.racleur]}).`)

    const outN = cloneStream(N)
    {
      const boues_concentration = (p.nominal_MES_bassin * (1 + p.recirculation_taux)) / p.recirculation_taux
      const MES_out = (p.sortie_MES * N.Q) / 1000
      const boues_MES = boues_produites_nom - MES_out
      const boues_Q = boues_MES / boues_concentration
      outN.Q = N.Q - boues_Q
      outN.MES = (p.sortie_MES * outN.Q) / 1000
      outN.DCO = N.DCO * (1 - H.rdt_DCO)
      outN.Pt = N.Pt - H.ratio_P_synthese * N.DBO * rdtDBO_nom
      outN.DBO = N.DBO * (1 - rdtDBO_nom)
      outN.Sh = 0
    }

    // ---------------- FONCTIONNEMENT REEL
    const R = ctx.inReel
    const stockage_Q = R.Q
    let G_eq_r = p.reel_G * Math.pow(H.correctif_T, Tr - H.T_reference)
    let Cm_eq_r = 1 / (H.a1_Cm * G_eq_r + H.a0_Cm)
    let Cm_r = Cm_eq_r * Math.pow(H.correctif_T, Tr - H.T_reference)
    let MVS_r = R.DBO / Cm_r
    let reel_MES_bassin = MVS_r / (p.reel_MV_MES * volume_bassins)
    if (forced.reel_MES_bassin != null) {
      reel_MES_bassin = forced.reel_MES_bassin
      MVS_r = volume_bassins * p.reel_MV_MES * reel_MES_bassin
      Cm_r = R.DBO / MVS_r
      Cm_eq_r = Cm_r / Math.pow(H.correctif_T, Tr - H.T_reference)
    }
    const rdtDBO_r = polyRdtDBO(coefRdt, Math.min(Cm_eq_r / p.reel_MV_MES, H.FtoM_lim))
    let sortie_DBO = ((R.DBO * (1 - rdtDBO_r)) / R.Q) * 1000
    if (forced.sortie_DBO != null) sortie_DBO = forced.sortie_DBO
    const boues_produites_r = MVS_r / (p.reel_MV_MES * p.reel_G)
    const DBO_elim = R.DBO - (sortie_DBO * stockage_Q) / 1000
    let O2_besoin = ratioO2(p.reel_G) * DBO_elim
    const besoins_O2_respiration = ((H.O2resp.a1 * p.reel_G * fT_O2) / (1 + H.O2resp.a2 * p.reel_G * fT_O2)) * DBO_elim
    O2_besoin += besoinsO2HS(R.Sh)
    if (forced.O2_besoin != null) O2_besoin = forced.O2_besoin

    let boues_concentration = (reel_MES_bassin * (1 + p.recirculation_taux)) / p.recirculation_taux
    if (forced.boues_concentration != null) boues_concentration = forced.boues_concentration
    const outR = cloneStream(R)
    let boues_MES = boues_produites_r - (p.sortie_MES * R.Q) / 1000
    if (forced.boues_MES != null) boues_MES = forced.boues_MES
    const boues_Q = boues_MES / boues_concentration
    outR.Q = R.Q - boues_Q
    outR.MES = (p.sortie_MES * outR.Q) / 1000
    outR.DCO = R.DCO * (1 - H.rdt_DCO)
    outR.Pt = R.Pt - H.ratio_P_synthese * (R.DBO - (sortie_DBO * outR.Q) / 1000)
    outR.DBO = (sortie_DBO * outR.Q) / 1000
    outR.Sh = 0

    // ---------------- ELECTRICITE
    let alpha
    if (choices.aerateur === 'fines') {
      alpha = H.alpha_fb.a0 + H.alpha_fb.aMVS * (MVS_r / volume_bassins) + H.alpha_fb.aG * (p.reel_G * Math.pow(H.correctif_T, Tr - H.T_reference))
      alpha = Math.min(H.alpha_fb.max, Math.max(H.alpha_fb.min, alpha))
    } else {
      alpha = H.alpha0_autres
      if (reel_MES_bassin > H.alpha_corr.Cref) alpha *= (H.alpha_corr.a0 + H.alpha_corr.a1 * reel_MES_bassin) / (H.alpha_corr.a0 + H.alpha_corr.a1 * H.alpha_corr.Cref)
    }
    if (forced.O2_facteur_alpha != null) alpha = forced.O2_facteur_alpha
    const K = facteurK({ alpha, T_eau: Tr, altitude: site.altitude, hauteur_bassin: p.hauteur_bassin, insufflation: !!a.insufflation, O2_dissous: p.O2_dissous })

    let electricite_aeration, air_Q_Nm3j = 0
    if (a.insufflation) {
      air_Q_Nm3j = O2_besoin / (K.K * (p.O2_rdt_transfert / 100) * (p.hauteur_bassin - HYP.insufflation_hauteur_diffuseur_m) * H.ratio_kgO2_Nm3air)
      if (forced.air_Q_Nm3j != null) air_Q_Nm3j = forced.air_Q_Nm3j
      if (choices.surpresseur === 'roots' && p.air_P_refoulement > H.Pmax_roots) warnings.push(`Pression de refoulement (${p.air_P_refoulement.toFixed(1)} mCE) supérieure au maximum conseillé pour des roots (${H.Pmax_roots} mCE).`)
      electricite_aeration = (air_Q_Nm3j * p.air_P_refoulement * p.surpresseur_conso_spec) / 1000
    } else {
      electricite_aeration = O2_besoin / (p.ASB_eau_claire * K.K)
    }
    const S_unit = clarif_surface / p.nb_clarificateurs
    const electricite_racleur = p.nb_clarificateurs * (S_unit < CONST.PI * Math.pow(H.clarif_diametre_limite / 2, 2) ? 0.55 : 0.75) * CONST.NOMBRE_HEURE_PAR_JOUR
    const ratio_elec_recirc = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * p.recirculation_pompe_rdt)
    const ratio_elec_extr = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * p.extraction_pompe_rdt)
    const electricite_recirculation = ratio_elec_recirc * p.recirculation_taux * stockage_Q * p.recirculation_P_refoulement
    const electricite_extraction = ratio_elec_extr * boues_Q * p.extraction_P_refoulement
    const total = electricite_aeration + electricite_racleur + electricite_recirculation + electricite_extraction
    const fixe = (O2_besoin > 0 ? besoins_O2_respiration / O2_besoin : 0) * electricite_aeration + electricite_racleur

    return {
      outNominal: outN,
      outReel: outR,
      sludge: { origine: 'II_forte', Q: boues_Q, MES: boues_MES, concentration: boues_concentration, MV_MES: p.reel_MV_MES },
      results: [
        { key: 'volume_bassins', label: 'Volume des bassins', unit: 'm³', value: volume_bassins },
        { key: 'Cm', label: 'Charge massique (nominal)', unit: 'kgDBO/(kgMVS·j)', value: Cm },
        { key: 'rdtDBO_nom', label: 'Rendement DBO5 (nominal)', unit: '-', value: rdtDBO_nom },
        { key: 'clarif_surface', label: 'Surface de clarification', unit: 'm²', value: clarif_surface },
        { key: 'clarif_vmax_recalc', label: 'Vitesse hydraulique max recalculée', unit: 'm/h', value: clarif_vmax_recalc },
        { key: 'charge_radier', label: 'Charge au radier (nominal)', unit: 'kg/(m²·h)', value: charge_radier },
        { key: 'reel_MES_bassin', label: 'MES réel dans les bassins', unit: 'g/L', value: reel_MES_bassin },
        { key: 'sortie_DBO', label: 'DBO5 sortie (réel)', unit: 'mg/L', value: sortie_DBO },
        { key: 'O2_besoin', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: O2_besoin },
        { key: 'alpha', label: 'Facteur alpha', unit: '-', value: alpha },
        { key: 'K', label: 'Facteur K', unit: '-', value: K.K },
        { key: 'air_Q', label: "Débit d'air", unit: 'Nm³/h', value: air_Q_Nm3j / 24 },
        { key: 'boues_MES', label: 'Boues extraites (réel)', unit: 'kg MES/j', value: boues_MES },
        { key: 'boues_conc', label: 'Concentration des boues', unit: 'g/L', value: boues_concentration },
        { key: 'MES_out', label: 'MES sortie (réel)', unit: 'mg/L', value: conc(outR, 'MES') },
      ],
      electricity: { total, fixed: fixe, detail: { aeration: electricite_aeration, racleur: electricite_racleur, recirculation: electricite_recirculation, extraction: electricite_extraction } },
      warnings,
    }
  },
})
