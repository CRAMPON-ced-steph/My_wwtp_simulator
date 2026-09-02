// Port de E2_BA_moyenne_charge.cls — même structure que E1 (traitement du
// carbone, pas de nitrification modélisée), constantes différentes :
//   G_reference = 4 j ; 1/Cm = 0,4959×G_eq + 1,3908 (relation E3, ≠ E1) ;
//   rdt_DCO = 0,9 ; alpha0 = 0,8 ; MES/MV_MES/Mohlman/sortie_MES propres.
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, HYP, besoinsO2HS, facteurK } from '../core/hypotheses.js'

const H = {
  T_reference: HYP.BA_bio_equiv_Tref,
  G_reference: 4,
  correctif_T: HYP.BA_bio_equiv_correctionT,
  a1_Cm: HYP.BA_charge_massique_a1, // 0.4959
  a0_Cm: HYP.BA_charge_massique_a0, // 1.3908
  FtoM_lim: HYP.BA_rdtDBO_FtoM_max,
  rdt_DCO: 0.9,
  O2resp: HYP.BA_besoinO2_DBO_resp,
  alpha0_autres: 0.8,
  alpha_fb: { a0: HYP.BA_alpha_finebulle_a0, aMVS: HYP.BA_alpha_finebulle_aMVS, aG: HYP.BA_alpha_finebulle_aGeq, max: 1, min: 0.05 },
  alpha_corr: { a0: HYP.BA_alpha_correction_MES_a0, a1: HYP.BA_alpha_correction_MES_a1, Cref: HYP.BA_alpha_correction_MES_Cref },
  ratio_kgO2_Nm3air: HYP.ratio_kgO2_Nm3air,
  Px2: HYP.surpresseur_Px2_mCE,
  ratio_P_synthese: HYP.assimilation_P_kgP_kgDBO,
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

const mesBassinDefaut = (c) => (c.upstream.primaire ? (c.site.T_eau_design < 12 ? 4 : c.site.T_eau_design < 18 ? 3.5 : 3) : 5)
const mvMesDefaut = (c, T) => (c.upstream.primaire ? (T < 16 ? 0.8 : 0.85) : 0.74)
function O2DissousDefaut(c) {
  const T = c.site.T_eau_exploit
  const t = T < 14 ? { horloge: 4, sans_variateur: 2, variateur: 1.5, avance: 1.5 } : { horloge: 4, sans_variateur: 1.5, variateur: 1, avance: 1 }
  return t[c.choices.regulation]
}
const polyRdtDBO = (coef, x) => coef.reduce((a, k, i) => a + k * Math.pow(x, i), 0)

export default defineNode({
  id: 'ba-moyenne-charge',
  label: 'Boue activée moyenne charge',
  short: 'BA moyenne charge',
  family: 'secondaire',
  vba: 'E2_BA_moyenne_charge.cls',
  description: 'Boue activée moyenne charge (âge de boues ≈ 4 j à 12 °C). Traitement du carbone ; nitrification non modélisée. Clarificateur dimensionné sur la pointe temps de pluie.',
  choices: [
    { key: 'aerateur', label: "Type d'aération", default: 'fines', options: AERATEURS.map(({ value, label }) => ({ value, label })) },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [{ value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' }] },
    { key: 'regulation', label: "Régulation de l'aération", default: 'variateur', options: [{ value: 'horloge', label: 'régulation sur horloge' }, { value: 'sans_variateur', label: 'sans variateur' }, { value: 'variateur', label: 'classique – avec variateur' }, { value: 'avance', label: 'avancé' }] },
    { key: 'racleur', label: 'Type de râcleur', default: 'racle', options: [{ value: 'racle', label: 'râclé' }, { value: 'racle_suce', label: 'râclé-sucé' }, { value: 'kruger', label: 'Kruger' }] },
  ],
  params: [
    { key: 'nominal_G', label: 'Âge de boues design', unit: 'j', group: 'Nominal', default: (c) => H.G_reference * Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_design) },
    { key: 'nominal_MES_bassin', label: 'MES design dans les bassins', unit: 'g/L', group: 'Nominal', default: mesBassinDefaut },
    { key: 'nominal_MV_MES', label: 'MV/MES des boues design', unit: '-', group: 'Nominal', default: (c) => mvMesDefaut(c, c.site.T_eau_design) },
    { key: 'volume_bassins', label: 'Volume des bassins', unit: 'm³', group: 'Nominal', default: undefined, hint: 'calculé si non forcé' },
    { key: 'reel_G', label: 'Âge de boues réel', unit: 'j', group: 'Réel', default: (c) => (c.p.nominal_G / Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_design)) * Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_exploit) },
    { key: 'reel_MV_MES', label: 'MV/MES des boues réel', unit: '-', group: 'Réel', default: (c) => mvMesDefaut(c, c.site.T_eau_exploit) },
    { key: 'reel_MES_bassin', label: 'MES réel dans les bassins', unit: 'g/L', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_dissous', label: 'O2 dissous moyen', unit: 'mg/L', group: 'Réel', default: O2DissousDefaut },
    { key: 'sortie_DBO', label: 'DBO5 en sortie (réel)', unit: 'mg/L', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
    { key: 'hauteur_bassin', label: "Hauteur d'eau du bassin", unit: 'm', group: 'Aération', default: (c) => aer(c).hauteur },
    { key: 'O2_besoin', label: 'Besoin en O2', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_facteur_alpha', label: 'Facteur alpha', unit: '-', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_rdt_transfert', label: 'Rendement de dissolution eau claire', unit: '%/m', group: 'Aération', default: (c) => (c.choices.aerateur === 'moyennes' ? HYP.BA_rdt_dissolution_O2_eau_claire.moyennes_bulles : HYP.BA_rdt_dissolution_O2_eau_claire.fines_bulles) },
    { key: 'air_Q_Nm3j', label: "Débit d'air", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'diffuseur_encrassement', label: 'Durée depuis dernier nettoyage des diffuseurs', unit: 'an', group: 'Aération', default: 0 },
    { key: 'air_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Aération', default: (c) => c.p.hauteur_bassin + 2 + 0.25 * c.p.diffuseur_encrassement },
    { key: 'surpresseur_conso_spec', label: 'Conso spécifique surpresseur', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'ASB_eau_claire', label: 'ASB eau claire (aérateurs de surface)', unit: 'kg O2/kWh', group: 'Aération', default: (c) => aer(c).ASB ?? 0 },
    { key: 'recirculation_taux', label: 'Taux de recirculation', unit: '-', group: 'Recirculation', default: 1 },
    { key: 'recirculation_P_refoulement', label: 'Pression de refoulement recirculation', unit: 'mCE', group: 'Recirculation', default: 5 },
    { key: 'recirculation_pompe_rdt', label: 'Rendement global pompes recirculation', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
    { key: 'nb_clarificateurs', label: 'Nombre de clarificateurs', unit: 'u', group: 'Clarificateur', default: 1 },
    { key: 'indice_Mohlman', label: 'Indice de Mohlman', unit: 'mL/g', group: 'Clarificateur', default: (c) => (c.upstream.primaire ? 185 : 160) },
    { key: 'sortie_MES', label: 'MES eau traitée', unit: 'mg/L', group: 'Clarificateur', default: 25 },
    { key: 'clarif_hauteur', label: 'Hauteur du clarificateur', unit: 'm', group: 'Clarificateur', default: 4 },
    { key: 'clarif_vitesse_max', label: 'Vitesse hydraulique maximale', unit: 'm/h', group: 'Clarificateur', default: (c) => (100 * c.p.clarif_hauteur * Math.sqrt(c.p.sortie_MES / 3.15)) / ((1 + c.p.recirculation_taux) * c.p.indice_Mohlman * c.p.nominal_MES_bassin) },
    { key: 'clarif_surface', label: 'Surface de radier du clarificateur', unit: 'm²', group: 'Clarificateur', default: undefined, hint: 'calculée si non forcée' },
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
    const fT_O2 = Math.pow(H.O2resp.correctionT, Tr - H.O2resp.Tref)
    const ratioO2 = (G) => H.O2resp.a0 + (H.O2resp.a1 * G * fT_O2) / (1 + H.O2resp.a2 * G * fT_O2)

    // ---------------- DIMENSIONNEMENT
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
    let clarif_surface = forced.clarif_surface ?? (site.Q_nominal * site.pointe_TP + (N.Q - site.Q_nominal)) / (CONST.NOMBRE_HEURE_PAR_JOUR * p.clarif_vitesse_max)
    const rdtDBO_nom = polyRdtDBO(coefRdt, Math.min(Cm_eq / p.nominal_MV_MES, H.FtoM_lim))
    const boues_produites_nom = MVS / (p.nominal_MV_MES * p.nominal_G)
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
    let sortie_DBO = forced.sortie_DBO ?? ((R.DBO * (1 - rdtDBO_r)) / R.Q) * 1000
    const boues_produites_r = MVS_r / (p.reel_MV_MES * p.reel_G)
    const DBO_elim = R.DBO - (sortie_DBO * stockage_Q) / 1000
    const besoins_O2_respiration = ((H.O2resp.a1 * p.reel_G * fT_O2) / (1 + H.O2resp.a2 * p.reel_G * fT_O2)) * DBO_elim
    let O2_besoin = ratioO2(p.reel_G) * DBO_elim + besoinsO2HS(R.Sh)
    if (forced.O2_besoin != null) O2_besoin = forced.O2_besoin
    let boues_concentration = forced.boues_concentration ?? (reel_MES_bassin * (1 + p.recirculation_taux)) / p.recirculation_taux
    const outR = cloneStream(R)
    let boues_MES = forced.boues_MES ?? boues_produites_r - (p.sortie_MES * R.Q) / 1000
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
      air_Q_Nm3j = forced.air_Q_Nm3j ?? O2_besoin / (K.K * (p.O2_rdt_transfert / 100) * (p.hauteur_bassin - HYP.insufflation_hauteur_diffuseur_m) * H.ratio_kgO2_Nm3air)
      if (choices.surpresseur === 'roots' && p.air_P_refoulement > H.Px2) warnings.push(`Pression de refoulement (${p.air_P_refoulement.toFixed(1)} mCE) trop élevée pour des surpresseurs roots.`)
      electricite_aeration = (air_Q_Nm3j * p.air_P_refoulement * p.surpresseur_conso_spec) / 1000
    } else {
      electricite_aeration = O2_besoin / (p.ASB_eau_claire * K.K)
    }
    const S_unit = clarif_surface / p.nb_clarificateurs
    const electricite_racleur = p.nb_clarificateurs * (S_unit < CONST.PI * Math.pow(H.clarif_diametre_limite / 2, 2) ? 0.55 : 0.75) * CONST.NOMBRE_HEURE_PAR_JOUR
    const ratioElec = (rdt) => CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * rdt)
    const electricite_recirculation = ratioElec(p.recirculation_pompe_rdt) * p.recirculation_taux * stockage_Q * p.recirculation_P_refoulement
    const electricite_extraction = ratioElec(p.extraction_pompe_rdt) * boues_Q * p.extraction_P_refoulement
    const total = electricite_aeration + electricite_racleur + electricite_recirculation + electricite_extraction
    const fixe = (O2_besoin > 0 ? besoins_O2_respiration / O2_besoin : 0) * electricite_aeration + electricite_racleur

    return {
      outNominal: outN,
      outReel: outR,
      sludge: { origine: 'II_moyenne', Q: boues_Q, MES: boues_MES, concentration: boues_concentration, MV_MES: p.reel_MV_MES },
      results: [
        { key: 'volume_bassins', label: 'Volume des bassins', unit: 'm³', value: volume_bassins },
        { key: 'Cm', label: 'Charge massique (nominal)', unit: 'kgDBO/(kgMVS·j)', value: Cm },
        { key: 'rdtDBO_nom', label: 'Rendement DBO5 (nominal)', unit: '-', value: rdtDBO_nom },
        { key: 'clarif_surface', label: 'Surface de clarification', unit: 'm²', value: clarif_surface },
        { key: 'clarif_vmax', label: 'Vitesse hydraulique max recalculée', unit: 'm/h', value: clarif_vmax_recalc },
        { key: 'charge_radier', label: 'Charge au radier (nominal)', unit: 'kg/(m²·h)', value: charge_radier },
        { key: 'reel_MES_bassin', label: 'MES réel dans les bassins', unit: 'g/L', value: reel_MES_bassin },
        { key: 'sortie_DBO', label: 'DBO5 sortie (réel)', unit: 'mg/L', value: sortie_DBO },
        { key: 'O2_besoin', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: O2_besoin },
        { key: 'alpha', label: 'Facteur alpha', unit: '-', value: alpha },
        { key: 'K', label: 'Facteur K', unit: '-', value: K.K },
        ...(a.insufflation ? [{ key: 'air_Q', label: "Débit d'air", unit: 'Nm³/h', value: air_Q_Nm3j / 24 }] : []),
        { key: 'boues_MES', label: 'Boues extraites (réel)', unit: 'kg MES/j', value: boues_MES },
        { key: 'boues_conc', label: 'Concentration des boues', unit: 'g/L', value: boues_concentration },
        { key: 'MES_out', label: 'MES sortie (réel)', unit: 'mg/L', value: conc(outR, 'MES') },
      ],
      electricity: { total, fixed: fixe, detail: { aeration: electricite_aeration, racleur: electricite_racleur, recirculation: electricite_recirculation, extraction: electricite_extraction } },
      warnings,
    }
  },
})
