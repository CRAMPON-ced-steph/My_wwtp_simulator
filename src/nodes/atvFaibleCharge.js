// ---------------------------------------------------------------------------
// Fabrique commune E3_BA_faible_charge.cls / E4_BA_aeration_prolongee.cls.
// Les deux classes VBA sont identiques à trois constantes près :
//   E3 : G_reference = 14 j, sortie_NH4 = 3 mg/L, sortie_NO3 = 5 mg/L
//   E4 : G_reference = 20 j, sortie_NH4 = 1 mg/L, sortie_NO3 = 6 mg/L
// Le MBR (absent du classeur VBA) est dérivé du même cœur biologique, membranes en
// remplacement du clarificateur (extension signalée, pas un port).
//
// Fidélités volontaires au VBA (voir README) :
// - boues_Q, boues_chimiques et boues_methanol sont des états de classe : le
//   dimensionnement les utilise à 0 (première passe) et le fonctionnement réel
//   réutilise ceux du dimensionnement avant de les recalculer.
// - branche bio-P : P_synthese n'est PAS soustrait (contrairement aux autres branches).
// - dimensionnement : NO3 de sortie n'est pas mis à jour (comportement VBA).
// ---------------------------------------------------------------------------
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, HYP, facteurK, ratioMolairePtbr } from '../core/hypotheses.js'

const AERATEURS = [
  { value: 'fines', label: 'diffuseur fines bulles', hauteur: 6, insufflation: true },
  { value: 'moyennes', label: 'diffuseur moyennes bulles', hauteur: 6, insufflation: true },
  { value: 'brosses', label: "brosses d'aération", hauteur: 3, ASB: HYP.BA_ASB.brosses },
  { value: 'turbines_lentes', label: 'turbines lentes', hauteur: 3, ASB: HYP.BA_ASB.turbines_lentes },
  { value: 'turbines_rapides', label: 'turbines rapides', hauteur: 3, ASB: HYP.BA_ASB.turbines_rapides },
]
const aer = (c) => AERATEURS.find((a) => a.value === c.choices.aerateur)

// hypothèses propres à E3/E4 (Sub hypotheses)
const H = {
  T_reference: HYP.BA_bio_equiv_Tref, // 12
  correctif_T: HYP.BA_bio_equiv_correctionT, // 1.072
  anaerobie_tps_passage_ref: 10, // Q / 10 → temps de séjour anaérobie ~2,4 h
  ratio_DCO_NK_10: 10,
  Pt_residuel: 0.1,
  a1_Cm: HYP.BA_charge_massique_a1, // 0.4959
  a0_Cm: HYP.BA_charge_massique_a0, // 1.3908
  ratio_P_synthese: HYP.assimilation_P_kgP_kgDBO, // 0.01
  a_Pbio_coef_K: 0.0276,
  b_Pbio_coef_K: 0.079,
  a_Kp_bio_1: 0.7,
  b_Kp_bio_1: 0.3,
  lim_Kp_bio_1: 1.15,
  rdt_Psol_max: 0.75,
  ratio_FeCl3_Pt: 5.242, // 162.5/31
  a_dosis_MMS: 3.447, // (55.85+3×17)/31 → boues Fe(OH)3
  b_dosis_MMS: 4.866, // (55.85+31+64)/31 → boues FePO4
  ratio_DCO_sol_EB: 0.33,
  DCO_methanol: 1.5, // kgDCO/kg méthanol
  ratio_boues_methanol: 0.25,
  rdt_DCO: 0.9,
  facteur_boues_C_lim: 1200,
  O2resp: HYP.BA_besoinO2_DBO_resp,
  alpha0_autres: 0.8, // ≠ E1 (0.75)
  alpha_fb: { a0: HYP.BA_alpha_finebulle_a0, aMVS: HYP.BA_alpha_finebulle_aMVS, aG: HYP.BA_alpha_finebulle_aGeq, max: 1, min: 0.05 },
  alpha_corr: { a0: HYP.BA_alpha_correction_MES_a0, a1: HYP.BA_alpha_correction_MES_a1, Cref: HYP.BA_alpha_correction_MES_Cref },
  Px2: HYP.surpresseur_Px2_mCE, // 8
  clarif_diametre_limite: 32,
}

function ratioCNLimite(T) {
  if (T >= 15) return 4
  if (T <= 5) return 5
  return 4 + ((5 - 4) * (T - 5)) / (15 - 5)
}
function O2DissousDefaut(c) {
  const T = c.site.T_eau_exploit
  const table = T < 14 ? { horloge: 4, sans_variateur: 2.5, variateur: 2, avance: 1.5 } : { horloge: 4, sans_variateur: 2, variateur: 1.5, avance: 0.8 }
  return table[c.choices.regulation]
}
/** rendement de dissolution eau claire (%/m) — E3/E4 : relation sur la hauteur
 *  si la hauteur de bassin est FORCÉE sous Hlim (6 m) : rdt = −0,18×h + 6,88 ;
 *  sinon 5,8 %/m (fines bulles) / 2,5 %/m (moyennes bulles). */
function rdtDissolution(c) {
  if (c.choices.aerateur === 'moyennes') return HYP.BA_rdt_dissolution_O2_eau_claire.moyennes_bulles
  const hf = c.forced?.hauteur_bassin
  if (hf != null && hf < HYP.BA_rdt_dissolution_Hlim) return HYP.BA_rdt_dissolution_pente * hf + HYP.BA_rdt_dissolution_ordonnee
  return HYP.BA_rdt_dissolution_O2_eau_claire.fines_bulles_Hsup
}

const dephos = (c) => c.choices.dephosphatation
const bioP = (c) => dephos(c) === 'bio' || dephos(c) === 'co_precipitation'
const MV_MES_ref = (c) => (c.upstream.primaire ? 0.77 : 0.67)
const polyRdtDBO = (coef, x) => coef.reduce((a, k, i) => a + k * Math.pow(x, i), 0)

/**
 * cfg : { id, label, short, vba, description, G_reference, sortie_NH4_def,
 *         sortie_NO3_def, origineEB, origineED, membrane? }
 * membrane (MBR uniquement) : remplace le clarificateur par des membranes.
 */
export function makeATVFaibleCharge(cfg) {
  const membrane = !!cfg.membrane
  return defineNode({
    id: cfg.id,
    label: cfg.label,
    short: cfg.short,
    family: 'secondaire',
    vba: cfg.vba,
    ported: true,
    extension: cfg.extension,
    description: cfg.description,
    choices: [
      { key: 'dephosphatation', label: 'Déphosphatation', default: 'co_precipitation', options: [
        { value: 'non', label: 'aucune' },
        { value: 'bio', label: 'biologique' },
        { value: 'precipitation', label: 'précipitation (FeCl3)' },
        { value: 'co_precipitation', label: 'bio + co-précipitation' },
      ] },
      ...(membrane ? [] : [{ key: 'config_bassin', label: 'Configuration des bassins', default: 'plugflow', options: [
        { value: 'plugflow', label: 'cuves séparées (plug-flow)' },
        { value: 'chenal', label: 'chenal (carrousel)' },
      ] }]),
      { key: 'aerateur', label: "Type d'aération", default: 'fines', options: (membrane ? AERATEURS.filter((a) => a.insufflation) : AERATEURS).map(({ value, label }) => ({ value, label })) },
      { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
        { value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' },
      ] },
      { key: 'regulation', label: "Régulation de l'aération", default: 'variateur', options: [
        { value: 'horloge', label: 'sur horloge' }, { value: 'sans_variateur', label: 'sans variateur' }, { value: 'variateur', label: 'classique – avec variateur' }, { value: 'avance', label: 'avancé' },
      ] },
      ...(membrane ? [] : [{ key: 'racleur', label: 'Type de râcleur', default: 'racle', options: [
        { value: 'racle', label: 'râclé' }, { value: 'racle_suce', label: 'râclé-sucé' }, { value: 'kruger', label: 'Kruger' },
      ] }]),
      { key: 'fct_SHUNT', label: 'Fonctionnement SHUNT (nitritation)', default: 'non', options: [{ value: 'non', label: 'non' }, { value: 'oui', label: 'oui' }] },
    ],
    params: [
      // nominal
      { key: 'nominal_G', label: 'Âge de boues design', unit: 'j', group: 'Nominal', default: (c) => (cfg.G_reference * (bioP(c) ? 1.08 : 1)) * Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_design) },
      { key: 'nominal_MES_bassin', label: 'MES design dans les bassins', unit: 'g/L', group: 'Nominal', default: (c) => (membrane ? cfg.membrane.MES_bassin_def : c.upstream.primaire ? 3.5 : 5) },
      { key: 'nominal_MV_MES', label: 'MV/MES des boues design', unit: '-', group: 'Nominal', default: undefined, hint: 'calculé (MV/MES référence + boues chimiques et méthanol)' },
      { key: 'volume_bassins', label: 'Volume total des bassins', unit: 'm³', group: 'Nominal', default: undefined, hint: 'calculé si non forcé' },
      { key: 'volume_anaerobie', label: 'Volume anaérobie (bio-P)', unit: 'm³', group: 'Zones', default: undefined, hint: 'calculé si non forcé' },
      { key: 'volume_anoxie', label: 'Volume anoxie', unit: 'm³', group: 'Zones', default: undefined, hint: 'calculé si non forcé' },
      { key: 'volume_aerobie', label: 'Volume aérobie', unit: 'm³', group: 'Zones', default: undefined, hint: 'calculé si non forcé' },
      ...(membrane ? [] : [{ key: 'volume_chenal', label: 'Volume chenal', unit: 'm³', group: 'Zones', default: undefined, hint: 'calculé si non forcé' }]),
      { key: 'ratio_elec_anaerobie', label: 'Agitation zone anaérobie', unit: 'W/m³', group: 'Zones', default: 10 },
      { key: 'ratio_elec_anoxie', label: 'Agitation zone anoxie', unit: 'W/m³', group: 'Zones', default: 3 },
      { key: 'ratio_elec_aerobie', label: 'Agitation zone aérobie', unit: 'W/m³', group: 'Zones', default: 0 },
      ...(membrane ? [] : [{ key: 'ratio_elec_chenal', label: 'Agitation chenal', unit: 'W/m³', group: 'Zones', default: 3 }]),
      // réel
      { key: 'reel_G', label: 'Âge de boues réel', unit: 'j', group: 'Réel', default: (c) => (c.p.nominal_G / Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_design)) * Math.pow(H.correctif_T, H.T_reference - c.site.T_eau_exploit) },
      { key: 'reel_MV_MES', label: 'MV/MES des boues réel', unit: '-', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
      { key: 'reel_MES_bassin', label: 'MES réel dans les bassins', unit: 'g/L', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
      { key: 'O2_dissous', label: 'O2 dissous moyen', unit: 'mg/L', group: 'Réel', default: (c) => (c.choices.fct_SHUNT === 'oui' ? 0.5 : O2DissousDefaut(c)) },
      { key: 'sortie_DBO', label: 'DBO5 en sortie (réel)', unit: 'mg/L', group: 'Sorties', default: undefined, hint: 'calculé si non forcé' },
      { key: 'sortie_NH4', label: 'N-NH4 en sortie', unit: 'mg/L', group: 'Sorties', default: cfg.sortie_NH4_def },
      { key: 'sortie_NO3', label: 'N-NO3 en sortie', unit: 'mg/L', group: 'Sorties', default: cfg.sortie_NO3_def },
      { key: 'rendement_bioP', label: 'Rendement de déphosphatation biologique', unit: '-', group: 'Sorties', default: undefined, hint: 'calculé si non forcé' },
      { key: 'sortie_Pt', label: 'Pt en sortie (cible précipitation)', unit: 'mg/L', group: 'Sorties', default: (c) => c.site.Pt_garantie },
      { key: 'methanol_pur', label: 'Méthanol pur (réel)', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculé si non forcé' },
      { key: 'FeCl3_pur', label: 'FeCl3 pur (réel)', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculé si non forcé' },
      // aération
      { key: 'hauteur_bassin', label: "Hauteur d'eau des bassins", unit: 'm', group: 'Aération', default: (c) => aer(c).hauteur },
      { key: 'O2_besoin', label: 'Besoin en O2', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
      { key: 'O2_facteur_alpha', label: 'Facteur alpha', unit: '-', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
      { key: 'O2_rdt_transfert', label: 'Rendement de dissolution eau claire', unit: '%/m', group: 'Aération', default: (c) => rdtDissolution(c) },
      { key: 'air_Q_Nm3j', label: "Débit d'air", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
      { key: 'diffuseur_encrassement', label: 'Durée depuis dernier nettoyage des diffuseurs', unit: 'an', group: 'Aération', default: 0 },
      { key: 'air_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Aération', default: (c) => c.p.hauteur_bassin + 2 + 0.25 * c.p.diffuseur_encrassement },
      { key: 'surpresseur_conso_spec', label: 'Conso spécifique surpresseur', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
      { key: 'ASB_eau_claire', label: 'ASB eau claire (aérateurs de surface)', unit: 'kg O2/kWh', group: 'Aération', default: (c) => aer(c).ASB ?? 0 },
      // recirculations
      { key: 'recirculation_MLSS_taux', label: 'Taux de recirculation de liqueur mixte', unit: '-', group: 'Recirculation', default: undefined, hint: 'calculé sur le bilan NO3 si non forcé' },
      { key: 'recirculation_MLSS_P_refoulement', label: 'Pression recirculation liqueur mixte', unit: 'mCE', group: 'Recirculation', default: 2 },
      { key: 'recirculation_MLSS_pompe_rdt', label: 'Rendement pompes liqueur mixte', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
      ...(membrane ? [] : [
        { key: 'recirculation_taux', label: 'Taux de recirculation des boues (clarif)', unit: '-', group: 'Recirculation', default: 1 },
        { key: 'recirculation_P_refoulement', label: 'Pression recirculation boues', unit: 'mCE', group: 'Recirculation', default: 5 },
        { key: 'recirculation_pompe_rdt', label: 'Rendement pompes recirculation', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
      ]),
      // clarificateur ou membranes
      ...(membrane
        ? [
            { key: 'flux_membranes', label: 'Flux de dimensionnement (pointe temps de pluie)', unit: 'L/(m²·h)', group: 'Membranes', default: cfg.membrane.flux_def, hint: 'flux instantané admissible en pointe' },
            { key: 'surface_membranes', label: 'Surface membranaire', unit: 'm²', group: 'Membranes', default: undefined, hint: 'calculée sur la pointe temps de pluie si non forcée' },
            { key: 'SADm', label: 'Air de décolmatage spécifique (SADm)', unit: 'Nm³/(m²·h)', group: 'Membranes', default: cfg.membrane.SADm_def },
            { key: 'membrane_cyclage', label: "Cyclage de l'air de décolmatage (part du temps)", unit: '-', group: 'Membranes', default: 0.5, hint: 'aération séquencée 10 s / 10 s' },
            { key: 'membrane_P_refoulement', label: "Pression de l'air de décolmatage", unit: 'mCE', group: 'Membranes', default: 6 },
            { key: 'permeat_P_aspiration', label: 'Dépression de perméation', unit: 'mCE', group: 'Membranes', default: 3 },
            { key: 'permeat_pompe_rdt', label: 'Rendement pompes de perméat', unit: '-', group: 'Membranes', default: 0.5 },
            { key: 'sortie_MES', label: 'MES eau traitée', unit: 'mg/L', group: 'Membranes', default: 1 },
          ]
        : [
            { key: 'nb_clarificateurs', label: 'Nombre de clarificateurs', unit: 'u', group: 'Clarificateur', default: 1 },
            { key: 'indice_Mohlman', label: 'Indice de Mohlman', unit: 'mL/g', group: 'Clarificateur', default: (c) => (c.upstream.primaire ? (dephos(c) === 'non' ? 165 : 185) : dephos(c) === 'non' ? 125 : 135) },
            { key: 'sortie_MES', label: 'MES eau traitée', unit: 'mg/L', group: 'Clarificateur', default: 20 },
            { key: 'clarif_hauteur', label: 'Hauteur du clarificateur', unit: 'm', group: 'Clarificateur', default: 4 },
            { key: 'clarif_vitesse_max', label: 'Vitesse hydraulique maximale', unit: 'm/h', group: 'Clarificateur', default: (c) => (100 * c.p.clarif_hauteur * Math.sqrt(c.p.sortie_MES / 3.15)) / ((1 + c.p.recirculation_taux) * c.p.indice_Mohlman * c.p.nominal_MES_bassin) },
            { key: 'clarif_surface', label: 'Surface de radier du clarificateur', unit: 'm²', group: 'Clarificateur', default: undefined, hint: 'calculée si non forcée' },
          ]),
      // boues
      { key: 'boues_concentration', label: 'Concentration des boues extraites', unit: 'g/L', group: 'Boues', default: undefined, hint: 'calculée si non forcée' },
      { key: 'boues_MES', label: 'Boues extraites', unit: 'kg MES/j', group: 'Boues', default: undefined, hint: 'calculé si non forcé' },
      { key: 'extraction_P_refoulement', label: "Pression de refoulement de l'extraction", unit: 'mCE', group: 'Boues', default: 5 },
      { key: 'extraction_pompe_rdt', label: "Rendement pompes d'extraction", unit: '-', group: 'Boues', default: 0.7 * 0.88 },
    ],
    compute(ctx) {
      const { p, site, forced, choices } = ctx
      const warnings = []
      const a = aer(ctx)
      const Td = site.T_eau_design
      const Tr = site.T_eau_exploit
      const SHUNT = choices.fct_SHUNT === 'oui'
      const configChenal = !membrane && choices.config_bassin === 'chenal'
      const coefRdt = ctx.upstream.primaire ? HYP.BA_rdtDBO_coef_ED : HYP.BA_rdtDBO_coef_EB
      const ratio_O2_nit = SHUNT ? 4.57 * 0.75 : HYP.ratio_O2_nit
      const ratio_O2_denit = SHUNT ? -2.86 * 0.6 : HYP.ratio_O2_denit
      const shuntCN = SHUNT ? 0.6 : 1
      const MVref = MV_MES_ref(ctx)
      const rdt_DCO_sol_primaire = ctx.upstream.primaire_reactif ? 0.08 : 0
      const recirculation_taux0 = membrane ? 0 : p.recirculation_taux
      const fT_O2 = Math.pow(H.O2resp.correctionT, Tr - H.O2resp.Tref)
      const ratioResp = (G) => H.O2resp.a0 + (H.O2resp.a1 * G * fT_O2) / (1 + H.O2resp.a2 * G * fT_O2)

      // état de classe partagé entre les deux passes (comportement VBA)
      const st = { boues_chimiques: 0, boues_methanol: 0, boues_Q: 0, methanol_pur: 0 }

      // ---- rendement bio-P (identique aux deux passes, T réelle utilisée dans les deux — comportement VBA)
      const rendementBioP = (Q, DBO, Pt, G_equivalent, volume_anaerobie) => {
        if (forced.rendement_bioP != null) return forced.rendement_bioP
        const K = H.a_Pbio_coef_K + H.b_Pbio_coef_K / G_equivalent
        const ratio_DBO_P = DBO / Pt
        const volume_optimal = (Q / H.anaerobie_tps_passage_ref) * Math.pow(H.correctif_T, H.T_reference - Tr)
        let Kp1 = volume_anaerobie / volume_optimal
        if (Kp1 <= 1) Kp1 = Kp1 * Kp1
        else if (Kp1 <= 1.5) Kp1 = H.a_Kp_bio_1 + H.b_Kp_bio_1 * Kp1
        else Kp1 = H.lim_Kp_bio_1
        return Math.min(H.rdt_Psol_max, K * Kp1 * 1 * ratio_DBO_P)
      }
      const ratio_Psol_P = ctx.upstream.primaire ? 0.85 : 0.7

      // ---- traitement du phosphore (bloc commun aux deux passes)
      const traitementP = (S, out, Q, DBO, rdt_DBO, G_equivalent, volume_anaerobie, reel) => {
        const P_synthese = H.ratio_P_synthese * DBO * rdt_DBO
        let rendement_bioP = 0, P_precipite = 0, FeCl3_pur = 0
        let sortie_Pt = p.sortie_Pt
        const precipitation = () => {
          let P_precipitation = out.Pt - (sortie_Pt * (Q - st.boues_Q)) / 1000
          let rm = 0
          if (P_precipitation > 0) {
            const Ptbr = (P_precipitation / Q) * 1000
            rm = ratioMolairePtbr(Ptbr, sortie_Pt)
            FeCl3_pur = rm * H.ratio_FeCl3_Pt * P_precipitation
          } else {
            P_precipitation = 0
            FeCl3_pur = 0
          }
          if (reel && forced.FeCl3_pur != null) FeCl3_pur = forced.FeCl3_pur
          const dose_MMS = (rm - 1) * H.a_dosis_MMS + H.b_dosis_MMS
          st.boues_chimiques = dose_MMS * P_precipitation
          out.Pt -= P_precipitation
          P_precipite = P_precipitation
        }
        if (dephos(ctx) === 'bio') {
          rendement_bioP = rendementBioP(Q, DBO, out.Pt, G_equivalent, volume_anaerobie)
          out.Pt -= rendement_bioP * ratio_Psol_P * out.Pt
          sortie_Pt = (out.Pt / (Q - st.boues_Q)) * 1000
          st.boues_chimiques = 0
        } else if (dephos(ctx) === 'co_precipitation') {
          rendement_bioP = rendementBioP(Q, DBO, out.Pt, G_equivalent, volume_anaerobie)
          out.Pt -= rendement_bioP * ratio_Psol_P * out.Pt
          out.Pt -= P_synthese
          precipitation()
        } else if (dephos(ctx) === 'precipitation') {
          out.Pt -= P_synthese
          precipitation()
        } else {
          out.Pt -= P_synthese
          st.boues_chimiques = 0
        }
        // Garde-fou (écart assumé) : le VBA déclare Pt_residuel = 0,1 mg/L sans
        // l'utiliser et peut produire un Pt négatif (bio-P + assimilation > Pt
        // entrant). On plafonne l'abattement au résiduel et on signale.
        const Pt_plancher = (H.Pt_residuel * (Q - st.boues_Q)) / 1000
        if (out.Pt < Pt_plancher) {
          if (reel) warnings.push(`Pt calculé sous le résiduel de ${H.Pt_residuel} mg/L (bio-P + assimilation supérieurs au Pt entrant) : sortie plafonnée au résiduel.`)
          out.Pt = Pt_plancher
        }
        return { rendement_bioP, P_precipite, FeCl3_pur, sortie_Pt, P_synthese }
      }

      // =================================================================
      // DIMENSIONNEMENT (eau nominale)
      // =================================================================
      const N = ctx.inNominal
      const outN = cloneStream(N)
      const G_eq = p.nominal_G * Math.pow(H.correctif_T, Td - H.T_reference)
      let Cm_eq = 1 / (H.a1_Cm * G_eq + H.a0_Cm)
      let Cm = Cm_eq * Math.pow(H.correctif_T, Td - H.T_reference)
      let MVS = N.DBO / Cm
      let boues_bassins_MES = MVS / MVref + (st.boues_chimiques + st.boues_methanol) * p.nominal_G
      let nominal_MV_MES = (MVS + st.boues_methanol * p.nominal_G) / boues_bassins_MES
      if (forced.nominal_MV_MES != null) nominal_MV_MES = forced.nominal_MV_MES
      let volume_bassins = MVS / (nominal_MV_MES * p.nominal_MES_bassin)
      if (forced.volume_bassins != null) {
        volume_bassins = forced.volume_bassins
        MVS = volume_bassins * nominal_MV_MES * p.nominal_MES_bassin
        Cm = N.DBO / MVS
        Cm_eq = Cm / Math.pow(H.correctif_T, Td - H.T_reference)
      }
      // zones
      let volume_anaerobie = 0
      if (bioP(ctx)) {
        volume_anaerobie = forced.volume_anaerobie ?? (N.Q / H.anaerobie_tps_passage_ref) * Math.pow(H.correctif_T, H.T_reference - Td)
      }
      let volume_anoxie = 0, volume_aerobie = 0, volume_chenal = 0
      if (!configChenal) {
        volume_anoxie = forced.volume_anoxie ?? (N.DCO / N.NK < H.ratio_DCO_NK_10 ? 0.4 : 0.2) * (volume_bassins - volume_anaerobie)
        volume_aerobie = forced.volume_aerobie ?? volume_bassins - (volume_anoxie + volume_anaerobie)
      } else {
        volume_chenal = forced.volume_chenal ?? volume_bassins - volume_anaerobie
      }
      if (Math.abs(volume_bassins - (volume_anaerobie + volume_anoxie + volume_aerobie + volume_chenal)) > 1e-6) {
        warnings.push('La somme des volumes de zones ne correspond pas au volume total des bassins.')
      }
      // clarificateur / membranes (surface sur la pointe temps de pluie)
      const Qpointe_h = (site.Q_nominal * site.pointe_TP + (N.Q - site.Q_nominal)) / CONST.NOMBRE_HEURE_PAR_JOUR
      let clarif_surface = 0, surface_membranes = 0
      if (membrane) {
        surface_membranes = forced.surface_membranes ?? (Qpointe_h * 1000) / p.flux_membranes
      } else {
        clarif_surface = forced.clarif_surface ?? Qpointe_h / p.clarif_vitesse_max
      }
      // fonctionnement nominal
      const rdtDBO_nom = polyRdtDBO(coefRdt, Cm_eq / nominal_MV_MES)
      const nitrification_nom = N.NH4 - (p.sortie_NH4 * N.Q) / 1000
      const denitrification_nom = N.NO3 + nitrification_nom - (p.sortie_NO3 * N.Q) / 1000
      // méthanol
      const DCO_soluble_nom = ((site.DCO_nominal * site.Q_nominal) / 1000) * H.ratio_DCO_sol_EB * (1 - rdt_DCO_sol_primaire)
      let methanol_nom = 0
      if (denitrification_nom > 0 && DCO_soluble_nom / denitrification_nom < ratioCNLimite(Td) * shuntCN) {
        methanol_nom = (ratioCNLimite(Td) * shuntCN * denitrification_nom - DCO_soluble_nom) / H.DCO_methanol
      }
      st.methanol_pur = methanol_nom
      let DCO_necessaire = methanol_nom * H.DCO_methanol
      st.boues_methanol = H.ratio_boues_methanol * DCO_necessaire
      outN.DCO = N.DCO + DCO_necessaire * (1 - H.ratio_boues_methanol)
      // phosphore
      const Pnom = traitementP(N, outN, N.Q, N.DBO, rdtDBO_nom, G_eq, volume_anaerobie, false)
      // O2
      let O2_besoin_nom = ratioResp(p.reel_G) * rdtDBO_nom * N.DBO
      O2_besoin_nom += ratio_O2_nit * nitrification_nom + ratio_O2_denit * denitrification_nom
      O2_besoin_nom += methanol_nom * H.DCO_methanol
      // boues + radier
      const boues_produites_nom = MVS / (nominal_MV_MES * p.nominal_G)
      const clarif_vmax_recalc = membrane ? 0 : Qpointe_h / clarif_surface
      const charge_radier = membrane ? 0 : (p.nominal_MES_bassin * N.Q * (1 + recirculation_taux0)) / (CONST.NOMBRE_HEURE_PAR_JOUR * clarif_surface)
      if (!membrane && charge_radier > HYP.BA_charge_radier_max[choices.racleur]) {
        warnings.push(`Charge au radier du clarificateur (${charge_radier.toFixed(1)} kg/m²/h) supérieure au maximum admissible (${HYP.BA_charge_radier_max[choices.racleur]}).`)
      }
      // sorties nominales
      outN.Q = N.Q - st.boues_Q
      outN.MES = (p.sortie_MES * outN.Q) / 1000
      outN.DCO *= 1 - H.rdt_DCO
      outN.DBO = N.DBO * (1 - rdtDBO_nom)
      outN.NH4 = N.NH4 - nitrification_nom
      outN.NK = outN.NH4
      outN.Sh = 0
      const boues_MES_nom = boues_produites_nom - outN.MES
      const boues_conc_nom = membrane ? p.nominal_MES_bassin : (p.nominal_MES_bassin * (1 + recirculation_taux0)) / recirculation_taux0
      st.boues_Q = boues_MES_nom / boues_conc_nom

      // =================================================================
      // FONCTIONNEMENT REEL
      // =================================================================
      const R = ctx.inReel
      const outR = cloneStream(R)
      const stockage_Q = R.Q
      const G_eq_r = p.reel_G * Math.pow(H.correctif_T, Tr - H.T_reference)
      let Cm_eq_r = 1 / (H.a1_Cm * G_eq_r + H.a0_Cm)
      let Cm_r = Cm_eq_r * Math.pow(H.correctif_T, Tr - H.T_reference)
      let MVS_r = R.DBO / Cm_r
      let boues_bassins_MES_r = MVS_r / MVref + (st.boues_chimiques + st.boues_methanol) * p.reel_G
      let reel_MV_MES = (MVS_r + p.reel_G * st.boues_methanol) / boues_bassins_MES_r
      if (forced.reel_MV_MES != null) reel_MV_MES = forced.reel_MV_MES
      let reel_MES_bassin = MVS_r / (reel_MV_MES * volume_bassins)
      if (forced.reel_MES_bassin != null) {
        reel_MES_bassin = forced.reel_MES_bassin
        MVS_r = volume_bassins * reel_MV_MES * reel_MES_bassin
        Cm_r = R.DBO / MVS_r
        Cm_eq_r = Cm_r / Math.pow(H.correctif_T, Tr - H.T_reference)
      }
      const rdtDBO_r = polyRdtDBO(coefRdt, Cm_eq_r / reel_MV_MES)
      let sortie_DBO = forced.sortie_DBO ?? ((R.DBO * (1 - rdtDBO_r)) / R.Q) * 1000
      const nitrification = R.NH4 - (p.sortie_NH4 * R.Q) / 1000
      const denitrification = R.NO3 + nitrification - (p.sortie_NO3 * R.Q) / 1000
      // recirculation de liqueur mixte
      let recirculation_MLSS_taux = 0
      if (!configChenal) {
        recirculation_MLSS_taux = (nitrification * 1000) / (R.Q * p.sortie_NO3) - (1 + recirculation_taux0)
        if (forced.recirculation_MLSS_taux != null) {
          if (recirculation_MLSS_taux > forced.recirculation_MLSS_taux) warnings.push('Le taux de recirculation de liqueur mixte forcé est inférieur au taux nécessaire pour atteindre le NO3 de sortie visé.')
          recirculation_MLSS_taux = forced.recirculation_MLSS_taux
        }
        if (recirculation_MLSS_taux < 0) recirculation_MLSS_taux = 0
      }
      // méthanol
      const DCO_soluble = ((site.DCO_nominal * site.Q_nominal) / 1000) * site.NC_DCO * H.ratio_DCO_sol_EB * (1 - rdt_DCO_sol_primaire)
      let methanol_pur = 0
      if (denitrification > 0 && DCO_soluble / denitrification < ratioCNLimite(Tr) * shuntCN) {
        methanol_pur = (ratioCNLimite(Tr) * shuntCN * denitrification - DCO_soluble) / H.DCO_methanol
      }
      if (forced.methanol_pur != null) methanol_pur = forced.methanol_pur
      DCO_necessaire = methanol_pur * H.DCO_methanol
      const DCO_apportee_methanol = DCO_necessaire
      st.boues_methanol = H.ratio_boues_methanol * DCO_necessaire
      outR.DCO = R.DCO + DCO_necessaire * (1 - H.ratio_boues_methanol)
      // phosphore
      const Preel = traitementP(R, outR, R.Q, R.DBO, rdtDBO_r, G_eq_r, volume_anaerobie, true)
      // O2
      const besoins_O2_respiration = ((H.O2resp.a1 * p.reel_G * fT_O2) / (1 + H.O2resp.a2 * p.reel_G * fT_O2)) * (R.DBO - (sortie_DBO * stockage_Q) / 1000)
      let O2_besoin = ratioResp(p.reel_G) * (R.DBO - (sortie_DBO * stockage_Q) / 1000)
      O2_besoin += ratio_O2_nit * nitrification + ratio_O2_denit * denitrification
      O2_besoin += methanol_pur * H.DCO_methanol
      if (forced.O2_besoin != null) O2_besoin = forced.O2_besoin
      // boues
      const boues_produites_r = (reel_MES_bassin * volume_bassins) / p.reel_G
      // sorties réelles
      outR.Q = R.Q - st.boues_Q
      outR.MES = (p.sortie_MES * outR.Q) / 1000
      outR.DCO *= 1 - H.rdt_DCO
      outR.DBO = (sortie_DBO * outR.Q) / 1000
      outR.NH4 = (p.sortie_NH4 * outR.Q) / 1000
      outR.NK = outR.NH4
      outR.NO3 = (p.sortie_NO3 * outR.Q) / 1000
      outR.Sh = 0
      let boues_MES = forced.boues_MES ?? boues_produites_r - outR.MES
      let boues_concentration = forced.boues_concentration ?? (membrane ? reel_MES_bassin : (reel_MES_bassin * (1 + recirculation_taux0)) / recirculation_taux0)
      const boues_Q = boues_MES / boues_concentration
      st.boues_Q = boues_Q
      if (!membrane && boues_concentration > H.facteur_boues_C_lim / p.indice_Mohlman) {
        warnings.push('Concentration des boues recirculées trop élevée pour une bonne décantation : augmenter le taux de recirculation.')
      }

      // =================================================================
      // ELECTRICITE
      // =================================================================
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
        air_Q_Nm3j = forced.air_Q_Nm3j ?? O2_besoin / (K.K * (p.O2_rdt_transfert / 100) * (p.hauteur_bassin - HYP.insufflation_hauteur_diffuseur_m) * HYP.ratio_kgO2_Nm3air)
        if (choices.surpresseur === 'roots' && p.air_P_refoulement > H.Px2) warnings.push(`Pression de refoulement (${p.air_P_refoulement.toFixed(1)} mCE) trop élevée pour des surpresseurs roots.`)
        electricite_aeration = (air_Q_Nm3j * p.air_P_refoulement * p.surpresseur_conso_spec) / 1000
      } else {
        electricite_aeration = O2_besoin / (p.ASB_eau_claire * K.K)
      }
      let electricite_racleur = 0
      if (!membrane) {
        const S_unit = clarif_surface / p.nb_clarificateurs
        electricite_racleur = p.nb_clarificateurs * (S_unit < CONST.PI * Math.pow(H.clarif_diametre_limite / 2, 2) ? 0.55 : 0.75) * CONST.NOMBRE_HEURE_PAR_JOUR
      }
      // agitation
      const agit = HYP.BA_agitation_tps_fct
      let electricite_agitation = 0
      if (bioP(ctx)) electricite_agitation += (p.ratio_elec_anaerobie * volume_anaerobie * agit) / 1000
      if (!configChenal) {
        electricite_agitation += (p.ratio_elec_anoxie * volume_anoxie * agit) / 1000
        electricite_agitation += (p.ratio_elec_aerobie * volume_aerobie * agit) / 1000
      } else {
        electricite_agitation += (p.ratio_elec_chenal * volume_chenal * agit) / 1000
      }
      // pompages
      const ratioElec = (rdt) => CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * rdt)
      const electricite_recirculation_MLSS = ratioElec(p.recirculation_MLSS_pompe_rdt) * recirculation_MLSS_taux * stockage_Q * p.recirculation_MLSS_P_refoulement
      const electricite_recirculation = membrane ? 0 : ratioElec(p.recirculation_pompe_rdt) * recirculation_taux0 * stockage_Q * p.recirculation_P_refoulement
      const electricite_extraction = ratioElec(p.extraction_pompe_rdt) * boues_Q * p.extraction_P_refoulement
      // membranes (extension MBR, hors classeur VBA)
      let electricite_membranes = 0, electricite_permeat = 0
      if (membrane) {
        electricite_membranes = (p.SADm * p.membrane_cyclage * surface_membranes * CONST.NOMBRE_HEURE_PAR_JOUR * p.membrane_P_refoulement * p.surpresseur_conso_spec) / 1000
        electricite_permeat = ratioElec(p.permeat_pompe_rdt) * outR.Q * p.permeat_P_aspiration
      }
      const total = electricite_aeration + electricite_racleur + electricite_agitation + electricite_recirculation_MLSS + electricite_recirculation + electricite_extraction + electricite_membranes + electricite_permeat
      const fixe = (O2_besoin > 0 ? besoins_O2_respiration / O2_besoin : 0) * electricite_aeration + electricite_racleur + electricite_agitation

      return {
        outNominal: outN,
        outReel: outR,
        sludge: { origine: ctx.upstream.primaire ? cfg.origineED : cfg.origineEB, Q: boues_Q, MES: boues_MES, concentration: boues_concentration, MV_MES: reel_MV_MES },
        reactifs: { methanol_kgj: methanol_pur, FeCl3_kgj: Preel.FeCl3_pur },
        results: [
          { key: 'volume_bassins', label: 'Volume total des bassins', unit: 'm³', value: volume_bassins },
          ...(volume_anaerobie ? [{ key: 'v_ana', label: 'Volume anaérobie', unit: 'm³', value: volume_anaerobie }] : []),
          ...(!configChenal ? [
            { key: 'v_anox', label: 'Volume anoxie', unit: 'm³', value: volume_anoxie },
            { key: 'v_aero', label: 'Volume aérobie', unit: 'm³', value: volume_aerobie },
          ] : [{ key: 'v_chenal', label: 'Volume chenal', unit: 'm³', value: volume_chenal }]),
          { key: 'Cm', label: 'Charge massique (nominal)', unit: 'kgDBO/(kgMVS·j)', value: Cm },
          { key: 'MV_MES', label: 'MV/MES calculé (réel)', unit: '-', value: reel_MV_MES },
          { key: 'reel_MES_bassin', label: 'MES réel dans les bassins', unit: 'g/L', value: reel_MES_bassin },
          ...(membrane
            ? [{ key: 'S_membranes', label: 'Surface membranaire', unit: 'm²', value: surface_membranes }]
            : [
                { key: 'clarif_surface', label: 'Surface de clarification', unit: 'm²', value: clarif_surface },
                { key: 'clarif_vmax', label: 'Vitesse hydraulique max recalculée', unit: 'm/h', value: clarif_vmax_recalc },
                { key: 'charge_radier', label: 'Charge au radier (nominal)', unit: 'kg/(m²·h)', value: charge_radier },
              ]),
          { key: 'nit', label: 'Nitrification (réel)', unit: 'kg N/j', value: nitrification },
          { key: 'denit', label: 'Dénitrification (réel)', unit: 'kg N/j', value: denitrification },
          { key: 'MLSS', label: 'Taux de recirculation liqueur mixte', unit: '-', value: recirculation_MLSS_taux },
          { key: 'methanol', label: 'Méthanol pur (réel)', unit: 'kg/j', value: methanol_pur },
          { key: 'FeCl3', label: 'FeCl3 pur (réel)', unit: 'kg/j', value: Preel.FeCl3_pur },
          ...(Preel.rendement_bioP ? [{ key: 'bioP', label: 'Rendement bio-P', unit: '-', value: Preel.rendement_bioP }] : []),
          { key: 'O2_besoin', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: O2_besoin },
          { key: 'alpha', label: 'Facteur alpha', unit: '-', value: alpha },
          { key: 'K', label: 'Facteur K', unit: '-', value: K.K },
          ...(a.insufflation ? [{ key: 'air_Q', label: "Débit d'air", unit: 'Nm³/h', value: air_Q_Nm3j / 24 }] : []),
          { key: 'sortie_DBO', label: 'DBO5 sortie (réel)', unit: 'mg/L', value: sortie_DBO },
          { key: 'boues_MES', label: 'Boues extraites (réel)', unit: 'kg MES/j', value: boues_MES },
          { key: 'boues_conc', label: 'Concentration des boues', unit: 'g/L', value: boues_concentration },
          { key: 'Pt_out', label: 'Pt sortie (réel)', unit: 'mg/L', value: conc(outR, 'Pt') },
        ],
        electricity: {
          total,
          fixed: fixe,
          detail: {
            aeration: electricite_aeration,
            ...(membrane ? { air_decolmatage: electricite_membranes, permeat: electricite_permeat } : { racleur: electricite_racleur }),
            agitation: electricite_agitation,
            recirculation_MLSS: electricite_recirculation_MLSS,
            ...(membrane ? {} : { recirculation: electricite_recirculation }),
            extraction: electricite_extraction,
          },
        },
        warnings,
      }
    },
  })
}
