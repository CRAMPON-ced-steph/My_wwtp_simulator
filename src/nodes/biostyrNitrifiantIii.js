// Port de F1_Biostyr_N_III.cls — Biostyr nitrifiant en tertiaire.
// Structure de E8 type N, sans recirculation, avec ses propres charges
// volumiques (nitrification tertiaire : Cv_nit 1,1–1,3 kg/(m³·j)) et sa
// conversion DCO propre (0,47 / 0,32 / 0,23 selon T).
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'
import { CONST, HYP, besoinsO2HS, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const CV_MAX = {
  3.6: { DCO: 6, nit: 1.3, MES: 1.5, Q: 12 },
  4: { DCO: 5.4, nit: 1.2, MES: 1.5, Q: 14 },
  4.5: { DCO: 4.8, nit: 1.1, MES: 2, Q: 16 },
}
const SORTIE_MES = { 3.6: 13, 4: 18, 4.5: 23, 5: 28 }
const IMPACT_REGULATION = { horloge: 0.4, sans_variateur: 0.25, variateur: 0.1, avance: 0 }
const H = {
  a_rdt_DCO: 0.91027, b_rdt_DCO: -0.0431,
  a_rdt_DBO: 0.1,
  a_rdt_nit: 0.99648, b_rdt_nit: -0.1762,
  ratio_O2_nit: 4.57,
  cellule_tps_filtration: 23.5,
  v_eau_limite: 6,
  ratio_kgO2_Nm3air: 0.3,
  correctif_T: 1.07, T_reference: 12,
  nb_cellules_file_max: 20, nb_cellules_file_min: 4, nb_cellules_lavage_file: 1,
  ES_tps_retour: 18,
  ratio_boues_MES_elimine: 0.86,
  ratio_boues_N_nitrifie: 0.2,
  a1_hep: 2.5, h_aere: 3.5,
  Px2: 8,
  facteur_conso_additionnelle: 0.02,
}
const convDCO = (T) => (T <= 16 ? 0.47 : T <= 20 ? 0.32 : 0.23)
const typeEau = (r) => (r < 0.175 ? 'concentree' : r < 0.325 ? 'standard' : 'diluee')
const ratioBouesDBO = (te) => (te === 'diluee' ? 0.2 : 0.4)

export default defineNode({
  id: 'biostyr-nitrifiant-iii',
  label: 'Biostyr nitrifiant (tertiaire)',
  short: 'Biostyr N III',
  family: 'tertiaire',
  vba: 'F1_Biostyr_N_III.cls',
  description: "Biofiltre Biostyr nitrifiant en position tertiaire (derrière un traitement du carbone). Charges volumiques de nitrification élevées, pas de recirculation. Les eaux de lavage doivent être renvoyées en tête.",
  choices: [
    { key: 'diametre_media', label: 'Diamètre des billes', default: '4', options: [
      { value: '3.6', label: '3,6 mm' }, { value: '4', label: '4 mm' }, { value: '4.5', label: '4,5 mm' },
    ] },
    { key: 'surface_unitaire', label: 'Surface unitaire des cellules', default: '100', options: ['50', '70', '100', '130', '170'].map((v) => ({ value: v, label: v + ' m²' })) },
    { key: 'regulation', label: "Régulation de l'aération", default: 'variateur', options: [
      { value: 'horloge', label: 'sur horloge' }, { value: 'sans_variateur', label: 'sans variateur' }, { value: 'variateur', label: 'classique – avec variateur' }, { value: 'avance', label: 'avancé' },
    ] },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
      { value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' },
    ] },
  ],
  params: [
    { key: 'hauteur_media', label: 'Hauteur de média', unit: 'm', group: 'Filtration', default: 3.5 },
    { key: 'surface_filtration_min', label: 'Surface de filtration minimale', unit: 'm²', group: 'Filtration', default: undefined, hint: 'calculée si non forcée' },
    { key: 'nb_cellules', label: 'Nombre total de cellules', unit: 'u', group: 'Filtration', default: undefined, hint: 'calculé si non forcé' },
    { key: 'Cv_DCO', label: 'Charge volumique DCO appliquée (réel)', unit: 'kg/(m³·j)', group: 'Filtration', default: undefined, hint: 'calculée si non forcée' },
    { key: 'Cv_nit', label: 'Charge volumique N-NH4 appliquée (réel)', unit: 'kg/(m³·j)', group: 'Filtration', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_DCO', label: 'DCO en sortie (réel)', unit: 'mg/L', group: 'Sorties', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_NH4', label: 'N-NH4 en sortie (réel)', unit: 'mg/L', group: 'Sorties', default: undefined, hint: 'calculée si non forcée' },
    { key: 'air_P_refoulement', label: 'Pression de refoulement des surpresseurs', unit: 'mCE', group: 'Aération', default: 10 },
    { key: 'O2_besoin', label: 'Besoin total en O2', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_rdt_transfert', label: "Rendement de transfert de l'O2", unit: '-', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'air_Q_Nm3j', label: "Débit d'air process", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'surpresseur_conso_spec', label: 'Conso spécifique surpresseur', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'ratio_ES_volume_media', label: 'Ratio volume eaux sales / volume de média', unit: '-', group: 'Eaux sales', default: 2.5 },
    { key: 'ES_concentration', label: 'Concentration MES des eaux sales', unit: 'mg/L', group: 'Eaux sales', default: undefined, hint: 'calculée si non forcée' },
    { key: 'ES_nb_pompe', label: 'Nombre de pompes eaux sales', unit: 'u', group: 'Eaux sales', default: 2 },
    { key: 'ES_P_refoulement', label: 'Pression refoulement eaux sales', unit: 'mCE', group: 'Eaux sales', default: 10 },
    { key: 'ES_tps_fonctionnement', label: 'Durée de fonctionnement pompes eaux sales', unit: 'h/j', group: 'Eaux sales', default: 18 },
    { key: 'ES_Q_unitaire', label: 'Débit unitaire pompes eaux sales', unit: 'm³/h', group: 'Eaux sales', default: undefined, hint: 'calculé si non forcé' },
    { key: 'lavage_air_Q', label: "Débit d'air de lavage", unit: 'Nm³/h', group: 'Lavage', default: (c) => 12 * Number(c.choices.surface_unitaire) },
    { key: 'lavage_air_P_refoulement', label: "Pression de l'air de lavage", unit: 'mCE', group: 'Lavage', default: 8 },
    { key: 'lavage_air_tps_fct', label: "Durée quotidienne de l'air de lavage (par cellule)", unit: 'h/j', group: 'Lavage', default: 0.1 },
  ],
  compute(ctx) {
    const { p, site, forced, choices } = ctx
    const warnings = []
    const diam = Number(choices.diametre_media)
    const S_unit = Number(choices.surface_unitaire)
    const cv = CV_MAX[diam]
    const sortie_MES = SORTIE_MES[diam]
    const facteur_media = 3.6 / diam
    const h_media = p.hauteur_media
    let h_aere = H.h_aere
    if (forced.hauteur_media != null && forced.hauteur_media < h_aere) h_aere = forced.hauteur_media

    // ---------------- surface (nominal)
    const N = ctx.inNominal
    const fT_nom = Math.pow(H.correctif_T, site.T_eau_design - H.T_reference)
    let nb_cellules = 10
    let surface_filtration_min
    if (forced.surface_filtration_min != null) surface_filtration_min = forced.surface_filtration_min
    else {
      let S = N.DCO / (cv.DCO * fT_nom * h_media)
      S = Math.max(S, N.MES / (cv.MES * fT_nom * h_media))
      S = Math.max(S, N.NH4 / (cv.nit * fT_nom * h_media))
      const nb_file = Math.ceil(nb_cellules / H.nb_cellules_file_max)
      const nb_lavage = H.nb_cellules_lavage_file * nb_file
      const Q_considere = (site.Q_nominal * site.pointe_TP) / CONST.NOMBRE_HEURE_PAR_JOUR + (N.Q - site.Q_nominal) / H.ES_tps_retour
      S = Math.max(S, (Q_considere / cv.Q) * (nb_cellules / (nb_cellules - nb_lavage)))
      surface_filtration_min = S
    }
    nb_cellules = forced.nb_cellules ?? Math.max(H.nb_cellules_file_min, Math.ceil(surface_filtration_min / S_unit))
    const surface_filtration = nb_cellules * S_unit
    const nb_lavage = H.nb_cellules_lavage_file * Math.ceil(nb_cellules / H.nb_cellules_file_max)
    const volume_media = surface_filtration * h_media

    const traiter = (s, T, te, reel) => {
      const out = cloneStream(s)
      const fT = Math.pow(H.correctif_T, T - H.T_reference)
      const ES_Q = p.ratio_ES_volume_media * volume_media
      let Cv_DCO = reel && forced.Cv_DCO != null ? forced.Cv_DCO : s.DCO / volume_media
      const Cv_DCO_eq = Cv_DCO / (facteur_media * fT)
      let rdt_DCO, sortie_DCO
      if (reel && forced.sortie_DCO != null) {
        sortie_DCO = forced.sortie_DCO
        rdt_DCO = 1 - ((sortie_DCO * (s.Q - ES_Q)) / 1000) / s.DCO
      } else {
        rdt_DCO = H.a_rdt_DCO * Math.exp(H.b_rdt_DCO * Cv_DCO_eq)
        rdt_DCO = 1 - ((s.Q - ES_Q) / s.Q) * (1 - rdt_DCO)
        sortie_DCO = ((s.DCO * (1 - rdt_DCO)) / (s.Q - ES_Q)) * 1000
      }
      // écart assumé : le VBA ne plafonne pas rdt_DBO = rdt_DCO + 0,1 (DBO négative possible)
      const rdt_DBO = Math.min(1, rdt_DCO + H.a_rdt_DBO)
      const DBO_elim = s.DBO * rdt_DBO
      const DCO_conso_bio = (1 - convDCO(T)) * rdt_DCO * s.DCO
      let Cv_nit = reel && forced.Cv_nit != null ? forced.Cv_nit : s.NH4 / volume_media
      const Cv_nit_eq = Cv_nit / (facteur_media * fT)
      let rdt_nit = H.a_rdt_nit * Math.exp(H.b_rdt_nit * Cv_nit_eq)
      let nitrification, sortie_NH4
      if (reel && forced.sortie_NH4 != null) {
        sortie_NH4 = forced.sortie_NH4
        nitrification = s.NH4 - (sortie_NH4 * s.Q) / 1000
      } else {
        nitrification = s.NH4 * rdt_nit
        sortie_NH4 = ((s.NH4 - nitrification) / s.Q) * 1000
      }
      const sortie_NO3 = ((s.NO3 + nitrification) / s.Q) * 1000
      const MES_elimine = s.MES - (sortie_MES * (s.Q - ES_Q)) / 1000
      let ES_MES
      if (reel && forced.ES_concentration != null) ES_MES = (forced.ES_concentration * ES_Q) / 1000
      else ES_MES = H.ratio_boues_MES_elimine * MES_elimine + ratioBouesDBO(te) * DBO_elim + H.ratio_boues_N_nitrifie * nitrification
      let O2_besoin = DCO_conso_bio + H.ratio_O2_nit * nitrification + besoinsO2HS(s.Sh)
      if (reel && forced.O2_besoin != null) O2_besoin = forced.O2_besoin
      let O2_rdt_transfert
      if (forced.O2_rdt_transfert != null) O2_rdt_transfert = forced.O2_rdt_transfert
      else {
        const v_eau = s.Q / (S_unit * (nb_cellules - nb_lavage) * H.cellule_tps_filtration)
        const hep = 0.5 + H.a1_hep * (h_media / nb_cellules)
        if (v_eau < H.v_eau_limite) O2_rdt_transfert = ((6.2 + 0.375 * v_eau) * h_aere * (10.7 + hep + h_aere / 2)) / 13.4
        else O2_rdt_transfert = (8.45 * h_aere * (10.7 + hep + h_aere / 2)) / 13.4
        O2_rdt_transfert = (O2_rdt_transfert / 100) * (1 - IMPACT_REGULATION[choices.regulation])
      }
      let air_Q_Nm3j = O2_besoin / (O2_rdt_transfert * H.ratio_kgO2_Nm3air)
      if (reel && forced.air_Q_Nm3j != null) air_Q_Nm3j = forced.air_Q_Nm3j
      out.Q = s.Q - ES_Q
      out.DCO = reel ? (sortie_DCO * out.Q) / 1000 : s.DCO * (1 - rdt_DCO)
      out.DBO = s.DBO * (1 - rdt_DBO)
      out.MES = (sortie_MES * out.Q) / 1000
      out.NH4 = (sortie_NH4 * out.Q) / 1000
      out.NK = out.NH4
      out.NO3 = (sortie_NO3 * out.Q) / 1000
      out.Sh = 0
      return { out, ES_Q, ES_MES, rdt_DCO, rdt_nit, Cv_DCO, Cv_nit, sortie_DCO, sortie_NH4, sortie_NO3, nitrification, O2_besoin, O2_rdt_transfert, air_Q_Nm3j }
    }
    const nom = traiter(ctx.inNominal, site.T_eau_design, typeEau(site.Q_nominal / site.Eq_hab), false)
    const reel = traiter(ctx.inReel, site.T_eau_exploit, typeEau((site.Q_nominal * site.NC_Q) / site.Eq_hab), true)

    if (choices.surpresseur === 'roots' && p.air_P_refoulement >= H.Px2) warnings.push(`Pression de refoulement (${p.air_P_refoulement.toFixed(1)} mCE) trop élevée pour des surpresseurs roots.`)
    const electricite_aeration = (reel.air_Q_Nm3j * p.air_P_refoulement * p.surpresseur_conso_spec) / 1000
    const electricite_lavage_air = (p.lavage_air_Q * p.lavage_air_tps_fct * nb_cellules * p.lavage_air_P_refoulement * p.surpresseur_conso_spec) / 1000
    const es = repartitionPompes(reel.ES_Q, { nb: p.ES_nb_pompe, tps: p.ES_tps_fonctionnement }, { nb: forced.ES_nb_pompe, tps: forced.ES_tps_fonctionnement, Q_unitaire: forced.ES_Q_unitaire }, H.ES_tps_retour)
    if (es.incoherence) warnings.push('Incohérence sur les pompes des eaux sales.')
    const rdt_es = rendementPompeGlobal(es.Q_unitaire, p.ES_P_refoulement, 0.7)
    const electricite_ES = electricitePompage(es.Q_unitaire, p.ES_P_refoulement, rdt_es, es.nb, es.tps)
    const k = 1 + H.facteur_conso_additionnelle
    const total = (electricite_aeration + electricite_lavage_air + electricite_ES) * k

    return {
      outNominal: nom.out,
      outReel: reel.out,
      eauxSales: { origine: 'III_biostyr_N', Q: reel.ES_Q, MES: reel.ES_MES, DCO: convDCO(site.T_eau_exploit) * reel.rdt_DCO * ctx.inReel.DCO, DBO: 0, NK: (reel.ES_Q * reel.sortie_NH4) / 1000, NH4: (reel.ES_Q * reel.sortie_NH4) / 1000, NO3: (reel.ES_Q * reel.sortie_NO3) / 1000, Pt: 0, MV_MES: 0.85 },
      results: [
        { key: 'S_min', label: 'Surface de filtration minimale', unit: 'm²', value: surface_filtration_min },
        { key: 'nb', label: 'Cellules de filtration', unit: 'u', value: nb_cellules },
        { key: 'S', label: 'Surface réelle de filtration', unit: 'm²', value: surface_filtration },
        { key: 'Cv_nit', label: 'Cv N-NH4 appliquée (réel)', unit: 'kg/(m³·j)', value: reel.Cv_nit },
        { key: 'rdt_nit', label: 'Rendement de nitrification (réel)', unit: '-', value: reel.rdt_nit },
        { key: 'nit', label: 'Nitrification (réel)', unit: 'kg N/j', value: reel.nitrification },
        { key: 'sortie_NH4', label: 'N-NH4 sortie (réel)', unit: 'mg/L', value: reel.sortie_NH4 },
        { key: 'sortie_NO3', label: 'N-NO3 sortie (réel)', unit: 'mg/L', value: reel.sortie_NO3 },
        { key: 'O2', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: reel.O2_besoin },
        { key: 'air', label: "Débit d'air (réel)", unit: 'Nm³/h', value: reel.air_Q_Nm3j / 24 },
        { key: 'ES_Q', label: 'Eaux sales (réel)', unit: 'm³/j', value: reel.ES_Q },
        { key: 'ES_MES', label: 'MES des eaux sales (réel)', unit: 'kg/j', value: reel.ES_MES },
      ],
      electricity: { total, fixed: electricite_lavage_air * k, detail: { aeration: electricite_aeration * k, lavage_air: electricite_lavage_air * k, eaux_sales: electricite_ES * k } },
      warnings,
    }
  },
})
