// Port de E8_Biostyr.cls — biofiltration Biostyr (lit immergé à flux ascendant,
// billes de polystyrène). Quatre modes : C (carbone), N (nitrification),
// NDNc (nit + pré-dénit classique avec recirculation), NDNs (nit-dénit simultanée).
// Dimensionnement : max des surfaces limitées par Cv_DCO, Cv_MES, Cv_nit et par
// la charge hydraulique de pointe (cellules en lavage déduites), corrigées de la
// température (1,07^(T−12)) et du diamètre des billes (3,6/Ø).
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'
import { CONST, HYP, besoinsO2HS, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const TYPES = [
  { value: 'C', label: 'Biostyr C (carbone)' },
  { value: 'N', label: 'Biostyr N (nitrification)' },
  { value: 'NDNc', label: 'Biostyr NDN classique (pré-dénit + recirculation)' },
  { value: 'NDNs', label: 'Biostyr NDN simultanée' },
]
// charges volumiques maximales admissibles par type et diamètre de billes
const CV_MAX = {
  C: { 4.5: { DCO: 8, MES: 2.8, Q: 6 }, 5: { DCO: 7, MES: 3.5, Q: 8 } },
  N: { 3.6: { DCO: 6, nit: 0.7, MES: 1.8, Q: 4 }, 4: { DCO: 5.4, nit: 0.6, MES: 2.3, Q: 6 }, 4.5: { DCO: 4.8, nit: 0.5, MES: 2.8, Q: 8 } },
  NDN: { 3.6: { DCO: 6, nit: 0.6, MES: 1.8, Q: 6 }, 4: { DCO: 5.4, nit: 0.55, MES: 2.3, Q: 8 }, 4.5: { DCO: 4.8, nit: 0.5, MES: 2.8, Q: 10 } },
}
const SORTIE_MES = { 3.6: 13, 4: 18, 4.5: 23, 5: 28 }
const IMPACT_REGULATION = { horloge: 0.4, sans_variateur: 0.25, variateur: 0.1, avance: 0 }
const H = {
  a_rdt_DCO: 0.91027, b_rdt_DCO: -0.0431,
  a_rdt_DBO: 0.1,
  a_rdt_nit: 0.99648, b_rdt_nit: -0.1762,
  rdt_nit_max_NDNs: 0.9,
  rdt_nit_C: 0.2,
  rdt_denit_max_NDN: 0.75,
  ratio_O2_nit: 4.57, ratio_O2_denit: -2.86,
  cellule_tps_filtration: 23.5,
  v_eau_limite: 6,
  ratio_kgO2_Nm3air: 0.3,
  correctif_T: 1.07, T_reference: 12,
  nb_cellules_file_max: 20, nb_cellules_file_min: 4, nb_cellules_lavage_file: 1,
  pollution_ratio_lim: 0.75,
  ES_tps_retour: 18,
  porosite_lit: 0.3, hauteur_sous_filtre: 0.55, ratio_eau_non_traite_lit: 0.5,
  ratio_Norga_dure_DCO: 0.25 / 100,
  Px2: 8,
  facteur_conso_additionnelle: 0.02,
}
const typeEau = (ratioQEH) => (ratioQEH < 0.175 ? 'concentree' : ratioQEH < 0.325 ? 'standard' : 'diluee')
const convDCO = (T) => (T <= 16 ? 0.46 : T <= 20 ? 0.32 : 0.26)
const ratioDBODenit = (te, type) => (type === 'NDNc' ? { diluee: 7, standard: 5, concentree: 4 }[te] : { diluee: 11, standard: 8, concentree: 5.5 }[te])
const ratioBouesDBO = (te) => (te === 'diluee' ? 0.2 : 0.4)
const isNDN = (c) => c.choices.type === 'NDNc' || c.choices.type === 'NDNs'
const cvTable = (c) => CV_MAX[c.choices.type === 'C' ? 'C' : c.choices.type === 'N' ? 'N' : 'NDN'][Number(c.choices.diametre_media)]

export default defineNode({
  id: 'biostyr',
  label: 'Biostyr',
  short: 'Biostyr',
  family: 'secondaire',
  vba: 'E8_Biostyr.cls',
  description: 'Biofiltration Biostyr en traitement secondaire : C, N, NDN classique (pré-dénitrification avec recirculation) ou NDN simultanée. Les eaux de lavage (eaux sales) doivent être renvoyées vers un traitement.',
  choices: [
    { key: 'type', label: 'Type de Biostyr', default: 'NDNs', options: TYPES },
    { key: 'diametre_media', label: 'Diamètre des billes', default: '4', options: [
      { value: '3.6', label: '3,6 mm (N / NDN)' }, { value: '4', label: '4 mm (N / NDN)' }, { value: '4.5', label: '4,5 mm' }, { value: '5', label: '5 mm (C uniquement)' },
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
    { key: 'Cv_nit', label: 'Charge volumique N nitrifié appliquée (réel)', unit: 'kg/(m³·j)', group: 'Filtration', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_DCO', label: 'DCO en sortie (réel)', unit: 'mg/L', group: 'Sorties', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_NH4', label: 'N-NH4 en sortie (réel)', unit: 'mg/L', group: 'Sorties', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_NO3', label: 'N-NO3 en sortie (réel)', unit: 'mg/L', group: 'Sorties', default: undefined, hint: 'calculée si non forcée' },
    { key: 'air_P_refoulement', label: 'Pression de refoulement des surpresseurs', unit: 'mCE', group: 'Aération', default: 10 },
    { key: 'O2_besoin', label: 'Besoin total en O2', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'O2_rdt_transfert', label: "Rendement de transfert de l'O2", unit: '-', group: 'Aération', default: undefined, hint: 'calculé (corrélation v_eau, hauteur aérée) si non forcé' },
    { key: 'air_Q_Nm3j', label: "Débit d'air process", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'surpresseur_conso_spec', label: 'Conso spécifique surpresseur', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'recirculation_taux', label: 'Taux de recirculation', unit: '-', group: 'Recirculation', default: (c) => (c.choices.type === 'NDNc' ? 1.5 : c.choices.type === 'NDNs' ? 0.8 : 0) },
    { key: 'recirculation_nb_pompe', label: 'Nombre de pompes recirculation', unit: 'u', group: 'Recirculation', default: 2 },
    { key: 'recirculation_P_refoulement', label: 'Pression refoulement recirculation', unit: 'mCE', group: 'Recirculation', default: 4 },
    { key: 'recirculation_tps_fct', label: 'Durée de fonctionnement recirculation', unit: 'h/j', group: 'Recirculation', default: 24 },
    { key: 'recirculation_Q_unitaire', label: 'Débit unitaire recirculation', unit: 'm³/h', group: 'Recirculation', default: undefined, hint: 'calculé si non forcé' },
    { key: 'ratio_ES_volume_media', label: 'Ratio volume eaux sales / volume de média', unit: '-', group: 'Eaux sales', default: (c) => (c.choices.type === 'C' ? 3.5 : 2.5) },
    { key: 'ES_concentration', label: 'Concentration MES des eaux sales', unit: 'mg/L', group: 'Eaux sales', default: undefined, hint: 'calculée (bilan boues) si non forcée' },
    { key: 'ES_nb_pompe', label: 'Nombre de pompes eaux sales', unit: 'u', group: 'Eaux sales', default: 2 },
    { key: 'ES_P_refoulement', label: 'Pression refoulement eaux sales', unit: 'mCE', group: 'Eaux sales', default: 10 },
    { key: 'ES_tps_fonctionnement', label: 'Durée de fonctionnement pompes eaux sales', unit: 'h/j', group: 'Eaux sales', default: 18 },
    { key: 'ES_Q_unitaire', label: 'Débit unitaire pompes eaux sales', unit: 'm³/h', group: 'Eaux sales', default: undefined, hint: 'calculé si non forcé' },
    { key: 'lavage_air_Q', label: "Débit d'air de lavage", unit: 'Nm³/h', group: 'Lavage', default: (c) => (c.choices.type === 'NDNc' ? 15 : 12) * Number(c.choices.surface_unitaire) },
    { key: 'lavage_air_P_refoulement', label: "Pression de l'air de lavage", unit: 'mCE', group: 'Lavage', default: 8 },
    { key: 'lavage_air_tps_fct', label: "Durée quotidienne de l'air de lavage (par cellule)", unit: 'h/j', group: 'Lavage', default: 0.1 },
  ],
  compute(ctx) {
    const { p, site, forced, choices } = ctx
    const warnings = []
    const type = choices.type
    const diam = Number(choices.diametre_media)
    const S_unit = Number(choices.surface_unitaire)
    const cv = cvTable(ctx)
    if (!cv) {
      warnings.push(`Diamètre de billes ${diam} mm indisponible pour le type ${type} — combinaisons valides : C → 4,5/5 mm ; N/NDN → 3,6/4/4,5 mm.`)
      return { outNominal: cloneStream(ctx.inNominal), outReel: cloneStream(ctx.inReel), results: [], electricity: { total: 0, detail: {} }, warnings }
    }
    const facteur_media = 3.6 / diam
    const sortie_MES = SORTIE_MES[diam]
    const h_media = p.hauteur_media
    let h_aere = type === 'NDNc' ? 2.5 : 3.5
    if (forced.hauteur_media != null && forced.hauteur_media < h_aere) h_aere = forced.hauteur_media
    const a1_hep = type === 'C' ? 3.5 : 2.5
    const NH4_min_NDNs = choices.regulation === 'avance' ? 5 : 4
    const Cv_Q_corr_diluee = type === 'N' ? 1.75 : 1.25

    // ---------------- DIMENSIONNEMENT : surface de filtration (eau nominale)
    const N = ctx.inNominal
    const te_nom = typeEau(site.Q_nominal / site.Eq_hab)
    const fT_nom = Math.pow(H.correctif_T, site.T_eau_design - H.T_reference)
    let nb_cellules = 10
    let surface_filtration_min
    if (forced.surface_filtration_min != null) {
      surface_filtration_min = forced.surface_filtration_min
    } else {
      const Cv_DCO_max_T = cv.DCO * fT_nom
      const Cv_nit_max_T = (cv.nit ?? 0) * fT_nom
      const Cv_MES_max_T = cv.MES * fT_nom
      let correction_Cv_Q = 1
      let surface_retenue = 0
      for (let i = 1; i <= 2; i++) {
        surface_retenue = N.DCO / (Cv_DCO_max_T * h_media)
        surface_retenue = Math.max(surface_retenue, N.MES / (Cv_MES_max_T * h_media))
        if (type !== 'C') surface_retenue = Math.max(surface_retenue, N.NH4 / (Cv_nit_max_T * h_media))
        let flag_Q = false
        const nb_file = Math.ceil(nb_cellules / H.nb_cellules_file_max)
        const nb_lavage = H.nb_cellules_lavage_file * nb_file
        let Q_considere = (site.Q_nominal * site.pointe_TP) / CONST.NOMBRE_HEURE_PAR_JOUR + (N.Q - site.Q_nominal) / H.ES_tps_retour
        let S_Q = (Q_considere / (cv.Q * correction_Cv_Q)) * (nb_cellules / (nb_cellules - nb_lavage))
        if (S_Q > surface_retenue) { surface_retenue = S_Q; flag_Q = true }
        if (isNDN(ctx)) {
          Q_considere = ((site.Q_nominal * site.pointe_TS) / CONST.NOMBRE_HEURE_PAR_JOUR + (N.Q - site.Q_nominal) / H.ES_tps_retour) * (1 + p.recirculation_taux)
          S_Q = (Q_considere / (cv.Q * correction_Cv_Q)) * (nb_cellules / (nb_cellules - nb_lavage))
          if (S_Q > surface_retenue) { surface_retenue = S_Q; flag_Q = true }
        }
        if (flag_Q && i === 1) {
          const Cv_DCO_i = N.DCO / (surface_retenue * h_media)
          const Cv_MES_i = N.MES / (surface_retenue * h_media)
          if ((Cv_DCO_i / Cv_DCO_max_T) * (Cv_MES_i / Cv_MES_max_T) < H.pollution_ratio_lim) correction_Cv_Q = Cv_Q_corr_diluee
        }
      }
      surface_filtration_min = surface_retenue
    }
    nb_cellules = forced.nb_cellules ?? Math.max(H.nb_cellules_file_min, Math.ceil(surface_filtration_min / S_unit))
    const surface_filtration = nb_cellules * S_unit
    const nb_file = Math.ceil(nb_cellules / H.nb_cellules_file_max)
    const nb_lavage = H.nb_cellules_lavage_file * nb_file
    const volume_media = surface_filtration * h_media

    // ---------------- passe de traitement (commune nominal / réel)
    const traiter = (s, T, te, reel) => {
      const out = cloneStream(s)
      const fT = Math.pow(H.correctif_T, T - H.T_reference)
      const Norga_dure = H.ratio_Norga_dure_DCO * ((site.DCO_nominal * site.Q_nominal * (reel ? site.NC_DCO : 1)) / (site.Q_nominal * (reel ? site.NC_Q : 1)))
      const MESin_mgL = (s.MES / s.Q) * 1000
      const ES_Q = p.ratio_ES_volume_media * volume_media
      // DCO
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
      // nitrification
      let rdt_nit = H.rdt_nit_C, Cv_nit = 0
      if (type !== 'C') {
        Cv_nit = reel && forced.Cv_nit != null ? forced.Cv_nit : s.NH4 / volume_media
        const Cv_nit_eq = Cv_nit / (facteur_media * fT)
        rdt_nit = H.a_rdt_nit * Math.exp(H.b_rdt_nit * Cv_nit_eq)
        if (type === 'NDNs') rdt_nit = Math.min(rdt_nit, H.rdt_nit_max_NDNs)
      }
      let nitrification, sortie_NH4
      if (reel && forced.sortie_NH4 != null) {
        sortie_NH4 = forced.sortie_NH4
        nitrification = s.NH4 - (sortie_NH4 * s.Q) / 1000
      } else {
        nitrification = s.NH4 * rdt_nit
        sortie_NH4 = ((s.NH4 - nitrification) / s.Q) * 1000
        if (type === 'NDNs' && sortie_NH4 < NH4_min_NDNs) {
          sortie_NH4 = NH4_min_NDNs
          nitrification = s.NH4 - (sortie_NH4 * s.Q) / 1000
        }
      }
      // dénitrification
      let denitrification, sortie_NO3
      if (reel && forced.sortie_NO3 != null) {
        sortie_NO3 = forced.sortie_NO3
        denitrification = s.NO3 + nitrification - (sortie_NO3 * s.Q) / 1000
      } else {
        if (isNDN(ctx)) {
          denitrification = DBO_elim / ratioDBODenit(te, type)
          denitrification = Math.min(denitrification, H.rdt_denit_max_NDN * (s.NO3 + nitrification))
        } else {
          denitrification = 0
        }
        sortie_NO3 = ((s.NO3 + nitrification - denitrification) / s.Q) * 1000
      }
      // eaux sales
      const ratio_boues_MES = type === 'C' ? 1 : 0.86
      const ratio_boues_N = type === 'C' ? 0 : 0.2
      let ES_MES
      if (reel && forced.ES_concentration != null) {
        ES_MES = (forced.ES_concentration * ES_Q) / 1000
      } else {
        ES_MES = ratio_boues_MES * s.MES + ratioBouesDBO(te) * DBO_elim + ratio_boues_N * nitrification
        ES_MES += (MESin_mgL * (H.hauteur_sous_filtre * surface_filtration + H.ratio_eau_non_traite_lit * H.porosite_lit * volume_media)) / 1000
      }
      const ES_concentration = (ES_MES / ES_Q) * 1000
      // O2 et air
      let O2_besoin = (1 - convDCO(reel ? site.T_eau_exploit : site.T_eau_design)) * rdt_DCO * s.DCO + H.ratio_O2_nit * nitrification + H.ratio_O2_denit * denitrification + besoinsO2HS(s.Sh)
      if (reel && forced.O2_besoin != null) O2_besoin = forced.O2_besoin
      let O2_rdt_transfert
      if (forced.O2_rdt_transfert != null) {
        O2_rdt_transfert = forced.O2_rdt_transfert
      } else {
        const v_eau = s.Q / (S_unit * (nb_cellules - nb_lavage) * H.cellule_tps_filtration)
        const hep = 0.5 + a1_hep * (h_media / nb_cellules)
        if (type === 'NDNc') O2_rdt_transfert = ((6.2 + 0.375 * v_eau) * (h_aere - 0.4) * (10.7 + hep + h_aere / 2)) / 12.7
        else if (v_eau < H.v_eau_limite) O2_rdt_transfert = ((6.2 + 0.375 * v_eau) * h_aere * (10.7 + hep + h_aere / 2)) / 13.4
        else O2_rdt_transfert = (8.45 * h_aere * (10.7 + hep + h_aere / 2)) / 13.4
        O2_rdt_transfert = (O2_rdt_transfert / 100) * (1 - IMPACT_REGULATION[choices.regulation])
      }
      let air_Q_Nm3j = O2_besoin / (O2_rdt_transfert * H.ratio_kgO2_Nm3air)
      if (reel && forced.air_Q_Nm3j != null) air_Q_Nm3j = forced.air_Q_Nm3j
      // sorties
      out.Q = s.Q - ES_Q
      out.DCO = reel ? (sortie_DCO * out.Q) / 1000 : s.DCO * (1 - rdt_DCO)
      out.DBO = s.DBO * (1 - rdt_DBO)
      out.MES = (sortie_MES * out.Q) / 1000
      out.NH4 = (sortie_NH4 * out.Q) / 1000
      out.NK = out.NH4 + (Norga_dure * out.Q) / 1000
      out.NO3 = (sortie_NO3 * out.Q) / 1000
      // Pt inchangé (fidèle au VBA : la part assimilée est réglée dans la file boues)
      out.Sh = 0
      return { out, ES_Q, ES_MES, ES_concentration, rdt_DCO, rdt_DBO, rdt_nit, Cv_DCO, Cv_nit, sortie_DCO, sortie_NH4, sortie_NO3, nitrification, denitrification, O2_besoin, O2_rdt_transfert, air_Q_Nm3j, DBO_elim }
    }
    const te_reel = typeEau((site.Q_nominal * site.NC_Q) / site.Eq_hab)
    const nom = traiter(ctx.inNominal, site.T_eau_design, te_nom, false)
    const reel = traiter(ctx.inReel, site.T_eau_exploit, te_reel, true)

    // ---------------- ELECTRICITE (sur le réel)
    if (choices.surpresseur === 'roots' && p.air_P_refoulement >= H.Px2) warnings.push(`Pression de refoulement (${p.air_P_refoulement.toFixed(1)} mCE) trop élevée pour des surpresseurs roots.`)
    const electricite_aeration = (reel.air_Q_Nm3j * p.air_P_refoulement * p.surpresseur_conso_spec) / 1000
    const electricite_lavage_air = (p.lavage_air_Q * p.lavage_air_tps_fct * nb_cellules * p.lavage_air_P_refoulement * p.surpresseur_conso_spec) / 1000
    let electricite_recirculation = 0
    if (isNDN(ctx) && p.recirculation_taux > 0) {
      const Qrec = p.recirculation_taux * ctx.inReel.Q
      const rep = repartitionPompes(Qrec, { nb: p.recirculation_nb_pompe, tps: p.recirculation_tps_fct }, { nb: forced.recirculation_nb_pompe, tps: forced.recirculation_tps_fct, Q_unitaire: forced.recirculation_Q_unitaire })
      if (rep.incoherence) warnings.push('Incohérence sur les pompes de recirculation.')
      const rdt = rendementPompeGlobal(rep.Q_unitaire, p.recirculation_P_refoulement, 0.7)
      electricite_recirculation = electricitePompage(rep.Q_unitaire, p.recirculation_P_refoulement, rdt, rep.nb, rep.tps)
    }
    const es = repartitionPompes(reel.ES_Q, { nb: p.ES_nb_pompe, tps: p.ES_tps_fonctionnement }, { nb: forced.ES_nb_pompe, tps: forced.ES_tps_fonctionnement, Q_unitaire: forced.ES_Q_unitaire }, H.ES_tps_retour)
    if (es.incoherence) warnings.push('Incohérence sur les pompes des eaux sales.')
    const rdt_es = rendementPompeGlobal(es.Q_unitaire, p.ES_P_refoulement, 0.7)
    const electricite_ES = electricitePompage(es.Q_unitaire, p.ES_P_refoulement, rdt_es, es.nb, es.tps)
    const k = 1 + H.facteur_conso_additionnelle
    const total = (electricite_aeration + electricite_lavage_air + electricite_recirculation + electricite_ES) * k

    return {
      outNominal: nom.out,
      outReel: reel.out,
      eauxSales: { origine: type === 'C' ? 'II_biostyr_C' : type === 'N' ? 'II_biostyr_N' : 'II_biostyr_NDN', Q: reel.ES_Q, MES: reel.ES_MES, DCO: convDCO(site.T_eau_exploit) * reel.rdt_DCO * ctx.inReel.DCO, DBO: 0, NK: (reel.ES_Q * reel.sortie_NH4) / 1000, NH4: (reel.ES_Q * reel.sortie_NH4) / 1000, NO3: (reel.ES_Q * reel.sortie_NO3) / 1000, Pt: 0, MV_MES: 0.85 },
      results: [
        { key: 'S_min', label: 'Surface de filtration minimale', unit: 'm²', value: surface_filtration_min },
        { key: 'nb_cellules', label: 'Cellules de filtration', unit: 'u', value: nb_cellules },
        { key: 'S', label: 'Surface réelle de filtration', unit: 'm²', value: surface_filtration },
        { key: 'V_media', label: 'Volume de média', unit: 'm³', value: volume_media },
        { key: 'Cv_DCO', label: 'Cv DCO appliquée (réel)', unit: 'kg/(m³·j)', value: reel.Cv_DCO },
        ...(type !== 'C' ? [{ key: 'Cv_nit', label: 'Cv N-NH4 appliquée (réel)', unit: 'kg/(m³·j)', value: reel.Cv_nit }] : []),
        { key: 'rdt_DCO', label: 'Rendement DCO (réel)', unit: '-', value: reel.rdt_DCO },
        ...(type !== 'C' ? [{ key: 'rdt_nit', label: 'Rendement de nitrification (réel)', unit: '-', value: reel.rdt_nit }] : []),
        { key: 'nit', label: 'Nitrification (réel)', unit: 'kg N/j', value: reel.nitrification },
        ...(isNDN(ctx) ? [{ key: 'denit', label: 'Dénitrification (réel)', unit: 'kg N/j', value: reel.denitrification }] : []),
        { key: 'sortie_DCO', label: 'DCO sortie (réel)', unit: 'mg/L', value: reel.sortie_DCO },
        { key: 'sortie_NH4', label: 'N-NH4 sortie (réel)', unit: 'mg/L', value: reel.sortie_NH4 },
        { key: 'sortie_NO3', label: 'N-NO3 sortie (réel)', unit: 'mg/L', value: reel.sortie_NO3 },
        { key: 'O2', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: reel.O2_besoin },
        { key: 'rdt_transfert', label: "Rendement de transfert de l'O2 (réel)", unit: '-', value: reel.O2_rdt_transfert },
        { key: 'air', label: "Débit d'air (réel)", unit: 'Nm³/h', value: reel.air_Q_Nm3j / 24 },
        { key: 'ES_Q', label: 'Eaux sales (réel)', unit: 'm³/j', value: reel.ES_Q },
        { key: 'ES_MES', label: 'MES des eaux sales (réel)', unit: 'kg/j', value: reel.ES_MES },
        { key: 'ES_conc', label: 'Concentration des eaux sales (réel)', unit: 'mg/L', value: reel.ES_concentration },
      ],
      electricity: { total, fixed: electricite_lavage_air * k, detail: { aeration: electricite_aeration * k, lavage_air: electricite_lavage_air * k, recirculation: electricite_recirculation * k, eaux_sales: electricite_ES * k } },
      warnings,
    }
  },
})
