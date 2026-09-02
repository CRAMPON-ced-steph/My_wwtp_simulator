// Port de E9_Biostyr_PDN.cls — Biostyr en post-dénitrification (méthanol).
// Une fraction du débit (inlet_ratio_admis) est admise sur le filtre, calculée
// pour respecter la garantie NGL en sortie de filière ; le reste by-passe.
// Le méthanol est dosé à facteur_DCO(T) × dénitrification ; l'O2 dissous entrant
// (6 mg/L) consomme une part de la capacité de dénitrification (ratio 0,35).
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'
import { CONST, HYP, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const CV_MAX = { 4.5: { DCO: 10, denit: 1.7, Q: 20 }, 5: { DCO: 12, denit: 1.5, Q: 25 } }
const RDT_MES = { 4.5: [0, 0.1, 0.2, 0.25], 5: [0, 0, 0.1, 0.13] } // gammes <15 / <20 / <25 / ≥25 mg/L
const H = {
  inlet_O2_mgL: 6,
  ratio_denit_O2dissous: 0.35,
  DCO_methanol: 1.5,
  DBO_methanol: 0.94,
  correctif_T: 1.07, T_reference: 12,
  ratio_boues_MES_elimine: 1,
  ratio_boues_DBO_elimine_std: 0.4,
  ES_tps_retour: 18,
  facteur_conso_additionnelle: 0.02,
  ratio_elec_methanol_pompe: 1.16, // Wh/m³ d'eau entrante (0,73 Angers, 1,62 Nantes, 1,12 Douarnenez)
}
const facteurDCO = (T) => (T < 5 ? 5 : T < 15 ? 4 + (T - 5) * 0.1 : 4)
const typeEau = (r) => (r < 0.175 ? 'concentree' : r < 0.325 ? 'standard' : 'diluee')
const gammeMES = (m) => (m < 15 ? 0 : m < 20 ? 1 : m < 25 ? 2 : 3)

export function makeBiostyrPdn(cfg) {
  return defineNode({
  id: cfg.id,
  label: cfg.label,
  short: cfg.short,
  family: cfg.family,
  vba: cfg.vba,
  ported: true,
  description: cfg.description,
  choices: [
    { key: 'diametre_media', label: 'Diamètre des billes', default: '4.5', options: [{ value: '4.5', label: '4,5 mm' }, { value: '5', label: '5 mm' }] },
    { key: 'surface_unitaire', label: 'Surface unitaire des cellules', default: '100', options: ['50', '70', '100', '130', '170'].map((v) => ({ value: v, label: v + ' m²' })) },
    { key: 'surpresseur', label: 'Type de surpresseur (air de lavage)', default: 'roots', options: [{ value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' }] },
  ],
  params: [
    { key: 'NO3_garantie', label: 'N-NO3 garanti en sortie de filtre', unit: 'mg/L', group: 'Objectif', default: 2 },
    { key: 'NO3_visee', label: 'N-NO3 visé après mélange avec le by-pass', unit: 'mg/L', group: 'Objectif', default: undefined, hint: 'calculé (garantie NGL − NK sortie) si non forcé' },
    { key: 'hauteur_media', label: 'Hauteur de média', unit: 'm', group: 'Filtration', default: 2.5 },
    { key: 'surface_filtration_min', label: 'Surface de filtration minimale', unit: 'm²', group: 'Filtration', default: undefined, hint: 'calculée si non forcée' },
    { key: 'nb_cellules', label: 'Nombre de cellules', unit: 'u', group: 'Filtration', default: undefined, hint: 'calculé si non forcé' },
    { key: 'methanol_flux', label: 'Méthanol pur (réel)', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculé si non forcé' },
    { key: 'ratio_ES_volume_media', label: 'Ratio volume eaux sales / volume de média', unit: '-', group: 'Eaux sales', default: 2.5 },
    { key: 'ES_concentration', label: 'Concentration MES des eaux sales', unit: 'mg/L', group: 'Eaux sales', default: undefined, hint: 'calculée si non forcée' },
    { key: 'ES_nb_pompe', label: 'Nombre de pompes eaux sales', unit: 'u', group: 'Eaux sales', default: 2 },
    { key: 'ES_P_refoulement', label: 'Pression refoulement eaux sales', unit: 'mCE', group: 'Eaux sales', default: 10 },
    { key: 'ES_tps_fonctionnement', label: 'Durée de fonctionnement pompes eaux sales', unit: 'h/j', group: 'Eaux sales', default: 18 },
    { key: 'ES_Q_unitaire', label: 'Débit unitaire pompes eaux sales', unit: 'm³/h', group: 'Eaux sales', default: undefined, hint: 'calculé si non forcé' },
    { key: 'lavage_air_Q', label: "Débit d'air de lavage", unit: 'Nm³/h', group: 'Lavage', default: (c) => 12 * Number(c.choices.surface_unitaire) },
    { key: 'lavage_air_P_refoulement', label: "Pression de l'air de lavage", unit: 'mCE', group: 'Lavage', default: 7 },
    { key: 'lavage_air_tps_fct', label: "Durée quotidienne de l'air de lavage (par cellule)", unit: 'h/j', group: 'Lavage', default: 0.1 },
    { key: 'surpresseur_conso_spec', label: 'Conso spécifique surpresseur', unit: 'Wh/(Nm³·mCE)', group: 'Lavage', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
  ],
  compute(ctx) {
    const { p, site, forced, choices } = ctx
    const warnings = []
    const diam = Number(choices.diametre_media)
    const S_unit = Number(choices.surface_unitaire)
    const cv = CV_MAX[diam]
    const h_media = p.hauteur_media

    // état partagé : ES_Q du dimensionnement est réutilisé dans le calcul du
    // ratio admis (première passe à 0, comportement VBA)
    const st = { ES_Q: 0, surface_filtration: 0, nb_cellules: 0 }

    const traiter = (s, T, te, reel) => {
      const out = cloneStream(s)
      const correction_denit_O2 = H.ratio_denit_O2dissous * H.inlet_O2_mgL
      let NO3_visee = reel && forced.NO3_visee != null ? forced.NO3_visee : site.NGL_garantie - (1000 * s.NK) / s.Q
      const inlet_NO3_mgL = (1000 * s.NO3) / s.Q + correction_denit_O2
      let ratio_admis = (NO3_visee * (s.Q - st.ES_Q) - inlet_NO3_mgL * s.Q + p.NO3_garantie * st.ES_Q) / (s.Q * (p.NO3_garantie - inlet_NO3_mgL))
      if (ratio_admis > 1) { ratio_admis = 1; NO3_visee = p.NO3_garantie }
      else if (ratio_admis < 0) { ratio_admis = 0.01; NO3_visee = inlet_NO3_mgL }
      const inlet_Q = ratio_admis * s.Q
      // surface (au dimensionnement uniquement)
      if (!reel) {
        let S_min
        if (forced.surface_filtration_min != null) S_min = forced.surface_filtration_min
        else {
          const fT = Math.pow(H.correctif_T, T - H.T_reference)
          S_min = (ratio_admis * s.DCO) / (cv.DCO * fT * h_media)
          S_min = Math.max(S_min, (s.NO3 * ratio_admis) / (cv.denit * fT * h_media))
          const Q_considere = ratio_admis * ((site.Q_nominal * site.pointe_TP) / CONST.NOMBRE_HEURE_PAR_JOUR + (s.Q - site.Q_nominal) / H.ES_tps_retour)
          S_min = Math.max(S_min, Q_considere / cv.Q)
          st.surface_filtration_min = S_min
        }
        st.nb_cellules = forced.nb_cellules ?? Math.ceil(S_min / S_unit)
        st.surface_filtration = st.nb_cellules * S_unit
        st.ES_Q = p.ratio_ES_volume_media * st.surface_filtration * h_media
      }
      const ES_Q = st.ES_Q
      const facteur_repartition = ES_Q / s.Q
      const ES = {
        Q: ES_Q,
        NO3: (p.NO3_garantie * ES_Q) / 1000,
        DCO: s.DCO * facteur_repartition,
        DBO: s.DBO * facteur_repartition,
        NK: s.NK * facteur_repartition,
        NH4: s.NH4 * facteur_repartition,
        Pt: s.Pt * facteur_repartition,
      }
      // dénitrification (y compris consommation de l'O2 dissous entrant)
      let denitrification = ratio_admis * s.NO3 - ES.NO3 - (p.NO3_garantie * (inlet_Q - ES_Q)) / 1000
      denitrification += (H.ratio_denit_O2dissous * H.inlet_O2_mgL * inlet_Q) / 1000
      // méthanol
      let methanol_flux = (facteurDCO(T) * denitrification) / H.DCO_methanol
      if (reel && forced.methanol_flux != null) methanol_flux = forced.methanol_flux
      const DCO_apportee = H.DCO_methanol * methanol_flux
      // MES
      const inlet_MES_mgL = (s.MES / s.Q) * 1000
      const rdt_MES = RDT_MES[diam][gammeMES(inlet_MES_mgL)]
      const sortie_MES_mgL = inlet_MES_mgL * (1 - rdt_MES)
      // F2 (tertiaire) calcule les MES éliminées sur (1 − ratio) — écart du VBA reproduit
      const MES_elimine = ((cfg.mesElimineSurBypass ? 1 - ratio_admis : ratio_admis) * (inlet_MES_mgL - sortie_MES_mgL) * s.Q) / 1000
      let ES_MES
      if (reel && forced.ES_concentration != null) ES_MES = (forced.ES_concentration * ES_Q) / 1000
      else ES_MES = H.ratio_boues_MES_elimine * MES_elimine + H.ratio_boues_DBO_elimine_std * (H.DBO_methanol / H.DCO_methanol) * DCO_apportee
      // sorties
      out.Q = s.Q - ES_Q
      out.DCO = s.DCO - ES.DCO
      out.DBO = s.DBO - ES.DBO
      out.MES = (1 - ratio_admis) * s.MES + (sortie_MES_mgL * (inlet_Q - ES_Q)) / 1000
      out.NH4 = s.NH4 - ES.NH4
      out.NK = s.NK - ES.NK
      out.NO3 = (NO3_visee * out.Q) / 1000
      out.Pt = s.Pt - ES.Pt
      return { out, ratio_admis, inlet_Q, NO3_visee, denitrification, methanol_flux, DCO_apportee, ES_Q, ES_MES, ES_concentration: (ES_MES / ES_Q) * 1000, ES, sortie_MES_mgL }
    }
    const nom = traiter(ctx.inNominal, site.T_eau_design, typeEau(site.Q_nominal / site.Eq_hab), false)
    const reel = traiter(ctx.inReel, site.T_eau_exploit, typeEau((site.Q_nominal * site.NC_Q) / site.Eq_hab), true)

    // électricité : lavage air + pompes eaux sales + dosage méthanol, ×1,02
    const electricite_lavage_air = (p.lavage_air_Q * p.lavage_air_tps_fct * st.nb_cellules * p.lavage_air_P_refoulement * p.surpresseur_conso_spec) / 1000
    const es = repartitionPompes(reel.ES_Q, { nb: p.ES_nb_pompe, tps: p.ES_tps_fonctionnement }, { nb: forced.ES_nb_pompe, tps: forced.ES_tps_fonctionnement, Q_unitaire: forced.ES_Q_unitaire }, H.ES_tps_retour)
    if (es.incoherence) warnings.push('Incohérence sur les pompes des eaux sales.')
    const rdt_es = rendementPompeGlobal(es.Q_unitaire, p.ES_P_refoulement, 0.7)
    const electricite_ES = electricitePompage(es.Q_unitaire, p.ES_P_refoulement, rdt_es, es.nb, es.tps)
    const electricite_methanol = (H.ratio_elec_methanol_pompe * reel.inlet_Q) / 1000
    const k = 1 + H.facteur_conso_additionnelle
    const total = (electricite_lavage_air + electricite_ES + electricite_methanol) * k

    return {
      outNominal: nom.out,
      outReel: reel.out,
      eauxSales: { origine: cfg.origine, Q: reel.ES_Q, MES: reel.ES_MES, DCO: reel.ES.DCO, DBO: reel.ES.DBO, NK: reel.ES.NK, NH4: reel.ES.NH4, NO3: reel.ES.NO3, Pt: reel.ES.Pt, MV_MES: 0.9 },
      reactifs: { methanol_kgj: reel.methanol_flux },
      results: [
        { key: 'ratio', label: 'Fraction du débit admise sur le filtre (réel)', unit: '-', value: reel.ratio_admis },
        { key: 'S_min', label: 'Surface de filtration minimale', unit: 'm²', value: forced.surface_filtration_min ?? st.surface_filtration_min },
        { key: 'nb', label: 'Cellules de filtration', unit: 'u', value: st.nb_cellules },
        { key: 'S', label: 'Surface réelle de filtration', unit: 'm²', value: st.surface_filtration },
        { key: 'NO3_visee', label: 'N-NO3 visé après mélange (réel)', unit: 'mg/L', value: reel.NO3_visee },
        { key: 'denit', label: 'Dénitrification (réel)', unit: 'kg N/j', value: reel.denitrification },
        { key: 'methanol', label: 'Méthanol pur (réel)', unit: 'kg/j', value: reel.methanol_flux },
        { key: 'ES_Q', label: 'Eaux sales (réel)', unit: 'm³/j', value: reel.ES_Q },
        { key: 'ES_MES', label: 'MES des eaux sales (réel)', unit: 'kg/j', value: reel.ES_MES },
      ],
      electricity: { total, fixed: electricite_lavage_air * k, detail: { lavage_air: electricite_lavage_air * k, eaux_sales: electricite_ES * k, dosage_methanol: electricite_methanol * k } },
      warnings,
    }
  },
  })
}

export default makeBiostyrPdn({
  id: 'biostyr-pdn',
  label: 'Biostyr post-dénitrification',
  short: 'Biostyr PDN',
  family: 'secondaire',
  vba: 'E9_Biostyr_PDN.cls',
  origine: 'II_biostyr_PDN',
  description: 'Biofiltre Biostyr en post-dénitrification au méthanol, placé derrière un étage nitrifiant. La fraction de débit admise est calculée pour tenir la garantie NGL ; le reste by-passe le filtre.',
})
