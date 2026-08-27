// Port de C2_Dessablage_Deshuilage.cls (Sub DessablageDeshuilage)
// Le procédé ne modifie pas la ligne d'eau (pas de rendement d'élimination
// dans OCEAN) : il produit un flux interne de graisses (Mat_graisse_int) et
// une consommation électrique (râcleurs + aération).
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'
import { CONST, HYP, ratio } from '../core/hypotheses.js'

const H = {
  graisse_MS_EH_an: 0.83, // kg MS / EH / an
  ratio_graisse_DBO_DCO: HYP.dessabl_graisse_DBO_DCO, // 0.9
  design_tps_sejour_min: 20, // min au débit nominal moyen
  PuissanceRacleurParOuvrage_kW: 0.65, // entre 0,3 et 1 selon DWA
}

export default defineNode({
  id: 'dessablage-deshuilage',
  label: 'Dessablage – déshuilage',
  short: 'Dessableur',
  family: 'pretraitement',
  vba: 'C2_Dessablage_Deshuilage.cls',
  description: 'Ouvrage combiné aéré. Production de graisses (0,83 kg MS/EH/an + 50 % des MES de vidanges), électricité râcleurs et surpresseurs.',
  params: [
    { key: 'volume_total_m3', label: 'Volume total des bassins', unit: 'm³', group: 'Ouvrage', default: (c) => (H.design_tps_sejour_min / (CONST.NOMBRE_MINUTE_PAR_HEURE * CONST.NOMBRE_HEURE_PAR_JOUR)) * c.site.Q_nominal },
    { key: 'nb_ouvrages', label: "Nombre d'ouvrages", unit: '-', group: 'Ouvrage', default: 2 },
    { key: 'Qair_spec', label: "Débit d'air spécifique", unit: '(Nm³/h)/m³', group: 'Aération', default: 0.5, hint: 'DWA A216 : 0,5 à 1,3' },
    { key: 'tps_fct_air', label: "Temps de fonctionnement de l'aération", unit: 'h/j', group: 'Aération', default: 24 },
    { key: 'P_refoulement', label: 'Pression de refoulement air', unit: 'mCE', group: 'Aération', default: 4 },
    { key: 'conso_spec_surpresseur', label: 'Conso spécifique surpresseurs', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: 4.5, hint: 'surpresseurs à lobes' },
    { key: 'graisse_concentration', label: 'Concentration MS des graisses', unit: 'g/L', group: 'Graisses', default: 92, hint: 'sans concentrateur' },
    { key: 'graisse_Q', label: 'Débit de graisses', unit: 'm³/j', group: 'Graisses', default: (c) => graisseMS(c) / c.p.graisse_concentration },
    { key: 'graisse_MV_MES', label: 'MV/MES des graisses', unit: '-', group: 'Graisses', default: 0.9 },
    { key: 'graisse_DCO', label: 'DCO des graisses', unit: 'kg/j', group: 'Graisses', default: (c) => ratio('codigestion_graisses', 'DCO_MV') * c.p.graisse_MV_MES * c.p.graisse_concentration * c.p.graisse_Q },
  ],
  compute(ctx) {
    const { p, site } = ctx
    const vidange_MES = (site.vidange_Q_nominal * site.vidange_MES_mgL_nominal) / 1000
    // graisses
    const graisse_MS = p.graisse_concentration * p.graisse_Q
    const graisse_DCO = p.graisse_DCO
    const graisse_DBO = graisse_DCO * H.ratio_graisse_DBO_DCO
    const graisses = {
      MS_kgj: graisse_MS,
      Q_m3j: p.graisse_Q,
      MV_MES: p.graisse_MV_MES,
      DCO_MES: graisse_MS > 0 ? graisse_DCO / graisse_MS : 0,
      DBO_MES: graisse_MS > 0 ? graisse_DBO / graisse_MS : 0,
      NK_MES: ratio('codigestion_graisses', 'NK_MV') * p.graisse_MV_MES,
      Pt_MES: ratio('codigestion_graisses', 'Pt_MES'),
    }
    // électricité
    const electricite_racleur = H.PuissanceRacleurParOuvrage_kW * p.nb_ouvrages * CONST.NOMBRE_HEURE_PAR_JOUR
    const electricite_aeration = (p.conso_spec_surpresseur / 1000) * (p.Qair_spec * p.volume_total_m3 * p.tps_fct_air) * p.P_refoulement
    const warnings = []
    if (vidange_MES > 0) warnings.push('50 % des MES de vidanges sont comptées dans les graisses (hypothèse OCEAN).')
    return {
      outNominal: cloneStream(ctx.inNominal),
      outReel: cloneStream(ctx.inReel),
      sludge: { origine: 'graisses', Q: graisses.Q_m3j, MES: graisses.MS_kgj, concentration: p.graisse_concentration, MV_MES: graisses.MV_MES, DCO: graisse_DCO, DBO: graisse_DBO },
      results: [
        { key: 'ts', label: 'Temps de séjour au nominal', unit: 'min', value: (p.volume_total_m3 / site.Q_nominal) * 1440 },
        { key: 'graisse_MS', label: 'Graisses produites', unit: 'kg MS/j', value: graisse_MS },
        { key: 'graisse_Q', label: 'Débit de graisses', unit: 'm³/j', value: p.graisse_Q },
        { key: 'graisse_DCO', label: 'DCO des graisses', unit: 'kg/j', value: graisse_DCO },
        { key: 'Qair', label: "Débit d'air", unit: 'Nm³/h', value: p.Qair_spec * p.volume_total_m3 },
      ],
      electricity: { total: electricite_racleur + electricite_aeration, fixed: electricite_racleur, detail: { racleurs: electricite_racleur, aeration: electricite_aeration } },
      warnings,
    }
  },
})

function graisseMS(c) {
  const vidange_MES = (c.site.vidange_Q_nominal * c.site.vidange_MES_mgL_nominal) / 1000
  return (H.graisse_MS_EH_an / CONST.NOMBRE_JOUR_PAR_AN) * c.site.Eq_hab + 0.5 * vidange_MES
}
