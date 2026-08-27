// Port de G2_Desinfection_Cl.cls — chloration (Cl2 gazeux ou hypochlorite)
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'

export default defineNode({
  id: 'chloration',
  label: 'Chloration',
  short: 'Chloration',
  family: 'desinfection',
  vba: 'G2_Desinfection_Cl.cls',
  description: "Consommation de chlore pur : dose fixe (rejet carbone) ou point de rupture (5 mg Cl2/mg N-NH4 + 5 mg/mg N-NO2 + résiduel). Pas d'électricité, pas d'effet sur les charges.",
  choices: [
    { key: 'reactif', label: 'Réactif', default: 'HClO', options: [{ value: 'Cl2', label: 'chlore gazeux' }, { value: 'HClO', label: 'hypochlorite (HClO)' }] },
  ],
  params: [
    { key: 'Q_ratio_admis', label: "Part du débit admise sur l'ouvrage", unit: '-', group: 'Hydraulique', default: 1 },
    { key: 'C_dosage_mgL', label: 'Dose (rejet carbone seul)', unit: 'mg Cl2/L', group: 'Dosage', default: 10 },
    { key: 'CN_dosage_mg_mgNH4', label: 'Dose sur N-NH4 (rejet C+N)', unit: 'mg Cl2/mg N', group: 'Dosage', default: 5 },
    { key: 'CN_NO2_mgL', label: 'N-NO2 dans l\'eau', unit: 'mg/L', group: 'Dosage', default: 0 },
    { key: 'CN_dosage_mg_mgNO2', label: 'Dose sur N-NO2', unit: 'mg Cl2/mg N', group: 'Dosage', default: 5 },
    { key: 'CN_residuel_mgL', label: 'Chlore résiduel', unit: 'mg/L', group: 'Dosage', default: 1 },
  ],
  compute(ctx) {
    const { p, site, choices } = ctx
    const calc = (s) => {
      let conso
      let dose_resultante = null
      if (site.qualite_rejet === 'C') {
        conso = (p.C_dosage_mgL * p.Q_ratio_admis * s.Q) / 1000
      } else {
        conso = p.CN_dosage_mg_mgNH4 * p.Q_ratio_admis * s.NH4
        conso += (p.CN_dosage_mg_mgNO2 * p.CN_NO2_mgL * p.Q_ratio_admis * s.Q) / 1000
        conso += (p.CN_residuel_mgL * p.Q_ratio_admis * s.Q) / 1000
        dose_resultante = (conso * 1000) / (p.Q_ratio_admis * s.Q)
      }
      return { conso, dose_resultante }
    }
    const nom = calc(ctx.inNominal)
    const reel = calc(ctx.inReel)
    return {
      outNominal: cloneStream(ctx.inNominal),
      outReel: cloneStream(ctx.inReel),
      reactifs: { [choices.reactif === 'Cl2' ? 'Cl2_gazeux_pur_kgj' : 'Cl2_HClO_pur_kgj']: reel.conso },
      results: [
        { key: 'NH4_in', label: 'N-NH4 entrée (réel)', unit: 'mg/L', value: conc(ctx.inReel, 'NH4') },
        { key: 'conso_nom', label: 'Chlore pur (nominal)', unit: 'kg/j', value: nom.conso },
        { key: 'conso_reel', label: 'Chlore pur (réel)', unit: 'kg/j', value: reel.conso },
        ...(reel.dose_resultante != null ? [{ key: 'dose', label: 'Dose résultante (réel)', unit: 'mg/L', value: reel.dose_resultante }] : []),
      ],
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings: [],
    }
  },
})
