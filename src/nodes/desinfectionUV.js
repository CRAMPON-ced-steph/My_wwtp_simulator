// Port de G1_Desinfection_UV.cls — désinfection UV (basse / moyenne pression)
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'

const H = {
  factor_reel: 0.6, // régulation d'intensité possible entre 60 % et 100 %
  temps_operation_jour: 24, // h
  BP: { dose_elec: 30, puissance_par_banc: 70 }, // Wh/m³ ; kW (10 modules × 9 paires × 400 W)
  MP: { dose_elec: 70, puissance_par_banc: 40 }, // Wh/m³ ; kW (6 lampes de 6 kW)
}
const tech = (c) => (c.choices.technologie === 'MP' ? H.MP : H.BP)

export default defineNode({
  id: 'desinfection-uv',
  label: 'Désinfection UV',
  short: 'UV',
  family: 'desinfection',
  vba: 'G1_Desinfection_UV.cls',
  description: "Dose électrique (Wh/m³) × débit horaire de pointe → puissance et nombre de bancs. Pas d'effet sur les charges. Consommation sur le débit réel.",
  choices: [
    { key: 'technologie', label: 'Technologie de lampes', default: 'BP', options: [{ value: 'BP', label: 'basse pression' }, { value: 'MP', label: 'moyenne pression' }] },
    { key: 'lampes', label: 'Adaptation de la puissance des lampes', default: 'automatique', options: [{ value: 'manuel', label: 'manuelle' }, { value: 'automatique', label: 'automatique' }] },
  ],
  params: [
    { key: 'dose_elec', label: 'Dose électrique', unit: 'Wh/m³', group: 'Dimensionnement', default: (c) => tech(c).dose_elec, hint: 'dépend de la qualité d\'eau (J. Pannejon)' },
    { key: 'debit_nominal', label: 'Débit nominal de dimensionnement', unit: 'm³/h', group: 'Dimensionnement', default: (c) => (c.site.Q_nominal * c.site.pointe_TP + (c.inNominal.Q - c.site.Q_nominal)) / H.temps_operation_jour },
    { key: 'nombre_bancs_prevu', label: 'Nombre de bancs prévus', unit: 'u', group: 'Dimensionnement', default: undefined, hint: 'calculé si non forcé' },
    { key: 'debit_reel_jour', label: 'Débit réel traité', unit: 'm³/j', group: 'Réel', default: (c) => c.site.Q_nominal * c.site.NC_Q + (c.inNominal.Q - c.site.Q_nominal) },
    { key: 'nombre_bancs_reel', label: 'Nombre de bancs en service', unit: 'u', group: 'Réel', default: undefined, hint: 'calculé si non forcé' },
  ],
  compute(ctx) {
    const { p, forced, choices } = ctx
    const warnings = []
    const T = tech(ctx)
    const puissance_nominal = (p.dose_elec / 1000) * p.debit_nominal
    let puissance_par_banc = T.puissance_par_banc
    let nombre_bancs_prevu
    if (forced.nombre_bancs_prevu != null) {
      nombre_bancs_prevu = forced.nombre_bancs_prevu
      puissance_par_banc = puissance_nominal / nombre_bancs_prevu
    } else {
      nombre_bancs_prevu = Math.ceil(puissance_nominal / puissance_par_banc - 1e-12)
    }
    // réel
    const debit_reel_heure = p.debit_reel_jour / H.temps_operation_jour
    let nombre_bancs_reel, puissance_ratio
    const puissance_elec_reel = (p.dose_elec / 1000) * debit_reel_heure
    if (forced.nombre_bancs_reel != null) {
      nombre_bancs_reel = forced.nombre_bancs_reel
      puissance_ratio = nombre_bancs_reel
    } else {
      puissance_ratio = puissance_elec_reel / puissance_par_banc
      nombre_bancs_reel = Math.ceil(puissance_ratio - 1e-12)
    }
    if (nombre_bancs_reel > nombre_bancs_prevu) warnings.push('Le nombre de bancs en fonctionnement réel dépasse le nombre de bancs prévus.')
    // électricité
    let electricite
    if (choices.lampes === 'manuel') {
      electricite = nombre_bancs_reel * H.temps_operation_jour * puissance_par_banc
    } else if (nombre_bancs_reel === 1 && puissance_ratio < H.factor_reel) {
      electricite = H.factor_reel * nombre_bancs_reel * H.temps_operation_jour * puissance_par_banc
    } else {
      electricite = (p.dose_elec / 1000) * p.debit_reel_jour // = dose × débit horaire × 24 h
    }
    return {
      outNominal: cloneStream(ctx.inNominal),
      outReel: cloneStream(ctx.inReel),
      results: [
        { key: 'P_nom', label: 'Puissance nominale', unit: 'kW', value: puissance_nominal },
        { key: 'P_banc', label: 'Puissance par banc', unit: 'kW', value: puissance_par_banc },
        { key: 'nb_prevu', label: 'Bancs prévus', unit: 'u', value: nombre_bancs_prevu },
        { key: 'nb_reel', label: 'Bancs en service (réel)', unit: 'u', value: nombre_bancs_reel },
        { key: 'P_reel', label: 'Puissance appelée (réel)', unit: 'kW', value: puissance_elec_reel },
      ],
      electricity: { total: electricite, fixed: 0, detail: { lampes: electricite } },
      warnings,
    }
  },
})
