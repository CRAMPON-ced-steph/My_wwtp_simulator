// MBR — ABSENT du classeur 140822_OCEAN_CCR.xlsm : aucune classe VBA dédiée
// (seul un jeu de ratios de boues "II_MBR — A REVOIR" existe dans AA_collection.ratio).
// Squelette à compléter à partir de la classe BA faible charge (E3) + membranes.
import { passthrough } from './_stub.js'

export default passthrough({
  id: 'mbr',
  label: 'MBR (bioréacteur à membranes)',
  short: 'MBR',
  family: 'secondaire',
  vba: '— (non modélisé dans OCEAN)',
  description: "Non modélisé dans le classeur source. À dériver de la boue activée faible charge : clarificateur remplacé par des membranes (flux, aération de décolmatage, perméat).",
  params: [
    { key: 'flux_membranes_L_m2h', label: 'Flux net de filtration', unit: 'L/(m²·h)', group: 'Membranes', default: 20 },
    { key: 'MES_bassin', label: 'MES dans le bassin', unit: 'g/L', group: 'Membranes', default: 8 },
    { key: 'SAD_m', label: "Air de décolmatage spécifique", unit: 'Nm³/(m²·h)', group: 'Membranes', default: 0.3 },
  ],
})
