// MBR — EXTENSION (pas un port) : aucune classe VBA dédiée dans le classeur
// 140822_OCEAN_CCR.xlsm (seul un jeu de ratios de boues "II_MBR — A REVOIR"
// existe dans AA_collection.ratio). Ce nœud dérive du cœur biologique
// E3/E4 (fabrique atvFaibleCharge.js) avec les adaptations classiques :
//  - clarificateur remplacé par des membranes immergées (surface = pointe / flux)
//  - MES bassin par défaut 8 g/L, MES sortie 1 mg/L
//  - boues extraites à la concentration du bassin (pas de recirculation clarif)
//  - électricité : air de décolmatage (SADm × surface) + pompes de perméat,
//    en plus de l'aération biologique, l'agitation et la recirculation interne
import { makeATVFaibleCharge } from './atvFaibleCharge.js'

export default makeATVFaibleCharge({
  id: 'mbr',
  label: 'MBR (bioréacteur à membranes)',
  short: 'MBR',
  vba: '— dérivé de E3 (extension, hors OCEAN)',
  extension: true,
  description: "EXTENSION hors classeur : cœur biologique de la BA faible charge (nitrification/dénitrification, bio-P, méthanol, FeCl3) avec séparation membranaire au lieu du clarificateur. Boues de ratios II_MBR ('A REVOIR' dans OCEAN).",
  G_reference: 14,
  sortie_NH4_def: 3,
  sortie_NO3_def: 5,
  origineEB: 'II_MBR',
  origineED: 'II_MBR',
  membrane: { MES_bassin_def: 8, flux_def: 40, SADm_def: 0.3 },
})
