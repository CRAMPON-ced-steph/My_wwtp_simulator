// Port de F2_Biostyr_PDN_III.cls — Biostyr post-dénitrification en tertiaire.
// Classe quasi identique à E9 (vérifié par diff) : mêmes charges volumiques,
// même dosage de méthanol ; seule différence de calcul : les MES éliminées
// sont comptées sur la fraction by-passée (1 − ratio) — écart du VBA reproduit.
import { makeBiostyrPdn } from './biostyrPdn.js'

export default makeBiostyrPdn({
  id: 'biostyr-pdn-iii',
  label: 'Biostyr PDN III (tertiaire)',
  short: 'Biostyr PDN III',
  family: 'tertiaire',
  vba: 'F2_Biostyr_PDN_III.cls',
  origine: 'III_biostyr_PDN',
  mesElimineSurBypass: true,
  description: 'Biofiltre Biostyr en post-dénitrification au méthanol, en position tertiaire derrière un étage nitrifiant. La fraction de débit admise est calculée pour tenir la garantie NGL.',
})
