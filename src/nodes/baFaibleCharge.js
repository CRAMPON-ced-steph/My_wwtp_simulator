// Port de E3_BA_faible_charge.cls via la fabrique commune atvFaibleCharge.js
import { makeATVFaibleCharge } from './atvFaibleCharge.js'

export default makeATVFaibleCharge({
  id: 'ba-faible-charge',
  label: 'Boue activée faible charge',
  short: 'BA faible charge',
  vba: 'E3_BA_faible_charge.cls',
  description: "Boue activée faible charge (âge de boues ≈ 14 j à 12 °C, ×1,08 si bio-P) : nitrification/dénitrification, zones anaérobie/anoxie/aérobie ou chenal, bio-P, méthanol et FeCl3, clarificateur.",
  G_reference: 14,
  sortie_NH4_def: 3,
  sortie_NO3_def: 5,
  origineEB: 'II_faible_EB',
  origineED: 'II_faible_ED',
})
