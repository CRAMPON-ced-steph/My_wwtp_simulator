// Port de E4_BA_aeration_prolongee.cls via la fabrique commune atvFaibleCharge.js
// (E4 est identique à E3 à trois constantes près : G_reference = 20 j,
//  sortie_NH4 = 1 mg/L, sortie_NO3 = 6 mg/L — vérifié par diff des deux classes)
import { makeATVFaibleCharge } from './atvFaibleCharge.js'

export default makeATVFaibleCharge({
  id: 'aeration-prolongee',
  label: 'Aération prolongée',
  short: 'Aération prolongée',
  vba: 'E4_BA_aeration_prolongee.cls',
  description: "Aération prolongée (âge de boues ≈ 20 j à 12 °C, ×1,08 si bio-P) : nitrification/dénitrification poussées, zones ou chenal, bio-P, méthanol et FeCl3, clarificateur.",
  G_reference: 20,
  sortie_NH4_def: 1,
  sortie_NO3_def: 6,
  origineEB: 'II_prolongee_EB',
  origineED: 'II_prolongee_ED',
})
