// Registre des nœuds d'utilités.
import { desodorisationBiologique, desodorisationChimique } from './desodorisation.js'
import { turbineHydraulique, gestionReactifs } from './turbineReactifs.js'
import { hvacAdmin, hvacExploitation, hvacElectrique } from './hvac.js'
import { photovoltaique } from './photovoltaique.js'
import { traitementSulfures } from './sulfures.js'
import { energieThermique } from './energieThermique.js'

export const UTILITY_NODE_LIST = [
  traitementSulfures,
  desodorisationBiologique,
  desodorisationChimique,
  hvacAdmin,
  hvacExploitation,
  hvacElectrique,
  photovoltaique,
  energieThermique,
  turbineHydraulique,
  gestionReactifs,
]

export const UTILITY_REGISTRY = Object.fromEntries(UTILITY_NODE_LIST.map((n) => [n.id, n]))
