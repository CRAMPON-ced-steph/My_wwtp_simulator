// Registre des nœuds de la file boues.
import epaississement from './epaississement.js'
import digestion from './digestion.js'
import deshydratation from './deshydratation.js'
import chaulage from './chaulage.js'
import { sechageThermique, sechageBioco, sechageInos } from './sechage.js'
import incineration from './incineration.js'
import athos from './athos.js'
import biothelys from './biothelys.js'
import exelys from './exelys.js'
import anitaMox from './anitaMox.js'
import anitaShunt from './anitaShunt.js'
import biolix from './biolix.js'
import retoursMap from './retoursMap.js'

export const SLUDGE_NODE_LIST = [
  epaississement,
  digestion,
  deshydratation,
  chaulage,
  sechageThermique,
  sechageBioco,
  sechageInos,
  incineration,
  athos,
  biothelys,
  exelys,
  anitaMox,
  anitaShunt,
  biolix,
  retoursMap,
]

export const SLUDGE_REGISTRY = Object.fromEntries(SLUDGE_NODE_LIST.map((n) => [n.id, n]))
