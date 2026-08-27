// Registre des nœuds procédé — ordre d'affichage dans la palette.
import dessablageDeshuilage from './dessablageDeshuilage.js'
import decantationSimple from './decantationSimple.js'
import decantationReactifs from './decantationReactifs.js'
import baForteCharge from './baForteCharge.js'
import baMoyenneCharge from './baMoyenneCharge.js'
import baFaibleCharge from './baFaibleCharge.js'
import aerationProlongee from './aerationProlongee.js'
import hybas from './hybas.js'
import mbbr from './mbbr.js'
import mbr from './mbr.js'
import biostyr from './biostyr.js'
import biostyrPdn from './biostyrPdn.js'
import biostyrNitrifiantIii from './biostyrNitrifiantIii.js'
import biostyrPdnIii from './biostyrPdnIii.js'
import decantationTertiaire from './decantationTertiaire.js'
import discfilter from './discfilter.js'
import filtrationSable from './filtrationSable.js'
import desinfectionUV from './desinfectionUV.js'
import chloration from './chloration.js'
import decantationEauxSales from './decantationEauxSales.js'

export const NODE_LIST = [
  dessablageDeshuilage,
  decantationSimple,
  decantationReactifs,
  baForteCharge,
  baMoyenneCharge,
  baFaibleCharge,
  aerationProlongee,
  hybas,
  mbbr,
  mbr,
  biostyr,
  biostyrPdn,
  biostyrNitrifiantIii,
  biostyrPdnIii,
  decantationTertiaire,
  discfilter,
  filtrationSable,
  desinfectionUV,
  chloration,
  decantationEauxSales,
]

export const REGISTRY = Object.fromEntries(NODE_LIST.map((n) => [n.id, n]))
