// Registre des nœuds transverses.
import { bilanElectrique, empreinteCO2 } from './bilan.js'
import { gestionOpex, retourInvestissement } from './economie.js'

export const TRANSVERSE_NODE_LIST = [bilanElectrique, empreinteCO2, gestionOpex, retourInvestissement]
export const TRANSVERSE_REGISTRY = Object.fromEntries(TRANSVERSE_NODE_LIST.map((n) => [n.id, n]))
