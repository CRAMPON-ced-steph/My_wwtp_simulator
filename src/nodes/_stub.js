// Fabrique de nœud "à porter" : l'eau traverse sans modification, le nœud est
// signalé ported:false dans l'interface. Remplacer compute() par le port de la
// classe VBA correspondante.
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'

export function passthrough(def) {
  return defineNode({
    ...def,
    ported: false,
    compute(ctx) {
      return {
        outNominal: cloneStream(ctx.inNominal),
        outReel: cloneStream(ctx.inReel),
        results: [],
        electricity: { total: 0, fixed: 0, detail: {} },
        warnings: [`Calcul non porté : ${def.vba}. L'eau traverse ce nœud sans modification.`],
      }
    },
  })
}
