// ---------------------------------------------------------------------------
// Moteur de simulation : enchaîne les nœuds procédé dans l'ordre de la filière,
// en reproduisant la séquence d'OCEAN : dimensionnement (eau nominale) puis
// fonctionnement réel (eau réelle = nominale × NC_*) puis consommation élec.
//
// Chaque nœud est un objet :
//   {
//     id, label, family, vba, ported,
//     choices: [{ key, label, options: [{value,label}], default }],
//     params:  [{ key, label, unit, group, default: (ctx) => number }],
//     compute: (ctx) => { outNominal, outReel, sludge?, eauxSales?, results, electricity, warnings }
//   }
// ctx = { site, inNominal, inReel, choices, p, forced, upstream }
//   p[key]      = valeur effective (forcée si présente, sinon défaut) — "Ve" dans OCEAN
//   forced[key] = valeur forcée ou undefined                          — "Vr"
// ---------------------------------------------------------------------------
import { cloneStream, nominalStream, reelStream } from './stream.js'

export const FAMILIES = [
  { id: 'pretraitement', label: 'Prétraitement' },
  { id: 'primaire', label: 'Traitement primaire' },
  { id: 'secondaire', label: 'Traitement secondaire' },
  { id: 'tertiaire', label: 'Traitement tertiaire' },
  { id: 'desinfection', label: 'Désinfection' },
  { id: 'retours', label: 'Eaux sales / retours' },
]

export function defineNode(def) {
  return {
    ported: true,
    choices: [],
    params: [],
    ...def,
  }
}

/** valeurs de choix effectives (défaut si non renseigné) */
export function resolveChoices(node, instChoices = {}) {
  const out = {}
  for (const c of node.choices) out[c.key] = instChoices[c.key] ?? c.default
  return out
}

/** construit p (Ve) à partir des défauts (Vp) et des forçages (Vr) */
export function resolveParams(node, ctx0, forced = {}) {
  const p = {}
  const defaults = {}
  for (const prm of node.params) {
    let d
    try {
      d = typeof prm.default === 'function' ? prm.default({ ...ctx0, p }) : prm.default
    } catch {
      d = undefined
    }
    defaults[prm.key] = d
    p[prm.key] = forced[prm.key] != null && forced[prm.key] !== '' ? Number(forced[prm.key]) : d
  }
  return { p, defaults }
}

/**
 * Exécute la filière.
 * chain = [{ uid, nodeId, choices, forced }]
 * registry = { [nodeId]: nodeDef }
 */
export function runChain(chain, registry, site) {
  let inNominal = nominalStream(site)
  let inReel = reelStream(site)
  const upstream = { primaire: false, BA_forte: false, secondaire: null }
  const steps = []
  let elecTotal = 0

  for (const inst of chain) {
    const node = registry[inst.nodeId]
    if (!node) continue
    const choices = resolveChoices(node, inst.choices)
    const forced = {}
    for (const [k, v] of Object.entries(inst.forced || {})) if (v !== '' && v != null) forced[k] = Number(v)
    const ctx0 = { site, inNominal: cloneStream(inNominal), inReel: cloneStream(inReel), choices, forced, upstream: { ...upstream } }
    const { p, defaults } = resolveParams(node, ctx0, forced)
    const ctx = { ...ctx0, p, defaults }
    let out
    try {
      out = node.compute(ctx)
    } catch (e) {
      out = { outNominal: cloneStream(inNominal), outReel: cloneStream(inReel), results: [], electricity: { total: 0 }, warnings: [`Erreur de calcul : ${e.message}`] }
    }
    const step = {
      uid: inst.uid,
      nodeId: inst.nodeId,
      label: node.label,
      inNominal: ctx0.inNominal,
      inReel: ctx0.inReel,
      outNominal: out.outNominal,
      outReel: out.outReel,
      sludge: out.sludge,
      eauxSales: out.eauxSales,
      results: out.results || [],
      electricity: out.electricity || { total: 0 },
      warnings: out.warnings || [],
      p,
      defaults,
      choices,
      ported: node.ported,
    }
    steps.push(step)
    elecTotal += step.electricity.total || 0
    inNominal = out.outNominal
    inReel = out.outReel
    if (node.family === 'primaire') upstream.primaire = true
    if (node.id === 'ba-forte-charge') upstream.BA_forte = true
    if (node.family === 'secondaire') upstream.secondaire = node.id
  }
  return { steps, outNominal: inNominal, outReel: inReel, electricityTotal: elecTotal }
}
