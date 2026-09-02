// ---------------------------------------------------------------------------
// Moteur de simulation : enchaîne les nœuds procédé dans l'ordre de la filière,
// en reproduisant la séquence du classeur VBA : dimensionnement (eau nominale) puis
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
//   p[key]      = valeur effective (forcée si présente, sinon défaut) — "Ve" dans le classeur VBA
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
 *
 * Rebouclage des eaux sales (boucle iteration_ES du classeur) : les nœuds qui
 * produisent un flux `eauxSales` (biofiltres, Discfilter, filtre sable) le
 * voient renvoyé en tête de filière. Si un nœud "decantation-eaux-sales" (H1)
 * est présent, les eaux sales y sont d'abord traitées et c'est son surnageant
 * qui retourne en tête (les boues partent en file boues). Le calcul itère
 * jusqu'à stabilisation du débit de retour (max 4 itérations).
 */
const ES_KEYS = ['Q', 'DCO', 'DBO', 'MES', 'NK', 'NH4', 'NO3', 'Pt']
const zeroES = () => ({ Q: 0, DCO: 0, DBO: 0, MES: 0, NK: 0, NH4: 0, NO3: 0, Pt: 0 })
const addES = (a, b) => { for (const k of ES_KEYS) a[k] += b[k] || 0 }

function runChainOnce(chain, registry, site, retourNominal, retourReel) {
  let inNominal = nominalStream(site)
  let inReel = reelStream(site)
  for (const k of ES_KEYS) {
    inNominal[k] += retourNominal[k]
    inReel[k] += retourReel[k]
  }
  const upstream = { primaire: false, primaire_reactif: false, BA_forte: false, secondaire: null }
  const steps = []
  let elecTotal = 0
  const esNominal = []
  const esReel = []

  for (const inst of chain) {
    const node = registry[inst.nodeId]
    if (!node) continue
    const choices = resolveChoices(node, inst.choices)
    const forced = {}
    for (const [k, v] of Object.entries(inst.forced || {})) if (v !== '' && v != null) forced[k] = Number(v)
    const ctx0 = { site, inNominal: cloneStream(inNominal), inReel: cloneStream(inReel), choices, forced, upstream: { ...upstream }, esAmontNominal: esNominal.slice(), esAmontReel: esReel.slice() }
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
      retourTraite: out.retourTraite,
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
    if (out.eauxSales) {
      esReel.push(out.eauxSales)
      esNominal.push(out.eauxSalesNominal || out.eauxSales)
    }
    if (node.family === 'primaire') upstream.primaire = true
    if (node.id === 'decantation-reactifs') upstream.primaire_reactif = true
    if (node.id === 'ba-forte-charge') upstream.BA_forte = true
    if (node.family === 'secondaire') upstream.secondaire = node.id
  }
  return { steps, outNominal: inNominal, outReel: inReel, electricityTotal: elecTotal, esNominal, esReel }
}

export function runChain(chain, registry, site) {
  const hasH1 = chain.some((i) => i.nodeId === 'decantation-eaux-sales')
  let retourNominal = zeroES()
  let retourReel = zeroES()
  let r = null
  const MAX_ITER = 4
  for (let iter = 1; iter <= MAX_ITER; iter++) {
    r = runChainOnce(chain, registry, site, retourNominal, retourReel)
    // flux de retour de la prochaine itération
    const nextNominal = zeroES()
    const nextReel = zeroES()
    if (hasH1) {
      let consommees = 0
      for (const s of r.steps) {
        if (s.nodeId === 'decantation-eaux-sales' && s.retourTraite) {
          addES(nextNominal, s.retourTraite.nominal)
          addES(nextReel, s.retourTraite.reel)
          consommees = Math.max(consommees, s.retourTraite.esConsommees || 0)
        }
      }
      // eaux sales émises en aval du décanteur d'eaux sales : retour direct
      for (let i = consommees; i < r.esNominal.length; i++) addES(nextNominal, r.esNominal[i])
      for (let i = consommees; i < r.esReel.length; i++) addES(nextReel, r.esReel[i])
    } else {
      for (const es of r.esNominal) addES(nextNominal, es)
      for (const es of r.esReel) addES(nextReel, es)
    }
    const converged = Math.abs(nextReel.Q - retourReel.Q) < Math.max(1, 0.005 * (site.Q_nominal || 1)) && Math.abs(nextNominal.Q - retourNominal.Q) < Math.max(1, 0.005 * (site.Q_nominal || 1))
    retourNominal = nextNominal
    retourReel = nextReel
    if (converged) break
  }
  r.retourNominal = retourNominal
  r.retourReel = retourReel
  return r
}
