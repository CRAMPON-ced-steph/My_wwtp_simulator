// ---------------------------------------------------------------------------
// Moteur de la file boues.
//
// Équivalent de la seconde moitié de MOD_ProgrammePrincipal.prgm_principal :
// les boues extraites par les nœuds de la file eau alimentent l'étape
// « entrée », puis chaque procédé de la file boues lit une ou plusieurs étapes
// amont, écrit son étape aval et alimente le vecteur de retour en tête.
//
// Un nœud de la file boues est un objet :
//   {
//     id, label, family, vba, etapeSortie,
//     choices, params,
//     compute: (ctx) => { results, electricity, warnings, reactifs? }
//   }
// ctx = { site, table, soluble, retour, choices, p, forced, index }
//   table   : flux de boues, modifié en place (voir sludge.js)
//   soluble : pollution soluble, modifiée en place
//   retour  : vecteur de retour en tête, modifié en place
//   index   : rang de l'instance parmi les nœuds de même type (épaississeur 1, 2…)
// ---------------------------------------------------------------------------
import { makeSludgeTable, makeSolubleTable, makeRetours, RET_ORIGINE, chargerEntrees, ETAPE, RET, P, NB_TYPES, MES_etape, Q_etape } from './sludge.js'
import { resolveChoices, resolveParams } from './engine.js'
import { ratio } from './hypotheses.js'

export const SLUDGE_FAMILIES = [
  { id: 'epaississement', label: 'Épaississement' },
  { id: 'stabilisation', label: 'Stabilisation / digestion' },
  { id: 'deshydratation', label: 'Déshydratation' },
  { id: 'hygienisation', label: 'Hygiénisation' },
  { id: 'sechage', label: 'Séchage' },
  { id: 'valorisation', label: 'Valorisation / élimination' },
  { id: 'retours', label: 'Traitement des retours' },
]

export function defineSludgeNode(def) {
  return { ported: true, choices: [], params: [], ...def }
}

/**
 * Exécute la file boues.
 * apports = [{ type, origine, Q, MES, MV_MES, DCO, DBO, NK, Pt }] issus de la file eau
 * chain   = [{ uid, nodeId, choices, forced }]
 */
export function runSludgeChain(chain, registry, site, apports = []) {
  const table = chargerEntrees(makeSludgeTable(), apports)
  const soluble = makeSolubleTable()
  // `retours` porte un vecteur par origine ; `retour` reste l'alias du total,
  // pour que les procédés déjà portés continuent d'y écrire directement.
  const retours = makeRetours()
  const retour = retours[RET_ORIGINE.total]
  const steps = []
  let elecTotal = 0
  const compteur = {}

  for (const inst of chain) {
    const node = registry[inst.nodeId]
    if (!node) continue
    compteur[inst.nodeId] = (compteur[inst.nodeId] || 0) + 1
    const index = compteur[inst.nodeId]
    const choices = resolveChoices(node, inst.choices)
    const forced = {}
    for (const [k, v] of Object.entries(inst.forced || {})) if (v !== '' && v != null) forced[k] = Number(v)
    const ctx0 = { site, table, soluble, retour, retours, choices, forced, index }
    const { p, defaults } = resolveParams(node, ctx0, forced)
    const ctx = { ...ctx0, p, defaults }
    let out
    try {
      out = node.compute(ctx)
    } catch (e) {
      out = { results: [], electricity: { total: 0 }, warnings: [`Erreur de calcul : ${e.message}`] }
    }
    const etape = out.etapeSortie ?? (typeof node.etapeSortie === 'function' ? node.etapeSortie(index) : node.etapeSortie)
    steps.push({
      uid: inst.uid,
      nodeId: inst.nodeId,
      label: node.label + (node.multiple ? ` ${index}` : ''),
      index,
      etapeSortie: etape,
      MES: etape ? MES_etape(table, etape) : 0,
      Q: etape ? Q_etape(table, etape) : 0,
      results: out.results || [],
      electricity: out.electricity || { total: 0 },
      reactifs: out.reactifs || {},
      energie: out.energie || null,
      dechets: out.dechets || null,
      produits: out.produits || null,
      warnings: out.warnings || [],
      p,
      defaults,
      choices,
      ported: node.ported,
    })
    elecTotal += out.electricity?.total || 0
  }

  // bilan matière : ce qui reste disponible à chaque étape part à l'évacuation
  const evacuation = { Q: 0, MES: 0, MV_MES: 0, detail: [] }
  let MV = 0
  for (let e = 1; e <= 23; e++) {
    if (e === ETAPE.evacuees) continue
    for (let j = 1; j <= NB_TYPES; j++) {
      const reste = table[e][j][P.verif_flux]
      if (!(reste > 1e-9) || !(table[e][j][P.MES] > 0)) continue
      const MES = table[e][j][P.MES] * reste
      const Q = table[e][j][P.Q] * reste
      evacuation.MES += MES
      evacuation.Q += Q
      MV += MES * table[e][j][P.MV_MES]
      evacuation.detail.push({ etape: e, type: j, MES, Q, origine: table[e][j][P.origine] })
    }
  }
  evacuation.MV_MES = evacuation.MES > 0 ? MV / evacuation.MES : 0
  evacuation.siccite = evacuation.Q > 0 ? evacuation.MES / evacuation.Q : 0

  return {
    steps,
    table,
    soluble,
    retour,
    retours,
    evacuation,
    electricityTotal: elecTotal,
    retourResume: {
      Q: retour[RET.Q], MES: retour[RET.MES], DCO: retour[RET.DCO],
      DBO: retour[RET.DBO], NK: retour[RET.NK], NH4: retour[RET.NH4],
      NO3: retour[RET.NO3], Pt: retour[RET.Pt],
    },
  }
}

/**
 * Collecte les flux `sludge` produits par une exécution de la file eau, sous la
 * forme attendue par runSludgeChain.
 */
export function apportsDepuisFileEau(resultatFileEau, registry) {
  const out = []
  for (const s of resultatFileEau.steps) {
    if (!s.sludge || !(s.sludge.MES > 0)) continue
    const node = registry[s.nodeId]
    const b = s.sludge
    const MES = b.MES
    const MV = MES * (b.MV_MES || 0)
    // Les nœuds qui ne renseignent pas la composition de leurs boues la voient
    // reconstituée depuis la table de ratios d'AA_collection, comme dans le classeur VBA :
    // DCO, DBO et NK sont rapportés aux matières volatiles, Pt aux MES.
    out.push({
      type: typeDepuisNoeud(node, s.nodeId, b.origine),
      origine: b.origine,
      Q: b.Q,
      MES,
      MV_MES: b.MV_MES || 0,
      DCO: b.DCO ?? ratio(b.origine, 'DCO_MV') * MV,
      DBO: b.DBO ?? ratio(b.origine, 'DBO_MV') * MV,
      NK: b.NK ?? ratio(b.origine, 'NK_MV') * MV,
      Pt: b.Pt ?? ratio(b.origine, 'Pt_MES') * MES,
    })
  }
  return out
}
function typeDepuisNoeud(node, nodeId, origine) {
  // les graisses du dégraisseur ont leur type propre : c'est lui qui permet au
  // Biolix de les retrouver dans la file boues
  if (origine === 'graisses' || origine === 'codigestion_graisses') return 6
  const fam = node?.family
  if (fam === 'primaire') return 1
  if (fam === 'secondaire') return 2
  if (fam === 'tertiaire' || nodeId === 'decantation-eaux-sales') return 3
  return 2
}
