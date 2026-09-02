// ---------------------------------------------------------------------------
// Moteur des utilités.
//
// Les utilités ne transforment ni l'eau ni les boues : elles consomment de
// l'électricité, des réactifs ou de la chaleur, ou en produisent, à partir de
// grandeurs déjà calculées par les deux files. La désodorisation part du débit
// nominal et des sulfures strippés, le photovoltaïque d'une surface, la turbine
// d'un débit et d'une chute.
//
// Un nœud d'utilité est un objet :
//   { id, label, family, vba, choices, params,
//     compute: (ctx) => { results, electricity, reactifs?, energie?, warnings } }
// ctx = { site, contexte, choices, p, forced }
//   contexte : grandeurs issues des deux files (débit nominal, sulfures
//              strippés, biogaz produit, besoins thermiques…), rassemblées par
//              `contexteDepuisFilieres()`.
// ---------------------------------------------------------------------------
import { resolveChoices, resolveParams } from './engine.js'
import { fluxDepuisFilieres } from '../nodes-utilites/energieThermique.js'

export const UTILITY_FAMILIES = [
  { id: 'desodorisation', label: 'Désodorisation' },
  { id: 'batiments', label: 'Bâtiments' },
  { id: 'production', label: "Production d'énergie" },
  { id: 'reactifs', label: 'Gestion des réactifs' },
]

export function defineUtilityNode(def) {
  return { ported: true, choices: [], params: [], ...def }
}

/**
 * Rassemble ce dont les utilités ont besoin à partir des résultats des deux
 * files, sans que chaque nœud ait à connaître leur structure interne.
 */
export function contexteDepuisFilieres(site, eau = null, boues = null) {
  const ctx = {
    Q_nominal: site.Q_nominal,
    Q_reel: eau?.outReel?.Q ?? site.Q_nominal,
    HS_strippe_kgj: 0,
    biogaz_Nm3j: 0,
    biogaz_CH4: 0,
    besoin_thermique_kWhj: 0,
    chaleur_recuperable_kWhj: 0,
    electricite_eau_kWhj: eau?.electricityTotal ?? 0,
    electricite_boues_kWhj: boues?.electricityTotal ?? 0,
    reactifs: {},
    // agrégats consommés par le bloc transverse
    DCO_eliminee: 0,
    DBO_eliminee: 0,
    N_elimine: 0,
    P_elimine: 0,
    boues_evacuees_Q: boues?.evacuation?.Q ?? 0,
    boues_evacuees_MES: boues?.evacuation?.MES ?? 0,
    gaz_naturel_kWhPCIj: 0,
    fioul_kWhPCIj: 0,
    cendres_Tj: 0,
    REFIB_Tj: 0,
    graisses_Tj: 0,
    struvite_kgj: 0,
    fluxThermiques: [],
  }
  // sous-produits valorisables ou à évacuer, issus de la file boues
  for (const s of boues?.steps ?? []) {
    ctx.cendres_Tj += (s.dechets?.cendres_kgj ?? 0) / 1000
    ctx.REFIB_Tj += (s.dechets?.REFIB_kgj ?? 0) / 1000
    ctx.struvite_kgj += s.produits?.struvite_kgj ?? 0
  }
  // pollution éliminée par la file eau, entrée moins sortie
  const entree = eau?.steps?.[0]?.inReel
  const sortie = eau?.outReel
  if (entree && sortie) {
    ctx.DCO_eliminee = Math.max(0, entree.DCO - sortie.DCO)
    ctx.DBO_eliminee = Math.max(0, entree.DBO - sortie.DBO)
    ctx.N_elimine = Math.max(0, entree.NK + entree.NO3 - sortie.NK - sortie.NO3)
    ctx.P_elimine = Math.max(0, entree.Pt - sortie.Pt)
  }
  // combustibles consommés par la file boues (séchage, incinération, Athos)
  for (const s of boues?.steps ?? []) {
    const src = s.energie?.source
    const kWh = s.energie?.besoin_thermique_kWhj ?? 0
    if (src === 'gaz_naturel' || src === 'combustible') ctx.gaz_naturel_kWhPCIj += kWh
    else if (src === 'fioul') ctx.fioul_kWhPCIj += kWh
  }
  // Sulfures strippés : ce sont eux qui chargent l'air vicié en H2S. On prend
  // la charge entrante du premier ouvrage de la file eau, avant abattement.
  ctx.HS_strippe_kgj = eau?.steps?.[0]?.inReel?.Sh ?? eau?.inReel?.Sh ?? 0

  for (const s of boues?.steps ?? []) {
    if (s.energie?.biogaz_Nm3j) {
      ctx.biogaz_Nm3j += s.energie.biogaz_Nm3j
      ctx.biogaz_CH4 = s.energie.biogaz_CH4 ?? ctx.biogaz_CH4
    }
    ctx.besoin_thermique_kWhj += s.energie?.besoin_thermique_kWhj ?? 0
    ctx.chaleur_recuperable_kWhj += s.energie?.recuperable_kWhj ?? 0
  }
  // flux thermiques déclarés par les procédés, pour l'analyse pincement
  ctx.fluxThermiques = fluxDepuisFilieres(boues, null)
  // cumul des réactifs consommés par les deux files
  for (const s of [...(eau?.steps ?? []), ...(boues?.steps ?? [])]) {
    for (const [k, v] of Object.entries(s.reactifs ?? {})) {
      if (v > 0) ctx.reactifs[k] = (ctx.reactifs[k] || 0) + v
    }
  }
  return ctx
}

/** Exécute une liste de nœuds d'utilités. */
export function runUtilities(chain, registry, site, contexte, options = {}) {
  const steps = []
  let elecTotal = 0
  let elecProduite = 0
  for (const inst of chain) {
    const node = registry[inst.nodeId]
    if (!node) continue
    const choices = resolveChoices(node, inst.choices)
    const forced = {}
    for (const [k, v] of Object.entries(inst.forced || {})) if (v !== '' && v != null) forced[k] = Number(v)
    const ctx0 = { site, contexte, choices, forced }
    const { p, defaults } = resolveParams(node, ctx0, forced)
    let out
    try {
      out = node.compute({ ...ctx0, p, defaults })
    } catch (e) {
      out = { results: [], electricity: { total: 0 }, warnings: [`Erreur de calcul : ${e.message}`] }
    }
    const step = {
      uid: inst.uid,
      nodeId: inst.nodeId,
      label: node.label,
      results: out.results || [],
      electricity: out.electricity || { total: 0 },
      reactifs: out.reactifs || {},
      energie: out.energie || null,
      sulfures: out.sulfures || null,
      pincement: out.pincement || null,
      bilan: out.bilan || null,
      co2: out.co2 || null,
      opex: out.opex || null,
      roi: out.roi || null,
      warnings: out.warnings || [],
      p, defaults, choices,
    }
    steps.push(step)
    // permet au bloc transverse de chaîner bilan électrique puis empreinte CO2
    options.apresChaqueEtape?.(step)
    // le traitement des sulfures fixe la charge strippée que la désodorisation
    // placée en aval retrouvera dans son air vicié
    if (out.sulfures) contexte.HS_strippe_kgj = out.sulfures.HS_strippe
    // un bâtiment chauffé ajoute son besoin aux flux soumis au pincement
    if (out.energie?.besoin_thermique_kWhj > 0 && node.family === 'batiments') {
      contexte.fluxThermiques = [
        ...contexte.fluxThermiques,
        ...fluxDepuisFilieres(null, { steps: [{ nodeId: node.id, label: node.label, energie: out.energie }] }),
      ]
    }
    // les réactifs consommés par une utilité alimentent le contexte, ce qui
    // permet à la gestion des réactifs placée en aval de les consolider
    for (const [k, v] of Object.entries(out.reactifs ?? {})) {
      if (v > 0 && node.family !== 'reactifs') contexte.reactifs[k] = (contexte.reactifs[k] || 0) + v
    }
    // Une production est portée comme un total négatif ; le champ `produite`
    // n'est qu'un doublon explicite, à ne pas cumuler par-dessus.
    const t = out.electricity?.total || 0
    if (t >= 0) {
      elecTotal += t
      elecProduite += out.electricity?.produite || 0
    } else {
      elecProduite -= t
    }
  }
  return { steps, electricityTotal: elecTotal, electricityProduite: elecProduite }
}
