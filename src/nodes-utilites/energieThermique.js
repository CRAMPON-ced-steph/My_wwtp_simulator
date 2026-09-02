// ---------------------------------------------------------------------------
// Port de z_Gestion_energie_thermique.cls et z_Gestion_energie_therm_PINCH.cls.
//
// Le nœud collecte les flux thermiques exposés par les procédés des trois
// files, les soumet au solveur d'analyse pincement, et en tire les cibles
// énergétiques ainsi que le placement des utilités.
//
// Chaque procédé porté expose déjà un objet `energie` : la digestion y déclare
// son besoin de maintien en température et son biogaz, les séchages leur
// consommation et la chaleur récupérable au refroidissement des buées,
// l'incinération son appoint et sa chaleur disponible sur les fumées, l'Athos
// son RTO, le HVAC son chauffage. Ces déclarations sont converties en flux
// chauds et froids avec leurs niveaux de température.
//
// Les niveaux de température retenus par défaut sont ceux du classeur, où ils
// sont écrits en dur dans les routines `transfert_flux_energie_vers_pinch` de
// chaque classe. Ils restent modifiables.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { analyserPincement } from '../core/pinch.js'
import { CONST } from '../core/hypotheses.js'

// Niveaux de température des flux thermiques, par procédé. `T_in` et `T_out`
// décrivent le fluide caloporteur, pas le procédé lui-même.
const NIVEAUX = {
  digestion: { froid: { T_in: 35, T_out: 55, dT: 3, h: 0.6, label: 'Maintien en température du digesteur' } },
  biothelys: {
    froid: { T_in: 90, T_out: 165, dT: 5, h: 0.8, label: 'Vapeur de lyse Biothelys' },
    chaud: { T_in: 165, T_out: 120, dT: 5, h: 0.8, label: 'Refroidissement des boues lysées' },
  },
  exelys: {
    froid: { T_in: 95, T_out: 165, dT: 5, h: 0.8, label: 'Vapeur de lyse Exelys' },
    chaud: { T_in: 165, T_out: 55, dT: 5, h: 0.8, label: 'Refroidissement des boues lysées' },
  },
  sechage: {
    froid: { T_in: 110, T_out: 150, dT: 8, h: 0.3, label: 'Séchage — apport de chaleur' },
    chaud: { T_in: 110, T_out: 40, dT: 8, h: 0.15, label: 'Condensation des buées' },
  },
  incineration: {
    chaud: { T_in: 890, T_out: 275, dT: 15, h: 0.1, label: 'Refroidissement des fumées' },
    froid: { T_in: 60, T_out: 550, dT: 15, h: 0.1, label: "Préchauffage de l'air de combustion" },
  },
  athos: { froid: { T_in: 100, T_out: 250, dT: 5, h: 0.8, label: 'Athos — RTO' } },
  hvac: { froid: { T_in: 35, T_out: 55, dT: 5, h: 0.4, label: 'Chauffage des bâtiments' } },
}

// Utilités disponibles, ordonnées par coût croissant : le solveur les place
// dans cet ordre sur la grande courbe composite.
const UTILITES = [
  { nom: 'Récupération sur fumées', type: 'chaude', T: 250, cout: 0, dT_contribution: 15 },
  { nom: 'Biogaz', type: 'chaude', T: 500, cout: 1, dT_contribution: 10 },
  { nom: 'Gaz naturel', type: 'chaude', T: 500, cout: 2, dT_contribution: 10 },
  { nom: 'Eau de refroidissement', type: 'froide', T: 25, cout: 1, dT_contribution: 5 },
  { nom: 'Aéroréfrigérant', type: 'froide', T: 35, cout: 2, dT_contribution: 10 },
]

/**
 * Convertit les déclarations `energie` des procédés en flux thermiques.
 * Une puissance moyenne est déduite du besoin journalier, réparti sur les
 * heures de fonctionnement du poste.
 */
export function fluxDepuisFilieres(boues, utilites, options = {}) {
  const flux = []
  const heures = options.heures_fonctionnement ?? CONST.NOMBRE_HEURE_PAR_JOUR
  const enKW = (kWhj) => kWhj / heures

  const ajouter = (cle, sens, puissance, nom) => {
    const n = NIVEAUX[cle]?.[sens]
    if (!n || !(puissance > 0)) return
    flux.push({
      nom: nom ?? n.label,
      type: sens,
      T_in: n.T_in,
      T_out: n.T_out,
      charge: puissance,
      dT_contribution: n.dT,
      h: n.h,
      origine: cle,
    })
  }

  for (const s of boues?.steps ?? []) {
    const e = s.energie
    if (!e) continue
    const id = s.nodeId
    if (id === 'digestion') ajouter('digestion', 'froid', enKW(e.besoin_thermique_kWhj), `${s.label} — maintien`)
    else if (id === 'biothelys' || id === 'exelys') {
      ajouter(id, 'froid', enKW(e.besoin_thermique_kWhj), `${s.label} — vapeur`)
      ajouter(id, 'chaud', enKW(e.recuperable_kWhj), `${s.label} — refroidissement`)
    } else if (id?.startsWith('sechage')) {
      ajouter('sechage', 'froid', enKW(e.besoin_thermique_kWhj), `${s.label} — apport`)
      ajouter('sechage', 'chaud', enKW(e.recuperable_kWhj), `${s.label} — buées`)
    } else if (id === 'incineration') {
      ajouter('incineration', 'chaud', enKW(e.recuperable_kWhj), `${s.label} — fumées`)
      ajouter('incineration', 'froid', enKW(e.besoin_thermique_kWhj), `${s.label} — air de combustion`)
    } else if (id === 'athos') ajouter('athos', 'froid', enKW(e.besoin_thermique_kWhj), `${s.label} — RTO`)
  }
  for (const s of utilites?.steps ?? []) {
    const e = s.energie
    if (!e || !(e.besoin_thermique_kWhj > 0)) continue
    if (s.nodeId?.startsWith('hvac')) ajouter('hvac', 'froid', enKW(e.besoin_thermique_kWhj), `${s.label} — chauffage`)
  }
  return flux
}

export const energieThermique = defineUtilityNode({
  id: 'energie-thermique',
  label: 'Intégration énergétique',
  short: 'Pincement',
  family: 'production',
  vba: 'z_Gestion_energie_thermique.cls, PINCH.cls',
  description:
    "Analyse pincement des flux thermiques de la filière : cibles de récupération, température de pincement, courbes composites et placement des utilités par niveau de coût.",
  choices: [
    { key: 'utilites_multiples', label: 'Placement des utilités', default: 'multiples', options: [
      { value: 'multiples', label: 'plusieurs niveaux, du moins cher au plus cher' },
      { value: 'unique', label: 'une seule utilité chaude et une seule froide' },
    ] },
  ],
  params: [
    { key: 'dT_contribution', label: 'Contribution ΔT par défaut', unit: 'K', group: 'Analyse', default: 5, hint: "l'écart entre deux flux est la somme de leurs contributions" },
    { key: 'heures_fonctionnement', label: 'Heures de fonctionnement des postes thermiques', unit: 'h/j', group: 'Analyse', default: 24 },
    { key: 'T_biogaz', label: 'Température de la chaleur issue du biogaz', unit: '°C', group: 'Utilités', default: 500 },
    { key: 'capacite_biogaz', label: 'Puissance disponible en biogaz', unit: 'kW', group: 'Utilités', default: undefined, hint: 'déduite du biogaz produit si non forcée' },
    { key: 'T_eau_refroidissement', label: "Température de l'eau de refroidissement", unit: '°C', group: 'Utilités', default: 25 },
  ],

  compute(ctx) {
    const { contexte, p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const heures = p.heures_fonctionnement ?? 24

    const fluxBruts = contexte.fluxThermiques ?? []
    if (!fluxBruts.length) {
      return {
        results: [],
        electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Aucun flux thermique déclaré : l'intégration énergétique suppose des procédés qui consomment ou produisent de la chaleur — digestion, séchage, incinération, hydrolyse ou chauffage de bâtiments."],
      }
    }

    // ---- utilités disponibles
    let utilites = UTILITES.map((u) => ({ ...u }))
    const T_biogaz = p.T_biogaz ?? 500
    utilites = utilites.map((u) => (u.nom === 'Biogaz' ? { ...u, T: T_biogaz } : u))
    // la puissance de biogaz disponible borne l'utilité correspondante
    const capacite_biogaz = f('capacite_biogaz')
      ?? (contexte.biogaz_Nm3j * (contexte.biogaz_CH4 || 0.63) * 9.94) / heures
    utilites = utilites.map((u) => (u.nom === 'Biogaz' ? { ...u, capacite: capacite_biogaz } : u))
    utilites = utilites.map((u) => (u.nom === 'Eau de refroidissement' ? { ...u, T: p.T_eau_refroidissement ?? 25 } : u))
    if (choices.utilites_multiples === 'unique') {
      utilites = [
        utilites.filter((u) => u.type === 'chaude').pop(),
        utilites.filter((u) => u.type === 'froide').pop(),
      ].filter(Boolean)
    }

    const analyse = analyserPincement(fluxBruts, {
      dT_contribution_defaut: p.dT_contribution ?? 5,
      utilites,
    })
    if (analyse.erreur) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: [analyse.erreur] }
    }

    const b = analyse.bilan
    const c = analyse.cibles
    for (const a of analyse.alertes) {
      if (a.niveau === 'attention') warnings.push(a.texte)
    }
    if (b.economie > 0.6) {
      warnings.push(`Potentiel de récupération élevé (${(b.economie * 100).toFixed(0)} %) : l'intégration mérite une étude de réseau d'échangeurs détaillée.`)
    }

    const jour = (kW) => kW * heures

    const results = [
      { key: 'nb_flux', label: 'Flux thermiques analysés', unit: 'u', value: analyse.flux.length },
      { key: 'charge_chaude', label: 'Chaleur disponible sur les flux chauds', unit: 'kW', value: b.charge_chaude },
      { key: 'charge_froide', label: 'Chaleur requise par les flux froids', unit: 'kW', value: b.charge_froide },
      { key: 'Qh', label: 'Utilité chaude minimale', unit: 'kW', value: b.Qh_min },
      { key: 'Qh_j', label: 'Utilité chaude minimale', unit: 'kWh/j', value: jour(b.Qh_min) },
      { key: 'Qc', label: 'Utilité froide minimale', unit: 'kW', value: b.Qc_min },
      { key: 'Qc_j', label: 'Utilité froide minimale', unit: 'kWh/j', value: jour(b.Qc_min) },
      { key: 'recup', label: 'Récupération maximale', unit: 'kW', value: b.recuperation },
      { key: 'recup_j', label: 'Récupération maximale', unit: 'kWh/j', value: jour(b.recuperation) },
      { key: 'sans_integration', label: 'Utilité chaude sans intégration', unit: 'kW', value: b.Qh_sans_integration },
      { key: 'economie', label: "Économie apportée par l'intégration", unit: '-', value: b.economie },
      ...(b.T_pincement_decale != null ? [
        { key: 'T_pincement', label: 'Température de pincement (échelle décalée)', unit: '°C', value: b.T_pincement_decale },
        { key: 'T_pincement_chaud', label: 'Pincement — côté chaud', unit: '°C', value: b.T_pincement_chaud },
        { key: 'T_pincement_froid', label: 'Pincement — côté froid', unit: '°C', value: b.T_pincement_froid },
      ] : []),
      { key: 'unites', label: "Nombre minimal d'échangeurs", unit: 'u', value: c.unites_min },
      { key: 'n_dessus', label: 'dont au-dessus du pincement', unit: 'u', value: Math.max(0, c.n_dessus - 1) },
      { key: 'n_dessous', label: 'dont en dessous du pincement', unit: 'u', value: Math.max(0, c.n_dessous - 1) },
      { key: 'surface', label: "Surface d'échange cible", unit: 'm²', value: c.surface },
    ]
    for (const u of analyse.utilites?.chaudes ?? []) {
      results.push({ key: `uc_${u.nom}`, label: `Utilité chaude — ${u.nom}`, unit: 'kW', value: u.charge })
    }
    for (const u of analyse.utilites?.froides ?? []) {
      results.push({ key: `uf_${u.nom}`, label: `Utilité froide — ${u.nom}`, unit: 'kW', value: u.charge })
    }

    // le besoin thermique résiduel est celui que les utilités doivent couvrir
    const biogaz_utilise = analyse.utilites?.chaudes.find((u) => u.nom === 'Biogaz')?.charge ?? 0
    const fossile = (analyse.utilites?.chaudes ?? [])
      .filter((u) => u.nom === 'Gaz naturel')
      .reduce((s, u) => s + u.charge, 0)

    return {
      results,
      pincement: analyse,
      energie: {
        besoin_thermique_kWhj: jour(fossile),
        source: 'gaz_naturel',
        recuperable_kWhj: jour(b.recuperation),
        biogaz_valorise_kWhj: jour(biogaz_utilise),
      },
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})

export { NIVEAUX, UTILITES }
