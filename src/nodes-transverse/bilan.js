// ---------------------------------------------------------------------------
// Port de z_Bilan_electrique.cls et z_Empreinte_CO2.cls.
//
// Ces deux modules ne calculent aucun procédé : ils agrègent ce que les trois
// files ont déjà produit. Le bilan électrique répartit la consommation par
// poste et la rapporte au débit, à la pollution éliminée et à l'équivalent
// habitant ; l'empreinte CO2 convertit les mêmes flux en émissions, réactif par
// réactif, en tenant compte du transport et de l'électricité verte produite.
//
// Deux notions structurent le bilan électrique :
//   part fixe      ce qui tourne indépendamment de la charge (agitation,
//                  brassage, racleurs) ;
//   part variable  ce qui suit la charge (aération, pompages).
// Les nœuds portés exposent déjà `electricity.fixed` : la part variable est le
// complément.
//
// L'électricité verte peut être autoconsommée ou vendue. Autoconsommée, elle
// **réduit** les émissions du site ; vendue, elle en **évite** ailleurs. Le
// classeur distingue les deux et ne les additionne pas dans le même total :
// les émissions réduites sont retranchées, les évitées sont exposées à part.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - le poste « autres » lu depuis une feuille du classeur n'existe pas ici,
//    il est remplacé par un paramètre forçable ;
//  - le protoxyde d'azote et le méthane fugitif ne figurent pas dans la classe
//    d'origine : l'empreinte ne porte que le CO2 des consommations, pas les
//    émissions directes du procédé ;
//  - les incertitudes sont portées et propagées, mais le classeur écrase
//    plusieurs fois `FE_reactifs_incertitude(H2O2)` en lieu et place de CuSO4
//    et MgCl2 — reproduit, sans effet puisque leurs facteurs sont nuls.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { CONST } from '../core/hypotheses.js'

// Postes du bilan électrique, dans l'ordre d'affichage du classeur.
export const POSTES = [
  { id: 'relevement', label: 'Relèvement', file: 'eau' },
  { id: 'pretraitement', label: 'Prétraitement', file: 'eau' },
  { id: 'primaire', label: 'Traitement primaire', file: 'eau' },
  { id: 'secondaire', label: 'Traitement secondaire', file: 'eau' },
  { id: 'tertiaire', label: 'Traitement tertiaire', file: 'eau' },
  { id: 'desinfection', label: 'Désinfection', file: 'eau' },
  { id: 'epaississement', label: 'Épaississement', file: 'boues' },
  { id: 'stabilisation', label: 'Stabilisation / digestion', file: 'boues' },
  { id: 'deshydratation', label: 'Déshydratation', file: 'boues' },
  { id: 'hygienisation', label: 'Hygiénisation', file: 'boues' },
  { id: 'sechage', label: 'Séchage', file: 'boues' },
  { id: 'valorisation', label: 'Valorisation / élimination', file: 'boues' },
  { id: 'retours', label: 'Traitement des retours', file: 'boues' },
  { id: 'desodorisation', label: 'Désodorisation', file: 'utilites' },
  { id: 'batiments', label: 'Bâtiments', file: 'utilites' },
  { id: 'production', label: "Production d'énergie", file: 'utilites' },
  { id: 'autres', label: 'Autres postes', file: 'utilites' },
]
const POSTE_PAR_FAMILLE = {
  relevement: 'relevement', pretraitement: 'pretraitement', primaire: 'primaire',
  secondaire: 'secondaire', tertiaire: 'tertiaire', desinfection: 'desinfection',
  epaississement: 'epaississement', stabilisation: 'stabilisation',
  deshydratation: 'deshydratation', hygienisation: 'hygienisation',
  sechage: 'sechage', valorisation: 'valorisation', retours: 'retours',
  desodorisation: 'desodorisation', batiments: 'batiments',
  production: 'production', reactifs: 'autres',
}

// Facteurs d'émission des réactifs, en kg CO2 par tonne de produit pur, avec
// leur incertitude relative. Sources citées dans le classeur : ADEME, ASTEE,
// Eco-Invent, Carbone 4, mise à jour du 08/11/2012.
const FE_REACTIFS = {
  polymere: { fe: 810, inc: 0.3, source: 'SNF/ASTEE v5.2010.2' },
  FeCl3: { fe: 800, inc: 0.25, source: 'Eco-Invent' },
  // le méthanol cumule les émissions de production et celles de sa combustion
  methanol: { fe: 521 + 1375, inc: (521 * 0.5 + 1375 * 0) / (521 + 1375), source: 'ADEME BC v7 + Carbone 4' },
  chaux_eteinte: { fe: 846, inc: 0.3, source: 'LHOIST v5.2010.2' },
  chaux_vive: { fe: 1110, inc: 0.3, source: 'LHOIST v5.2010.2/ASTEE' },
  H2SO4: { fe: 148, inc: 0.5, source: 'ADEME BC v7' },
  NaOH: { fe: 1174, inc: 0.5, source: 'ADEME BC v7' },
  soude: { fe: 1174, inc: 0.5, source: 'ADEME BC v7' },
  NaOCl: { fe: 2884, inc: 0.5, source: 'Arkema' },
  NaHSO3: { fe: 420, inc: 0.1, source: 'ASTEE 5.2010.2' },
  NaHCO3: { fe: 1166, inc: 0.1, source: 'Carbone 4' },
  bicarbonate_sodium: { fe: 1166, inc: 0.1, source: 'Carbone 4' },
  Ca_2NO3: { fe: 640, inc: 0.3, source: 'YARA / ASTEE 05.2010.2' },
  nitrate_calcium: { fe: 640, inc: 0.3, source: 'YARA / ASTEE 05.2010.2' },
  oxygene_liquide: { fe: 408, inc: 0.27, source: 'Eco-Invent v2.2' },
  H2O2: { fe: 1145, inc: 0.16, source: 'Eco-Invent v2.2' },
  charbon_actif: { fe: 7000, inc: 0.1, source: 'ASTEE 5.2010.2' },
  Cl2: { fe: 947.5, inc: 0.25, source: 'Carbone 4' },
  ammoniaque: { fe: 2110, inc: 0.5, source: "Ec-Eau 2.8 / bilan produit ADEME" },
  uree: { fe: 1230, inc: 0.5, source: "Ec-Eau 2.8 / bilan produit ADEME" },
  // le classeur ne dispose pas de facteur pour ces deux réactifs
  CuSO4: { fe: 0, inc: 0.5, source: "pas de donnée dans le classeur" },
  MgCl2: { fe: 0, inc: 0.5, source: "pas de donnée dans le classeur" },
  azote: { fe: 2110, inc: 0.5, source: 'assimilé à l\'ammoniaque' },
  phosphore: { fe: 0, inc: 0.5, source: 'pas de donnée dans le classeur' },
}

const H = {
  FE_transport: { camion_15T: 0.238, camion_20T: 0.223, semi_remorque_25T: 0.119 },
  FE_transport_inc: 0.1,
  FE_transport_defaut: 0.129, // boues, graisses, cendres, refus
  FE_gaz_naturel: 0.03 + 0.18, // amont + combustion, kg CO2 / kWh PCI
  FE_gaz_naturel_inc: 0.05,
  FE_fioul: 0.06 + 0.27,
  FE_fioul_inc: 0.05,
  densite_boues: 1, // t/m³
  emissions_reference_EH_an: 75, // kg CO2, vol Paris-Toulouse par voyageur
  DBO_par_EH_gj: 60,
}
// Facteur d'émission de l'électricité, kg CO2 / kWh (ADEME).
const FE_ELECTRICITE = {
  france: { fe: 0.091, inc: 0.1, label: 'France' },
  europe: { fe: 0.42, inc: 0.15, label: 'Europe' },
  monde: { fe: 0.6, inc: 0.2, label: 'Moyenne mondiale' },
}

// ---------------------------------------------------------------------------
/**
 * Répartit la consommation électrique par poste et par part fixe/variable, à
 * partir des trois moteurs. C'est la base commune du bilan et de l'empreinte.
 */
export function repartitionElectrique(eau, boues, utilites, registres) {
  const postes = {}
  for (const p of POSTES) postes[p.id] = { ...p, fixe: 0, variable: 0, total: 0, details: [] }

  const verser = (steps, registre, familleParDefaut) => {
    for (const s of steps ?? []) {
      const node = registre?.[s.nodeId]
      const famille = node?.family ?? familleParDefaut
      const posteId = POSTE_PAR_FAMILLE[famille] ?? 'autres'
      const poste = postes[posteId] ?? postes.autres
      const total = s.electricity?.total ?? 0
      const fixe = s.electricity?.fixed ?? 0
      // une production nette est portée au poste « production d'énergie »
      poste.fixe += fixe
      poste.variable += total - fixe
      poste.total += total
      poste.details.push({
        label: s.label,
        total,
        fixe,
        variable: total - fixe,
        detail: s.electricity?.detail ?? {},
      })
    }
  }
  verser(eau?.steps, registres?.eau, 'secondaire')
  verser(boues?.steps, registres?.boues, 'valorisation')
  verser(utilites?.steps, registres?.utilites, 'autres')

  const liste = POSTES.map((p) => postes[p.id]).filter((p) => Math.abs(p.total) > 1e-9)
  const consommee = liste.reduce((s, p) => s + Math.max(0, p.total), 0)
  const produite = liste.reduce((s, p) => s + Math.max(0, -p.total), 0)
  return { postes: liste, consommee, produite }
}

// ---------------------------------------------------------------------------
export const bilanElectrique = defineUtilityNode({
  id: 'bilan-electrique',
  label: 'Bilan électrique',
  short: 'Bilan élec',
  family: 'transverse',
  vba: 'z_Bilan_electrique.cls',
  description:
    "Répartit la consommation électrique par poste et par part fixe et variable, et la rapporte au débit traité, à la pollution éliminée et à l'équivalent habitant.",
  choices: [
    { key: 'valorisation_verte', label: "Valorisation de l'électricité verte", default: 'autoconsommee', options: [
      { value: 'autoconsommee', label: 'autoconsommée en priorité' },
      { value: 'vendue', label: 'intégralement vendue' },
    ] },
  ],
  params: [
    { key: 'electricite_autre', label: 'Consommation des postes non modélisés', unit: 'kWh/j', group: 'Compléments', default: 0 },
    { key: 'electricite_verte_biogaz', label: 'Électricité produite par le biogaz', unit: 'kWh/j', group: 'Production verte', default: undefined, hint: 'calculée sur le biogaz disponible' },
    { key: 'rendement_cogeneration', label: 'Rendement électrique de la cogénération', unit: '-', group: 'Production verte', default: 0.38 },
    { key: 'PCI_CH4', label: 'PCI du méthane', unit: 'kWh/Nm³', group: 'Production verte', default: 9.94 },
    { key: 'capacite_EH', label: 'Capacité nominale de la station', unit: 'EH', group: 'Ratios', default: undefined, hint: 'déduite de la DBO nominale' },
  ],

  compute(ctx) {
    const { site, contexte, p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const bilan = contexte.bilan
    if (!bilan) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Le bilan électrique doit être calculé après les trois files."] }
    }

    const autre = p.electricite_autre ?? 0
    const consommee = bilan.consommee + autre
    // le biogaz disponible se convertit en électricité par cogénération
    const verte_biogaz = f('electricite_verte_biogaz')
      ?? contexte.biogaz_Nm3j * (contexte.biogaz_CH4 || 0.63) * (p.PCI_CH4 ?? 9.94) * (p.rendement_cogeneration ?? 0.38)
    const verte_autre = bilan.produite
    const verte = verte_biogaz + verte_autre

    // autoconsommation : la production verte s'impute d'abord sur les besoins
    let verte_consommee, verte_vendue
    if (choices.valorisation_verte === 'autoconsommee') {
      verte_consommee = Math.min(verte, consommee)
      verte_vendue = verte - verte_consommee
    } else {
      verte_consommee = 0
      verte_vendue = verte
    }
    const autosuffisance = consommee > 0 ? verte / consommee : 0
    const autosuffisance_biogaz = consommee > 0 ? verte_biogaz / consommee : 0

    // ---- ratios
    const Q = contexte.Q_reel
    const capacite_EH = f('capacite_EH')
      ?? (site.DBO_nominal > 0 ? (site.DBO_nominal * 1000) / H.DBO_par_EH_gj : 0)
    const DCO_eliminee = contexte.DCO_eliminee ?? 0
    const DBO_eliminee = contexte.DBO_eliminee ?? 0
    const ratio_Q = Q > 0 ? consommee / Q : 0
    const ratio_EH = capacite_EH > 0 ? (consommee * CONST.NOMBRE_JOUR_PAR_AN) / capacite_EH : 0
    const ratio_DCO = DCO_eliminee > 0 ? consommee / DCO_eliminee : 0
    const ratio_DBO = DBO_eliminee > 0 ? consommee / DBO_eliminee : 0

    const fixe_total = bilan.postes.reduce((s, x) => s + Math.max(0, x.fixe), 0)
    const variable_total = consommee - fixe_total
    if (autosuffisance > 1) {
      warnings.push(`La production verte (${verte.toFixed(0)} kWh/j) dépasse la consommation : la station est excédentaire.`)
    }

    const results = [
      { key: 'consommee', label: 'Électricité consommée', unit: 'kWh/j', value: consommee },
      { key: 'consommee_an', label: 'Électricité consommée', unit: 'MWh/an', value: (consommee * CONST.NOMBRE_JOUR_PAR_AN) / 1000 },
      { key: 'fixe', label: 'dont part fixe', unit: 'kWh/j', value: fixe_total },
      { key: 'variable', label: 'dont part variable', unit: 'kWh/j', value: variable_total },
      { key: 'part_fixe', label: 'Part fixe de la consommation', unit: '-', value: consommee > 0 ? fixe_total / consommee : 0 },
      { key: 'verte', label: 'Électricité verte produite', unit: 'kWh/j', value: verte },
      { key: 'verte_biogaz', label: 'dont cogénération biogaz', unit: 'kWh/j', value: verte_biogaz },
      { key: 'verte_autre', label: 'dont turbine et photovoltaïque', unit: 'kWh/j', value: verte_autre },
      { key: 'verte_consommee', label: 'Électricité verte autoconsommée', unit: 'kWh/j', value: verte_consommee },
      { key: 'verte_vendue', label: 'Électricité verte vendue', unit: 'kWh/j', value: verte_vendue },
      { key: 'autosuff', label: 'Taux d\'autosuffisance', unit: '-', value: autosuffisance },
      { key: 'autosuff_biogaz', label: 'dont apporté par le biogaz', unit: '-', value: autosuffisance_biogaz },
      { key: 'bilan_net', label: 'Consommation nette du site', unit: 'kWh/j', value: consommee - verte_consommee },
      { key: 'ratio_Q', label: 'Consommation rapportée au débit', unit: 'kWh/m³', value: ratio_Q },
      { key: 'ratio_EH', label: 'Consommation par équivalent habitant', unit: 'kWh/(EH·an)', value: ratio_EH },
      ...(ratio_DCO > 0 ? [{ key: 'ratio_DCO', label: 'Consommation par kg de DCO éliminée', unit: 'kWh/kg', value: ratio_DCO }] : []),
      ...(ratio_DBO > 0 ? [{ key: 'ratio_DBO', label: 'Consommation par kg de DBO éliminée', unit: 'kWh/kg', value: ratio_DBO }] : []),
    ]
    for (const poste of bilan.postes) {
      results.push({
        key: `poste_${poste.id}`,
        label: `Poste — ${poste.label}`,
        unit: 'kWh/j',
        value: poste.total,
      })
    }

    return {
      results,
      bilan: {
        consommee, verte, verte_consommee, verte_vendue, verte_biogaz,
        fixe: fixe_total, variable: variable_total,
        autosuffisance, ratio_Q, ratio_EH, capacite_EH,
        postes: bilan.postes,
      },
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})

// ---------------------------------------------------------------------------
export const empreinteCO2 = defineUtilityNode({
  id: 'empreinte-co2',
  label: 'Empreinte CO2',
  short: 'CO2',
  family: 'transverse',
  vba: 'z_Empreinte_CO2.cls',
  description:
    "Convertit les consommations d'électricité, de réactifs et de combustibles en émissions de CO2, transport compris, et déduit les émissions réduites par l'électricité verte autoconsommée.",
  choices: [
    { key: 'pays', label: "Facteur d'émission de l'électricité", default: 'france', options: Object.entries(FE_ELECTRICITE).map(([value, x]) => ({ value, label: `${x.label} — ${x.fe} kg CO2/kWh` })) },
    { key: 'transport_reactifs', label: 'Mode de transport des réactifs', default: 'camion_20T', options: [
      { value: 'camion_15T', label: 'camion 15 t' },
      { value: 'camion_20T', label: 'camion 20 t' },
      { value: 'semi_remorque_25T', label: 'semi-remorque 25 t' },
    ] },
  ],
  params: [
    { key: 'distance_reactifs', label: 'Distance de transport des réactifs', unit: 'km', group: 'Transport', default: 200 },
    { key: 'distance_boues', label: 'Distance de transport des boues', unit: 'km', group: 'Transport', default: 50 },
    { key: 'FE_transport_boues', label: "Facteur d'émission du transport des boues", unit: 'kg CO2/(t·km)', group: 'Transport', default: 0.129 },
    { key: 'gaz_naturel', label: 'Consommation de gaz naturel', unit: 'kWh PCI/j', group: 'Combustibles', default: 0 },
    { key: 'fioul', label: 'Consommation de fioul', unit: 'kWh PCI/j', group: 'Combustibles', default: 0 },
    { key: 'FE_electricite', label: "Facteur d'émission de l'électricité", unit: 'kg CO2/kWh', group: 'Électricité', default: undefined, hint: 'selon le pays retenu' },
  ],

  compute(ctx) {
    const { site, contexte, p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const bilan = contexte.bilanElectrique
    if (!bilan) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["L'empreinte CO2 doit être calculée après le bilan électrique."] }
    }

    const FE_elec = f('FE_electricite') ?? FE_ELECTRICITE[choices.pays].fe
    const FE_elec_inc = FE_ELECTRICITE[choices.pays].inc
    const FE_transport = H.FE_transport[choices.transport_reactifs]
    const distance = p.distance_reactifs ?? 200
    const an = CONST.NOMBRE_JOUR_PAR_AN

    // les compartiments d'émission, en tonnes de CO2 par an
    const compartiments = { electricite: 0, reactifs: 0, transport: 0, gaz_naturel: 0, fioul: 0 }
    let incertitude_absolue = 0
    const lignes = []

    // ---- réactifs : production puis transport
    for (const [cle, kgj] of Object.entries(contexte.reactifs ?? {})) {
      if (!(kgj > 0)) continue
      const ref = FE_REACTIFS[cle]
      const tonnes_an = (kgj / 1000) * an
      if (!ref) {
        warnings.push(`Réactif « ${cle} » sans facteur d'émission de référence : non compté dans l'empreinte.`)
        continue
      }
      const production = (tonnes_an * ref.fe) / 1000
      const transport = (tonnes_an * FE_transport * distance) / 1000
      compartiments.reactifs += production
      compartiments.transport += transport
      incertitude_absolue += ref.inc * production + H.FE_transport_inc * transport
      lignes.push({ label: ref.label ?? cle, cle, tonnes_an, production, transport, total: production + transport, source: ref.source })
    }

    // ---- transport des boues évacuées
    const boues_m3j = contexte.boues_evacuees_Q ?? 0
    const FE_boues = p.FE_transport_boues ?? H.FE_transport_defaut
    const transport_boues = (boues_m3j * an * H.densite_boues * FE_boues * (p.distance_boues ?? 50)) / 1000
    compartiments.transport += transport_boues
    incertitude_absolue += H.FE_transport_inc * transport_boues

    // ---- combustibles
    const gaz = (p.gaz_naturel ?? 0) + (contexte.gaz_naturel_kWhPCIj ?? 0)
    const fioul = (p.fioul ?? 0) + (contexte.fioul_kWhPCIj ?? 0)
    compartiments.gaz_naturel = (H.FE_gaz_naturel * gaz * an) / 1000
    compartiments.fioul = (H.FE_fioul * fioul * an) / 1000
    incertitude_absolue += H.FE_gaz_naturel_inc * compartiments.gaz_naturel + H.FE_fioul_inc * compartiments.fioul

    // ---- électricité consommée
    compartiments.electricite = (FE_elec * bilan.consommee * an) / 1000
    incertitude_absolue += FE_elec_inc * compartiments.electricite

    const emissions_brutes = Object.values(compartiments).reduce((s, x) => s + x, 0)
    // l'électricité verte autoconsommée réduit les émissions du site,
    // celle qui est vendue en évite ailleurs — les deux ne se cumulent pas
    const reduites = (FE_elec * bilan.verte_consommee * an) / 1000
    const evitees = (FE_elec * bilan.verte_vendue * an) / 1000
    const emissions_nettes = emissions_brutes - reduites
    const incertitude = emissions_brutes > 0 ? incertitude_absolue / emissions_brutes : 0

    // ---- ratios
    const Q = contexte.Q_reel
    const capacite_EH = bilan.capacite_EH ?? 0
    const ratio_EH = capacite_EH > 0 ? (emissions_nettes * 1000) / capacite_EH : 0
    const ratio_Q = Q > 0 ? ((emissions_nettes / an) * 1e6) / Q : 0
    const CNP = (contexte.DCO_eliminee ?? 0) + (contexte.N_elimine ?? 0) + (contexte.P_elimine ?? 0)
    const ratio_CNP = CNP > 0 ? ((emissions_nettes / an) * 1e6) / CNP : 0
    const equivalent_vols = ratio_EH / H.emissions_reference_EH_an

    const repartition = Object.entries(compartiments)
      .map(([id, v]) => ({ id, valeur: v, part: emissions_nettes > 0 ? v / emissions_nettes : 0 }))
      .filter((x) => Math.abs(x.part) > 1e-4)
      .sort((a, b) => b.valeur - a.valeur)
    lignes.sort((a, b) => b.total - a.total)

    const LIB = {
      electricite: 'Électricité', reactifs: 'Réactifs', transport: 'Transport',
      gaz_naturel: 'Gaz naturel', fioul: 'Fioul',
    }
    const results = [
      { key: 'brutes', label: 'Émissions brutes', unit: 't CO2/an', value: emissions_brutes },
      { key: 'reduites', label: "Émissions réduites par l'électricité verte autoconsommée", unit: 't CO2/an', value: reduites },
      { key: 'nettes', label: 'Émissions nettes du site', unit: 't CO2/an', value: emissions_nettes },
      { key: 'evitees', label: "Émissions évitées ailleurs par l'électricité vendue", unit: 't CO2/an', value: evitees },
      { key: 'incertitude', label: 'Incertitude sur les émissions', unit: '-', value: incertitude },
      ...repartition.map((x) => ({ key: `comp_${x.id}`, label: `Compartiment — ${LIB[x.id] ?? x.id}`, unit: 't CO2/an', value: x.valeur })),
      { key: 'ratio_EH', label: 'Émissions par équivalent habitant', unit: 'kg CO2/(EH·an)', value: ratio_EH },
      { key: 'ratio_Q', label: 'Émissions rapportées au débit', unit: 'g CO2/m³', value: ratio_Q },
      ...(ratio_CNP > 0 ? [{ key: 'ratio_CNP', label: 'Émissions par kg de C, N et P éliminés', unit: 'g CO2/kg', value: ratio_CNP }] : []),
      { key: 'vols', label: "Équivalent en vols Paris-Toulouse par habitant et par an", unit: 'vol', value: equivalent_vols },
    ]

    return {
      results,
      co2: { emissions_brutes, reduites, evitees, emissions_nettes, incertitude, repartition, lignes, ratio_EH, ratio_Q, FE_elec },
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})

export { FE_REACTIFS, FE_ELECTRICITE }
