// ---------------------------------------------------------------------------
// File boues — modèle de flux.
//
// OCEAN véhicule la file boues dans deux tableaux à trois dimensions, déclarés
// dans MOD_ProgrammePrincipal :
//   TableauRecapitulatifFluxBoues(étape, type_boue, paramètre)   23 × 6 × 10
//   boues_pollution_soluble(étape, type_boue, repère_mgL)        23 × 6 × 5
// L'étape repère la position dans la filière (entrée, épaissies 1..5, digérées,
// déshydratées 1..3, chaulées, séchées, incinérées, évacuées) et le type de
// boue son origine (primaire, secondaire, tertiaire, externes, graisses).
//
// Le port conserve cette double indexation : c'est elle qui permet à
// l'épaississement de retrouver la concentration de référence propre à chaque
// origine de boue, et au bilan matière de se vérifier étape par étape.
//
// Conventions d'unités, identiques à OCEAN :
//   Q   en m³/j
//   MES en kg/j                        (matières en suspension, part solide)
//   MV_MES sans dimension              (en réalité MV/MS)
//   ratio_*_MES sans dimension         (DCO, DBO, NK, Pt rapportés aux MES)
//   pollution soluble en mg/L
// ---------------------------------------------------------------------------

/** étapes de la file boues (constantes boues_* du VBA) */
export const ETAPE = {
  inlet: 1,
  epaissies: [null, 2, 3, 4, 5, 6], // epaissies[1..5]
  digerees: 7,
  athos: 8,
  deshydratees: [null, 9, 10, 11],
  chaulees: [null, 12, 13, 14],
  sechees: [null, 15, 16],
  sechees_bioco: [null, 17, 18],
  sechees_inos: [null, 19, 20],
  incinerees: 21,
  graisses_biolix: 22,
  evacuees: 23,
}
export const NB_ETAPES = 23

/** types de boue en entrée de filière (constantes boues_I… du VBA) */
export const TYPE = {
  I: 1, // primaire
  II: 2, // secondaire
  III: 3, // tertiaire
  externes_1: 4,
  externes_2: 5,
  graisses: 6,
}
export const NB_TYPES = 6
export const TYPE_LABELS = {
  1: 'Boues primaires',
  2: 'Boues secondaires',
  3: 'Boues tertiaires',
  4: 'Boues externes 1',
  5: 'Boues externes 2',
  6: 'Graisses',
}

/** paramètres portés par un flux de boues (constantes repere_* du VBA) */
export const P = {
  origine: 1, // clé de ratio() : 'I_simple', 'II_faible_EB'…
  MES: 2, // kg/j
  Q: 3, // m³/j
  MV_MES: 4, // -
  verif_flux: 5, // bilan matière : part de flux encore disponible
  flux_in: 6, // part de flux entrée dans l'étape
  ratio_DCO_MES: 7,
  ratio_DBO_MES: 8,
  ratio_NK_MES: 9,
  ratio_Pt_MES: 10,
}
export const NB_PARAM = 10

/** repères de la pollution soluble véhiculée avec l'eau interstitielle (mg/L) */
export const SOL = { DCO: 1, NK: 2, Pt: 3, MS_soluble: 4, MV_soluble: 5 }
export const NB_SOL = 5

/** repères du flux de retour en tête de station */
export const RET = { Q: 1, MES: 2, DCO: 3, DBO: 4, NK: 5, NH4: 6, NO3: 7, Pt: 8 }
export const NB_RET = 8

/**
 * Origines de retour tenues séparément. OCEAN distingue ces vecteurs parce que
 * le traitement des retours (ANITA Mox, Shunt, MAP) ne s'applique qu'aux jus
 * les plus chargés en azote — ceux de la digestion et de l'Athos — et non aux
 * surnageants d'épaississement ou aux condensats de séchage.
 * `total` est la somme, tenue à jour en parallèle.
 */
export const RET_ORIGINE = {
  total: 'total',
  digestion: 'digestion',
  athos: 'athos',
  autres: 'autres',
}

// ---------------------------------------------------------------------------
// Fabriques
// ---------------------------------------------------------------------------
const cube = (a, b, c, init = 0) =>
  Array.from({ length: a + 1 }, () => Array.from({ length: b + 1 }, () => new Array(c + 1).fill(init)))

/** tableau des flux de boues, indexé [étape][type][paramètre] */
export function makeSludgeTable() {
  const t = cube(NB_ETAPES, NB_TYPES, NB_PARAM)
  // repere_origine porte une chaîne, pas un nombre
  for (let e = 1; e <= NB_ETAPES; e++) for (let j = 1; j <= NB_TYPES; j++) t[e][j][P.origine] = ''
  return t
}
/** tableau de la pollution soluble, indexé [étape][type][repère] */
export const makeSolubleTable = () => cube(NB_ETAPES, NB_TYPES, NB_SOL)
/** vecteur de retour en tête de station */
export const makeRetour = () => new Array(NB_RET + 1).fill(0)

/**
 * Jeu de vecteurs de retour, un par origine, plus la part soluble de chacun.
 * Les procédés écrivent dans `total` et dans leur origine propre ; le nœud de
 * gestion des retours s'en sert pour décider ce qui part au traitement.
 */
export function makeRetours() {
  const r = {}
  for (const o of Object.values(RET_ORIGINE)) {
    r[o] = makeRetour()
    r[`${o}_soluble`] = makeRetour()
  }
  return r
}

/**
 * Ajoute un flux au vecteur total et au vecteur de son origine.
 * `flux` et `soluble` sont indexés par les repères RET.
 */
export function ajouterRetour(retours, origine, flux, soluble = null) {
  for (let i = 1; i <= NB_RET; i++) {
    const v = flux[i] || 0
    retours[RET_ORIGINE.total][i] += v
    if (origine && retours[origine]) retours[origine][i] += v
    if (soluble) {
      const sv = soluble[i] || 0
      retours[`${RET_ORIGINE.total}_soluble`][i] += sv
      if (origine && retours[`${origine}_soluble`]) retours[`${origine}_soluble`][i] += sv
    }
  }
  return retours
}

/**
 * Alimente l'étape « entrée » du tableau à partir des flux `sludge` émis par
 * les nœuds de la file eau. Chaque nœud secondaire/primaire/tertiaire déclare
 * son origine, qui détermine ensuite sa concentration de référence à
 * l'épaississement.
 */
export function chargerEntrees(table, apports) {
  for (const a of apports) {
    const j = a.type
    if (!j) continue
    const MES = a.MES || 0
    const Q = a.Q || 0
    if (!(MES > 0)) continue
    const cible = table[ETAPE.inlet][j]
    const MES_cumul = cible[P.MES] + MES
    // moyennes pondérées par les MES pour les ratios et le MV/MES
    const pond = (ancien, nouveau) => (cible[P.MES] * ancien + MES * nouveau) / MES_cumul
    cible[P.MV_MES] = MES_cumul > 0 ? pond(cible[P.MV_MES], a.MV_MES || 0) : 0
    cible[P.ratio_DCO_MES] = MES_cumul > 0 ? pond(cible[P.ratio_DCO_MES], MES > 0 ? (a.DCO || 0) / MES : 0) : 0
    cible[P.ratio_DBO_MES] = MES_cumul > 0 ? pond(cible[P.ratio_DBO_MES], MES > 0 ? (a.DBO || 0) / MES : 0) : 0
    cible[P.ratio_NK_MES] = MES_cumul > 0 ? pond(cible[P.ratio_NK_MES], MES > 0 ? (a.NK || 0) / MES : 0) : 0
    cible[P.ratio_Pt_MES] = MES_cumul > 0 ? pond(cible[P.ratio_Pt_MES], MES > 0 ? (a.Pt || 0) / MES : 0) : 0
    cible[P.MES] = MES_cumul
    cible[P.Q] += Q
    cible[P.origine] = a.origine || cible[P.origine]
    cible[P.flux_in] = 1
    cible[P.verif_flux] = 1
    // l'origine est propagée sur toutes les étapes, comme dans le VBA
    for (let e = 1; e <= NB_ETAPES; e++) if (!table[e][j][P.origine]) table[e][j][P.origine] = a.origine || ''
  }
  return table
}

/** somme des MES d'une étape, tous types confondus (kg/j) */
export const MES_etape = (table, etape) => {
  let s = 0
  for (let j = 1; j <= NB_TYPES; j++) s += table[etape][j][P.MES]
  return s
}
/** somme des débits d'une étape (m³/j) */
export const Q_etape = (table, etape) => {
  let s = 0
  for (let j = 1; j <= NB_TYPES; j++) s += table[etape][j][P.Q]
  return s
}
/** siccité d'une étape (g/L, équivalent à kg/m³) */
export const siccite_etape = (table, etape) => {
  const Q = Q_etape(table, etape)
  return Q > 0 ? MES_etape(table, etape) / Q : 0
}

/**
 * Détermine le type de boue (I, II, III…) à partir de la famille du nœud de la
 * file eau qui l'a produite.
 */
export function typeDepuisFamille(family, nodeId) {
  if (family === 'primaire') return TYPE.I
  if (family === 'secondaire') return TYPE.II
  if (family === 'tertiaire' || nodeId === 'decantation-eaux-sales') return TYPE.III
  return TYPE.II
}
