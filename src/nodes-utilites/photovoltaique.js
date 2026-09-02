// ---------------------------------------------------------------------------
// Port de Photovoltaique.cls.
//
// Bonne surprise : le modèle du classeur est **astronomique**, pas tabulé. Il
// calcule l'irradiation extraterrestre à partir de la seule latitude et du jour
// de l'année — déclinaison solaire, angle horaire au lever, distance
// Terre-Soleil. Aucune base de données n'est nécessaire pour cette partie, qui
// est reprise telle quelle.
//
// La seule grandeur que le classeur va chercher dans sa base de stations est
// l'indice de clarté, c'est-à-dire la fraction du rayonnement extraterrestre
// qui atteint effectivement le sol après traversée de l'atmosphère. Le port en
// fait un paramètre, avec des préréglages par type de climat : de 0,40 sous un
// ciel océanique à 0,65 en climat désertique. C'est une donnée de site que
// l'utilisateur connaît en général mieux qu'une station météo choisie par
// identifiant.
//
// Le rendement des modules est corrigé de leur température de fonctionnement,
// estimée à partir de la NOCT et de l'irradiation — un module chauffe et perd
// entre 0,25 et 0,45 % de rendement par degré au-dessus de 25 °C, ce qui n'est
// pas négligeable en été.
//
// Écarts au VBA, assumés et signalés (voir README) :
//  - l'indice de clarté est saisi ou déduit d'un préréglage plutôt que lu dans
//    une base de 10 896 stations ;
//  - la répartition horaire de l'irradiation, calculée par pas de 0,1 h dans le
//    classeur pour établir un profil de production, n'est pas portée : seules
//    les productions mensuelles et annuelle le sont, qui sont les grandeurs
//    utilisées en aval par le bilan électrique.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { CONST } from '../core/hypotheses.js'

const H = {
  Gs: 0.082, // MJ/(m²·min), constante solaire
  MJ_vers_kWh: 1 / 3.6,
  T_reference_module: 25, // °C, conditions standard de test
  irradiation_NOCT: 800, // W/m², conditions de mesure de la NOCT
  T_air_NOCT: 20, // °C
}
// jour moyen de chaque mois, au sens de la moyenne d'irradiation
// (Al-Hallaj & Kiszynski, Hybrid Hydrogen Systems, Springer 2011)
const JOUR_MOYEN = [17, 47, 75, 105, 135, 162, 198, 228, 258, 288, 318, 344]
const JOURS_MOIS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
const MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']

// Indice de clarté : fraction du rayonnement extraterrestre atteignant le sol.
const CLARTE = {
  oceanique: { label: 'Océanique, ciel souvent couvert', valeur: 0.40 },
  tempere: { label: 'Tempéré', valeur: 0.45 },
  continental: { label: 'Continental', valeur: 0.48 },
  mediterraneen: { label: 'Méditerranéen', valeur: 0.55 },
  desertique: { label: 'Désertique, ciel très dégagé', valeur: 0.65 },
  tropical: { label: 'Tropical humide', valeur: 0.44 },
}
// Technologies de modules : rendement, NOCT et coefficient de température.
const TECHNOLOGIES = {
  mono: { label: 'Silicium monocristallin', rendement: 0.20, NOCT: 45, coefTemp: -0.0038 },
  poly: { label: 'Silicium polycristallin', rendement: 0.17, NOCT: 46, coefTemp: -0.0040 },
  amorphe: { label: 'Silicium amorphe', rendement: 0.08, NOCT: 48, coefTemp: -0.0020 },
  CdTe: { label: 'Tellurure de cadmium', rendement: 0.16, NOCT: 45, coefTemp: -0.0025 },
}

/**
 * Irradiation extraterrestre journalière sur plan horizontal, en kWh/(m²·j).
 * Reprise du calcul du classeur : distance Terre-Soleil, déclinaison solaire,
 * angle horaire au coucher.
 */
export function irradiationExtraterrestre(latitude_deg, jour) {
  const lat = (latitude_deg * Math.PI) / 180
  const d_invTS = 1 + 0.033 * Math.cos((2 * Math.PI * jour) / CONST.NOMBRE_JOUR_PAR_AN)
  const decl = 0.409 * Math.sin((2 * Math.PI * jour) / CONST.NOMBRE_JOUR_PAR_AN - 1.39)
  // angle horaire au coucher : borné aux jours et nuits polaires
  let cos_w = -Math.tan(lat) * Math.tan(decl)
  if (cos_w > 1) cos_w = 1
  if (cos_w < -1) cos_w = -1
  const w = Math.acos(cos_w)
  const MJ = ((24 * 60 * H.Gs * d_invTS) / Math.PI)
    * (w * Math.sin(lat) * Math.sin(decl) + Math.cos(lat) * Math.cos(decl) * Math.sin(w))
  return Math.max(0, MJ) * H.MJ_vers_kWh
}

/**
 * Facteur d'inclinaison : rapport entre l'irradiation reçue sur un plan incliné
 * et celle reçue à l'horizontale.
 *
 * Le rapport des angles d'incidence à midi solaire donne le gain **maximal** de
 * la journée, atteint au seul moment où le soleil est au plus haut ; intégré sur
 * la journée, le gain réel est plus faible, et la part diffuse du rayonnement ne
 * bénéficie pas de l'inclinaison. L'écart au facteur unité est donc atténué de
 * moitié, ce qui restitue le gain de 12 à 15 % observé aux latitudes moyennes
 * pour une inclinaison égale à la latitude.
 */
const ATTENUATION_INCLINAISON = 0.5

function facteurInclinaison(latitude_deg, inclinaison_deg, orientation_deg, jour) {
  const lat = (latitude_deg * Math.PI) / 180
  const beta = (inclinaison_deg * Math.PI) / 180
  const gamma = (orientation_deg * Math.PI) / 180
  const decl = 0.409 * Math.sin((2 * Math.PI * jour) / CONST.NOMBRE_JOUR_PAR_AN - 1.39)
  // hauteur du soleil à midi solaire
  const hauteur = Math.PI / 2 - Math.abs(lat - decl)
  if (hauteur <= 0) return 0
  // incidence sur le plan incliné, orientation comprise
  const cos_incidence = Math.sin(hauteur) * Math.cos(beta)
    + Math.cos(hauteur) * Math.sin(beta) * Math.cos(gamma)
  const facteur_midi = cos_incidence / Math.sin(hauteur)
  const facteur = 1 + ATTENUATION_INCLINAISON * (facteur_midi - 1)
  return Math.max(0, Math.min(1.5, facteur))
}

export const photovoltaique = defineUtilityNode({
  id: 'photovoltaique',
  label: 'Photovoltaïque',
  short: 'PV',
  family: 'production',
  vba: 'Photovoltaique.cls',
  description:
    "Production solaire sur la surface disponible. L'irradiation est calculée par modèle astronomique à partir de la latitude, corrigée d'un indice de clarté et de l'inclinaison des modules ; le rendement tient compte de leur échauffement.",
  choices: [
    { key: 'technologie', label: 'Technologie des modules', default: 'poly', options: Object.entries(TECHNOLOGIES).map(([value, t]) => ({ value, label: `${t.label} — ${(t.rendement * 100).toFixed(0)} %` })) },
    { key: 'clarte', label: 'Clarté du ciel', default: 'tempere', options: Object.entries(CLARTE).map(([value, c]) => ({ value, label: `${c.label} — ${c.valeur}` })) },
  ],
  params: [
    { key: 'surface_installee', label: 'Surface de modules installée', unit: 'm²', group: 'Installation', default: 500 },
    { key: 'latitude', label: 'Latitude du site', unit: '°', group: 'Installation', default: 48.85, hint: 'positive au nord, négative au sud' },
    { key: 'inclinaison', label: 'Inclinaison des modules', unit: '°', group: 'Installation', default: undefined, hint: 'latitude arrondie si non forcée' },
    { key: 'orientation', label: 'Écart à la direction du sud', unit: '°', group: 'Installation', default: 0, hint: '0 pour plein sud (plein nord dans l\'hémisphère sud)' },
    { key: 'indice_clarte', label: 'Indice de clarté', unit: '-', group: 'Rayonnement', default: undefined, hint: 'selon la clarté du ciel retenue' },
    { key: 'T_air_moyenne', label: "Température moyenne de l'air", unit: '°C', group: 'Rayonnement', default: 12 },
    { key: 'rendement_module', label: 'Rendement des modules', unit: '-', group: 'Modules', default: undefined, hint: 'selon la technologie' },
    { key: 'NOCT', label: 'Température de fonctionnement nominale', unit: '°C', group: 'Modules', default: undefined, hint: 'selon la technologie' },
    { key: 'coef_temperature', label: 'Coefficient de température', unit: '/°C', group: 'Modules', default: undefined, hint: 'perte de rendement par degré au-dessus de 25 °C' },
    { key: 'pertes_systeme', label: 'Pertes du système', unit: '-', group: 'Modules', default: 0.14, hint: 'onduleur, câblage, salissures' },
  ],

  compute(ctx) {
    const { p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)

    const surface = p.surface_installee ?? 500
    if (!(surface > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Surface de modules nulle.'] }
    }
    const latitude = p.latitude ?? 48.85
    if (Math.abs(latitude) > 90) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Latitude hors de la plage −90° à +90°.'] }
    }
    // l'inclinaison optimale est proche de la latitude sous nos climats
    const inclinaison = f('inclinaison') ?? Math.min(60, Math.abs(latitude))
    const orientation = p.orientation ?? 0
    const Kt = f('indice_clarte') ?? CLARTE[choices.clarte].valeur
    const techno = TECHNOLOGIES[choices.technologie]
    const rendement_module = f('rendement_module') ?? techno.rendement
    const NOCT = f('NOCT') ?? techno.NOCT
    const coefTemp = f('coef_temperature') ?? techno.coefTemp
    const pertes = p.pertes_systeme ?? 0.14
    const T_air = p.T_air_moyenne ?? 12

    // ---- irradiation mensuelle, jour moyen par jour moyen
    const mensuel = JOUR_MOYEN.map((jour, m) => {
      const H0 = irradiationExtraterrestre(latitude, jour)
      const horizontal = Kt * H0
      const facteur = facteurInclinaison(latitude, inclinaison, orientation, jour)
      const incline = horizontal * facteur
      // température de cellule : le module chauffe proportionnellement à
      // l'irradiation reçue, l'écart NOCT − 20 °C servant de référence
      const irradiance_moyenne = (incline * 1000) / 8 // W/m², sur 8 h d'ensoleillement utile
      const T_cellule = T_air + ((NOCT - H.T_air_NOCT) / H.irradiation_NOCT) * irradiance_moyenne
      const rendement = rendement_module * (1 + coefTemp * (T_cellule - H.T_reference_module))
      const production_j = incline * surface * rendement * (1 - pertes)
      return {
        mois: MOIS[m],
        jours: JOURS_MOIS[m],
        H0,
        horizontal,
        incline,
        T_cellule,
        rendement,
        production_j,
        production_mois: production_j * JOURS_MOIS[m],
      }
    })

    const production_an = mensuel.reduce((s, x) => s + x.production_mois, 0)
    const production_j = production_an / CONST.NOMBRE_JOUR_PAR_AN
    const irradiation_an = mensuel.reduce((s, x) => s + x.incline * x.jours, 0)
    const puissance_crete = (surface * rendement_module * 1000) / 1000 // kWc, à 1 000 W/m²
    const productible = puissance_crete > 0 ? production_an / puissance_crete : 0
    const rendement_moyen = irradiation_an * surface > 0 ? production_an / (irradiation_an * surface) : 0

    const meilleur = mensuel.reduce((a, b) => (b.production_mois > a.production_mois ? b : a))
    const pire = mensuel.reduce((a, b) => (b.production_mois < a.production_mois ? b : a))
    if (pire.production_mois > 0 && meilleur.production_mois / pire.production_mois > 8) {
      warnings.push(`Forte saisonnalité : la production de ${meilleur.mois} vaut ${(meilleur.production_mois / pire.production_mois).toFixed(0)} fois celle de ${pire.mois}.`)
    }
    if (productible < 700) {
      warnings.push(`Productible faible (${productible.toFixed(0)} kWh/kWc/an) : vérifier la latitude, l'inclinaison et la clarté du ciel.`)
    }

    return {
      results: [
        { key: 'surface', label: 'Surface installée', unit: 'm²', value: surface },
        { key: 'Pc', label: 'Puissance crête', unit: 'kWc', value: puissance_crete },
        { key: 'latitude', label: 'Latitude', unit: '°', value: latitude },
        { key: 'inclinaison', label: 'Inclinaison retenue', unit: '°', value: inclinaison },
        { key: 'Kt', label: 'Indice de clarté', unit: '-', value: Kt },
        { key: 'irr_an', label: 'Irradiation annuelle sur plan incliné', unit: 'kWh/(m²·an)', value: irradiation_an },
        { key: 'rendement_moyen', label: 'Rendement moyen du système', unit: '-', value: rendement_moyen },
        { key: 'prod_an', label: 'Production annuelle', unit: 'kWh/an', value: production_an },
        { key: 'prod_j', label: 'Production journalière moyenne', unit: 'kWh/j', value: production_j },
        { key: 'productible', label: 'Productible', unit: 'kWh/(kWc·an)', value: productible },
        { key: 'prod_max', label: `Meilleur mois — ${meilleur.mois}`, unit: 'kWh', value: meilleur.production_mois },
        { key: 'prod_min', label: `Mois le plus faible — ${pire.mois}`, unit: 'kWh', value: pire.production_mois },
        { key: 'T_cell_max', label: 'Température de cellule maximale', unit: '°C', value: Math.max(...mensuel.map((x) => x.T_cellule)) },
      ],
      pv: { mensuel, production_an, production_j, productible, puissance_crete, irradiation_an },
      electricity: {
        total: -production_j,
        produite: production_j,
        fixed: 0,
        detail: { production: -production_j },
      },
      warnings,
    }
  },
})

export { CLARTE, TECHNOLOGIES }
