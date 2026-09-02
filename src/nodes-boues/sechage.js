// ---------------------------------------------------------------------------
// Port de Sechage_thermique.cls, Sechage_bioco.cls et sechage_inos.cls.
//
// Les trois classes partagent le même cœur de calcul — évaporation d'eau,
// besoins thermiques proportionnels à l'eau évaporée, condensation des buées et
// retour des condensats en tête — et ne diffèrent que sur trois points :
//
//   Séchage thermique  quatre technologies (direct, indirect, mixte, CMV) avec
//                      chacune son couple (électricité, chaleur) par tonne
//                      d'eau évaporée ; la CMV ne consomme pas de chaleur.
//   Séchage bioco      technologie unique, 43 kWh et 880 kWh par TEE.
//   Séchage Inos       la déshydratation est intégrée au sécheur : la siccité
//                      intermédiaire se calcule par origine de boue comme dans
//                      un filtre à plateaux, et l'évaporation ne porte que sur
//                      l'écart entre cette siccité et la siccité finale.
//                      L'électricité suit une corrélation sur le tonnage traité
//                      et la part de boues digérées, non un ratio par TEE.
//
// La factory ci-dessous produit les trois nœuds ; les écarts sont portés par le
// descripteur `variante`, la structure commune n'est écrite qu'une fois.
//
// La condensation des buées existe en trois montages : indirecte (échangeur,
// la chaleur latente est récupérable), directe (lavage à l'eau froide, qui
// consomme de l'eau et produit un retour important), ou indirecte puis directe.
// L'énergie récupérée en condensation indirecte est exposée mais aucun
// consommateur ne s'en saisit tant que la gestion d'énergie n'est pas portée.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - l'ajustement automatique de la siccité de sortie pour rendre les boues
//    autocombustibles en incinération (`incineration_siccite_ajustable`) n'est
//    pas porté tant que l'incinération ne l'est pas ;
//  - le VBA ajoute la pollution soluble cumulée à chaque itération de la boucle
//    sur les types de boue, ce qui la compte autant de fois qu'il y a de types
//    présents. Le port ne la compte qu'une fois ;
//  - le second poste de pompage d'alimentation (bâche à deux étages) est porté
//    mais les deux pompes voient le même débit, comme dans le VBA.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, RET, NB_TYPES, RET_ORIGINE, ajouterRetour } from '../core/sludge.js'
import { CONST } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees, repartitionPompage, rdtPompeBoues, elecPompage } from './_commun.js'

const H = {
  // enthalpie de vaporisation de l'eau, 530 kcal/kg vers 110-115 °C
  deltaH_vap_kWh: (530 * CONST.CONVERSION_kJ_PAR_kcal) / CONST.NOMBRE_SECONDE_PAR_HEURE,
  Cp_eau_kWh: 4.18 / CONST.NOMBRE_SECONDE_PAR_HEURE, // kWh/(kg·°C)
  masse_volumique_boues: 1000, // kg/m³
  condensat_direct_Teau_inlet: 20, // °C
  condensat_direct_Teau_outlet: 40, // °C
  condensat_direct_Tair_inlet: 110, // °C, air en sortie de sécheur
  // part des pollutions que l'on retrouve dans les condensats
  ratio_DCOret_MVin: 0.015,
  ratio_NKret_NKin: 0.01,
  ratio_Pret_Pin: 0,
  ratio_retour_DBO_DCO: 0.3,
  tps_fct_pompe_alim: 5, // h/j
  capture_MS_defaut: 0.98,
  ratio_condensation_indirecte: 0.5, // part de la vapeur condensée côté indirect
}

// Siccité intermédiaire du Inos après déshydratation intégrée (g/L), par
// origine de boue — mêmes valeurs qu'un filtre à plateaux.
const CONC_INTERMEDIAIRE_INOS = {
  I_simple: 420, I_reactif: 400,
  II_forte: 360, II_moyenne: 290, II_faible_EB: 260, II_faible_ED: 260,
  II_prolongee_EB: 260, II_prolongee_ED: 260, II_MBR: 240, II_MBBR: 250, II_HybAS: 250,
  II_biostyr_C: 320, II_biostyr_N: 300, II_biostyr_NDN: 270, II_biostyr_PDN: 240,
  III_decantation: 240, III_biostyr_N: 240, III_biostyr_PDN: 240,
  codigestion_graisses: 200,
}
const CORRECTION_INOS_DIGEREES = 20 // g/L de mieux sur des boues digérées

// Couples (électricité, chaleur) par tonne d'eau évaporée, séchage thermique.
const TECHNO_THERMIQUE = {
  direct: { label: 'Sécheur direct', ratio_elec: 50, ratio_chal: 1000 },
  indirect: { label: 'Sécheur indirect', ratio_elec: 50, ratio_chal: 930 },
  mixte: { label: 'Sécheur mixte', ratio_elec: 150, ratio_chal: 1070 },
  cmv: { label: 'Compression mécanique de vapeur', ratio_elec: 150, ratio_chal: 0 },
}

const ENTREES_SECHAGE = [
  { value: 'deshydratees_1', label: 'sortie déshydratation 1' },
  { value: 'deshydratees_2', label: 'sortie déshydratation 2' },
  { value: 'deshydratees_3', label: 'sortie déshydratation 3' },
  { value: 'chaulees_1', label: 'sortie chaulage 1' },
  { value: 'digerees', label: 'boues digérées' },
  { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
]
const OPT_SOURCE = [
  { value: 'eau_HT', label: 'eau chaude haute température' },
  { value: 'vapeur', label: 'vapeur' },
  { value: 'biogaz', label: 'biogaz' },
  { value: 'combustible', label: 'combustible fossile' },
]
const OPT_CONDENSATION = [
  { value: 'indirecte', label: 'indirecte (échangeur)' },
  { value: 'directe', label: 'directe (lavage à l\'eau)' },
  { value: 'indirecte_directe', label: 'indirecte puis directe' },
]

/**
 * Fabrique un nœud de séchage. `variante` décrit les seuls points sur lesquels
 * les trois classes du classeur diffèrent.
 */
function makeSechage(variante) {
  const {
    id, label, short, vba, description, etapes,
    technologies, // null si technologie unique
    ratio_elec_fixe, ratio_chal_fixe,
    deshydratation_integree,
    sechage_total_seulement,
  } = variante

  const choices = [
    ...(technologies ? [{
      key: 'technologie', label: 'Technologie de séchage', default: 'indirect',
      options: Object.entries(technologies).map(([value, t]) => ({ value, label: t.label })),
    }] : []),
    ...(sechage_total_seulement ? [] : [{
      key: 'performance', label: 'Performance visée', default: 'total', options: [
        { value: 'partiel', label: 'séchage partiel (65 % de siccité)' },
        { value: 'total', label: 'séchage total (90 % de siccité)' },
      ],
    }]),
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: deshydratation_integree ? 'digerees' : 'deshydratees_1',
      options: [...ENTREES_SECHAGE, { value: 'aucune', label: 'aucune' }] },
    { key: 'entree_2', label: 'Boues admises (entrée 2)', default: 'aucune',
      options: [{ value: 'aucune', label: 'aucune' }, ...ENTREES_SECHAGE] },
    { key: 'source_energie', label: 'Source de chaleur', default: 'biogaz', options: OPT_SOURCE },
    { key: 'condensation', label: 'Condensation des buées', default: 'indirecte', options: OPT_CONDENSATION },
    { key: 'bache_amont', label: "Bâche d'alimentation en amont", default: 'non', options: [
      { value: 'non', label: 'non (alimentation directe)' },
      { value: 'un_etage', label: 'oui, un poste de pompage' },
      { value: 'deux_etages', label: 'oui, deux postes de pompage' },
    ] },
    ...(deshydratation_integree ? [{
      key: 'conditionnement', label: 'Conditionnement', default: 'polymere', options: [
        { value: 'polymere', label: 'polymère' },
        { value: 'polymere_FeCl3', label: 'polymère + FeCl3' },
        { value: 'chaux_FeCl3', label: 'chaux + FeCl3' },
      ],
    }] : []),
    ...(technologies ? [{
      key: 'niveau_temperature', label: 'Niveau de température du sécheur', default: 'basse', options: [
        { value: 'basse', label: 'basse température' }, { value: 'haute', label: 'haute température' },
      ],
    }] : []),
  ]

  const params = [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'outlet_siccite', label: 'Siccité des boues séchées', unit: '-', group: 'Séchage', default: undefined, hint: sechage_total_seulement ? '0,9' : '0,65 en partiel, 0,9 en total' },
    { key: 'capture_MS', label: 'Taux de capture des MS', unit: '-', group: 'Séchage', default: 0.98 },
    { key: 'ratio_elec', label: "Consommation électrique spécifique", unit: 'kWh/t eau évaporée', group: 'Séchage', default: undefined, hint: deshydratation_integree ? 'corrélation sur le tonnage traité' : 'selon la technologie' },
    { key: 'ratio_chal', label: 'Consommation thermique spécifique', unit: 'kWh/t eau évaporée', group: 'Séchage', default: undefined, hint: 'selon la technologie' },
    { key: 'besoins_thermiques', label: 'Besoin thermique total', unit: 'kWh/j', group: 'Séchage', default: undefined, hint: 'eau évaporée × consommation spécifique' },
    { key: 'ratio_condensation_indirecte', label: 'Part de vapeur condensée côté indirect', unit: '-', group: 'Condensation', default: 0.5 },
    { key: 'condensat_Teau_inlet', label: "Température de l'eau de lavage en entrée", unit: '°C', group: 'Condensation', default: 20 },
    { key: 'condensat_Teau_outlet', label: "Température de l'eau de lavage en sortie", unit: '°C', group: 'Condensation', default: 40 },
    { key: 'alimentation_pompe_nb', label: "Nombre de pompes d'alimentation", unit: 'u', group: 'Alimentation', default: 1 },
    { key: 'alimentation_P_refoulement', label: 'Pression de refoulement en alimentation', unit: 'mCE', group: 'Alimentation', default: 20 },
    { key: 'alimentation_tps_fonctionnement', label: 'Durée de fonctionnement en alimentation', unit: 'h/j', group: 'Alimentation', default: 5 },
    { key: 'alimentation_Q_unitaire', label: "Débit unitaire des pompes d'alimentation", unit: 'm³/h', group: 'Alimentation', default: undefined, hint: 'calculé si non forcé' },
    { key: 'alimentation2_pompe_nb', label: "Nombre de pompes (2e étage)", unit: 'u', group: 'Alimentation', default: 1 },
    { key: 'alimentation2_P_refoulement', label: 'Pression de refoulement (2e étage)', unit: 'mCE', group: 'Alimentation', default: 20 },
    { key: 'alimentation2_tps_fonctionnement', label: 'Durée de fonctionnement (2e étage)', unit: 'h/j', group: 'Alimentation', default: 5 },
    ...(deshydratation_integree ? [
      { key: 'concentration_intermediaire', label: 'Siccité après déshydratation intégrée', unit: 'g/L', group: 'Déshydratation intégrée', default: undefined, hint: 'moyenne des siccités de référence' },
      { key: 'dose_polymere', label: 'Dosage de polymère', unit: 'kg m.a./tMS', group: 'Déshydratation intégrée', default: 8 },
    ] : []),
  ]

  return defineSludgeNode({
    id, label, short, vba, description,
    family: 'sechage',
    multiple: true,
    maxInstances: 2,
    etapeSortie: (index) => etapes[Math.min(index, 2)],
    choices,
    params,

    compute(ctx) {
      const { table, soluble, retour, choices: ch, forced, p, index } = ctx
      const warnings = []
      const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
      const etapeOut = etapes[Math.min(index, 2)]

      const entrees = entreesDepuisChoix(ch, p, ['entree_1', 'entree_2'])
      if (!entrees.length) {
        return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
      }

      // ---- hypotheses / attribution_valeur_par_defaut
      const techno = technologies ? technologies[ch.technologie] : null
      const capture_MS = p.capture_MS ?? H.capture_MS_defaut
      const partiel = !sechage_total_seulement && ch.performance === 'partiel'
      const outlet_siccite = f('outlet_siccite') ?? (partiel ? 0.65 : 0.9)
      const ratio_chal = f('ratio_chal') ?? (techno ? techno.ratio_chal : ratio_chal_fixe)
      const ratio_condensation_indirecte = p.ratio_condensation_indirecte ?? H.ratio_condensation_indirecte
      const Teau_in = p.condensat_Teau_inlet ?? H.condensat_direct_Teau_inlet
      const Teau_out = p.condensat_Teau_outlet ?? H.condensat_direct_Teau_outlet

      // ---- lecture des entrées et écriture de l'étape aval
      const lu = lireEntrees(table, soluble, entrees)
      let inlet_MES = 0, inlet_Q = 0, inlet_MV = 0
      let inlet_DCO = 0, inlet_DBO = 0, inlet_NK = 0, inlet_Pt = 0
      let outlet_MES = 0, outlet_MV = 0
      let outlet_DCO = 0, outlet_DBO = 0, outlet_NK = 0, outlet_Pt = 0
      let MES_digerees = 0
      const soluble_kg = new Array(6).fill(0)
      const parType = {}

      for (const e of lu) {
        const { j, Q, MV_MES, ratios, sol, src } = e
        let MES = e.MES
        if (!(MES > 0)) continue
        let MV = MES * MV_MES
        let DCO = MES * ratios.DCO
        let DBO = MES * ratios.DBO
        let NK = MES * ratios.NK
        let Pt = MES * ratios.Pt
        // la pollution soluble n'est comptée qu'une fois (voir en-tête)
        for (let k = 1; k <= 5; k++) soluble_kg[k] += (sol[k] * Q) / 1000
        inlet_MES += MES; inlet_Q += Q; inlet_MV += MV
        inlet_DCO += DCO; inlet_DBO += DBO; inlet_NK += NK; inlet_Pt += Pt
        if (e.etape === ETAPE.digerees) MES_digerees += MES

        // les fines entraînées avec les buées et une part de la pollution
        // organique partent aux condensats
        MES *= capture_MS
        DCO -= H.ratio_DCOret_MVin * MV
        MV *= capture_MS
        DBO = Math.max(0, DBO - H.ratio_retour_DBO_DCO * H.ratio_DCOret_MVin * MV)
        NK -= H.ratio_NKret_NKin * NK
        Pt *= 1 - H.ratio_Pret_Pin
        outlet_MES += MES; outlet_MV += MV
        outlet_DCO += DCO; outlet_DBO += DBO; outlet_NK += NK; outlet_Pt += Pt

        if (!parType[j]) parType[j] = { MES: 0, MV: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, origine: src[P.origine] }
        const t = parType[j]
        t.MES += MES; t.MV += MV
        t.DCO += DCO; t.DBO += DBO; t.NK += NK; t.Pt += Pt
        t.flux_in += e.flux_in
      }

      if (!(inlet_MES > 0) || !(inlet_Q > 0)) {
        return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée du sécheur."] }
      }
      // l'azote soluble part intégralement aux condensats
      outlet_NK -= soluble_kg[SOL.NK]
      inlet_DCO += soluble_kg[SOL.DCO]
      inlet_NK += soluble_kg[SOL.NK]
      inlet_Pt += soluble_kg[SOL.Pt]

      const inlet_siccite = inlet_MES / (inlet_Q * H.masse_volumique_boues)
      const inlet_MV_MES = inlet_MV / inlet_MES
      const part_digeree = MES_digerees / inlet_MES

      // ---- siccité d'entrée du sécheur proprement dit
      // Le Inos déshydrate avant de sécher : l'évaporation part de la siccité
      // intermédiaire, pas de celle des boues reçues.
      let siccite_avant_sechage = inlet_siccite
      let concentration_intermediaire = 0
      let polymere_flux = 0
      let eau_deshydratation = 0
      if (deshydratation_integree) {
        let conc = 0
        for (let j = 1; j <= NB_TYPES; j++) {
          const t = parType[j]
          if (!t || !(t.MES > 0) || !(outlet_MES > 0)) continue
          const c = CONC_INTERMEDIAIRE_INOS[t.origine]
          if (c == null) {
            warnings.push(`Origine de boue « ${t.origine || 'inconnue'} » sans siccité de référence : 240 g/L retenus.`)
            conc += (t.MES / outlet_MES) * 240
          } else conc += (t.MES / outlet_MES) * c
        }
        conc += CORRECTION_INOS_DIGEREES * part_digeree
        concentration_intermediaire = f('concentration_intermediaire') ?? conc
        siccite_avant_sechage = concentration_intermediaire / H.masse_volumique_boues
        polymere_flux = ((p.dose_polymere ?? 8) / 1000) * inlet_MES
        // Le VBA ne verse pas au retour en tête l'eau retirée par la
        // déshydratation intégrée : seule la condensation des buées y figure.
        // Le comportement est reproduit, mais le volume manquant est exposé.
        eau_deshydratation = Math.max(0, inlet_Q * H.masse_volumique_boues
          - (outlet_MES * (1 - siccite_avant_sechage)) / siccite_avant_sechage - outlet_MES)
        if (eau_deshydratation > 0) {
          warnings.push(`L'eau retirée par la déshydratation intégrée (${(eau_deshydratation / 1000).toFixed(0)} m³/j) n'est pas comptée au retour en tête, conformément au classeur d'origine.`)
        }
      }

      if (outlet_siccite <= siccite_avant_sechage) {
        warnings.push(`Siccité visée (${(outlet_siccite * 100).toFixed(0)} %) inférieure ou égale à la siccité en amont du sécheur (${(siccite_avant_sechage * 100).toFixed(0)} %) : aucune évaporation.`)
      }

      // ---- eau évaporée et besoins thermiques
      const eau_evaporee = Math.max(0,
        (outlet_MES * (1 - siccite_avant_sechage)) / siccite_avant_sechage
        - (outlet_MES * (1 - outlet_siccite)) / outlet_siccite)
      const besoins_thermiques = f('besoins_thermiques') ?? (eau_evaporee / 1000) * ratio_chal

      // ---- condensation des buées
      // En condensation directe, l'air chargé de vapeur est lavé à l'eau
      // froide : le débit d'eau se déduit d'un bilan enthalpique entre la
      // chaleur latente à évacuer et l'échauffement admis sur l'eau de lavage.
      let eau_lavage = 0
      let recup_kWh = 0
      let Q_retour_kg
      const eauLavagePour = (masse) =>
        (masse * (H.deltaH_vap_kWh + H.Cp_eau_kWh * (H.condensat_direct_Tair_inlet - Teau_out)))
        / (H.Cp_eau_kWh * (Teau_out - Teau_in))

      if (ch.condensation === 'directe') {
        eau_lavage = eauLavagePour(eau_evaporee)
        Q_retour_kg = eau_evaporee + eau_lavage
      } else if (ch.condensation === 'indirecte') {
        Q_retour_kg = ratio_condensation_indirecte * eau_evaporee
        recup_kWh = Q_retour_kg * H.deltaH_vap_kWh
      } else {
        const condense_indirect = ratio_condensation_indirecte * eau_evaporee
        recup_kWh = condense_indirect * H.deltaH_vap_kWh
        eau_lavage = eauLavagePour(eau_evaporee - condense_indirect)
        Q_retour_kg = eau_evaporee + eau_lavage
      }
      const Q_retour = Q_retour_kg / 1000 // kg/j → m³/j

      // ---- écriture de l'étape aval
      const outlet_Q = outlet_siccite > 0 ? outlet_MES / outlet_siccite / H.masse_volumique_boues : 0
      for (const [jStr, t] of Object.entries(parType)) {
        const j = Number(jStr)
        const dst = table[etapeOut][j]
        dst[P.origine] = t.origine
        dst[P.MES] = t.MES
        dst[P.Q] = outlet_siccite > 0 ? t.MES / outlet_siccite / H.masse_volumique_boues : 0
        dst[P.MV_MES] = t.MES > 0 ? t.MV / t.MES : 0
        dst[P.ratio_DCO_MES] = t.MES > 0 ? t.DCO / t.MES : 0
        dst[P.ratio_DBO_MES] = t.MES > 0 ? t.DBO / t.MES : 0
        dst[P.ratio_NK_MES] = t.MES > 0 ? Math.max(0, t.NK) / t.MES : 0
        dst[P.ratio_Pt_MES] = t.MES > 0 ? t.Pt / t.MES : 0
        dst[P.flux_in] = t.flux_in
        dst[P.verif_flux] = t.flux_in
        // les boues séchées ne portent plus d'eau interstitielle
        for (let k = 1; k <= 5; k++) soluble[etapeOut][j][k] = 0
      }
      for (const e of lu) e.src[P.verif_flux] -= e.flux_in

      // ---- condensats retournés en tête
      // Condensats de buées : volume faible, non éligibles au traitement.
      const vecteur = []
      vecteur[RET.Q] = Q_retour
      vecteur[RET.MES] = inlet_MES - outlet_MES
      vecteur[RET.DCO] = inlet_DCO - outlet_DCO
      vecteur[RET.DBO] = inlet_DBO - outlet_DBO
      vecteur[RET.NK] = inlet_NK - outlet_NK
      vecteur[RET.NH4] = soluble_kg[SOL.NK]
      vecteur[RET.Pt] = inlet_Pt - outlet_Pt
      if (ctx.retours) ajouterRetour(ctx.retours, RET_ORIGINE.autres, vecteur, null)
      else for (let i = 1; i <= 8; i++) retour[i] += vecteur[i] || 0

      // ---- électricité
      let ratio_elec
      if (deshydratation_integree) {
        // Corrélation Inos : la consommation par tonne de MS décroît avec le
        // tonnage traité, et dépend du conditionnement pour des boues fraîches.
        const t_MS = inlet_MES / 1000
        if (part_digeree > 0) {
          ratio_elec = 143.24 - 0.0135 * t_MS
        } else if (ch.conditionnement === 'polymere') {
          ratio_elec = 249.84 - 0.0044 * t_MS
        } else if (ch.conditionnement === 'polymere_FeCl3') {
          ratio_elec = 180.84 - 0.0067 * t_MS
        } else {
          ratio_elec = 178.72 - 0.0141 * t_MS
        }
      } else {
        ratio_elec = techno ? techno.ratio_elec : ratio_elec_fixe
      }
      const rf = f('ratio_elec')
      if (rf != null) ratio_elec = rf
      // le Inos rapporte sa consommation au tonnage de MS, les autres à l'eau évaporée
      const electricite_specifique = deshydratation_integree
        ? (inlet_MES / 1000) * ratio_elec
        : (eau_evaporee / 1000) * ratio_elec

      let electricite_alimentation = 0
      const postes = ch.bache_amont === 'non' ? 0 : ch.bache_amont === 'un_etage' ? 1 : 2
      for (let poste = 1; poste <= postes; poste++) {
        const prefixe = poste === 1 ? 'alimentation' : 'alimentation2'
        const r = repartitionPompage(inlet_Q, p[`${prefixe}_pompe_nb`], p[`${prefixe}_tps_fonctionnement`], forced, prefixe, H.tps_fct_pompe_alim)
        if (r.incoherence) warnings.push(`Incohérence sur le pompage d'alimentation${poste === 2 ? ' (2e étage)' : ''}.`)
        const rdt = rdtPompeBoues(r.Qu, p[`${prefixe}_P_refoulement`])
        electricite_alimentation += elecPompage(r.Qu, r.nb, r.tps, p[`${prefixe}_P_refoulement`], rdt)
      }

      const total = electricite_specifique + electricite_alimentation

      const reactifs = {}
      if (polymere_flux > 0) reactifs.polymere = polymere_flux

      return {
        etapeSortie: etapeOut,
        results: [
          { key: 'in_MES', label: 'MS en entrée', unit: 'kg/j', value: inlet_MES },
          { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: inlet_Q },
          { key: 'in_siccite', label: 'Siccité en entrée', unit: '%', value: inlet_siccite * 100 },
          { key: 'in_MV_MES', label: 'MV/MS en entrée', unit: '-', value: inlet_MV_MES },
          ...(deshydratation_integree ? [
            { key: 'conc_interm', label: 'Siccité après déshydratation intégrée', unit: 'g/L', value: concentration_intermediaire },
            { key: 'eau_deshy', label: 'Eau retirée par la déshydratation intégrée', unit: 'm³/j', value: eau_deshydratation / 1000 },
            { key: 'poly', label: 'Polymère (matière active)', unit: 'kg/j', value: polymere_flux },
          ] : []),
          { key: 'out_MES', label: 'MS des boues séchées', unit: 'kg/j', value: outlet_MES },
          { key: 'out_Q', label: 'Débit de boues séchées', unit: 'm³/j', value: outlet_Q },
          { key: 'out_siccite', label: 'Siccité des boues séchées', unit: '%', value: outlet_siccite * 100 },
          { key: 'out_t', label: 'Tonnage de boues séchées', unit: 't/j', value: (outlet_MES / outlet_siccite) / 1000 },
          { key: 'evap', label: 'Eau évaporée', unit: 'kg/j', value: eau_evaporee },
          { key: 'evap_t', label: 'Eau évaporée', unit: 't/j', value: eau_evaporee / 1000 },
          { key: 'therm', label: `Besoin thermique (${ch.source_energie.replace('_', ' ')})`, unit: 'kWh/j', value: besoins_thermiques },
          { key: 'therm_spec', label: 'Consommation thermique spécifique', unit: 'kWh/t eau évaporée', value: ratio_chal },
          { key: 'elec_spec', label: 'Consommation électrique spécifique', unit: deshydratation_integree ? 'kWh/tMS' : 'kWh/t eau évaporée', value: ratio_elec },
          ...(recup_kWh > 0 ? [{ key: 'recup', label: 'Chaleur récupérable à la condensation', unit: 'kWh/j', value: recup_kWh }] : []),
          ...(eau_lavage > 0 ? [{ key: 'lavage', label: 'Eau de lavage des buées', unit: 'm³/j', value: eau_lavage / 1000 }] : []),
          { key: 'ret_Q', label: 'Condensats retournés en tête', unit: 'm³/j', value: Q_retour },
          { key: 'ret_MES', label: 'MES au retour en tête', unit: 'kg/j', value: inlet_MES - outlet_MES },
          { key: 'ret_NK', label: 'NK au retour en tête', unit: 'kg/j', value: inlet_NK - outlet_NK },
          { key: 'ret_DCO', label: 'DCO au retour en tête', unit: 'kg/j', value: inlet_DCO - outlet_DCO },
        ],
        reactifs,
        energie: {
          besoin_thermique_kWhj: besoins_thermiques,
          source: ch.source_energie,
          recuperable_kWhj: recup_kWh,
        },
        electricity: {
          total,
          fixed: 0,
          detail: { procede: electricite_specifique, alimentation: electricite_alimentation },
        },
        warnings,
      }
    },
  })
}

export const sechageThermique = makeSechage({
  id: 'sechage-thermique',
  label: 'Séchage thermique',
  short: 'Séchage',
  vba: 'Sechage_thermique.cls',
  etapes: ETAPE.sechees,
  technologies: TECHNO_THERMIQUE,
  description:
    "Séchage thermique des boues déshydratées, direct, indirect, mixte ou par compression mécanique de vapeur. L'eau évaporée détermine les besoins en chaleur ; les buées sont condensées par échangeur, par lavage à l'eau ou par les deux, et les condensats repartent en tête de station.",
})

export const sechageBioco = makeSechage({
  id: 'sechage-bioco',
  label: 'Séchage Bioco',
  short: 'Bioco',
  vba: 'Sechage_bioco.cls',
  etapes: ETAPE.sechees_bioco,
  technologies: null,
  ratio_elec_fixe: 43,
  ratio_chal_fixe: 880,
  description:
    "Séchage Bioco : technologie unique à 43 kWh électriques et 880 kWh thermiques par tonne d'eau évaporée, moins gourmande en chaleur qu'un sécheur thermique classique.",
})

export const sechageInos = makeSechage({
  id: 'sechage-inos',
  label: 'Séchage Inos',
  short: 'Inos',
  vba: 'sechage_inos.cls',
  etapes: ETAPE.sechees_inos,
  technologies: null,
  ratio_chal_fixe: 1100,
  deshydratation_integree: true,
  sechage_total_seulement: true,
  description:
    "Séchage Inos, qui intègre la déshydratation au sécheur : les boues sont reçues liquides, la siccité intermédiaire est calculée par origine comme sur un filtre à plateaux, et l'évaporation ne porte que sur l'écart jusqu'à 90 % de siccité.",
})
