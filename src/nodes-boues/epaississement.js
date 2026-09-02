// ---------------------------------------------------------------------------
// Port de z_Epaississement.cls — épaississement des boues.
//
// Quatre technologies : gravitaire (sans polymère), centrifuge, table
// d'égouttage et flottation. Le procédé lit une ou deux étapes amont de la file
// boues, applique un taux de capture des MES, et recalcule la siccité de sortie
// comme la moyenne des concentrations de référence propres à chaque origine de
// boue, pondérée par les MES.
//
// La fraction non capturée, augmentée de la pollution soluble entraînée par
// l'eau interstitielle rejetée, part au retour en tête de station.
//
// Le classeur autorise jusqu'à cinq épaississeurs en parallèle ; l'instance est
// repérée par `ctx.index`, qui détermine l'étape de sortie (boues_epaissies).
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, RET, NB_TYPES, TYPE, RET_ORIGINE, ajouterRetour } from '../core/sludge.js'
import { CONST } from '../core/hypotheses.js'
import { repartitionPompage, rdtPompeBoues, elecPompage, H_POMPE } from './_commun.js'

const TYPES_EPAISSISSEUR = [
  { value: 'gravitaire', label: 'Gravitaire' },
  { value: 'centrifuge', label: 'Centrifugeuse' },
  { value: 'table_egouttage', label: "Table d'égouttage" },
  { value: 'flottation', label: 'Flottation' },
]

// Concentration de sortie attendue (g/L) par technologie et par origine de boue
// (outlet_concentration_reference du VBA).
const CONC_REF = {
  gravitaire: {
    I_simple: 60, I_reactif: 53,
    II_forte: 45, II_moyenne: 28, II_faible_EB: 20, II_faible_ED: 20,
    II_prolongee_EB: 25, II_prolongee_ED: 25, II_MBR: 20, II_MBBR: 20, II_HybAS: 25,
    II_biostyr_C: 55, II_biostyr_N: 45, II_biostyr_NDN: 40, II_biostyr_PDN: 35,
    III_decantation: 19, III_biostyr_N: 19, III_biostyr_PDN: 19,
    codigestion_graisses: 5,
  },
  centrifuge: {
    I_simple: 71, I_reactif: 67,
    II_forte: 62, II_moyenne: 57, II_faible_EB: 40, II_faible_ED: 40,
    II_prolongee_EB: 50, II_prolongee_ED: 50, II_MBR: 39, II_MBBR: 50, II_HybAS: 50,
    II_biostyr_C: 62, II_biostyr_N: 57, II_biostyr_NDN: 52, II_biostyr_PDN: 48,
    III_decantation: 34, III_biostyr_N: 34, III_biostyr_PDN: 34,
    codigestion_graisses: 29,
  },
  table_egouttage: {
    I_simple: 75, I_reactif: 70,
    II_forte: 65, II_moyenne: 60, II_faible_EB: 45, II_faible_ED: 45,
    II_prolongee_EB: 53, II_prolongee_ED: 53, II_MBR: 41, II_MBBR: 45, II_HybAS: 45,
    II_biostyr_C: 65, II_biostyr_N: 60, II_biostyr_NDN: 55, II_biostyr_PDN: 50,
    III_decantation: 36, III_biostyr_N: 36, III_biostyr_PDN: 36,
    codigestion_graisses: 30,
  },
  flottation: {
    I_simple: 55, I_reactif: 50,
    II_forte: 53, II_moyenne: 47, II_faible_EB: 45, II_faible_ED: 45,
    II_prolongee_EB: 45, II_prolongee_ED: 45, II_MBR: 38, II_MBBR: 45, II_HybAS: 45,
    II_biostyr_C: 53, II_biostyr_N: 48, II_biostyr_NDN: 48, II_biostyr_PDN: 45,
    III_decantation: 30, III_biostyr_N: 30, III_biostyr_PDN: 30,
    codigestion_graisses: 60,
  },
}
const CONC_DEFAUT = { gravitaire: 50, centrifuge: 40, table_egouttage: 70, flottation: 40 }
const CAPTURE_DEFAUT = { gravitaire: 0.95, centrifuge: 0.97, table_egouttage: 0.9, flottation: 0.98 }
const DOSE_POLYMERE = { gravitaire: 0, centrifuge: 2, table_egouttage: 5, flottation: 2 }

const H = {
  polymere_concentration: 1.5, // g/L de polymère injecté
  tps_fct_pompe: 5, // h/j
  critere_incoherence: 0.1,
  rdt_pompe: 0.4,
  conso_polymere: 0.76, // kWh/kg de polymère
  // flottation (Predimboo)
  ratio_recirc_eau_blanche_L_g: 0.4,
  ratio_air_eau_blanche: 0.14,
  P_eau_blanche_mCE: 50,
  conso_pompe_eau_blanche: 2.725 / 0.6, // Wh/(m³·mCE)
  conso_compresseur: 2, // Wh/(Nm³·mCE)
  // gravitaire
  grav_puissance_fixe: 0.1, // Wh/m³ sur le nominal
  grav_puissance_debit: 1, // Wh/m³ sur le réel
  conso_centrifuge: 0.9, // kWh/m³
  conso_table: 4, // kWh/tMS
}

export default defineSludgeNode({
  id: 'epaississement',
  label: 'Épaississement',
  short: 'Épaiss.',
  family: 'epaississement',
  vba: 'z_Epaississement.cls',
  multiple: true,
  maxInstances: 5,
  etapeSortie: (index) => ETAPE.epaissies[Math.min(index, 5)],
  description:
    "Épaississement des boues par gravité, centrifugation, table d'égouttage ou flottation. La siccité de sortie est déduite des concentrations de référence propres à chaque origine de boue ; les fines non capturées et la pollution soluble entraînée repartent en tête de station.",
  choices: [
    { key: 'type', label: "Type d'épaississeur", default: 'gravitaire', options: TYPES_EPAISSISSEUR },
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: 'toutes', options: [
      { value: 'toutes', label: 'toutes les boues extraites' },
      { value: 'I', label: 'boues primaires seules' },
      { value: 'II', label: 'boues secondaires seules' },
      { value: 'III', label: 'boues tertiaires seules' },
      { value: 'graisses', label: 'graisses seules' },
      { value: 'aucune', label: 'aucune' },
    ] },
    { key: 'entree_2', label: 'Boues admises (entrée 2)', default: 'aucune', options: [
      { value: 'aucune', label: 'aucune' },
      { value: 'I', label: 'boues primaires' },
      { value: 'II', label: 'boues secondaires' },
      { value: 'III', label: 'boues tertiaires' },
      { value: 'graisses', label: 'graisses' },
    ] },
    { key: 'bache_alimentation', label: "Bâche d'alimentation en amont", default: 'oui', options: [
      { value: 'oui', label: 'oui (pompage des boues)' }, { value: 'non', label: 'non (alimentation gravitaire)' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'alimentation_pompe_nb', label: "Nombre de pompes d'alimentation", unit: 'u', group: 'Alimentation', default: 1 },
    { key: 'alimentation_P_refoulement', label: 'Pression de refoulement en alimentation', unit: 'mCE', group: 'Alimentation', default: 10 },
    { key: 'alimentation_tps_fonctionnement', label: 'Durée de fonctionnement en alimentation', unit: 'h/j', group: 'Alimentation', default: 5 },
    { key: 'alimentation_Q_unitaire', label: 'Débit unitaire des pompes d\'alimentation', unit: 'm³/h', group: 'Alimentation', default: undefined, hint: 'calculé si non forcé' },
    { key: 'alimentation_pompe_rdt', label: "Rendement global des pompes d'alimentation", unit: '-', group: 'Alimentation', default: undefined, hint: 'machine 0,4 × moteur' },
    { key: 'capture_MES', label: 'Taux de capture des MES', unit: '-', group: 'Séparation', default: undefined, hint: 'selon la technologie' },
    { key: 'outlet_concentration', label: 'Siccité des boues épaissies', unit: 'g/L', group: 'Séparation', default: undefined, hint: 'moyenne des concentrations de référence' },
    { key: 'dose_polymere', label: 'Dosage de polymère', unit: 'kg m.a./tMS', group: 'Réactifs', default: undefined, hint: 'selon la technologie, nul en gravitaire' },
    { key: 'extraction_pompe_nb', label: "Nombre de pompes d'extraction", unit: 'u', group: 'Extraction', default: 1 },
    { key: 'extraction_P_refoulement', label: 'Pression de refoulement en extraction', unit: 'mCE', group: 'Extraction', default: 10 },
    { key: 'extraction_tps_fonctionnement', label: 'Durée de fonctionnement en extraction', unit: 'h/j', group: 'Extraction', default: 5 },
    { key: 'extraction_Q_unitaire', label: "Débit unitaire des pompes d'extraction", unit: 'm³/h', group: 'Extraction', default: undefined, hint: 'calculé si non forcé' },
    { key: 'extraction_pompe_rdt', label: "Rendement global des pompes d'extraction", unit: '-', group: 'Extraction', default: undefined, hint: 'machine 0,4 × moteur' },
  ],

  compute(ctx) {
    const { table, soluble, retour, choices, forced, p, index } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const type = choices.type
    const etapeOut = ETAPE.epaissies[Math.min(index, 5)]

    // ---- lecture_choix : quelles boues sont admises, dans quelle proportion
    const typeDe = (v) => ({ I: TYPE.I, II: TYPE.II, III: TYPE.III, graisses: TYPE.graisses })[v]
    const entrees = []
    for (const [cle, ratioCle] of [['entree_1', 'ratio_admis_1'], ['entree_2', 'ratio_admis_2']]) {
      const v = choices[cle]
      if (!v || v === 'aucune') continue
      entrees.push({ etape: ETAPE.inlet, types: v === 'toutes' ? null : [typeDe(v)], ratio: p[ratioCle] ?? 1 })
    }
    if (!entrees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucune boue admise : configurer au moins une entrée."] }
    }

    // ---- attribution_valeur_par_defaut
    const capture_MES = f('capture_MES') ?? CAPTURE_DEFAUT[type]
    const dose_polymere = f('dose_polymere') ?? DOSE_POLYMERE[type]
    // valeur d'amorçage : le VBA calcule Q de sortie avant de connaître la
    // siccité finale, puis la recalcule. Reproduit à l'identique.
    let outlet_concentration = CONC_DEFAUT[type]

    // ---- première passe : capture des MES et alimentation de l'étape aval
    let inlet_MES = 0, inlet_Q = 0, inlet_DCO = 0, inlet_DBO = 0, inlet_NK = 0, inlet_Pt = 0
    let outlet_MES = 0, outlet_Q = 0, outlet_DCO = 0, outlet_DBO = 0, outlet_NK = 0, outlet_Pt = 0
    const soluble_kg = new Array(6).fill(0)

    for (const e of entrees) {
      const types = e.types ?? Array.from({ length: NB_TYPES }, (_, i) => i + 1)
      for (const j of types) {
        const src = table[e.etape][j]
        if (!(src[P.MES] > 0)) continue
        let MES = e.ratio * src[P.MES]
        let Q = e.ratio * src[P.Q]
        let DCO = src[P.ratio_DCO_MES] * MES
        let DBO = src[P.ratio_DBO_MES] * MES
        let NK = src[P.ratio_NK_MES] * MES
        let Pt = src[P.ratio_Pt_MES] * MES
        inlet_MES += MES; inlet_Q += Q; inlet_DCO += DCO; inlet_DBO += DBO; inlet_NK += NK; inlet_Pt += Pt
        for (let k = 1; k <= 5; k++) soluble_kg[k] += (soluble[e.etape][j][k] * Q) / 1000
        // séparation : on ne retient que la part capturée
        MES *= capture_MES; DCO *= capture_MES; DBO *= capture_MES; NK *= capture_MES; Pt *= capture_MES
        Q = outlet_concentration > 0 ? MES / outlet_concentration : 0
        outlet_MES += MES; outlet_Q += Q; outlet_DCO += DCO; outlet_DBO += DBO; outlet_NK += NK; outlet_Pt += Pt
        const dst = table[etapeOut][j]
        dst[P.MES] += MES
        dst[P.Q] += Q
        dst[P.MV_MES] = src[P.MV_MES]
        dst[P.ratio_DCO_MES] = MES > 0 ? DCO / MES : 0
        dst[P.ratio_DBO_MES] = MES > 0 ? DBO / MES : 0
        dst[P.ratio_NK_MES] = MES > 0 ? NK / MES : 0
        dst[P.ratio_Pt_MES] = MES > 0 ? Pt / MES : 0
        dst[P.flux_in] += src[P.flux_in] * e.ratio
        dst[P.verif_flux] = dst[P.flux_in]
        // le flux consommé n'est plus disponible en amont (bilan matière)
        src[P.verif_flux] -= src[P.flux_in] * e.ratio
      }
    }

    const inlet_MS = soluble_kg[SOL.MS_soluble] + inlet_MES
    if (!(inlet_MS > 0) || !(inlet_Q > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée de l'épaississeur."] }
    }
    const inlet_concentration = inlet_MS / inlet_Q

    // ---- polymère
    let polymere_flux = 0
    let polymere_Q = 0
    if (type !== 'gravitaire') {
      polymere_flux = (dose_polymere / 1000) * inlet_MS
      polymere_Q = polymere_flux / H.polymere_concentration
    }
    const Q_traite = inlet_Q + polymere_Q

    // ---- siccité de sortie : moyenne des concentrations de référence,
    // pondérée par les MES de chaque origine
    const ref = CONC_REF[type]
    let conc = 0
    for (let j = 1; j <= NB_TYPES; j++) {
      const MES = table[etapeOut][j][P.MES]
      if (!(MES > 0) || !(outlet_MES > 0)) continue
      const origine = table[etapeOut][j][P.origine]
      const c = ref[origine]
      if (c == null) {
        warnings.push(`Origine de boue « ${origine || 'inconnue'} » sans concentration de référence : valeur par défaut ${CONC_DEFAUT[type]} g/L retenue.`)
        conc += (MES / outlet_MES) * CONC_DEFAUT[type]
      } else conc += (MES / outlet_MES) * c
    }
    outlet_concentration = f('outlet_concentration') ?? conc
    if (outlet_concentration < inlet_concentration) {
      // un épaississeur ne peut pas diluer : plancher à 1 % au-dessus de l'entrée
      outlet_concentration = inlet_concentration * 1.01
      warnings.push("Siccité de sortie inférieure à celle d'entrée : plafonnée à 1 % au-dessus de l'entrée.")
    }

    // ---- reprise des débits de sortie avec la siccité définitive
    outlet_Q = 0
    for (let j = 1; j <= NB_TYPES; j++) {
      const MES = table[etapeOut][j][P.MES]
      const Q = outlet_concentration > 0 ? MES / outlet_concentration : 0
      table[etapeOut][j][P.Q] = Q
      outlet_Q += Q
    }

    // ---- retours en tête : eau séparée + pollution soluble entraînée
    const Q_retour = inlet_Q + polymere_Q - outlet_Q
    // Surnageant d'épaississement : peu chargé en azote, il n'est pas éligible
    // au traitement des retours et rejoint directement la tête de station.
    const vecteur = []
    vecteur[RET.Q] = Q_retour
    vecteur[RET.MES] = inlet_MES - outlet_MES
    vecteur[RET.DCO] = inlet_DCO - outlet_DCO
    vecteur[RET.DBO] = inlet_DBO - outlet_DBO
    vecteur[RET.NK] = inlet_NK - outlet_NK
    vecteur[RET.Pt] = inlet_Pt - outlet_Pt

    // la pollution soluble se répartit entre le surnageant et l'eau restée dans
    // les boues, au prorata des débits
    const part = (kg) => (Q_retour > 0 ? kg / (1 + outlet_Q / Q_retour) : 0)
    const sol_DCO = part(soluble_kg[SOL.DCO])
    const sol_NK = part(soluble_kg[SOL.NK])
    const sol_Pt = part(soluble_kg[SOL.Pt])
    const sol_MS = part(soluble_kg[SOL.MS_soluble])
    const sol_MV = part(soluble_kg[SOL.MV_soluble])
    vecteur[RET.DCO] += sol_DCO
    vecteur[RET.NK] += sol_NK
    vecteur[RET.Pt] += sol_Pt
    const vecteur_soluble = []
    vecteur_soluble[RET.Q] = Q_retour
    vecteur_soluble[RET.DCO] = sol_DCO
    vecteur_soluble[RET.NK] = sol_NK
    vecteur_soluble[RET.Pt] = sol_Pt
    if (ctx.retours) ajouterRetour(ctx.retours, RET_ORIGINE.autres, vecteur, vecteur_soluble)
    else for (let i = 1; i <= 8; i++) retour[i] += vecteur[i] || 0
    for (let j = 1; j <= NB_TYPES; j++) {
      const s = soluble[etapeOut][j]
      s[SOL.DCO] = outlet_Q > 0 ? ((soluble_kg[SOL.DCO] - sol_DCO) / outlet_Q) * 1000 : 0
      s[SOL.NK] = outlet_Q > 0 ? ((soluble_kg[SOL.NK] - sol_NK) / outlet_Q) * 1000 : 0
      s[SOL.Pt] = outlet_Q > 0 ? ((soluble_kg[SOL.Pt] - sol_Pt) / outlet_Q) * 1000 : 0
      s[SOL.MS_soluble] = outlet_Q > 0 ? ((soluble_kg[SOL.MS_soluble] - sol_MS) / outlet_Q) * 1000 : 0
      s[SOL.MV_soluble] = outlet_Q > 0 ? ((soluble_kg[SOL.MV_soluble] - sol_MV) / outlet_Q) * 1000 : 0
    }
    const outlet_MS = outlet_MES + (soluble_kg[SOL.MS_soluble] - sol_MS)

    // ---- calcul_consommation_electrique
    let electricite_alimentation = 0
    let alim = { nb: 0, tps: 0, Qu: 0 }
    if (choices.bache_alimentation === 'oui') {
      alim = repartitionPompage(inlet_Q, p.alimentation_pompe_nb, p.alimentation_tps_fonctionnement, forced, 'alimentation', H.tps_fct_pompe)
      if (alim.incoherence) warnings.push("Incohérence sur le pompage d'alimentation (débit unitaire × nombre × durée ≠ débit à traiter).")
      const rdt = f('alimentation_pompe_rdt') ?? rdtPompeBoues(alim.Qu, p.alimentation_P_refoulement)
      electricite_alimentation = elecPompage(alim.Qu, alim.nb, alim.tps, p.alimentation_P_refoulement, rdt)
    }

    let electricite_specifique = 0
    let eau_blanche_Q = 0
    if (type === 'gravitaire') {
      electricite_specifique = (H.grav_puissance_fixe * Q_traite + H.grav_puissance_debit * Q_traite) / 1000
    } else if (type === 'centrifuge') {
      electricite_specifique = H.conso_centrifuge * Q_traite
    } else if (type === 'table_egouttage') {
      electricite_specifique = (H.conso_table * outlet_MES) / 1000
    } else {
      // flottation : pompe d'eau blanche + compresseur d'air (Predimboo)
      const taux = H.ratio_recirc_eau_blanche_L_g * inlet_concentration
      eau_blanche_Q = taux * inlet_Q
      const air = H.ratio_air_eau_blanche * eau_blanche_Q
      electricite_specifique = (eau_blanche_Q * H.P_eau_blanche_mCE * H.conso_pompe_eau_blanche) / 1000
        + (air * H.P_eau_blanche_mCE * H.conso_compresseur) / 1000
    }
    const electricite_polymere = type === 'gravitaire' ? 0 : H.conso_polymere * polymere_flux

    const extr = repartitionPompage(outlet_Q, p.extraction_pompe_nb, p.extraction_tps_fonctionnement, forced, 'extraction', H.tps_fct_pompe)
    if (extr.incoherence) warnings.push("Incohérence sur le pompage d'extraction.")
    const rdt_extr = f('extraction_pompe_rdt') ?? rdtPompeBoues(extr.Qu, p.extraction_P_refoulement)
    const electricite_extraction = elecPompage(extr.Qu, extr.nb, extr.tps, p.extraction_P_refoulement, rdt_extr)

    const total = electricite_alimentation + electricite_extraction + electricite_specifique + electricite_polymere

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MES en entrée', unit: 'kg/j', value: inlet_MES },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: inlet_Q },
        { key: 'in_conc', label: 'Siccité en entrée', unit: 'g/L', value: inlet_concentration },
        { key: 'capture', label: 'Taux de capture des MES', unit: '-', value: capture_MES },
        { key: 'out_MES', label: 'MES épaissies', unit: 'kg/j', value: outlet_MES },
        { key: 'out_MS', label: 'MS épaissies (avec soluble)', unit: 'kg/j', value: outlet_MS },
        { key: 'out_Q', label: 'Débit de boues épaissies', unit: 'm³/j', value: outlet_Q },
        { key: 'out_conc', label: 'Siccité des boues épaissies', unit: 'g/L', value: outlet_concentration },
        ...(type !== 'gravitaire' ? [
          { key: 'poly', label: 'Polymère (matière active)', unit: 'kg/j', value: polymere_flux },
          { key: 'poly_Q', label: 'Solution de polymère injectée', unit: 'm³/j', value: polymere_Q },
        ] : []),
        ...(type === 'flottation' ? [{ key: 'eau_blanche', label: "Débit d'eau blanche", unit: 'm³/j', value: eau_blanche_Q }] : []),
        { key: 'ret_Q', label: 'Retour en tête', unit: 'm³/j', value: Q_retour },
        { key: 'ret_MES', label: 'MES au retour en tête', unit: 'kg/j', value: inlet_MES - outlet_MES },
        { key: 'ret_DCO', label: 'DCO au retour en tête', unit: 'kg/j', value: inlet_DCO - outlet_DCO + sol_DCO },
        { key: 'ret_NK', label: 'NK au retour en tête', unit: 'kg/j', value: inlet_NK - outlet_NK + sol_NK },
      ],
      reactifs: type !== 'gravitaire' ? { polymere: polymere_flux } : {},
      electricity: {
        total,
        fixed: type === 'gravitaire' ? electricite_specifique : 0,
        detail: {
          alimentation: electricite_alimentation,
          procede: electricite_specifique,
          polymere: electricite_polymere,
          extraction: electricite_extraction,
        },
      },
      warnings,
    }
  },
})
