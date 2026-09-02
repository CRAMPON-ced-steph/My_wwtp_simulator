// ---------------------------------------------------------------------------
// Port de z_Deshydratation.cls — déshydratation mécanique des boues.
//
// Trois technologies : centrifugeuse, filtre à plateaux et filtre à bandes.
// Comme à l'épaississement, la siccité de sortie est la moyenne des siccités de
// référence propres à chaque origine de boue, pondérée par les MES ; s'y ajoute
// un bonus lorsque les boues ont été digérées au préalable (environ 20 g/L,
// proratisé à la part de boues digérées dans l'alimentation).
//
// Le dosage de polymère n'est pas une constante par technologie mais dépend de
// l'origine de chaque boue et du fait qu'elle ait été digérée ou non : une boue
// primaire se conditionne à 5 kg/tMS en centrifugation, une boue d'aération
// prolongée à 12. Le filtre à plateaux se conditionne à la chaux et au chlorure
// ferrique plutôt qu'au polymère.
//
// Le classeur autorise jusqu'à trois déshydratations en parallèle ; l'instance
// est repérée par `ctx.index`.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - le circuit technosable (filtre à plateaux en aval d'un Athos, siccité
//    600 g/L, retour dédié) n'est pas porté tant qu'Athos ne l'est pas ;
//  - le VBA distingue un vecteur `retour_digestion` du retour général, pour
//    isoler les retours chargés en azote issus de la digestion. Le port ne tient
//    qu'un vecteur ; la part digérée est exposée en résultat pour permettre le
//    dimensionnement d'un traitement des retours.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, RET, NB_TYPES, RET_ORIGINE, ajouterRetour } from '../core/sludge.js'
import { CONST } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees, repartitionPompage, rdtPompeBoues, elecPompage } from './_commun.js'

// Siccité de sortie attendue (g/L) par technologie et par origine de boue.
const CONC_REF = {
  centrifuge: {
    I_simple: 326, I_reactif: 308,
    II_forte: 282, II_moyenne: 220, II_faible_EB: 200, II_faible_ED: 195,
    II_prolongee_EB: 200, II_prolongee_ED: 195, II_MBR: 180, II_MBBR: 190, II_HybAS: 190,
    II_biostyr_C: 246, II_biostyr_N: 229, II_biostyr_NDN: 202, II_biostyr_PDN: 185,
    III_decantation: 185, III_biostyr_N: 185, III_biostyr_PDN: 185,
    codigestion_graisses: 150,
  },
  filtre_plateaux: {
    I_simple: 420, I_reactif: 400,
    II_forte: 360, II_moyenne: 290, II_faible_EB: 260, II_faible_ED: 260,
    II_prolongee_EB: 260, II_prolongee_ED: 260, II_MBR: 240, II_MBBR: 250, II_HybAS: 250,
    II_biostyr_C: 320, II_biostyr_N: 300, II_biostyr_NDN: 270, II_biostyr_PDN: 240,
    III_decantation: 240, III_biostyr_N: 240, III_biostyr_PDN: 240,
    codigestion_graisses: 200,
  },
  filtre_bandes: {
    I_simple: 278, I_reactif: 263,
    II_forte: 240, II_moyenne: 188, II_faible_EB: 170, II_faible_ED: 170,
    II_prolongee_EB: 170, II_prolongee_ED: 170, II_MBR: 154, II_MBBR: 170, II_HybAS: 170,
    II_biostyr_C: 210, II_biostyr_N: 195, II_biostyr_NDN: 173, II_biostyr_PDN: 158,
    III_decantation: 158, III_biostyr_N: 158, III_biostyr_PDN: 158,
    codigestion_graisses: 128,
  },
}
// Dosage de polymère (kg de matière active par tonne de MS), boues non digérées.
const DOSE_POLY = {
  centrifuge: {
    I_simple: 5, I_reactif: 6,
    II_forte: 7, II_moyenne: 8, II_faible_EB: 12, II_faible_ED: 11.5,
    II_prolongee_EB: 10.5, II_prolongee_ED: 12, II_MBR: 12, II_MBBR: 10.5, II_HybAS: 11.5,
    II_biostyr_C: 8.5, II_biostyr_N: 9.5, II_biostyr_NDN: 10.5, II_biostyr_PDN: 11.5,
    III_decantation: 8.5, III_biostyr_N: 9.5, III_biostyr_PDN: 11.5,
    codigestion_graisses: 16.5,
  },
  filtre_plateaux: {
    I_simple: 3.6, I_reactif: 3.8,
    II_forte: 4.3, II_moyenne: 7.3, II_faible_EB: 7.9, II_faible_ED: 8.1,
    II_prolongee_EB: 7.9, II_prolongee_ED: 8.1, II_MBR: 7.3, II_MBBR: 7.5, II_HybAS: 7.9,
    II_biostyr_C: 5, II_biostyr_N: 5.5, II_biostyr_NDN: 6.4, II_biostyr_PDN: 7.1,
    III_decantation: 7.1, III_biostyr_N: 5.5, III_biostyr_PDN: 7.1,
    codigestion_graisses: 15,
  },
  filtre_bandes: {
    I_simple: 3, I_reactif: 3.9,
    II_forte: 6.5, II_moyenne: 7.5, II_faible_EB: 11.5, II_faible_ED: 11.5,
    II_prolongee_EB: 10, II_prolongee_ED: 11.5, II_MBR: 8, II_MBBR: 10.5, II_HybAS: 11.5,
    II_biostyr_C: 8, II_biostyr_N: 9, II_biostyr_NDN: 10, II_biostyr_PDN: 11.5,
    III_decantation: 8, III_biostyr_N: 9, III_biostyr_PDN: 11.5,
    codigestion_graisses: 16,
  },
}
// Dosage de polymère après digestion : les boues digérées se conditionnent plus
// difficilement, et l'écart entre origines s'efface.
const DOSE_POLY_DIG = {
  centrifuge: {
    I_simple: 6.5, I_reactif: 6.5,
    II_forte: 9.5, II_moyenne: 9.5, II_faible_EB: 9.5, II_faible_ED: 9.5,
    II_prolongee_EB: 9.5, II_prolongee_ED: 9.5, II_MBR: 9.5, II_MBBR: 9.5, II_HybAS: 9.5,
    II_biostyr_C: 9.5, II_biostyr_N: 9.5, II_biostyr_NDN: 9.5, II_biostyr_PDN: 9.5,
    III_decantation: 8.5, III_biostyr_N: 9.5, III_biostyr_PDN: 9.5,
    codigestion_graisses: 16.5,
  },
  filtre_plateaux: {
    I_simple: 5.7, I_reactif: 5.7,
    II_forte: 8.4, II_moyenne: 8.4, II_faible_EB: 8.4, II_faible_ED: 8.4,
    II_prolongee_EB: 8.4, II_prolongee_ED: 8.4, II_MBR: 8.4, II_MBBR: 8.4, II_HybAS: 8.4,
    II_biostyr_C: 8.4, II_biostyr_N: 8.4, II_biostyr_NDN: 8.4, II_biostyr_PDN: 8.4,
    III_decantation: 7.1, III_biostyr_N: 5.5, III_biostyr_PDN: 7.1,
    codigestion_graisses: 15,
  },
  filtre_bandes: {
    I_simple: 6, I_reactif: 6,
    II_forte: 9, II_moyenne: 9, II_faible_EB: 9, II_faible_ED: 9,
    II_prolongee_EB: 9, II_prolongee_ED: 9, II_MBR: 9, II_MBBR: 9, II_HybAS: 9,
    II_biostyr_C: 9, II_biostyr_N: 9, II_biostyr_NDN: 9, II_biostyr_PDN: 9,
    III_decantation: 8, III_biostyr_N: 9, III_biostyr_PDN: 9,
    codigestion_graisses: 16,
  },
}
const CONC_DEFAUT = { centrifuge: 270, filtre_plateaux: 400, filtre_bandes: 230 }
const CAPTURE_DEFAUT = { centrifuge: 0.95, filtre_plateaux: 0.98, filtre_bandes: 0.96 }

const H = {
  correction_conc_boues_digerees: 20, // g/L de mieux sur des boues digérées
  ratio_soluble_NH4_NK: 0.95,
  ratio_soluble_DBO_DCO_dig_simple: 0.3 * 0.8,
  ratio_soluble_DBO_DCO_dig_avancee: 0.2 * 0.8,
  polymere_concentration: 1.5, // g/L de solution injectée
  conso_polymere: 0.76, // kWh/kg de polymère
  tps_fct_pompe: 5, // h/j
  dose_FeCl3_defaut: 50, // kg/tMS
  dose_chaux_kgCaO_tMV: (190 * 56.08) / 74.08, // EOLIA : 190 kg Ca(OH)2 par tMV
  MM_CaO: 56.08,
  MM_CaOH2: 74.08,
  // centrifugeuse : consommation interpolée entre deux siccités de référence
  centri_conc_x1: 150, centri_conso_x1: 1, // g/L → kWh/m³
  centri_conc_x2: 350, centri_conso_x2: 2.5,
  conso_filtre_plateaux: 50, // kWh/tMS
  conso_filtre_bandes: 30, // kWh/tMS
}

export default defineSludgeNode({
  id: 'deshydratation',
  label: 'Déshydratation',
  short: 'Déshy.',
  family: 'deshydratation',
  vba: 'z_Deshydratation.cls',
  multiple: true,
  maxInstances: 3,
  etapeSortie: (index) => ETAPE.deshydratees[Math.min(index, 3)],
  description:
    "Déshydratation mécanique par centrifugeuse, filtre à plateaux ou filtre à bandes. La siccité obtenue dépend de l'origine des boues et d'un éventuel passage en digestion ; le dosage de polymère suit la même logique. Les filtrats repartent en tête de station.",
  choices: [
    { key: 'type', label: 'Type de déshydratation', default: 'centrifuge', options: [
      { value: 'centrifuge', label: 'Centrifugeuse' },
      { value: 'filtre_plateaux', label: 'Filtre à plateaux' },
      { value: 'filtre_bandes', label: 'Filtre à bandes' },
    ] },
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: 'digerees', options: [
      { value: 'digerees', label: 'boues digérées' },
      { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
      { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
      { value: 'epaississeur_3', label: 'sortie épaississeur 3' },
      { value: 'toutes', label: 'boues extraites, toutes origines' },
      { value: 'I', label: 'boues primaires brutes' },
      { value: 'II', label: 'boues secondaires brutes' },
      { value: 'aucune', label: 'aucune' },
    ] },
    { key: 'entree_2', label: 'Boues admises (entrée 2)', default: 'aucune', options: [
      { value: 'aucune', label: 'aucune' },
      { value: 'digerees', label: 'boues digérées' },
      { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
      { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
      { value: 'epaississeur_3', label: 'sortie épaississeur 3' },
      { value: 'I', label: 'boues primaires brutes' },
      { value: 'II', label: 'boues secondaires brutes' },
    ] },
    { key: 'conditionnement', label: 'Conditionnement (filtre à plateaux)', default: 'chaux_FeCl3', options: [
      { value: 'polymere_FeCl3', label: 'polymère + FeCl3' },
      { value: 'chaux_FeCl3', label: 'chaux + FeCl3' },
      { value: 'aucun', label: 'sans conditionnement' },
    ] },
    { key: 'type_chaux', label: 'Nature de la chaux', default: 'vive', options: [
      { value: 'vive', label: 'chaux vive (CaO)' }, { value: 'eteinte', label: 'chaux éteinte (Ca(OH)2)' },
    ] },
    { key: 'digestion_amont', label: 'Type de digestion en amont', default: 'simple', options: [
      { value: 'simple', label: 'digestion simple' },
      { value: 'avancee', label: 'digestion avancée (Biothelys, Exelys)' },
      { value: 'aucune', label: 'pas de digestion' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'capture_MES', label: 'Taux de capture des MES', unit: '-', group: 'Séparation', default: undefined, hint: 'selon la technologie' },
    { key: 'outlet_concentration', label: 'Siccité des boues déshydratées', unit: 'g/L', group: 'Séparation', default: undefined, hint: 'moyenne des siccités de référence' },
    { key: 'dose_polymere', label: 'Dosage de polymère', unit: 'kg m.a./tMS', group: 'Réactifs', default: undefined, hint: 'selon origine et digestion amont' },
    { key: 'dose_chaux', label: 'Dosage de chaux', unit: 'kg/tMS', group: 'Réactifs', default: undefined, hint: '190 kg Ca(OH)2 par tMV' },
    { key: 'dose_FeCl3', label: 'Dosage de chlorure ferrique', unit: 'kg/tMS', group: 'Réactifs', default: undefined, hint: '50 kg/tMS en filtre à plateaux' },
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
    const etapeOut = ETAPE.deshydratees[Math.min(index, 3)]

    const entrees = entreesDepuisChoix(choices, p, ['entree_1', 'entree_2'])
    if (!entrees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    // ---- attribution_valeur_par_defaut
    const capture_MES = f('capture_MES') ?? CAPTURE_DEFAUT[type]
    let outlet_concentration = CONC_DEFAUT[type] // amorçage, recalculé plus bas
    const avecPolymere = type !== 'filtre_plateaux' || choices.conditionnement === 'polymere_FeCl3'
    const avecChaux = type === 'filtre_plateaux' && choices.conditionnement === 'chaux_FeCl3'
    const avecFeCl3 = type === 'filtre_plateaux' && choices.conditionnement !== 'aucun'
    const ratio_soluble_DBO_DCO = choices.digestion_amont === 'avancee'
      ? H.ratio_soluble_DBO_DCO_dig_avancee
      : H.ratio_soluble_DBO_DCO_dig_simple

    // ---- première passe : capture des MES, dosage de polymère par origine
    const lu = lireEntrees(table, soluble, entrees)
    let inlet_MES = 0, inlet_Q = 0, inlet_MV = 0, inlet_DCO = 0, inlet_DBO = 0, inlet_NK = 0, inlet_Pt = 0
    let outlet_MES = 0, outlet_Q = 0, outlet_DCO = 0, outlet_DBO = 0, outlet_NK = 0, outlet_Pt = 0
    let polymere_cumul = 0
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
      inlet_MES += MES; inlet_Q += Q; inlet_MV += MV
      inlet_DCO += DCO; inlet_DBO += DBO; inlet_NK += NK; inlet_Pt += Pt
      for (let k = 1; k <= 5; k++) soluble_kg[k] += (sol[k] * Q) / 1000

      // le dosage de polymère dépend de l'origine et d'un éventuel passage en
      // digestion ; il porte sur les MS, part soluble comprise
      const digeree = e.etape === ETAPE.digerees
      if (digeree) MES_digerees += MES
      const origine = src[P.origine]
      const bareme = digeree ? DOSE_POLY_DIG[type] : DOSE_POLY[type]
      const dose = bareme[origine]
      if (dose == null && origine) {
        warnings.push(`Origine de boue « ${origine} » sans dosage de polymère de référence : dosage nul retenu.`)
      }
      const MS = MES + (sol[SOL.MS_soluble] * Q) / 1000
      polymere_cumul += (dose ?? 0) * MS

      MES *= capture_MES; MV *= capture_MES; DCO *= capture_MES
      DBO *= capture_MES; NK *= capture_MES; Pt *= capture_MES
      const Qout = outlet_concentration > 0 ? MES / outlet_concentration : 0
      outlet_MES += MES; outlet_Q += Qout
      outlet_DCO += DCO; outlet_DBO += DBO; outlet_NK += NK; outlet_Pt += Pt

      if (!parType[j]) parType[j] = { MES: 0, MV: 0, Q: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, origine }
      const t = parType[j]
      t.MES += MES; t.MV += MV; t.Q += Qout
      t.DCO += DCO; t.DBO += DBO; t.NK += NK; t.Pt += Pt
      t.flux_in += e.flux_in
    }

    if (!(inlet_MES > 0) || !(inlet_Q > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée de la déshydratation."] }
    }
    const inlet_MS = inlet_MES + soluble_kg[SOL.MS_soluble]
    const inlet_concentration = inlet_MS / inlet_Q

    // ---- réactifs
    const dose_polymere = avecPolymere ? (f('dose_polymere') ?? (inlet_MS > 0 ? polymere_cumul / inlet_MS : 0)) : 0
    const polymere_flux = (dose_polymere / 1000) * inlet_MS
    const polymere_Q = polymere_flux / H.polymere_concentration
    let dose_chaux = 0
    let chaux_flux = 0
    if (avecChaux) {
      // le barème est exprimé en kg de CaO par tonne de MV : on le ramène aux MS,
      // et à la chaux éteinte le cas échéant
      const base = H.dose_chaux_kgCaO_tMV * (inlet_MES > 0 ? inlet_MV / inlet_MES : 0)
      dose_chaux = f('dose_chaux') ?? (choices.type_chaux === 'eteinte' ? (base * H.MM_CaOH2) / H.MM_CaO : base)
      chaux_flux = (dose_chaux / 1000) * inlet_MES
    }
    const dose_FeCl3 = avecFeCl3 ? (f('dose_FeCl3') ?? H.dose_FeCl3_defaut) : 0
    const FeCl3_flux = (dose_FeCl3 / 1000) * inlet_MES

    // ---- siccité de sortie
    const ref = CONC_REF[type]
    let conc = 0
    for (let j = 1; j <= NB_TYPES; j++) {
      const t = parType[j]
      if (!t || !(t.MES > 0) || !(outlet_MES > 0)) continue
      const c = ref[t.origine]
      if (c == null) {
        warnings.push(`Origine de boue « ${t.origine || 'inconnue'} » sans siccité de référence : valeur par défaut ${CONC_DEFAUT[type]} g/L retenue.`)
        conc += (t.MES / outlet_MES) * CONC_DEFAUT[type]
      } else conc += (t.MES / outlet_MES) * c
    }
    // les boues digérées se déshydratent mieux : bonus au prorata de leur part
    const part_digeree = inlet_MES > 0 ? MES_digerees / inlet_MES : 0
    conc += H.correction_conc_boues_digerees * part_digeree
    outlet_concentration = f('outlet_concentration') ?? conc
    if (outlet_concentration < inlet_concentration) {
      outlet_concentration = inlet_concentration * 1.01
      warnings.push("Siccité de sortie inférieure à celle d'entrée : plafonnée à 1 % au-dessus de l'entrée.")
    }

    // ---- reprise des débits avec la siccité définitive et écriture aval
    outlet_Q = 0
    for (const [jStr, t] of Object.entries(parType)) {
      const j = Number(jStr)
      t.Q = outlet_concentration > 0 ? t.MES / outlet_concentration : 0
      outlet_Q += t.Q
      const dst = table[etapeOut][j]
      dst[P.origine] = t.origine
      dst[P.MES] = t.MES
      dst[P.Q] = t.Q
      dst[P.MV_MES] = t.MES > 0 ? t.MV / t.MES : 0
      dst[P.ratio_DCO_MES] = t.MES > 0 ? t.DCO / t.MES : 0
      dst[P.ratio_DBO_MES] = t.MES > 0 ? t.DBO / t.MES : 0
      dst[P.ratio_NK_MES] = t.MES > 0 ? t.NK / t.MES : 0
      dst[P.ratio_Pt_MES] = t.MES > 0 ? t.Pt / t.MES : 0
      dst[P.flux_in] = t.flux_in
      dst[P.verif_flux] = t.flux_in
    }
    for (const e of lu) e.src[P.verif_flux] -= e.flux_in

    // ---- filtrats : eau séparée et pollution soluble entraînée
    const Q_retour = inlet_Q + polymere_Q - outlet_Q
    // Les filtrats de boues digérées sont les plus chargés en azote : ce sont
    // eux que le traitement des retours ira chercher. On les range à part.
    const origine_retour = part_digeree > 0 ? RET_ORIGINE.digestion : RET_ORIGINE.autres
    const vecteur = []
    vecteur[RET.Q] = Q_retour
    vecteur[RET.MES] = inlet_MES - outlet_MES
    vecteur[RET.DCO] = inlet_DCO - outlet_DCO
    vecteur[RET.DBO] = inlet_DBO - outlet_DBO
    vecteur[RET.NK] = inlet_NK - outlet_NK
    vecteur[RET.Pt] = inlet_Pt - outlet_Pt

    const part = (kg) => (Q_retour > 0 ? kg / (1 + outlet_Q / Q_retour) : 0)
    const sol_DCO = part(soluble_kg[SOL.DCO])
    const sol_NK = part(soluble_kg[SOL.NK])
    const sol_Pt = part(soluble_kg[SOL.Pt])
    const sol_MS = part(soluble_kg[SOL.MS_soluble])
    const sol_MV = part(soluble_kg[SOL.MV_soluble])
    vecteur[RET.DCO] += sol_DCO
    vecteur[RET.DBO] += ratio_soluble_DBO_DCO * sol_DCO
    vecteur[RET.NK] += sol_NK
    vecteur[RET.NH4] = H.ratio_soluble_NH4_NK * sol_NK
    vecteur[RET.Pt] += sol_Pt
    const vecteur_soluble = []
    vecteur_soluble[RET.Q] = Q_retour
    vecteur_soluble[RET.DCO] = sol_DCO
    vecteur_soluble[RET.NK] = sol_NK
    vecteur_soluble[RET.NH4] = H.ratio_soluble_NH4_NK * sol_NK
    vecteur_soluble[RET.Pt] = sol_Pt
    if (ctx.retours) ajouterRetour(ctx.retours, origine_retour, vecteur, vecteur_soluble)
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
    let electricite_specifique = 0
    if (type === 'centrifuge') {
      // consommation interpolée linéairement sur la siccité obtenue
      const ratio = H.centri_conso_x1
        + ((outlet_concentration - H.centri_conc_x1) / (H.centri_conc_x2 - H.centri_conc_x1)) * (H.centri_conso_x2 - H.centri_conso_x1)
      electricite_specifique = ratio * (inlet_Q + polymere_Q)
    } else if (type === 'filtre_plateaux') {
      electricite_specifique = (H.conso_filtre_plateaux * inlet_MES) / 1000
    } else {
      electricite_specifique = (H.conso_filtre_bandes * inlet_MES) / 1000
    }
    const electricite_polymere = type === 'filtre_plateaux' ? 0 : H.conso_polymere * polymere_flux

    // le filtre à plateaux évacue par gravité : pas de pompe d'extraction
    let electricite_extraction = 0
    let extr = { nb: 0, tps: 0, Qu: 0 }
    if (type !== 'filtre_plateaux') {
      extr = repartitionPompage(outlet_Q, p.extraction_pompe_nb, p.extraction_tps_fonctionnement, forced, 'extraction', H.tps_fct_pompe)
      if (extr.incoherence) warnings.push("Incohérence sur le pompage d'extraction.")
      const rdt = f('extraction_pompe_rdt') ?? rdtPompeBoues(extr.Qu, p.extraction_P_refoulement)
      electricite_extraction = elecPompage(extr.Qu, extr.nb, extr.tps, p.extraction_P_refoulement, rdt)
    }

    const total = electricite_specifique + electricite_polymere + electricite_extraction

    const reactifs = {}
    if (polymere_flux > 0) reactifs.polymere = polymere_flux
    if (chaux_flux > 0) reactifs[choices.type_chaux === 'eteinte' ? 'chaux_eteinte' : 'chaux_vive'] = chaux_flux
    if (FeCl3_flux > 0) reactifs.FeCl3 = FeCl3_flux

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MES en entrée', unit: 'kg/j', value: inlet_MES },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: inlet_Q },
        { key: 'in_conc', label: 'Siccité en entrée', unit: 'g/L', value: inlet_concentration },
        { key: 'part_dig', label: 'Part de boues digérées', unit: '-', value: part_digeree },
        { key: 'capture', label: 'Taux de capture des MES', unit: '-', value: capture_MES },
        { key: 'out_MES', label: 'MES déshydratées', unit: 'kg/j', value: outlet_MES },
        { key: 'out_MS', label: 'MS déshydratées (avec soluble)', unit: 'kg/j', value: outlet_MS },
        { key: 'out_Q', label: 'Débit de boues déshydratées', unit: 'm³/j', value: outlet_Q },
        { key: 'out_conc', label: 'Siccité des boues déshydratées', unit: 'g/L', value: outlet_concentration },
        { key: 'siccite_pct', label: 'Siccité des boues déshydratées', unit: '%', value: outlet_concentration / 10 },
        ...(avecPolymere ? [
          { key: 'dose_poly', label: 'Dosage de polymère', unit: 'kg m.a./tMS', value: dose_polymere },
          { key: 'poly', label: 'Polymère (matière active)', unit: 'kg/j', value: polymere_flux },
        ] : []),
        ...(avecChaux ? [
          { key: 'dose_chaux', label: `Dosage de chaux ${choices.type_chaux}`, unit: 'kg/tMS', value: dose_chaux },
          { key: 'chaux', label: `Chaux ${choices.type_chaux}`, unit: 'kg/j', value: chaux_flux },
        ] : []),
        ...(avecFeCl3 ? [{ key: 'FeCl3', label: 'Chlorure ferrique', unit: 'kg/j', value: FeCl3_flux }] : []),
        { key: 'ret_Q', label: 'Filtrats retournés en tête', unit: 'm³/j', value: Q_retour },
        { key: 'ret_MES', label: 'MES au retour en tête', unit: 'kg/j', value: inlet_MES - outlet_MES },
        { key: 'ret_NK', label: 'NK au retour en tête', unit: 'kg/j', value: inlet_NK - outlet_NK + sol_NK },
        { key: 'ret_NH4', label: 'dont N-NH4', unit: 'kg/j', value: H.ratio_soluble_NH4_NK * sol_NK },
        { key: 'ret_NK_mgL', label: 'Concentration en NK des filtrats', unit: 'mg/L', value: Q_retour > 0 ? ((inlet_NK - outlet_NK + sol_NK) / Q_retour) * 1000 : 0 },
      ],
      reactifs,
      electricity: {
        total,
        fixed: 0,
        detail: {
          procede: electricite_specifique,
          polymere: electricite_polymere,
          extraction: electricite_extraction,
        },
      },
      warnings,
    }
  },
})
