// ---------------------------------------------------------------------------
// Port de z_CLS_RetoursMAP.cls — cristallisation de struvite sur les retours.
//
// Troisième voie de traitement des jus, à côté d'ANITA Mox et d'ANITA Shunt,
// mais d'une nature différente : au lieu d'éliminer l'azote, la MAP
// (phosphate ammoniaco-magnésien, ou struvite) le **récupère** avec le
// phosphore sous forme d'un engrais commercialisable.
//
//   Mg²⁺ + NH₄⁺ + PO₄³⁻ + 6 H₂O  →  MgNH₄PO₄·6H₂O
//
// La stœchiométrie est équimolaire : chaque mole de phosphate précipitée
// emporte une mole d'ammonium. C'est ce qui limite l'intérêt du procédé sur
// l'azote — les jus de digestion contiennent bien plus d'azote que de
// phosphore, si bien que la struvite n'en capte qu'une fraction. En revanche
// elle élimine 80 % du phosphate, là où les deux ANITA n'y touchent pas.
//
// Le magnésium est le réactif limitant : il n'est pas présent dans les jus en
// quantité suffisante et doit être apporté sous forme de MgCl₂, à raison de
// 1,2 mole par mole de phosphate à précipiter.
//
// Quatre conditions d'applicabilité sont vérifiées, toutes portées :
//   P-PO₄ d'entrée au moins 70 mg/L, sans quoi la cristallisation ne s'amorce
//   pas ; MES au plus 2 000 mg/L ; rapport molaire N-NH₄/P-PO₄ au moins 2 ;
//   rapport molaire Ca/P-PO₄ au plus 1, le calcium entrant en concurrence avec
//   le magnésium.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - les quatre contrôles d'applicabilité et celui du pH existent dans le
//    classeur mais leurs messages y sont vides : le port émet de vrais
//    avertissements ;
//  - le VBA fait deux passes ; le port n'en fait qu'une, la file boues n'ayant
//    qu'un jeu de flux, et expose les charges nominales en paramètres ;
//  - la struvite produite quitte la filière comme produit valorisé : elle est
//    exposée en résultat mais ne rejoint pas la table des boues.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { RET, NB_RET } from '../core/sludge.js'
import { retoursAdmis, remplacerRetoursAdmis } from './anitaMox.js'

// masses molaires (kg/mol)
const MM = { Mg: 0.02431, Cl: 0.03545, N: 0.014, H: 0.001, P: 0.031, O: 0.016, Ca: 0.04008 }
MM.MgCl2 = MM.Mg + 2 * MM.Cl
// struvite hexahydratée MgNH4PO4·6H2O
MM.struvite = MM.Mg + MM.N + 4 * MM.H + MM.P + 4 * MM.O + 12 * MM.H + 6 * MM.O

const H = {
  rendement_PO4_defaut: 0.8,
  outlet_PPO4_mini_mgL: 10, // plancher de solubilité
  inlet_PPO4_mini_mgL: 70, // sous ce seuil la cristallisation ne s'amorce pas
  inlet_MES_maxi_mgL: 2000,
  ratio_molaire_mini_NH4_PO4: 2,
  ratio_molaire_maxi_Ca_PO4: 1, // le calcium concurrence le magnésium
  ratio_molaire_Mg_P_defaut: 1.2,
  pH_mini: 7.5,
  // gamme Ostara PEARL
  production_max_PEARL500: 500, // kg/j
  production_max_PEARL2000: 2000,
  // relevés d'exploitation : Chartres 0,68 et Perpignan 0,77 kWh/kg sur PEARL 500 ;
  // Bruxelles 0,35 kWh/kg sur PEARL 2000
  elec_PEARL500_kWh_kg: 0.7,
  elec_PEARL2000_kWh_kg: 0.35,
}

export default defineSludgeNode({
  id: 'retours-map',
  label: 'Cristallisation MAP',
  short: 'MAP',
  family: 'retours',
  vba: 'z_CLS_RetoursMAP.cls',
  etapeSortie: null,
  description:
    "Cristallisation de struvite sur les jus de digestion. Le phosphate et l'ammonium précipitent avec du magnésium apporté sous forme de MgCl2, formant un engrais valorisable. Le procédé élimine 80 % du phosphate mais n'agit sur l'azote qu'à hauteur de la stœchiométrie.",
  choices: [
    { key: 'jus_traites', label: 'Jus dirigés vers le traitement', default: 'digestion', options: [
      { value: 'digestion', label: 'jus de digestion seuls' },
      { value: 'athos', label: "jus d'Athos seuls" },
      { value: 'les_deux', label: "jus de digestion et d'Athos" },
    ] },
    { key: 'process', label: 'Technologie', default: 'ostara', options: [
      { value: 'ostara', label: 'Ostara PEARL' },
      { value: 'autre', label: 'autre technologie' },
    ] },
  ],
  params: [
    { key: 'Q_nominal', label: 'Débit nominal de jus', unit: 'm³/j', group: 'Conditions', default: undefined, hint: 'débit admis si non forcé' },
    { key: 'NH4_nominal', label: 'N-NH4 des jus', unit: 'mg/L', group: 'Conditions', default: undefined, hint: 'concentration admise si non forcée' },
    { key: 'PPO4_nominal', label: 'P-PO4 des jus', unit: 'mg/L', group: 'Conditions', default: undefined, hint: 'phosphore soluble admis si non forcé' },
    { key: 'inlet_Mg', label: 'Magnésium présent dans les jus', unit: 'mg/L', group: 'Conditions', default: 0, hint: 'nul par défaut, ce qui surestime la consommation de réactif' },
    { key: 'inlet_Ca', label: 'Calcium présent dans les jus', unit: 'mg/L', group: 'Conditions', default: 0 },
    { key: 'inlet_pH', label: 'pH des jus', unit: '-', group: 'Conditions', default: 7.5 },
    { key: 'inlet_temperature', label: 'Température des jus', unit: '°C', group: 'Conditions', default: 25 },
    { key: 'rendement_PO4', label: "Rendement d'élimination du phosphate", unit: '-', group: 'Cristallisation', default: undefined, hint: '80 %, plafonné par le plancher de 10 mg/L en sortie' },
    { key: 'ratio_molaire_Mg_P', label: 'Rapport molaire Mg/P visé', unit: '-', group: 'Cristallisation', default: 1.2 },
    { key: 'MgCl2_pur', label: 'Consommation de MgCl2 pur', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: 'calculée sur le rapport molaire' },
  ],

  compute(ctx) {
    const { retours, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    if (!retours) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Les vecteurs de retour ne sont pas disponibles."] }
    }

    const lu = retoursAdmis(retours, choices.jus_traites)
    if (!(lu.Q > 0)) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Aucun jus admissible en amont : le traitement des retours suppose une digestion ou un Athos dans la filière boues."],
      }
    }

    // ---- charges admises
    const Q = f('Q_nominal') ?? lu.Q
    const facteurDebit = Q > 0 ? lu.Q / Q : 1
    const c = lu.concentrations
    const DCO = (c.DCO * facteurDebit * Q) / 1000
    const DBO = (c.DBO * facteurDebit * Q) / 1000
    const MES = (c.MES * facteurDebit * Q) / 1000
    let NH4 = f('NH4_nominal') != null ? (f('NH4_nominal') * Q) / 1000 : (c.NH4 * facteurDebit * Q) / 1000
    let NK = ((c.NK - c.NH4) * facteurDebit * Q) / 1000 + NH4
    const NO3 = (c.NO3 * facteurDebit * Q) / 1000
    let PPO4 = f('PPO4_nominal') != null ? (f('PPO4_nominal') * Q) / 1000 : (c.PO4 * facteurDebit * Q) / 1000
    let Pt = ((c.Pt - c.PO4) * facteurDebit * Q) / 1000 + PPO4

    const inlet_Mg = p.inlet_Mg ?? 0
    const inlet_Ca = p.inlet_Ca ?? 0
    const pH = p.inlet_pH ?? 7.5
    const PPO4_inlet_mgL = Q > 0 ? (PPO4 / Q) * 1000 : 0
    const MES_mgL = Q > 0 ? (MES / Q) * 1000 : 0

    if (!(PPO4 > 0)) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Les jus ne contiennent pas de phosphore soluble : la cristallisation de struvite est sans objet."],
      }
    }

    // ---- conditions d'applicabilité (messages vides dans le classeur)
    if (PPO4_inlet_mgL < H.inlet_PPO4_mini_mgL) {
      warnings.push(`P-PO4 des jus (${PPO4_inlet_mgL.toFixed(0)} mg/L) inférieur au minimum de ${H.inlet_PPO4_mini_mgL} mg/L : la cristallisation ne s'amorce pas dans de bonnes conditions.`)
    }
    if (MES_mgL > H.inlet_MES_maxi_mgL) {
      warnings.push(`MES des jus (${MES_mgL.toFixed(0)} mg/L) supérieures au maximum de ${H.inlet_MES_maxi_mgL} mg/L : une clarification amont est nécessaire.`)
    }
    const ratio_N_P = (NH4 / MM.N) / (PPO4 / MM.P)
    if (ratio_N_P < H.ratio_molaire_mini_NH4_PO4) {
      warnings.push(`Rapport molaire N-NH4/P-PO4 (${ratio_N_P.toFixed(1)}) inférieur au minimum de ${H.ratio_molaire_mini_NH4_PO4} : l'ammonium risque de devenir limitant.`)
    }
    const ratio_Ca_P = (inlet_Ca / MM.Ca) / (PPO4_inlet_mgL / MM.P)
    if (ratio_Ca_P > H.ratio_molaire_maxi_Ca_PO4) {
      warnings.push(`Rapport molaire Ca/P-PO4 (${ratio_Ca_P.toFixed(1)}) supérieur au maximum de ${H.ratio_molaire_maxi_Ca_PO4} : le calcium concurrence le magnésium et dégrade la qualité du cristal.`)
    }
    if (pH < H.pH_mini) {
      warnings.push(`pH des jus (${pH}) inférieur à ${H.pH_mini} : une correction alcaline est nécessaire pour amorcer la précipitation.`)
    }

    // ---- calcul_cristallisation
    let rendement_PO4 = f('rendement_PO4')
    let PPO4_outlet_mgL
    if (rendement_PO4 != null) {
      PPO4_outlet_mgL = PPO4_inlet_mgL * (1 - rendement_PO4)
    } else {
      rendement_PO4 = H.rendement_PO4_defaut
      PPO4_outlet_mgL = PPO4_inlet_mgL * (1 - rendement_PO4)
      // le phosphate résiduel ne peut descendre sous le plancher de solubilité
      if (PPO4_outlet_mgL < H.outlet_PPO4_mini_mgL) {
        PPO4_outlet_mgL = H.outlet_PPO4_mini_mgL
        rendement_PO4 = (PPO4_inlet_mgL - PPO4_outlet_mgL) / PPO4_inlet_mgL
      }
    }
    const PPO4_precipite = PPO4 * rendement_PO4
    const PO4_precipite_mol = PPO4_precipite / MM.P

    // Le magnésium est apporté au prorata du phosphate **entrant**, pas du seul
    // phosphate précipité : le VBA divise par le rendement, ce qui revient à
    // doser sur la charge totale reçue.
    let ratio_Mg_P = f('ratio_molaire_Mg_P') ?? H.ratio_molaire_Mg_P_defaut
    let MgCl2_mol = ratio_Mg_P * (PO4_precipite_mol / rendement_PO4)
    // on déduit le magnésium déjà présent dans les jus
    MgCl2_mol -= ((inlet_Mg * Q) / 1000) / MM.Mg
    if (MgCl2_mol < 0) {
      if (f('ratio_molaire_Mg_P') == null) {
        // les jus apportent déjà plus de magnésium que nécessaire
        ratio_Mg_P = (((inlet_Mg * Q) / 1000) / MM.Mg) / (PO4_precipite_mol / rendement_PO4)
        warnings.push('Le magnésium présent dans les jus suffit à la précipitation : aucun apport de MgCl2.')
      }
      MgCl2_mol = 0
    }
    const MgCl2_pur = f('MgCl2_pur') ?? MgCl2_mol * MM.MgCl2
    // stœchiométrie équimolaire : une mole de phosphate emporte une mole d'ammonium
    const struvite = PO4_precipite_mol * MM.struvite
    const NH4_precipite = PO4_precipite_mol * MM.N

    Pt -= PPO4_precipite
    PPO4 -= PPO4_precipite
    NH4 -= NH4_precipite
    NK -= NH4_precipite

    // ---- gamme de réacteurs Ostara PEARL
    let reacteurs = '—'
    if (choices.process === 'ostara') {
      const production_mini = 0.6 * H.production_max_PEARL500
      if (struvite < production_mini) {
        reacteurs = '—'
        warnings.push(`Production de struvite (${struvite.toFixed(0)} kg/j) inférieure au minimum de ${production_mini.toFixed(0)} kg/j : aucun réacteur PEARL n'est adapté.`)
      } else if (struvite < H.production_max_PEARL500) reacteurs = '1 PEARL 500'
      else if (struvite < 2 * H.production_max_PEARL500) reacteurs = '2 PEARL 500'
      else {
        let nb = 0
        do { nb += 1 } while (nb < struvite / H.production_max_PEARL2000)
        reacteurs = `${nb} PEARL 2000`
      }
    }

    // ---- flux de sortie
    const sortie = new Array(NB_RET + 1).fill(0)
    sortie[RET.Q] = Q
    sortie[RET.DCO] = DCO
    sortie[RET.DBO] = DBO
    sortie[RET.MES] = MES
    sortie[RET.NK] = Math.max(0, NK)
    sortie[RET.NH4] = Math.max(0, NH4)
    sortie[RET.NO3] = NO3
    sortie[RET.Pt] = Math.max(0, Pt)

    const sortie_soluble = lu.admis_soluble.slice()
    sortie_soluble[RET.Q] = Q
    sortie_soluble[RET.NK] = Math.max(0, sortie_soluble[RET.NK] - NH4_precipite)
    sortie_soluble[RET.NH4] = Math.max(0, sortie_soluble[RET.NH4] - NH4_precipite)
    sortie_soluble[RET.Pt] = Math.max(0, PPO4)
    remplacerRetoursAdmis(retours, lu, sortie, sortie_soluble)

    // ---- électricité, d'après les relevés d'exploitation
    let electricite = 0
    if (choices.process === 'ostara') {
      electricite = struvite < 2 * H.production_max_PEARL500
        ? H.elec_PEARL500_kWh_kg * struvite
        : H.elec_PEARL2000_kWh_kg * struvite
    }

    return {
      etapeSortie: null,
      results: [
        { key: 'jus_Q', label: 'Débit de jus traités', unit: 'm³/j', value: Q },
        { key: 'jus_PPO4', label: 'P-PO4 des jus', unit: 'mg/L', value: PPO4_inlet_mgL },
        { key: 'jus_NH4', label: 'N-NH4 des jus', unit: 'mg/L', value: Q > 0 ? ((NH4 + NH4_precipite) / Q) * 1000 : 0 },
        { key: 'ratio_N_P', label: 'Rapport molaire N-NH4 / P-PO4', unit: '-', value: ratio_N_P },
        { key: 'rdt_PO4', label: "Rendement d'élimination du phosphate", unit: '-', value: rendement_PO4 },
        { key: 'P_precip', label: 'Phosphore précipité', unit: 'kg P/j', value: PPO4_precipite },
        { key: 'N_precip', label: 'Azote précipité', unit: 'kg N/j', value: NH4_precipite },
        { key: 'ratio_Mg_P', label: 'Rapport molaire Mg/P retenu', unit: '-', value: ratio_Mg_P },
        { key: 'MgCl2', label: 'MgCl2 pur', unit: 'kg/j', value: MgCl2_pur },
        { key: 'struvite', label: 'Struvite produite', unit: 'kg/j', value: struvite },
        { key: 'struvite_t', label: 'Struvite produite', unit: 't/an', value: (struvite * 365) / 1000 },
        { key: 'reacteurs', label: 'Réacteurs installés', unit: '', value: reacteurs },
        { key: 'out_PPO4', label: 'P-PO4 en sortie', unit: 'mg/L', value: PPO4_outlet_mgL },
        { key: 'out_NH4', label: 'N-NH4 en sortie', unit: 'mg/L', value: Q > 0 ? (sortie[RET.NH4] / Q) * 1000 : 0 },
        { key: 'rdt_N', label: "Rendement d'élimination de l'azote", unit: '-', value: NH4 + NH4_precipite > 0 ? NH4_precipite / (NH4 + NH4_precipite) : 0 },
        { key: 'elec_spec', label: 'Consommation par kg de struvite', unit: 'kWh/kg', value: struvite > 0 ? electricite / struvite : 0 },
      ],
      reactifs: MgCl2_pur > 0 ? { MgCl2: MgCl2_pur } : {},
      produits: { struvite_kgj: struvite },
      electricity: { total: electricite, fixed: 0, detail: { cristallisation: electricite } },
      warnings,
    }
  },
})
