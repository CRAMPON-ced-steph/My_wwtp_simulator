// ---------------------------------------------------------------------------
// Port de z_Chaulage_boues.cls — chaulage des boues déshydratées.
//
// Le chaulage poursuit deux objectifs, qui conduisent à deux dosages :
//  - l'hygiénisation, qui demande un taux fixe de 30 % de CaO rapporté aux MS ;
//  - le relèvement de siccité à 300 g/L, dont le taux se déduit d'un bilan de
//    masse entre la chaux ajoutée, l'eau qu'elle consomme en s'hydratant et
//    les MS supplémentaires qu'elle apporte.
// Le choix « les deux » retient le plus élevé des deux taux.
//
// Contrairement à l'épaississement et à la déshydratation, le chaulage
// n'extrait pas d'eau : il n'y a donc pas de retour en tête. La masse de boues
// augmente de la masse de chaux apportée, et la siccité monte parce que la
// chaux vive consomme de l'eau en s'éteignant.
//
// Écart au VBA, volontaire et signalé (voir README) : le VBA ajoute au flux
// sortant une contribution `stockage_boues_pdtes` issue d'un stockage amont
// non renseigné dans le port ; elle est ignorée.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, NB_TYPES } from '../core/sludge.js'
import { entreesDepuisChoix, lireEntrees } from './_commun.js'

const H = {
  ratio_elec: 1, // kWh par tonne de boue brute (malaxeur)
  conc_requise: 300, // g/L visés en relèvement de siccité
  taux_chaux_hygiene: 0.3, // kg CaO par kg de MS
  MM_CaOH2: 74.08,
  MM_CaO: 56.08,
}

export default defineSludgeNode({
  id: 'chaulage',
  label: 'Chaulage',
  short: 'Chaulage',
  family: 'hygienisation',
  vba: 'z_Chaulage_boues.cls',
  multiple: true,
  maxInstances: 3,
  etapeSortie: (index) => ETAPE.chaulees[Math.min(index, 3)],
  description:
    "Chaulage des boues déshydratées, pour hygiénisation ou pour relever la siccité à 300 g/L. La chaux vive consomme de l'eau en s'éteignant, ce qui augmente à la fois la masse et la siccité des boues. Aucun retour en tête.",
  choices: [
    { key: 'objectif', label: 'Objectif du chaulage', default: 'hygiene', options: [
      { value: 'hygiene', label: 'hygiénisation (30 % CaO/MS)' },
      { value: 'siccite', label: 'relèvement de siccité à 300 g/L' },
      { value: 'les_deux', label: 'les deux (dosage le plus élevé)' },
    ] },
    { key: 'type_chaux', label: 'Nature de la chaux', default: 'vive', options: [
      { value: 'vive', label: 'chaux vive (CaO)' }, { value: 'eteinte', label: 'chaux éteinte (Ca(OH)2)' },
    ] },
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: 'deshydratees_1', options: [
      { value: 'deshydratees_1', label: 'sortie déshydratation 1' },
      { value: 'deshydratees_2', label: 'sortie déshydratation 2' },
      { value: 'deshydratees_3', label: 'sortie déshydratation 3' },
      { value: 'digerees', label: 'boues digérées' },
      { value: 'aucune', label: 'aucune' },
    ] },
    { key: 'entree_2', label: 'Boues admises (entrée 2)', default: 'aucune', options: [
      { value: 'aucune', label: 'aucune' },
      { value: 'deshydratees_1', label: 'sortie déshydratation 1' },
      { value: 'deshydratees_2', label: 'sortie déshydratation 2' },
      { value: 'deshydratees_3', label: 'sortie déshydratation 3' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'conc_requise', label: 'Siccité visée', unit: 'g/L', group: 'Dosage', default: 300 },
    { key: 'taux_chaux', label: 'Taux de chaux', unit: 'kg CaO/kg MS', group: 'Dosage', default: undefined, hint: '0,3 en hygiénisation, calculé en relèvement de siccité' },
    { key: 'ratio_elec', label: 'Consommation spécifique du malaxeur', unit: 'kWh/t de boue brute', group: 'Électricité', default: 1 },
  ],

  compute(ctx) {
    const { table, soluble, choices, forced, p, index } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.chaulees[Math.min(index, 3)]
    const eteinte = choices.type_chaux === 'eteinte'
    const rapport_MM = H.MM_CaOH2 / H.MM_CaO

    const entrees = entreesDepuisChoix(choices, p, ['entree_1', 'entree_2'])
    if (!entrees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    // ---- lecture des entrées ; le chaulage ne sépare rien, tout ressort
    const lu = lireEntrees(table, soluble, entrees)
    let inlet_MES = 0, inlet_Q = 0, inlet_MV = 0, inlet_DCO = 0, inlet_DBO = 0, inlet_NK = 0, inlet_Pt = 0
    const parType = {}
    for (const e of lu) {
      const { j, MES, Q, MV_MES, ratios, src, sol } = e
      if (!(MES > 0)) continue
      inlet_MES += MES
      inlet_Q += Q
      inlet_MV += MES * MV_MES
      inlet_DCO += MES * ratios.DCO
      inlet_DBO += MES * ratios.DBO
      inlet_NK += MES * ratios.NK
      inlet_Pt += MES * ratios.Pt
      if (!parType[j]) parType[j] = { MES: 0, MV: 0, Q: 0, DCO: 0, DBO: 0, NK: 0, Pt: 0, flux_in: 0, origine: src[P.origine], sol }
      const t = parType[j]
      t.MES += MES
      t.Q += Q
      t.MV += MES * MV_MES
      t.DCO += MES * ratios.DCO
      t.DBO += MES * ratios.DBO
      t.NK += MES * ratios.NK
      t.Pt += MES * ratios.Pt
      t.flux_in += e.flux_in
    }
    if (!(inlet_MES > 0) || !(inlet_Q > 0)) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée du chaulage."] }
    }
    const inlet_concentration = inlet_MES / inlet_Q
    const inlet_MV_MES = inlet_MV / inlet_MES
    const conc_requise = p.conc_requise ?? H.conc_requise

    /**
     * Taux de chaux nécessaire pour atteindre la siccité visée. Le bilan tient
     * compte de l'eau consommée par l'extinction de la chaux vive, ce qui
     * explique la différence de dénominateur entre les deux natures de chaux.
     */
    function tauxPourSiccite() {
      if (inlet_concentration > conc_requise) return 0
      const num = 1 - conc_requise / inlet_concentration
      const den = eteinte
        ? (conc_requise / 1000 - 1) * rapport_MM
        : conc_requise / 1000 - rapport_MM
      return den !== 0 ? num / den : 0
    }

    let taux_chaux
    const tf = f('taux_chaux')
    if (tf != null) taux_chaux = tf
    else if (choices.objectif === 'hygiene') taux_chaux = H.taux_chaux_hygiene
    else if (choices.objectif === 'siccite') taux_chaux = tauxPourSiccite()
    else taux_chaux = Math.max(tauxPourSiccite(), H.taux_chaux_hygiene)

    // ---- masse de chaux et siccité obtenue
    // Le VBA écrit ce bilan sous forme d'un rapport : le numérateur est la masse
    // sèche finale par kg de MS entrante, le dénominateur le volume final.
    const dose = taux_chaux // kg CaO par kg de MS entrante
    let chaux_flux, outlet_concentration, outlet_MES
    if (eteinte) {
      chaux_flux = dose * inlet_MES * rapport_MM
      const r = chaux_flux / inlet_MES
      outlet_concentration = (1000 * (r + 1)) / (1000 / inlet_concentration + r)
      outlet_MES = (inlet_MES * outlet_concentration) / 1000 * (1000 / inlet_concentration + r)
    } else {
      chaux_flux = dose * inlet_MES
      const r = chaux_flux / inlet_MES
      outlet_concentration = (1000 * (rapport_MM * r + 1)) / (1000 / inlet_concentration + r)
      outlet_MES = (inlet_MES * outlet_concentration) / 1000 * (1000 / inlet_concentration + r)
    }
    const outlet_Q = outlet_concentration > 0 ? outlet_MES / outlet_concentration : 0
    const outlet_MV_MES = outlet_MES > 0 ? (inlet_MV_MES * inlet_MES) / outlet_MES : 0

    if (taux_chaux === 0 && choices.objectif === 'siccite') {
      warnings.push(`Siccité d'entrée (${inlet_concentration.toFixed(0)} g/L) déjà supérieure à la cible : aucun apport de chaux.`)
    }
    if (choices.objectif !== 'hygiene' && taux_chaux < H.taux_chaux_hygiene) {
      warnings.push("Taux de chaux inférieur à 30 % de CaO : l'hygiénisation des boues n'est pas assurée.")
    }

    // ---- écriture de l'étape aval ; l'apport de chaux est réparti au prorata
    // des MES de chaque type de boue
    const facteur = inlet_MES > 0 ? outlet_MES / inlet_MES : 1
    for (const [jStr, t] of Object.entries(parType)) {
      const j = Number(jStr)
      const MES = t.MES * facteur
      const dst = table[etapeOut][j]
      dst[P.origine] = t.origine
      dst[P.MES] = MES
      dst[P.Q] = outlet_concentration > 0 ? MES / outlet_concentration : 0
      // la chaux est minérale : elle dilue les matières volatiles et les ratios
      dst[P.MV_MES] = MES > 0 ? t.MV / MES : 0
      dst[P.ratio_DCO_MES] = MES > 0 ? t.DCO / MES : 0
      dst[P.ratio_DBO_MES] = MES > 0 ? t.DBO / MES : 0
      dst[P.ratio_NK_MES] = MES > 0 ? t.NK / MES : 0
      dst[P.ratio_Pt_MES] = MES > 0 ? t.Pt / MES : 0
      dst[P.flux_in] = t.flux_in
      dst[P.verif_flux] = t.flux_in
      // la pollution soluble traverse le chaulage sans modification
      for (let k = 1; k <= 5; k++) soluble[etapeOut][j][k] = t.sol[k]
    }
    for (const e of lu) e.src[P.verif_flux] -= e.flux_in

    // ---- électricité : malaxeur, rapporté au tonnage de boue brute
    const ratio_elec = p.ratio_elec ?? H.ratio_elec
    const electricite = taux_chaux > 0 ? inlet_Q * ratio_elec : 0

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MES en entrée', unit: 'kg/j', value: inlet_MES },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: inlet_Q },
        { key: 'in_conc', label: 'Siccité en entrée', unit: 'g/L', value: inlet_concentration },
        { key: 'in_siccite_pct', label: "Siccité en entrée", unit: '%', value: inlet_concentration / 10 },
        { key: 'taux', label: 'Taux de chaux', unit: 'kg CaO/kg MS', value: taux_chaux },
        { key: 'chaux', label: `Chaux ${eteinte ? 'éteinte' : 'vive'}`, unit: 'kg/j', value: chaux_flux },
        { key: 'out_MES', label: 'MS des boues chaulées', unit: 'kg/j', value: outlet_MES },
        { key: 'out_Q', label: 'Débit de boues chaulées', unit: 'm³/j', value: outlet_Q },
        { key: 'out_conc', label: 'Siccité des boues chaulées', unit: 'g/L', value: outlet_concentration },
        { key: 'out_siccite_pct', label: 'Siccité des boues chaulées', unit: '%', value: outlet_concentration / 10 },
        { key: 'out_MV_MES', label: 'MV/MS des boues chaulées', unit: '-', value: outlet_MV_MES },
        { key: 'gain', label: 'Gain de masse sèche', unit: 'kg/j', value: outlet_MES - inlet_MES },
      ],
      reactifs: chaux_flux > 0 ? { [eteinte ? 'chaux_eteinte' : 'chaux_vive']: chaux_flux } : {},
      electricity: { total: electricite, fixed: 0, detail: { malaxeur: electricite } },
      warnings,
    }
  },
})
