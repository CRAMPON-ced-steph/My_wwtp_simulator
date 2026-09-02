// ---------------------------------------------------------------------------
// Port de z_Digestion_simple.cls — digestion anaérobie mésophile ou thermophile.
//
// Le cœur du calcul est un taux de réduction des matières volatiles propre à
// chaque origine de boue (0,64 pour une boue primaire, 0,30 pour une boue
// d'aération prolongée…), corrigé de la température du digesteur et du temps de
// séjour. Les MV détruites donnent le biogaz ; le reste de l'azote et du
// phosphore libéré part au retour en tête via la pollution soluble.
//
// Deux passes, comme dans OCEAN :
//   dimensionnement()     à SRT_design       → volume du digesteur
//   fonctionnement_reel() à SRT = V / Q_réel → production réelle de biogaz
// La seconde passe reprend le volume de la première ; le port les enchaîne dans
// le même compute(), la file boues n'ayant qu'un jeu de flux.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - la codigestion de coferments (BMP, facteur de sécurité) n'est pas portée :
//    elle dépend d'une collection saisie dans un formulaire du classeur ;
//  - le détail par eaux sales (boues_I_detail…) n'existe pas dans le port : le
//    coefficient de réduction est pris sur l'origine principale de chaque type
//    de boue, ce qui revient au même hors filière biofiltration séparée ;
//  - l'énergie thermique est calculée et exposée, mais aucun consommateur ne la
//    récupère tant que la gestion d'énergie n'est pas portée.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, NB_TYPES, TYPE } from '../core/sludge.js'
import { CONST, rendementMoteur } from '../core/hypotheses.js'
import { entreesDepuisChoix, lireEntrees, repartitionPompage, rdtPompeBoues, H_POMPE } from './_commun.js'

// Taux de réduction des MV de référence, par origine de boue
// (reduction_MV_reference du VBA, valeurs révisées par DCA le 31/05/11).
const REDUCTION_MV = {
  I_simple: 0.638, I_reactif: 0.614,
  II_forte: 0.546, II_moyenne: 0.501, II_faible_EB: 0.378, II_faible_ED: 0.32,
  II_prolongee_EB: 0.358, II_prolongee_ED: 0.3, II_MBR: 0.3, II_MBBR: 0.358, II_HybAS: 0.358,
  II_biostyr_C: 0.48, II_biostyr_N: 0.48, II_biostyr_NDN: 0.48, II_biostyr_PDN: 0.48,
  III_decantation: 0.157, III_biostyr_N: 0.157, III_biostyr_PDN: 0.157,
  codigestion_graisses: 0.629,
}

const H = {
  SRT_reference: { mesophile: 20, thermophile: 12 }, // j
  T_reference: { mesophile: 36, thermophile: 56 }, // °C
  a_correction_SRT_gain: 1.25,
  a_correction_SRT_n: 0.7,
  a_densite_biogaz: [44 / 22.4, (16 - 44) / 22.4], // densité = a0 + a1·%CH4
  ratio_CH4_DCO: 0.35, // Nm³ CH4 par kg de DCO méthanisée
  ratio_NK_MVdeg: 0.9, // part du NK solubilisé par MV dégradée
  ratio_P_Pinlet: 0.16, // part du phosphore particulaire solubilisé
  ratio_DCOsol_outlet: 0.05, // part de DCO soluble en sortie
  // pertes thermiques du digesteur, kWh/(m³·j), selon climat et isolation
  pertes_thermiques: {
    froid: { oui: 0.6, non: 3.72 },
    tempere: { oui: 0.36, non: 2.2 },
    chaud: { oui: 0.24, non: 1.44 },
  },
  efficacite_echangeur: 0.95,
  Cp_boues: 1.163, // kWh/(m³·°C)
  delta_T_boues: 3, // °C dans l'échangeur, côté boues
  tps_fct_pompe_extraction: 14, // h/j
  T_limite_BT_HT: 50, // °C : au-delà, la chaleur demandée est de la haute température
  agitation_W_m3: 7, // W/m³, agitation mécanique
  brassage_biogaz_Nm_h: 1.25, // Nm/h par m² de section
  conso_surpresseur: 4, // Wh/(Nm³·mCE)
}

/**
 * Correctif de vitesse lié à la température du digesteur (correction_T).
 * Courbe en deux paliers : optimum mésophile vers 36 °C, creux entre les deux
 * régimes, optimum thermophile vers 56 °C, effondrement au-delà de 80 °C.
 */
function correctionT(T) {
  if (T < 0) return 0
  if (T < 35) return T / 35
  if (T < 38) return 1
  if (T < 45) return -T / 14 + (3 + 5 / 7)
  if (T < 55) return (5 * T) / 91 - 1.975
  if (T < 58) return 1.05
  if (T < 80) return -T / 21 + (3 + 9 / 11)
  return 0
}
/** correctif lié au temps de séjour, rapporté au temps de séjour de référence */
function correctionSRT(SRT, SRT_ref) {
  const r = SRT_ref > 0 ? SRT / SRT_ref : 0
  const x = Math.pow(r, H.a_correction_SRT_n)
  return (H.a_correction_SRT_gain * x) / (x + H.a_correction_SRT_gain - 1)
}

export default defineSludgeNode({
  id: 'digestion',
  label: 'Digestion anaérobie',
  short: 'Digestion',
  family: 'stabilisation',
  vba: 'z_Digestion_simple.cls',
  etapeSortie: ETAPE.digerees,
  description:
    "Digestion anaérobie mésophile ou thermophile. La réduction des matières volatiles dépend de l'origine des boues, de la température et du temps de séjour ; elle produit le biogaz et libère l'azote et le phosphore qui repartent en tête via les eaux de déshydratation.",
  choices: [
    { key: 'type', label: 'Régime de digestion', default: 'mesophile', options: [
      { value: 'mesophile', label: 'Mésophile (36 °C)' }, { value: 'thermophile', label: 'Thermophile (56 °C)' },
    ] },
    { key: 'entree_1', label: 'Boues admises (entrée 1)', default: 'epaississeur_1', options: [
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
      { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
      { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
      { value: 'epaississeur_3', label: 'sortie épaississeur 3' },
      { value: 'I', label: 'boues primaires brutes' },
      { value: 'II', label: 'boues secondaires brutes' },
      { value: 'graisses', label: 'graisses' },
    ] },
    { key: 'entree_3', label: 'Boues admises (entrée 3)', default: 'aucune', options: [
      { value: 'aucune', label: 'aucune' },
      { value: 'epaississeur_1', label: 'sortie épaississeur 1' },
      { value: 'epaississeur_2', label: 'sortie épaississeur 2' },
      { value: 'epaississeur_3', label: 'sortie épaississeur 3' },
      { value: 'graisses', label: 'graisses' },
    ] },
    { key: 'chauffage_boues', label: 'Chauffage des boues', default: 'amont', options: [
      { value: 'amont', label: 'en amont du digesteur' },
      { value: 'recirculation', label: 'sur boucle de recirculation' },
    ] },
    { key: 'agitation', label: 'Agitation du digesteur', default: 'biogaz', options: [
      { value: 'mecanique', label: 'mécanique' },
      { value: 'recirculation', label: 'par recirculation des boues' },
      { value: 'biogaz', label: 'par injection de biogaz' },
    ] },
    { key: 'isolation', label: 'Digesteur isolé', default: 'oui', options: [
      { value: 'oui', label: 'oui' }, { value: 'non', label: 'non' },
    ] },
    { key: 'climat', label: 'Climat du site', default: 'tempere', options: [
      { value: 'froid', label: 'froid' }, { value: 'tempere', label: 'tempéré' }, { value: 'chaud', label: 'chaud' },
    ] },
    { key: 'biogaz_stock', label: 'Gazomètre de stockage', default: 'oui', options: [
      { value: 'oui', label: 'oui' }, { value: 'non', label: 'non' },
    ] },
    { key: 'valorisation_biogaz', label: 'Valorisation du biogaz', default: 'cogeneration', options: [
      { value: 'chaudiere', label: 'chaudière' },
      { value: 'cogeneration', label: 'cogénération' },
      { value: 'microturbine', label: 'microturbine' },
      { value: 'torchere', label: 'torchère seule' },
    ] },
  ],
  params: [
    { key: 'ratio_admis_1', label: 'Part du flux amont admise (entrée 1)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_2', label: 'Part du flux amont admise (entrée 2)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'ratio_admis_3', label: 'Part du flux amont admise (entrée 3)', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'SRT_design', label: 'Temps de séjour de dimensionnement', unit: 'j', group: 'Dimensionnement', default: undefined, hint: '20 j en mésophile, 12 j en thermophile' },
    { key: 'volume_total', label: 'Volume total de digestion', unit: 'm³', group: 'Dimensionnement', default: undefined, hint: 'débit × temps de séjour' },
    { key: 'T_digesteur', label: 'Température du digesteur', unit: '°C', group: 'Dimensionnement', default: undefined, hint: '36 °C en mésophile, 56 °C en thermophile' },
    { key: 'T_inlet', label: 'Température des boues en entrée', unit: '°C', group: 'Dimensionnement', default: undefined, hint: "température de l'eau en conditions réelles" },
    { key: 'nb_digesteur', label: 'Nombre de digesteurs', unit: 'u', group: 'Dimensionnement', default: 1 },
    { key: 'hauteur_digesteur', label: 'Hauteur du digesteur', unit: 'm', group: 'Dimensionnement', default: 10 },
    { key: 'biogaz_teneur_CH4', label: 'Teneur en méthane du biogaz', unit: '-', group: 'Biogaz', default: 0.63 },
    { key: 'ratio_biogaz_valorise', label: 'Part du biogaz valorisée', unit: '-', group: 'Biogaz', default: 1 },
    { key: 'energie_maintien', label: 'Énergie thermique de maintien en température', unit: 'kWh/j', group: 'Thermique', default: undefined, hint: 'pertes + chauffage des boues' },
    { key: 'tx_recirculation_agitation', label: 'Taux de recirculation pour agitation', unit: 'vol/j', group: 'Agitation', default: 5 },
    { key: 'extraction_pompe_nb', label: "Nombre de pompes d'extraction", unit: 'u', group: 'Extraction', default: 2 },
    { key: 'extraction_P_refoulement', label: 'Pression de refoulement en extraction', unit: 'mCE', group: 'Extraction', default: 50 },
    { key: 'extraction_tps_fonctionnement', label: 'Durée de fonctionnement en extraction', unit: 'h/j', group: 'Extraction', default: 14 },
    { key: 'extraction_Q_unitaire', label: "Débit unitaire des pompes d'extraction", unit: 'm³/h', group: 'Extraction', default: undefined, hint: 'calculé si non forcé' },
    { key: 'extraction_pompe_rdt', label: "Rendement global des pompes d'extraction", unit: '-', group: 'Extraction', default: undefined, hint: 'machine 0,4 × moteur' },
  ],

  compute(ctx) {
    const { site, table, soluble, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.digerees
    const regime = choices.type

    // ---- attribution_valeur_par_defaut
    const SRT_reference = H.SRT_reference[regime]
    const T_reference = H.T_reference[regime]
    const T_digesteur = f('T_digesteur') ?? T_reference
    const T_inlet = f('T_inlet') ?? site.T_eau_exploit
    const SRT_design = f('SRT_design') ?? SRT_reference
    const biogaz_teneur_CH4 = p.biogaz_teneur_CH4 ?? 0.63
    const correction_T = correctionT(T_digesteur)

    // densité du biogaz en fonction de sa teneur en méthane, puis volume de
    // biogaz produit par kg de MV détruite
    let biogaz_densite = 0
    for (let i = 0; i <= 1; i++) biogaz_densite += H.a_densite_biogaz[i] * Math.pow(biogaz_teneur_CH4, i)
    const ratio_biogaz_MV = biogaz_densite > 0 ? 1 / biogaz_densite : 0

    const entrees = entreesDepuisChoix(choices, p, ['entree_1', 'entree_2', 'entree_3'])
    if (!entrees.length) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Aucune boue admise : configurer au moins une entrée.'] }
    }

    /**
     * Une passe de digestion : lit les entrées, détruit les MV et remplit
     * l'étape aval. Employée deux fois — au SRT de dimensionnement puis au SRT
     * réel — comme dimensionnement() et fonctionnement_reel() dans le VBA.
     */
    function passe(SRT, ecrire) {
      const correction_SRT = correctionSRT(SRT, SRT_reference)
      const lu = lireEntrees(table, soluble, entrees)
      let inlet_MES = 0, inlet_MV = 0, inlet_MS = 0, inlet_Q = 0
      let outlet_MES = 0, outlet_MS = 0, outlet_MV = 0, outlet_Q = 0
      let NK_soluble_in = 0, Pt_soluble_in = 0
      const MV_reduits = new Array(NB_TYPES + 1).fill(0)
      const parType = {}

      for (const e of lu) {
        const { j, MV_MES, ratios, sol, src } = e
        const Q = e.Q
        let MES = e.MES
        if (!(MES > 0)) continue
        let MV = MES * MV_MES
        let DCO = MES * ratios.DCO
        let NK = MES * ratios.NK
        let Pt = MES * ratios.Pt
        NK_soluble_in += (sol[SOL.NK] * Q) / 1000
        Pt_soluble_in += (sol[SOL.Pt] * Q) / 1000
        inlet_MES += MES; inlet_MS += MES; inlet_MV += MV; inlet_Q += Q

        const origine = src[P.origine]
        const red_ref = REDUCTION_MV[origine]
        if (red_ref == null && origine) {
          warnings.push(`Origine de boue « ${origine} » sans taux de réduction des MV de référence : boue considérée non dégradable.`)
        }
        const reduction_MV = (red_ref ?? 0) * correction_T * correction_SRT * MV
        MV_reduits[j] += reduction_MV

        let MS = MES
        MV -= reduction_MV
        MES -= reduction_MV
        MS -= reduction_MV
        // les MV détruites quittent la phase solide : le volume diminue
        // d'autant (masse volumique des boues prise à 1 000 kg/m³)
        const Qout = Q - reduction_MV / 1000
        outlet_MES += MES; outlet_MS += MS; outlet_MV += MV; outlet_Q += Qout

        if (!parType[j]) parType[j] = { MES: 0, MV: 0, Q: 0, DCO: 0, NK: 0, Pt: 0, flux_in: 0, origine }
        const t = parType[j]
        t.MES += MES; t.MV += MV; t.Q += Qout; t.DCO += DCO; t.NK += NK; t.Pt += Pt
        t.flux_in += e.flux_in
      }

      if (!(inlet_MES > 0) || !(inlet_Q > 0)) return null
      const reduction_globale = inlet_MV > 0 ? (inlet_MV - outlet_MV) / inlet_MV : 0
      const biogaz_Q = ratio_biogaz_MV * (inlet_MV - outlet_MV)
      const DCO_biogaz = H.ratio_CH4_DCO > 0 ? (biogaz_teneur_CH4 * biogaz_Q) / H.ratio_CH4_DCO : 0
      // part de MV détruites imputable à chaque type de boue
      const deltaMV = inlet_MV - outlet_MV
      for (let j = 1; j <= NB_TYPES; j++) MV_reduits[j] = deltaMV > 0 ? MV_reduits[j] / deltaMV : 0

      const res = {
        inlet_MES, inlet_MV, inlet_MS, inlet_Q,
        outlet_MES, outlet_MS, outlet_MV, outlet_Q,
        inlet_concentration: inlet_MS / inlet_Q,
        inlet_MV_MES: inlet_MS > 0 ? inlet_MV / inlet_MS : 0,
        outlet_MV_MES: outlet_MS > 0 ? outlet_MV / outlet_MS : 0,
        outlet_concentration: outlet_Q > 0 ? outlet_MS / outlet_Q : 0,
        reduction_globale, biogaz_Q, DCO_biogaz, correction_SRT,
      }
      if (!ecrire) return res

      // ---- écriture de l'étape aval et compartimentation de la pollution
      const NK_sol_mgL = outlet_Q > 0 ? (NK_soluble_in / outlet_Q) * 1000 : 0
      const Pt_sol_mgL = outlet_Q > 0 ? (Pt_soluble_in / outlet_Q) * 1000 : 0
      for (const [jStr, t] of Object.entries(parType)) {
        const j = Number(jStr)
        const dst = table[etapeOut][j]
        dst[P.origine] = t.origine
        dst[P.MES] = t.MES
        dst[P.Q] = t.Q
        dst[P.MV_MES] = t.MES > 0 ? t.MV / t.MES : 0
        dst[P.flux_in] = t.flux_in
        dst[P.verif_flux] = t.flux_in
        if (!(t.MES > 0) || !(t.Q > 0)) continue
        const s = soluble[etapeOut][j]
        s[SOL.MS_soluble] = outlet_Q > 0 ? ((outlet_MS - outlet_MES) / outlet_Q) * 1000 : 0
        s[SOL.MV_soluble] = res.outlet_MV_MES * s[SOL.MS_soluble]
        // la DCO du biogaz est retranchée au prorata des MV détruites par type
        let DCO = t.DCO - MV_reduits[j] * DCO_biogaz
        let dissous = H.ratio_DCOsol_outlet * DCO * correction_T
        s[SOL.DCO] = (dissous / t.Q) * 1000
        dst[P.ratio_DCO_MES] = (DCO - dissous) / t.MES
        dst[P.ratio_DBO_MES] = 0 // hypothèse d'OCEAN : boues digérées non biodégradables
        // azote libéré au prorata des MV dégradées de ce type de boue
        const src_in = table[ETAPE.inlet][j]
        const ratio_NK_MV = src_in[P.MV_MES] > 0 ? src_in[P.ratio_NK_MES] / src_in[P.MV_MES] : 0
        dissous = H.ratio_NK_MVdeg * MV_reduits[j] * deltaMV * ratio_NK_MV
        s[SOL.NK] = NK_sol_mgL + (dissous / t.Q) * 1000
        dst[P.ratio_NK_MES] = Math.max(0, (t.NK - dissous) / t.MES)
        dissous = t.Pt * H.ratio_P_Pinlet * correction_T
        s[SOL.Pt] = Pt_sol_mgL + (dissous / t.Q) * 1000
        dst[P.ratio_Pt_MES] = Math.max(0, (t.Pt - dissous) / t.MES)
      }
      // consommation du flux amont (bilan matière)
      for (const e of lu) e.src[P.verif_flux] -= e.flux_in
      return res
    }

    // ---- dimensionnement : volume déduit du temps de séjour visé
    const nominal = passe(SRT_design, false)
    if (!nominal) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Aucun flux de boues à l'entrée du digesteur."] }
    }
    let volume_total = nominal.inlet_Q * SRT_design
    const Vf = f('volume_total')
    if (Vf != null) {
      if (Math.abs((volume_total - Vf) / Vf) > H_POMPE.critere_incoherence) {
        warnings.push(`Volume forcé (${Math.round(Vf)} m³) éloigné de plus de 10 % du volume nécessaire au temps de séjour visé (${Math.round(volume_total)} m³).`)
      }
      volume_total = Vf
    }

    // ---- fonctionnement réel : le temps de séjour découle du volume retenu
    const SRT_reel = nominal.inlet_Q > 0 ? volume_total / nominal.inlet_Q : SRT_reference
    const reel = passe(SRT_reel, true)

    // ---- bilan thermique : pertes de l'enceinte + chauffage des boues
    const pertes = H.pertes_thermiques[choices.climat][choices.isolation] * volume_total
    const chauffage = (reel.inlet_Q * H.Cp_boues * (T_digesteur - T_inlet)) / H.efficacite_echangeur
    const energie_maintien = f('energie_maintien') ?? pertes + chauffage
    const niveau_chaleur = T_digesteur < H.T_limite_BT_HT ? 'basse température' : 'haute température'

    // ---- calcul_consommation_electrique
    let electricite_agitation = 0
    if (choices.agitation === 'mecanique') {
      electricite_agitation = (H.agitation_W_m3 * CONST.NOMBRE_HEURE_PAR_JOUR * volume_total) / 1000
    } else if (choices.agitation === 'recirculation') {
      const tps = CONST.NOMBRE_HEURE_PAR_JOUR
      const HMT = 10
      const Q = (p.tx_recirculation_agitation * volume_total) / tps
      const puissance = (Q / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * HMT / 0.4
      electricite_agitation = (puissance / rendementMoteur(puissance)) * tps
    } else {
      const tps = CONST.NOMBRE_HEURE_PAR_JOUR
      const hauteur = p.hauteur_digesteur
      const Q = hauteur > 0 ? (H.brassage_biogaz_Nm_h * tps * volume_total) / hauteur : 0
      electricite_agitation = (H.conso_surpresseur * Q * (hauteur + 2)) / 1000
    }

    let electricite_recirculation = 0
    if (choices.chauffage_boues === 'recirculation') {
      const tps = CONST.NOMBRE_HEURE_PAR_JOUR
      const HMT = 10
      const Q = energie_maintien / (H.Cp_boues * H.delta_T_boues) / tps
      const puissance = (Q / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * HMT / 0.4
      electricite_recirculation = (puissance / rendementMoteur(puissance)) * tps
    }

    const extr = repartitionPompage(reel.outlet_Q, p.extraction_pompe_nb, p.extraction_tps_fonctionnement, forced, 'extraction', H.tps_fct_pompe_extraction)
    if (extr.incoherence) warnings.push("Incohérence sur le pompage d'extraction.")
    const rdt = f('extraction_pompe_rdt') ?? rdtPompeBoues(extr.Qu, p.extraction_P_refoulement)
    const electricite_extraction = rdt > 0
      ? (extr.Qu / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * p.extraction_P_refoulement / rdt * extr.nb * extr.tps
      : 0

    let electricite_stockage = 0
    let volume_gazometre = 0
    if (choices.biogaz_stock === 'oui') {
      const SRT_gaz = choices.valorisation_biogaz === 'chaudiere' ? 3 : 8 // h
      volume_gazometre = (nominal.biogaz_Q / CONST.NOMBRE_HEURE_PAR_JOUR) * SRT_gaz
      const soufflage = volume_gazometre < 1225 ? 300 : 500 // Nm³/h
      electricite_stockage = (H.conso_surpresseur * soufflage * 0.25 * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
    }
    const soufflage_valo = (reel.biogaz_Q * p.ratio_biogaz_valorise) / CONST.NOMBRE_HEURE_PAR_JOUR
    const HMT_valo = choices.valorisation_biogaz === 'chaudiere' ? 1 : 3
    const electricite_surpresseur = (H.conso_surpresseur * soufflage_valo * HMT_valo * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000

    const total = electricite_agitation + electricite_recirculation + electricite_extraction + electricite_stockage + electricite_surpresseur

    if (SRT_reel < 12) warnings.push(`Temps de séjour réel faible (${SRT_reel.toFixed(1)} j) : risque de lessivage de la biomasse méthanogène.`)
    if (correction_T === 0) warnings.push(`Température de digestion (${T_digesteur} °C) hors du domaine de fonctionnement : réduction des MV nulle.`)

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'MES en entrée', unit: 'kg/j', value: reel.inlet_MES },
        { key: 'in_Q', label: 'Débit en entrée', unit: 'm³/j', value: reel.inlet_Q },
        { key: 'in_conc', label: 'Siccité en entrée', unit: 'g/L', value: reel.inlet_concentration },
        { key: 'in_MV', label: 'MV en entrée', unit: 'kg/j', value: reel.inlet_MV },
        { key: 'V', label: 'Volume total de digestion', unit: 'm³', value: volume_total },
        { key: 'SRT_design', label: 'Temps de séjour de dimensionnement', unit: 'j', value: SRT_design },
        { key: 'SRT_reel', label: 'Temps de séjour réel', unit: 'j', value: SRT_reel },
        { key: 'T', label: 'Température du digesteur', unit: '°C', value: T_digesteur },
        { key: 'corr_T', label: 'Correctif de température', unit: '-', value: correction_T },
        { key: 'corr_SRT', label: 'Correctif de temps de séjour', unit: '-', value: reel.correction_SRT },
        { key: 'red_MV', label: 'Réduction des matières volatiles', unit: '-', value: reel.reduction_globale },
        { key: 'out_MES', label: 'MES digérées', unit: 'kg/j', value: reel.outlet_MES },
        { key: 'out_Q', label: 'Débit de boues digérées', unit: 'm³/j', value: reel.outlet_Q },
        { key: 'out_conc', label: 'Siccité des boues digérées', unit: 'g/L', value: reel.outlet_concentration },
        { key: 'out_MV_MES', label: 'MV/MS des boues digérées', unit: '-', value: reel.outlet_MV_MES },
        { key: 'biogaz', label: 'Production de biogaz', unit: 'Nm³/j', value: reel.biogaz_Q },
        { key: 'biogaz_nom', label: 'Production de biogaz (dimensionnement)', unit: 'Nm³/j', value: nominal.biogaz_Q },
        { key: 'CH4', label: 'Teneur en méthane', unit: '-', value: biogaz_teneur_CH4 },
        { key: 'CH4_Q', label: 'Méthane produit', unit: 'Nm³/j', value: reel.biogaz_Q * biogaz_teneur_CH4 },
        { key: 'DCO_biogaz', label: 'DCO méthanisée', unit: 'kg/j', value: reel.DCO_biogaz },
        ...(volume_gazometre > 0 ? [{ key: 'V_gaz', label: 'Volume du gazomètre', unit: 'm³', value: volume_gazometre }] : []),
        { key: 'therm', label: `Énergie thermique de maintien (${niveau_chaleur})`, unit: 'kWh/j', value: energie_maintien },
        { key: 'therm_pertes', label: "dont pertes de l'enceinte", unit: 'kWh/j', value: pertes },
        { key: 'therm_chauffage', label: 'dont chauffage des boues', unit: 'kWh/j', value: chauffage },
      ],
      energie: {
        biogaz_Nm3j: reel.biogaz_Q,
        biogaz_CH4: biogaz_teneur_CH4,
        besoin_thermique_kWhj: energie_maintien,
        niveau: T_digesteur < H.T_limite_BT_HT ? 'BT' : 'HT',
      },
      electricity: {
        total,
        fixed: electricite_agitation + electricite_recirculation,
        detail: {
          agitation: electricite_agitation,
          recirculation_chauffage: electricite_recirculation,
          extraction: electricite_extraction,
          stockage_biogaz: electricite_stockage,
          surpresseur_biogaz: electricite_surpresseur,
        },
      },
      warnings,
    }
  },
})
