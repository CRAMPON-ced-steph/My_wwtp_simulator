// ---------------------------------------------------------------------------
// Port de E6_MBBR.cls — réacteur à lit mobile (Moving Bed Biofilm Reactor).
//
// Le procédé est une chaîne de 10 cuves fixes, dont chacune peut être active ou
// non (constantes cuve_* du VBA) :
//   1 pré-dénit 1   2 pré-dénit 2   3 C1   4 C2   5 N1   6 N2
//   7 dé-ox         8 post-dénit 1  9 post-dénit 2      10 ré-ox
// et de 14 flux intermédiaires (constantes flux_* du VBA), le flux 9 étant la
// recirculation de liqueur mixte prélevée en sortie de dé-ox.
//
// La biomasse est entièrement fixée sur le média en suspension : le
// dimensionnement se fait sur des vitesses surfaciques (g/(m²·j)) et non sur
// une charge massique. Chaque sous-routine dimensionne_* calcule un
// « volume de média nécessaire » puis résout le couple (volume de cuve, taux de
// remplissage) par le même arbre de décision — factorisé ici dans
// resoudreVolumeFilling().
//
// Deux passes, comme dans le classeur VBA :
//   dimensionnement()      sur l'eau nominale, à T_design  → volumes, remplissages
//   fonctionnement_reel()  sur l'eau réelle, à T_exploit   → performances, air
// Chaque passe contient une boucle de point fixe sur la recirculation de MLSS,
// et les sous-routines simule_* résolvent leur équation par Newton-Raphson.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - garde d'itérations sur toutes les boucles (VBA : boucles non bornées) ;
//  - le choix du média / de l'aérateur / du mode d'agitation est proposé par
//    étape et non cuve par cuve (le calcul, lui, reste par cuve) ;
//  - les variables de classe utilisées avant d'être calculées (boues_Q,
//    reox_DBOremoved, boues_postdenit, deox_nit, flux post-séparation) valent 0
//    au premier passage : reproduit fidèlement.
// ---------------------------------------------------------------------------
import { defineNode } from '../core/engine.js'
import { makeStream } from '../core/stream.js'
import { CONST, HYP, besoinsO2HS, ratio } from '../core/hypotheses.js'

// --- repères de cuves (Private Const cuve_*) -------------------------------
const C = { predenit1: 1, predenit2: 2, C1: 3, C2: 4, N1: 5, N2: 6, deox: 7, postdenit1: 8, postdenit2: 9, reox: 10 }
const NB_CUVES = 10
// --- repères de flux (Private Const flux_*) --------------------------------
const F = {
  inlet: 0, inlet_recirculation: 1, predenit1_predenit2: 2, predenit2_C1: 3,
  C1_C2: 4, C2_N1: 5, N1_N2: 6, N2_deox: 7, deox_separationMLSS: 8,
  recirculation_MLSS: 9, separationMLSS_postdenit1: 10, postdenit1_postdenit2: 11,
  postdenit2_reox: 12, outlet: 13,
}
const NB_FLUX = 13
// --- repères de paramètres de flux (Private Const repere_*) ----------------
const R = { Q: 1, DCO: 2, DBO: 3, MES: 4, NK: 5, NH4: 6, NO3: 7, Pt: 8, O2: 9 }
const NB_PARAM = 9

// --- médias : surface spécifique (m²/m³) et vitesse d'air mini d'agitation --
const MEDIA = {
  K1: { label: 'AnoxKaldnes K1', surface: 500, Vmin: 10 },
  K1_heavy: { label: 'AnoxKaldnes K1 heavy', surface: 500, Vmin: 10 },
  K3: { label: 'AnoxKaldnes K3', surface: 500, Vmin: 10 },
  K5: { label: 'AnoxKaldnes K5', surface: 800, Vmin: 10 },
  chip_P: { label: 'Chip P', surface: 900, Vmin: 15 },
  chip_M: { label: 'Chip M', surface: 1200, Vmin: 10 },
}
const MEDIA_OPTIONS = Object.entries(MEDIA).map(([value, m]) => ({ value, label: m.label }))
const estK = (m) => m === 'K1' || m === 'K1_heavy' || m === 'K3' || m === 'K5'
const estChip = (m) => m === 'chip_P' || m === 'chip_M'

// --- hypothèses figées de la classe (Sub hypotheses) -----------------------
const H = {
  critere_convergence: 1e-10,
  critere_volume_nul: 1e-5,
  ratio_O2_nit: 4.57,
  ratio_O2_denit: -2.86,
  ratio_O2_DBO: 0.75,
  ratio_NO3eq_O2: 0.35,
  ratio_N_assimile: 0.05,
  ratio_N_assimile_Capporte: 0.08,
  ratio_P_assimile: 0.01,
  ratio_DCO_dure_total: 0.33 * 0.13,
  ratio_boues_DBO_apportee: 0.25, // kg MES / kg O2
  rate_DBO_source_C: 0.96, // méthanol
  ratio_DCO_DBO_source_C: 1.5 / 0.96,
  DBO_apportee_fuite_mgL: 5,
  delta_NO3_post_DN: 10,
  predenit_NO3_limite: 3,
  postdenit_NO3_limite: 2,
  hyp_rate_DBOin_appliquee_10_CN: 3.9,
  hyp_outlet_DBO_mgL_mini: 10,
  cuve_HRTmini_h: 15 / CONST.NOMBRE_MINUTE_PAR_HEURE,
  ratio_elec_agitation: 7.5, // W/m³ des cuves anoxiques
  agitation_fct: 24, // h/j
  O2_facteur_beta: 0.95,
  hauteur_diffuseur: 0.25,
  correctif_T_K: 1.024,
  T_ref_K: 20,
  ratio_kgO2_Nm3air: 0.3,
  // Mox
  rate_N_30: { K3: 2.8, K5: 2.3, chip_M: 2 },
  ratio_NO3f_nit_Mox: 0.11,
  ratio_O2_nit_Mox: 1.94,
  cuve_NH4_outlet_Mox_mgL: 10,
}
const EXCES_O2_REGULATION = { horloge: 2, sans_variateur: 1, variateur: 0.5, avance: 0 }
const SEP = { clarif: 'clarif', clarif_polymere: 'clarif_polymere', clarif_polymere_coag: 'clarif_polymere_coag' }

// --- corrections de température (motifs répétés de la classe) --------------
/** 1,07 sous 10 °C puis 1,06, plafonné à 25 °C — DBO, dénitrification */
const corrT_DBO = (T) => (T <= 10 ? Math.pow(1.07, T - 10) : T <= 25 ? Math.pow(1.06, T - 10) : Math.pow(1.06, 15))
/** 1,07 plafonné à 25 °C — nitrification */
const corrT_nit = (T) => (T <= 25 ? Math.pow(1.07, T - 10) : Math.pow(1.07, 15))
/** respiration endogène de la biomasse fixée (kg O2/j) pour une surface S (m²) */
const respiration = (T, S) => (T >= 24 ? 5e-3 * 0.012 * 24 * S : 2e-3 * Math.pow(1.07, T - 10) * 0.012 * 24 * S)

// ---------------------------------------------------------------------------
// Génération de la liste des paramètres forçables, cuve par cuve.
// ---------------------------------------------------------------------------
const NOM_CUVE = {
  1: 'Pré-dénit 1', 2: 'Pré-dénit 2', 3: 'Carbone 1', 4: 'Carbone 2',
  5: 'Nitrif 1', 6: 'Nitrif 2', 7: 'Dé-Ox', 8: 'Post-dénit 1', 9: 'Post-dénit 2', 10: 'Ré-Ox',
}
const EST_AEREE = (i) => (i >= C.C1 && i <= C.N2) || i === C.reox
function paramsCuves() {
  const out = []
  for (let i = 1; i <= NB_CUVES; i++) {
    const g = `Cuve ${i} — ${NOM_CUVE[i]}`
    out.push({ key: `cuve${i}_volume`, label: 'Volume de la cuve', unit: 'm³', group: g, default: undefined, hint: 'dimensionné si non forcé' })
    out.push({ key: `cuve${i}_filling`, label: 'Taux de remplissage en média', unit: '-', group: g, default: undefined, hint: 'dimensionné si non forcé' })
    if (EST_AEREE(i) || i === C.deox) out.push({ key: `cuve${i}_O2_dissous`, label: 'O2 dissous', unit: 'mg/L', group: g, default: undefined })
    if (EST_AEREE(i)) {
      out.push({ key: `cuve${i}_alpha`, label: 'Facteur alpha', unit: '-', group: g, default: undefined })
      out.push({ key: `cuve${i}_rdt_transfert`, label: "Rendement de transfert de l'O2", unit: '%/m', group: g, default: undefined, hint: 'corrélation OTR média si non forcé' })
      out.push({ key: `cuve${i}_hauteur`, label: "Hauteur d'eau", unit: 'm', group: g, default: undefined })
      out.push({ key: `cuve${i}_Vair_agitation`, label: "Débit d'air surfacique de brassage", unit: 'Nm³/(m²·h)', group: g, default: undefined })
      out.push({ key: `cuve${i}_brassage_air_hj`, label: 'Durée journalière de brassage par air', unit: 'h/j', group: g, default: undefined })
    }
    out.push({ key: `cuve${i}_agitation_W_m3`, label: "Ratio d'agitation mécanique", unit: 'W/m³', group: g, default: undefined })
  }
  return out
}

const OUI_NON = [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }]

// ---------------------------------------------------------------------------
export default defineNode({
  id: 'mbbr',
  label: 'MBBR',
  short: 'MBBR',
  family: 'secondaire',
  vba: 'E6_MBBR.cls',
  description:
    "Réacteur à lit mobile : cuves de pré-dénitrification, d'élimination du carbone, de nitrification, de dé-oxygénation, de post-dénitrification au méthanol et de ré-oxygénation, avec biomasse fixée sur média en suspension et séparation liquide/solide en aval.",
  choices: [
    { key: 'etape_predenit', label: 'Pré-dénitrification', default: 'oui', options: OUI_NON },
    { key: 'etape_C', label: 'Élimination du carbone', default: 'oui', options: OUI_NON },
    { key: 'etape_N', label: 'Nitrification', default: 'oui', options: OUI_NON },
    { key: 'etape_deox', label: 'Dé-oxygénation', default: 'oui', options: OUI_NON },
    { key: 'etape_postdenit', label: 'Post-dénitrification (méthanol)', default: 'non', options: OUI_NON },
    { key: 'etape_reox', label: 'Ré-oxygénation', default: 'non', options: OUI_NON },
    { key: 'mode_Mox', label: 'Fonctionnement Mox (nitritation sur N1)', default: 'non', options: [{ value: 'non', label: 'non' }, { value: 'oui', label: 'oui' }] },
    { key: 'media_predenit', label: 'Média des cuves de pré-dénit', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'media_C', label: 'Média des cuves carbone', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'media_N', label: 'Média des cuves de nitrification', default: 'K5', options: MEDIA_OPTIONS },
    { key: 'media_deox', label: 'Média de la cuve de dé-ox', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'media_postdenit', label: 'Média des cuves de post-dénit', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'media_reox', label: 'Média de la cuve de ré-ox', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'aerateur', label: 'Type de diffuseurs', default: 'fine_bulle', options: [{ value: 'fine_bulle', label: 'fines bulles' }, { value: 'moyenne_bulle', label: 'moyennes bulles' }] },
    { key: 'agitation_cuve_aeree', label: 'Agitation des cuves aérées', default: 'air', options: [{ value: 'air', label: "par l'air" }, { value: 'mecanique', label: 'mécanique' }] },
    { key: 'type_separation', label: 'Séparation liquide/solide aval', default: 'clarif', options: [
      { value: 'clarif', label: 'clarificateur seul' },
      { value: 'clarif_polymere', label: 'clarificateur + polymère' },
      { value: 'clarif_polymere_coag', label: 'clarificateur + polymère + coagulant' },
    ] },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
      { value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' },
    ] },
    { key: 'regulation', label: "Régulation de l'aération", default: 'variateur', options: [
      { value: 'horloge', label: 'sur horloge' }, { value: 'sans_variateur', label: 'sans variateur' }, { value: 'variateur', label: 'classique – avec variateur' }, { value: 'avance', label: 'avancé' },
    ] },
  ],
  params: [
    { key: 'nominal_outlet_NH4_mgL', label: 'N-NH4 en sortie (nominal)', unit: 'mg/L', group: 'Objectifs de traitement', default: 3 },
    { key: 'nominal_outlet_NO3_mgL', label: 'N-NO3 en sortie (nominal)', unit: 'mg/L', group: 'Objectifs de traitement', default: undefined, hint: '5 avec post-dénit, sinon 15' },
    { key: 'nominal_outlet_DBO_mgL', label: 'DBO5 en sortie (nominal)', unit: 'mg/L', group: 'Objectifs de traitement', default: 20 },
    { key: 'Norga_dure_soluble', label: 'N organique dur et soluble', unit: 'mg/L', group: 'Objectifs de traitement', default: 1.5 },
    { key: 'nominal_predenit_NO3_mgL', label: 'N-NO3 en sortie de pré-dénitrification', unit: 'mg/L', group: 'Objectifs de traitement', default: 3 },
    { key: 'nominal_last_aerated_tank_NO3_mgL', label: 'N-NO3 en sortie de la dernière cuve aérée', unit: 'mg/L', group: 'Objectifs de traitement', default: undefined, hint: 'NO3 sortie + 10 avec post-dénit' },
    { key: 'ratio_elimination_predenit1_predenit12', label: 'Répartition pré-dénit 1 / (1+2)', unit: '-', group: 'Répartition entre cuves', default: undefined, hint: '0,5 si deux cuves' },
    { key: 'ratio_elimination_C1_C12', label: 'Répartition carbone 1 / (1+2)', unit: '-', group: 'Répartition entre cuves', default: undefined, hint: '0,7 si deux cuves' },
    { key: 'ratio_elimination_N1_N12', label: 'Répartition nitrif 1 / (1+2)', unit: '-', group: 'Répartition entre cuves', default: undefined, hint: '0,5 si deux cuves' },
    { key: 'ratio_elimination_postdenit1_postdenit12', label: 'Répartition post-dénit 1 / (1+2)', unit: '-', group: 'Répartition entre cuves', default: undefined, hint: '0,5 si deux cuves' },
    { key: 'nominal_recirculation_MLSS_taux', label: 'Taux de recirculation de liqueur mixte (nominal)', unit: '-', group: 'Recirculation', default: undefined, hint: 'calculé sur le bilan NO3' },
    { key: 'reel_recirculation_MLSS_taux', label: 'Taux de recirculation de liqueur mixte (réel)', unit: '-', group: 'Recirculation', default: undefined, hint: 'repris du nominal si non forcé' },
    { key: 'recirculation_MLSS_P_refoulement', label: 'Pression de refoulement de la recirculation', unit: 'mCE', group: 'Recirculation', default: 5 },
    { key: 'recirculation_MLSS_pompe_rdt', label: 'Rendement global des pompes de recirculation', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
    { key: 'rate_N_T_design', label: 'Vitesse de nitritation Mox (dimensionnement)', unit: 'g N/(m²·j)', group: 'Mox', default: undefined, hint: 'rate_N_30 × 1,07^(T−30)' },
    { key: 'rate_N_T_reel', label: 'Vitesse de nitritation Mox (réel)', unit: 'g N/(m²·j)', group: 'Mox', default: undefined, hint: 'rate_N_30 × 1,07^(T−30)' },
    ...paramsCuves(),
    { key: 'volume_total_bassins', label: 'Volume total des cuves', unit: 'm³', group: 'Post-dénitrification', default: undefined, hint: 'somme des volumes si non forcé' },
    { key: 'postdenit_carbone_apporte_flux', label: 'Consommation de méthanol', unit: 'kg/j', group: 'Post-dénitrification', default: undefined, hint: 'calculée si non forcée' },
    { key: 'diffuseur_encrassement', label: 'Ancienneté des diffuseurs', unit: 'an(s)', group: 'Aération', default: 0 },
    { key: 'air_P_refoulement_moyenne', label: 'Pression de refoulement des surpresseurs', unit: 'mCE', group: 'Aération', default: undefined, hint: 'hauteur max + 2 + 0,25 × encrassement' },
    { key: 'surpresseur_conso_spec', label: 'Consommation spécifique des surpresseurs', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'O2_besoin_total', label: 'Besoin total en O2 (réel)', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'air_Q_Nm3j', label: "Débit d'air process (réel)", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'reel_outlet_DBO_mgL', label: 'DBO5 en sortie (réel)', unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'calculée si non forcée' },
    { key: 'reel_outlet_NH4_mgL', label: 'N-NH4 en sortie (réel)', unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'calculée si non forcée' },
    { key: 'reel_outlet_NO3_mgL', label: 'N-NO3 en sortie (réel)', unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_MES', label: "MES de l'effluent", unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'corrélation sur la DBO de sortie' },
    { key: 'boues_concentration', label: 'Concentration des boues extraites', unit: 'g/L', group: 'Boues', default: 20 },
    { key: 'boues_MES', label: 'Boues à extraire (réel)', unit: 'kg MES/j', group: 'Boues', default: undefined, hint: 'calculées si non forcées' },
    { key: 'boues_MV_MES', label: 'Rapport MV/MES des boues', unit: '-', group: 'Boues', default: undefined, hint: 'calculé si non forcé' },
    { key: 'extraction_P_refoulement', label: "Pression de refoulement de l'extraction", unit: 'mCE', group: 'Boues', default: 5 },
    { key: 'extraction_pompe_rdt', label: "Rendement global des pompes d'extraction", unit: '-', group: 'Boues', default: 0.7 * 0.88 },
  ],

  compute(ctx) {
    const { site, choices, forced } = ctx
    const warnings = []
    const T_design = site.T_eau_design
    const T_reel = site.T_eau_exploit
    /** valeur forcée (équivalent des drapeaux *_force du VBA) ou undefined */
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)

    // =======================================================================
    // lecture_choix — cuves actives et médias, cuve par cuve
    // =======================================================================
    const on = (v) => v === 'oui'
    const choix_cuve = new Array(NB_CUVES + 1).fill(0)
    choix_cuve[C.predenit1] = on(choices.etape_predenit) ? 1 : 0
    choix_cuve[C.C1] = on(choices.etape_C) ? 1 : 0
    choix_cuve[C.N1] = on(choices.etape_N) ? 1 : 0
    choix_cuve[C.deox] = on(choices.etape_deox) ? 1 : 0
    choix_cuve[C.postdenit1] = on(choices.etape_postdenit) ? 1 : 0
    choix_cuve[C.reox] = on(choices.etape_reox) ? 1 : 0
    // nb_etapes_traitement vaut 1 quand seul le carbone est traité, ce qui
    // bascule la classe en configuration « C seul » (deux cuves C en série).
    const nb_etapes_traitement = [C.predenit1, C.C1, C.N1, C.deox, C.postdenit1, C.reox].reduce((n, i) => n + choix_cuve[i], 0)
    const choix_media = new Array(NB_CUVES + 1).fill('K3')
    choix_media[C.predenit1] = choix_media[C.predenit2] = choices.media_predenit
    choix_media[C.C1] = choix_media[C.C2] = choices.media_C
    choix_media[C.N1] = choix_media[C.N2] = choices.media_N
    choix_media[C.deox] = choices.media_deox
    choix_media[C.postdenit1] = choix_media[C.postdenit2] = choices.media_postdenit
    choix_media[C.reox] = choices.media_reox
    const Smedia = (i) => MEDIA[choix_media[i]].surface
    const choix_fct_Mox = on(choices.mode_Mox)

    // contrôles de cohérence bloquants de lecture_choix
    const erreurs = []
    if (choix_cuve[C.predenit1] && !choix_cuve[C.N1]) erreurs.push('Une pré-dénitrification sans nitrification est impossible.')
    if (choix_cuve[C.postdenit1] && !choix_cuve[C.N1]) erreurs.push('Une post-dénitrification sans nitrification est impossible.')
    if (choix_cuve[C.reox] && !choix_cuve[C.postdenit1]) erreurs.push('Une ré-oxygénation sans post-dénitrification est impossible.')
    if (choix_fct_Mox && (choix_cuve[C.predenit1] || choix_cuve[C.C1])) erreurs.push("La configuration Mox est incompatible avec la pré-dénitrification et l'abattement du carbone.")
    if (choix_fct_Mox && !['K3', 'K5', 'chip_M'].includes(choix_media[C.N1])) erreurs.push('Média non utilisable en configuration Mox — choix possibles : K3, K5, Chip M.')
    if (erreurs.length) {
      return {
        outNominal: makeStream(ctx.inNominal), outReel: makeStream(ctx.inReel),
        results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: erreurs,
      }
    }

    // =======================================================================
    // hypotheses — compartimentation de l'eau brute, constante de nitrification
    // =======================================================================
    const primaire = ctx.upstream?.primaire === true
    const primaire_simple = primaire && ctx.upstream?.primaire_reactif !== true
    let EB_rate_DBOsol_DBO, EB_rate_Norgasol_Norga
    if (primaire) {
      if (primaire_simple) { EB_rate_DBOsol_DBO = 0.482; EB_rate_Norgasol_Norga = 0.489 }
      else { EB_rate_DBOsol_DBO = 0.699; EB_rate_Norgasol_Norga = 0.71 }
    } else { EB_rate_DBOsol_DBO = 0.406; EB_rate_Norgasol_Norga = 0.386 }
    const rate_Norgapart_hydrolyse = 0.95
    const rate_DBOpart_hydrolyse = 0.3
    let k_nit
    if (choix_cuve[C.predenit1]) k_nit = primaire ? (primaire_simple ? 0.53 : 0.58) : 0.47
    else k_nit = primaire ? (primaire_simple ? 0.47 : 0.58) : 0.4
    const rateDBOdenit = (T) => (T < 5 ? 5 : T > 15 ? 4 : 5 - 0.1 * (T - 5))
    const rate_DBO_apportee_denit_nominal = rateDBOdenit(T_design)
    const rate_DBO_apportee_denit_reel = rateDBOdenit(T_reel)
    const hyp_cuve_C1_DBOsol_mgL = choix_fct_Mox ? 10 : 5
    const hyp_cuve_N_deox_DBOsol_mgL = choix_fct_Mox ? 10 : 2
    const ratio_NO3f_nit = choix_fct_Mox ? H.ratio_NO3f_nit_Mox : 1

    // taux de remplissage maximum admissible par cuve et par média
    const fillingMax = (i, m) => {
      if (i === C.predenit1 || i === C.predenit2 || i === C.deox || i === C.postdenit1 || i === C.postdenit2) return estK(m) ? 0.5 : 0
      if (i === C.C1 || i === C.C2) return m === 'K3' || m === 'K5' ? 0.65 : m === 'chip_P' ? 0.55 : 0
      if (i === C.N1 || i === C.N2 || i === C.reox) return m === 'K3' || m === 'K5' ? 0.65 : estChip(m) ? 0.55 : 0
      return 0
    }
    const fillingGuide = fillingMax // hyp_cuve_filling = media_filling_max dans le VBA

    // =======================================================================
    // attribution_valeur_par_defaut
    // =======================================================================
    const exces_O2 = EXCES_O2_REGULATION[choices.regulation] ?? 0
    const nominal_outlet_NH4_mgL = f('nominal_outlet_NH4_mgL') ?? 3
    const nominal_outlet_NO3_mgL = f('nominal_outlet_NO3_mgL') ?? (choix_cuve[C.postdenit1] ? 5 : 15)
    const nominal_outlet_DBO_mgL = f('nominal_outlet_DBO_mgL') ?? 20
    const Norga_dure_soluble = f('Norga_dure_soluble') ?? 1.5
    let nominal_last_aerated_tank_NO3_mgL = 0
    if (choix_cuve[C.predenit1]) {
      nominal_last_aerated_tank_NO3_mgL = choix_cuve[C.postdenit1]
        ? (f('nominal_last_aerated_tank_NO3_mgL') ?? nominal_outlet_NO3_mgL + H.delta_NO3_post_DN)
        : nominal_outlet_NO3_mgL
    }
    const nominal_predenit_NO3_mgL = choix_cuve[C.predenit1] ? (f('nominal_predenit_NO3_mgL') ?? 3) : 0

    const cuve_ratio_elimination = new Array(NB_CUVES + 1).fill(0)
    /** répartition entre les deux cuves d'une étape ; ratio = 1 → une seule cuve */
    function repartition(cuve1, cuve2, ratioKey, defautDeuxCuves, deuxCuvesParDefaut) {
      const rf = f(ratioKey)
      let r12
      if (rf != null) {
        r12 = rf
        choix_cuve[cuve2] = r12 === 1 ? 0 : 1
      } else {
        const vf = f(`cuve${cuve2}_volume`)
        if (vf != null) choix_cuve[cuve2] = vf <= H.critere_volume_nul ? 0 : 1
        else choix_cuve[cuve2] = deuxCuvesParDefaut ? 1 : 0
        r12 = choix_cuve[cuve2] ? defautDeuxCuves : 1
      }
      cuve_ratio_elimination[cuve1] = r12
      cuve_ratio_elimination[cuve2] = 1 - r12
    }
    if (choix_cuve[C.predenit1]) repartition(C.predenit1, C.predenit2, 'ratio_elimination_predenit1_predenit12', 0.5, nominal_predenit_NO3_mgL < H.predenit_NO3_limite)
    if (choix_cuve[C.C1] && nb_etapes_traitement === 1) repartition(C.C1, C.C2, 'ratio_elimination_C1_C12', 0.7, true)
    else choix_cuve[C.C2] = 0
    if (choix_cuve[C.N1]) repartition(C.N1, C.N2, 'ratio_elimination_N1_N12', 0.5, true)
    if (choix_cuve[C.postdenit1]) repartition(C.postdenit1, C.postdenit2, 'ratio_elimination_postdenit1_postdenit12', 0.5, nominal_outlet_NO3_mgL < H.postdenit_NO3_limite)

    const boues_concentration = f('boues_concentration') ?? 20

    // état par cuve
    const cuve_volume = new Array(NB_CUVES + 1).fill(0)
    const cuve_filling = new Array(NB_CUVES + 1).fill(0)
    const cuve_O2_dissous = new Array(NB_CUVES + 1).fill(0)
    const cuve_alfa = new Array(NB_CUVES + 1).fill(0)
    const cuve_rdt_transfert = new Array(NB_CUVES + 1).fill(0)
    const cuve_hauteur = new Array(NB_CUVES + 1).fill(0)
    const cuve_agitation_W_m3 = new Array(NB_CUVES + 1).fill(0)
    const cuve_Vair_agitation = new Array(NB_CUVES + 1).fill(0)
    const cuve_brassage_air_hj = new Array(NB_CUVES + 1).fill(0)
    const cuve_debit_air_agitation = new Array(NB_CUVES + 1).fill(0)
    const cuve_suraeration_brassage = new Array(NB_CUVES + 1).fill(0)

    let hauteur_bassin_max = 0
    for (let i = C.C1; i <= C.C2; i++) {
      if (choix_cuve[i]) {
        cuve_O2_dissous[i] = f(`cuve${i}_O2_dissous`) ??
          (T_design <= 11 ? 3 + exces_O2 : T_design < 16 ? 2.5 + exces_O2 : 2 + exces_O2)
      }
      cuve_alfa[i] = f(`cuve${i}_alpha`) ?? (choices.aerateur === 'fine_bulle' ? 0.65 : 0.8)
      cuve_hauteur[i] = f(`cuve${i}_hauteur`) ?? 6
      if (cuve_hauteur[i] > hauteur_bassin_max) hauteur_bassin_max = cuve_hauteur[i]
      cuve_rdt_transfert[i] = f(`cuve${i}_rdt_transfert`) ?? 4
    }
    for (let i = C.N1; i <= C.N2; i++) {
      if (choix_cuve[i]) {
        cuve_O2_dissous[i] = f(`cuve${i}_O2_dissous`) ??
          (choix_fct_Mox && i === C.N1 ? 1.5 : T_design <= 20 ? 5 + exces_O2 : 4 + exces_O2)
      }
      cuve_alfa[i] = f(`cuve${i}_alpha`) ?? (choices.aerateur === 'fine_bulle' ? 0.75 : 0.8)
      cuve_hauteur[i] = f(`cuve${i}_hauteur`) ?? (choix_fct_Mox ? 7 : 6)
      if (cuve_hauteur[i] > hauteur_bassin_max) hauteur_bassin_max = cuve_hauteur[i]
      cuve_rdt_transfert[i] = f(`cuve${i}_rdt_transfert`) ?? 4
    }
    if (choix_cuve[C.deox]) cuve_O2_dissous[C.deox] = f(`cuve${C.deox}_O2_dissous`) ?? 2
    if (choix_cuve[C.reox]) {
      const i = C.reox
      cuve_O2_dissous[i] = f(`cuve${i}_O2_dissous`) ?? 1.5 + exces_O2
      cuve_alfa[i] = f(`cuve${i}_alpha`) ?? (choices.aerateur === 'fine_bulle' ? 0.65 : 0.8)
      cuve_hauteur[i] = f(`cuve${i}_hauteur`) ?? 6
      if (cuve_hauteur[i] > hauteur_bassin_max) hauteur_bassin_max = cuve_hauteur[i]
      cuve_rdt_transfert[i] = f(`cuve${i}_rdt_transfert`) ?? 4
    }
    for (let i = 1; i <= NB_CUVES; i++) {
      if (EST_AEREE(i) && choices.agitation_cuve_aeree === 'air') {
        cuve_Vair_agitation[i] = f(`cuve${i}_Vair_agitation`) ?? (cuve_hauteur[i] >= 5 ? 7 : MEDIA[choix_media[i]].Vmin)
        cuve_brassage_air_hj[i] = f(`cuve${i}_brassage_air_hj`) ?? 24
      }
    }
    for (let i = 1; i <= NB_CUVES; i++) {
      if (!EST_AEREE(i)) cuve_agitation_W_m3[i] = choix_cuve[i] ? (f(`cuve${i}_agitation_W_m3`) ?? H.ratio_elec_agitation) : 0
      else cuve_agitation_W_m3[i] = choices.agitation_cuve_aeree === 'mecanique' ? (f(`cuve${i}_agitation_W_m3`) ?? 3) : 0
    }
    const diffuseur_encrassement = f('diffuseur_encrassement') ?? 0
    const air_P_refoulement_moyenne = f('air_P_refoulement_moyenne') ?? hauteur_bassin_max + 2 + 0.25 * diffuseur_encrassement
    const recirculation_MLSS_P_refoulement = f('recirculation_MLSS_P_refoulement') ?? 5
    const recirculation_MLSS_pompe_rdt = f('recirculation_MLSS_pompe_rdt') ?? 0.7 * 0.88
    const ratio_elec_recirculation_MLSS = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * recirculation_MLSS_pompe_rdt)
    const extraction_P_refoulement = f('extraction_P_refoulement') ?? 5
    const extraction_pompe_rdt = f('extraction_pompe_rdt') ?? 0.7 * 0.88
    const ratio_elec_extraction = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * extraction_pompe_rdt)

    // =======================================================================
    // Fonctions internes
    // =======================================================================
    /** MV/MES de l'eau brute (Private Function MV_MES_influent) */
    const MV_MES_influent = (MESin, DCOin, DBOin) => {
      if (primaire) return primaire_simple ? 0.782 : 0.716
      if (!(DCOin > 0) || !(DBOin > 0)) return 0.4
      const v = 1.213 - MESin / (DCOin * 3) - 0.1429 * (DCOin / DBOin)
      return Math.min(0.8, Math.max(0.4, v))
    }
    /** production de boues aérobies (fct_boues_pdtes) */
    const fct_boues_pdtes = (DBO_eliminee, Nnitrifie) => 0.2 * DBO_eliminee + 0.42 * Nnitrifie

    /** rendement de transfert de l'O2 en présence de média (aeration_rdt_transfert) */
    function aerationRdtTransfert(i) {
      const m = choix_media[i]
      const fine = choices.aerateur === 'fine_bulle'
      let OTR0, a1, a2, a3, a4, fmin, fmax
      if (fine) {
        OTR0 = 17
        if (estChip(m)) { a1 = -12; a2 = 20; a3 = 2; a4 = 70; fmin = 10; fmax = 55 }
        else { a1 = -2; a2 = 7; a3 = 3; a4 = 10; fmin = 10; fmax = m === 'K5' ? 60 : 65 }
      } else {
        OTR0 = 9
        if (estChip(m)) { a1 = -4; a2 = 20; a3 = 7; a4 = 70; fmin = 10; fmax = 55 }
        else if (m === 'K5') { a1 = 1; a2 = 7; a3 = 3; a4 = 30; fmin = 10; fmax = 60 }
        else if (m === 'K3') { a1 = 5.5; a2 = 7; a3 = 3; a4 = 30; fmin = 10; fmax = 65 }
        else { a1 = 5; a2 = 7; a3 = 3; a4 = 30; fmin = 10; fmax = 65 }
      }
      const pct = cuve_filling[i] * 100
      const retenu = pct < fmin ? fmin : pct > fmax ? fmax : pct
      const OTR = OTR0 + (a1 * Math.pow(retenu + a2, a3)) / (Math.pow(retenu + a2, a3) + Math.pow(a4, a3))
      const MF = 1 + (0.2 * Math.pow(pct + 10, 5)) / (Math.pow(pct + 10, 5) + Math.pow(20, 5))
      return ((OTR * MF) / (H.ratio_kgO2_Nm3air * 1000)) * 100 // gO2/(Nm³·m) → %/m
    }

    /**
     * Arbre de décision commun aux six routines dimensionne_* : résout le couple
     * (volume de cuve, taux de remplissage) à partir du volume de média requis.
     */
    function resoudreVolumeFilling(i, volume_media_necessaire, Qin, HRTmini = H.cuve_HRTmini_h) {
      const pointe_m3h = (Qin * site.pointe_TP) / CONST.NOMBRE_HEURE_PAR_JOUR
      const filling_HRTmini = pointe_m3h > 0 ? volume_media_necessaire / (pointe_m3h * HRTmini) : Infinity
      const fmax = fillingMax(i, choix_media[i])
      const Vf = f(`cuve${i}_volume`)
      const Ff = f(`cuve${i}_filling`)
      const dire = (msg) => warnings.push(`${NOM_CUVE[i]} : ${msg}`)
      if (Vf != null) {
        cuve_volume[i] = Vf
        if (Ff != null) {
          cuve_filling[i] = Ff
          if (cuve_filling[i] > fmax) dire('taux de remplissage supérieur au maximum admissible pour ce média.')
          if (cuve_volume[i] < volume_media_necessaire / filling_HRTmini) dire('temps de séjour insuffisant au débit de pointe.')
          if (cuve_volume[i] * cuve_filling[i] < volume_media_necessaire) dire('volume de média insuffisant pour le traitement visé.')
        } else {
          cuve_filling[i] = cuve_volume[i] > 0 ? volume_media_necessaire / cuve_volume[i] : 0
          if (cuve_filling[i] > fmax) {
            cuve_filling[i] = fmax
            if (cuve_volume[i] < volume_media_necessaire / filling_HRTmini) dire('temps de séjour insuffisant au débit de pointe.')
            if (cuve_volume[i] * cuve_filling[i] < volume_media_necessaire) dire('volume de média insuffisant pour le traitement visé.')
          } else if (cuve_volume[i] < volume_media_necessaire / filling_HRTmini) {
            dire('temps de séjour insuffisant au débit de pointe.')
          }
        }
        return
      }
      if (Ff != null) {
        cuve_filling[i] = Ff
        cuve_volume[i] = Ff > 0 ? volume_media_necessaire / Ff : 0
        if (cuve_filling[i] > fmax) dire('taux de remplissage supérieur au maximum admissible pour ce média.')
        if (cuve_volume[i] < volume_media_necessaire / filling_HRTmini) dire('temps de séjour insuffisant au débit de pointe.')
      } else {
        const guide = fillingGuide(i, choix_media[i])
        if (filling_HRTmini < guide) {
          cuve_filling[i] = filling_HRTmini
          dire('taux de remplissage abaissé pour respecter le temps de séjour minimal.')
        } else cuve_filling[i] = guide
      }
      cuve_volume[i] = cuve_filling[i] > 0 ? volume_media_necessaire / cuve_filling[i] : 0
    }

    /**
     * Aération d'une cuve (Sub aeration_cuve) : besoins en O2 (DBO, nitrification,
     * sulfures, respiration, variation d'O2 dissous), facteur K, débit d'air et
     * suraération éventuelle pour assurer le brassage du média.
     */
    function aerationCuve(i, flux, fin, T, DBO_removed, N_removed, Sh) {
      const S = cuve_volume[i] * cuve_filling[i] * Smedia(i)
      let besoin = choix_fct_Mox && i === C.N1
        ? H.ratio_O2_DBO * DBO_removed + H.ratio_O2_nit_Mox * N_removed
        : H.ratio_O2_DBO * DBO_removed + H.ratio_O2_nit * N_removed
      if (!choix_cuve[C.predenit1]) besoin += besoinsO2HS(Sh)
      besoin += respiration(T, S)
      besoin += flux[fin + 1][R.O2] - flux[fin][R.O2]

      const Patm = (HYP.Patm_P0 * Math.pow(HYP.Patm_a0 + HYP.Patm_a1 * site.altitude, HYP.Patm_a2)) / 100
      const delta_P = (0.35 / 10.33) * (HYP.Patm_P0 / 100) * (cuve_hauteur[i] - H.hauteur_diffuseur)
      let O2sat20 = 0, O2satT = 0
      for (let k = 0; k <= 4; k++) {
        O2sat20 += HYP.O2sat_coef[k] * Math.pow(20, k)
        O2satT += HYP.O2sat_coef[k] * Math.pow(T, k)
      }
      const corr = (Patm + delta_P) / (HYP.Patm_P0 / 100)
      O2sat20 *= corr; O2satT *= corr
      let K = (cuve_alfa[i] * (H.O2_facteur_beta * O2satT - cuve_O2_dissous[i])) / O2sat20
      K *= Math.pow(H.correctif_T_K, T - H.T_ref_K)

      cuve_rdt_transfert[i] = f(`cuve${i}_rdt_transfert`) ?? aerationRdtTransfert(i)
      const denom = K * (cuve_rdt_transfert[i] / 100) * (cuve_hauteur[i] - H.hauteur_diffuseur) * H.ratio_kgO2_Nm3air
      let air = denom > 0 ? besoin / denom : 0

      if (choices.agitation_cuve_aeree === 'air') {
        const section = cuve_hauteur[i] > 0 ? cuve_volume[i] / cuve_hauteur[i] : 0
        const mini = cuve_Vair_agitation[i] * section
        if (air / CONST.NOMBRE_HEURE_PAR_JOUR < mini) {
          cuve_suraeration_brassage[i] = (mini - air / CONST.NOMBRE_HEURE_PAR_JOUR) * cuve_brassage_air_hj[i]
          air += cuve_suraeration_brassage[i]
          warnings.push(`${NOM_CUVE[i]} : débit d'air augmenté pour assurer le brassage du média.`)
        } else cuve_suraeration_brassage[i] = 0
        cuve_debit_air_agitation[i] = mini * cuve_brassage_air_hj[i]
      } else {
        cuve_debit_air_agitation[i] = 0
        cuve_suraeration_brassage[i] = 0
      }
      return { besoin, air }
    }

    /** nitrification résiduelle dans une cuve carbone (nitrification_cuve_C) */
    function nitrificationCuveC(i, DBO_out_mgL, T) {
      const S = cuve_volume[i] * cuve_filling[i] * Smedia(i)
      const correction = Math.max(0, 1 - 0.125 * DBO_out_mgL)
      const base = (cuve_O2_dissous[i] - 0.5) / 3.2
      if (!(base > 0)) return 0
      return (S / 1000) * correction * k_nit * Math.pow(base, 0.7) * corrT_nit(T)
    }

    const nouveauFlux = () => Array.from({ length: NB_FLUX + 1 }, () => new Array(NB_PARAM + 1).fill(0))
    const copierFlux = (flux, dst, src) => { for (let j = 1; j <= NB_PARAM; j++) flux[dst][j] = flux[src][j] }

    // état de classe partagé entre les deux passes (0 au premier usage, comme en VBA)
    let boues_Q = 0
    let boues_postdenit = 0
    const reox_DBOremoved = 0 // déclaré mais jamais affecté dans le VBA
    const nominal_flux = nouveauFlux()

    // =======================================================================
    // dimensionnement (eau nominale, T de dimensionnement)
    // =======================================================================
    let nominal_recirculation_MLSS_taux = 0
    let nominal_boues_MV_MES = 0
    let nominal_sortie_MES = 0
    let surface_media_totale = 0
    let volume_total_bassins = 0
    let nominal_O2_besoin_total = 0
    let nominal_air_Q = 0
    let nominal_postdenit_carbone = 0
    let deox_nit = 0 // variable de classe : 0 à la première itération

    /** dimensionne_C_C : surface de média imposée par la DBO soluble visée */
    function dimensionneCC(flux) {
      let DBOsol, rate10
      const sep = choices.type_separation
      if (sep === SEP.clarif) {
        DBOsol = 0.5 * nominal_outlet_DBO_mgL - 5
        rate10 = nominal_outlet_DBO_mgL <= 21.7 ? 3.9 : nominal_outlet_DBO_mgL >= 37.5 ? 11.5 : 0.8346 * DBOsol - 0.273
      } else if (sep === SEP.clarif_polymere) {
        DBOsol = 0.5275 * nominal_outlet_DBO_mgL - 3
        rate10 = nominal_outlet_DBO_mgL <= 11.33 ? 3.9 : nominal_outlet_DBO_mgL >= 34.029 ? 11.5 : 0.7688 * DBOsol + 0.056
      } else {
        DBOsol = 0.3936 * nominal_outlet_DBO_mgL - 1.1553
        rate10 = nominal_outlet_DBO_mgL <= 10 ? 3.9 : nominal_outlet_DBO_mgL >= 24.895 ? 11.5 : 1.281 * DBOsol - 2.505
      }
      const rateT = rate10 * corrT_DBO(T_design)
      const DBO_eliminee = flux[F.predenit2_C1][R.DBO] - (DBOsol * flux[F.predenit2_C1][R.Q]) / 1000
      // charge appliquée = celle de l'eau brute, hors surcroît dû à la recirculation
      const DBO_appliquee = flux[F.inlet][R.DBO] - (flux[F.inlet_recirculation][R.DBO] - flux[F.predenit2_C1][R.DBO])
      const volume_media_total = rateT > 0 ? (DBO_appliquee * 1000) / rateT : 0
      return { DBO_eliminee, volume_media_total }
    }
    /** dimensionne_C_configCN : cuve carbone unique en amont d'une nitrification */
    function dimensionneCconfigCN(i, rate10, flux, Qin) {
      const rateT = rate10 * corrT_DBO(T_design)
      const V = rateT > 0 ? (1000 * flux[F.predenit2_C1][R.DBO]) / (rateT * Smedia(i)) : 0
      resoudreVolumeFilling(i, V, Qin)
    }
    /** dimensionne_predenit : vitesse de dénit fonction du rapport DBO disponible / NO3 */
    function dimensionnePredenit(i, NO3_out_mgL, DBOsol_EB, denit, Qin) {
      const DBO_conso = -H.ratio_O2_denit * denit
      const correction_NO3 = NO3_out_mgL / (NO3_out_mgL + 0.4)
      const rate10 = correction_NO3 * Math.exp(Math.log(0.3) * Math.pow(DBO_conso / DBOsol_EB, 2))
      const rateT = rate10 * corrT_DBO(T_design)
      const V = rateT > 0 ? (1000 * denit) / (rateT * Smedia(i)) : 0
      resoudreVolumeFilling(i, V, Qin)
    }
    /** dimensionne_N : vitesse de nitrification limitée par l'O2 ou par le NH4 */
    function dimensionneN(i, cuve_nit, NH4_out_mgL, Qin) {
      if (choix_fct_Mox && i === C.N1) {
        const rate = f('rate_N_T_design') ?? (H.rate_N_30[choix_media[i]] ?? 0) * Math.pow(1.07, T_design - 30)
        const V = rate > 0 ? (1000 * cuve_nit) / (rate * Smedia(i)) : 0
        resoudreVolumeFilling(i, V, Qin, 0.2)
        return
      }
      const lim = (cuve_O2_dissous[i] - 0.5) / 3.2
      const base = lim <= NH4_out_mgL ? lim : NH4_out_mgL
      const rate10 = base > 0 ? k_nit * Math.pow(base, 0.7) : 0
      const rateT = rate10 * corrT_nit(T_design)
      const V = rateT > 0 ? (1000 * cuve_nit) / (rateT * Smedia(i)) : 0
      resoudreVolumeFilling(i, V, Qin)
    }
    /**
     * dimensionne_deox : la surface de média est bornée par l'O2 disponible —
     * dichotomie pour que nitrification + respiration consomment exactement
     * l'oxygène apporté par l'eau.
     */
    function dimensionneDeox(nit, Qin) {
      const i = C.deox
      const base = (cuve_O2_dissous[i] - 0.5) / 3.2
      const rateT = (base > 0 ? k_nit * Math.pow(base, 0.7) : 0) * corrT_nit(T_design)
      const O2_conso_max = nit * H.ratio_O2_nit
      let Smin = 0
      let Smax = rateT > 0 ? (1000 * nit) / rateT : 0
      let S = (Smin + Smax) / 2
      let garde = 0
      while (S > 0 && (Smax - Smin) / S > H.critere_convergence && garde++ < 300) {
        const conso = (H.ratio_O2_nit * rateT * S) / 1000 + respiration(T_design, S)
        if (conso < O2_conso_max) Smin = S
        else if (conso > O2_conso_max) Smax = S
        else { Smin = S; Smax = S }
        S = (Smin + Smax) / 2
        deox_nit = (rateT * S) / 1000
      }
      resoudreVolumeFilling(i, S / Smedia(i), Qin)
    }
    /** dimensionne_postdenit : vitesse Monod 2,29·NO3/(NO3+2,9) */
    function dimensionnePostdenit(i, denit, NO3_mgL, Qin) {
      const rateT = ((2.29 * NO3_mgL) / (NO3_mgL + 2.9)) * corrT_DBO(T_design)
      const V = rateT > 0 ? (1000 * denit) / (rateT * Smedia(i)) : 0
      resoudreVolumeFilling(i, V, Qin)
    }
    /** dimensionne_reox : vitesse d'élimination de DBO fixée à 4 g/(m²·j) à 10 °C */
    function dimensionneReox(DBO_eliminee, Qin) {
      const rateT = 4 * corrT_DBO(T_design)
      const V = rateT > 0 ? (1000 * DBO_eliminee) / (rateT * Smedia(C.reox)) : 0
      resoudreVolumeFilling(C.reox, V, Qin)
    }

    /** separation_boues : MES de sortie, boues extraites et MV/MES */
    function separationBoues(flux, outlet_DBO_mgL, outlet_DBOsol_mgL, boues_aerobie, MV_MES_cumule, reel) {
      let sortie_MES = f('sortie_MES')
      if (sortie_MES == null) {
        if (nb_etapes_traitement === 1 && choix_cuve[C.C1]) {
          const sep = choices.type_separation
          sortie_MES = sep === SEP.clarif ? 0.5356 * outlet_DBO_mgL + 16.668
            : sep === SEP.clarif_polymere ? 0.5431 * outlet_DBO_mgL + 10.23
              : 0.8268 * outlet_DBO_mgL + 0.7594
        } else sortie_MES = (outlet_DBO_mgL - outlet_DBOsol_mgL) / 0.5555
      }
      const boues_bio = boues_aerobie + boues_postdenit
      const MV_MES = boues_bio > 0 ? MV_MES_cumule / boues_bio : 0
      const Qin = flux[F.inlet][R.Q]
      const force = reel ? f('boues_MES') : undefined
      const boues_MES = force ?? boues_bio - (sortie_MES * Qin) / 1000
      return { sortie_MES, boues_MES, MV_MES, boues_Q: boues_concentration > 0 ? boues_MES / boues_concentration : 0 }
    }

    function dimensionnement() {
      const s = ctx.inNominal
      const flux = nominal_flux
      nominal_boues_MV_MES = MV_MES_influent(s.MES, s.DCO, s.DBO) * s.MES
      flux[F.inlet][R.Q] = s.Q; flux[F.inlet][R.DCO] = s.DCO; flux[F.inlet][R.DBO] = s.DBO
      flux[F.inlet][R.MES] = s.MES; flux[F.inlet][R.NK] = s.NK; flux[F.inlet][R.NH4] = s.NH4
      flux[F.inlet][R.NO3] = s.NO3; flux[F.inlet][R.Pt] = s.Pt; flux[F.inlet][R.O2] = 0

      // charge à nitrifier au total ; les termes non encore calculés valent 0
      let N_nit_total = s.NH4 - ((nominal_outlet_NH4_mgL + Norga_dure_soluble) * (s.Q - boues_Q)) / 1000
        + EB_rate_Norgasol_Norga * (s.NK - s.NH4)
        + (1 - EB_rate_Norgasol_Norga) * rate_Norgapart_hydrolyse * (s.NK - s.NH4)
      N_nit_total -= H.ratio_N_assimile * (s.DBO - flux[F.separationMLSS_postdenit1][R.DBO])
        + H.ratio_N_assimile_Capporte * boues_postdenit
        + H.ratio_N_assimile * reox_DBOremoved

      const deltaNO3 = nominal_last_aerated_tank_NO3_mgL - nominal_predenit_NO3_mgL
      if (choix_cuve[C.predenit1]) {
        const base = (s.Q / 1000) * deltaNO3
        const taux0 = base > 0 ? (N_nit_total - base) / base : 0
        const Qr = taux0 * s.Q
        flux[F.recirculation_MLSS][R.Q] = Qr
        flux[F.recirculation_MLSS][R.DBO] = (nominal_outlet_DBO_mgL * Qr) / 1000
        flux[F.recirculation_MLSS][R.DCO] = flux[F.recirculation_MLSS][R.DBO]
        flux[F.recirculation_MLSS][R.NK] = ((nominal_outlet_NH4_mgL + Norga_dure_soluble) * Qr) / 1000
        flux[F.recirculation_MLSS][R.NH4] = (nominal_outlet_NH4_mgL * Qr) / 1000
        flux[F.recirculation_MLSS][R.NO3] = (nominal_last_aerated_tank_NO3_mgL * Qr) / 1000
        flux[F.recirculation_MLSS][R.Pt] = s.Q > 0 ? (s.Pt / s.Q) * Qr : 0
        flux[F.recirculation_MLSS][R.O2] = 0
      }

      const boues_pdtes = new Array(NB_CUVES + 1).fill(0)
      let nit_cuve_C = 0
      let I = 0 // reproduit la variable de boucle VBA réutilisée hors de sa boucle
      let eps = 1 + H.critere_convergence
      let garde = 0
      while (eps > H.critere_convergence && garde++ < 200) {
        surface_media_totale = 0
        nominal_O2_besoin_total = 0
        nominal_air_Q = 0
        let epsDBO = flux[F.recirculation_MLSS][R.DBO]
        let epsNH4 = flux[F.recirculation_MLSS][R.NH4]
        let epsNO3 = flux[F.recirculation_MLSS][R.NO3]
        for (let j = 1; j <= NB_PARAM; j++) flux[F.inlet_recirculation][j] = flux[F.inlet][j] + flux[F.recirculation_MLSS][j]

        // ---- PRE-DENITRIFICATION
        if (choix_cuve[C.predenit1]) {
          nominal_recirculation_MLSS_taux = f('nominal_recirculation_MLSS_taux') ??
            (deltaNO3 > 0 ? (N_nit_total - (s.Q / 1000) * deltaNO3) / ((s.Q / 1000) * deltaNO3) : 0)
          const predenit_denit = flux[F.inlet_recirculation][R.NO3]
            + H.ratio_NO3eq_O2 * flux[F.inlet_recirculation][R.O2]
            - (nominal_predenit_NO3_mgL * s.Q * (1 + nominal_recirculation_MLSS_taux)) / 1000
          let elim = 0
          for (I = C.predenit1; I <= C.predenit2; I++) {
            const fin = F.inlet_recirculation + (I - C.predenit1)
            const DBOsol_EB = flux[F.inlet][R.DBO] * (EB_rate_DBOsol_DBO + (1 - EB_rate_DBOsol_DBO) * rate_DBOpart_hydrolyse)
              + flux[F.recirculation_MLSS][R.DBO] - elim
            if (choix_cuve[I]) {
              const part = cuve_ratio_elimination[I] * predenit_denit
              let NO3out = flux[fin][R.NO3] + H.ratio_NO3eq_O2 * flux[fin][R.O2] - part
              NO3out = (NO3out * 1000) / ((1 + nominal_recirculation_MLSS_taux) * s.Q)
              dimensionnePredenit(I, NO3out, DBOsol_EB, part, flux[fin][R.Q])
              flux[fin + 1][R.Q] = flux[fin][R.Q]
              flux[fin + 1][R.NO3] = (NO3out * flux[fin + 1][R.Q]) / 1000
              elim = -H.ratio_O2_denit * part
              flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
              flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elim
              flux[fin + 1][R.NK] = flux[fin][R.NK] - H.ratio_N_assimile * elim
              flux[fin + 1][R.NH4] = flux[fin][R.NH4] - H.ratio_N_assimile * elim
              flux[fin + 1][R.Pt] = flux[fin][R.Pt] - H.ratio_P_assimile * elim
              flux[fin + 1][R.O2] = 0
              surface_media_totale += cuve_volume[I] * cuve_filling[I] * Smedia(I)
            } else copierFlux(flux, fin + 1, fin)
          }
        } else {
          nominal_recirculation_MLSS_taux = 0
          for (let k = F.predenit1_predenit2; k <= F.predenit2_C1; k++) copierFlux(flux, k, F.inlet_recirculation)
        }

        // ---- TRAITEMENT DU CARBONE
        nit_cuve_C = 0
        if (choix_cuve[C.C1] && nb_etapes_traitement === 1) {
          const { DBO_eliminee, volume_media_total } = dimensionneCC(flux)
          for (I = C.C1; I <= C.C2; I++) {
            const fin = F.predenit2_C1 + (I - C.C1)
            const elim = cuve_ratio_elimination[I] * DBO_eliminee
            resoudreVolumeFilling(I, (volume_media_total * 0.5) / Smedia(I), flux[fin][R.Q])
            flux[fin + 1][R.Q] = flux[fin][R.Q]
            flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
            flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elim
            const cn = nitrificationCuveC(I, (flux[fin + 1][R.DBO] / flux[fin + 1][R.Q]) * 1000, T_design)
            nit_cuve_C += cn
            flux[fin + 1][R.NK] = flux[fin][R.NK] - H.ratio_N_assimile * elim - cn
            flux[fin + 1][R.NH4] = flux[fin][R.NH4] - H.ratio_N_assimile * elim - cn
            flux[fin + 1][R.NO3] = flux[fin][R.NO3] + cn
            flux[fin + 1][R.Pt] = flux[fin][R.Pt] - H.ratio_P_assimile * elim
            flux[fin + 1][R.O2] = (cuve_O2_dissous[I] * flux[fin + 1][R.Q]) / 1000
            surface_media_totale += cuve_volume[I] * cuve_filling[I] * Smedia(I)
            boues_pdtes[I] = fct_boues_pdtes(elim, cn)
            const a = aerationCuve(I, flux, fin, T_design, elim, cn, s.Sh)
            nominal_O2_besoin_total += a.besoin
            nominal_air_Q += a.air
          }
        } else if (choix_cuve[C.C1]) {
          const fin = F.predenit2_C1
          const elim = flux[fin][R.DBO] - (hyp_cuve_C1_DBOsol_mgL * flux[fin][R.Q]) / 1000
          dimensionneCconfigCN(C.C1, H.hyp_rate_DBOin_appliquee_10_CN, flux, flux[fin][R.Q])
          nit_cuve_C = nitrificationCuveC(C.C1, hyp_cuve_C1_DBOsol_mgL, T_design)
          for (let k = F.C1_C2; k <= F.C2_N1; k++) {
            flux[k][R.Q] = flux[fin][R.Q]
            flux[k][R.DBO] = flux[fin][R.DBO] - elim
            flux[k][R.DCO] = flux[fin][R.DCO] - elim
            flux[k][R.NK] = flux[fin][R.NK] - H.ratio_N_assimile * elim - nit_cuve_C
            flux[k][R.NH4] = flux[fin][R.NH4] - H.ratio_N_assimile * elim - nit_cuve_C
            flux[k][R.NO3] = flux[fin][R.NO3] + nit_cuve_C
            flux[k][R.Pt] = flux[fin][R.Pt] - H.ratio_P_assimile * elim
            flux[k][R.O2] = (cuve_O2_dissous[k] * flux[fin + 1][R.Q]) / 1000
          }
          boues_pdtes[C.C1] = fct_boues_pdtes(elim, nit_cuve_C)
          surface_media_totale += cuve_volume[C.C1] * cuve_filling[C.C1] * Smedia(C.C1)
          const a = aerationCuve(C.C1, flux, fin, T_design, elim, nit_cuve_C, s.Sh)
          nominal_O2_besoin_total += a.besoin
          nominal_air_Q += a.air
        } else {
          for (let k = F.C1_C2; k <= F.C2_N1; k++) copierFlux(flux, k, F.predenit2_C1)
        }

        // ---- NITRIFICATION
        if (choix_cuve[C.N1]) {
          for (I = C.N1; I <= C.N2; I++) {
            const fin = F.C2_N1 + (I - C.N1)
            if (choix_cuve[I]) {
              const elim = flux[fin][R.DBO] - (hyp_cuve_N_deox_DBOsol_mgL * flux[fin][R.Q]) / 1000
              flux[fin][R.NH4] -= H.ratio_N_assimile * elim
              flux[fin][R.NK] -= H.ratio_N_assimile * elim
              let cuve_nit, NH4_removed, NH4_out_mgL
              if (choix_fct_Mox && I === C.N1 && f('ratio_elimination_N1_N12') == null) {
                NH4_removed = flux[fin][R.NH4] - (H.cuve_NH4_outlet_Mox_mgL / 1000) * flux[fin][R.Q]
                cuve_ratio_elimination[C.N1] = (1 / flux[F.C2_N1][R.NH4]) *
                  (NH4_removed + (1 + nominal_recirculation_MLSS_taux) * ((s.Q / 1000) * nominal_outlet_NH4_mgL))
                cuve_ratio_elimination[C.N2] = 1 - cuve_ratio_elimination[C.N1]
                cuve_nit = cuve_ratio_elimination[I] * (N_nit_total - deox_nit - nit_cuve_C)
                NH4_out_mgL = H.cuve_NH4_outlet_Mox_mgL
              } else {
                cuve_nit = cuve_ratio_elimination[I] * (N_nit_total - deox_nit - nit_cuve_C)
                NH4_removed = cuve_ratio_elimination[I] *
                  (flux[F.C2_N1][R.NH4] - (1 + nominal_recirculation_MLSS_taux) * ((s.Q / 1000) * nominal_outlet_NH4_mgL))
                NH4_out_mgL = ((flux[fin][R.NH4] - NH4_removed) * 1000) / flux[fin][R.Q]
              }
              dimensionneN(I, cuve_nit, NH4_out_mgL, flux[fin][R.Q])
              flux[fin + 1][R.Q] = flux[fin][R.Q]
              flux[fin + 1][R.NO3] = flux[fin][R.NO3] + (choix_fct_Mox && I === C.N1 ? cuve_nit * ratio_NO3f_nit : cuve_nit)
              flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
              flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elim
              flux[fin + 1][R.NK] = flux[fin][R.NK] - cuve_nit
              flux[fin + 1][R.NH4] = flux[fin][R.NH4] - NH4_removed
              flux[fin + 1][R.Pt] = flux[fin][R.Pt]
              flux[fin + 1][R.O2] = (cuve_O2_dissous[I] * flux[fin + 1][R.Q]) / 1000
              boues_pdtes[I] = fct_boues_pdtes(elim, cuve_nit)
              surface_media_totale += cuve_volume[I] * cuve_filling[I] * Smedia(I)
              const a = aerationCuve(I, flux, fin, T_design, elim, cuve_nit, s.Sh)
              nominal_O2_besoin_total += a.besoin
              nominal_air_Q += a.air
            } else copierFlux(flux, fin + 1, fin)
          }
        } else {
          for (I = F.N1_N2; I <= F.N2_deox; I++) copierFlux(flux, I, F.C2_N1)
        }

        // ---- DE-OXYGENATION : nitrification sur l'O2 résiduel de l'eau
        if (choix_cuve[C.deox]) {
          const O2_mgL = (flux[F.N2_deox][R.O2] / flux[F.N2_deox][R.Q]) * 1000
          deox_nit = ((O2_mgL - cuve_O2_dissous[C.deox]) * s.Q * (1 + nominal_recirculation_MLSS_taux)) / 1000 / H.ratio_O2_nit
          const fin = F.N2_deox
          const elim = flux[fin][R.DBO] - (hyp_cuve_N_deox_DBOsol_mgL * flux[fin][R.Q]) / 1000
          flux[fin][R.NH4] -= H.ratio_N_assimile * elim
          flux[fin][R.NK] -= H.ratio_N_assimile * elim
          dimensionneDeox(deox_nit, flux[fin][R.Q])
          flux[fin + 1][R.Q] = flux[fin][R.Q]
          flux[fin + 1][R.NO3] = flux[fin][R.NO3] + deox_nit
          flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
          flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elim
          flux[fin + 1][R.NK] = flux[fin][R.NK] - deox_nit
          flux[fin + 1][R.NH4] = flux[fin][R.NH4] - deox_nit
          flux[fin + 1][R.Pt] = flux[fin][R.Pt]
          // le VBA utilise ici l'indice I hérité de la boucle précédente
          flux[fin + 1][R.O2] = ((cuve_O2_dissous[I] ?? 0) * flux[fin + 1][R.Q]) / 1000
          boues_pdtes[C.deox] = fct_boues_pdtes(elim, deox_nit)
          surface_media_totale += cuve_volume[C.deox] * cuve_filling[C.deox] * Smedia(C.deox)
        } else copierFlux(flux, F.deox_separationMLSS, F.N2_deox)

        // ---- flux recirculé, prélevé en sortie de dé-ox
        for (let j = 1; j <= NB_PARAM; j++) {
          flux[F.recirculation_MLSS][j] = (flux[F.deox_separationMLSS][j] * nominal_recirculation_MLSS_taux) / (1 + nominal_recirculation_MLSS_taux)
        }
        epsDBO = Math.pow(epsDBO - flux[F.recirculation_MLSS][R.DBO], 2)
        epsNH4 = Math.pow(epsNH4 - flux[F.recirculation_MLSS][R.NH4], 2)
        epsNO3 = Math.pow(epsNO3 - flux[F.recirculation_MLSS][R.NO3], 2)
        eps = epsDBO + epsNH4 + epsNO3
      }
      if (garde >= 200) warnings.push('Dimensionnement : la boucle de recirculation de liqueur mixte n\'a pas convergé en 200 itérations.')

      for (let j = 1; j <= NB_PARAM; j++) {
        flux[F.separationMLSS_postdenit1][j] = flux[F.deox_separationMLSS][j] - flux[F.recirculation_MLSS][j]
      }
      const outlet_DBO_soluble_mgL = (flux[F.separationMLSS_postdenit1][R.DBO] / flux[F.separationMLSS_postdenit1][R.Q]) * 1000

      // ---- POST-DENITRIFICATION au méthanol
      if (choix_cuve[C.postdenit1]) {
        let denit_total = flux[F.separationMLSS_postdenit1][R.NO3] + H.ratio_NO3eq_O2 * flux[F.separationMLSS_postdenit1][R.O2]
        denit_total -= (nominal_outlet_NO3_mgL * flux[F.separationMLSS_postdenit1][R.Q]) / 1000
        const DBO_apportee = rate_DBO_apportee_denit_nominal * denit_total
        nominal_postdenit_carbone = f('postdenit_carbone_apporte_flux') ?? DBO_apportee / H.rate_DBO_source_C
        boues_postdenit = H.ratio_boues_DBO_apportee * DBO_apportee
        for (let i = C.postdenit1; i <= C.postdenit2; i++) {
          const fin = F.separationMLSS_postdenit1 + (i - C.postdenit1)
          if (choix_cuve[i]) {
            let NO3 = flux[fin][R.NO3] + H.ratio_NO3eq_O2 * flux[fin][R.O2]
            NO3 = ((NO3 - denit_total * cuve_ratio_elimination[i]) * 1000) / flux[F.separationMLSS_postdenit1][R.Q]
            dimensionnePostdenit(i, denit_total * cuve_ratio_elimination[i], NO3, flux[fin][R.Q])
            flux[fin + 1][R.Q] = flux[fin][R.Q]
            flux[fin + 1][R.DBO] = flux[fin][R.DBO]
            flux[fin + 1][R.DCO] = flux[fin][R.DCO]
            const Nassim = cuve_ratio_elimination[i] * boues_postdenit * H.ratio_N_assimile_Capporte
            if (flux[fin][R.NH4] < Nassim) {
              // NH4 insuffisant : le complément d'azote assimilé est pris sur les nitrates
              flux[fin + 1][R.NO3] = flux[fin][R.NO3] - (Nassim - flux[fin][R.NH4]) - cuve_ratio_elimination[i] * denit_total
              flux[fin + 1][R.NK] = flux[fin][R.NK] - flux[fin][R.NH4]
              flux[fin + 1][R.NH4] = 0
            } else {
              flux[fin + 1][R.NH4] = flux[fin][R.NH4] - Nassim
              flux[fin + 1][R.NO3] = flux[fin][R.NO3] - cuve_ratio_elimination[i] * denit_total
              flux[fin + 1][R.NK] = flux[fin][R.NK] - Nassim
            }
            flux[fin + 1][R.Pt] = flux[fin][R.Pt]
            flux[fin + 1][R.O2] = 0
            surface_media_totale += cuve_volume[i] * cuve_filling[i] * Smedia(i)
          } else copierFlux(flux, fin + 1, fin)
        }
        // fuite de carbone non consommé
        flux[F.postdenit2_reox][R.DBO] = flux[F.separationMLSS_postdenit1][R.DBO] + (H.DBO_apportee_fuite_mgL * flux[F.postdenit2_reox][R.Q]) / 1000
        flux[F.postdenit2_reox][R.DCO] = flux[F.separationMLSS_postdenit1][R.DCO] + (H.ratio_DCO_DBO_source_C * H.DBO_apportee_fuite_mgL * flux[F.postdenit2_reox][R.Q]) / 1000
      } else {
        for (let k = F.postdenit1_postdenit2; k <= F.postdenit2_reox; k++) copierFlux(flux, k, F.separationMLSS_postdenit1)
      }

      // ---- RE-OXYGENATION : reprise de la DBO apportée par le méthanol
      if (choix_cuve[C.reox]) {
        const fin = F.postdenit2_reox
        const elim = flux[fin][R.DBO] - (outlet_DBO_soluble_mgL * flux[fin][R.Q]) / 1000
        dimensionneReox(elim, flux[fin][R.Q])
        flux[fin + 1][R.Q] = flux[fin][R.Q]
        flux[fin + 1][R.NO3] = flux[fin][R.NO3]
        flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
        flux[fin + 1][R.DCO] = flux[fin][R.DCO]
        flux[fin + 1][R.NK] = flux[fin][R.NK]
        flux[fin + 1][R.NH4] = flux[fin][R.NH4]
        flux[fin + 1][R.Pt] = flux[fin][R.Pt]
        flux[fin + 1][R.O2] = (cuve_O2_dissous[C.reox] * flux[fin + 1][R.Q]) / 1000
        boues_pdtes[C.reox] = fct_boues_pdtes(elim, 0)
        surface_media_totale += cuve_volume[C.reox] * cuve_filling[C.reox] * Smedia(C.reox)
        const a = aerationCuve(C.reox, flux, fin, T_design, elim, 0, s.Sh)
        nominal_O2_besoin_total += a.besoin
        nominal_air_Q += a.air
      } else copierFlux(flux, F.outlet, F.postdenit2_reox)

      volume_total_bassins = f('volume_total_bassins') ?? cuve_volume.reduce((a, b) => a + b, 0)
      let boues_aerobie = boues_pdtes.reduce((a, b) => a + b, 0)
      nominal_boues_MV_MES = nominal_boues_MV_MES * (1 - rate_DBOpart_hydrolyse) + boues_aerobie + boues_postdenit
      boues_aerobie += flux[F.inlet][R.MES] * (1 - rate_DBOpart_hydrolyse)
      const sep = separationBoues(flux, nominal_outlet_DBO_mgL, (flux[F.outlet][R.DBO] / flux[F.outlet][R.Q]) * 1000, boues_aerobie, nominal_boues_MV_MES, false)
      nominal_sortie_MES = sep.sortie_MES
      nominal_boues_MV_MES = f('boues_MV_MES') ?? sep.MV_MES
      boues_Q = sep.boues_Q

      const out = makeStream()
      out.Q = flux[F.outlet][R.Q]
      out.DBO = flux[F.outlet][R.DBO]
      const DCO_dure_mgL = site.Q_nominal > 0 ? ((site.DCO_nominal * H.ratio_DCO_dure_total) / site.Q_nominal) * 1000 : 0
      out.DCO = out.DBO + ((DCO_dure_mgL + 1.45 * nominal_sortie_MES * nominal_boues_MV_MES) * out.Q) / 1000
      out.MES = (nominal_sortie_MES * out.Q) / 1000
      out.NK = flux[F.outlet][R.NK]
      out.NH4 = flux[F.outlet][R.NH4]
      out.NO3 = flux[F.outlet][R.NO3]
      out.Pt = flux[F.outlet][R.Pt]
      out.Sh = 0
      return out
    }

    // =======================================================================
    // fonctionnement_reel (eau réelle, T d'exploitation)
    // =======================================================================
    const reel_flux = nouveauFlux()
    let reel_recirculation_MLSS_taux = 0
    let reel_O2_besoin_total = 0
    let reel_air_Q = 0
    let reel_sortie_MES = 0
    let reel_boues_MES = 0
    let reel_boues_MV_MES = 0
    let reel_boues_Q = 0
    let reel_outlet_DBO_mgL = 0
    let reel_outlet_NH4_mgL = 0
    let reel_outlet_NO3_mgL = 0
    let reel_postdenit_carbone = 0
    let suraeration_brassage_Nm3j = 0
    let debit_air_agitation = 0

    /** simule_C : DBO soluble de sortie déduite de la charge surfacique appliquée */
    function simuleC(cuve, flux, fin) {
      const S = cuve === C.C1
        ? cuve_volume[C.C1] * cuve_filling[C.C1] * Smedia(C.C1)
        : cuve_volume[C.C1] * cuve_filling[C.C1] * Smedia(C.C1) + cuve_volume[C.C2] * cuve_filling[C.C2] * Smedia(C.C2)
      if (!(S > 0)) return 0
      let rateT = flux[F.inlet][R.DBO] - (flux[F.inlet_recirculation][R.DBO] - flux[F.predenit2_C1][R.DBO])
      rateT = (1000 * rateT) / S
      const rate10 = T_reel <= 10 ? rateT / Math.pow(1.07, T_reel - 10)
        : T_reel <= 25 ? rateT / Math.pow(1.06, T_reel - 10)
          : rateT / Math.pow(1.06, 15)
      const sep = choices.type_separation
      let DBOsol = sep === SEP.clarif ? (rate10 + 0.273) / 0.8346
        : sep === SEP.clarif_polymere ? (rate10 - 0.056) / 0.7688
          : rate10 <= 3.9 ? (5 / 3.9) * rate10 : (rate10 + 2.5705) / 1.281
      if (DBOsol < 0) DBOsol = 0
      return flux[fin][R.DBO] - (DBOsol * flux[fin][R.Q]) / 1000
    }

    /** simule_N : nitrification, bascule sur Newton-Raphson si le NH4 devient limitant */
    function simuleN(i, flux, fin) {
      const S = cuve_volume[i] * cuve_filling[i] * Smedia(i)
      const Qin = flux[fin][R.Q]
      const NH4in = flux[fin][R.NH4]
      let NH4out
      if (choix_fct_Mox && i === C.N1) {
        const rate = f('rate_N_T_reel') ?? (H.rate_N_30[choix_media[i]] ?? 0) * Math.pow(1.07, T_reel - 30)
        NH4out = ((NH4in - (S / 1000) * rate) / Qin) * 1000
      } else {
        const corr = corrT_nit(T_reel)
        const seuil = (cuve_O2_dissous[i] - 0.5) / 3.2
        let nit = seuil > 0 ? (S / 1000) * k_nit * Math.pow(seuil, 0.7) * corr : 0
        NH4out = ((NH4in - nit) / Qin) * 1000
        if (NH4out < seuil) {
          nit = NH4in - (nominal_outlet_NH4_mgL * Qin) / 1000
          let eps = 1
          let garde = 0
          while (Math.abs(eps) > H.critere_convergence && garde++ < 200) {
            const conc = ((NH4in - nit) / Qin) * 1000
            if (!(conc > 0)) break
            const fn = nit - (S / 1000) * k_nit * Math.pow(conc, 0.7) * corr
            const fp = 1 + ((S * corr) / 1000) * k_nit * Math.pow(1000 / Qin, 0.7) * Math.pow(NH4in - nit, -0.3)
            if (!Number.isFinite(fp) || fp === 0) break
            nit -= fn / fp
            NH4out = ((NH4in - nit) / Qin) * 1000
            eps = fn
          }
        }
      }
      return NH4out < 0 ? 0 : NH4out
    }

    /** simule_predenit : Newton-Raphson sur la concentration de NO3 en sortie */
    function simulePredenit(i, flux, fin, DBOsol_EB) {
      const corr = corrT_DBO(T_reel)
      const S = cuve_volume[i] * cuve_filling[i] * Smedia(i)
      const Qin = flux[fin][R.Q]
      const NO3in = ((flux[fin][R.NO3] + H.ratio_NO3eq_O2 * flux[fin][R.O2]) / Qin) * 1000
      const L = Math.log(0.3)
      const kQ = Qin / 1000
      let NO3out = 1
      let eps = 1
      let garde = 0
      while (Math.abs(eps) > H.critere_convergence && garde++ < 200) {
        const corrNO3 = NO3out / (NO3out + 0.4)
        const denit = (NO3in - NO3out) * kQ
        const u = (-H.ratio_O2_denit * denit) / DBOsol_EB
        const E = Math.exp(L * u * u)
        const fn = denit - (corrNO3 * E * corr * S) / 1000
        const dcorr = 0.4 / Math.pow(NO3out + 0.4, 2)
        const dE = L * Math.pow((-H.ratio_O2_denit * kQ) / DBOsol_EB, 2) * -2 * (NO3in - NO3out) * E
        const fp = -kQ - (dcorr * E * corr * S) / 1000 - ((corrNO3 * corr * S) / 1000) * dE
        if (!Number.isFinite(fp) || fp === 0) break
        NO3out -= fn / fp
        eps = fn
      }
      return NO3out >= 0 && Number.isFinite(NO3out) ? NO3out : 0
    }

    /** simule_deox : Newton-Raphson sur l'O2 dissous résiduel */
    function simuleDeox(flux) {
      const corr = corrT_nit(T_reel)
      const S = cuve_volume[C.deox] * cuve_filling[C.deox] * Smedia(C.deox)
      const Q = flux[F.N2_deox][R.Q]
      let inletO2 = flux[F.N2_deox][R.O2] - respiration(T_reel, S)
      inletO2 = (inletO2 / Q) * 1000
      let O2 = 1
      let eps = 1
      let garde = 0
      while (Math.abs(eps) > H.critere_convergence && garde++ < 200) {
        if (O2 <= 0.5) { O2 = 0.5; break }
        const fn = ((inletO2 - O2) * Q) / 1000 / H.ratio_O2_nit - (S / 1000) * k_nit * Math.pow((O2 - 0.5) / 3.2, 0.7) * corr
        const fp = -(Q / 1000) / H.ratio_O2_nit - (((S / 1000) * k_nit * corr) / Math.pow(3.2, 0.7)) * 0.7 * Math.pow(O2 - 0.5, -0.3)
        if (!Number.isFinite(fp) || fp === 0) break
        O2 -= fn / fp
        eps = fn
      }
      return Number.isFinite(O2) ? O2 : 0
    }

    /** simule_postdenit : Newton-Raphson, vitesse Monod 2,29·NO3/(NO3+2,9) */
    function simulePostdenit(i, flux, fin) {
      const corr = corrT_DBO(T_reel)
      const S = cuve_volume[i] * cuve_filling[i] * Smedia(i)
      const Q = flux[fin][R.Q]
      const NO3in = flux[fin][R.NO3] + H.ratio_NO3eq_O2 * flux[fin][R.O2]
      let NO3 = nominal_outlet_NO3_mgL
      let eps = 1
      let garde = 0
      while (Math.abs(eps) > H.critere_convergence && garde++ < 200) {
        const fn = NO3in - (NO3 * Q) / 1000 - ((S * corr) / 1000) * ((2.29 * NO3) / (NO3 + 2.9))
        const fp = -Q / 1000 - ((S * corr) / 1000) * ((2.29 * 2.9) / Math.pow(NO3 + 2.9, 2))
        if (!Number.isFinite(fp) || fp === 0) break
        NO3 -= fn / fp
        eps = fn
      }
      return NO3 >= 0 && Number.isFinite(NO3) ? NO3 : 0
    }

    /** simule_reox : élimination de DBO à vitesse fixe de 4 g/(m²·j) à 10 °C */
    function simuleReox() {
      const S = cuve_volume[C.reox] * cuve_filling[C.reox] * Smedia(C.reox)
      return (S / 1000) * 4 * corrT_DBO(T_reel)
    }

    function fonctionnementReel() {
      const s = ctx.inReel
      const flux = reel_flux
      let boues_MV_MES = MV_MES_influent(s.MES, s.DCO, s.DBO) * s.MES
      flux[F.inlet][R.Q] = s.Q; flux[F.inlet][R.DCO] = s.DCO; flux[F.inlet][R.DBO] = s.DBO
      flux[F.inlet][R.MES] = s.MES; flux[F.inlet][R.NK] = s.NK
      // au réel, l'ammonification du N organique est intégrée dès l'entrée
      flux[F.inlet][R.NH4] = s.NH4 + EB_rate_Norgasol_Norga * (s.NK - s.NH4)
        + (1 - EB_rate_Norgasol_Norga) * rate_Norgapart_hydrolyse * (s.NK - s.NH4)
        - (Norga_dure_soluble * s.Q) / 1000
      flux[F.inlet][R.NO3] = s.NO3; flux[F.inlet][R.Pt] = s.Pt; flux[F.inlet][R.O2] = 0

      reel_recirculation_MLSS_taux = (choix_cuve[C.predenit1] && f('reel_recirculation_MLSS_taux') != null)
        ? f('reel_recirculation_MLSS_taux')
        : nominal_recirculation_MLSS_taux

      if (choix_cuve[C.predenit1]) {
        const Qr = reel_recirculation_MLSS_taux * s.Q
        flux[F.recirculation_MLSS][R.Q] = Qr
        flux[F.recirculation_MLSS][R.DBO] = (nominal_outlet_DBO_mgL * Qr) / 1000
        flux[F.recirculation_MLSS][R.DCO] = flux[F.recirculation_MLSS][R.DBO]
        flux[F.recirculation_MLSS][R.NK] = ((nominal_outlet_NH4_mgL + Norga_dure_soluble) * Qr) / 1000
        flux[F.recirculation_MLSS][R.NH4] = (nominal_outlet_NH4_mgL * Qr) / 1000
        flux[F.recirculation_MLSS][R.NO3] = (nominal_last_aerated_tank_NO3_mgL * Qr) / 1000
        flux[F.recirculation_MLSS][R.Pt] = s.Q > 0 ? (s.Pt / s.Q) * Qr : 0
        flux[F.recirculation_MLSS][R.O2] = 0
      }

      const boues_pdtes = new Array(NB_CUVES + 1).fill(0)
      let I = 0
      let eps = 1 + H.critere_convergence
      let garde = 0
      while (eps > H.critere_convergence && garde++ < 200) {
        let epsDBO = flux[F.recirculation_MLSS][R.DBO]
        let epsNH4 = flux[F.recirculation_MLSS][R.NH4]
        let epsNO3 = flux[F.recirculation_MLSS][R.NO3]
        reel_O2_besoin_total = 0
        reel_air_Q = 0
        for (let j = 1; j <= NB_PARAM; j++) flux[F.inlet_recirculation][j] = flux[F.inlet][j] + flux[F.recirculation_MLSS][j]

        // ---- PRE-DENIT
        let elimPre = 0
        for (I = C.predenit1; I <= C.predenit2; I++) {
          const fin = F.inlet_recirculation + (I - C.predenit1)
          if (choix_cuve[I]) {
            const DBOsol_EB = flux[F.inlet][R.DBO] * (EB_rate_DBOsol_DBO + (1 - EB_rate_DBOsol_DBO) * rate_DBOpart_hydrolyse)
              + flux[F.recirculation_MLSS][R.DBO] - elimPre
            const NO3out = simulePredenit(I, flux, fin, DBOsol_EB)
            const denit = flux[fin][R.NO3] + H.ratio_NO3eq_O2 * flux[fin][R.O2] - (NO3out * flux[fin][R.Q]) / 1000
            elimPre = -H.ratio_O2_denit * denit
            flux[fin + 1][R.Q] = flux[fin][R.Q]
            flux[fin + 1][R.NO3] = flux[fin][R.NO3] + H.ratio_NO3eq_O2 * flux[fin][R.O2] - denit
            flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elimPre
            flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elimPre
            flux[fin + 1][R.NK] = flux[fin][R.NK] - H.ratio_N_assimile * elimPre
            flux[fin + 1][R.NH4] = flux[fin][R.NH4] - H.ratio_N_assimile * elimPre
            flux[fin + 1][R.Pt] = flux[fin][R.Pt]
            flux[fin + 1][R.O2] = 0
          } else copierFlux(flux, fin + 1, fin)
        }

        // ---- CARBONE
        if (choix_cuve[C.C1] && nb_etapes_traitement === 1) {
          // configuration C seul : la DBO éliminée est calculée globalement sur
          // les deux cuves puis répartie selon cuve_ratio_elimination
          const DBO_removed = simuleC(C.C2, flux, F.predenit2_C1)
          for (let k = 0; k < 2; k++) {
            const i = C.C1 + k
            const fin = F.predenit2_C1 + k
            const part = cuve_ratio_elimination[i] * DBO_removed
            flux[fin + 1][R.Q] = flux[fin][R.Q]
            flux[fin + 1][R.DCO] = flux[fin][R.DCO] - part
            flux[fin + 1][R.DBO] = flux[fin][R.DBO] - part
            const cn = nitrificationCuveC(i, (flux[fin + 1][R.DBO] / flux[fin + 1][R.Q]) * 1000, T_reel)
            flux[fin + 1][R.NK] = flux[fin][R.NK] - H.ratio_N_assimile * part - cn
            flux[fin + 1][R.NH4] = flux[fin][R.NH4] - H.ratio_N_assimile * part - cn
            flux[fin + 1][R.NO3] = flux[fin][R.NO3] + cn
            flux[fin + 1][R.Pt] = flux[fin][R.Pt] - H.ratio_P_assimile * part
            flux[fin + 1][R.O2] = (cuve_O2_dissous[i] * flux[fin + 1][R.Q]) / 1000
            boues_pdtes[i] = fct_boues_pdtes(DBO_removed, cn)
            const a = aerationCuve(i, flux, fin, T_reel, part, cn, s.Sh)
            reel_O2_besoin_total += a.besoin
            reel_air_Q += a.air
          }
        } else if (choix_cuve[C.C1]) {
          const i = C.C1
          const fin = F.predenit2_C1
          const DBO_removed = simuleC(i, flux, fin)
          flux[fin + 1][R.Q] = flux[fin][R.Q]
          flux[fin + 1][R.DCO] = flux[fin][R.DCO] - DBO_removed
          flux[fin + 1][R.DBO] = flux[fin][R.DBO] - DBO_removed
          let cn = nitrificationCuveC(i, (flux[fin + 1][R.DBO] / flux[fin + 1][R.Q]) * 1000, T_reel)
          const dispo = flux[fin][R.NH4] - H.ratio_N_assimile * DBO_removed
          if (cn > dispo) cn = dispo
          flux[fin + 1][R.NK] = flux[fin][R.NK] - H.ratio_N_assimile * DBO_removed - cn
          flux[fin + 1][R.NH4] = flux[fin][R.NH4] - H.ratio_N_assimile * DBO_removed - cn
          flux[fin + 1][R.NO3] = flux[fin][R.NO3] + cn
          flux[fin + 1][R.Pt] = flux[fin][R.Pt] - H.ratio_P_assimile * DBO_removed
          flux[fin + 1][R.O2] = (cuve_O2_dissous[i] * flux[fin + 1][R.Q]) / 1000
          boues_pdtes[i] = fct_boues_pdtes(DBO_removed, cn)
          const a = aerationCuve(i, flux, fin, T_reel, DBO_removed, cn, s.Sh)
          reel_O2_besoin_total += a.besoin
          reel_air_Q += a.air
          copierFlux(flux, fin + 2, fin + 1)
        } else {
          copierFlux(flux, F.predenit2_C1 + 1, F.predenit2_C1)
          copierFlux(flux, F.predenit2_C1 + 2, F.predenit2_C1)
        }

        // ---- NITRIFICATION
        for (I = C.N1; I <= C.N2; I++) {
          const fin = F.C2_N1 + (I - C.N1)
          if (choix_cuve[I]) {
            const elim = flux[fin][R.DBO] - (hyp_cuve_N_deox_DBOsol_mgL * flux[fin][R.Q]) / 1000
            flux[fin][R.NH4] -= H.ratio_N_assimile * elim
            flux[fin][R.NK] -= H.ratio_N_assimile * elim
            const NH4out = simuleN(I, flux, fin)
            const cuve_nit = flux[fin][R.NH4] - (NH4out * flux[fin][R.Q]) / 1000
            flux[fin + 1][R.Q] = flux[fin][R.Q]
            flux[fin + 1][R.NO3] = flux[fin][R.NO3] + (choix_fct_Mox && I === C.N1 ? cuve_nit * ratio_NO3f_nit : cuve_nit)
            flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elim
            flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
            flux[fin + 1][R.NK] = flux[fin][R.NK] - cuve_nit
            flux[fin + 1][R.NH4] = flux[fin][R.NH4] - cuve_nit
            flux[fin + 1][R.Pt] = flux[fin][R.Pt]
            flux[fin + 1][R.O2] = (cuve_O2_dissous[I] * flux[fin + 1][R.Q]) / 1000
            boues_pdtes[I] = fct_boues_pdtes(elim, cuve_nit)
            const a = aerationCuve(I, flux, fin, T_reel, elim, cuve_nit, s.Sh)
            reel_O2_besoin_total += a.besoin
            reel_air_Q += a.air
          } else copierFlux(flux, fin + 1, fin)
        }

        // ---- DE-OX
        if (choix_cuve[C.deox]) {
          const elim = flux[F.N2_deox][R.DBO] - (hyp_cuve_N_deox_DBOsol_mgL * flux[F.N2_deox][R.Q]) / 1000
          flux[F.N2_deox][R.NH4] -= H.ratio_N_assimile * elim
          flux[F.N2_deox][R.NK] -= H.ratio_N_assimile * elim
          const O2out = simuleDeox(flux)
          let cuve_nit = (flux[F.N2_deox][R.O2] - (O2out * flux[F.N2_deox][R.Q]) / 1000) / H.ratio_O2_nit
          if (cuve_nit > flux[F.N2_deox][R.NH4]) cuve_nit = flux[F.N2_deox][R.NH4]
          const d = F.deox_separationMLSS
          flux[d][R.Q] = flux[F.N2_deox][R.Q]
          flux[d][R.NO3] = flux[F.N2_deox][R.NO3] + cuve_nit
          flux[d][R.DCO] = flux[F.N2_deox][R.DCO] - elim
          flux[d][R.DBO] = flux[F.N2_deox][R.DBO] - elim
          flux[d][R.NK] = flux[F.N2_deox][R.NK] - cuve_nit
          flux[d][R.NH4] = flux[F.N2_deox][R.NH4] - cuve_nit
          flux[d][R.Pt] = flux[F.N2_deox][R.Pt]
          flux[d][R.O2] = (O2out * flux[d][R.Q]) / 1000
          // le VBA emploie ici une variable locale deox_nit restée nulle
          boues_pdtes[C.deox] = fct_boues_pdtes(elim, 0)
        } else copierFlux(flux, F.deox_separationMLSS, F.N2_deox)

        for (let j = 1; j <= NB_PARAM; j++) {
          flux[F.recirculation_MLSS][j] = (flux[F.deox_separationMLSS][j] * reel_recirculation_MLSS_taux) / (1 + reel_recirculation_MLSS_taux)
        }
        epsDBO = Math.pow(epsDBO - flux[F.recirculation_MLSS][R.DBO], 2)
        epsNH4 = Math.pow(epsNH4 - flux[F.recirculation_MLSS][R.NH4], 2)
        epsNO3 = Math.pow(epsNO3 - flux[F.recirculation_MLSS][R.NO3], 2)
        eps = epsDBO + epsNH4 + epsNO3
      }
      if (garde >= 200) warnings.push("Fonctionnement réel : la boucle de recirculation de liqueur mixte n'a pas convergé en 200 itérations.")

      for (let j = 1; j <= NB_PARAM; j++) {
        flux[F.separationMLSS_postdenit1][j] = flux[F.deox_separationMLSS][j] - flux[F.recirculation_MLSS][j]
      }

      // ---- POST-DENIT
      let DBO_apportee = 0
      for (let i = C.postdenit1; i <= C.postdenit2; i++) {
        const fin = F.separationMLSS_postdenit1 + (i - C.postdenit1)
        if (choix_cuve[i]) {
          const NO3out = simulePostdenit(i, flux, fin)
          const denit = flux[fin][R.NO3] - (NO3out * flux[fin][R.Q]) / 1000
          DBO_apportee += rate_DBO_apportee_denit_reel * denit
          flux[fin + 1][R.Q] = flux[fin][R.Q]
          flux[fin + 1][R.NO3] = flux[fin][R.NO3] - denit
          if (i === C.postdenit2 || !choix_cuve[C.postdenit2]) {
            flux[fin + 1][R.DCO] = flux[fin][R.DCO] + (H.ratio_DCO_DBO_source_C * H.DBO_apportee_fuite_mgL * flux[fin + 1][R.Q]) / 1000
            flux[fin + 1][R.DBO] = flux[fin][R.DBO] + (H.DBO_apportee_fuite_mgL * flux[fin + 1][R.Q]) / 1000
          } else {
            flux[fin + 1][R.DCO] = flux[fin][R.DCO]
            flux[fin + 1][R.DBO] = flux[fin][R.DBO]
          }
          const Nassim = H.ratio_N_assimile_Capporte * H.ratio_boues_DBO_apportee * rate_DBO_apportee_denit_reel * denit
          flux[fin + 1][R.NK] = Math.max(0, flux[fin][R.NK] - Nassim)
          flux[fin + 1][R.NH4] = Math.max(0, flux[fin][R.NH4] - Nassim)
          flux[fin + 1][R.Pt] = flux[fin][R.Pt]
          flux[fin + 1][R.O2] = 0
        } else copierFlux(flux, fin + 1, fin)
      }
      reel_postdenit_carbone = f('postdenit_carbone_apporte_flux') ?? DBO_apportee / H.rate_DBO_source_C
      boues_postdenit = H.ratio_boues_DBO_apportee * DBO_apportee

      // ---- RE-OX
      if (choix_cuve[C.reox]) {
        const fin = F.postdenit2_reox
        let elim = simuleReox()
        if (elim > flux[fin][R.DBO]) elim = flux[fin][R.DBO]
        flux[fin + 1][R.Q] = flux[fin][R.Q]
        flux[fin + 1][R.DCO] = flux[fin][R.DCO] - elim
        flux[fin + 1][R.DBO] = flux[fin][R.DBO] - elim
        flux[fin + 1][R.NK] = flux[fin][R.NK]
        flux[fin + 1][R.NH4] = flux[fin][R.NH4]
        flux[fin + 1][R.NO3] = flux[fin][R.NO3]
        flux[fin + 1][R.Pt] = flux[fin][R.Pt] - H.ratio_P_assimile * elim
        flux[fin + 1][R.O2] = (cuve_O2_dissous[C.reox] * flux[fin + 1][R.Q]) / 1000
        boues_pdtes[C.reox] = fct_boues_pdtes(elim, 0)
        const a = aerationCuve(C.reox, flux, fin, T_reel, elim, 0, s.Sh)
        reel_O2_besoin_total += a.besoin
        reel_air_Q += a.air
      } else copierFlux(flux, F.outlet, F.postdenit2_reox)

      // DBO totale de sortie, reconstruite à partir de la DBO soluble
      const DBOsol_out = (flux[F.outlet][R.DBO] / flux[F.outlet][R.Q]) * 1000
      if (nb_etapes_traitement === 1) {
        const sep = choices.type_separation
        reel_outlet_DBO_mgL = sep === SEP.clarif ? (DBOsol_out + 5) / 0.5
          : sep === SEP.clarif_polymere ? (DBOsol_out + 3) / 0.5275
            : (DBOsol_out + 1.1553) / 0.3936
      } else {
        const force = f('reel_outlet_DBO_mgL')
        if (force != null) reel_outlet_DBO_mgL = force
        else {
          const nomSol = (nominal_flux[F.outlet][R.DBO] / nominal_flux[F.outlet][R.Q]) * 1000
          reel_outlet_DBO_mgL = nomSol > 0 ? nominal_outlet_DBO_mgL * (DBOsol_out / nomSol) : H.hyp_outlet_DBO_mgL_mini
          if (reel_outlet_DBO_mgL <= H.hyp_outlet_DBO_mgL_mini) reel_outlet_DBO_mgL = H.hyp_outlet_DBO_mgL_mini
        }
      }

      let boues_aerobie = boues_pdtes.reduce((a, b) => a + b, 0)
      boues_MV_MES = boues_MV_MES * (1 - rate_DBOpart_hydrolyse) + boues_aerobie + boues_postdenit
      boues_aerobie += flux[F.inlet][R.MES] * (1 - rate_DBOpart_hydrolyse)
      const sep = separationBoues(flux, reel_outlet_DBO_mgL, DBOsol_out, boues_aerobie, boues_MV_MES, true)
      reel_sortie_MES = sep.sortie_MES
      reel_boues_MES = sep.boues_MES
      reel_boues_MV_MES = f('boues_MV_MES') ?? sep.MV_MES
      reel_boues_Q = sep.boues_Q

      const Q = flux[F.outlet][R.Q]
      reel_outlet_NH4_mgL = (flux[F.outlet][R.NH4] / Q) * 1000
      reel_outlet_NO3_mgL = (flux[F.outlet][R.NO3] / Q) * 1000
      let NK = flux[F.outlet][R.NK]

      // Forçages sur les sorties : besoin en O2 et débit d'air réajustés au
      // prorata (reproduction fidèle de l'arbre du VBA).
      const NH4f = f('reel_outlet_NH4_mgL')
      const O2f = f('O2_besoin_total')
      const airf = f('air_Q_Nm3j')
      if (NH4f != null) {
        const delta = (H.ratio_O2_nit * (reel_outlet_NH4_mgL - NH4f) * Q) / 1000
        if (O2f != null) {
          if (airf != null) { reel_O2_besoin_total = O2f; reel_air_Q = airf }
          else { reel_air_Q = reel_O2_besoin_total > 0 ? (O2f / reel_O2_besoin_total) * reel_air_Q : 0; reel_O2_besoin_total = O2f }
        } else if (airf != null) { reel_air_Q = airf; reel_O2_besoin_total += delta }
        else {
          reel_air_Q = reel_O2_besoin_total > 0 ? ((reel_O2_besoin_total + delta) / reel_O2_besoin_total) * reel_air_Q : 0
          reel_O2_besoin_total += delta
        }
        reel_outlet_NH4_mgL = NH4f
      } else if (O2f != null) {
        reel_air_Q = airf != null ? airf : (reel_O2_besoin_total > 0 ? (O2f / reel_O2_besoin_total) * reel_air_Q : 0)
        reel_O2_besoin_total = O2f
      } else if (airf != null) reel_air_Q = airf
      const NO3f = f('reel_outlet_NO3_mgL')
      if (NO3f != null) reel_outlet_NO3_mgL = NO3f

      // recomposition cohérente des charges de sortie
      NK -= flux[F.outlet][R.NH4]
      const out = makeStream()
      out.Q = Q
      out.DBO = (reel_outlet_DBO_mgL * Q) / 1000
      const Qn = site.Q_nominal * site.NC_Q
      const DCO_dure_mgL = Qn > 0 ? ((site.DCO_nominal * site.NC_DCO * H.ratio_DCO_dure_total) / Qn) * 1000 : 0
      out.DCO = out.DBO + ((DCO_dure_mgL + 1.45 * reel_sortie_MES * reel_boues_MV_MES) * Q) / 1000
      out.MES = (reel_sortie_MES * Q) / 1000
      out.NH4 = (reel_outlet_NH4_mgL * Q) / 1000
      out.NK = NK + out.NH4
      out.NO3 = (reel_outlet_NO3_mgL * Q) / 1000
      out.Pt = flux[F.outlet][R.Pt]
      out.Sh = 0
      suraeration_brassage_Nm3j = cuve_suraeration_brassage.reduce((a, b) => a + b, 0)
      debit_air_agitation = cuve_debit_air_agitation.reduce((a, b) => a + b, 0)
      return out
    }

    const outNominal = dimensionnement()
    const outReel = fonctionnementReel()

    // =======================================================================
    // calcul_consommation_electrique (sur le fonctionnement réel)
    // =======================================================================
    if (choices.surpresseur === 'roots' && air_P_refoulement_moyenne > HYP.surpresseur_Px2) {
      warnings.push(`Pression de refoulement (${air_P_refoulement_moyenne.toFixed(1)} mCE) élevée pour des surpresseurs roots.`)
    }
    const surpresseur_conso_spec = f('surpresseur_conso_spec') ?? HYP.surpresseur_conso_spec_Wh_Nm3mCE[choices.surpresseur]
    const electricite_aeration = (reel_air_Q * air_P_refoulement_moyenne * surpresseur_conso_spec) / 1000
    const electricite_recirculation_MLSS = ratio_elec_recirculation_MLSS * reel_flux[F.recirculation_MLSS][R.Q] * recirculation_MLSS_P_refoulement
    const electricite_extraction = ratio_elec_extraction * reel_boues_Q * extraction_P_refoulement
    let electricite_agitation = 0
    for (let i = 1; i <= NB_CUVES; i++) {
      if (!choix_cuve[i]) continue
      const mecanique = !EST_AEREE(i) || choices.agitation_cuve_aeree === 'mecanique'
      if (mecanique) electricite_agitation += (cuve_agitation_W_m3[i] * cuve_volume[i] * H.agitation_fct) / 1000
    }
    const total = electricite_aeration + electricite_agitation + electricite_recirculation_MLSS + electricite_extraction
    const fixed = electricite_agitation + (reel_air_Q > 0 ? (debit_air_agitation / reel_air_Q) * electricite_aeration : 0)

    // =======================================================================
    const resultsCuves = []
    for (let i = 1; i <= NB_CUVES; i++) {
      if (!choix_cuve[i]) continue
      resultsCuves.push({ key: `V${i}`, label: `Volume ${NOM_CUVE[i]}`, unit: 'm³', value: cuve_volume[i] })
      resultsCuves.push({ key: `fill${i}`, label: `Taux de remplissage ${NOM_CUVE[i]}`, unit: '-', value: cuve_filling[i] })
    }

    return {
      outNominal,
      outReel,
      sludge: {
        origine: 'II_MBBR',
        Q: reel_boues_Q,
        MES: reel_boues_MES,
        concentration: boues_concentration,
        MV_MES: reel_boues_MV_MES,
        NK: ratio('II_MBBR', 'NK_MV') * reel_boues_MES * reel_boues_MV_MES,
        Pt: ratio('II_MBBR', 'Pt_MES') * reel_boues_MES,
        DCO: ratio('II_MBBR', 'DCO_MV') * reel_boues_MES * reel_boues_MV_MES,
        DBO: ratio('II_MBBR', 'DBO_MV') * reel_boues_MES * reel_boues_MV_MES,
      },
      results: [
        { key: 'V_total', label: 'Volume total des cuves', unit: 'm³', value: volume_total_bassins },
        { key: 'S_media', label: 'Surface de média développée', unit: 'm²', value: surface_media_totale },
        ...resultsCuves,
        { key: 'rec_nom', label: 'Recirculation de liqueur mixte (nominal)', unit: '-', value: nominal_recirculation_MLSS_taux },
        { key: 'rec_reel', label: 'Recirculation de liqueur mixte (réel)', unit: '-', value: reel_recirculation_MLSS_taux },
        { key: 'O2_nom', label: 'Besoin en O2 (nominal)', unit: 'kg O2/j', value: nominal_O2_besoin_total },
        { key: 'O2_reel', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: reel_O2_besoin_total },
        { key: 'air_nom', label: "Débit d'air process (nominal)", unit: 'Nm³/h', value: nominal_air_Q / 24 },
        { key: 'air', label: "Débit d'air process (réel)", unit: 'Nm³/h', value: reel_air_Q / 24 },
        { key: 'air_brassage', label: 'dont suraération de brassage (réel)', unit: 'Nm³/h', value: suraeration_brassage_Nm3j / 24 },
        { key: 'P_air', label: 'Pression de refoulement retenue', unit: 'mCE', value: air_P_refoulement_moyenne },
        ...(choix_cuve[C.postdenit1] ? [
          { key: 'MeOH_nom', label: 'Méthanol (nominal)', unit: 'kg/j', value: nominal_postdenit_carbone },
          { key: 'MeOH', label: 'Méthanol (réel)', unit: 'kg/j', value: reel_postdenit_carbone },
        ] : []),
        { key: 'DBO_out', label: 'DBO5 en sortie (réel)', unit: 'mg/L', value: reel_outlet_DBO_mgL },
        { key: 'NH4_out', label: 'N-NH4 en sortie (réel)', unit: 'mg/L', value: reel_outlet_NH4_mgL },
        { key: 'NO3_out', label: 'N-NO3 en sortie (réel)', unit: 'mg/L', value: reel_outlet_NO3_mgL },
        { key: 'MES_out', label: 'MES en sortie (réel)', unit: 'mg/L', value: reel_sortie_MES },
        { key: 'boues', label: 'Boues à extraire (réel)', unit: 'kg MES/j', value: reel_boues_MES },
        { key: 'boues_MV', label: 'MV/MES des boues (réel)', unit: '-', value: reel_boues_MV_MES },
        { key: 'boues_Q', label: 'Débit de boues extraites (réel)', unit: 'm³/j', value: reel_boues_Q },
      ],
      electricity: {
        total,
        fixed,
        detail: {
          aeration: electricite_aeration,
          agitation: electricite_agitation,
          recirculation_MLSS: electricite_recirculation_MLSS,
          extraction: electricite_extraction,
        },
      },
      warnings,
    }
  },
})
