// ---------------------------------------------------------------------------
// Port de E5_HybAS.cls — procédé hybride (IFAS) : boues activées en suspension
// + biofilm fixé sur média mobile dans une partie des cuves aérées.
//
// Chaîne de traitement (chaque cuve amont/aval est optionnelle) :
//   pré-anoxie → anaérobie (bio-P + FeCl3) → [retour liqueur mixte] → pré-dénit
//   → cuves aérées centrales (1 à 3, de type C/N sans média ou H avec média)
//   → dé-ox → post-dénit (méthanol, avec ou sans média) → ré-ox → clarificateur
//
// Six configurations de cuves centrales, qui déterminent la routine de
// dimensionnement employée :
//   1 C/N-H   2 C/N-H-C/N   3 C/N-H-H   4 H-H   5 H-C/N   6 H
// Les configurations sans cuve C/N (4 et 6) sont dimensionnées par dichotomie
// sur le volume total ; les autres par point fixe sur la charge massique
// appliquée, la nitrification se répartissant entre le média (vitesse
// surfacique) et la liqueur mixte (vitesse massique).
//
// Deux passes, comme dans le classeur VBA :
//   dimensionnement()     sur l'eau nominale, à T_design  → volumes
//   fonctionnement_reel() sur l'eau réelle, à T_exploit   → performances, air
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - dimensionnement() est relancé jusqu'à stabilisation des variables qu'il
//    corrige d'un appel à l'autre (reox_nit, boues_methanol, DBO_sortie,
//    recirculation_taux, NO3_last_aerated_tank). Dans le classeur VBA ces corrections ne
//    prennent effet qu'à l'itération suivante du programme principal ;
//  - la borne haute de la dichotomie de dimensionne_config_4_6 est corrigée :
//    le VBA écrit `1000 * nit / 24 * MES * rate`, c'est-à-dire une
//    multiplication là où une division est attendue, ce qui donne une borne
//    inférieure au volume cherché et bloque la dichotomie ;
//  - au fonctionnement réel, le VBA ajoute les besoins en O2 dus aux sulfures
//    à `cuve_aeree_O2_besoins(I)` alors que la variable de boucle est `cuve` :
//    la contribution est donc perdue. Reproduit, et signalé ;
//  - le choix du média et de l'aérateur est proposé globalement et non cuve par
//    cuve (le calcul, lui, reste par cuve) ;
//  - gardes d'itérations sur toutes les boucles (VBA : boucles non bornées).
// ---------------------------------------------------------------------------
import { defineNode } from '../core/engine.js'
import { makeStream } from '../core/stream.js'
import { CONST, HYP, besoinsO2HS, ratio } from '../core/hypotheses.js'

// --- médias : surface spécifique (m²/m³) et déplacement de liqueur ---------
const MEDIA = {
  K1: { label: 'AnoxKaldnes K1', surface: 500, deplacement: 0.14, fillingMaxH: 0, fillingMaxPDN: 0.5 },
  K1_heavy: { label: 'AnoxKaldnes K1 heavy', surface: 500, deplacement: 0.14, fillingMaxH: 0, fillingMaxPDN: 0.5 },
  K3: { label: 'AnoxKaldnes K3', surface: 500, deplacement: 0.11, fillingMaxH: 0.65, fillingMaxPDN: 0.5 },
  K5: { label: 'AnoxKaldnes K5', surface: 800, deplacement: 0.123, fillingMaxH: 0.6, fillingMaxPDN: 0.5 },
  chip_P: { label: 'Chip P', surface: 900, deplacement: 0.173, fillingMaxH: 0.55, fillingMaxPDN: 0 },
  chip_M: { label: 'Chip M', surface: 1200, deplacement: 0.23, fillingMaxH: 0.55, fillingMaxPDN: 0 },
}
const MEDIA_OPTIONS = Object.entries(MEDIA).map(([value, m]) => ({ value, label: m.label }))

// --- configurations de cuves aérées (config_*) -----------------------------
const CONFIGS = {
  CN_H: { label: 'C/N – H', types: ['CN', 'H'] },
  CN_H_CN: { label: 'C/N – H – C/N', types: ['CN', 'H', 'CN'] },
  CN_H_H: { label: 'C/N – H – H', types: ['CN', 'H', 'H'] },
  H_H: { label: 'H – H', types: ['H', 'H'] },
  H_CN: { label: 'H – C/N', types: ['H', 'CN'] },
  H: { label: 'H (une seule cuve)', types: ['H'] },
}

// --- hypothèses figées de la classe (Sub hypotheses) -----------------------
const H = {
  critere_convergence: 1e-6,
  delta_NO3_post_DN: 6,
  ratio_EB_MV_MES: [1.213, -1 / 3, -0.1429],
  ratio_EB_MV_MES_min: 0.4,
  ratio_EB_MV_MES_max: 0.8,
  ratio_ED_MV_MES: 0.75,
  ratio_N_assimile: 0.06,
  ratio_N_assimile_methanol: 0.08,
  ratio_DCO_dure_total: 0.33 * 0.13,
  ratio_boues_nit: 0.125,
  ratio_boues_DBOin: 0.6,
  recirculation_O2: 1, // mg/L d'O2 dissous dans les boues recirculées
  ratio_NO3eq_O2: 0.35,
  pre_anoxie_NO3: 0.3,
  ratio_DBO_denit_max: 15,
  ratio_O2_denit: -2.86,
  ratio_O2_DBO: 0.75,
  anaerobie_P_limite: 10, // mg/L
  anaerobie_HRT_P_low: 1.3, // h
  anaerobie_HRT_P_high: 2, // h
  ratio_P_assimile: 0.01,
  anaerobie_DBO_Pt: 5, // kg DBO soluble éliminée par kg de P traité en bio-P
  MM_P: 31,
  MM_FeO3H3: 55.85 + 3 * 17,
  MM_FePO4: 55.85 + 31 + 64,
  MM_FeCl3: 162.5,
  deox_HRT_mini: 5 / 60, // h
  rate_DCO_methanol: 1.5, // kg O2 / kg méthanol
  ratio_boues_DCO_methanol: 0.25, // kg MES / kg O2
  hyp_cinetique_post_denit_15: 6, // g N-NO3/(kg MVS·h) sans média
  cinetique_post_denit_15_mini: 3,
  a_cinetique_postDN_media_10: [2.29, 2.9],
  charge_radier_max: { racle: 5, racle_suce: 7.5, kruger: 9 },
  facteur_boues_C_lim: 1200,
  O2_facteur_beta: 0.99,
  hauteur_diffuseur: 0.25,
  correctif_T_K: 1.024,
  T_ref_K: 20,
  ratio_kgO2_Nm3air: 0.3,
  agitation_fct: 24,
  ratio_Norgaduresoluble_DCOtot: 0.25 / 100,
  clarif_diametre_limite: 32,
  // Mox
  rate_N_30: { K3: 5.6, K5: 4.6, chip_M: 4 }, // × 2 par rapport au MBBR
}
const EXCES_O2_REGULATION = { horloge: 2, sans_variateur: 1, variateur: 0.5, avance: 0 }
const OUI_NON = [{ value: 'non', label: 'non' }, { value: 'oui', label: 'oui' }]

/** respiration endogène de la liqueur mixte, en kg O2/(j·m³ de bassin) */
const respiration_gL = (T, MLSS) => 0.002 * MLSS * CONST.NOMBRE_HEURE_PAR_JOUR * Math.pow(1.07, T - 10)

// paramètres forçables par cuve aérée (jusqu'à 3)
function paramsCuvesAerees() {
  const out = []
  for (let i = 1; i <= 3; i++) {
    const g = `Cuve aérée ${i}`
    out.push({ key: `cuve${i}_volume`, label: 'Volume de la cuve', unit: 'm³', group: g, default: undefined, hint: 'dimensionné si non forcé' })
    out.push({ key: `cuve${i}_filling`, label: 'Taux de remplissage en média (cuve H)', unit: '-', group: g, default: undefined, hint: '0,5 par défaut' })
    out.push({ key: `cuve${i}_O2_dissous`, label: 'O2 dissous', unit: 'mg/L', group: g, default: undefined })
    out.push({ key: `cuve${i}_alpha`, label: 'Facteur alpha', unit: '-', group: g, default: undefined })
    out.push({ key: `cuve${i}_rdt_transfert`, label: "Rendement de transfert de l'O2", unit: '%/m', group: g, default: undefined })
    out.push({ key: `cuve${i}_hauteur`, label: "Hauteur d'eau", unit: 'm', group: g, default: undefined })
  }
  return out
}

// ---------------------------------------------------------------------------
export default defineNode({
  id: 'hybas',
  label: 'HybAS',
  short: 'HybAS',
  family: 'secondaire',
  vba: 'E5_HybAS.cls',
  description:
    "Procédé hybride boues activées + média fixé (IFAS) : pré-anoxie, zone anaérobie de déphosphatation biologique, pré-dénitrification, cuves aérées C/N et H, dé-oxygénation, post-dénitrification au méthanol, ré-oxygénation et clarificateur.",
  choices: [
    { key: 'configuration', label: 'Configuration des cuves aérées', default: 'CN_H', options: Object.entries(CONFIGS).map(([value, c]) => ({ value, label: c.label })) },
    { key: 'media_cuves_H', label: 'Média des cuves H', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'pre_anoxie', label: 'Cuve de pré-anoxie', default: 'non', options: OUI_NON },
    { key: 'anaerobie', label: 'Cuve anaérobie (bio-P)', default: 'non', options: OUI_NON },
    { key: 'pre_denit', label: 'Cuve de pré-dénitrification', default: 'oui', options: OUI_NON },
    { key: 'deox', label: 'Cuve de dé-oxygénation', default: 'oui', options: OUI_NON },
    { key: 'post_denit', label: 'Cuve de post-dénitrification', default: 'non', options: [
      { value: 'non', label: 'non' }, { value: 'sans_media', label: 'oui, sans média' }, { value: 'avec_media', label: 'oui, avec média' },
    ] },
    { key: 'media_post_denit', label: 'Média de la post-dénitrification', default: 'K1', options: MEDIA_OPTIONS },
    { key: 'reox', label: 'Cuve de ré-oxygénation', default: 'non', options: OUI_NON },
    { key: 'mode_Mox', label: 'Fonctionnement Mox', default: 'non', options: [{ value: 'non', label: 'non' }, { value: 'oui', label: 'oui' }] },
    { key: 'aerateur', label: 'Type de diffuseurs', default: 'fine_bulle', options: [{ value: 'fine_bulle', label: 'fines bulles' }, { value: 'moyenne_bulle', label: 'moyennes bulles' }] },
    { key: 'racleur', label: 'Type de racleur du clarificateur', default: 'racle_suce', options: [
      { value: 'racle', label: 'raclé' }, { value: 'racle_suce', label: 'raclé-sucé' }, { value: 'kruger', label: 'Kruger' },
    ] },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
      { value: 'roots', label: 'tri-lobes type roots' }, { value: 'vis', label: 'surpresseurs à vis' }, { value: 'turbo', label: 'centrifuge type turbo' },
    ] },
    { key: 'regulation', label: "Régulation de l'aération", default: 'variateur', options: [
      { value: 'horloge', label: 'sur horloge' }, { value: 'sans_variateur', label: 'sans variateur' }, { value: 'variateur', label: 'classique – avec variateur' }, { value: 'avance', label: 'avancé' },
    ] },
  ],
  params: [
    { key: 'outlet_NH4', label: 'N-NH4 en sortie', unit: 'mg/L', group: 'Objectifs de traitement', default: 3 },
    { key: 'outlet_NO3', label: 'N-NO3 en sortie', unit: 'mg/L', group: 'Objectifs de traitement', default: undefined, hint: '5 avec post-dénit, sinon 10' },
    { key: 'outlet_DBO_soluble', label: 'DBO5 soluble en sortie', unit: 'mg/L', group: 'Objectifs de traitement', default: 2 },
    { key: 'Norga_dure_soluble', label: 'N organique dur et soluble', unit: 'mg/L', group: 'Objectifs de traitement', default: undefined, hint: '0,25 % de la DCO d\'entrée' },
    { key: 'pre_denit_NO3', label: 'N-NO3 en sortie de pré-dénitrification', unit: 'mg/L', group: 'Objectifs de traitement', default: 0.1 },
    { key: 'NO3_last_aerated_tank', label: 'N-NO3 en sortie de la dernière cuve aérée', unit: 'mg/L', group: 'Objectifs de traitement', default: undefined, hint: 'NO3 sortie + 6 avec post-dénit' },
    { key: 'nominal_MES_concentration', label: 'MES dans les bassins (nominal)', unit: 'g/L', group: 'Liqueur mixte', default: undefined, hint: '4 g/L (1 g/L en Mox)' },
    { key: 'nominal_MV_MES', label: 'MV/MES (nominal)', unit: '-', group: 'Liqueur mixte', default: undefined, hint: 'corrélation sur MES/DCO et DCO/DBO' },
    { key: 'reel_MES_concentration', label: 'MES dans les bassins (réel)', unit: 'g/L', group: 'Liqueur mixte', default: undefined, hint: 'proratisée sur la production de boues' },
    { key: 'reel_MV_MES', label: 'MV/MES (réel)', unit: '-', group: 'Liqueur mixte', default: undefined },
    { key: 'recirculation_taux', label: 'Taux de recirculation des boues du clarificateur', unit: '-', group: 'Recirculation', default: 1 },
    { key: 'recirculation_P_refoulement', label: 'Pression de refoulement de la recirculation des boues', unit: 'mCE', group: 'Recirculation', default: 5 },
    { key: 'recirculation_pompe_rdt', label: 'Rendement global des pompes de recirculation des boues', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
    { key: 'nominal_recirculation_MLSS_taux', label: 'Taux de recirculation de liqueur mixte (nominal)', unit: '-', group: 'Recirculation', default: undefined, hint: 'calculé sur le bilan NO3' },
    { key: 'reel_recirculation_MLSS_taux', label: 'Taux de recirculation de liqueur mixte (réel)', unit: '-', group: 'Recirculation', default: undefined, hint: 'repris du nominal si non forcé' },
    { key: 'recirculation_MLSS_P_refoulement', label: 'Pression de refoulement de la recirculation de liqueur mixte', unit: 'mCE', group: 'Recirculation', default: 5 },
    { key: 'recirculation_MLSS_pompe_rdt', label: 'Rendement global des pompes de liqueur mixte', unit: '-', group: 'Recirculation', default: 0.7 * 0.88 },
    { key: 'pre_anoxie_volume', label: 'Volume de la cuve de pré-anoxie', unit: 'm³', group: 'Cuves amont', default: undefined, hint: 'dimensionné si non forcé' },
    { key: 'anaerobie_volume', label: 'Volume de la cuve anaérobie', unit: 'm³', group: 'Cuves amont', default: undefined, hint: 'HRT 1,3 h (2 h si Pt ≥ 10 mg/L)' },
    { key: 'anaerobie_FeCl3_flux', label: 'Consommation de FeCl3 pur', unit: 'kg/j', group: 'Cuves amont', default: undefined, hint: 'calculée si non forcée' },
    { key: 'pre_denit_volume', label: 'Volume de la cuve de pré-dénitrification', unit: 'm³', group: 'Cuves amont', default: undefined, hint: 'dimensionné si non forcé' },
    { key: 'ratio_Nnit_media_Nnit_total', label: 'Part de N nitrifié sur le média', unit: '-', group: 'Répartition de la nitrification', default: undefined, hint: '0,6 en présence de cuves C/N' },
    { key: 'ratio_Nnit_mediaH1_Nnit_media', label: 'Part nitrifiée sur le média de H1 / média total', unit: '-', group: 'Répartition de la nitrification', default: undefined, hint: '0,5 si deux cuves H' },
    { key: 'ratio_Nnit_CN1_Nnit_CN12', label: 'Part nitrifiée en C/N1 / (C/N1 + C/N2)', unit: '-', group: 'Répartition de la nitrification', default: undefined, hint: '0,6 si deux cuves C/N' },
    ...paramsCuvesAerees(),
    { key: 'deox_volume', label: 'Volume de la cuve de dé-ox', unit: 'm³', group: 'Dé-Ox / Ré-Ox', default: undefined, hint: 'dimensionné si non forcé' },
    { key: 'deox_O2_dissous', label: 'O2 dissous en dé-ox', unit: 'mg/L', group: 'Dé-Ox / Ré-Ox', default: 1 },
    { key: 'reox_volume', label: 'Volume de la cuve de ré-ox', unit: 'm³', group: 'Dé-Ox / Ré-Ox', default: undefined, hint: 'HRT fonction du coefficient de pointe' },
    { key: 'reox_O2_dissous', label: 'O2 dissous en ré-ox', unit: 'mg/L', group: 'Dé-Ox / Ré-Ox', default: undefined },
    { key: 'reox_O2_alpha', label: 'Facteur alpha en ré-ox', unit: '-', group: 'Dé-Ox / Ré-Ox', default: undefined },
    { key: 'reox_O2_rdt_transfert', label: "Rendement de transfert de l'O2 en ré-ox", unit: '%/m', group: 'Dé-Ox / Ré-Ox', default: undefined },
    { key: 'reox_hauteur_bassin', label: "Hauteur d'eau de la cuve de ré-ox", unit: 'm', group: 'Dé-Ox / Ré-Ox', default: 8 },
    { key: 'post_denit_volume', label: 'Volume de la cuve de post-dénitrification', unit: 'm³', group: 'Post-dénitrification', default: undefined, hint: 'dimensionné si non forcé' },
    { key: 'post_denit_media_filling', label: 'Taux de remplissage en média de la post-dénit', unit: '-', group: 'Post-dénitrification', default: 0.4 },
    { key: 'post_denit_methanol_flux', label: 'Consommation de méthanol', unit: 'kg/j', group: 'Post-dénitrification', default: undefined, hint: 'calculée si non forcée' },
    { key: 'diffuseur_encrassement', label: 'Ancienneté des diffuseurs', unit: 'an(s)', group: 'Aération', default: 0 },
    { key: 'air_P_refoulement_moyenne', label: 'Pression de refoulement des surpresseurs', unit: 'mCE', group: 'Aération', default: undefined, hint: 'hauteur max + 2 + 0,25 × encrassement' },
    { key: 'surpresseur_conso_spec', label: 'Consommation spécifique des surpresseurs', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'O2_besoin_total', label: 'Besoin total en O2 (réel)', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'air_Q_Nm3j', label: "Débit d'air process (réel)", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'ratio_elec_pre_anoxie', label: "Ratio d'agitation en pré-anoxie", unit: 'W/m³', group: 'Agitation', default: 3 },
    { key: 'ratio_elec_anaerobie', label: "Ratio d'agitation en anaérobie", unit: 'W/m³', group: 'Agitation', default: 10 },
    { key: 'ratio_elec_pre_denit', label: "Ratio d'agitation en pré-dénit", unit: 'W/m³', group: 'Agitation', default: 3 },
    { key: 'ratio_elec_deox', label: "Ratio d'agitation en dé-ox", unit: 'W/m³', group: 'Agitation', default: 3 },
    { key: 'ratio_elec_post_denit', label: "Ratio d'agitation en post-dénit", unit: 'W/m³', group: 'Agitation', default: undefined, hint: '3 sans média, 7,5 avec média' },
    { key: 'outlet_reel_DBO', label: 'DBO5 en sortie (réel)', unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'calculée si non forcée' },
    { key: 'outlet_reel_NH4', label: 'N-NH4 en sortie (réel)', unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'calculée si non forcée' },
    { key: 'outlet_reel_NO3', label: 'N-NO3 en sortie (réel)', unit: 'mg/L', group: 'Sorties réelles', default: undefined, hint: 'calculée si non forcée' },
    { key: 'sortie_MES', label: "MES de l'effluent", unit: 'mg/L', group: 'Clarificateur', default: 20 },
    { key: 'boues_indice_Mohlman', label: 'Indice de boues', unit: 'mL/g', group: 'Clarificateur', default: 100 },
    { key: 'clarif_hauteur', label: 'Hauteur du clarificateur', unit: 'm', group: 'Clarificateur', default: 4 },
    { key: 'clarif_vitesse_max', label: 'Vitesse hydraulique maximale', unit: 'm/h', group: 'Clarificateur', default: undefined, hint: 'calculée sur le volume corrigé' },
    { key: 'clarif_surface', label: 'Surface de décantation', unit: 'm²', group: 'Clarificateur', default: undefined, hint: 'calculée si non forcée' },
    { key: 'boues_concentration', label: 'Concentration des boues extraites', unit: 'g/L', group: 'Boues', default: undefined, hint: 'calculée sur le bilan MES' },
    { key: 'boues_MES', label: 'Boues à extraire (réel)', unit: 'kg MES/j', group: 'Boues', default: undefined, hint: 'calculées si non forcées' },
    { key: 'extraction_P_refoulement', label: "Pression de refoulement de l'extraction", unit: 'mCE', group: 'Boues', default: 5 },
    { key: 'extraction_pompe_rdt', label: "Rendement global des pompes d'extraction", unit: '-', group: 'Boues', default: 0.7 * 0.88 },
  ],

  compute(ctx) {
    const { site, choices, forced } = ctx
    const warnings = []
    const dire = (m) => { if (!warnings.includes(m)) warnings.push(m) }
    const T_design = site.T_eau_design
    const T_reel = site.T_eau_exploit
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)

    // =======================================================================
    // lecture_choix
    // =======================================================================
    const types = CONFIGS[choices.configuration].types
    const nb_cuves_aerees = types.length
    const estH = (i) => types[i - 1] === 'H'
    const nb_cuves_H = types.filter((t) => t === 'H').length
    const nb_cuves_CN = types.filter((t) => t === 'CN').length
    const choix_pre_anoxie = choices.pre_anoxie === 'oui'
    const choix_anaerobie = choices.anaerobie === 'oui'
    const choix_pre_denit = choices.pre_denit === 'oui'
    const choix_deox = choices.deox === 'oui'
    const choix_reox = choices.reox === 'oui'
    const choix_post_denit = choices.post_denit // 'non' | 'sans_media' | 'avec_media'
    const postDN = choix_post_denit !== 'non'
    const choix_fct_Mox = choices.mode_Mox === 'oui'
    const media_H = choices.media_cuves_H
    const media_PDN = choices.media_post_denit
    const Smedia = (i) => (estH(i) ? MEDIA[media_H].surface : 0)
    const deplacement = (i) => (estH(i) ? MEDIA[media_H].deplacement : 0)

    const erreurs = []
    if (choix_anaerobie && !choix_pre_anoxie) erreurs.push("Une cuve anaérobie sans pré-anoxie en amont n'est pas admise.")
    const nbCuvesAmontAval = [choix_pre_anoxie, choix_anaerobie, choix_pre_denit, choix_deox, choix_reox, postDN].filter(Boolean).length
    if (choix_fct_Mox && nbCuvesAmontAval + nb_cuves_aerees > 1) erreurs.push('Le mode Mox ne peut être utilisé qu\'en configuration H seule, sans cuve en amont ni en aval.')
    if (choix_fct_Mox && !['K3', 'K5', 'chip_M'].includes(media_H)) erreurs.push('Média non utilisable en configuration Mox — choix possibles : K3, K5, Chip M.')
    if (erreurs.length) {
      return { outNominal: makeStream(ctx.inNominal), outReel: makeStream(ctx.inReel), results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: erreurs }
    }

    // =======================================================================
    // hypotheses
    // =======================================================================
    const primaire = ctx.upstream?.primaire === true
    const primaire_simple = primaire && ctx.upstream?.primaire_reactif !== true
    const primaire_reactif = ctx.upstream?.primaire_reactif === true
    const ratio_boues_MESin = choix_fct_Mox ? 0.35 : primaire ? 0.35 : 0.5
    const ratio_NO3f_nit = choix_fct_Mox ? 0.11 : 1
    const ratio_O2_nit = choix_fct_Mox ? 1.94 : 4.57
    const rateDCOdenit = (T) => (T < 5 ? 5 : T > 15 ? 4 : 5 - 0.1 * (T - 5))
    const rate_DCO_apportee_denit_nominal = rateDCOdenit(T_design)
    const rate_DCO_apportee_denit_reel = rateDCOdenit(T_reel)
    const a_cinetique_postDN_nominal = Math.pow(1.08, T_design - 15)
    const a_cinetique_postDN_reel = Math.pow(1.08, T_reel - 15)
    const a_cinetique_postDN_media_nominal = Math.pow(1.06, T_design - 10)
    const a_cinetique_postDN_media_reel = Math.pow(1.06, T_reel - 10)
    const post_DN_SRT_mini = 1.35 * Math.pow(1.07, 14 - T_design)
    const pointe = site.pointe_TP
    const reox_HRT = pointe <= 1 ? 20 / 1440 : pointe >= 3 ? (25 - 5 * pointe) / 1440 : 10 / 1440 // j
    const cuve_H_V_Nmh_mini = choix_fct_Mox ? 0 : 7 // Nm³/(m²·h)

    /** vitesse de dénitrification (g N/(kg MVS·h)) fonction du rapport DBO/NO3 */
    function cinetique_denit(ratio_DBO_denit, T) {
      let r = Math.min(ratio_DBO_denit, H.ratio_DBO_denit_max)
      let v
      if (primaire_simple) {
        const lim = 5
        v = r < lim ? (0.8333 * r - 1.6667) / 0.77 : (0.8333 * lim - 1.6667 + 0.05 * (r - lim)) / 0.77
      } else if (primaire_reactif) {
        const lim = 4
        v = r < lim ? (1.25 * r - 2.5) / 0.77 : (1.25 * lim - 2.5 + 0.0455 * (r - lim)) / 0.77
      } else {
        const lim = 6.5
        v = r < lim ? (0.5556 * r - 1.1111) / 0.77 : (0.5556 * lim - 1.1111 + 0.0588 * (r - lim)) / 0.77
      }
      return v * Math.pow(1.08, T - 15)
    }
    /** vitesse de nitrification de la liqueur mixte (g N/(kg MES·h)) */
    const cinetique_nit_MLSS = (DOi, T, rate_bacteries_nit) =>
      (DOi <= 2.7 ? 63 * (DOi / (0.7 + DOi)) : 50) * rate_bacteries_nit * Math.pow(1.103, T - 15)
    /** rendement d'élimination de la DBO d'une cuve aérée, fonction de la charge massique */
    function rdt_DBO(Cm) {
      let r
      if (primaire_simple) r = 25 + 71 * Math.exp(-0.6 * Math.pow(Cm + 0.1, 2))
      else if (primaire_reactif) r = 14 + 80 * Math.exp(-0.6 * Math.pow(Cm + 0.1, 2))
      else r = 44.57 + 53 * Math.exp(-0.75 * Math.pow(Cm + 0.1, 2))
      return r / 100
    }
    /** coefficient de déphosphatation biologique Kp (kg P / kg Pt entrant) */
    function anaerobie_Kp(V_Vopt, NO3in_mgL, Pt, DBO) {
      const Kp1 = V_Vopt <= 1 ? V_Vopt * V_Vopt : V_Vopt <= 1.5 ? 0.7 + 0.3 * V_Vopt : 1.15
      const Kp2 = NO3in_mgL >= 0.3 ? Math.pow(0.3 / NO3in_mgL, 2) : 1.1 - NO3in_mgL / 3
      return 0.028 * Kp1 * Kp2 * (Pt > 0 ? DBO / Pt : 0)
    }
    /** ratio molaire Fe/P nécessaire pour précipiter deltaP */
    function ratio_molaire_precipitation(deltaP, Q) {
      const Ps = site.Pt_garantie
      const Ptbr = (deltaP / Q) * 1000
      let r = Math.pow(Ptbr + Ps, -0.5) - (Math.pow(Ps, -0.5) - 0.1)
      r = r / (Ptbr + Ps - (Ps - 0.1))
      return 1.1 - 6 * r
    }

    // =======================================================================
    // attribution_valeur_par_defaut
    // =======================================================================
    const exces_O2 = EXCES_O2_REGULATION[choices.regulation] ?? 0
    const outlet_NH4 = f('outlet_NH4') ?? 3
    let outlet_NO3 = f('outlet_NO3') ?? (postDN ? 5 : 10)
    const pre_denit_NO3 = choix_pre_denit ? (f('pre_denit_NO3') ?? 0.1) : 0
    let NO3_last_aerated_tank = f('NO3_last_aerated_tank') ?? (postDN ? outlet_NO3 + H.delta_NO3_post_DN : outlet_NO3)
    const outlet_DBO_soluble = f('outlet_DBO_soluble') ?? 2
    let recirculation_taux = f('recirculation_taux') ?? 1
    const nominal_MES_concentration = f('nominal_MES_concentration') ?? (choix_fct_Mox ? 1 : 4)
    const deox_O2_dissous = f('deox_O2_dissous') ?? 1
    const reox_O2_dissous = f('reox_O2_dissous') ?? 3 + exces_O2
    const fine = choices.aerateur === 'fine_bulle'
    const reox_O2_alpha = f('reox_O2_alpha') ?? (fine ? 0.65 : 0.8)
    const reox_O2_rdt_transfert = f('reox_O2_rdt_transfert') ?? ((fine ? 16 : 12) / 300) * 100
    const reox_hauteur_bassin = f('reox_hauteur_bassin') ?? 8
    let hauteur_bassin_max = reox_hauteur_bassin

    let post_denit_media_filling = f('post_denit_media_filling') ?? 0.4
    if (f('post_denit_media_filling') != null && post_denit_media_filling > MEDIA[media_PDN].fillingMaxPDN) {
      dire('Post-dénitrification : taux de remplissage supérieur au maximum admissible pour ce média.')
    }
    let ratio_Nnit_media_Nnit_total = f('ratio_Nnit_media_Nnit_total') ?? (nb_cuves_CN !== 0 ? 0.6 : 0)
    const ratio_Nnit_media_force = f('ratio_Nnit_media_Nnit_total') != null
    const ratio_Nnit_mediaH1 = nb_cuves_H === 2 ? (f('ratio_Nnit_mediaH1_Nnit_media') ?? 0.5) : nb_cuves_CN === 0 ? 1 : 1
    const ratio_Nnit_CN1 = nb_cuves_CN === 2 ? (f('ratio_Nnit_CN1_Nnit_CN12') ?? 0.6) : 1

    const cuve_volume = new Array(nb_cuves_aerees + 1).fill(0)
    const cuve_filling = new Array(nb_cuves_aerees + 1).fill(0)
    const cuve_O2_dissous = new Array(nb_cuves_aerees + 1).fill(0)
    const cuve_alfa = new Array(nb_cuves_aerees + 1).fill(0)
    const cuve_rdt_transfert = new Array(nb_cuves_aerees + 1).fill(0)
    const cuve_hauteur = new Array(nb_cuves_aerees + 1).fill(0)
    for (let i = 1; i <= nb_cuves_aerees; i++) {
      if (estH(i)) {
        cuve_filling[i] = f(`cuve${i}_filling`) ?? 0.5
        if (f(`cuve${i}_filling`) != null && cuve_filling[i] > MEDIA[media_H].fillingMaxH) {
          dire(`Cuve aérée ${i} : taux de remplissage supérieur au maximum admissible pour ce média.`)
        }
      }
      cuve_O2_dissous[i] = f(`cuve${i}_O2_dissous`) ??
        (i === 1 && !estH(i) ? 2 + exces_O2 : choix_fct_Mox ? 0.5 : 3 + exces_O2)
      cuve_hauteur[i] = f(`cuve${i}_hauteur`) ?? (choix_fct_Mox ? 7 : 8)
      if (cuve_hauteur[i] > hauteur_bassin_max) hauteur_bassin_max = cuve_hauteur[i]
      cuve_alfa[i] = f(`cuve${i}_alpha`) ?? (fine ? (estH(i) ? 0.75 : 0.65) : 0.8)
      cuve_rdt_transfert[i] = f(`cuve${i}_rdt_transfert`) ?? ((fine ? 16 : 12) / 300) * 100
    }
    if (choix_deox && deox_O2_dissous > cuve_O2_dissous[nb_cuves_aerees]) {
      return {
        outNominal: makeStream(ctx.inNominal), outReel: makeStream(ctx.inReel), results: [],
        electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["L'O2 dissous visé en dé-ox dépasse celui de la dernière cuve aérée."],
      }
    }

    const ratio_elec_pre_anoxie = f('ratio_elec_pre_anoxie') ?? 3
    const ratio_elec_anaerobie = f('ratio_elec_anaerobie') ?? 10
    const ratio_elec_pre_denit = f('ratio_elec_pre_denit') ?? 3
    const ratio_elec_deox = f('ratio_elec_deox') ?? 3
    const ratio_elec_post_denit = f('ratio_elec_post_denit') ?? (choix_post_denit === 'avec_media' ? 7.5 : 3)
    const diffuseur_encrassement = f('diffuseur_encrassement') ?? 0
    const air_P_refoulement_moyenne = f('air_P_refoulement_moyenne') ?? hauteur_bassin_max + 2 + 0.25 * diffuseur_encrassement
    const recirculation_MLSS_P_refoulement = f('recirculation_MLSS_P_refoulement') ?? 5
    const recirculation_MLSS_pompe_rdt = f('recirculation_MLSS_pompe_rdt') ?? 0.7 * 0.88
    const ratio_elec_recirculation_MLSS = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * recirculation_MLSS_pompe_rdt)
    const recirculation_P_refoulement = f('recirculation_P_refoulement') ?? 5
    const recirculation_pompe_rdt = f('recirculation_pompe_rdt') ?? 0.7 * 0.88
    const ratio_elec_recirculation = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * recirculation_pompe_rdt)
    const boues_indice_Mohlman = f('boues_indice_Mohlman') ?? 100
    const sortie_MES = f('sortie_MES') ?? 20
    const clarif_hauteur = f('clarif_hauteur') ?? 4
    const extraction_P_refoulement = f('extraction_P_refoulement') ?? 5
    const extraction_pompe_rdt = f('extraction_pompe_rdt') ?? 0.7 * 0.88
    const ratio_elec_extraction = CONST.ACCELERATION_PESANTEUR_m_s2 / (CONST.NOMBRE_SECONDE_PAR_HEURE * extraction_pompe_rdt)

    /** débit d'air d'une cuve (Private Function air_Q_Nm3j_cuve) */
    function air_Q_cuve(DOi, O2_besoin, T, volume, avecMedia, hauteur, alfa, rdt) {
      const Patm = (HYP.Patm_P0 * Math.pow(HYP.Patm_a0 + HYP.Patm_a1 * site.altitude, HYP.Patm_a2)) / 100
      const delta_P = (0.35 / 10.33) * (HYP.Patm_P0 / 100) * (hauteur - H.hauteur_diffuseur)
      let O2sat20 = 0, O2satT = 0
      for (let k = 0; k <= 4; k++) {
        O2sat20 += HYP.O2sat_coef[k] * Math.pow(20, k)
        O2satT += HYP.O2sat_coef[k] * Math.pow(T, k)
      }
      const corr = (Patm + delta_P) / (HYP.Patm_P0 / 100)
      O2sat20 *= corr; O2satT *= corr
      let K = (alfa * (H.O2_facteur_beta * O2satT - DOi)) / O2sat20
      K *= Math.pow(H.correctif_T_K, T - H.T_ref_K)
      const denom = K * (rdt / 100) * (hauteur - H.hauteur_diffuseur) * H.ratio_kgO2_Nm3air
      let air = denom > 0 ? O2_besoin / denom : 0
      if (avecMedia && hauteur > 0) {
        const section = volume / hauteur
        if (air / CONST.NOMBRE_HEURE_PAR_JOUR < cuve_H_V_Nmh_mini * section) {
          air = cuve_H_V_Nmh_mini * section * CONST.NOMBRE_HEURE_PAR_JOUR
          dire("Débit d'air augmenté pour assurer le brassage du média dans les cuves H.")
        }
      }
      return air
    }

    /** vitesse surfacique de nitrification sur média (g N/(m²·j)) */
    function surface_rate_media(i, DBO_appliquee, T) {
      if (choix_fct_Mox) return (H.rate_N_30[media_H] ?? 0) * Math.pow(1.07, T - 30)
      const Cm10 = DBO_appliquee * Math.pow(1.06, 10 - T)
      let k
      if (Cm10 < 0.13) k = 0.55
      else if (Cm10 < 0.48) k = (-0.11 / 0.35) * Cm10 + 0.55 + 0.0143 / 0.35
      else k = 0.44 * Math.exp(-1000 * Math.pow(Cm10 - 0.48, 2))
      const base = (cuve_O2_dissous[i] - 1) / 3.2
      if (!(base > 0)) return 0
      return k * Math.pow(base, 0.7) * Math.pow(1.06, T - 10)
    }
    /** volume d'une cuve H déduit de la charge à nitrifier sur le média */
    function volume_cuve_H(i, Nnit_media, DBO_appliquee, T) {
      const rate = surface_rate_media(i, DBO_appliquee, T)
      if (!(rate > 0) || !(cuve_filling[i] > 0)) return 0
      return (1000 * Nnit_media) / (rate * Smedia(i)) / cuve_filling[i]
    }
    /** nitrification totale d'une cuve H : média + liqueur mixte */
    function nitrification_H_kgj(i, DBO_appliquee, T, rate_N_MLSS, MES) {
      const rate = surface_rate_media(i, DBO_appliquee, T)
      const surMedia = (rate * Smedia(i) * cuve_volume[i] * cuve_filling[i]) / 1000
      const surMLSS = (rate_N_MLSS * cuve_volume[i] * (1 - cuve_filling[i] * deplacement(i)) * CONST.NOMBRE_HEURE_PAR_JOUR * MES) / 1000
      return { total: surMedia + surMLSS, surMedia }
    }

    // =======================================================================
    // dimensionnement — relancé jusqu'à stabilisation (voir en-tête)
    // =======================================================================
    // état de classe, nul au premier passage comme en VBA
    let reox_nit = 0
    let boues_methanol = 0
    let DBO_sortie = 0
    let boues_Q = 0
    let boues_concentration = 0
    let nominal_MV_MES = 0
    let boues_pdtes_nominal = 0
    let rate_bacteries_nit = 0
    let pre_anoxie_volume = 0
    let pre_anoxie_denit = 0
    let anaerobie_volume = 0
    let anaerobie_FeCl3_flux = 0
    let boues_minerales_pdtes = 0
    let pre_denit_volume = 0
    let pre_denit_denit = 0
    let deox_volume = 0
    let deox_nit = 0
    let reox_volume = 0
    let post_denit_volume = 0
    let post_denit_methanol_flux = 0
    let post_denit_DCO_apportee = 0
    let cinetique_post_DN_15 = H.hyp_cinetique_post_denit_15
    let nominal_recirculation_MLSS_taux = 0
    let volume_total_bassins = 0
    let clarif_surface = 0
    let clarif_vitesse_max = 0
    let clarif_vitesse_max_recalc = 0
    let clarif_charge_radier_nominal = 0
    let nominal_O2_besoin_total = 0
    let nominal_air_Q = 0
    let outNominal = makeStream(ctx.inNominal)
    let Norga_dure_soluble = 1.5

    /** nitrification portée par la liqueur mixte d'une cuve (kg N/j) */
    function nitMLSS(i, rate) {
      return (rate * cuve_volume[i] * (1 - cuve_filling[i] * deplacement(i)) * CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration) / 1000
    }

    /** dichotomie sur la dénitrification d'une cuve de volume imposé */
    function denitVolumeImpose(volume, NO3_dispo, DBO, MES, MV_MES, T) {
      let min = 0
      let max = NO3_dispo
      let d = NO3_dispo / 2
      let garde = 0
      while (d > 0 && (max - min) / d > H.critere_convergence && garde++ < 300) {
        const capacite = (volume * (CONST.NOMBRE_HEURE_PAR_JOUR * cinetique_denit(DBO / d, T) * MES * MV_MES)) / 1000
        if (capacite < d) max = d
        else min = d
        d = (max + min) / 2
      }
      return d
    }

    for (let passe = 1; passe <= 25; passe++) {
      const s = ctx.inNominal
      const Q = s.Q
      Norga_dure_soluble = f('Norga_dure_soluble') ?? H.ratio_Norgaduresoluble_DCOtot * ((site.DCO_nominal / site.Q_nominal) * 1000)
      nominal_air_Q = 0
      nominal_O2_besoin_total = 0

      // eau entrante + recirculation des boues du clarificateur
      let interm_Q = Q + recirculation_taux * Q
      let interm_DBO = s.DBO + ((recirculation_taux * Q) / 1000) * outlet_DBO_soluble
      let interm_NO3 = s.NO3 + (outlet_NO3 * recirculation_taux * Q) / 1000
      let interm_Pt = s.Pt + (site.Pt_garantie * recirculation_taux * Q) / 1000
      let O2_dissous_flux = (H.recirculation_O2 * recirculation_taux * Q) / 1000

      nominal_MV_MES = f('nominal_MV_MES') ?? (primaire ? H.ratio_ED_MV_MES
        : Math.min(H.ratio_EB_MV_MES_max, Math.max(H.ratio_EB_MV_MES_min,
          H.ratio_EB_MV_MES[0] + H.ratio_EB_MV_MES[1] * (s.MES / s.DCO) + H.ratio_EB_MV_MES[2] * (s.DCO / s.DBO))))

      const N_nit_total = s.NK - (Norga_dure_soluble * Q) / 1000 - (outlet_NH4 * Q) / 1000
        - H.ratio_N_assimile * (s.DBO - (outlet_DBO_soluble * Q) / 1000)
        - H.ratio_N_assimile_methanol * boues_methanol
      const N_nit_aerees = N_nit_total - reox_nit

      // taux de recirculation de liqueur mixte, sur le bilan NO3
      const deltaNO3 = NO3_last_aerated_tank - pre_denit_NO3
      if (f('nominal_recirculation_MLSS_taux') != null) {
        nominal_recirculation_MLSS_taux = f('nominal_recirculation_MLSS_taux')
        if (f('NO3_last_aerated_tank') == null) {
          NO3_last_aerated_tank = pre_denit_NO3 + (ratio_NO3f_nit * N_nit_aerees) / (((1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q) / 1000)
        }
      } else {
        nominal_recirculation_MLSS_taux = deltaNO3 > 0
          ? (N_nit_aerees * ratio_NO3f_nit - ((1 + recirculation_taux) * Q * deltaNO3) / 1000) / ((deltaNO3 * Q) / 1000)
          : -1
        if (nominal_recirculation_MLSS_taux < 0) {
          // objectif inatteignable : on abaisse la consigne de 10 % et on rejoue
          if (postDN) { if (f('NO3_last_aerated_tank') == null) NO3_last_aerated_tank *= 0.9 }
          else if (f('outlet_NO3') == null) { outlet_NO3 *= 0.9; NO3_last_aerated_tank = outlet_NO3 }
        }
      }

      const boues_pdtes = ratio_boues_MESin * s.MES + H.ratio_boues_DBOin * s.DBO + H.ratio_boues_nit * N_nit_total + boues_methanol
      boues_pdtes_nominal = boues_pdtes
      rate_bacteries_nit = boues_pdtes > 0 ? (H.ratio_boues_nit * N_nit_total) / boues_pdtes : 0

      // ---- PRE-ANOXIE
      if (choix_pre_anoxie) {
        if (f('pre_anoxie_volume') != null) {
          pre_anoxie_volume = f('pre_anoxie_volume')
          pre_anoxie_denit = denitVolumeImpose(pre_anoxie_volume, interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux, s.DBO, nominal_MES_concentration, nominal_MV_MES, T_design)
        } else {
          pre_anoxie_denit = interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux - (H.pre_anoxie_NO3 * interm_Q) / 1000
          const r = Math.min(H.ratio_DBO_denit_max, s.DBO / pre_anoxie_denit)
          const cin = cinetique_denit(r, T_design)
          pre_anoxie_volume = cin > 0 ? (1000 * pre_anoxie_denit) / (CONST.NOMBRE_HEURE_PAR_JOUR * cin * nominal_MES_concentration * nominal_MV_MES) : 0
        }
        interm_NO3 = interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux - pre_anoxie_denit
        interm_DBO += H.ratio_O2_denit * pre_anoxie_denit
        O2_dissous_flux = 0
      } else { pre_anoxie_volume = 0; pre_anoxie_denit = 0 }

      // ---- ANAEROBIE (déphosphatation biologique + précipitation d'appoint)
      const Pt_mgL_EB = (s.Pt / Q) * 1000
      const anaerobie_V_optimal = ((Pt_mgL_EB < H.anaerobie_P_limite ? H.anaerobie_HRT_P_low : H.anaerobie_HRT_P_high) * Q) / CONST.NOMBRE_HEURE_PAR_JOUR
      let deltaP_synthese, deltaP_precipitation = 0, deltaP_bio = 0
      if (choix_anaerobie) {
        anaerobie_volume = f('anaerobie_volume') ?? anaerobie_V_optimal
        const V_Vopt = anaerobie_volume / anaerobie_V_optimal
        deltaP_bio = interm_Pt * anaerobie_Kp(V_Vopt, (interm_NO3 / interm_Q) * 1000, interm_Pt, s.DBO)
        interm_DBO -= H.anaerobie_DBO_Pt * deltaP_bio
        deltaP_synthese = H.ratio_P_assimile * (s.DBO - DBO_sortie)
        deltaP_precipitation = interm_Pt - (deltaP_bio + deltaP_synthese + (site.Pt_garantie * Q * (1 + recirculation_taux)) / 1000)
        if (deltaP_precipitation > 0) {
          const rm = ratio_molaire_precipitation(deltaP_precipitation, Q)
          anaerobie_FeCl3_flux = f('anaerobie_FeCl3_flux') ?? (rm * H.MM_FeCl3 * deltaP_precipitation) / H.MM_P
          boues_minerales_pdtes = (((rm - 1) * H.MM_FeO3H3) / H.MM_P + H.MM_FePO4 / H.MM_P) * deltaP_precipitation
        } else { deltaP_precipitation = 0; anaerobie_FeCl3_flux = 0; boues_minerales_pdtes = 0 }
        interm_Pt -= deltaP_synthese + deltaP_precipitation + deltaP_bio
        O2_dissous_flux = 0
      } else {
        anaerobie_volume = 0
        deltaP_synthese = H.ratio_P_assimile * (s.DBO - DBO_sortie)
        interm_Pt -= deltaP_synthese
      }

      // ---- RETOUR DE LIQUEUR MIXTE
      const O2_MLSS = choix_deox ? deox_O2_dissous : cuve_O2_dissous[nb_cuves_aerees]
      let O2_dissous_MLSS_flux = (O2_MLSS * nominal_recirculation_MLSS_taux * Q) / 1000
      O2_dissous_flux += O2_dissous_MLSS_flux
      interm_NO3 += (NO3_last_aerated_tank * nominal_recirculation_MLSS_taux * Q) / 1000
      interm_DBO += ((nominal_recirculation_MLSS_taux * Q) / 1000) * outlet_DBO_soluble

      // ---- PRE-DENITRIFICATION
      if (choix_pre_denit) {
        if (f('pre_denit_volume') != null) {
          pre_denit_volume = f('pre_denit_volume')
          pre_denit_denit = denitVolumeImpose(pre_denit_volume, interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux, s.DBO, nominal_MES_concentration, nominal_MV_MES, T_design)
        } else {
          pre_denit_denit = interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux
            - (pre_denit_NO3 * Q * (1 + recirculation_taux + nominal_recirculation_MLSS_taux)) / 1000
          const r = Math.min(H.ratio_DBO_denit_max, s.DBO / pre_denit_denit)
          const cin = cinetique_denit(r, T_design)
          pre_denit_volume = cin > 0 ? (1000 * pre_denit_denit) / (CONST.NOMBRE_HEURE_PAR_JOUR * cin * nominal_MES_concentration * nominal_MV_MES) : 0
        }
        interm_NO3 = interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux - pre_denit_denit
        interm_DBO += H.ratio_O2_denit * pre_denit_denit
        O2_dissous_flux = 0
      } else { pre_denit_volume = 0; pre_denit_denit = 0 }

      let cuves_centrales_DBOin = interm_DBO
      const O2_flux_entree_cuves = O2_dissous_flux

      // ---- DE-OXYGENATION (l'ordre du VBA place ce bloc avant les cuves centrales)
      if (choix_deox) {
        const deox_O2_conso = (((1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q) / 1000) * (cuve_O2_dissous[nb_cuves_aerees] - deox_O2_dissous)
        if (f('deox_volume') != null) deox_volume = f('deox_volume')
        else {
          const den = (ratio_O2_nit * nominal_MES_concentration * CONST.NOMBRE_HEURE_PAR_JOUR / 1000) * cinetique_nit_MLSS(deox_O2_dissous, T_design, rate_bacteries_nit)
            + respiration_gL(T_design, nominal_MES_concentration)
          deox_volume = den > 0 ? deox_O2_conso / den : 0
        }
        deox_nit = (deox_O2_conso - respiration_gL(T_design, nominal_MES_concentration) * deox_volume) / ratio_O2_nit
        if (deox_nit < 0) deox_nit = 0
        const deox_V_mini = ((1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q * H.deox_HRT_mini) / CONST.NOMBRE_HEURE_PAR_JOUR
        if (deox_volume < deox_V_mini) dire(`Dé-Ox : volume inférieur au minimum conseillé (${Math.round(deox_V_mini)} m³).`)
      } else { deox_volume = 0; deox_nit = 0 }

      // ---- RE-OXYGENATION
      if (choix_reox) {
        const O2_in_reox = postDN ? 0 : choix_deox ? deox_O2_dissous : cuve_O2_dissous[nb_cuves_aerees]
        const O2_exces = ((reox_O2_dissous - O2_in_reox) * Q * (1 + recirculation_taux)) / 1000
        reox_volume = f('reox_volume') ?? reox_HRT * Q
        reox_nit = (cinetique_nit_MLSS(reox_O2_dissous, T_design, rate_bacteries_nit) * reox_volume * CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration) / 1000
        const max_nit = (((1 + recirculation_taux) * Q) / 1000) * (reox_O2_dissous - O2_in_reox) / ratio_O2_nit
        if (reox_nit > max_nit) reox_nit = max_nit
        const O2_besoin_cuve = ratio_O2_nit * reox_nit + O2_exces
        nominal_O2_besoin_total += O2_besoin_cuve
        nominal_air_Q += air_Q_cuve(reox_O2_dissous, O2_besoin_cuve, T_design, reox_volume, false, reox_hauteur_bassin, reox_O2_alpha, reox_O2_rdt_transfert)
      } else { reox_volume = 0; reox_nit = 0 }

      // ---- POST-DENITRIFICATION
      if (postDN) {
        if (nominal_recirculation_MLSS_taux > 0) O2_dissous_MLSS_flux = (O2_dissous_MLSS_flux * (1 + recirculation_taux)) / nominal_recirculation_MLSS_taux
        const post_denit_denit = ((NO3_last_aerated_tank - outlet_NO3) * (1 + recirculation_taux) * Q) / 1000 + reox_nit + H.ratio_NO3eq_O2 * O2_dissous_MLSS_flux
        post_denit_DCO_apportee = rate_DCO_apportee_denit_nominal * post_denit_denit
        post_denit_methanol_flux = f('post_denit_methanol_flux') ?? post_denit_DCO_apportee / H.rate_DCO_methanol
        const Vf = f('post_denit_volume')
        if (choix_post_denit === 'sans_media') {
          post_denit_media_filling = 0
          let cin15 = H.hyp_cinetique_post_denit_15 * a_cinetique_postDN_nominal
          if (Vf != null) post_denit_volume = Vf
          else {
            post_denit_volume = (post_denit_denit * 1000) / (CONST.NOMBRE_HEURE_PAR_JOUR * cin15 * nominal_MES_concentration * nominal_MV_MES)
            if ((post_denit_volume * nominal_MES_concentration) / boues_pdtes < post_DN_SRT_mini) {
              post_denit_volume = (post_DN_SRT_mini * boues_pdtes) / nominal_MES_concentration
            }
          }
          if ((post_denit_volume * nominal_MES_concentration) / boues_pdtes < post_DN_SRT_mini) {
            dire(`Post-dénitrification : âge de boues insuffisant, volume conseillé ${Math.round((post_DN_SRT_mini * boues_pdtes) / nominal_MES_concentration)} m³.`)
          }
          cinetique_post_DN_15 = post_denit_volume > 0
            ? ((post_denit_denit * 1000) / (CONST.NOMBRE_HEURE_PAR_JOUR * post_denit_volume * nominal_MES_concentration * nominal_MV_MES)) / a_cinetique_postDN_nominal
            : 0
          if (cinetique_post_DN_15 < H.cinetique_post_denit_15_mini) {
            dire("Post-dénitrification : cinétique faible, l'ajout de média est conseillé.")
          }
        } else {
          const NO3_out = outlet_NO3 - reox_nit / ((Q / 1000) * (1 + recirculation_taux))
          let cin = ((H.a_cinetique_postDN_media_10[0] * NO3_out) / (NO3_out + H.a_cinetique_postDN_media_10[1])) * a_cinetique_postDN_media_nominal
          if (Vf != null) {
            post_denit_volume = Vf
            if (f('post_denit_media_filling') == null) {
              post_denit_media_filling = (1000 * post_denit_denit) / (post_denit_volume * MEDIA[media_PDN].surface * cin)
              if (post_denit_media_filling > MEDIA[media_PDN].fillingMaxPDN) post_denit_media_filling = MEDIA[media_PDN].fillingMaxPDN
            }
          } else {
            post_denit_volume = (1000 * post_denit_denit) / (post_denit_media_filling * MEDIA[media_PDN].surface * cin)
          }
        }
        boues_methanol = H.ratio_boues_DCO_methanol * post_denit_DCO_apportee
      } else { post_denit_volume = 0; post_denit_media_filling = 0; post_denit_methanol_flux = 0; post_denit_DCO_apportee = 0 }

      // ---- CUVES CENTRALES
      const cuves_centrales_nit = N_nit_aerees - deox_nit
      const cuves_centrales_DBOout = ((outlet_DBO_soluble * Q) / 1000) * (1 + recirculation_taux + nominal_recirculation_MLSS_taux)
      const cuve_O2_besoins = new Array(nb_cuves_aerees + 1).fill(0)
      dimensionneCuvesCentrales(cuves_centrales_nit, cuves_centrales_DBOin, N_nit_aerees, cuve_O2_besoins)

      // volumes forcés
      for (let i = 1; i <= nb_cuves_aerees; i++) if (f(`cuve${i}_volume`) != null) cuve_volume[i] = f(`cuve${i}_volume`)

      // besoins en O2 : nitrification (déjà calculée) + DBO éliminée + respiration + excès d'O2
      let O2_flux_amont = O2_flux_entree_cuves
      let DBO_courante = cuves_centrales_DBOin
      const cuve_DBO_eliminee = new Array(nb_cuves_aerees + 1).fill(0)
      for (let i = 1; i <= nb_cuves_aerees; i++) {
        let O2_exces = -O2_flux_amont
        const V_liqueur = estH(i) ? cuve_volume[i] * (1 - cuve_filling[i] * deplacement(i)) : cuve_volume[i]
        const Cm = V_liqueur > 0 ? DBO_courante / (V_liqueur * nominal_MES_concentration) : 0
        let DBO_sortante = DBO_courante * (1 - rdt_DBO(Cm))
        if (DBO_sortante < cuves_centrales_DBOout) DBO_sortante = cuves_centrales_DBOout
        cuve_DBO_eliminee[i] = DBO_courante - DBO_sortante
        O2_flux_amont = (cuve_O2_dissous[i] * Q * (1 + recirculation_taux + nominal_recirculation_MLSS_taux)) / 1000
        O2_exces += O2_flux_amont
        cuve_O2_besoins[i] += O2_exces + H.ratio_O2_DBO * cuve_DBO_eliminee[i] + respiration_gL(T_design, nominal_MES_concentration) * cuve_volume[i]
        if (!choix_pre_anoxie && !choix_pre_denit && !choix_anaerobie) cuve_O2_besoins[i] += besoinsO2HS(s.Sh)
        DBO_courante = DBO_sortante
      }
      for (let i = 1; i <= nb_cuves_aerees; i++) {
        nominal_O2_besoin_total += cuve_O2_besoins[i]
        nominal_air_Q += air_Q_cuve(cuve_O2_dissous[i], cuve_O2_besoins[i], T_design, cuve_volume[i], estH(i), cuve_hauteur[i], cuve_alfa[i], cuve_rdt_transfert[i])
      }
      if (DBO_courante > cuves_centrales_DBOout) dire("Traitement de la DBO insuffisant dans les cuves aérées : la consigne de sortie n'est pas atteinte.")

      volume_total_bassins = pre_anoxie_volume + anaerobie_volume + pre_denit_volume + deox_volume + post_denit_volume + reox_volume
      for (let i = 1; i <= nb_cuves_aerees; i++) volume_total_bassins += cuve_volume[i]

      // ---- CLARIFICATEUR
      clarif_vitesse_max = f('clarif_vitesse_max') ??
        (100 * clarif_hauteur * Math.pow(sortie_MES / 3.15, 0.5)) / ((1 + recirculation_taux) * boues_indice_Mohlman * nominal_MES_concentration)
      clarif_surface = f('clarif_surface') ??
        (site.Q_nominal * pointe + (Q - site.Q_nominal)) / (CONST.NOMBRE_HEURE_PAR_JOUR * clarif_vitesse_max)
      clarif_vitesse_max_recalc = (site.Q_nominal * pointe + (Q - site.Q_nominal)) / (CONST.NOMBRE_HEURE_PAR_JOUR * clarif_surface)
      clarif_charge_radier_nominal = (nominal_MES_concentration * Q * (1 + recirculation_taux)) / (CONST.NOMBRE_HEURE_PAR_JOUR * clarif_surface)
      if (clarif_charge_radier_nominal > H.charge_radier_max[choices.racleur]) {
        dire('Charge au radier du clarificateur supérieure au maximum admissible pour ce type de racleur.')
      }
      // capacité à décanter : si les boues sont trop concentrées, on recircule davantage
      const C_lim = H.facteur_boues_C_lim / boues_indice_Mohlman
      let corrige = false
      if (boues_concentration > C_lim && f('recirculation_taux') == null) { recirculation_taux += 0.1; corrige = true }

      DBO_sortie = DBO_courante

      // ---- SORTIE EAU
      const Q_out = Q - boues_Q
      const out = makeStream()
      out.Q = Q_out
      out.MES = (sortie_MES * Q_out) / 1000
      out.DBO = ((outlet_DBO_soluble + 0.6 * sortie_MES * nominal_MV_MES) * Q_out) / 1000
      const DCO_dure_mgL = ((site.DCO_nominal * H.ratio_DCO_dure_total) / site.Q_nominal) * 1000
      out.DCO = ((DCO_dure_mgL + outlet_DBO_soluble + 1.45 * sortie_MES * nominal_MV_MES) * Q_out) / 1000
      out.NH4 = (outlet_NH4 * Q_out) / 1000
      out.NK = ((outlet_NH4 + Norga_dure_soluble) * Q_out) / 1000
      out.NO3 = (outlet_NO3 * Q_out) / 1000
      out.Pt = interm_Pt
      out.Sh = 0
      outNominal = out

      const boues_MES_nom = boues_pdtes + boues_minerales_pdtes - out.MES
      const conc_prev = boues_concentration
      boues_concentration = recirculation_taux > 0
        ? ((1 + recirculation_taux) * nominal_MES_concentration - sortie_MES / 1000) / recirculation_taux
        : nominal_MES_concentration
      const Q_prev = boues_Q
      boues_Q = boues_concentration > 0 ? boues_MES_nom / boues_concentration : 0

      // convergence de la passe de dimensionnement
      const stable = !corrige
        && Math.abs(boues_Q - Q_prev) < 1e-6 * Math.max(1, Math.abs(boues_Q))
        && Math.abs(boues_concentration - conc_prev) < 1e-6 * Math.max(1, Math.abs(boues_concentration))
        && nominal_recirculation_MLSS_taux >= 0
      if (stable && passe > 1) break
      if (passe === 25) dire("Dimensionnement : stabilisation non atteinte en 25 passes (recirculation ou consigne de NO3).")
    }

    /** aiguillage vers la routine de dimensionnement des cuves centrales */
    function dimensionneCuvesCentrales(nit_cible, DBOin, N_nit_aerees, O2_besoins) {
      if (nb_cuves_CN === 0) return config_4_6(nit_cible, DBOin, O2_besoins)
      if (nb_cuves_H === 2) return config_3(nit_cible, DBOin, N_nit_aerees, O2_besoins)
      if (nb_cuves_CN === 2) return config_2(nit_cible, DBOin, N_nit_aerees, O2_besoins)
      return config_1_5(nit_cible, DBOin, N_nit_aerees, O2_besoins)
    }

    /** configurations 1 (C/N–H) et 5 (H–C/N) : une cuve H, une cuve C/N */
    function config_1_5(nit_cible, DBOin, N_nit_aerees, O2_besoins) {
      let Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
      const rate = []
      for (let i = 1; i <= nb_cuves_aerees; i++) rate[i] = cinetique_nit_MLSS(cuve_O2_dissous[i], T_design, rate_bacteries_nit)
      let DBO_appliquee = 0.5
      let eps = H.critere_convergence + 1
      let garde = 0
      while (eps > H.critere_convergence && garde++ < 300) {
        let V_total = 0
        eps = 0
        let restant = nit_cible - Nnit_media_total
        for (let i = 1; i <= nb_cuves_aerees; i++) {
          if (!estH(i)) continue
          const prev = cuve_volume[i]
          cuve_volume[i] = volume_cuve_H(i, Nnit_media_total, DBO_appliquee, T_design)
          eps += cuve_volume[i] > 0 ? Math.pow(prev - cuve_volume[i], 2) / cuve_volume[i] : 0
          V_total += cuve_volume[i]
          restant -= nitMLSS(i, rate[i])
        }
        if (restant < 0) {
          if (ratio_Nnit_media_force) { dire('La nitrification affectée au média dépasse la capacité totale : réduire la part sur média.'); break }
          ratio_Nnit_media_Nnit_total *= 0.9
          Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
          eps = H.critere_convergence + 1
        }
        for (let i = 1; i <= nb_cuves_aerees; i++) {
          if (estH(i)) continue
          const prev = cuve_volume[i]
          cuve_volume[i] = (1000 * restant) / (CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate[i])
          eps += cuve_volume[i] > 0 ? Math.pow(prev - cuve_volume[i], 2) / cuve_volume[i] : 0
          V_total += cuve_volume[i]
        }
        DBO_appliquee = V_total > 0 ? DBOin / (V_total * nominal_MES_concentration) : 0
      }
      for (let i = 1; i <= nb_cuves_aerees; i++) {
        O2_besoins[i] = ratio_O2_nit * (estH(i) ? Nnit_media_total + nitMLSS(i, rate[i]) : (rate[i] * cuve_volume[i] * CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration) / 1000)
      }
    }

    /** configuration 2 : C/N – H – C/N */
    function config_2(nit_cible, DBOin, N_nit_aerees, O2_besoins) {
      let Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
      const partCN = { 1: ratio_Nnit_CN1, 3: 1 - ratio_Nnit_CN1 }
      const rate = []
      for (let i = 1; i <= nb_cuves_aerees; i++) rate[i] = cinetique_nit_MLSS(cuve_O2_dissous[i], T_design, rate_bacteries_nit)
      let DBO_appliquee = 0.5
      let eps = H.critere_convergence + 1
      let garde = 0
      while (eps > H.critere_convergence && garde++ < 300) {
        let V_total = 0
        eps = 0
        let restant = nit_cible - Nnit_media_total
        const prev2 = cuve_volume[2]
        cuve_volume[2] = volume_cuve_H(2, Nnit_media_total, DBO_appliquee, T_design)
        eps += cuve_volume[2] > 0 ? Math.pow(prev2 - cuve_volume[2], 2) / cuve_volume[2] : 0
        V_total += cuve_volume[2]
        restant -= nitMLSS(2, rate[2])
        if (restant < 0) {
          if (ratio_Nnit_media_force) { dire('La nitrification affectée au média dépasse la capacité totale : réduire la part sur média.'); break }
          ratio_Nnit_media_Nnit_total *= 0.9
          Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
          eps = H.critere_convergence + 1
        }
        for (const i of [1, 3]) {
          const prev = cuve_volume[i]
          cuve_volume[i] = (1000 * restant * partCN[i]) / (CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate[i])
          eps += cuve_volume[i] > 0 ? Math.pow(prev - cuve_volume[i], 2) / cuve_volume[i] : 0
          V_total += cuve_volume[i]
        }
        DBO_appliquee = V_total > 0 ? DBOin / (V_total * nominal_MES_concentration) : 0
      }
      for (const i of [1, 3]) O2_besoins[i] = ratio_O2_nit * ((rate[i] * cuve_volume[i] * CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration) / 1000)
      O2_besoins[2] = ratio_O2_nit * (Nnit_media_total + nitMLSS(2, rate[2]))
    }

    /** configuration 3 : C/N – H – H */
    function config_3(nit_cible, DBOin, N_nit_aerees, O2_besoins) {
      let Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
      const Nnit_media = { 2: ratio_Nnit_mediaH1 * Nnit_media_total, 3: (1 - ratio_Nnit_mediaH1) * Nnit_media_total }
      const rate = []
      for (let i = 1; i <= nb_cuves_aerees; i++) rate[i] = cinetique_nit_MLSS(cuve_O2_dissous[i], T_design, rate_bacteries_nit)
      let DBO_appliquee = 0.5
      let eps = H.critere_convergence + 1
      let garde = 0
      while (eps > H.critere_convergence && garde++ < 300) {
        let V_total = 0
        eps = 0
        let restant = nit_cible - Nnit_media_total
        for (const i of [2, 3]) {
          const prev = cuve_volume[i]
          cuve_volume[i] = volume_cuve_H(i, Nnit_media[i], DBO_appliquee, T_design)
          eps += cuve_volume[i] > 0 ? Math.pow(prev - cuve_volume[i], 2) / cuve_volume[i] : 0
          V_total += cuve_volume[i]
          restant -= nitMLSS(i, rate[i])
        }
        if (restant < 0) {
          if (ratio_Nnit_media_force) { dire('La nitrification affectée au média dépasse la capacité totale : réduire la part sur média.'); break }
          ratio_Nnit_media_Nnit_total *= 0.9
          Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
          Nnit_media[2] = ratio_Nnit_mediaH1 * Nnit_media_total
          Nnit_media[3] = (1 - ratio_Nnit_mediaH1) * Nnit_media_total
          eps = H.critere_convergence + 1
        }
        const prev1 = cuve_volume[1]
        cuve_volume[1] = (1000 * restant) / (CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate[1])
        eps += cuve_volume[1] > 0 ? Math.pow(prev1 - cuve_volume[1], 2) / cuve_volume[1] : 0
        V_total += cuve_volume[1]
        DBO_appliquee = V_total > 0 ? DBOin / (V_total * nominal_MES_concentration) : 0
      }
      for (let i = 1; i <= nb_cuves_aerees; i++) {
        O2_besoins[i] = ratio_O2_nit * (estH(i) ? (Nnit_media[i] ?? 0) + nitMLSS(i, rate[i]) : (rate[i] * cuve_volume[i] * CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration) / 1000)
      }
    }

    /**
     * configurations 4 (H–H) et 6 (H) : dichotomie sur le volume total, réparti
     * entre les cuves au prorata inverse de leur capacité cinétique sur média.
     */
    function config_4_6(nit_cible, DBOin, O2_besoins) {
      const rate = []
      const ratio_cinetique = []
      let somme = 0
      const partMedia = { 1: ratio_Nnit_mediaH1, 2: 1 - ratio_Nnit_mediaH1 }
      for (let i = 1; i <= nb_cuves_aerees; i++) {
        rate[i] = cinetique_nit_MLSS(cuve_O2_dissous[i], T_design, rate_bacteries_nit)
        const base = cuve_O2_dissous[i] - 1
        ratio_cinetique[i] = choix_fct_Mox ? 0
          : (partMedia[i] ?? 0) * Smedia(i) * cuve_filling[i] * Math.pow(Math.max(0, base), 0.7)
        somme += ratio_cinetique[i]
      }
      const partVolume = []
      if (nb_cuves_aerees > 1 && somme > 0) {
        for (let i = 1; i <= nb_cuves_aerees; i++) {
          let autres = 0
          for (let j = 1; j <= nb_cuves_aerees; j++) if (j !== i) autres += ratio_cinetique[j]
          partVolume[i] = autres / somme
        }
      } else for (let i = 1; i <= nb_cuves_aerees; i++) partVolume[i] = 1 / nb_cuves_aerees
      let rate_moyen = 0
      for (let i = 1; i <= nb_cuves_aerees; i++) rate_moyen += partVolume[i] * rate[i]
      // Borne haute : volume qu'il faudrait sans média. Le VBA multiplie là où
      // il faut diviser ; on retient la formule correcte, avec marge.
      let Vmin = 1
      let Vmax = rate_moyen > 0
        ? (5 * 1000 * nit_cible) / (CONST.NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate_moyen)
        : 1e6
      let V = (Vmin + Vmax) / 2
      let garde = 0
      const O2_interm = new Array(nb_cuves_aerees + 1).fill(0)
      while (V > 0 && (Vmax - Vmin) / V > H.critere_convergence && garde++ < 300) {
        for (let i = 1; i <= nb_cuves_aerees; i++) cuve_volume[i] = partVolume[i] * V
        let V_media = 0
        for (let i = 1; i <= nb_cuves_aerees; i++) V_media += cuve_filling[i] * cuve_volume[i] * deplacement(i)
        const V_liqueur = V - V_media + deox_volume + reox_volume
        const DBO_appliquee = V_liqueur > 0 ? DBOin / (nominal_MES_concentration * V_liqueur) : 0
        let nit_total = 0
        ratio_Nnit_media_Nnit_total = 0
        for (let i = 1; i <= nb_cuves_aerees; i++) {
          const n = nitrification_H_kgj(i, DBO_appliquee, T_design, rate[i], nominal_MES_concentration)
          nit_total += n.total
          ratio_Nnit_media_Nnit_total += n.surMedia
          O2_interm[i] = ratio_O2_nit * n.total
        }
        if (nit_total < nit_cible) Vmin = V
        else Vmax = V
        V = (Vmin + Vmax) / 2
      }
      for (let i = 1; i <= nb_cuves_aerees; i++) O2_besoins[i] = O2_interm[i]
      // le VBA divise ici le cumul de nitrification sur média par N_nit_aerees
      ratio_Nnit_media_Nnit_total = nit_cible > 0 ? ratio_Nnit_media_Nnit_total / nit_cible : 0
    }

    // =======================================================================
    // fonctionnement_reel
    // =======================================================================
    const r = ctx.inReel
    const Qr = r.Q
    const Qn = site.Q_nominal * site.NC_Q
    const Norga_dure_soluble_reel = Qn > 0 ? H.ratio_Norgaduresoluble_DCOtot * ((site.DCO_nominal * site.NC_DCO) / Qn) * 1000 : 1.5
    const DCO_dure_mgL_reel = Qn > 0 ? ((site.DCO_nominal * site.NC_DCO * H.ratio_DCO_dure_total) / Qn) * 1000 : 0
    const reel_MV_MES = f('reel_MV_MES') ?? (primaire ? H.ratio_ED_MV_MES
      : Math.min(H.ratio_EB_MV_MES_max, Math.max(H.ratio_EB_MV_MES_min,
        H.ratio_EB_MV_MES[0] + H.ratio_EB_MV_MES[1] * (r.MES / r.DCO) + H.ratio_EB_MV_MES[2] * (r.DCO / r.DBO))))
    const reel_recirculation_MLSS_taux = f('reel_recirculation_MLSS_taux') ?? nominal_recirculation_MLSS_taux

    let reel_MES_concentration = nominal_MES_concentration
    let reel_air_Q = 0
    let reel_O2_besoin_total = 0
    let outlet_reel_NH4 = outlet_NH4
    let outlet_reel_NO3 = outlet_NO3
    let outlet_reel_DBO = outlet_DBO_soluble
    let outlet_reel_NK = outlet_NH4 + Norga_dure_soluble_reel
    let outlet_reel_Pt = 0
    let reel_boues_pdtes = 0
    let reel_boues_minerales = 0
    let reel_FeCl3 = 0
    let reel_methanol = 0
    let reel_deox_nit = 0
    let reel_reox_nit = 0
    let NO3_last_aerated_tank_reel = NO3_last_aerated_tank
    let NH4_last_aerated_tank = outlet_NH4
    let deox_O2_dissous_reel = deox_O2_dissous
    let DBO_sortie_reel = DBO_sortie
    let boues_methanol_reel = boues_methanol
    let sulfuresIgnores = false

    fonctionnementReel()

    function fonctionnementReel() {
      let N_nit_total = r.NK - (Norga_dure_soluble_reel * Qr) / 1000 - (outlet_NH4 * Qr) / 1000
        - H.ratio_N_assimile * (r.DBO - (outlet_DBO_soluble * Qr) / 1000)
        - H.ratio_N_assimile_methanol * boues_methanol_reel
      let N_nit_aerees = N_nit_total - reox_nit
      reel_boues_pdtes = ratio_boues_MESin * r.MES + H.ratio_boues_DBOin * r.DBO + H.ratio_boues_nit * N_nit_total + boues_methanol_reel

      // amorce des flux recirculés depuis le clarificateur
      let rec_DCO = ((recirculation_taux * Qr) / 1000) * (outlet_reel_DBO + DCO_dure_mgL_reel)
      let rec_DBO = ((recirculation_taux * Qr) / 1000) * outlet_reel_DBO
      let rec_NK = (outlet_reel_NK * recirculation_taux * Qr) / 1000
      let rec_NH4 = (outlet_reel_NH4 * recirculation_taux * Qr) / 1000
      let rec_NO3 = (outlet_reel_NO3 * recirculation_taux * Qr) / 1000
      let rec_Pt = (site.Pt_garantie * recirculation_taux * Qr) / 1000

      let interm_Q = 0, interm_DCO = 0, interm_DBO = 0, interm_NK = 0, interm_NH4 = 0, interm_NO3 = 0, interm_Pt = 0
      let epsBoues = H.critere_convergence + 1
      let gardeBoues = 0
      while (epsBoues > H.critere_convergence && gardeBoues++ < 60) {
        reel_air_Q = 0
        reel_O2_besoin_total = 0
        const eps0_NH4 = outlet_reel_NH4
        const eps0_NO3 = outlet_reel_NO3
        const eps0_DBO = outlet_reel_DBO
        reel_MES_concentration = f('reel_MES_concentration') ??
          (boues_pdtes_nominal > 0 ? (reel_boues_pdtes / boues_pdtes_nominal) * nominal_MES_concentration : nominal_MES_concentration)

        interm_Q = Qr + recirculation_taux * Qr
        const k = recirculation_taux / (1 + recirculation_taux)
        interm_DCO = r.DCO + k * rec_DCO
        interm_DBO = r.DBO + k * rec_DBO
        interm_NK = r.NK + k * rec_NK
        interm_NH4 = r.NK - (Norga_dure_soluble_reel * Qr) / 1000 + k * rec_NH4
        interm_NO3 = r.NO3 + k * rec_NO3
        interm_Pt = r.Pt + k * rec_Pt
        let O2_dissous_flux = (H.recirculation_O2 * recirculation_taux * Qr) / 1000

        // ---- PRE-ANOXIE
        if (choix_pre_anoxie) {
          let denit = denitVolumeImpose(pre_anoxie_volume, interm_NO3 + H.ratio_NO3eq_O2 * O2_dissous_flux, r.DBO, reel_MES_concentration, reel_MV_MES, T_reel)
          interm_NO3 += H.ratio_NO3eq_O2 * O2_dissous_flux
          if (interm_NO3 < denit) { denit = interm_NO3; interm_NO3 = 0 } else interm_NO3 -= denit
          interm_DBO += H.ratio_O2_denit * denit
          interm_DCO += H.ratio_O2_denit * denit
          interm_NK += H.ratio_N_assimile * H.ratio_O2_denit * denit
          interm_NH4 += H.ratio_N_assimile * H.ratio_O2_denit * denit
          O2_dissous_flux = 0
        }

        // ---- ANAEROBIE
        let deltaP_bio = 0
        if (choix_anaerobie) {
          const Pt_mgL = (r.Pt / Qr) * 1000
          const V_opt = ((Pt_mgL < H.anaerobie_P_limite ? H.anaerobie_HRT_P_low : H.anaerobie_HRT_P_high) * Qr) / CONST.NOMBRE_HEURE_PAR_JOUR
          deltaP_bio = interm_Pt * anaerobie_Kp(anaerobie_volume / V_opt, (interm_NO3 / interm_Q) * 1000, interm_Pt, r.DBO)
          interm_DBO -= H.anaerobie_DBO_Pt * deltaP_bio
          interm_DCO -= H.anaerobie_DBO_Pt * deltaP_bio
          interm_NH4 -= H.ratio_N_assimile * H.anaerobie_DBO_Pt * deltaP_bio
          interm_NK -= H.ratio_N_assimile * H.anaerobie_DBO_Pt * deltaP_bio
          const deltaP_synthese = H.ratio_P_assimile * (r.DBO - DBO_sortie_reel)
          let deltaP_prec = interm_Pt - (deltaP_bio + deltaP_synthese + (site.Pt_garantie * interm_Q) / 1000)
          if (deltaP_prec > 0) {
            const rm = ratio_molaire_precipitation(deltaP_prec, Qr)
            reel_FeCl3 = f('anaerobie_FeCl3_flux') ?? (rm * H.MM_FeCl3 * deltaP_prec) / H.MM_P
            reel_boues_minerales = (((rm - 1) * H.MM_FeO3H3) / H.MM_P + H.MM_FePO4 / H.MM_P) * deltaP_prec
          } else { deltaP_prec = 0; reel_FeCl3 = f('anaerobie_FeCl3_flux') ?? 0; reel_boues_minerales = 0 }
          interm_Pt -= deltaP_synthese + deltaP_prec + deltaP_bio
          O2_dissous_flux = 0
        } else {
          interm_Pt -= H.ratio_P_assimile * (r.DBO - DBO_sortie_reel)
        }

        const st = { Q: interm_Q, DCO: interm_DCO, DBO: interm_DBO, NK: interm_NK, NH4: interm_NH4, NO3: interm_NO3, Pt: interm_Pt, O2: O2_dissous_flux }

        // ---- boucle sur la recirculation de liqueur mixte
        let epsMLSS = H.critere_convergence + 1
        let iterMLSS = 0
        let air_MLSS = 0
        let O2_MLSS = 0
        let nit_aerees_courant = 0
        while (epsMLSS > H.critere_convergence && iterMLSS++ < 60) {
          air_MLSS = 0
          O2_MLSS = 0
          const eps0_NO3_MLSS = NO3_last_aerated_tank_reel
          const eps0_NH4_MLSS = NH4_last_aerated_tank
          const ratio_MLSS = reel_recirculation_MLSS_taux / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
          if (iterMLSS > 1) {
            interm_Q = Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
            interm_DCO = st.DCO + ratio_MLSS * interm_DCO
            interm_DBO = st.DBO + ratio_MLSS * interm_DBO
            interm_NK = st.NK + ratio_MLSS * interm_NK
            interm_NH4 = st.NH4 + ratio_MLSS * interm_NH4
            interm_NO3 = st.NO3 + ratio_MLSS * interm_NO3
            interm_Pt = st.Pt + ratio_MLSS * interm_Pt
          } else {
            interm_Q = Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
            interm_DCO = st.DCO + ((reel_recirculation_MLSS_taux * Qr) / 1000) * (outlet_reel_DBO + DCO_dure_mgL_reel)
            interm_DBO = st.DBO + ((reel_recirculation_MLSS_taux * Qr) / 1000) * outlet_reel_DBO
            interm_NK = st.NK + (outlet_reel_NK * reel_recirculation_MLSS_taux * Qr) / 1000
            interm_NH4 = st.NH4 + (outlet_reel_NH4 * reel_recirculation_MLSS_taux * Qr) / 1000
            interm_NO3 = st.NO3 + (NO3_last_aerated_tank_reel * reel_recirculation_MLSS_taux * Qr) / 1000
            interm_Pt = st.Pt + (site.Pt_garantie * reel_recirculation_MLSS_taux * Qr) / 1000
          }
          const O2_ret = choix_deox ? deox_O2_dissous_reel : cuve_O2_dissous[nb_cuves_aerees]
          let O2_flux = st.O2 + (O2_ret * reel_recirculation_MLSS_taux * Qr) / 1000

          // ---- PRE-DENIT
          if (choix_pre_denit) {
            let denit = denitVolumeImpose(pre_denit_volume, interm_NO3 + H.ratio_NO3eq_O2 * O2_flux, r.DBO, reel_MES_concentration, reel_MV_MES, T_reel)
            interm_NO3 += H.ratio_NO3eq_O2 * O2_flux
            if (interm_NO3 < denit) { denit = interm_NO3; interm_NO3 = 0 } else interm_NO3 -= denit
            interm_DBO += H.ratio_O2_denit * denit
            interm_DCO += H.ratio_O2_denit * denit
            interm_NK += H.ratio_N_assimile * H.ratio_O2_denit * denit
            interm_NH4 += H.ratio_N_assimile * H.ratio_O2_denit * denit
            O2_flux = 0
          }

          // ---- CUVES CENTRALES
          let V_central = 0
          for (let i = 1; i <= nb_cuves_aerees; i++) V_central += cuve_volume[i]
          const DBO_appliquee = V_central > 0 ? interm_DBO / (V_central * reel_MES_concentration) : 0
          let DBOin = interm_DBO
          const DBOout_min = ((outlet_DBO_soluble * Qr) / 1000) * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
          let nit_cumul = 0
          for (let i = 1; i <= nb_cuves_aerees; i++) {
            let O2_exces = -O2_flux
            const rate_N_MLSS = cinetique_nit_MLSS(cuve_O2_dissous[i], T_reel, rate_bacteries_nit)
            let nit = estH(i)
              ? nitrification_H_kgj(i, DBO_appliquee, T_reel, rate_N_MLSS, reel_MES_concentration).total
              : (CONST.NOMBRE_HEURE_PAR_JOUR / 1000) * rate_N_MLSS * reel_MES_concentration * cuve_volume[i]
            const V_liq = estH(i) ? cuve_volume[i] * (1 - cuve_filling[i] * deplacement(i)) : cuve_volume[i]
            const Cm = V_liq > 0 ? DBOin / (V_liq * reel_MES_concentration) : 0
            let DBOout = DBOin * (1 - rdt_DBO(Cm))
            if (DBOout < DBOout_min) DBOout = DBOout_min
            const elim = DBOin - DBOout
            DBOin = DBOout
            interm_DCO -= elim
            interm_NH4 -= H.ratio_N_assimile * elim
            interm_NK -= H.ratio_N_assimile * elim
            if (interm_NH4 < nit) { nit = interm_NH4; interm_NH4 = 0 } else interm_NH4 -= nit
            interm_NK -= nit
            interm_NO3 += nit * ratio_NO3f_nit
            nit_cumul += nit
            let besoin = nit * ratio_O2_nit + elim * H.ratio_O2_DBO + cuve_volume[i] * respiration_gL(T_reel, reel_MES_concentration)
            O2_flux = (cuve_O2_dissous[i] * Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux)) / 1000
            O2_exces += O2_flux
            besoin += O2_exces
            // Le VBA ajoute ici besoins_O2_HS à un index de cuve non initialisé :
            // la contribution des sulfures est perdue. Reproduit tel quel.
            if (!choix_pre_anoxie && !choix_pre_denit && !choix_anaerobie && r.Sh > 0) sulfuresIgnores = true
            O2_MLSS += besoin
            air_MLSS += air_Q_cuve(cuve_O2_dissous[i], besoin, T_reel, cuve_volume[i], estH(i), cuve_hauteur[i], cuve_alfa[i], cuve_rdt_transfert[i])
          }
          interm_DBO = DBOin
          O2_flux = (cuve_O2_dissous[nb_cuves_aerees] / 1000) * Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
          nit_aerees_courant = nit_cumul

          // ---- DE-OX
          if (choix_deox) {
            const resp = respiration_gL(T_reel, reel_MES_concentration) * deox_volume
            let min = 0
            let max = (cuve_O2_dissous[nb_cuves_aerees] * interm_Q) / 1000 - resp
            if (max > 0) {
              let d = (min + max) / 2
              let garde = 0
              deox_O2_dissous_reel = cuve_O2_dissous[nb_cuves_aerees] - (resp + d) / (interm_Q / 1000)
              while (d > 0 && (max - min) / d > H.critere_convergence && garde++ < 300) {
                reel_deox_nit = (CONST.NOMBRE_HEURE_PAR_JOUR / 1000) * reel_MES_concentration * cinetique_nit_MLSS(deox_O2_dissous_reel, T_reel, rate_bacteries_nit) * deox_volume
                if (reel_deox_nit * ratio_O2_nit < d) max = d
                else min = d
                d = (min + max) / 2
                deox_O2_dissous_reel = cuve_O2_dissous[nb_cuves_aerees] - (resp + d) / (interm_Q / 1000)
              }
            }
            // Écart au VBA : la nitrification en dé-ox est plafonnée au NH4
            // disponible, faute de quoi l'ammonium de sortie peut devenir négatif.
            if (reel_deox_nit > interm_NH4) reel_deox_nit = Math.max(0, interm_NH4)
            interm_NK -= reel_deox_nit
            interm_NH4 -= reel_deox_nit
            interm_NO3 += reel_deox_nit
            nit_aerees_courant += reel_deox_nit
            O2_flux = (deox_O2_dissous_reel / 1000) * Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
          }
          st.O2sortie = O2_flux

          NO3_last_aerated_tank_reel = (interm_NO3 / (Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux))) * 1000
          NH4_last_aerated_tank = (interm_NH4 / (Qr * (1 + recirculation_taux + reel_recirculation_MLSS_taux))) * 1000
          N_nit_aerees = nit_aerees_courant
          N_nit_total = N_nit_aerees + reel_reox_nit
          reel_boues_pdtes = ratio_boues_MESin * r.MES + H.ratio_boues_DBOin * r.DBO + H.ratio_boues_nit * N_nit_total + boues_methanol_reel
          epsMLSS = Math.pow(eps0_NO3_MLSS - NO3_last_aerated_tank_reel, 2) + Math.pow(eps0_NH4_MLSS - NH4_last_aerated_tank, 2)
        }
        if (iterMLSS >= 60) dire("Fonctionnement réel : la boucle de recirculation de liqueur mixte n'a pas convergé.")
        reel_air_Q += air_MLSS
        reel_O2_besoin_total += O2_MLSS

        // retour au débit hors recirculation de liqueur mixte
        const kk = (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_Q *= kk; interm_DCO *= kk; interm_DBO *= kk; interm_NK *= kk
        interm_NH4 *= kk; interm_NO3 *= kk; interm_Pt *= kk
        let O2_flux = (st.O2sortie ?? 0) * kk

        // ---- POST-DENIT
        if (postDN) {
          interm_NO3 += H.ratio_NO3eq_O2 * O2_flux
          let denit
          if (choix_post_denit === 'sans_media') {
            const cin = cinetique_post_DN_15 * a_cinetique_postDN_reel
            denit = (post_denit_volume * CONST.NOMBRE_HEURE_PAR_JOUR * cin * reel_MES_concentration * reel_MV_MES) / 1000
          } else {
            let NO3_out = outlet_NO3
            let eps = 1
            let garde = 0
            while (eps > H.critere_convergence && garde++ < 300) {
              const prev = NO3_out
              const cin = ((H.a_cinetique_postDN_media_10[0] * NO3_out) / (NO3_out + H.a_cinetique_postDN_media_10[1])) * a_cinetique_postDN_media_reel
              NO3_out = interm_NO3 - (post_denit_volume * post_denit_media_filling * MEDIA[media_PDN].surface * cin) / 1000
              NO3_out = (NO3_out / interm_Q) * 1000
              eps = Math.pow(prev - NO3_out, 2)
            }
            denit = interm_NO3 - (NO3_out * interm_Q) / 1000
          }
          if (denit > interm_NO3) denit = interm_NO3
          if (denit < 0) denit = 0
          const DCO_apportee = rate_DCO_apportee_denit_reel * denit
          reel_methanol = f('post_denit_methanol_flux') ?? DCO_apportee / H.rate_DCO_methanol
          boues_methanol_reel = H.ratio_boues_DCO_methanol * DCO_apportee
          interm_NO3 -= denit
          const Nassim = H.ratio_N_assimile_methanol * boues_methanol_reel
          if (Nassim > interm_NH4) {
            interm_NO3 -= Nassim - interm_NH4
            interm_NK -= interm_NH4
            interm_NH4 = 0
            // Écart au VBA : l'azote assimilé peut excéder NH4 + NO3 disponibles
            // et rendre le nitrate négatif. On plafonne à zéro en le signalant.
            if (interm_NO3 < 0) {
              interm_NO3 = 0
              dire("Post-dénitrification : l'azote assimilé par la biomasse sur méthanol dépasse l'azote disponible ; le nitrate de sortie est plafonné à zéro.")
            }
          } else { interm_NH4 -= Nassim; interm_NK -= Nassim }
          O2_flux = 0
        }

        // ---- RE-OX
        if (choix_reox) {
          let O2_exces = -O2_flux
          reel_reox_nit = (cinetique_nit_MLSS(reox_O2_dissous, T_reel, rate_bacteries_nit) * reox_volume * CONST.NOMBRE_HEURE_PAR_JOUR * reel_MES_concentration) / 1000
          const max_nit = (((1 + recirculation_taux) * Qr * reox_O2_dissous) / 1000 - O2_flux) / ratio_O2_nit
          if (reel_reox_nit > max_nit) reel_reox_nit = max_nit
          if (reel_reox_nit > interm_NH4) reel_reox_nit = interm_NH4
          interm_NO3 += reel_reox_nit
          interm_NH4 -= reel_reox_nit
          interm_NK -= reel_reox_nit
          O2_flux = (reox_O2_dissous * Qr * (1 + recirculation_taux)) / 1000
          O2_exces += O2_flux
          const besoin = ratio_O2_nit * reel_reox_nit + O2_exces
          reel_O2_besoin_total += besoin
          reel_air_Q += air_Q_cuve(reox_O2_dissous, besoin, T_reel, reox_volume, false, reox_hauteur_bassin, reox_O2_alpha, reox_O2_rdt_transfert)
        }

        const Qtot = Qr * (1 + recirculation_taux)
        if (interm_NH4 < 0) {
          // l'assimilation d'azote par la biomasse peut dépasser le NH4 restant
          interm_NK -= interm_NH4
          interm_NH4 = 0
          dire("L'azote assimilé dépasse l'ammonium disponible : le NH4 de sortie est plafonné à zéro.")
        }
        if (interm_NK < 0) interm_NK = 0
        outlet_reel_NH4 = (interm_NH4 / Qtot) * 1000
        outlet_reel_DBO = (interm_DBO / Qtot) * 1000
        outlet_reel_NK = (interm_NK / Qtot) * 1000
        outlet_reel_NO3 = (interm_NO3 / Qtot) * 1000
        outlet_reel_Pt = (interm_Pt / Qtot) * 1000
        N_nit_total = N_nit_aerees + reel_reox_nit
        reel_boues_pdtes = ratio_boues_MESin * r.MES + H.ratio_boues_DBOin * r.DBO + H.ratio_boues_nit * N_nit_total + boues_methanol_reel
        rec_DCO = interm_DCO; rec_DBO = interm_DBO; rec_NK = interm_NK
        rec_NH4 = interm_NH4; rec_NO3 = interm_NO3; rec_Pt = interm_Pt

        epsBoues = Math.pow(eps0_NH4 - outlet_reel_NH4, 2) + Math.pow(eps0_NO3 - outlet_reel_NO3, 2) + Math.pow(eps0_DBO - outlet_reel_DBO, 2)
      }
      if (gardeBoues >= 60) dire("Fonctionnement réel : la boucle sur la production de boues n'a pas convergé.")
      DBO_sortie_reel = outlet_reel_DBO
    }
    if (sulfuresIgnores) {
      dire("Les besoins en O2 dus aux sulfures ne sont pas comptés au fonctionnement réel (indice de cuve erroné dans le VBA d'origine).")
    }

    // ---- boues extraites et clarificateur au réel
    const reel_boues_MES = f('boues_MES') ?? (reel_boues_pdtes + reel_boues_minerales - (sortie_MES * Qr) / 1000)
    const reel_boues_concentration = f('boues_concentration') ?? (recirculation_taux > 0
      ? ((1 + recirculation_taux) * reel_MES_concentration - sortie_MES / 1000) / recirculation_taux
      : reel_MES_concentration)
    const reel_boues_Q = reel_boues_concentration > 0 ? reel_boues_MES / reel_boues_concentration : 0
    if (reel_boues_concentration > H.facteur_boues_C_lim / boues_indice_Mohlman) {
      dire('Concentration des boues extraites supérieure à la limite de bonne décantation : augmenter le taux de recirculation.')
    }

    // ---- forçages sur les sorties réelles
    const Q_out_reel = Qr - reel_boues_Q
    const NH4f = f('outlet_reel_NH4')
    const O2f = f('O2_besoin_total')
    const airf = f('air_Q_Nm3j')
    if (NH4f != null) {
      const delta = (ratio_O2_nit * (outlet_reel_NH4 - NH4f) * Qr) / 1000
      outlet_reel_NH4 = NH4f
      if (O2f != null) {
        if (airf != null) { reel_O2_besoin_total = O2f; reel_air_Q = airf }
        else { reel_air_Q = reel_O2_besoin_total > 0 ? (O2f / reel_O2_besoin_total) * reel_air_Q : 0; reel_O2_besoin_total = O2f }
      } else if (airf != null) { reel_air_Q = airf; reel_O2_besoin_total += delta }
      else {
        reel_air_Q = reel_O2_besoin_total > 0 ? ((reel_O2_besoin_total + delta) / reel_O2_besoin_total) * reel_air_Q : 0
        reel_O2_besoin_total += delta
      }
    } else if (O2f != null) {
      reel_air_Q = airf != null ? airf : (reel_O2_besoin_total > 0 ? (O2f / reel_O2_besoin_total) * reel_air_Q : 0)
      reel_O2_besoin_total = O2f
    } else if (airf != null) reel_air_Q = airf
    const NO3f = f('outlet_reel_NO3')
    if (NO3f != null) outlet_reel_NO3 = NO3f

    const outReel = makeStream()
    outReel.Q = Q_out_reel
    outReel.MES = (sortie_MES * Q_out_reel) / 1000
    outReel.DCO = ((outlet_reel_DBO + DCO_dure_mgL_reel + 1.45 * sortie_MES * reel_MV_MES) * Q_out_reel) / 1000
    const DBOf = f('outlet_reel_DBO')
    if (DBOf != null) { outlet_reel_DBO = DBOf; outReel.DBO = (DBOf * Q_out_reel) / 1000 }
    else {
      outReel.DBO = ((outlet_reel_DBO + 0.6 * sortie_MES * reel_MV_MES) * Q_out_reel) / 1000
      outlet_reel_DBO = Q_out_reel > 0 ? (outReel.DBO / Q_out_reel) * 1000 : 0
    }
    outReel.NK = (outlet_reel_NK * Q_out_reel) / 1000
    outReel.NH4 = (outlet_reel_NH4 * Q_out_reel) / 1000
    outReel.NO3 = (outlet_reel_NO3 * Q_out_reel) / 1000
    outReel.Pt = (outlet_reel_Pt * Q_out_reel) / 1000
    outReel.Sh = 0

    // =======================================================================
    // calcul_consommation_electrique
    // =======================================================================
    if (choices.surpresseur === 'roots' && air_P_refoulement_moyenne > HYP.surpresseur_Px2) {
      dire(`Pression de refoulement (${air_P_refoulement_moyenne.toFixed(1)} mCE) élevée pour des surpresseurs roots.`)
    }
    const surpresseur_conso_spec = f('surpresseur_conso_spec') ?? HYP.surpresseur_conso_spec_Wh_Nm3mCE[choices.surpresseur]
    const electricite_aeration = (reel_air_Q * air_P_refoulement_moyenne * surpresseur_conso_spec) / 1000
    const electricite_racleur = (clarif_surface < Math.PI * Math.pow(H.clarif_diametre_limite / 2, 2) ? 0.55 : 0.75) * CONST.NOMBRE_HEURE_PAR_JOUR
    let electricite_agitation = (ratio_elec_anaerobie * anaerobie_volume + ratio_elec_pre_anoxie * pre_anoxie_volume
      + ratio_elec_pre_denit * pre_denit_volume + ratio_elec_deox * deox_volume
      + ratio_elec_post_denit * post_denit_volume) * H.agitation_fct / 1000
    if (choix_fct_Mox) electricite_agitation += (3 * cuve_volume[1] * H.agitation_fct) / 1000
    const electricite_recirculation_MLSS = ratio_elec_recirculation_MLSS * reel_recirculation_MLSS_taux * Qr * recirculation_MLSS_P_refoulement
    const electricite_recirculation = ratio_elec_recirculation * recirculation_taux * Qr * recirculation_P_refoulement
    const electricite_extraction = ratio_elec_extraction * reel_boues_Q * extraction_P_refoulement
    const total = electricite_aeration + electricite_racleur + electricite_agitation
      + electricite_recirculation_MLSS + electricite_recirculation + electricite_extraction
    const fixed = electricite_agitation + electricite_racleur
      + (reel_O2_besoin_total > 0 ? (respiration_gL(T_reel, reel_MES_concentration) * volume_total_bassins / reel_O2_besoin_total) * electricite_aeration : 0)

    // =======================================================================
    const resultsCuves = []
    for (let i = 1; i <= nb_cuves_aerees; i++) {
      resultsCuves.push({ key: `V${i}`, label: `Volume cuve aérée ${i} (${types[i - 1]})`, unit: 'm³', value: cuve_volume[i] })
      if (estH(i)) resultsCuves.push({ key: `fill${i}`, label: `Taux de remplissage cuve ${i}`, unit: '-', value: cuve_filling[i] })
    }

    return {
      outNominal,
      outReel,
      sludge: {
        origine: 'II_HybAS',
        Q: reel_boues_Q,
        MES: reel_boues_MES,
        concentration: reel_boues_concentration,
        MV_MES: reel_MV_MES,
        NK: ratio('II_HybAS', 'NK_MV') * reel_boues_MES * reel_MV_MES,
        Pt: ratio('II_HybAS', 'Pt_MES') * reel_boues_MES,
        DCO: ratio('II_HybAS', 'DCO_MV') * reel_boues_MES * reel_MV_MES,
        DBO: ratio('II_HybAS', 'DBO_MV') * reel_boues_MES * reel_MV_MES,
      },
      results: [
        { key: 'V_total', label: 'Volume total des bassins', unit: 'm³', value: volume_total_bassins },
        ...(choix_pre_anoxie ? [{ key: 'V_preanox', label: 'Volume de pré-anoxie', unit: 'm³', value: pre_anoxie_volume }] : []),
        ...(choix_anaerobie ? [
          { key: 'V_anaer', label: 'Volume anaérobie', unit: 'm³', value: anaerobie_volume },
          { key: 'FeCl3', label: 'FeCl3 pur (réel)', unit: 'kg/j', value: reel_FeCl3 },
        ] : []),
        ...(choix_pre_denit ? [{ key: 'V_predn', label: 'Volume de pré-dénitrification', unit: 'm³', value: pre_denit_volume }] : []),
        ...resultsCuves,
        ...(choix_deox ? [
          { key: 'V_deox', label: 'Volume de dé-ox', unit: 'm³', value: deox_volume },
          { key: 'deox_nit', label: 'Nitrification en dé-ox (réel)', unit: 'kg N/j', value: reel_deox_nit },
        ] : []),
        ...(postDN ? [
          { key: 'V_pdn', label: 'Volume de post-dénitrification', unit: 'm³', value: post_denit_volume },
          { key: 'fill_pdn', label: 'Taux de remplissage post-dénit', unit: '-', value: post_denit_media_filling },
          { key: 'MeOH_nom', label: 'Méthanol (nominal)', unit: 'kg/j', value: post_denit_methanol_flux },
          { key: 'MeOH', label: 'Méthanol (réel)', unit: 'kg/j', value: reel_methanol },
        ] : []),
        ...(choix_reox ? [
          { key: 'V_reox', label: 'Volume de ré-ox', unit: 'm³', value: reox_volume },
          { key: 'reox_nit', label: 'Nitrification en ré-ox (réel)', unit: 'kg N/j', value: reel_reox_nit },
        ] : []),
        { key: 'part_media', label: 'Part de N nitrifié sur le média', unit: '-', value: ratio_Nnit_media_Nnit_total },
        { key: 'rec_MLSS_nom', label: 'Recirculation de liqueur mixte (nominal)', unit: '-', value: nominal_recirculation_MLSS_taux },
        { key: 'rec_boues', label: 'Recirculation des boues du clarificateur', unit: '-', value: recirculation_taux },
        { key: 'MES_bassin', label: 'MES dans les bassins (réel)', unit: 'g/L', value: reel_MES_concentration },
        { key: 'O2_nom', label: 'Besoin en O2 (nominal)', unit: 'kg O2/j', value: nominal_O2_besoin_total },
        { key: 'O2_reel', label: 'Besoin en O2 (réel)', unit: 'kg O2/j', value: reel_O2_besoin_total },
        { key: 'air_nom', label: "Débit d'air process (nominal)", unit: 'Nm³/h', value: nominal_air_Q / 24 },
        { key: 'air', label: "Débit d'air process (réel)", unit: 'Nm³/h', value: reel_air_Q / 24 },
        { key: 'P_air', label: 'Pression de refoulement retenue', unit: 'mCE', value: air_P_refoulement_moyenne },
        { key: 'clarif_S', label: 'Surface du clarificateur', unit: 'm²', value: clarif_surface },
        { key: 'clarif_v', label: 'Vitesse hydraulique maximale recalculée', unit: 'm/h', value: clarif_vitesse_max_recalc },
        { key: 'clarif_radier', label: 'Charge au radier (nominal)', unit: 'kg/(m²·h)', value: clarif_charge_radier_nominal },
        { key: 'DBO_out', label: 'DBO5 en sortie (réel)', unit: 'mg/L', value: outlet_reel_DBO },
        { key: 'NH4_out', label: 'N-NH4 en sortie (réel)', unit: 'mg/L', value: outlet_reel_NH4 },
        { key: 'NO3_out', label: 'N-NO3 en sortie (réel)', unit: 'mg/L', value: outlet_reel_NO3 },
        { key: 'Pt_out', label: 'Pt en sortie (réel)', unit: 'mg/L', value: outlet_reel_Pt },
        { key: 'boues', label: 'Boues à extraire (réel)', unit: 'kg MES/j', value: reel_boues_MES },
        { key: 'boues_conc', label: 'Concentration des boues extraites', unit: 'g/L', value: reel_boues_concentration },
        { key: 'boues_Q', label: 'Débit de boues extraites (réel)', unit: 'm³/j', value: reel_boues_Q },
      ],
      electricity: {
        total,
        fixed,
        detail: {
          aeration: electricite_aeration,
          agitation: electricite_agitation,
          racleur: electricite_racleur,
          recirculation_boues: electricite_recirculation,
          recirculation_MLSS: electricite_recirculation_MLSS,
          extraction: electricite_extraction,
        },
      },
      warnings,
    }
  },
})
