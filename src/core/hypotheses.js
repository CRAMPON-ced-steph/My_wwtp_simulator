// ---------------------------------------------------------------------------
// Hypothèses globales — extraites de l'onglet "Hypothèses" du classeur VBA
// et des constantes pd_* de MOD_ProgrammePrincipal.
// Le nom de la cellule d'origine est indiqué en commentaire.
// ---------------------------------------------------------------------------
export const CONST = {
  NOMBRE_HEURE_PAR_JOUR: 24,
  NOMBRE_MINUTE_PAR_HEURE: 60,
  NOMBRE_SECONDE_PAR_HEURE: 3600,
  CONVERSION_kJ_PAR_kcal: 4.18,
  QUANTITE_OXYGENE_DANS_AIR_kgO2_Nm3: 0.3,
  NOMBRE_JOUR_PAR_AN: 365,
  NOMBRE_JOUR_PAR_SEMAINE: 7,
  ACCELERATION_PESANTEUR_m_s2: 9.81,
  PI: Math.PI,
  critere_incoherence: 0.01,
}

export const HYP = {
  // --- surpresseurs d'air (hyp_surpresseur_*)
  surpresseur_Px1_mCE: 6,
  surpresseur_Px2_mCE: 8,
  surpresseur_conso_spec_Wh_Nm3mCE: { roots: 4.5, vis: 3.8, turbo: 3.5 }, // valeurs codées dans E1..E4 (feuille: roots=5)
  surpresseur_roots_Pmax_conseillee: 8, // mCE
  // --- diffusion O2 / facteur K (hyp_insufflation_*, hyp_calcul_*)
  insufflation_hauteur_diffuseur_m: 0.25,
  insufflation_deltaP_a0: 0.35,
  insufflation_deltaP_a1: 10.33,
  Patm_P0: 101325,
  Patm_a0: 1,
  Patm_a1: -0.0000225577,
  Patm_a2: 5.25588,
  O2sat_coef: [14.628, -0.4223, 0.0102, -0.0002, 0.000001],
  facteurK_beta: 0.95,
  facteurK_correction_T: 1.024,
  facteurK_Tref: 20,
  // --- paramètres biologiques
  assimilation_N_kgN_kgDBO: 0.05,
  assimilation_P_kgP_kgDBO: 0.01,
  ratio_O2_nit: 4.57,
  ratio_O2_denit: -2.86,
  ratio_kgO2_Nm3air: 0.3,
  masse_volumique_effluent: 1000,
  // --- dessablage-déshuilage
  dessabl_graisse_DBO_DCO: 0.9,
  // --- primaire
  dec_I_boues_MV_MES: 0.68,
  dec_I_diametre_limite_m: 32,
  dec_I_Pw_racleur_inf_Dlim_kW: 0.55,
  dec_I_Pw_racleur_sup_Dlim_kW: 0.75,
  dec_I_ratio_Psol_P: 0.7,
  dec_I_ratio_PO4_Psol: 0.8,
  // --- BA (tous types) (hyp_II_BA_*)
  BA_charge_radier_max: { racle: 5, racle_suce: 7.5, kruger: 9 }, // kg/(m²·h)
  BA_rdt_dissolution_O2_eau_claire: { fines_bulles_Hsup: 5.8, moyennes_bulles: 2.5, fines_bulles: 5.8 }, // %/m
  BA_rdt_dissolution_Hlim: 6,
  BA_rdt_dissolution_pente: -0.18,
  BA_rdt_dissolution_ordonnee: 6.88,
  BA_ASB: { brosses: 1.7, turbines_lentes: 1.65, turbines_rapides: 1.2 }, // kgO2/kWh
  BA_rdtDBO_FtoM_max: 1.9,
  BA_rdtDBO_coef_ED: [0.9549, -0.0966, -0.4705, 0.1798], // après primaire
  BA_rdtDBO_coef_EB: [0.9854, -0.2053, -0.2472, 0.1104], // eau brute
  BA_alpha_finebulle_a0: 0.51,
  BA_alpha_finebulle_aMVS: -0.062,
  BA_alpha_finebulle_aGeq: 0.019,
  BA_alpha_correction_MES_a0: 1.0007,
  BA_alpha_correction_MES_a1: -0.035,
  BA_alpha_correction_MES_Cref: 5,
  BA_facteur_boues_Clim: 120,
  BA_agitation_W_m3: { aeree: 0, anoxie: 3, chenal: 3, anaerobie: 10 },
  BA_agitation_tps_fct: 24,
  BA_bio_equiv_Tref: 12,
  BA_bio_equiv_correctionT: 1.072,
  BA_charge_massique_a0: 1.3908,
  BA_charge_massique_a1: 0.4959,
  BA_besoinO2_DBO_resp: { a0: 0.56, a1: 0.15, a2: 0.17, correctionT: 1.072, Tref: 15 },
  // --- filtration tertiaire sur sable (hyp_III_filtrasable_*), index = matériau
  filtrasable: {
    materiaux: ['ponce', 'schiste_1m', 'schiste_2m', 'sable_fin', 'sable', 'bicouche_1m', 'bicouche_2m'],
    labels: ['Ponce', 'Schiste TE<2,5 H 1 m', 'Schiste TE<2,5 H 2 m', 'Sable fin TE~1', 'Sable TE>1,35', 'Bicouche sable/ponce 1 m', 'Bicouche sable/ponce 2 m'],
    P_eau_lavage: [8.2, 7.2, 8.4, 7.8, 10, 7.4, 8.6],
    P_air_lavage: [4.1, 3.2, 4.4, 3.5, 6.1, 3.7, 4.8],
    eau_lavage_m3_m2_lavage: [6, 11, 13, 4, 9, 7, 9],
    capacite_retention: [4.5, 3, 4.5, 1.2, 3.5, 2, 4],
    vitesse_filtration: [15, 10, 15, 6, 15, 10, 15],
    hauteur_materiau: [2, 1, 2, 1, 2, 1, 2],
    eau_lavage_vitesse: [35, 60, 80, 25, 60, 35, 35],
    air_lavage_vitesse: [55, 50, 55, 50, 55, 50, 50],
    surface_unitaire_max: 125,
    nb_filtres_mini: 4,
    air_lavage_tps_fct_min_lavage: 9,
    pompe_ES_tps_fct: 18,
  },
}

// ---------------------------------------------------------------------------
// AA_collection.ratio(type_boue, type_ratio) — ratios de composition des boues
// ---------------------------------------------------------------------------
const RATIOS = {
  I_simple: { NK_MV: 0.059, Pt_MES: 0.009, DCO_MV: 1.71, DBO_MV: 0.7 },
  I_reactif: { NK_MV: 0.059, Pt_MES: 0.009, DCO_MV: 1.71, DBO_MV: 0.7 },
  II_forte: { NK_MV: 0.065, Pt_MES: 0.014, DCO_MV: 1.71, DBO_MV: 0.9 },
  II_moyenne: { NK_MV: 0.08, Pt_MES: 0.02, DCO_MV: 1.6, DBO_MV: 0.75 },
  II_faible_EB: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.5, DBO_MV: 0.5 },
  II_HybAS: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.5, DBO_MV: 0.5 },
  II_faible_ED: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.45, DBO_MV: 0.5 },
  II_prolongee_EB: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.45, DBO_MV: 0.5 },
  II_prolongee_ED: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.45, DBO_MV: 0.5 },
  II_MBR: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.5, DBO_MV: 0.5 }, // "A REVOIR" dans le VBA
  II_MBBR: { NK_MV: 0.095, Pt_MES: 0.02, DCO_MV: 1.5, DBO_MV: 0.5 },
  II_biostyr_C: { NK_MV: 0.01, Pt_MES: 0.025, DCO_MV: 1.6, DBO_MV: 0.9 },
  II_biostyr_N: { NK_MV: 0.01, Pt_MES: 0.025, DCO_MV: 1.55, DBO_MV: 0.9 },
  II_biostyr_NDN: { NK_MV: 0.01, Pt_MES: 0.025, DCO_MV: 1.5, DBO_MV: 0.9 },
  II_biostyr_PDN: { NK_MV: 0.01, Pt_MES: 0.015, DCO_MV: 1.45, DBO_MV: 0.9 },
  III_decantation: { NK_MV: 0.041, Pt_MES: 0.04, DCO_MV: 1.45, DBO_MV: 0.45 },
  III_biostyr_N: { NK_MV: 0.041, Pt_MES: 0.04, DCO_MV: 1.45, DBO_MV: 0.45 },
  III_biostyr_PDN: { NK_MV: 0.041, Pt_MES: 0.04, DCO_MV: 1.45, DBO_MV: 0.45 },
  codigestion_graisses: { NK_MV: 0.002, Pt_MES: 0.002, DCO_MV: 2.8, DBO_MV: 2.8 },
}
export const ratio = (typeBoue, typeRatio) => RATIOS[typeBoue]?.[typeRatio] ?? 0

// ---------------------------------------------------------------------------
// Fonctions utilitaires communes à plusieurs classes VBA
// ---------------------------------------------------------------------------

/** MOD_FonctionsPubliques.besoins_O2_HS : 2 gO2/gS, tout le sulfure est consommé */
export const besoinsO2HS = (Sh) => 2 * Sh

/** rendement moteur en fonction de la puissance (kW) — formule répétée dans D1, F3, F4, F5 */
export function rendementMoteur(puissance_kW) {
  if (puissance_kW <= 0) return 0.5
  return puissance_kW < 60 ? (5.7195 * Math.log(puissance_kW) + 72.682) / 100 : 0.961
}

/**
 * Rendement global d'une pompe (machine × moteur) selon la logique du classeur :
 * puissance = Q/3600 × g × P / rdt_machine, puis rendement moteur par la loi log.
 */
export function rendementPompeGlobal(Q_unitaire_m3h, P_mCE, rdt_machine = 0.7) {
  const puissance = (Q_unitaire_m3h / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * P_mCE / rdt_machine
  return rdt_machine * rendementMoteur(puissance)
}

/** énergie de pompage kWh/j = Q_unit/3600 × g × P / rdt × nb × tps */
export function electricitePompage(Q_unitaire_m3h, P_mCE, rdt_global, nb_pompes, tps_h_j) {
  if (!(rdt_global > 0)) return 0
  return (Q_unitaire_m3h / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * P_mCE / rdt_global * nb_pompes * tps_h_j
}

/**
 * Répartition débit / nombre de pompes / temps de fonctionnement — motif répété
 * dans les blocs "EXTRACTION DES BOUES" (D1, F3, F5). Retourne l'ensemble cohérent.
 * forced : { Q_unitaire?, tps?, nb? } (valeurs forcées par l'utilisateur)
 */
export function repartitionPompes(Q_jour, defaults, forced = {}, tpsMax = CONST.NOMBRE_HEURE_PAR_JOUR) {
  let nb = forced.nb ?? defaults.nb
  let tps = forced.tps ?? defaults.tps
  let Qu
  const ceilDiv = (a, b) => (b > 0 ? Math.ceil(a / b - 1e-12) : 0)
  if (forced.Q_unitaire != null) {
    Qu = forced.Q_unitaire
    if (forced.tps != null) {
      if (forced.nb == null) nb = ceilDiv(Q_jour, tps * Qu)
    } else {
      tps = nb * Qu > 0 ? Q_jour / (nb * Qu) : 0
      if (tps > tpsMax) {
        tps = tpsMax
        if (forced.nb == null) {
          nb = ceilDiv(Q_jour, tps * Qu)
          tps = nb * Qu > 0 ? Q_jour / (nb * Qu) : 0
        }
      }
    }
  } else {
    Qu = nb * tps > 0 ? Q_jour / (nb * tps) : 0
  }
  const incoherence = Q_jour > 0 ? (Qu * nb * tps - Q_jour) / Q_jour : 0
  return { nb, tps, Q_unitaire: Qu, incoherence: Math.abs(incoherence) > CONST.critere_incoherence }
}

/** pression atmosphérique (hPa) en fonction de l'altitude (m) */
export function pressionAtmospherique_hPa(altitude) {
  return (HYP.Patm_P0 * Math.pow(HYP.Patm_a0 + HYP.Patm_a1 * altitude, HYP.Patm_a2)) / 100
}

/** saturation O2 (mg/L) à T (°C) et pression P0 — polynôme degré 4 */
export function O2SaturationP0(T) {
  return HYP.O2sat_coef.reduce((acc, a, i) => acc + a * Math.pow(T, i), 0)
}

/**
 * Facteur K de transfert d'oxygène (bloc AERATION des classes BA).
 * @returns {{K, O2sat_T_P, O2sat_20_P, deltaP, Patm}}
 */
export function facteurK({ alpha, T_eau, altitude, hauteur_bassin, insufflation, O2_dissous }) {
  const Patm = pressionAtmospherique_hPa(altitude)
  const deltaP = insufflation
    ? (HYP.insufflation_deltaP_a0 / HYP.insufflation_deltaP_a1) * (HYP.Patm_P0 / 100) * (hauteur_bassin - HYP.insufflation_hauteur_diffuseur_m)
    : 0
  const corr = (Patm + deltaP) / (HYP.Patm_P0 / 100)
  const O2sat_20_P = O2SaturationP0(20) * corr
  const O2sat_T_P = O2SaturationP0(T_eau) * corr
  let K = (alpha * (HYP.facteurK_beta * O2sat_T_P - O2_dissous)) / O2sat_20_P
  K *= Math.pow(HYP.facteurK_correction_T, T_eau - HYP.facteurK_Tref)
  return { K, O2sat_T_P, O2sat_20_P, deltaP, Patm }
}

/**
 * Précipitation du P par FeCl3 — calcul du rendement sur P soluble par dichotomie
 * (Private Function coagulation_floculation_rdt_P_soluble dans D2, F3, F4).
 */
export function coagulationFloculationRdtPSoluble(Q, Pt, dosage_FeCl3, ratio_Psol_P, ratio_PO4_Psol, ratio_FeCl3_Pt) {
  if (!(dosage_FeCl3 > 0) || !(Q > 0)) return 0
  const Psol_in = (Pt / Q) * 1000 * ratio_Psol_P * ratio_PO4_Psol
  if (!(Psol_in > 0.1)) return 0
  let x_min = 0.1
  let x_max = Psol_in
  let Ptbr = (x_min + x_max) / 2
  let guard = 0
  while ((x_max - x_min) / Ptbr > 0.00001 && guard++ < 200) {
    const Psol_out = Psol_in - Ptbr
    const rm = 1.1 - (6 * (Math.pow(Ptbr + Psol_out, -0.5) - (Math.pow(Psol_out, -0.5) - 0.1))) / (Ptbr + Psol_out - (Psol_out - 0.1))
    const dose = Ptbr * rm * ratio_FeCl3_Pt
    if (dose < dosage_FeCl3) x_min = Ptbr
    else x_max = Ptbr
    Ptbr = (x_min + x_max) / 2
  }
  return Ptbr / Psol_in
}

/** ratio molaire Fe/P en fonction de Ptbr (P à abattre, mg/L) et de la garantie */
export function ratioMolairePtbr(Ptbr, Pt_garantie) {
  return 1.1 - (6 * (Math.pow(Ptbr + Pt_garantie, -0.5) - (Math.pow(Pt_garantie, -0.5) - 0.1))) / (Ptbr + Pt_garantie - (Pt_garantie - 0.1))
}

/**
 * Bloc commun "calcul de rdt_P et dosage_FeCl3" (F3 Décantation III, F4 Discfilter).
 * mode : { rdt_P_force, rdt_P_f, dosage_FeCl3_force, dosage_FeCl3 }
 */
export function precipitationP({ Q_traite, Pt_traite, Q, Pt, Pt_garantie, rdt_P_f, dosage_FeCl3_f, hyp }) {
  const { ratio_Psol_P, ratio_PO4_Psol, ratio_FeCl3_Pt, hyp_rdt_P_particulaire } = hyp
  let rdt_P, dosage_FeCl3 = dosage_FeCl3_f ?? 0, rdt_P_soluble
  const doseFromRdt = () => {
    rdt_P_soluble = (rdt_P - (1 - ratio_Psol_P) * hyp_rdt_P_particulaire) / ratio_Psol_P
    if (rdt_P_soluble < 0) { dosage_FeCl3 = 0; return }
    if (rdt_P_soluble > 1) rdt_P_soluble = 1
    const P_eliminer = rdt_P_soluble * ratio_Psol_P * Pt_traite
    const Ptbr = Q_traite > 0 ? (P_eliminer / Q_traite) * 1000 : 0
    dosage_FeCl3 = Ptbr > 0 ? ratioMolairePtbr(Ptbr, Pt_garantie) * ratio_FeCl3_Pt * Ptbr : 0
    if (dosage_FeCl3 < 0) dosage_FeCl3 = 0
  }
  if (rdt_P_f != null) {
    rdt_P = rdt_P_f
    if (dosage_FeCl3_f == null) doseFromRdt()
  } else if (dosage_FeCl3_f != null) {
    rdt_P_soluble = Math.min(1, coagulationFloculationRdtPSoluble(Q_traite, Pt_traite, dosage_FeCl3, ratio_Psol_P, ratio_PO4_Psol, ratio_FeCl3_Pt))
    rdt_P = ratio_Psol_P * rdt_P_soluble + (1 - ratio_Psol_P) * hyp_rdt_P_particulaire
  } else {
    const P_eliminer = Pt - (Pt_garantie * Q) / 1000
    if (P_eliminer < 0) {
      rdt_P = (1 - ratio_Psol_P) * hyp_rdt_P_particulaire
      dosage_FeCl3 = 0
    } else {
      rdt_P = Pt_traite > 0 ? P_eliminer / Pt_traite : 0
      doseFromRdt()
    }
  }
  return { rdt_P, dosage_FeCl3 }
}
