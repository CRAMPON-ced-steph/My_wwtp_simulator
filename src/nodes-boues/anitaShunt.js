// ---------------------------------------------------------------------------
// Port de z_ANITA_Shunt.cls — traitement des jus par nitritation-dénitritation.
//
// Le Shunt est l'alternative à ANITA Mox : au lieu de la voie anammox, il
// exploite le raccourci par les nitrites. L'oxydation s'arrête à NO2 au lieu
// d'aller jusqu'à NO3, puis la dénitritation ramène le nitrite à l'azote gazeux.
//
// Les deux économies portées ici tiennent à ce raccourci :
//
//   3,43 kg O2 par kg de N nitrité, soit 75 % de la valeur classique de 4,57 ;
//   2,4 kg de DCO par kg de N dénitrité, soit 60 % des 4 kg d'une
//   dénitrification passant par les nitrates.
//
// Le procédé fonctionne en réacteur séquentiel (SBR) précédé d'une bâche
// tampon, dimensionnée sur le décalage entre le rythme d'alimentation en jus et
// le fonctionnement continu du réacteur.
//
// Différence majeure avec ANITA Mox : le Shunt travaille en biomasse libre et
// **produit des boues**, extraites à l'âge de boues visé de 15 jours. Ces boues
// rejoignent la file boues. Il consomme aussi du méthanol dès que la DBO
// soluble des jus ne suffit pas à la dénitritation.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - le VBA fait deux passes, la seconde reprenant les volumes de la première ;
//    le port n'en fait qu'une, la file boues n'ayant qu'un jeu de flux ;
//  - `boues_Q` est utilisé dans le calcul de `boues_MES` avant d'être calculé :
//    il vaut 0 au premier passage. Le port résout la boucle par un point fixe,
//    ce qui donne le résultat vers lequel le VBA converge sur ses itérations ;
//  - les contrôles de MES d'entrée et de rendement portent des messages vides
//    dans le classeur : le port émet de vrais avertissements.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { RET, NB_RET } from '../core/sludge.js'
import { CONST, HYP } from '../core/hypotheses.js'
import { retoursAdmis, remplacerRetoursAdmis } from './anitaMox.js'

const H = {
  tampon_tps_mini_retention: 4, // h
  tampon_securite: 1.1,
  G_SBR: 15, // j, âge de boues visé
  inlet_MES_max: 1.2, // g/L
  rdt_nit_vise: 0.95,
  rdt_DBO_sol: 0.95,
  rdt_NO2: 1,
  // 6 kg N-NH4 par kg de MVS et par heure d'aération, sur 13,5 h/j
  rate_nit_MVS: (6 / 1000) * 13.5,
  tps_anoxie: (1.5 / 8) * 24, // h/j
  agitation_tampon_W_m3: 3,
  agitation_SBR_W_m3: 3,
  SBR_concentration_MES: 5, // g/L
  // le raccourci par les nitrites économise 25 % de l'oxygène
  ratio_O2_nit: 4.57 * 0.75,
  ratio_O2_MES: 24 * 0.002,
  respO2_correctif_T: 1.07,
  respO2_T_ref: 10,
  // et 40 % du carbone : 4 kg DCO/kg N en dénitrification classique
  facteur_DCO_apportee: 4 * 0.6,
  DCO_methanol: 1.5, // kg DCO par kg de méthanol
  air_tps_fonctionnement: 13.5, // h/j
  hauteur_diffuseur: 0.25,
  correctif_T_K: 1.024,
  T_ref_K: 20,
  ratio_kgO2_Nm3air: 0.3,
  O2_facteur_beta: 0.99,
  // l'alpha se dégrade avec la concentration en MES de la liqueur
  a0_correctif_alfa_MES: 1.0007,
  a1_correctif_alfa_MES: -0.035,
  MES_reference_alfa: 5, // g/L
  critere_convergence: 1e-9,
}

export default defineSludgeNode({
  id: 'anita-shunt',
  label: 'ANITA Shunt',
  short: 'Shunt',
  family: 'retours',
  vba: 'z_ANITA_Shunt.cls',
  etapeSortie: null,
  description:
    "Traitement des jus par nitritation-dénitritation en réacteur séquentiel. En arrêtant l'oxydation au nitrite, le procédé économise 25 % de l'oxygène et 40 % du carbone d'une nitrification-dénitrification classique, au prix d'une production de boues.",
  choices: [
    { key: 'jus_traites', label: 'Jus dirigés vers le traitement', default: 'digestion', options: [
      { value: 'digestion', label: 'jus de digestion seuls' },
      { value: 'athos', label: "jus d'Athos seuls" },
      { value: 'les_deux', label: "jus de digestion et d'Athos" },
    ] },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
      { value: 'roots', label: 'tri-lobes type roots' },
      { value: 'vis', label: 'surpresseurs à vis' },
      { value: 'turbo', label: 'centrifuge type turbo' },
    ] },
  ],
  params: [
    { key: 'temperature', label: 'Température des jus', unit: '°C', group: 'Conditions', default: 25 },
    { key: 'alim_jour_semaine', label: "Jours d'alimentation en jus par semaine", unit: 'j/sem', group: 'Bâche tampon', default: 7 },
    { key: 'alim_heure_jour', label: "Heures d'alimentation en jus par jour", unit: 'h/j', group: 'Bâche tampon', default: 24 },
    { key: 'volume_tampon', label: 'Volume de la bâche tampon', unit: 'm³', group: 'Bâche tampon', default: undefined, hint: 'calculé sur le décalage d\'alimentation' },
    { key: 'volume_SBR', label: 'Volume du réacteur séquentiel', unit: 'm³', group: 'Réacteur', default: undefined, hint: 'dimensionné sur la charge azotée' },
    { key: 'SBR_concentration_MES', label: 'MES dans le réacteur', unit: 'g/L', group: 'Réacteur', default: 5 },
    { key: 'rdt_nit_vise', label: 'Rendement de nitritation visé', unit: '-', group: 'Réacteur', default: 0.95 },
    { key: 'boues_MV_MES', label: 'MV/MES des boues produites', unit: '-', group: 'Boues', default: 0.8 },
    { key: 'boues_concentration', label: 'Concentration des boues extraites', unit: 'g/L', group: 'Boues', default: undefined, hint: '1 000 / indice de boues, soit 6,67 g/L' },
    { key: 'outlet_MES', label: "MES de l'effluent traité", unit: 'mg/L', group: 'Boues', default: 100 },
    { key: 'methanol_flux', label: 'Consommation de méthanol', unit: 'kg/j', group: 'Réactifs', default: undefined, hint: "calculée sur le déficit de carbone" },
    { key: 'O2_dissous', label: 'O2 dissous', unit: 'mg/L', group: 'Aération', default: 1 },
    { key: 'O2_facteur_alfa', label: 'Facteur alpha', unit: '-', group: 'Aération', default: 0.69 },
    { key: 'O2_rdt_transfert', label: "Rendement de transfert de l'O2", unit: '%/m', group: 'Aération', default: 5.5 },
    { key: 'hauteur_bassin', label: "Hauteur d'eau du réacteur", unit: 'm', group: 'Aération', default: 6 },
    { key: 'air_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Aération', default: undefined, hint: 'hauteur de bassin + 2' },
    { key: 'air_Q_Nm3j', label: "Débit d'air", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'surpresseur_conso_spec', label: 'Consommation spécifique des surpresseurs', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'alim_nb_pompe', label: "Nombre de pompes d'alimentation", unit: 'u', group: 'Pompages', default: 2 },
    { key: 'alim_P_refoulement', label: "Pression de refoulement d'alimentation", unit: 'mCE', group: 'Pompages', default: 5 },
    { key: 'alim_tps_fct_pompe', label: "Durée de fonctionnement d'alimentation", unit: 'h/j', group: 'Pompages', default: 24 },
    { key: 'extraction_P_refoulement', label: "Pression de refoulement d'extraction", unit: 'mCE', group: 'Pompages', default: 5 },
  ],

  compute(ctx) {
    const { site, retours, choices, forced, p } = ctx
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

    // ---- attribution_valeur_par_defaut
    const T = p.temperature ?? 25
    const alim_jour_semaine = p.alim_jour_semaine ?? 7
    const alim_heure_jour = p.alim_heure_jour ?? 24
    const SBR_MES = p.SBR_concentration_MES ?? H.SBR_concentration_MES
    const rdt_nit_vise = p.rdt_nit_vise ?? H.rdt_nit_vise
    const boues_MV_MES = p.boues_MV_MES ?? 0.8
    const boues_concentration = f('boues_concentration') ?? 1000 / 150
    const outlet_MES = p.outlet_MES ?? 100
    const O2_dissous = p.O2_dissous ?? 1
    const O2_alfa_base = p.O2_facteur_alfa ?? 0.69
    const O2_rdt = p.O2_rdt_transfert ?? 5.5
    const hauteur = p.hauteur_bassin ?? 6
    const air_P_refoulement = f('air_P_refoulement') ?? hauteur + 2

    // ---- charges admises
    const inlet_Q = lu.Q
    const c = lu.concentrations
    const inlet_DCO = (c.DCO * inlet_Q) / 1000
    const inlet_DBO = (c.DBO * inlet_Q) / 1000
    const inlet_MES = (c.MES * inlet_Q) / 1000
    const inlet_NK = (c.NK * inlet_Q) / 1000
    const inlet_NH4 = (c.NH4 * inlet_Q) / 1000
    const inlet_NO3 = (c.NO3 * inlet_Q) / 1000
    const inlet_Pt = (c.Pt * inlet_Q) / 1000
    const DCO_soluble = lu.admis_soluble[RET.DCO]
    const DBO_soluble = lu.admis_soluble[RET.DBO]
    const NK_soluble = lu.admis_soluble[RET.NK]

    if (inlet_Q > 0 && inlet_MES / inlet_Q > H.inlet_MES_max) {
      warnings.push(`MES des jus (${(inlet_MES / inlet_Q).toFixed(2)} g/L) supérieures au maximum admissible de ${H.inlet_MES_max} g/L : une clarification amont est nécessaire.`)
    }

    // ---- bâche tampon : elle absorbe le décalage entre le rythme
    // d'alimentation en jus et le fonctionnement continu du réacteur
    let volume_tampon = f('volume_tampon')
    let tampon_tps_retention = 0
    if (volume_tampon == null) {
      const Q_m3h = ((inlet_Q / alim_jour_semaine) * CONST.NOMBRE_JOUR_PAR_SEMAINE) / alim_heure_jour
      tampon_tps_retention = (CONST.NOMBRE_HEURE_PAR_JOUR - alim_heure_jour) <= H.tampon_tps_mini_retention
        ? H.tampon_tps_mini_retention
        : CONST.NOMBRE_HEURE_PAR_JOUR - alim_heure_jour
      volume_tampon = tampon_tps_retention * Q_m3h * H.tampon_securite
    }

    // ---- réacteur séquentiel : dimensionné sur la charge azotée
    let volume_SBR, SBR_MVS, NH4_nit
    const Vf = f('volume_SBR')
    if (Vf != null) {
      volume_SBR = Vf
      SBR_MVS = SBR_MES * boues_MV_MES * volume_SBR
      NH4_nit = H.rate_nit_MVS * SBR_MVS
      if (NH4_nit > inlet_NH4) NH4_nit = inlet_NH4
    } else {
      NH4_nit = rdt_nit_vise * inlet_NH4
      SBR_MVS = H.rate_nit_MVS > 0 ? NH4_nit / H.rate_nit_MVS : 0
      volume_SBR = SBR_MES * boues_MV_MES > 0 ? SBR_MVS / (SBR_MES * boues_MV_MES) : 0
    }

    // ---- carbone : la DBO soluble des jus est consommée en priorité,
    // le méthanol ne comble que le déficit
    const DBO_traitee = DBO_soluble * H.rdt_DBO_sol
    const DCO_necessaire = H.facteur_DCO_apportee * H.rdt_NO2 * NH4_nit
    const methanol_flux = f('methanol_flux')
      ?? (DCO_necessaire > DBO_traitee ? (DCO_necessaire - DBO_traitee) / H.DCO_methanol : 0)

    const O2_besoin = H.ratio_O2_nit * NH4_nit
      + H.ratio_O2_MES * Math.pow(H.respO2_correctif_T, T - H.respO2_T_ref) * (boues_MV_MES > 0 ? SBR_MVS / boues_MV_MES : 0)

    // ---- extraction des boues à l'âge de boues visé
    // Le VBA emploie `boues_Q` avant de l'avoir calculé : il vaut 0 au premier
    // passage puis converge sur les itérations du programme principal. Le port
    // résout directement le point fixe.
    let boues_Q = 0
    let boues_MES = 0
    let garde = 0
    let ecart = Infinity
    while (ecart > H.critere_convergence && garde++ < 200) {
      const prev = boues_Q
      boues_MES = (volume_SBR * SBR_MES) / H.G_SBR - (outlet_MES * (inlet_Q - boues_Q)) / 1000
      boues_Q = boues_concentration > 0 ? boues_MES / boues_concentration : 0
      ecart = Math.abs(boues_Q - prev)
    }
    if (boues_MES < 0) {
      warnings.push("Production de boues négative : les MES de l'effluent dépassent ce que l'âge de boues permet d'extraire.")
      boues_MES = 0
      boues_Q = 0
    }

    // ---- flux de sortie
    const Q = inlet_Q - boues_Q
    const MES = (outlet_MES * Q) / 1000
    const partParticulaire = inlet_MES > 0 ? MES / inlet_MES : 0
    const DCO = (DCO_soluble - DBO_traitee) + partParticulaire * (inlet_DCO - DCO_soluble)
    const DBO = (DBO_soluble - DBO_traitee) + partParticulaire * (inlet_DBO - DBO_soluble)
    const NK = (NK_soluble - NH4_nit) + partParticulaire * (inlet_NK - NK_soluble)
    const NH4 = inlet_NH4 - NH4_nit
    const NO3 = inlet_Q > 0 ? (inlet_NO3 * Q) / inlet_Q : 0

    const sortie = new Array(NB_RET + 1).fill(0)
    sortie[RET.Q] = Q
    sortie[RET.DCO] = Math.max(0, DCO)
    sortie[RET.DBO] = Math.max(0, DBO)
    sortie[RET.MES] = MES
    sortie[RET.NK] = Math.max(0, NK)
    sortie[RET.NH4] = Math.max(0, NH4)
    sortie[RET.NO3] = NO3
    sortie[RET.Pt] = inlet_Pt

    const sortie_soluble = lu.admis_soluble.slice()
    sortie_soluble[RET.Q] = Q
    sortie_soluble[RET.NK] = Math.max(0, sortie_soluble[RET.NK] - sortie_soluble[RET.NH4] + sortie[RET.NH4])
    sortie_soluble[RET.NH4] = sortie[RET.NH4]
    sortie_soluble[RET.DCO] = Math.max(0, DCO_soluble - DBO_traitee)
    sortie_soluble[RET.DBO] = Math.max(0, DBO_soluble - DBO_traitee)
    sortie_soluble[RET.NO3] = NO3
    remplacerRetoursAdmis(retours, lu, sortie, sortie_soluble)

    // Le VBA calcule les boues extraites sur l'âge de boues visé (V × MES / G),
    // et non par différence entre l'entrée et la sortie : le bilan solide ne
    // boucle donc pas. L'écart est calculé et signalé.
    const MES_non_bouclees = inlet_MES - MES - boues_MES
    if (Math.abs(MES_non_bouclees) > 1e-6) {
      warnings.push(`Le bilan solide ne boucle pas de ${MES_non_bouclees.toFixed(0)} kg MES/j : les boues extraites sont calculées sur l'âge de boues visé et non par bilan matière, conformément au classeur d'origine.`)
    }

    // ---- boues produites, renvoyées vers la file boues
    const sludge = boues_MES > 0 ? {
      origine: 'II_shunt',
      Q: boues_Q,
      MES: boues_MES,
      concentration: boues_concentration,
      MV_MES: boues_MV_MES,
      DCO: Math.max(0, inlet_DCO - (DCO + DBO_traitee)),
      DBO: Math.max(0, inlet_DBO - (DBO + DBO_traitee)),
      NK: Math.max(0, inlet_NK - (NK + NH4_nit)),
      Pt: Math.max(0, inlet_Pt - sortie[RET.Pt]),
    } : null

    // =====================================================================
    // calcul_consommation_electrique
    // =====================================================================
    let air_Q = f('air_Q_Nm3j')
    if (air_Q == null) {
      const Patm = (HYP.Patm_P0 * Math.pow(HYP.Patm_a0 + HYP.Patm_a1 * site.altitude, HYP.Patm_a2)) / 100
      const delta_P = (0.35 / 10.33) * (HYP.Patm_P0 / 100) * (hauteur - H.hauteur_diffuseur)
      let O2sat20 = 0, O2satT = 0
      for (let k = 0; k <= 4; k++) {
        O2sat20 += HYP.O2sat_coef[k] * Math.pow(20, k)
        O2satT += HYP.O2sat_coef[k] * Math.pow(T, k)
      }
      const corr = (Patm + delta_P) / (HYP.Patm_P0 / 100)
      O2sat20 *= corr; O2satT *= corr
      // l'alpha se dégrade quand la liqueur s'épaissit
      const alfa = O2_alfa_base * (H.a0_correctif_alfa_MES + H.a1_correctif_alfa_MES * (SBR_MES - H.MES_reference_alfa))
      let K = (alfa * (H.O2_facteur_beta * O2satT - O2_dissous)) / O2sat20
      K *= Math.pow(H.correctif_T_K, T - H.T_ref_K)
      const denom = K * (O2_rdt / 100) * (hauteur - H.hauteur_diffuseur) * H.ratio_kgO2_Nm3air
      air_Q = denom > 0 ? O2_besoin / denom : 0
    }
    if (choices.surpresseur === 'roots' && air_P_refoulement > HYP.surpresseur_Px2) {
      warnings.push(`Pression de refoulement (${air_P_refoulement.toFixed(1)} mCE) élevée pour des surpresseurs roots.`)
    }
    const conso_spec = f('surpresseur_conso_spec') ?? HYP.surpresseur_conso_spec_Wh_Nm3mCE[choices.surpresseur]
    const electricite_aeration = (air_Q * air_P_refoulement * conso_spec) / 1000
    const electricite_agitation =
      (H.agitation_tampon_W_m3 * volume_tampon * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
      + (H.agitation_SBR_W_m3 * volume_SBR * H.tps_anoxie) / 1000
    const g = CONST.ACCELERATION_PESANTEUR_m_s2 / CONST.NOMBRE_SECONDE_PAR_HEURE
    const rdt_pompe = 0.7 * 0.88
    const electricite_alimentation = (g * inlet_Q * (p.alim_P_refoulement ?? 5)) / rdt_pompe
    const electricite_extraction = (g * boues_Q * (p.extraction_P_refoulement ?? 5)) / rdt_pompe
    const total = electricite_aeration + electricite_agitation + electricite_alimentation + electricite_extraction

    // ---- comparaison avec une nitrification-dénitrification classique
    const O2_classique = 4.57 * NH4_nit
    const economie_O2 = O2_classique > 0 ? 1 - (H.ratio_O2_nit * NH4_nit) / O2_classique : 0
    const DCO_classique = 4 * NH4_nit
    const economie_C = DCO_classique > 0 ? 1 - DCO_necessaire / DCO_classique : 0

    return {
      etapeSortie: null,
      sludge,
      results: [
        { key: 'jus_Q', label: 'Débit de jus traités', unit: 'm³/j', value: inlet_Q },
        { key: 'jus_NH4', label: 'N-NH4 des jus', unit: 'mg/L', value: c.NH4 },
        { key: 'jus_N_kg', label: 'Charge azotée admise', unit: 'kg N/j', value: inlet_NK },
        { key: 'V_tampon', label: 'Volume de la bâche tampon', unit: 'm³', value: volume_tampon },
        ...(tampon_tps_retention > 0 ? [{ key: 'tampon_HRT', label: 'Temps de rétention du tampon', unit: 'h', value: tampon_tps_retention }] : []),
        { key: 'V_SBR', label: 'Volume du réacteur séquentiel', unit: 'm³', value: volume_SBR },
        { key: 'SBR_MVS', label: 'Masse de MVS dans le réacteur', unit: 'kg', value: SBR_MVS },
        { key: 'HRT_SBR', label: 'Temps de séjour du réacteur', unit: 'h', value: inlet_Q > 0 ? (volume_SBR / inlet_Q) * 24 : 0 },
        { key: 'N_nit', label: 'Azote nitrité puis dénitrité', unit: 'kg N/j', value: NH4_nit },
        { key: 'rdt_N', label: "Rendement d'élimination de l'azote", unit: '-', value: inlet_NK > 0 ? NH4_nit / inlet_NK : 0 },
        { key: 'O2', label: "Besoin en O2", unit: 'kg O2/j', value: O2_besoin },
        { key: 'O2_spec', label: "O2 par kg d'azote nitrité", unit: 'kg O2/kg N', value: H.ratio_O2_nit },
        { key: 'eco_O2', label: "Économie d'O2 par rapport à la voie nitrate", unit: '-', value: economie_O2 },
        { key: 'DBO_jus', label: 'DBO soluble des jus utilisée comme carbone', unit: 'kg/j', value: DBO_traitee },
        { key: 'DCO_nec', label: 'DCO nécessaire à la dénitritation', unit: 'kg/j', value: DCO_necessaire },
        { key: 'eco_C', label: 'Économie de carbone par rapport à la voie nitrate', unit: '-', value: economie_C },
        { key: 'meoh', label: 'Consommation de méthanol', unit: 'kg/j', value: methanol_flux },
        { key: 'air', label: "Débit d'air", unit: 'Nm³/h', value: air_Q / H.air_tps_fonctionnement },
        { key: 'boues', label: 'Boues produites', unit: 'kg MES/j', value: boues_MES },
        { key: 'MES_non_bouclees', label: 'MES non bouclées au bilan solide', unit: 'kg/j', value: MES_non_bouclees },
        { key: 'boues_Q', label: 'Débit de boues extraites', unit: 'm³/j', value: boues_Q },
        { key: 'out_NH4', label: 'N-NH4 en sortie', unit: 'mg/L', value: Q > 0 ? (sortie[RET.NH4] / Q) * 1000 : 0 },
        { key: 'out_N_kg', label: 'Azote résiduel renvoyé en tête', unit: 'kg N/j', value: sortie[RET.NK] + sortie[RET.NO3] },
      ],
      reactifs: methanol_flux > 0 ? { methanol: methanol_flux } : {},
      electricity: {
        total,
        fixed: electricite_agitation,
        detail: {
          aeration: electricite_aeration,
          agitation: electricite_agitation,
          alimentation: electricite_alimentation,
          extraction: electricite_extraction,
        },
      },
      warnings,
    }
  },
})
