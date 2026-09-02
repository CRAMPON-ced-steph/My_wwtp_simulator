// ---------------------------------------------------------------------------
// Port de z_ANITA_Mox.cls et de z_CLS_RetoursAdmisTraitement.cls.
//
// ANITA Mox traite les jus de digestion et d'Athos par déammonification sur
// biofilm : nitritation partielle puis oxydation anaérobie de l'ammonium
// (procédé anammox). L'intérêt tient à trois chiffres, tous portés ici :
//
//   1,94 kg O2 par kg de N éliminé, contre 4,57 en nitrification classique ;
//   0,11 kg de NO3 formé par kg de N traité, contre 1 ;
//   aucune source de carbone à ajouter, là où une dénitrification classique
//   consommerait environ 4 kg de DCO par kg de NO3.
//
// La chaîne comporte au plus deux cuves :
//
//   jus admis → [cuve C, facultative] → cuve Mox → retour en tête
//
// La cuve C n'existe que si le rapport DCO soluble dégradable sur N-NH4 dépasse
// 0,4 : au-delà, les hétérotrophes concurrenceraient les bactéries anammox dans
// la cuve Mox. Elle est alors dimensionnée pour ramener ce rapport à 0,2.
//
// Le module de gestion des retours (`retoursAdmis`) décide en amont quels jus
// sont dirigés vers le traitement — digestion seule, Athos seul, ou les deux —
// et laisse le reste rejoindre directement la tête de station.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - le VBA appelle `dimensionnement` puis `fonctionnement_reel`, la seconde
//    passe reprenant les volumes de la première avec le débit réel ; la file
//    boues n'ayant qu'un jeu de flux, le port ne fait qu'une passe et expose la
//    charge nominale comme paramètre forçable ;
//  - le contrôle du temps de séjour minimal de 12 h est présent dans le VBA
//    mais son message d'alerte y est vide : le port émet un avertissement ;
//  - la gamme de débit du surpresseur n'influence pas la consommation
//    spécifique dans le classeur (les trois gammes portent la même valeur) : le
//    port ne retient que le type de surpresseur.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { RET, NB_RET, RET_ORIGINE } from '../core/sludge.js'
import { CONST, HYP } from '../core/hypotheses.js'

const MEDIA = {
  K3: { label: 'AnoxKaldnes K3', surface: 500, filling_max: 0.65 },
  K5: { label: 'AnoxKaldnes K5', surface: 800, filling_max: 0.55 },
  chip_M: { label: 'Chip M', surface: 1200, filling_max: 0.55 },
}
const MEDIA_OPTIONS = Object.entries(MEDIA).map(([value, m]) => ({ value, label: m.label }))

const H = {
  T_rdt_max: 30, // °C : au-delà, les vitesses ne progressent plus
  ratio_DCOsolDeg_NH4_maxi: 0.4, // seuil au-delà duquel une cuve C est nécessaire
  ratio_DCOsolDeg_NH4_sortie_cuveC: 0.2, // rapport visé en sortie de cuve C
  cuve_Mox_tps_sejour_min: 12, // h
  rdt_nit: 0.9,
  rate_NO3f_Nnit: 0.11, // kg NO3 formé par kg de N traité
  assimilation_N_DCOdegradee: 0.04,
  rdt_DCOsolDeg_lim: 0.9,
  rate_DCOsolDeg_30_max: 20, // g/(m²·j)
  DCOsolDeg_correction_T: 1.06,
  DCOsolDeg_T_ref: 30,
  N_correction_T: 1.07,
  N_T_ref: 30,
  rate_NorgParticulaire_ammonifie: 0.5,
  ratio_O2_DCOsoluble_degradable: 1.05, // dont 5 % pour la respiration endogène
  ratio_O2_nit: 1.94, // contre 4,57 en nitrification classique
  ratio_O2_MES: 0, // mis à zéro dans le classeur, « à revoir »
  respO2_correctif_T: 1.07,
  respO2_T_ref: 10,
  air_Qmin_agitation: 4, // Nm³/(m²·h)
  air_tps_fonctionnement: 24, // h/j
  hauteur_diffuseur: 0.25,
  correctif_T_K: 1.024,
  T_ref_K: 20,
  ratio_kgO2_Nm3air: 0.3,
  O2_facteur_beta: 0.95,
}
// Vitesses de déammonification à 30 °C (g N/(m²·j)), nettement plus élevées
// sur des jus de digestion simple que sur des jus d'hydrolyse thermique.
const RATE_N_30 = {
  digestion_simple: { K3: 2.6, K5: 2.6, chip_M: 2.25 },
  digestion_avancee: { K3: 1.35, K5: 1.35, chip_M: 1.12 },
}

/**
 * Sélectionne les jus admis au traitement et renvoie leurs caractéristiques,
 * équivalent de z_CLS_RetoursAdmisTraitement.dimensionnement.
 */
export function retoursAdmis(retours, choix) {
  const O = RET_ORIGINE
  const origines = choix === 'digestion' ? [O.digestion]
    : choix === 'athos' ? [O.athos]
      : [O.digestion, O.athos]
  const admis = new Array(NB_RET + 1).fill(0)
  const admis_soluble = new Array(NB_RET + 1).fill(0)
  for (const o of origines) {
    for (let i = 1; i <= NB_RET; i++) {
      admis[i] += retours[o]?.[i] || 0
      admis_soluble[i] += retours[`${o}_soluble`]?.[i] || 0
    }
  }
  const Q = admis[RET.Q]
  const mgL = (v) => (Q > 0 ? (v / Q) * 1000 : 0)
  return {
    origines,
    admis,
    admis_soluble,
    Q,
    concentrations: {
      DCO: mgL(admis[RET.DCO]), DBO: mgL(admis[RET.DBO]), MES: mgL(admis[RET.MES]),
      NK: mgL(admis[RET.NK]), NH4: mgL(admis[RET.NH4]), NO3: mgL(admis[RET.NO3]),
      Pt: mgL(admis[RET.Pt]), PO4: mgL(admis_soluble[RET.Pt]),
    },
  }
}

/**
 * Remplace, dans le jeu de vecteurs de retour, les flux admis au traitement par
 * les flux sortant du procédé — équivalent de `flux_sortie_process`.
 */
export function remplacerRetoursAdmis(retours, lu, sortie, sortie_soluble) {
  const O = RET_ORIGINE
  for (let i = 1; i <= NB_RET; i++) {
    // on retire les flux admis du total, puis on y verse les flux traités
    retours[O.total][i] += (sortie[i] || 0) - (lu.admis[i] || 0)
    retours[`${O.total}_soluble`][i] += (sortie_soluble[i] || 0) - (lu.admis_soluble[i] || 0)
    for (const o of lu.origines) {
      retours[o][i] = 0
      retours[`${o}_soluble`][i] = 0
    }
  }
  // les flux traités sont rattachés à une origine dédiée, non retraitable
  for (let i = 1; i <= NB_RET; i++) {
    retours[O.autres][i] += sortie[i] || 0
    retours[`${O.autres}_soluble`][i] += sortie_soluble[i] || 0
  }
  return retours
}

export default defineSludgeNode({
  id: 'anita-mox',
  label: 'ANITA Mox',
  short: 'ANITA Mox',
  family: 'retours',
  vba: 'z_ANITA_Mox.cls',
  etapeSortie: null,
  description:
    "Traitement des jus de digestion et d'Athos par déammonification sur biofilm. L'azote est éliminé par voie anammox : 1,94 kg d'O2 par kg de N contre 4,57 en nitrification classique, et sans apport de carbone.",
  choices: [
    { key: 'jus_traites', label: 'Jus dirigés vers le traitement', default: 'digestion', options: [
      { value: 'digestion', label: 'jus de digestion seuls' },
      { value: 'athos', label: 'jus d\'Athos seuls' },
      { value: 'les_deux', label: 'jus de digestion et d\'Athos' },
    ] },
    { key: 'type_digestion', label: 'Type de digestion en amont', default: 'simple', options: [
      { value: 'simple', label: 'digestion simple' },
      { value: 'avancee', label: 'digestion avancée (Biothelys, Exelys)' },
    ] },
    { key: 'media_Mox', label: 'Média de la cuve Mox', default: 'K5', options: MEDIA_OPTIONS },
    { key: 'media_cuve_C', label: 'Média de la cuve carbone', default: 'K3', options: MEDIA_OPTIONS },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
      { value: 'roots', label: 'tri-lobes type roots' },
      { value: 'vis', label: 'surpresseurs à vis' },
      { value: 'turbo', label: 'centrifuge type turbo' },
    ] },
  ],
  params: [
    { key: 'temperature', label: 'Température des jus', unit: '°C', group: 'Conditions', default: 25, hint: 'les vitesses plafonnent à 30 °C' },
    { key: 'Q_nominal', label: 'Débit nominal de jus', unit: 'm³/j', group: 'Conditions', default: undefined, hint: 'débit admis si non forcé' },
    { key: 'NH4_nominal', label: 'N-NH4 des jus', unit: 'mg/L', group: 'Conditions', default: undefined, hint: 'concentration admise si non forcée' },
    { key: 'MES_nominal', label: 'MES des jus', unit: 'mg/L', group: 'Conditions', default: undefined, hint: 'concentration admise si non forcée' },
    { key: 'DCO_soluble_degradable', label: 'DCO soluble dégradable', unit: 'mg/L', group: 'Conditions', default: undefined, hint: '50 % de la DCO soluble en digestion simple, 30 % sinon' },
    { key: 'cuve_C_volume', label: 'Volume de la cuve carbone', unit: 'm³', group: 'Cuve carbone', default: undefined, hint: 'dimensionnée si le rapport DCO/NH4 dépasse 0,4' },
    { key: 'cuve_C_filling', label: 'Taux de remplissage de la cuve carbone', unit: '-', group: 'Cuve carbone', default: 0.5 },
    { key: 'cuve_Mox_volume', label: 'Volume de la cuve Mox', unit: 'm³', group: 'Cuve Mox', default: undefined, hint: 'dimensionnée sur la charge azotée' },
    { key: 'cuve_Mox_filling', label: 'Taux de remplissage de la cuve Mox', unit: '-', group: 'Cuve Mox', default: undefined, hint: '0,4 avec du Chip M, 0,5 sinon' },
    { key: 'rdt_nit', label: 'Rendement de déammonification', unit: '-', group: 'Cuve Mox', default: 0.9 },
    { key: 'O2_dissous', label: 'O2 dissous', unit: 'mg/L', group: 'Aération', default: 2 },
    { key: 'O2_facteur_alfa', label: 'Facteur alpha', unit: '-', group: 'Aération', default: 0.7 },
    { key: 'O2_rdt_transfert', label: "Rendement de transfert de l'O2", unit: '%/m', group: 'Aération', default: 4 },
    { key: 'hauteur_bassin', label: "Hauteur d'eau des cuves", unit: 'm', group: 'Aération', default: 6 },
    { key: 'air_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Aération', default: undefined, hint: 'hauteur de bassin + 2' },
    { key: 'air_Q_Nm3j', label: "Débit d'air", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'surpresseur_conso_spec', label: 'Consommation spécifique des surpresseurs', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
  ],

  compute(ctx) {
    const { site, retours, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    if (!retours) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Les vecteurs de retour ne sont pas disponibles."] }
    }

    // ---- sélection des jus admis au traitement
    const lu = retoursAdmis(retours, choices.jus_traites)
    if (!(lu.Q > 0)) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Aucun jus admissible en amont : le traitement des retours suppose une digestion ou un Athos dans la filière boues."],
      }
    }

    // ---- attribution_valeur_par_defaut
    const T = p.temperature ?? 25
    if (T < 20) warnings.push(`Température des jus faible (${T.toFixed(0)} °C) : la déammonification perd fortement en vitesse sous 20 °C.`)
    const digestion_avancee = choices.type_digestion === 'avancee'
    const media_Mox = choices.media_Mox
    const media_C = choices.media_cuve_C
    const S_Mox = MEDIA[media_Mox].surface
    const S_C = MEDIA[media_C].surface
    const filling_max_Mox = MEDIA[media_Mox].filling_max
    let cuve_C_filling = f('cuve_C_filling') ?? 0.5
    let cuve_Mox_filling = f('cuve_Mox_filling') ?? (media_Mox === 'chip_M' ? 0.4 : 0.5)
    if (cuve_Mox_filling > filling_max_Mox) {
      warnings.push(`Taux de remplissage de la cuve Mox (${cuve_Mox_filling}) supérieur au maximum admissible pour ce média (${filling_max_Mox}).`)
    }
    const O2_dissous = p.O2_dissous ?? 2
    const O2_alfa = p.O2_facteur_alfa ?? 0.7
    const O2_rdt = p.O2_rdt_transfert ?? 4
    const hauteur = p.hauteur_bassin ?? 6
    const air_P_refoulement = f('air_P_refoulement') ?? hauteur + 2

    // ---- charges admises
    const Q = f('Q_nominal') ?? lu.Q
    const facteurDebit = Q > 0 ? lu.Q / Q : 1
    const c = lu.concentrations
    let DCO = (c.DCO * facteurDebit * Q) / 1000
    let DBO = (c.DBO * facteurDebit * Q) / 1000
    const MES = f('MES_nominal') != null ? (f('MES_nominal') * Q) / 1000 : (c.MES * facteurDebit * Q) / 1000
    const NH4_in = f('NH4_nominal') != null ? (f('NH4_nominal') * Q) / 1000 : (c.NH4 * facteurDebit * Q) / 1000
    const NK_in = ((c.NK - c.NH4) * facteurDebit * Q) / 1000 + NH4_in
    const Pt = (c.Pt * Q) / 1000

    // la fraction dégradable de la DCO soluble dépend du type de digestion :
    // l'hydrolyse thermique laisse une DCO plus réfractaire
    const ratio_DCOsolDeg = digestion_avancee ? 0.3 : 0.5
    let DCOsoluble_degradable = f('DCO_soluble_degradable') != null
      ? (f('DCO_soluble_degradable') * Q) / 1000
      : ratio_DCOsolDeg * lu.admis_soluble[RET.DCO]
    if (DCOsoluble_degradable > DCO) DCO = DCOsoluble_degradable

    /** vitesse de dégradation de la DCO soluble, plafonnée à 30 °C */
    const rate_DCO_T = T > H.T_rdt_max
      ? H.rate_DCOsolDeg_30_max
      : H.rate_DCOsolDeg_30_max * Math.pow(H.DCOsolDeg_correction_T, T - H.DCOsolDeg_T_ref)
    /** vitesse de déammonification, plafonnée à 30 °C */
    const rate_N_30 = RATE_N_30[digestion_avancee ? 'digestion_avancee' : 'digestion_simple'][media_Mox]
    const rate_N_T = T > H.T_rdt_max
      ? rate_N_30
      : rate_N_30 * Math.pow(H.N_correction_T, T - H.N_T_ref)

    // =====================================================================
    // Cuve carbone — n'existe que si les hétérotrophes menacent l'anammox
    // =====================================================================
    const ratio_DCOsolDeg_NH4 = NH4_in > 0 ? DCOsoluble_degradable / NH4_in : 0
    let cuve_C_volume = 0
    let DCOsolubleDegradee = 0
    if (ratio_DCOsolDeg_NH4 > H.ratio_DCOsolDeg_NH4_maxi) {
      const Vf = f('cuve_C_volume')
      if (Vf != null) {
        cuve_C_volume = Vf
        if (f('cuve_C_filling') == null) {
          DCOsolubleDegradee = DCOsoluble_degradable - H.ratio_DCOsolDeg_NH4_sortie_cuveC * NH4_in
          if (DCOsolubleDegradee > 0) {
            const volume_media = (DCOsolubleDegradee * 1000) / (rate_DCO_T * S_C)
            cuve_C_filling = cuve_C_volume > 0 ? volume_media / cuve_C_volume : 0
          } else { DCOsolubleDegradee = 0; cuve_C_filling = 0 }
        } else {
          DCOsolubleDegradee = (rate_DCO_T * S_C * cuve_C_volume * cuve_C_filling) / 1000
          if (DCOsoluble_degradable > 0 && DCOsolubleDegradee / DCOsoluble_degradable > H.rdt_DCOsolDeg_lim) {
            DCOsolubleDegradee = DCOsoluble_degradable * H.rdt_DCOsolDeg_lim
          }
        }
      } else {
        DCOsolubleDegradee = DCOsoluble_degradable - H.ratio_DCOsolDeg_NH4_sortie_cuveC * NH4_in
        if (DCOsolubleDegradee > 0) {
          const volume_media = (DCOsolubleDegradee * 1000) / (rate_DCO_T * S_C)
          cuve_C_volume = cuve_C_filling > 0 ? volume_media / cuve_C_filling : 0
        } else { DCOsolubleDegradee = 0; cuve_C_volume = 0 }
      }
    } else {
      cuve_C_volume = 0
      cuve_C_filling = 0
    }
    const N_assimile = H.assimilation_N_DCOdegradee * DCOsolubleDegradee

    // =====================================================================
    // Cuve Mox — déammonification
    // =====================================================================
    const rdt_nit = p.rdt_nit ?? H.rdt_nit
    // l'azote organique particulaire s'ammonifie pour moitié dans la cuve
    const N_disponible = NH4_in + H.rate_NorgParticulaire_ammonifie * (NK_in - NH4_in) - N_assimile
    let cuve_Mox_volume
    let NNH4_abattu
    const VMoxf = f('cuve_Mox_volume')
    if (VMoxf != null) {
      cuve_Mox_volume = VMoxf
      if (f('cuve_Mox_filling') == null) {
        NNH4_abattu = rdt_nit * N_disponible
        const volume_media = (NNH4_abattu * 1000) / (rate_N_T * S_Mox)
        cuve_Mox_filling = cuve_Mox_volume > 0 ? volume_media / cuve_Mox_volume : 0
        if (cuve_Mox_filling > filling_max_Mox) {
          cuve_Mox_filling = filling_max_Mox
          NNH4_abattu = (rate_N_T * cuve_Mox_volume * cuve_Mox_filling * S_Mox) / 1000
        }
      } else {
        NNH4_abattu = (rate_N_T * cuve_Mox_volume * cuve_Mox_filling * S_Mox) / 1000
        const max = rdt_nit * N_disponible
        if (NNH4_abattu > max) NNH4_abattu = max
      }
    } else {
      NNH4_abattu = rdt_nit * N_disponible
      const volume_media = (NNH4_abattu * 1000) / (rate_N_T * S_Mox)
      cuve_Mox_volume = cuve_Mox_filling > 0 ? volume_media / cuve_Mox_filling : 0
      // temps de séjour minimal de 12 h, quitte à surdimensionner la cuve
      const V_mini = (Q * H.cuve_Mox_tps_sejour_min) / CONST.NOMBRE_HEURE_PAR_JOUR
      if (cuve_Mox_volume < V_mini) {
        cuve_Mox_volume = V_mini
        if (f('cuve_Mox_filling') == null) {
          cuve_Mox_filling = cuve_Mox_volume > 0 ? volume_media / cuve_Mox_volume : 0
        } else {
          NNH4_abattu = (rate_N_T * cuve_Mox_volume * cuve_Mox_filling * S_Mox) / 1000
          const max = rdt_nit * N_disponible
          if (NNH4_abattu > max) NNH4_abattu = max
        }
      }
    }
    const HRT = Q > 0 ? (cuve_Mox_volume / Q) * CONST.NOMBRE_HEURE_PAR_JOUR : 0
    if (HRT < H.cuve_Mox_tps_sejour_min - 1e-6) {
      warnings.push(`Temps de séjour de la cuve Mox (${HRT.toFixed(1)} h) inférieur au minimum de ${H.cuve_Mox_tps_sejour_min} h.`)
    }

    // ---- besoins en O2 et flux de sortie
    const MES_concentration = Q > 0 ? MES / Q : 0
    const O2_besoin = H.ratio_O2_DCOsoluble_degradable * DCOsolubleDegradee
      + H.ratio_O2_nit * NNH4_abattu
      + H.ratio_O2_MES * Math.pow(H.respO2_correctif_T, T - H.respO2_T_ref) * MES_concentration * (cuve_C_volume + cuve_Mox_volume)
    const ratio_O2_Mox = O2_besoin > 0 ? (H.ratio_O2_nit * NNH4_abattu) / O2_besoin : 1

    const sortie = new Array(NB_RET + 1).fill(0)
    sortie[RET.Q] = Q
    sortie[RET.DCO] = DCO - DCOsolubleDegradee
    sortie[RET.DBO] = DBO > DCOsolubleDegradee ? DBO - DCOsolubleDegradee : 0
    sortie[RET.MES] = MES
    sortie[RET.NH4] = NH4_in + H.rate_NorgParticulaire_ammonifie * (NK_in - NH4_in) - NNH4_abattu - N_assimile
    sortie[RET.NK] = NK_in - NNH4_abattu - N_assimile
    sortie[RET.NO3] = H.rate_NO3f_Nnit * NNH4_abattu
    sortie[RET.Pt] = Pt

    const sortie_soluble = lu.admis_soluble.slice()
    sortie_soluble[RET.NK] = sortie_soluble[RET.NK] - sortie_soluble[RET.NH4] + sortie[RET.NH4]
    sortie_soluble[RET.NH4] = sortie[RET.NH4]
    sortie_soluble[RET.DCO] = Math.max(0, sortie_soluble[RET.DCO] - DCOsolubleDegradee)
    sortie_soluble[RET.DBO] = sortie_soluble[RET.DBO] > DCOsolubleDegradee
      ? sortie_soluble[RET.DBO] - DCOsolubleDegradee : 0
    sortie_soluble[RET.NO3] = sortie[RET.NO3]
    sortie_soluble[RET.Q] = Q

    remplacerRetoursAdmis(retours, lu, sortie, sortie_soluble)

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
      let K = (O2_alfa * (H.O2_facteur_beta * O2satT - O2_dissous)) / O2sat20
      K *= Math.pow(H.correctif_T_K, T - H.T_ref_K)
      const denom = K * (O2_rdt / 100) * (hauteur - H.hauteur_diffuseur) * H.ratio_kgO2_Nm3air

      // chaque cuve reçoit sa part du besoin, avec un plancher d'agitation
      let air_C = 0
      if (cuve_C_volume > 0 && denom > 0) {
        air_C = ((1 - ratio_O2_Mox) * O2_besoin) / denom
        const plancher = (H.air_Qmin_agitation * H.air_tps_fonctionnement * cuve_C_volume) / hauteur
        if (air_C < plancher) {
          air_C = plancher
          warnings.push("Débit d'air de la cuve carbone relevé pour assurer le brassage du média.")
        }
      }
      let air_Mox = denom > 0 ? (ratio_O2_Mox * O2_besoin) / denom : 0
      const plancher_Mox = (H.air_Qmin_agitation * H.air_tps_fonctionnement * cuve_Mox_volume) / hauteur
      if (air_Mox < plancher_Mox) {
        air_Mox = plancher_Mox
        warnings.push("Débit d'air de la cuve Mox relevé pour assurer le brassage du média.")
      }
      air_Q = air_C + air_Mox
    }
    if (choices.surpresseur === 'roots' && air_P_refoulement > HYP.surpresseur_Px2) {
      warnings.push(`Pression de refoulement (${air_P_refoulement.toFixed(1)} mCE) élevée pour des surpresseurs roots.`)
    }
    const conso_spec = f('surpresseur_conso_spec') ?? HYP.surpresseur_conso_spec_Wh_Nm3mCE[choices.surpresseur]
    const electricite = (air_Q * air_P_refoulement * conso_spec) / 1000

    // ---- comparaison avec une nitrification-dénitrification classique
    const O2_classique = 4.57 * NNH4_abattu
    const economie_O2 = O2_classique > 0 ? 1 - (H.ratio_O2_nit * NNH4_abattu) / O2_classique : 0
    const methanol_evite = 4 * NNH4_abattu * (1 - H.rate_NO3f_Nnit)

    return {
      etapeSortie: null,
      results: [
        { key: 'jus_Q', label: 'Débit de jus traités', unit: 'm³/j', value: Q },
        { key: 'jus_NH4', label: 'N-NH4 des jus', unit: 'mg/L', value: Q > 0 ? (NH4_in / Q) * 1000 : 0 },
        { key: 'jus_NK', label: 'NK des jus', unit: 'mg/L', value: Q > 0 ? (NK_in / Q) * 1000 : 0 },
        { key: 'jus_N_kg', label: 'Charge azotée admise', unit: 'kg N/j', value: NK_in },
        { key: 'ratio_C_N', label: 'Rapport DCO soluble dégradable / N-NH4', unit: '-', value: ratio_DCOsolDeg_NH4 },
        { key: 'V_C', label: 'Volume de la cuve carbone', unit: 'm³', value: cuve_C_volume },
        { key: 'fill_C', label: 'Taux de remplissage de la cuve carbone', unit: '-', value: cuve_C_filling },
        { key: 'DCO_deg', label: 'DCO soluble dégradée', unit: 'kg/j', value: DCOsolubleDegradee },
        { key: 'V_Mox', label: 'Volume de la cuve Mox', unit: 'm³', value: cuve_Mox_volume },
        { key: 'fill_Mox', label: 'Taux de remplissage de la cuve Mox', unit: '-', value: cuve_Mox_filling },
        { key: 'HRT', label: 'Temps de séjour de la cuve Mox', unit: 'h', value: HRT },
        { key: 'rate_N', label: 'Vitesse de déammonification', unit: 'g N/(m²·j)', value: rate_N_T },
        { key: 'N_abattu', label: 'Azote éliminé', unit: 'kg N/j', value: NNH4_abattu },
        { key: 'rdt_N', label: "Rendement d'élimination de l'azote", unit: '-', value: NK_in > 0 ? NNH4_abattu / NK_in : 0 },
        { key: 'NO3_forme', label: 'Nitrates formés', unit: 'kg N/j', value: sortie[RET.NO3] },
        { key: 'O2', label: "Besoin en O2", unit: 'kg O2/j', value: O2_besoin },
        { key: 'O2_spec', label: "O2 total par kg d'azote éliminé (azote + carbone)", unit: 'kg O2/kg N', value: NNH4_abattu > 0 ? O2_besoin / NNH4_abattu : 0 },
        { key: 'O2_N', label: "dont O2 de la voie azote", unit: 'kg O2/kg N', value: H.ratio_O2_nit },
        { key: 'eco_O2', label: "Économie d'O2 sur la voie azote seule", unit: '-', value: economie_O2 },
        { key: 'meoh', label: 'Méthanol évité par rapport à une dénitrification classique', unit: 'kg/j', value: methanol_evite },
        { key: 'air', label: "Débit d'air", unit: 'Nm³/h', value: air_Q / 24 },
        { key: 'out_NH4', label: 'N-NH4 en sortie', unit: 'mg/L', value: Q > 0 ? (sortie[RET.NH4] / Q) * 1000 : 0 },
        { key: 'out_NO3', label: 'N-NO3 en sortie', unit: 'mg/L', value: Q > 0 ? (sortie[RET.NO3] / Q) * 1000 : 0 },
        { key: 'out_N_kg', label: 'Azote résiduel renvoyé en tête', unit: 'kg N/j', value: sortie[RET.NK] + sortie[RET.NO3] },
      ],
      electricity: { total: electricite, fixed: 0, detail: { aeration: electricite } },
      warnings,
    }
  },
})
