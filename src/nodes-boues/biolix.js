// ---------------------------------------------------------------------------
// Port de z_Biolix_graisses.cls — traitement biologique aérobie des graisses.
//
// Les graisses extraites au dégraisseur sont diluées à 75 g/L de DCO, puis
// dégradées en aérobie dans un bassin à très long temps de séjour — 20 jours,
// contre 1 jour pour la cuve d'homogénéisation qui le précède. Cette durée est
// la signature du procédé : les graisses sont lentement biodégradables, et
// c'est justement ce qui interdit de les envoyer telles quelles au traitement
// secondaire.
//
//   graisses du dégraisseur → dilution à 75 g DCO/L
//                           → cuve d'homogénéisation (1 j)
//                           → bassin biologique aéré (20 j)
//                           → boues Biolix, dirigées vers le traitement secondaire
//
// Le procédé consomme des nutriments : les graisses sont riches en carbone mais
// pauvres en azote et en phosphore, si bien qu'il faut en ajouter à raison de
// 25 g de N et 5 g de P par kg de DCO entrante.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - `md_PartBouesBiolixVersTraitementSecondaireEau` vaut toujours 1 dans le
//    classeur, et le test « si pas de traitement secondaire » y a un corps
//    vide : les boues partent donc systématiquement au secondaire. Le port
//    signale le cas où aucun secondaire n'existe ;
//  - le VBA fait deux passes ; le port n'en fait qu'une, la file boues n'ayant
//    qu'un jeu de flux, et expose la charge de dimensionnement en paramètre.
// ---------------------------------------------------------------------------
import { defineSludgeNode } from '../core/sludgeEngine.js'
import { ETAPE, P, SOL, TYPE } from '../core/sludge.js'
import { CONST, HYP, rendementMoteur } from '../core/hypotheses.js'
import { repartitionPompage, rdtPompeBoues, elecPompage } from './_commun.js'

const H = {
  // concentration de la charge diluée entrant au procédé (note 3NP60 p.21/40)
  entree_diluee_DCO_gL: 75,
  besoin_O2_kgO2_kgDCO: 1,
  // facteur K de transfert : moyenne entre la note procédé (0,30) et le
  // tableur de dimensionnement (0,45)
  aeration_facteur_K: (0.3 + 0.45) / 2,
  agitation_homogeneisation_hj: 24,
  HRT_homogeneisation_j: 1,
  HRT_bassin_biologique_j: 20,
  besoin_P_kgP_kgDCO: 0.005,
  besoin_N_kgN_kgDCO: 0.025,
  hauteur_diffuseur: 0.25,
  production_boues_kgMVS_kgDCO: 0.3,
  // rendement de transfert : moyenne note procédé (5,25 %/m) et tableur (4 %/m)
  rdt_transfert_defaut: (5.25 + 4) / 2,
  rdt_elimination_DCO_defaut: 0.7,
  boues_MV_MES_defaut: 0.76, // FNDAE 24 p.32
  agitation_W_m3_defaut: 25, // 20 à 30 selon la note procédé
  hauteur_bassin_defaut: 6,
}

export default defineSludgeNode({
  id: 'biolix',
  label: 'Biolix graisses',
  short: 'Biolix',
  family: 'retours',
  vba: 'z_Biolix_graisses.cls',
  etapeSortie: ETAPE.graisses_biolix,
  description:
    "Traitement biologique aérobie des graisses extraites au dégraisseur, à très long temps de séjour. Les graisses sont diluées, homogénéisées puis dégradées en 20 jours ; les boues produites rejoignent le traitement secondaire.",
  choices: [
    { key: 'agitation_homogeneisation', label: "Agitation de la cuve d'homogénéisation", default: 'oui', options: [
      { value: 'oui', label: 'oui' }, { value: 'non', label: 'non' },
    ] },
    { key: 'agitation_bassin', label: 'Agitation du bassin biologique', default: 'non', options: [
      { value: 'non', label: "non (l'air suffit au brassage)" },
      { value: 'oui', label: 'oui (bassin sous-chargé)' },
    ] },
    { key: 'surpresseur', label: 'Type de surpresseur', default: 'roots', options: [
      { value: 'roots', label: 'tri-lobes type roots' },
      { value: 'vis', label: 'surpresseurs à vis' },
      { value: 'turbo', label: 'centrifuge type turbo' },
    ] },
  ],
  params: [
    { key: 'part_flux_admis', label: 'Part des graisses admises', unit: '-', group: 'Alimentation', default: 1 },
    { key: 'DCO_dimensionnement', label: 'Charge en DCO de dimensionnement', unit: 'kg/j', group: 'Alimentation', default: undefined, hint: 'DCO des graisses admises si non forcée' },
    { key: 'entree_diluee_DCO', label: 'Concentration de la charge diluée', unit: 'g DCO/L', group: 'Alimentation', default: 75 },
    { key: 'rdt_elimination_DCO', label: "Rendement d'élimination de la DCO", unit: '-', group: 'Procédé', default: 0.7 },
    { key: 'volume_homogeneisation', label: "Volume de la cuve d'homogénéisation", unit: 'm³', group: 'Procédé', default: undefined, hint: 'un jour de temps de passage' },
    { key: 'volume_bassin', label: 'Volume du bassin biologique', unit: 'm³', group: 'Procédé', default: undefined, hint: 'vingt jours de temps de passage' },
    { key: 'agitation_homogeneisation_W_m3', label: "Puissance d'agitation de la cuve d'homogénéisation", unit: 'W/m³', group: 'Procédé', default: 25 },
    { key: 'agitation_bassin_W_m3', label: "Puissance d'agitation du bassin biologique", unit: 'W/m³', group: 'Procédé', default: 25 },
    { key: 'agitation_bassin_hj', label: "Durée d'agitation du bassin biologique", unit: 'h/j', group: 'Procédé', default: 24 },
    { key: 'besoin_O2', label: 'Besoin en oxygène', unit: 'kg O2/j', group: 'Aération', default: undefined, hint: '1 kg O2 par kg de DCO éliminée' },
    { key: 'debit_air', label: "Débit d'air process", unit: 'Nm³/j', group: 'Aération', default: undefined, hint: 'calculé si non forcé' },
    { key: 'hauteur_bassin', label: "Hauteur d'eau du bassin", unit: 'm', group: 'Aération', default: 6 },
    { key: 'rdt_transfert', label: "Rendement de transfert de l'O2 en eau claire", unit: '%/m', group: 'Aération', default: undefined, hint: '4,625 %/m' },
    { key: 'air_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Aération', default: undefined, hint: "hauteur d'eau + 1" },
    { key: 'surpresseur_conso_spec', label: 'Consommation spécifique des surpresseurs', unit: 'Wh/(Nm³·mCE)', group: 'Aération', default: (c) => HYP.surpresseur_conso_spec_Wh_Nm3mCE[c.choices.surpresseur] },
    { key: 'consommation_N', label: "Consommation d'azote", unit: 'kg N/j', group: 'Nutriments', default: undefined, hint: '25 g N par kg de DCO' },
    { key: 'consommation_P', label: 'Consommation de phosphore', unit: 'kg P/j', group: 'Nutriments', default: undefined, hint: '5 g P par kg de DCO' },
    { key: 'boues_MES', label: 'Boues produites', unit: 'kg MES/j', group: 'Boues', default: undefined, hint: '0,3 kg MVS par kg de DCO éliminée' },
    { key: 'boues_MV_MES', label: 'MV/MES des boues', unit: '-', group: 'Boues', default: 0.76 },
    { key: 'boues_concentration', label: 'Concentration des boues', unit: 'g/L', group: 'Boues', default: undefined, hint: 'déduite du débit de sortie' },
    { key: 'pompage_nb', label: 'Nombre de pompes de boues', unit: 'u', group: 'Pompages', default: 2 },
    { key: 'pompage_P_refoulement', label: 'Pression de refoulement des pompes de boues', unit: 'mCE', group: 'Pompages', default: 5 },
    { key: 'pompage_tps_fonctionnement', label: 'Durée de fonctionnement des pompes de boues', unit: 'h/j', group: 'Pompages', default: 6 },
    { key: 'pompage_Q_unitaire', label: 'Débit unitaire des pompes de boues', unit: 'm³/h', group: 'Pompages', default: undefined, hint: 'calculé si non forcé' },
  ],

  compute(ctx) {
    const { site, table, soluble, choices, forced, p } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const etapeOut = ETAPE.graisses_biolix
    const j = TYPE.graisses

    // ---- calcul_cocktail_boues_entrees : seules les graisses sont admises
    const part = p.part_flux_admis ?? 1
    const src = table[ETAPE.inlet][j]
    const sol = soluble[ETAPE.inlet][j]
    const MES = src[P.MES] * part
    const Q = src[P.Q] * part
    if (!(MES > 0) || !(Q > 0)) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Aucune graisse en entrée : le Biolix suppose un dégraisseur en amont qui extrait des graisses vers la file boues."],
      }
    }
    const MV = src[P.MV_MES] * MES
    // la pollution soluble des graisses s'ajoute à la pollution particulaire
    const DCO = src[P.ratio_DCO_MES] * MES + (sol[SOL.DCO] * Q) / 1000
    const DBO = src[P.ratio_DBO_MES] * MES
    const NK = src[P.ratio_NK_MES] * MES + (sol[SOL.NK] * Q) / 1000
    const Pt = src[P.ratio_Pt_MES] * MES + (sol[SOL.Pt] * Q) / 1000

    // ---- attribution_valeur_par_defaut
    const rdt_DCO = p.rdt_elimination_DCO ?? H.rdt_elimination_DCO_defaut
    const boues_MV_MES = p.boues_MV_MES ?? H.boues_MV_MES_defaut
    const hauteur = p.hauteur_bassin ?? H.hauteur_bassin_defaut
    const rdt_transfert = f('rdt_transfert') ?? H.rdt_transfert_defaut
    const air_P_refoulement = f('air_P_refoulement') ?? hauteur + 1
    const conc_diluee = p.entree_diluee_DCO ?? H.entree_diluee_DCO_gL

    // ---- dilution : le débit traité découle de la charge en DCO
    const DCO_dim = f('DCO_dimensionnement') ?? DCO
    const DCO_traitee = f('DCO_dimensionnement') != null ? DCO_dim : DCO
    let Q_traite = conc_diluee > 0 ? DCO_dim / conc_diluee : 0
    if (Q_traite < Q) Q_traite = Q
    const eau_dilution = Math.max(0, Q_traite - Q)

    // ---- volumes, sur les temps de passage hydrauliques
    const avecHomog = choices.agitation_homogeneisation === 'oui'
    const avecAgitBassin = choices.agitation_bassin === 'oui'
    const volume_homogeneisation = avecHomog
      ? (f('volume_homogeneisation') ?? Q_traite * H.HRT_homogeneisation_j)
      : 0
    const volume_bassin = f('volume_bassin') ?? Q_traite * H.HRT_bassin_biologique_j

    // ---- calcul_aeration_et_nutriments
    const DCO_eliminee = rdt_DCO * DCO_traitee
    const besoin_O2 = f('besoin_O2') ?? H.besoin_O2_kgO2_kgDCO * DCO_eliminee
    const denom = H.aeration_facteur_K * (rdt_transfert / 100) * (hauteur - H.hauteur_diffuseur) * CONST.QUANTITE_OXYGENE_DANS_AIR_kgO2_Nm3
    const debit_air = f('debit_air') ?? (denom > 0 ? besoin_O2 / denom : 0)
    // les graisses sont pauvres en azote et en phosphore : il faut en apporter
    const consommation_N = f('consommation_N') ?? H.besoin_N_kgN_kgDCO * DCO_traitee
    const consommation_P = f('consommation_P') ?? H.besoin_P_kgP_kgDCO * DCO_traitee

    // ---- calcul_boues_produites
    const boues_MES = f('boues_MES')
      ?? (boues_MV_MES > 0 ? (H.production_boues_kgMVS_kgDCO * DCO_eliminee) / boues_MV_MES : 0)
    const boues_concentration = f('boues_concentration') ?? (Q_traite > 0 ? boues_MES / Q_traite : 0)
    const boues_Q = boues_concentration > 0 ? boues_MES / boues_concentration : Q_traite

    // ---- écriture de l'étape aval
    const dst = table[etapeOut][j]
    dst[P.origine] = src[P.origine] || 'codigestion_graisses'
    dst[P.MES] = boues_MES
    dst[P.Q] = boues_Q
    dst[P.MV_MES] = boues_MV_MES
    dst[P.ratio_DCO_MES] = boues_MES > 0 ? ((1 - rdt_DCO) * DCO_traitee) / boues_MES : 0
    dst[P.ratio_DBO_MES] = boues_MES > 0 ? ((1 - rdt_DCO) * DBO) / boues_MES : 0
    // l'azote et le phosphore ajoutés se retrouvent dans les boues
    dst[P.ratio_NK_MES] = boues_MES > 0 ? (NK + consommation_N) / boues_MES : 0
    dst[P.ratio_Pt_MES] = boues_MES > 0 ? (Pt + consommation_P) / boues_MES : 0
    dst[P.flux_in] = src[P.flux_in] * part
    // Le classeur dirige toujours la totalité des boues Biolix vers le
    // traitement secondaire : le flux est donc entièrement consommé ici.
    dst[P.verif_flux] = 0
    for (let k = 1; k <= 5; k++) soluble[etapeOut][j][k] = 0
    src[P.verif_flux] -= src[P.flux_in] * part
    if (!ctx.site?.choix_secondaire && ctx.site?.choix_secondaire !== undefined) {
      warnings.push("Les boues Biolix sont dirigées vers le traitement secondaire, mais la filière n'en comporte pas.")
    }

    // ---- calcul_consommation_electrique
    const conso_spec = f('surpresseur_conso_spec') ?? HYP.surpresseur_conso_spec_Wh_Nm3mCE[choices.surpresseur]
    if (choices.surpresseur === 'roots' && air_P_refoulement > HYP.surpresseur_Px2) {
      warnings.push(`Pression de refoulement (${air_P_refoulement.toFixed(1)} mCE) élevée pour des surpresseurs roots.`)
    }
    const electricite_aeration = (debit_air * air_P_refoulement * conso_spec) / 1000
    const electricite_homogeneisation = avecHomog
      ? (H.agitation_homogeneisation_hj * volume_homogeneisation * (p.agitation_homogeneisation_W_m3 ?? H.agitation_W_m3_defaut)) / 1000
      : 0
    const electricite_agitation_bassin = avecAgitBassin
      ? ((p.agitation_bassin_hj ?? 24) * volume_bassin * (p.agitation_bassin_W_m3 ?? H.agitation_W_m3_defaut)) / 1000
      : 0
    const pompage = repartitionPompage(boues_Q, p.pompage_nb, p.pompage_tps_fonctionnement, forced, 'pompage', CONST.NOMBRE_HEURE_PAR_JOUR)
    if (pompage.incoherence) warnings.push('Incohérence sur le pompage des boues Biolix.')
    const rdt_pompe = rdtPompeBoues(pompage.Qu, p.pompage_P_refoulement)
    const electricite_pompage = elecPompage(pompage.Qu, pompage.nb, pompage.tps, p.pompage_P_refoulement, rdt_pompe)

    const total = electricite_aeration + electricite_homogeneisation + electricite_agitation_bassin + electricite_pompage

    return {
      etapeSortie: etapeOut,
      results: [
        { key: 'in_MES', label: 'Graisses admises', unit: 'kg MES/j', value: MES },
        { key: 'in_Q', label: 'Débit de graisses admises', unit: 'm³/j', value: Q },
        { key: 'in_DCO', label: 'DCO des graisses admises', unit: 'kg/j', value: DCO },
        { key: 'in_MV_MES', label: 'MV/MES des graisses', unit: '-', value: MES > 0 ? MV / MES : 0 },
        { key: 'dilution', label: "Eau de dilution", unit: 'm³/j', value: eau_dilution },
        { key: 'Q_traite', label: 'Débit traité après dilution', unit: 'm³/j', value: Q_traite },
        { key: 'conc_diluee', label: 'Concentration de la charge diluée', unit: 'g DCO/L', value: conc_diluee },
        ...(avecHomog ? [{ key: 'V_homog', label: "Volume de la cuve d'homogénéisation", unit: 'm³', value: volume_homogeneisation }] : []),
        { key: 'V_bassin', label: 'Volume du bassin biologique', unit: 'm³', value: volume_bassin },
        { key: 'HRT', label: 'Temps de passage dans le bassin biologique', unit: 'j', value: Q_traite > 0 ? volume_bassin / Q_traite : 0 },
        { key: 'rdt_DCO', label: "Rendement d'élimination de la DCO", unit: '-', value: rdt_DCO },
        { key: 'DCO_elim', label: 'DCO éliminée', unit: 'kg/j', value: DCO_eliminee },
        { key: 'O2', label: 'Besoin en oxygène', unit: 'kg O2/j', value: besoin_O2 },
        { key: 'air', label: "Débit d'air process", unit: 'Nm³/h', value: debit_air / CONST.NOMBRE_HEURE_PAR_JOUR },
        { key: 'N', label: "Azote à apporter", unit: 'kg N/j', value: consommation_N },
        { key: 'P', label: 'Phosphore à apporter', unit: 'kg P/j', value: consommation_P },
        { key: 'boues', label: 'Boues Biolix produites', unit: 'kg MES/j', value: boues_MES },
        { key: 'boues_Q', label: 'Débit de boues Biolix', unit: 'm³/j', value: boues_Q },
        { key: 'boues_conc', label: 'Concentration des boues Biolix', unit: 'g/L', value: boues_concentration },
        { key: 'destination', label: 'Boues dirigées vers le traitement secondaire', unit: '-', value: 1 },
      ],
      reactifs: {
        ...(consommation_N > 0 ? { azote: consommation_N } : {}),
        ...(consommation_P > 0 ? { phosphore: consommation_P } : {}),
      },
      electricity: {
        total,
        fixed: electricite_homogeneisation + electricite_agitation_bassin,
        detail: {
          aeration: electricite_aeration,
          agitation_homogeneisation: electricite_homogeneisation,
          agitation_bassin: electricite_agitation_bassin,
          pompage: electricite_pompage,
        },
      },
      warnings,
    }
  },
})
