// ---------------------------------------------------------------------------
// Port de Gestion_OPEX.cls et Retour_investissement.cls.
//
// Le volet économique clôt le bloc transverse. Il ne calcule rien de nouveau :
// il valorise ce que les trois files ont déjà produit.
//
//   OPEX  : électricité achetée, réactifs, combustibles, évacuation des sous-
//           produits, moins les recettes de vente d'électricité, de biogaz et
//           de struvite. La distinction fixe/variable du bilan électrique est
//           conservée, parce qu'elle ne se comporte pas de la même façon
//           lorsque la charge évolue.
//   ROI   : comparaison de plusieurs variantes de filière par rapport à une
//           référence. Trois indicateurs : le temps de retour simple, le CAPEX
//           majoré de N années d'OPEX, et la valeur actuelle nette année par
//           année — c'est cette dernière série qui alimente le graphique.
//
// La convention du classeur sur le temps de retour mérite d'être explicitée :
// une variante ne « rembourse » que si elle coûte **moins cher à exploiter** que
// la référence. Si son OPEX est supérieur ou égal, aucun temps de retour n'est
// calculé, quel que soit l'écart d'investissement.
//
// Écarts au VBA, volontaires et signalés (voir README) :
//  - le classeur gère deux devises en parallèle (devise1 et devise2) avec un
//    taux de change ; le port n'en tient qu'une, la conversion étant une
//    multiplication triviale à appliquer en aval ;
//  - les prix par défaut sont lus depuis une feuille du classeur, non depuis le
//    VBA : les valeurs retenues ici sont des ordres de grandeur européens de
//    2024, tous forçables, et signalés comme tels ;
//  - la répartition des réactifs entre file eau, file boues et désodorisation
//    est faite par réactif dans le classeur ; le port l'établit par procédé
//    consommateur, ce qui est plus juste quand un même réactif sert aux deux
//    files.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { CONST } from '../core/hypotheses.js'

// Prix unitaires par défaut, en euros. Ordres de grandeur européens ; le
// classeur les lit depuis une feuille de saisie, ils sont donc tous forçables.
const PRIX = {
  electricite_achat: 0.15, // €/kWh
  electricite_vente: 0.08,
  biogaz_vente: 0.06, // €/kWh PCI
  gaz_naturel: 0.05, // €/kWh PCI
  fioul: 0.09,
  evacuation_boues: 45, // €/t
  evacuation_graisses: 90,
  evacuation_cendres: 60,
  evacuation_REFIB: 250, // déchet dangereux
  evacuation_refus: 120,
  struvite_vente: 250, // €/t
}
// Prix des réactifs en produit commercial, €/t.
const PRIX_REACTIFS = {
  polymere: 3200, FeCl3: 250, methanol: 450, chaux_eteinte: 150, chaux_vive: 130,
  H2SO4: 180, NaOH: 400, soude: 400, NaOCl: 300, NaHSO3: 350, NaHCO3: 450,
  bicarbonate_sodium: 450, Ca_2NO3: 300, nitrate_calcium: 300,
  oxygene_liquide: 120, H2O2: 600, charbon_actif: 2500, Cl2: 500,
  ammoniaque: 350, uree: 400, CuSO4: 1800, MgCl2: 280,
  azote: 350, phosphore: 900,
}
const LIBELLE_REACTIF = {
  polymere: 'Polymère', FeCl3: 'Chlorure ferrique', methanol: 'Méthanol',
  chaux_eteinte: 'Chaux éteinte', chaux_vive: 'Chaux vive', H2SO4: 'Acide sulfurique',
  NaOH: 'Soude', soude: 'Soude', NaOCl: 'Hypochlorite', NaHSO3: 'Bisulfite de sodium',
  NaHCO3: 'Bicarbonate de sodium', bicarbonate_sodium: 'Bicarbonate de sodium',
  Ca_2NO3: 'Nitrate de calcium', nitrate_calcium: 'Nitrate de calcium',
  oxygene_liquide: 'Oxygène liquide', H2O2: "Peroxyde d'hydrogène",
  charbon_actif: 'Charbon actif', Cl2: 'Chlore', ammoniaque: 'Ammoniaque',
  uree: 'Urée', CuSO4: 'Sulfate de cuivre', MgCl2: 'Chlorure de magnésium',
  azote: 'Azote (nutriment)', phosphore: 'Phosphore (nutriment)',
}

export const gestionOpex = defineUtilityNode({
  id: 'gestion-opex',
  label: 'Coûts d\'exploitation',
  short: 'OPEX',
  family: 'transverse',
  vba: 'Gestion_OPEX.cls',
  description:
    "Valorise les consommations des trois files : électricité, réactifs, combustibles et évacuation des sous-produits, moins les recettes de vente. La distinction entre coûts fixes et variables est conservée.",
  choices: [
    { key: 'valorisation_biogaz', label: 'Valorisation du biogaz', default: 'cogeneration', options: [
      { value: 'cogeneration', label: 'cogénération sur site' },
      { value: 'vendu', label: 'injecté ou vendu' },
    ] },
  ],
  params: [
    { key: 'prix_electricite_achat', label: "Prix d'achat de l'électricité", unit: '€/kWh', group: 'Énergie', default: 0.15 },
    { key: 'prix_electricite_vente', label: "Prix de vente de l'électricité", unit: '€/kWh', group: 'Énergie', default: 0.08 },
    { key: 'prix_biogaz_vente', label: 'Prix de vente du biogaz', unit: '€/kWh PCI', group: 'Énergie', default: 0.06 },
    { key: 'prix_gaz_naturel', label: 'Prix du gaz naturel', unit: '€/kWh PCI', group: 'Énergie', default: 0.05 },
    { key: 'prix_fioul', label: 'Prix du fioul', unit: '€/kWh PCI', group: 'Énergie', default: 0.09 },
    { key: 'prix_evacuation_boues', label: 'Coût d\'évacuation des boues', unit: '€/t', group: 'Sous-produits', default: 45 },
    { key: 'prix_evacuation_graisses', label: "Coût d'évacuation des graisses", unit: '€/t', group: 'Sous-produits', default: 90 },
    { key: 'prix_evacuation_cendres', label: "Coût d'évacuation des cendres", unit: '€/t', group: 'Sous-produits', default: 60 },
    { key: 'prix_evacuation_REFIB', label: "Coût d'évacuation des REFIB", unit: '€/t', group: 'Sous-produits', default: 250 },
    { key: 'prix_struvite_vente', label: 'Prix de vente de la struvite', unit: '€/t', group: 'Sous-produits', default: 250 },
    { key: 'depense_autre_fixe', label: 'Autres dépenses fixes', unit: '€/j', group: 'Compléments', default: 0 },
    { key: 'depense_autre_variable', label: 'Autres dépenses variables', unit: '€/j', group: 'Compléments', default: 0 },
    { key: 'benefice_autre', label: 'Autres recettes', unit: '€/j', group: 'Compléments', default: 0 },
    ...Object.entries(PRIX_REACTIFS)
      .filter(([k]) => !['soude', 'bicarbonate_sodium', 'nitrate_calcium'].includes(k))
      .map(([k, v]) => ({
        key: `prix_${k}`,
        label: `Prix — ${LIBELLE_REACTIF[k] ?? k}`,
        unit: '€/t commercial',
        group: 'Prix des réactifs',
        default: v,
      })),
  ],

  compute(ctx) {
    const { contexte, p, forced, choices } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const bilan = contexte.bilanElectrique
    if (!bilan) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Les coûts d'exploitation doivent être calculés après le bilan électrique."] }
    }
    const an = CONST.NOMBRE_JOUR_PAR_AN
    const prixDe = (cle, defaut) => f(cle) ?? p[cle] ?? defaut

    // ---- électricité : le prix ne s'applique qu'à l'électricité achetée
    const prix_achat = prixDe('prix_electricite_achat', PRIX.electricite_achat)
    const achetee = Math.max(0, bilan.consommee - bilan.verte_consommee)
    // la part fixe et la part variable sont valorisées au prorata de l'achat
    const part_achetee = bilan.consommee > 0 ? achetee / bilan.consommee : 0
    const cout_elec_fixe = bilan.fixe * part_achetee * prix_achat
    const cout_elec_variable = bilan.variable * part_achetee * prix_achat
    const cout_electricite = cout_elec_fixe + cout_elec_variable

    // répartition du coût électrique par file, depuis les postes du bilan
    const coutParFile = { eau: 0, boues: 0, utilites: 0 }
    for (const poste of bilan.postes ?? []) {
      const file = poste.file ?? 'utilites'
      coutParFile[file] = (coutParFile[file] || 0) + Math.max(0, poste.total) * part_achetee * prix_achat
    }

    // ---- réactifs, en produit commercial
    const lignesReactifs = []
    let cout_reactifs = 0
    for (const [cle, kgj] of Object.entries(contexte.reactifs ?? {})) {
      if (!(kgj > 0)) continue
      const prix = prixDe(`prix_${cle}`, PRIX_REACTIFS[cle])
      if (prix == null) {
        warnings.push(`Réactif « ${cle} » sans prix de référence : non valorisé.`)
        continue
      }
      const tonnes_an = (kgj / 1000) * an
      const cout_j = (kgj / 1000) * prix
      cout_reactifs += cout_j
      lignesReactifs.push({ cle, label: LIBELLE_REACTIF[cle] ?? cle, kgj, tonnes_an, prix, cout_j, cout_an: cout_j * an })
    }
    lignesReactifs.sort((a, b) => b.cout_j - a.cout_j)

    // ---- combustibles
    const cout_gaz = (contexte.gaz_naturel_kWhPCIj ?? 0) * prixDe('prix_gaz_naturel', PRIX.gaz_naturel)
    const cout_fioul = (contexte.fioul_kWhPCIj ?? 0) * prixDe('prix_fioul', PRIX.fioul)
    const cout_combustibles = cout_gaz + cout_fioul

    // ---- évacuation des sous-produits
    const boues_Tj = contexte.boues_evacuees_MES > 0 ? contexte.boues_evacuees_Q : 0
    const cout_boues = boues_Tj * prixDe('prix_evacuation_boues', PRIX.evacuation_boues)
    const cout_cendres = (contexte.cendres_Tj ?? 0) * prixDe('prix_evacuation_cendres', PRIX.evacuation_cendres)
    const cout_REFIB = (contexte.REFIB_Tj ?? 0) * prixDe('prix_evacuation_REFIB', PRIX.evacuation_REFIB)
    const cout_graisses = (contexte.graisses_Tj ?? 0) * prixDe('prix_evacuation_graisses', PRIX.evacuation_graisses)
    const cout_evacuation = cout_boues + cout_cendres + cout_REFIB + cout_graisses

    // ---- recettes
    const benefice_electricite = bilan.verte_vendue * prixDe('prix_electricite_vente', PRIX.electricite_vente)
    const biogaz_vendu_kWh = choices.valorisation_biogaz === 'vendu'
      ? contexte.biogaz_Nm3j * (contexte.biogaz_CH4 || 0.63) * 9.94
      : 0
    const benefice_biogaz = biogaz_vendu_kWh * prixDe('prix_biogaz_vente', PRIX.biogaz_vente)
    const benefice_struvite = ((contexte.struvite_kgj ?? 0) / 1000) * prixDe('prix_struvite_vente', PRIX.struvite_vente)
    const benefices = benefice_electricite + benefice_biogaz + benefice_struvite + (p.benefice_autre ?? 0)

    const cout_fixe = cout_elec_fixe + cout_combustibles + (p.depense_autre_fixe ?? 0)
    const cout_variable = cout_elec_variable + cout_reactifs + cout_evacuation + (p.depense_autre_variable ?? 0)
    const cout_total = cout_fixe + cout_variable
    const opex_net = cout_total - benefices

    const Q = contexte.Q_reel
    const ratio_Q = Q > 0 ? opex_net / Q : 0
    const capacite_EH = bilan.capacite_EH ?? 0
    const ratio_EH = capacite_EH > 0 ? (opex_net * an) / capacite_EH : 0

    const postes = [
      { id: 'electricite', label: 'Électricité', valeur: cout_electricite },
      { id: 'reactifs', label: 'Réactifs', valeur: cout_reactifs },
      { id: 'combustibles', label: 'Combustibles', valeur: cout_combustibles },
      { id: 'evacuation', label: 'Évacuation des sous-produits', valeur: cout_evacuation },
      { id: 'autres', label: 'Autres dépenses', valeur: (p.depense_autre_fixe ?? 0) + (p.depense_autre_variable ?? 0) },
    ].filter((x) => x.valeur > 1e-6)

    return {
      results: [
        { key: 'total_j', label: "Coûts d'exploitation", unit: '€/j', value: cout_total },
        { key: 'total_an', label: "Coûts d'exploitation", unit: 'k€/an', value: (cout_total * an) / 1000 },
        { key: 'fixe', label: 'dont coûts fixes', unit: '€/j', value: cout_fixe },
        { key: 'variable', label: 'dont coûts variables', unit: '€/j', value: cout_variable },
        { key: 'elec', label: 'Poste — électricité', unit: '€/j', value: cout_electricite },
        { key: 'reactifs', label: 'Poste — réactifs', unit: '€/j', value: cout_reactifs },
        { key: 'combustibles', label: 'Poste — combustibles', unit: '€/j', value: cout_combustibles },
        { key: 'evacuation', label: 'Poste — évacuation des sous-produits', unit: '€/j', value: cout_evacuation },
        { key: 'benefices', label: 'Recettes totales', unit: '€/j', value: benefices },
        { key: 'benef_elec', label: "dont vente d'électricité", unit: '€/j', value: benefice_electricite },
        { key: 'benef_biogaz', label: 'dont vente de biogaz', unit: '€/j', value: benefice_biogaz },
        { key: 'benef_struvite', label: 'dont vente de struvite', unit: '€/j', value: benefice_struvite },
        { key: 'net_j', label: 'OPEX net', unit: '€/j', value: opex_net },
        { key: 'net_an', label: 'OPEX net', unit: 'k€/an', value: (opex_net * an) / 1000 },
        { key: 'ratio_Q', label: 'OPEX rapporté au débit', unit: '€/m³', value: ratio_Q },
        { key: 'ratio_EH', label: 'OPEX par équivalent habitant', unit: '€/(EH·an)', value: ratio_EH },
        { key: 'elec_eau', label: 'Électricité — file eau', unit: '€/j', value: coutParFile.eau },
        { key: 'elec_boues', label: 'Électricité — file boues', unit: '€/j', value: coutParFile.boues },
        { key: 'elec_utilites', label: 'Électricité — utilités', unit: '€/j', value: coutParFile.utilites },
      ],
      opex: {
        cout_total, cout_fixe, cout_variable, benefices, opex_net,
        opex_net_an: opex_net * an,
        ratio_Q, ratio_EH, postes, lignesReactifs, coutParFile,
        detail_benefices: { electricite: benefice_electricite, biogaz: benefice_biogaz, struvite: benefice_struvite },
      },
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})

// ---------------------------------------------------------------------------
export const retourInvestissement = defineUtilityNode({
  id: 'retour-investissement',
  label: 'Retour sur investissement',
  short: 'ROI',
  family: 'transverse',
  vba: 'Retour_investissement.cls',
  description:
    "Compare la filière simulée à une variante de référence : temps de retour simple, coût complet sur la durée du contrat, et valeur actuelle nette année par année.",
  choices: [],
  params: [
    { key: 'capex', label: 'Investissement de la filière simulée', unit: 'k€', group: 'Filière simulée', default: 0 },
    { key: 'capex_reference', label: 'Investissement de la référence', unit: 'k€', group: 'Référence', default: 0 },
    { key: 'opex_reference', label: 'OPEX annuel de la référence', unit: 'k€/an', group: 'Référence', default: undefined, hint: "OPEX de la filière simulée si non forcé" },
    { key: 'duree_contrat', label: 'Durée du contrat', unit: 'an', group: 'Hypothèses', default: 20 },
    { key: 'taux_actualisation', label: "Taux d'actualisation", unit: '-', group: 'Hypothèses', default: 0.04 },
  ],

  compute(ctx) {
    const { contexte, p, forced } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
    const opex = contexte.opex
    if (!opex) {
      return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ["Le retour sur investissement doit être calculé après les coûts d'exploitation."] }
    }

    const capex = p.capex ?? 0
    const capex_ref = p.capex_reference ?? 0
    const opex_an = opex.opex_net_an / 1000 // k€/an
    const opex_ref_an = f('opex_reference') ?? opex_an
    const duree = Math.max(1, Math.round(p.duree_contrat ?? 20))
    const taux = p.taux_actualisation ?? 0.04

    if (capex === 0 && capex_ref === 0) {
      warnings.push("Aucun investissement renseigné : saisir les CAPEX de la filière et de la référence pour comparer.")
    }

    // ---- temps de retour simple
    // Convention du classeur : une variante ne rembourse que si elle coûte
    // moins cher à exploiter que la référence.
    const economie_an = opex_ref_an - opex_an
    let tps_retour = null
    if (economie_an > 0) tps_retour = (capex - capex_ref) / economie_an
    else warnings.push("La filière simulée coûte autant ou plus à exploiter que la référence : aucun temps de retour ne peut être calculé.")

    // ---- coût complet sur la durée du contrat
    const capex_x_opex = capex + duree * opex_an
    const capex_x_opex_ref = capex_ref + duree * opex_ref_an

    // ---- valeur actuelle nette, année par année
    const serie = [capex_ref - capex]
    let tps_retour_actualise = null
    for (let annee = 1; annee <= duree; annee++) {
      const precedent = serie[annee - 1]
      const valeur = precedent + economie_an / Math.pow(1 + taux, annee)
      serie.push(valeur)
      // franchissement de zéro : on interpole entre les deux années
      if (valeur >= 0 && precedent < 0 && tps_retour_actualise == null) {
        tps_retour_actualise = annee - 1 + Math.abs(precedent) / (valeur + Math.abs(precedent))
      }
    }
    const van = serie[duree]

    return {
      results: [
        { key: 'capex', label: 'Investissement de la filière', unit: 'k€', value: capex },
        { key: 'capex_ref', label: 'Investissement de la référence', unit: 'k€', value: capex_ref },
        { key: 'opex', label: 'OPEX de la filière', unit: 'k€/an', value: opex_an },
        { key: 'opex_ref', label: 'OPEX de la référence', unit: 'k€/an', value: opex_ref_an },
        { key: 'economie', label: 'Économie annuelle', unit: 'k€/an', value: economie_an },
        ...(tps_retour != null ? [{ key: 'tps', label: 'Temps de retour simple', unit: 'an', value: tps_retour }] : []),
        ...(tps_retour_actualise != null ? [{ key: 'tps_act', label: 'Temps de retour actualisé', unit: 'an', value: tps_retour_actualise }] : []),
        { key: 'cout_complet', label: `Coût complet sur ${duree} ans — filière`, unit: 'k€', value: capex_x_opex },
        { key: 'cout_complet_ref', label: `Coût complet sur ${duree} ans — référence`, unit: 'k€', value: capex_x_opex_ref },
        { key: 'van', label: `Valeur actuelle nette à ${duree} ans`, unit: 'k€', value: van },
      ],
      roi: {
        capex, capex_ref, opex_an, opex_ref_an, economie_an, duree, taux,
        tps_retour, tps_retour_actualise, van,
        capex_x_opex, capex_x_opex_ref,
        serie,
      },
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})

export { PRIX, PRIX_REACTIFS, LIBELLE_REACTIF }
