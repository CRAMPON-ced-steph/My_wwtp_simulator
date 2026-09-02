// ---------------------------------------------------------------------------
// Port de Turbine_hydraulique.cls et Gestion_reactifs.cls.
//
// Deux utilités courtes, réunies ici parce qu'elles partagent le même rôle :
// elles ne transforment rien, elles convertissent. La turbine convertit une
// chute d'eau en électricité, la gestion des réactifs convertit des flux de
// produit pur en flux de produit commercial.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { CONST } from '../core/hypotheses.js'

// ---------------------------------------------------------------------------
// Turbine hydraulique
// ---------------------------------------------------------------------------
const HT = {
  rho_eau: 1000, // kg/m³
  rendement_turbine: 0.93,
  rendement_generateur: 0.97,
  // rendement global constaté à 70 % de charge, plus faible que le produit des
  // deux rendements nominaux : la moyenne des deux sert de valeur de référence
  rendement_global_70: 0.67,
  duree_fonctionnement_defaut: 8500, // h/an
  duree_fonctionnement_max: 8760,
  hauteur_defaut: 5, // m
  perte_charge_defaut: 1, // mCE
}

export const turbineHydraulique = defineUtilityNode({
  id: 'turbine-hydraulique',
  label: 'Turbine hydraulique',
  short: 'Turbine',
  family: 'production',
  vba: 'Turbine_hydraulique.cls',
  description:
    "Récupération d'énergie sur la chute disponible en sortie de station. La puissance nette retranche les pertes de charge de la hauteur brute ; le rendement global retenu tient compte du fonctionnement à charge partielle.",
  choices: [],
  params: [
    { key: 'pourcentage_debit', label: 'Part du débit turbinée', unit: '-', group: 'Hydraulique', default: 1 },
    { key: 'hauteur_totale', label: 'Hauteur totale de chute', unit: 'm', group: 'Hydraulique', default: 5 },
    { key: 'perte_de_charge', label: 'Perte de charge hydraulique', unit: 'mCE', group: 'Hydraulique', default: 1 },
    { key: 'puissance_nette', label: 'Puissance nette', unit: 'kW', group: 'Hydraulique', default: undefined, hint: 'calculée sur la hauteur nette' },
    { key: 'rendement_global', label: 'Rendement global moyen', unit: '-', group: 'Machine', default: undefined, hint: 'moyenne entre le rendement nominal et celui à 70 % de charge' },
    { key: 'duree_fonctionnement', label: 'Durée de fonctionnement', unit: 'h/an', group: 'Machine', default: 8500, hint: 'plafonnée à 8 760 h' },
  ],

  compute(ctx) {
    const { contexte, p, forced } = ctx
    const warnings = []
    const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)

    const part = p.pourcentage_debit ?? 1
    const hauteur = p.hauteur_totale ?? HT.hauteur_defaut
    const perte = p.perte_de_charge ?? HT.perte_charge_defaut
    const rendement = f('rendement_global')
      ?? (HT.rendement_global_70 + HT.rendement_generateur * HT.rendement_turbine) / 2
    let duree = p.duree_fonctionnement ?? HT.duree_fonctionnement_defaut
    if (duree > HT.duree_fonctionnement_max) {
      duree = HT.duree_fonctionnement_max
      warnings.push(`Durée de fonctionnement plafonnée à ${HT.duree_fonctionnement_max} h/an.`)
    }
    if (perte >= hauteur) {
      warnings.push(`Les pertes de charge (${perte} mCE) atteignent ou dépassent la hauteur de chute (${hauteur} m) : aucune énergie récupérable.`)
    }

    // le débit turbiné suit la charge réelle, pas le débit nominal
    const Q_m3h = (contexte.Q_reel * part) / CONST.NOMBRE_HEURE_PAR_JOUR
    const debit_m3s = Q_m3h / CONST.NOMBRE_SECONDE_PAR_HEURE
    const puissance_brute = (HT.rho_eau * CONST.ACCELERATION_PESANTEUR_m_s2 * debit_m3s * hauteur) / 1000
    const puissance_nette = f('puissance_nette')
      ?? (HT.rho_eau * CONST.ACCELERATION_PESANTEUR_m_s2 * debit_m3s * Math.max(0, hauteur - perte)) / 1000
    const puissance_electrique = puissance_nette * rendement
    const production_annuelle = duree * puissance_electrique
    const production_kWhj = production_annuelle / CONST.NOMBRE_JOUR_PAR_AN

    return {
      results: [
        { key: 'Q', label: 'Débit turbiné', unit: 'm³/h', value: Q_m3h },
        { key: 'H', label: 'Hauteur de chute nette', unit: 'm', value: Math.max(0, hauteur - perte) },
        { key: 'P_brute', label: 'Puissance hydraulique brute', unit: 'kW', value: puissance_brute },
        { key: 'P_nette', label: 'Puissance hydraulique nette', unit: 'kW', value: puissance_nette },
        { key: 'rdt', label: 'Rendement global moyen', unit: '-', value: rendement },
        { key: 'P_elec', label: 'Puissance électrique moyenne', unit: 'kW', value: puissance_electrique },
        { key: 'prod_an', label: 'Production annuelle', unit: 'kWh/an', value: production_annuelle },
        { key: 'prod', label: 'Production journalière moyenne', unit: 'kWh/j', value: production_kWhj },
      ],
      // une production est une consommation négative dans le bilan électrique
      electricity: { total: -production_kWhj, produite: production_kWhj, fixed: 0, detail: { production: -production_kWhj } },
      warnings,
    }
  },
})

// ---------------------------------------------------------------------------
// Gestion des réactifs
// ---------------------------------------------------------------------------
// Pureté du produit commercial, par réactif. Un flux de produit pur se convertit
// en flux commercial par simple division — sauf l'hypochlorite, dont le titre
// s'exprime en degrés chlorométriques et non en fraction massique.
const REACTIFS = {
  polymere: { label: 'Polymère', purete: 0.98, note: 'poudre ; 0,5 en émulsion liquide' },
  FeCl3: { label: 'Chlorure ferrique', purete: 0.41 },
  methanol: { label: 'Méthanol', purete: 1 },
  chaux_eteinte: { label: 'Chaux éteinte', purete: 0.9 },
  chaux_vive: { label: 'Chaux vive', purete: 0.95 },
  H2SO4: { label: 'Acide sulfurique', purete: 0.98, note: '98 % en France, 96 % à l\'étranger' },
  NaOH: { label: 'Soude', purete: 0.3 },
  NaOCl: { label: 'Hypochlorite de sodium', purete: 48, beaume: true },
  NaHSO3: { label: 'Bisulfite de sodium', purete: 0.35 },
  NaHCO3: { label: 'Bicarbonate de sodium', purete: 1 },
  bicarbonate_sodium: { label: 'Bicarbonate de sodium', purete: 1 },
  Ca_2NO3: { label: 'Nitrate de calcium', purete: 0.45 },
  nitrate_calcium: { label: 'Nitrate de calcium', purete: 0.45 },
  oxygene_liquide: { label: 'Oxygène liquide', purete: 1 },
  H2O2: { label: 'Peroxyde d\'hydrogène', purete: 0.3 },
  charbon_actif: { label: 'Charbon actif', purete: 1, note: 'en poudre' },
  CuSO4: { label: 'Sulfate de cuivre', purete: 0.98 },
  MgCl2: { label: 'Chlorure de magnésium', purete: 0.3 },
  Cl2: { label: 'Chlore gazeux', purete: 1 },
  ammoniaque: { label: 'Ammoniaque', purete: 0.25, note: 'en solution' },
  uree: { label: 'Urée', purete: 0.45, note: 'en solution' },
  azote: { label: 'Azote (nutriment)', purete: 1 },
  phosphore: { label: 'Phosphore (nutriment)', purete: 1 },
  soude: { label: 'Soude', purete: 0.3 },
}
// conversion du titre en degrés chlorométriques en fraction massique
const BEAUME_VERS_MASSIQUE = 0.152 / 48

export const gestionReactifs = defineUtilityNode({
  id: 'gestion-reactifs',
  label: 'Gestion des réactifs',
  short: 'Réactifs',
  family: 'reactifs',
  vba: 'Gestion_reactifs.cls',
  description:
    "Consolide les consommations de réactifs des deux files et les convertit en produit commercial selon la pureté de chaque produit.",
  choices: [],
  params: Object.entries(REACTIFS)
    .filter(([k]) => !['bicarbonate_sodium', 'nitrate_calcium', 'soude'].includes(k))
    .map(([k, r]) => ({
      key: `purete_${k}`,
      label: `Pureté — ${r.label}`,
      unit: r.beaume ? '°chl' : '-',
      group: 'Puretés commerciales',
      default: r.purete,
      hint: r.note,
    })),

  compute(ctx) {
    const { contexte, p, forced } = ctx
    const warnings = []
    const consommations = contexte.reactifs ?? {}
    if (!Object.keys(consommations).length) {
      return {
        results: [], electricity: { total: 0, fixed: 0, detail: {} },
        warnings: ["Aucun réactif consommé par les filières : rien à consolider."],
      }
    }

    const lignes = []
    let total_pur = 0
    let total_commercial = 0
    for (const [cle, pur] of Object.entries(consommations)) {
      if (!(pur > 0)) continue
      const ref = REACTIFS[cle]
      if (!ref) {
        warnings.push(`Réactif « ${cle} » sans pureté de référence : compté en produit pur.`)
        lignes.push({ cle, label: cle, pur, purete: 1, commercial: pur })
        total_pur += pur
        total_commercial += pur
        continue
      }
      const purete = forced[`purete_${cle}`] ?? p[`purete_${cle}`] ?? ref.purete
      // l'hypochlorite se titre en degrés chlorométriques
      const commercial = ref.beaume
        ? pur / (BEAUME_VERS_MASSIQUE * purete)
        : purete > 0 ? pur / purete : 0
      lignes.push({ cle, label: ref.label, pur, purete, commercial })
      total_pur += pur
      total_commercial += commercial
    }
    lignes.sort((a, b) => b.commercial - a.commercial)

    const results = []
    for (const l of lignes) {
      results.push({ key: `${l.cle}_pur`, label: `${l.label} — produit pur`, unit: 'kg/j', value: l.pur })
      results.push({ key: `${l.cle}_com`, label: `${l.label} — produit commercial`, unit: 'kg/j', value: l.commercial })
      results.push({ key: `${l.cle}_an`, label: `${l.label} — produit commercial`, unit: 't/an', value: (l.commercial * CONST.NOMBRE_JOUR_PAR_AN) / 1000 })
    }
    results.push({ key: 'total_pur', label: 'Total en produit pur', unit: 'kg/j', value: total_pur })
    results.push({ key: 'total_com', label: 'Total en produit commercial', unit: 'kg/j', value: total_commercial })
    results.push({ key: 'total_an', label: 'Total en produit commercial', unit: 't/an', value: (total_commercial * CONST.NOMBRE_JOUR_PAR_AN) / 1000 })

    return {
      results,
      reactifs: Object.fromEntries(lignes.map((l) => [l.cle, l.commercial])),
      electricity: { total: 0, fixed: 0, detail: {} },
      warnings,
    }
  },
})
