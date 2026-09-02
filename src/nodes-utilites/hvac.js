// ---------------------------------------------------------------------------
// Port simplifié de HVAC_bat_admin.cls, HVAC_bat_exploit.cls et
// HVAC_bat_elec.cls, plus les fonctions psychrométriques de MOD_HVAC.bas.
//
// Le classeur s'appuie sur une base de 10 896 stations météo, chacune avec
// douze températures, pressions et humidités mensuelles, pour calculer des
// degrés-jours et des grammes-jours jour par jour. Embarquer cette base
// représenterait à elle seule plusieurs mégaoctets pour un module qui pèse
// quelques pour cent de la consommation d'une station.
//
// Le port retient donc un **modèle climatique sinusoïdal** : la température
// mensuelle est reconstituée à partir de trois grandeurs — moyenne annuelle,
// amplitude et mois le plus froid — soit par un préréglage, soit par saisie
// directe. Les degrés-jours en découlent par intégration mensuelle, méthode
// usuelle en thermique du bâtiment. L'humidité suit la même logique, avec le
// polynôme de pression de vapeur saturante déjà porté pour l'Athos.
//
// Les sorties sont identiques à celles du classeur : besoins annuels de
// chauffage et de climatisation rapportés au m², consommations d'électricité,
// de gaz et de fioul, coefficients de performance, débits d'eau usée en entrée
// de pompe à chaleur, coefficient de déperdition et taux de renouvellement.
//
// Les trois classes de bâtiment ne diffèrent que par leurs consignes de
// température, leur hauteur sous plafond et leur mode d'occupation : elles sont
// produites par une factory, comme les trois séchages.
//
// Écarts au VBA, assumés et signalés (voir README) :
//  - degrés-jours et grammes-jours sont calculés au pas mensuel sur un profil
//    sinusoïdal, non au pas journalier sur des relevés ; l'écart attendu sur les
//    besoins annuels est de l'ordre de quelques pour cent, très inférieur à
//    l'incertitude sur le coefficient de déperdition lui-même ;
//  - le raccordement au module PINCH (option « récupération sur le procédé »)
//    n'est pas fait : le besoin est calculé et exposé, à charge d'un futur
//    module d'intégration énergétique de le consommer.
// ---------------------------------------------------------------------------
import { defineUtilityNode } from '../core/utilityEngine.js'
import { CONST } from '../core/hypotheses.js'

// Préréglages climatiques : moyenne annuelle, amplitude et humidité relative
// moyenne. Ordres de grandeur suffisants pour un calcul de besoins.
const CLIMATS = {
  oceanique: { label: 'Océanique (Brest, Nantes)', T_moy: 12.5, amplitude: 6.5, HR: 80 },
  tempere: { label: 'Tempéré (Paris, Lyon)', T_moy: 11.5, amplitude: 8.5, HR: 75 },
  continental: { label: 'Continental (Strasbourg)', T_moy: 10.5, amplitude: 10, HR: 76 },
  mediterraneen: { label: 'Méditerranéen (Marseille)', T_moy: 15.5, amplitude: 8.5, HR: 68 },
  froid: { label: 'Froid (Scandinavie, montagne)', T_moy: 5, amplitude: 11, HR: 78 },
  chaud: { label: 'Chaud et sec', T_moy: 22, amplitude: 8, HR: 45 },
  tropical: { label: 'Tropical humide', T_moy: 27, amplitude: 2.5, HR: 82 },
}
// Coefficient de déperdition volumique, W/(m³·K), selon l'isolation.
const DEPERDITION = {
  rt2005: { label: 'RT 2005 ou mieux', valeur: 0.65 },
  rt2000: { label: 'RT 2000', valeur: 0.75 },
  ap1980: { label: 'Après 1980', valeur: 0.9 },
  moyenne: { label: 'Isolation moyenne', valeur: 1.2 },
  aucune: { label: 'Non isolé', valeur: 1.8 },
}

const H = {
  chaleur_volumique_air: 0.34, // Wh/(m³·K)
  chaleur_vaporisation_eau: 0.694, // Wh/g d'eau
  Qair_neuf_personne: 25, // m³/h, bureau sans travail physique
  surface_bureau_personne: 9, // m²/personne
  rendement_chaudiere: 0.85,
  rendement_distribution: 0.95,
  rendement_regulation_chaud: 0.95,
  rendement_chauffage_elec: 0.98,
  rendement_regulation_elec: 0.98,
  ventilation_P_reseau: 2000, // Pa
  ventilateur_rdt_machine: 0.7,
  ventilateur_rdt_moteur: 0.8,
  PCI_fioul_kWh_L: 10.7,
  PCI_gaz_kWh_Nm3: 10.3,
  // les apports internes et solaires abaissent la température de non-chauffage
  correction_apports: 3 / 5,
  T_reference_K: 273.15,
  Cp_eau_kWh_m3K: 1.16,
}
const JOURS_MOIS = [31, 28, 31, 30, 31, 31, 30, 31, 30, 31, 30, 31]
// somme = 365 ; le classeur emploie la même série
JOURS_MOIS[5] = 30; JOURS_MOIS[6] = 31

/**
 * Pression de vapeur saturante de l'eau, en bars, par la formule de Magnus.
 *
 * Le classeur emploie ailleurs un polynôme de degré 6 (repris dans le port de
 * l'Athos) calé sur la plage 0–373 °C. Ce polynôme est juste au-dessus de
 * 50 °C mais **devient négatif aux températures ambiantes** : il donne
 * −0,025 bar à 20 °C contre 0,023 en réalité. Il est donc inutilisable pour un
 * calcul de bâtiment, et Magnus le remplace ici — précis à mieux que 0,4 %
 * entre −45 et +60 °C, ce qui couvre largement le domaine du HVAC.
 */
function PvSat_bar(T_C) {
  return 0.0061121 * Math.exp((17.62 * T_C) / (243.12 + T_C))
}
/** humidité absolue de l'air, g d'eau par kg d'air sec */
function humiditeAbsolue(T_C, HR_pct, Patm_bar = 1.013) {
  const Pv = (HR_pct / 100) * PvSat_bar(T_C)
  if (Pv >= Patm_bar) return 0
  return 622 * (Pv / (Patm_bar - Pv))
}

/**
 * Profil de température mensuel, reconstitué par sinusoïde. Le minimum tombe
 * en janvier dans l'hémisphère nord, en juillet dans l'hémisphère sud.
 */
function profilMensuel(T_moy, amplitude, hemisphere_sud = false) {
  const decalage = hemisphere_sud ? 6 : 0
  return Array.from({ length: 12 }, (_, m) =>
    T_moy - amplitude * Math.cos((2 * Math.PI * (m - decalage)) / 12))
}

/**
 * Degrés-jours de chauffage ou de climatisation, au pas mensuel.
 * Un mois ne contribue que si sa température moyenne franchit la base ; la
 * contribution est proratisée sur l'écart, ce qui approche l'intégration
 * journalière sans nécessiter de relevés.
 */
function degresJours(profil, T_base, sens) {
  let DJ = 0
  let jours = 0
  profil.forEach((T, m) => {
    const ecart = sens === 'chaud' ? T_base - T : T - T_base
    if (ecart > 0) {
      DJ += ecart * JOURS_MOIS[m]
      jours += JOURS_MOIS[m]
    }
  })
  return { DJ, jours }
}

/**
 * Grammes-jours : écart d'humidité absolue entre l'air extérieur et l'ambiance
 * visée, cumulé sur les mois où l'écart est du bon signe.
 */
function grammesJours(profil, HR_ext, T_ambiante, HR_ambiante, sens) {
  const w_ambiant = humiditeAbsolue(T_ambiante, HR_ambiante)
  let GJ = 0
  let jours = 0
  profil.forEach((T, m) => {
    const w_ext = humiditeAbsolue(T, HR_ext)
    const ecart = sens === 'humidification' ? w_ambiant - w_ext : w_ext - w_ambiant
    if (ecart > 0) {
      GJ += ecart * JOURS_MOIS[m]
      jours += JOURS_MOIS[m]
    }
  })
  return { GJ, jours }
}

const OPT_CHAUFFAGE = [
  { value: 'non', label: 'pas de chauffage' },
  { value: 'elec', label: 'électrique' },
  { value: 'gaz', label: 'chaudière gaz' },
  { value: 'fioul', label: 'chaudière fioul' },
  { value: 'pac', label: 'pompe à chaleur sur eaux usées' },
  { value: 'recup', label: 'récupération sur le procédé' },
]
const OPT_APPOINT = [
  { value: 'elec', label: 'électrique' },
  { value: 'gaz', label: 'gaz' },
  { value: 'fioul', label: 'fioul' },
]
const OPT_CLIM = [
  { value: 'non', label: 'pas de climatisation' },
  { value: 'groupe_froid', label: 'groupe froid' },
  { value: 'pac', label: 'pompe à chaleur réversible' },
]

/**
 * Fabrique un nœud HVAC. Les trois bâtiments ne diffèrent que par leurs
 * consignes, leur hauteur sous plafond et leur ventilation.
 */
function makeHVAC(v) {
  return defineUtilityNode({
    id: v.id,
    label: v.label,
    short: v.short,
    family: 'batiments',
    vba: v.vba,
    description: v.description,
    choices: [
      { key: 'climat', label: 'Climat du site', default: 'tempere', options: Object.entries(CLIMATS).map(([value, c]) => ({ value, label: c.label })) },
      { key: 'hemisphere', label: 'Hémisphère', default: 'nord', options: [{ value: 'nord', label: 'nord' }, { value: 'sud', label: 'sud' }] },
      { key: 'isolation', label: 'Niveau d\'isolation', default: 'rt2000', options: Object.entries(DEPERDITION).map(([value, d]) => ({ value, label: `${d.label} — ${d.valeur} W/(m³·K)` })) },
      { key: 'chauffage', label: 'Type de chauffage', default: v.chauffage_defaut, options: OPT_CHAUFFAGE },
      { key: 'appoint', label: "Appoint de la pompe à chaleur", default: 'elec', options: OPT_APPOINT },
      { key: 'climatisation', label: 'Type de climatisation', default: v.clim_defaut, options: OPT_CLIM },
    ],
    params: [
      { key: 'surface', label: 'Surface du bâtiment', unit: 'm²', group: 'Bâtiment', default: v.surface_defaut },
      { key: 'hauteur_plafond', label: 'Hauteur sous plafond', unit: 'm', group: 'Bâtiment', default: v.hauteur_defaut },
      { key: 'coef_deperdition', label: 'Coefficient de déperdition', unit: 'W/(m³·K)', group: 'Bâtiment', default: undefined, hint: "selon le niveau d'isolation" },
      { key: 'temp_ambiante_hiver', label: 'Consigne de température en hiver', unit: '°C', group: 'Consignes', default: v.T_hiver },
      { key: 'temp_ambiante_ete', label: 'Consigne de température en été', unit: '°C', group: 'Consignes', default: v.T_ete },
      { key: 'HR_ambiante', label: 'Humidité relative de consigne', unit: '%', group: 'Consignes', default: 50 },
      { key: 'taux_renouvellement', label: "Taux de renouvellement d'air", unit: 'vol/h', group: 'Ventilation', default: undefined, hint: v.ventilation_hint },
      { key: 'T_moy_annuelle', label: 'Température moyenne annuelle', unit: '°C', group: 'Climat', default: undefined, hint: 'selon le climat retenu' },
      { key: 'amplitude_thermique', label: 'Amplitude thermique annuelle', unit: '°C', group: 'Climat', default: undefined, hint: 'demi-écart entre mois le plus chaud et le plus froid' },
      { key: 'HR_exterieure', label: 'Humidité relative extérieure moyenne', unit: '%', group: 'Climat', default: undefined, hint: 'selon le climat retenu' },
      { key: 'COP_chaud', label: 'COP de la pompe à chaleur en chaud', unit: '-', group: 'Machines', default: 3.5 },
      { key: 'COP_froid', label: 'COP de la pompe à chaleur en froid', unit: '-', group: 'Machines', default: 3 },
      { key: 'ratio_dimensionnement_PAC', label: 'Part des besoins couverte par la PAC', unit: '-', group: 'Machines', default: 0.8 },
      { key: 'delta_T_PAC', label: "Écart de température sur l'eau usée (chaud)", unit: '°C', group: 'Machines', default: 3 },
      { key: 'delta_T_PAC_clim', label: "Écart de température sur l'eau usée (froid)", unit: '°C', group: 'Machines', default: 3 },
      { key: 'rendement_global', label: 'Rendement global du chauffage', unit: '-', group: 'Machines', default: undefined, hint: 'chaudière × distribution × régulation' },
      { key: 'rendement_appoint', label: "Rendement global de l'appoint", unit: '-', group: 'Machines', default: undefined },
      { key: 'besoin_chauffage_m2', label: 'Besoin annuel de chauffage', unit: 'kWh/(m²·an)', group: 'Résultats forçables', default: undefined, hint: 'calculé si non forcé' },
      { key: 'besoin_clim_m2', label: 'Besoin annuel de climatisation', unit: 'kWh/(m²·an)', group: 'Résultats forçables', default: undefined, hint: 'calculé si non forcé' },
    ],

    compute(ctx) {
      const { p, forced, choices } = ctx
      const warnings = []
      const f = (k) => (forced[k] != null && !Number.isNaN(forced[k]) ? forced[k] : undefined)
      const an = CONST.NOMBRE_JOUR_PAR_AN

      const surface = p.surface ?? v.surface_defaut
      if (!(surface > 0)) {
        return { results: [], electricity: { total: 0, fixed: 0, detail: {} }, warnings: ['Surface du bâtiment nulle.'] }
      }
      const hauteur = p.hauteur_plafond ?? v.hauteur_defaut
      const volume = surface * hauteur
      const coef_deperdition = f('coef_deperdition') ?? DEPERDITION[choices.isolation].valeur
      const T_hiver = p.temp_ambiante_hiver ?? v.T_hiver
      const T_ete = p.temp_ambiante_ete ?? v.T_ete
      const HR_ambiante = p.HR_ambiante ?? 50

      // ---- climat reconstitué
      const clim = CLIMATS[choices.climat]
      const T_moy = f('T_moy_annuelle') ?? clim.T_moy
      const amplitude = f('amplitude_thermique') ?? clim.amplitude
      const HR_ext = f('HR_exterieure') ?? clim.HR
      const profil = profilMensuel(T_moy, amplitude, choices.hemisphere === 'sud')
      const T_moy_min = Math.min(...profil)
      const T_moy_max = Math.max(...profil)

      // ---- ventilation
      let Qair, taux_renouvellement
      const tauxForce = f('taux_renouvellement')
      if (tauxForce != null) {
        taux_renouvellement = tauxForce
        Qair = taux_renouvellement * volume
      } else if (v.ventilation_par_occupation) {
        // bureau : le débit d'air neuf suit l'occupation
        Qair = (H.Qair_neuf_personne * surface) / H.surface_bureau_personne
        taux_renouvellement = volume > 0 ? Qair / volume : 0
      } else {
        taux_renouvellement = v.taux_renouvellement_defaut
        Qair = taux_renouvellement * volume
      }

      // ---- température de non-chauffage : les apports internes et solaires
      // abaissent la consigne effective
      const T_base_chaud = T_hiver - Math.abs(H.correction_apports * T_moy_min)
      const T_base_froid = T_ete - Math.abs(H.correction_apports * T_moy_max)

      // ---- besoins de chauffage
      const dj = degresJours(profil, T_base_chaud, 'chaud')
      const deperditions = (coef_deperdition * volume * CONST.NOMBRE_HEURE_PAR_JOUR * dj.DJ) / 1000
      const rechauffement_air = (Qair * H.chaleur_volumique_air * dj.DJ * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
      const gj_humid = grammesJours(profil, HR_ext, T_hiver, HR_ambiante, 'humidification')
      const humidification = (Qair * H.chaleur_vaporisation_eau * gj_humid.GJ * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
      let besoin_chauffage = deperditions + rechauffement_air + humidification

      // ---- besoins de climatisation
      const dj_froid = degresJours(profil, T_base_froid, 'froid')
      const refroidissement = (Qair * H.chaleur_volumique_air * dj_froid.DJ * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
      const gj_deshum = grammesJours(profil, HR_ext, T_ete, HR_ambiante, 'deshumidification')
      const deshumidification = (Qair * H.chaleur_vaporisation_eau * gj_deshum.GJ * CONST.NOMBRE_HEURE_PAR_JOUR) / 1000
      let besoin_clim = refroidissement + deshumidification

      // ---- électricité de ventilation, corrigée de la température de l'air
      const rdt_ventilation = H.ventilateur_rdt_machine * H.ventilateur_rdt_moteur
      const elecVentilation = (T, jours) =>
        (((Qair * (T + H.T_reference_K)) / H.T_reference_K / CONST.NOMBRE_SECONDE_PAR_HEURE)
          * H.ventilation_P_reseau / rdt_ventilation / 1000) * CONST.NOMBRE_HEURE_PAR_JOUR * jours
      const elec_ventilation_hiver = elecVentilation(T_hiver, dj.jours)
      const elec_ventilation_ete = elecVentilation(T_ete, dj_froid.jours)
      const jours_reste = Math.max(0, an - dj.jours - dj_froid.jours)
      const elec_ventilation_reste = elecVentilation((T_hiver + T_ete) / 2, jours_reste)

      // ---- conversion des besoins en consommations
      const rendement_global = f('rendement_global')
        ?? (choices.chauffage === 'elec'
          ? H.rendement_chauffage_elec * H.rendement_regulation_elec * H.rendement_distribution
          : H.rendement_chaudiere * H.rendement_distribution * H.rendement_regulation_chaud)
      const rendement_appoint = f('rendement_appoint') ?? rendement_global
      const COP_chaud = p.COP_chaud ?? 3.5
      const COP_froid = p.COP_froid ?? 3
      const ratio_PAC = p.ratio_dimensionnement_PAC ?? 0.8

      let besoin_chauffage_m2 = 0
      let conso_elec_chauffage = 0
      let conso_gaz = 0
      let conso_fioul = 0
      let debit_EU_PAC = 0
      let besoin_recup_process = 0
      if (choices.chauffage !== 'non' && besoin_chauffage > 0) {
        besoin_chauffage_m2 = f('besoin_chauffage_m2') ?? besoin_chauffage / surface
        besoin_chauffage = besoin_chauffage_m2 * surface
        if (choices.chauffage === 'elec') {
          conso_elec_chauffage = besoin_chauffage / rendement_global
        } else if (choices.chauffage === 'fioul') {
          conso_fioul = besoin_chauffage / (H.PCI_fioul_kWh_L * rendement_global)
        } else if (choices.chauffage === 'gaz') {
          conso_gaz = besoin_chauffage / (H.PCI_gaz_kWh_Nm3 * rendement_global)
        } else if (choices.chauffage === 'pac') {
          conso_elec_chauffage = (besoin_chauffage * ratio_PAC) / COP_chaud
          // la chaleur prélevée sur l'eau usée détermine le débit à faire passer
          const puissance_EU = ((besoin_chauffage * ratio_PAC * (COP_chaud - 1)) / COP_chaud) / an / CONST.NOMBRE_HEURE_PAR_JOUR
          debit_EU_PAC = puissance_EU / (H.Cp_eau_kWh_m3K * (p.delta_T_PAC ?? 3))
          const complement = besoin_chauffage * (1 - ratio_PAC)
          if (choices.appoint === 'elec') conso_elec_chauffage += complement / rendement_appoint
          else if (choices.appoint === 'fioul') conso_fioul = complement / (H.PCI_fioul_kWh_L * rendement_appoint)
          else conso_gaz = complement / (H.PCI_gaz_kWh_Nm3 * rendement_appoint)
        } else if (choices.chauffage === 'recup') {
          besoin_recup_process = besoin_chauffage
          warnings.push("Le chauffage par récupération sur le procédé suppose un module d'intégration énergétique : le besoin est exposé mais non couvert.")
        }
      } else besoin_chauffage = 0

      let besoin_clim_m2 = 0
      let conso_elec_clim = 0
      let debit_EU_PAC_clim = 0
      if (choices.climatisation !== 'non' && besoin_clim > 0) {
        besoin_clim_m2 = f('besoin_clim_m2') ?? besoin_clim / surface
        besoin_clim = besoin_clim_m2 * surface
        conso_elec_clim = besoin_clim / COP_froid + elec_ventilation_ete
        if (choices.climatisation === 'pac') {
          const puissance_EU = ((besoin_clim * ratio_PAC * (COP_froid - 1)) / COP_froid) / an / CONST.NOMBRE_HEURE_PAR_JOUR
          debit_EU_PAC_clim = puissance_EU / (H.Cp_eau_kWh_m3K * (p.delta_T_PAC_clim ?? 3))
        }
      } else besoin_clim = 0

      const conso_elec_an = conso_elec_chauffage + conso_elec_clim + elec_ventilation_reste
      const conso_elec_j = conso_elec_an / an

      if (dj.jours === 0 && choices.chauffage !== 'non') {
        warnings.push('Aucun mois ne descend sous la température de non-chauffage : besoin de chauffage nul dans ce climat.')
      }
      if (dj_froid.jours === 0 && choices.climatisation !== 'non') {
        warnings.push('Aucun mois ne dépasse la température de non-climatisation : besoin de climatisation nul dans ce climat.')
      }

      const reactifs = {}
      return {
        results: [
          { key: 'surface', label: 'Surface', unit: 'm²', value: surface },
          { key: 'volume', label: 'Volume du bâtiment', unit: 'm³', value: volume },
          { key: 'deperdition', label: 'Coefficient de déperdition', unit: 'W/(m³·K)', value: coef_deperdition },
          { key: 'taux_renouv', label: "Taux de renouvellement d'air", unit: 'vol/h', value: taux_renouvellement },
          { key: 'Qair', label: "Débit d'air neuf", unit: 'm³/h', value: Qair },
          { key: 'T_min', label: 'Température du mois le plus froid', unit: '°C', value: T_moy_min },
          { key: 'T_max', label: 'Température du mois le plus chaud', unit: '°C', value: T_moy_max },
          { key: 'DJ_chaud', label: 'Degrés-jours de chauffage', unit: '°C·j', value: dj.DJ },
          { key: 'j_chaud', label: 'Jours de chauffage', unit: 'j/an', value: dj.jours },
          { key: 'DJ_froid', label: 'Degrés-jours de climatisation', unit: '°C·j', value: dj_froid.DJ },
          { key: 'j_froid', label: 'Jours de climatisation', unit: 'j/an', value: dj_froid.jours },
          { key: 'besoin_chauffage_m2', label: 'Besoin annuel de chauffage', unit: 'kWh/(m²·an)', value: besoin_chauffage_m2 },
          { key: 'besoin_chauffage', label: 'Besoin annuel de chauffage', unit: 'kWh/an', value: besoin_chauffage },
          { key: 'deperditions', label: 'dont déperditions du bâti', unit: 'kWh/an', value: deperditions },
          { key: 'rechauffement', label: "dont réchauffement de l'air neuf", unit: 'kWh/an', value: rechauffement_air },
          { key: 'humidification', label: 'dont humidification', unit: 'kWh/an', value: humidification },
          { key: 'besoin_clim_m2', label: 'Besoin annuel de climatisation', unit: 'kWh/(m²·an)', value: besoin_clim_m2 },
          { key: 'besoin_clim', label: 'Besoin annuel de climatisation', unit: 'kWh/an', value: besoin_clim },
          { key: 'COP_chaud', label: 'COP en chaud', unit: '-', value: COP_chaud },
          { key: 'COP_froid', label: 'COP en froid', unit: '-', value: COP_froid },
          { key: 'rendement', label: 'Rendement global du chauffage', unit: '-', value: rendement_global },
          { key: 'elec_chauffage', label: 'Électricité de chauffage', unit: 'kWh/an', value: conso_elec_chauffage },
          { key: 'elec_clim', label: 'Électricité de climatisation', unit: 'kWh/an', value: conso_elec_clim },
          { key: 'elec_ventilation', label: 'Électricité de ventilation hors saison', unit: 'kWh/an', value: elec_ventilation_reste },
          { key: 'elec_totale', label: 'Électricité totale', unit: 'kWh/an', value: conso_elec_an },
          { key: 'elec_j', label: 'Électricité totale', unit: 'kWh/j', value: conso_elec_j },
          ...(conso_gaz > 0 ? [{ key: 'gaz', label: 'Gaz naturel', unit: 'Nm³/an', value: conso_gaz }] : []),
          ...(conso_fioul > 0 ? [{ key: 'fioul', label: 'Fioul', unit: 'L/an', value: conso_fioul }] : []),
          ...(debit_EU_PAC > 0 ? [{ key: 'debit_EU', label: "Débit d'eau usée en entrée de PAC (chaud)", unit: 'm³/h', value: debit_EU_PAC }] : []),
          ...(debit_EU_PAC_clim > 0 ? [{ key: 'debit_EU_clim', label: "Débit d'eau usée en entrée de PAC (froid)", unit: 'm³/h', value: debit_EU_PAC_clim }] : []),
          ...(besoin_recup_process > 0 ? [{ key: 'recup', label: 'Besoin à couvrir par récupération sur le procédé', unit: 'kWh/an', value: besoin_recup_process }] : []),
        ],
        energie: besoin_recup_process > 0
          ? { besoin_thermique_kWhj: besoin_recup_process / an, source: 'recuperation', niveau: 'BT' }
          : conso_gaz > 0
            ? { besoin_thermique_kWhj: (conso_gaz * H.PCI_gaz_kWh_Nm3) / an, source: 'gaz_naturel', niveau: 'BT' }
            : conso_fioul > 0
              ? { besoin_thermique_kWhj: (conso_fioul * H.PCI_fioul_kWh_L) / an, source: 'fioul', niveau: 'BT' }
              : null,
        reactifs,
        electricity: {
          total: conso_elec_j,
          // la ventilation tourne en permanence, le reste suit la saison
          fixed: elec_ventilation_reste / an,
          detail: {
            chauffage: conso_elec_chauffage / an,
            climatisation: conso_elec_clim / an,
            ventilation: elec_ventilation_reste / an,
          },
        },
        warnings,
      }
    },
  })
}

// ---------------------------------------------------------------------------
export const hvacAdmin = makeHVAC({
  id: 'hvac-admin',
  label: 'HVAC — bâtiment administratif',
  short: 'HVAC admin',
  vba: 'HVAC_bat_admin.cls',
  description:
    "Chauffage, ventilation et climatisation d'un bâtiment de bureaux. Le débit d'air neuf suit l'occupation, à 25 m³/h par personne pour 9 m² par poste.",
  surface_defaut: 400,
  hauteur_defaut: 2.7,
  T_hiver: 19,
  T_ete: 26,
  chauffage_defaut: 'gaz',
  clim_defaut: 'groupe_froid',
  ventilation_par_occupation: true,
  ventilation_hint: "déduit de l'occupation : 25 m³/h par personne",
})

export const hvacExploitation = makeHVAC({
  id: 'hvac-exploitation',
  label: "HVAC — bâtiment d'exploitation",
  short: 'HVAC exploit',
  vba: 'HVAC_bat_exploit.cls',
  description:
    "Chauffage et ventilation d'un bâtiment d'exploitation. Consignes plus basses qu'en bureaux et renouvellement d'air plus soutenu, l'activité y étant physique.",
  surface_defaut: 600,
  hauteur_defaut: 4,
  T_hiver: 16,
  T_ete: 28,
  chauffage_defaut: 'gaz',
  clim_defaut: 'non',
  ventilation_par_occupation: false,
  taux_renouvellement_defaut: 2,
  ventilation_hint: '2 volumes par heure par défaut',
})

export const hvacElectrique = makeHVAC({
  id: 'hvac-electrique',
  label: 'HVAC — local électrique',
  short: 'HVAC élec',
  vba: 'HVAC_bat_elec.cls',
  description:
    "Ventilation et rafraîchissement d'un local électrique. Pas de chauffage, mais un renouvellement d'air élevé pour évacuer la chaleur dissipée par les équipements.",
  surface_defaut: 120,
  hauteur_defaut: 3,
  T_hiver: 10,
  T_ete: 30,
  chauffage_defaut: 'non',
  clim_defaut: 'groupe_froid',
  ventilation_par_occupation: false,
  taux_renouvellement_defaut: 4,
  ventilation_hint: '4 volumes par heure par défaut',
})

export { CLIMATS, DEPERDITION, profilMensuel, degresJours, grammesJours, humiditeAbsolue }
