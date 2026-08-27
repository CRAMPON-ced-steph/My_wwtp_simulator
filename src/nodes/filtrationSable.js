// Port de F5_Filtration_sable.cls — filtration tertiaire sur matériau granulaire
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, HYP, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const F = HYP.filtrasable
const mat = (c) => F.materiaux.indexOf(c.choices.materiau)
const H = { surpresseur_conso_spec: 4.5 } // Wh/(m³·mCE)

export default defineNode({
  id: 'filtration-sable',
  label: 'Filtration sur sable',
  short: 'Filtre sable',
  family: 'tertiaire',
  vba: 'F5_Filtration_sable.cls',
  description: 'Filtre granulaire (ponce, schiste, sable, bicouche). Surface sur la pointe temps de pluie, fréquence de lavage sur la capacité de rétention, eaux sales renvoyées en tête.',
  choices: [
    { key: 'materiau', label: 'Matériau filtrant', default: 'sable', options: F.materiaux.map((m, i) => ({ value: m, label: F.labels[i] })) },
    { key: 'retour_ES', label: 'Retour des eaux sales', default: 'pompage', options: [{ value: 'pompage', label: 'par pompage' }, { value: 'gravitaire', label: 'gravitaire' }] },
  ],
  params: [
    { key: 'ratio_Q_admis', label: 'Part du débit admise sur les filtres', unit: '-', group: 'Hydraulique', default: 1 },
    { key: 'Q_design', label: 'Débit de dimensionnement', unit: 'm³/h', group: 'Hydraulique', default: (c) => (c.inNominal.Q * c.p.ratio_Q_admis * c.site.pointe_TP) / CONST.NOMBRE_HEURE_PAR_JOUR },
    { key: 'surface_necessaire', label: 'Surface de filtration nécessaire', unit: 'm²', group: 'Filtres', default: (c) => c.p.Q_design / F.vitesse_filtration[mat(c)] },
    { key: 'nb_filtres', label: 'Nombre de filtres', unit: 'u', group: 'Filtres', default: undefined, hint: 'calculé si non forcé' },
    { key: 'surface_unitaire', label: 'Surface unitaire', unit: 'm²', group: 'Filtres', default: undefined, hint: 'calculée si non forcée' },
    { key: 'hauteur_materiau', label: 'Hauteur de matériau', unit: 'm', group: 'Filtres', default: (c) => F.hauteur_materiau[mat(c)] },
    { key: 'outlet_MES_mgL', label: 'MES en sortie', unit: 'mg/L', group: 'Filtres', default: 2 },
    { key: 'capacite_retention', label: 'Capacité de rétention', unit: 'kg MES/m³', group: 'Lavage', default: (c) => F.capacite_retention[mat(c)] },
    { key: 'conso_eau_lavage', label: "Consommation d'eau de lavage", unit: 'm³/(m²·lavage)', group: 'Lavage', default: (c) => F.eau_lavage_m3_m2_lavage[mat(c)] },
    { key: 'pompe_lavage_P', label: 'Pression pompe eau de lavage', unit: 'mCE', group: 'Lavage', default: (c) => F.P_eau_lavage[mat(c)] },
    { key: 'pompe_lavage_rdt', label: 'Rendement global pompe de lavage', unit: '-', group: 'Lavage', default: undefined, hint: 'calculé si non forcé' },
    { key: 'air_lavage_vitesse', label: "Vitesse air de lavage", unit: 'Nm/h', group: 'Lavage', default: (c) => F.air_lavage_vitesse[mat(c)] },
    { key: 'air_lavage_P', label: "Pression air de lavage", unit: 'mCE', group: 'Lavage', default: (c) => F.P_air_lavage[mat(c)] },
    { key: 'air_lavage_tps', label: "Durée d'air par lavage", unit: 'min', group: 'Lavage', default: F.air_lavage_tps_fct_min_lavage },
    { key: 'pompe_ES_nb', label: 'Nombre de pompes eaux sales', unit: 'u', group: 'Eaux sales', default: 2 },
    { key: 'pompe_ES_P', label: 'Pression pompes eaux sales', unit: 'mCE', group: 'Eaux sales', default: 10 },
    { key: 'pompe_ES_tps', label: 'Durée de fonctionnement pompes ES', unit: 'h/j', group: 'Eaux sales', default: F.pompe_ES_tps_fct },
    { key: 'pompe_ES_Qunitaire', label: 'Débit unitaire pompes ES', unit: 'm³/h', group: 'Eaux sales', default: undefined, hint: 'calculé si non forcé' },
    { key: 'pompe_ES_rdt', label: 'Rendement global pompes ES', unit: '-', group: 'Eaux sales', default: undefined, hint: 'calculé si non forcé' },
  ],
  compute(ctx) {
    const { p, forced, choices } = ctx
    const m = mat(ctx)
    const warnings = []

    // ---- dimensionnement des filtres (nominal)
    let nb_filtres, surface_unitaire
    if (forced.nb_filtres != null) {
      nb_filtres = forced.nb_filtres
      surface_unitaire = forced.surface_unitaire ?? p.surface_necessaire / nb_filtres
    } else if (forced.surface_unitaire != null) {
      surface_unitaire = forced.surface_unitaire
      nb_filtres = Math.ceil(p.surface_necessaire / surface_unitaire - 1e-12)
    } else if (F.surface_unitaire_max * F.nb_filtres_mini <= p.surface_necessaire) {
      surface_unitaire = F.surface_unitaire_max
      nb_filtres = Math.ceil(p.surface_necessaire / surface_unitaire - 1e-12)
    } else {
      nb_filtres = F.nb_filtres_mini
      surface_unitaire = p.surface_necessaire / nb_filtres
    }
    const surface_reelle = nb_filtres * surface_unitaire

    let ES_Q_prev = 0
    const pass = (s, reel) => {
      const r = p.ratio_Q_admis
      const Qt = s.Q * r
      const t = { DCO: s.DCO * r, DBO: s.DBO * r, MES: s.MES * r, NK: s.NK * r, NH4: s.NH4 * r, NO3: s.NO3 * r, Pt: s.Pt * r }
      let outlet_MES = p.outlet_MES_mgL
      let ES_MES
      if ((t.MES / Qt) * 1000 < outlet_MES) {
        outlet_MES = (t.MES / Qt) * 1000
        ES_MES = 0
        if (reel) warnings.push('MES entrante déjà inférieure à la consigne de sortie : la filtration sur sable est inutile ici.')
      } else {
        ES_MES = t.MES - (outlet_MES * (Qt - ES_Q_prev)) / 1000 // le VBA utilise ES_Q de l'itération précédente
      }
      const nb_lavage = ES_MES / (p.capacite_retention * surface_reelle * p.hauteur_materiau)
      const ES_Q = nb_lavage * surface_reelle * p.conso_eau_lavage
      ES_Q_prev = ES_Q
      const ES = { origine: 'filtrasable', Q: ES_Q, MES: ES_MES, NH4: (s.NH4 * ES_Q) / s.Q, NO3: (s.NO3 * ES_Q) / s.Q, DCO: (s.DCO * ES_Q) / s.Q, DBO: (s.DBO * ES_Q) / s.Q, NK: (s.NK * ES_Q) / s.Q, Pt: (s.Pt * ES_Q) / s.Q }
      const out = cloneStream(s)
      out.Q = (1 - r) * s.Q + (Qt - ES.Q)
      out.MES = (1 - r) * s.MES + (t.MES - ES.MES)
      out.DCO = (1 - r) * s.DCO + (t.DCO - ES.DCO)
      out.DBO = (1 - r) * s.DBO + (t.DBO - ES.DBO)
      out.NK = (1 - r) * s.NK + (t.NK - ES.NK)
      out.NH4 = (1 - r) * s.NH4 + (t.NH4 - ES.NH4)
      out.NO3 = (1 - r) * s.NO3 + (t.NO3 - ES.NO3)
      out.Pt = (1 - r) * s.Pt + (t.Pt - ES.Pt)
      return { out, ES, nb_lavage, outlet_MES }
    }
    const nom = pass(ctx.inNominal, false)
    const reel = pass(ctx.inReel, true)

    // ---- électricité (réel)
    let elec_EL = 0, elec_ES = 0
    if (reel.nb_lavage > 0) {
      const EL_Qu = F.eau_lavage_vitesse[m] * surface_reelle
      const rdt_EL = forced.pompe_lavage_rdt ?? rendementPompeGlobal(EL_Qu, p.pompe_lavage_P, 0.7)
      elec_EL = (reel.ES.Q / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * p.pompe_lavage_P / rdt_EL
      if (choices.retour_ES === 'pompage') {
        const rep = repartitionPompes(reel.ES.Q, { nb: p.pompe_ES_nb, tps: p.pompe_ES_tps }, { nb: forced.pompe_ES_nb, tps: forced.pompe_ES_tps, Q_unitaire: forced.pompe_ES_Qunitaire }, F.pompe_ES_tps_fct)
        if (rep.incoherence) warnings.push('Incohérence sur les pompes de retour des eaux sales.')
        const rdt_ES = forced.pompe_ES_rdt ?? rendementPompeGlobal(rep.Q_unitaire, p.pompe_ES_P, 0.7)
        elec_ES = electricitePompage(rep.Q_unitaire, p.pompe_ES_P, rdt_ES, rep.nb, rep.tps)
      }
    }
    const elec_air = (p.air_lavage_vitesse * surface_reelle * (p.air_lavage_tps / CONST.NOMBRE_MINUTE_PAR_HEURE) * reel.nb_lavage * p.air_lavage_P * H.surpresseur_conso_spec) / 1000

    return {
      outNominal: nom.out,
      outReel: reel.out,
      eauxSales: reel.ES,
      results: [
        { key: 'nb_filtres', label: 'Nombre de filtres', unit: 'u', value: nb_filtres },
        { key: 'surface_unitaire', label: 'Surface unitaire', unit: 'm²', value: surface_unitaire },
        { key: 'surface_reelle', label: 'Surface installée', unit: 'm²', value: surface_reelle },
        { key: 'v_filtration', label: 'Vitesse de filtration (pointe)', unit: 'm/h', value: p.Q_design / surface_reelle },
        { key: 'nb_lavage', label: 'Lavages par filtre et par jour (réel)', unit: '/j', value: reel.nb_lavage },
        { key: 'ES_Q', label: 'Eaux sales (réel)', unit: 'm³/j', value: reel.ES.Q },
        { key: 'ES_MES', label: 'MES retenues (réel)', unit: 'kg/j', value: reel.ES.MES },
        { key: 'MES_out', label: 'MES sortie (réel)', unit: 'mg/L', value: conc(reel.out, 'MES') },
      ],
      electricity: { total: elec_EL + elec_ES + elec_air, fixed: 0, detail: { eau_lavage: elec_EL, retour_ES: elec_ES, air_lavage: elec_air } },
      warnings,
    }
  },
})
