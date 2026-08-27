// Port de D1_decanteur_simple.cls (Decanteur_I_simple + calcul_consommation_electrique)
// Décantation primaire simple / lamellaire (MULTIFLO MONO).
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, HYP, electricitePompage, rendementPompeGlobal, repartitionPompes } from '../core/hypotheses.js'

const H = {
  MV_MES_reference: HYP.dec_I_boues_MV_MES, // 0.68
  hyp_rdt_P_NoDiscfilter: 0.11,
  dec_diametre_limite: HYP.dec_I_diametre_limite_m,
  Pw_racleur_sup: HYP.dec_I_Pw_racleur_sup_Dlim_kW,
  Pw_racleur_inf: HYP.dec_I_Pw_racleur_inf_Dlim_kW,
  rendement_pompe_extraction: 0.7,
}

const lam = (c) => c.choices.lamellaire === 'oui'

export default defineNode({
  id: 'decantation-simple',
  label: 'Décantation simple / lamellaire',
  short: 'Décanteur I',
  family: 'primaire',
  vba: 'D1_decanteur_simple.cls',
  description: 'Décanteur primaire sans réactif (MULTIFLO MONO en lamellaire). Rendements fixes par paramètre, surface au miroir sur la pointe temps de pluie.',
  choices: [
    { key: 'lamellaire', label: 'Décantation lamellaire', default: 'non', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'dimensionnement', label: 'Dimensionnement du poste', default: '100', options: [{ value: '100', label: 'sur 100 % du débit' }, { value: 'bypass', label: 'sur le % non by-passé' }] },
  ],
  params: [
    { key: 'bypass', label: 'Eau by-passée au primaire', unit: '%', group: 'Hydraulique', default: 0 },
    { key: 'nb_ouvrages', label: "Nombre d'ouvrages", unit: '-', group: 'Ouvrage', default: 1 },
    { key: 'V_miroir_limite', label: 'Vitesse au miroir limite (temps de pluie)', unit: 'm/h', group: 'Ouvrage', default: (c) => (lam(c) ? 13 : 5), hint: 'Aquademy / CAM' },
    { key: 'S_miroir', label: 'Surface au miroir', unit: 'm²', group: 'Ouvrage', default: (c) => surfaceMiroir(c, c.p.V_miroir_limite) },
    { key: 'hauteur', label: 'Hauteur du décanteur', unit: 'm', group: 'Ouvrage', default: 3.5 },
    { key: 'rdt_MES', label: 'Rendement MES', unit: '-', group: 'Rendements', default: (c) => (lam(c) ? 0.6 : 0.5) },
    { key: 'rdt_DCO', label: 'Rendement DCO', unit: '-', group: 'Rendements', default: (c) => (lam(c) ? 0.3 : 0.28) },
    { key: 'rdt_DBO', label: 'Rendement DBO5', unit: '-', group: 'Rendements', default: (c) => (lam(c) ? 0.28 : 0.25) },
    { key: 'rdt_NK', label: 'Rendement NK', unit: '-', group: 'Rendements', default: 0.09 },
    { key: 'rdt_P', label: 'Rendement P', unit: '-', group: 'Rendements', default: H.hyp_rdt_P_NoDiscfilter },
    { key: 'MV_MES', label: 'MV/MES des boues', unit: '-', group: 'Boues', default: H.MV_MES_reference },
    { key: 'boues_concentration', label: 'Concentration des boues extraites', unit: 'g/L', group: 'Boues', default: (c) => (lam(c) ? 20 : 10) },
    { key: 'extraction_nb_pompe', label: "Nombre de pompes d'extraction", unit: 'u', group: 'Extraction', default: 2 },
    { key: 'extraction_P_refoulement', label: 'Pression de refoulement', unit: 'mCE', group: 'Extraction', default: 15 },
    { key: 'extraction_tps_fct', label: 'Durée de fonctionnement', unit: 'h/j', group: 'Extraction', default: CONST.NOMBRE_HEURE_PAR_JOUR },
    { key: 'extraction_Q_unitaire', label: 'Débit de pompage unitaire', unit: 'm³/h', group: 'Extraction', default: undefined, hint: 'calculé si non forcé' },
    { key: 'extraction_pompe_rdt', label: 'Rendement global pompes (machine+moteur)', unit: '-', group: 'Extraction', default: undefined, hint: 'calculé si non forcé' },
  ],
  compute(ctx) {
    const { p, site, forced } = ctx
    const bypass = (p.bypass || 0) / 100
    const warnings = []

    const pass = (s) => {
      const f = ctx.choices.dimensionnement === '100' ? 1 : 1 - bypass
      const Qt = f * s.Q
      const traite = { DCO: f * s.DCO, DBO: f * s.DBO, MES: f * s.MES, NK: f * s.NK, Pt: f * s.Pt }
      const boues_MES = traite.MES * p.rdt_MES
      const boues_Q = p.boues_concentration > 0 ? boues_MES / p.boues_concentration : 0
      const out = cloneStream(s)
      out.Q = s.Q - boues_Q
      out.DCO = s.DCO - traite.DCO * p.rdt_DCO
      out.DBO = s.DBO - traite.DBO * p.rdt_DBO
      out.MES = s.MES - traite.MES * p.rdt_MES
      out.NK = s.NK - traite.NK * p.rdt_NK
      out.NH4 = (s.NH4 * out.Q) / s.Q
      out.NO3 = (s.NO3 * out.Q) / s.Q
      out.Pt = s.Pt - traite.Pt * p.rdt_P
      return { out, boues_MES, boues_Q, Qt }
    }
    const nom = pass(ctx.inNominal)
    const reel = pass(ctx.inReel)

    // vitesse au miroir recalculée
    const Qpointe = site.pointe_TP * (nom.Qt - site.Q_retour) + site.Q_retour
    const V_miroir_recalc = p.S_miroir > 0 ? Qpointe / (CONST.NOMBRE_HEURE_PAR_JOUR * p.S_miroir) : 0
    const V_moyenne = p.S_miroir > 0 ? nom.Qt / (CONST.NOMBRE_HEURE_PAR_JOUR * p.S_miroir) : 0
    const tps_retention = V_moyenne > 0 ? p.hauteur / V_moyenne : 0

    // électricité — râcleur
    const S_unit = p.S_miroir / p.nb_ouvrages
    const Pw = S_unit < CONST.PI * Math.pow(H.dec_diametre_limite / 2, 2) ? H.Pw_racleur_inf : H.Pw_racleur_sup
    const electricite_racleur = p.nb_ouvrages * Pw * CONST.NOMBRE_HEURE_PAR_JOUR
    // électricité — extraction (sur le fonctionnement réel)
    const rep = repartitionPompes(reel.boues_Q, { nb: p.extraction_nb_pompe, tps: p.extraction_tps_fct }, { nb: forced.extraction_nb_pompe, tps: forced.extraction_tps_fct, Q_unitaire: forced.extraction_Q_unitaire })
    if (rep.incoherence) warnings.push("Incohérence entre nombre, débit et durée des pompes d'extraction.")
    const rdt = forced.extraction_pompe_rdt ?? rendementPompeGlobal(rep.Q_unitaire, p.extraction_P_refoulement, H.rendement_pompe_extraction)
    const electricite_extraction = electricitePompage(rep.Q_unitaire, p.extraction_P_refoulement, rdt, rep.nb, rep.tps)

    return {
      outNominal: nom.out,
      outReel: reel.out,
      sludge: { origine: 'I_simple', Q: reel.boues_Q, MES: reel.boues_MES, concentration: p.boues_concentration, MV_MES: p.MV_MES },
      results: [
        { key: 'V_miroir_recalc', label: 'Vitesse au miroir limite recalculée', unit: 'm/h', value: V_miroir_recalc },
        { key: 'V_moyenne', label: 'Vitesse moyenne au miroir', unit: 'm/h', value: V_moyenne },
        { key: 'tps_retention', label: 'Temps de rétention', unit: 'h', value: tps_retention },
        { key: 'boues_MES_nom', label: 'Boues à extraire (nominal)', unit: 'kg MES/j', value: nom.boues_MES },
        { key: 'boues_MES', label: 'Boues à extraire (réel)', unit: 'kg MES/j', value: reel.boues_MES },
        { key: 'boues_Q', label: 'Boues à extraire (réel)', unit: 'm³/j', value: reel.boues_Q },
        { key: 'ext_nb', label: "Pompes d'extraction", unit: 'u', value: rep.nb },
        { key: 'ext_Qu', label: 'Débit unitaire extraction', unit: 'm³/h', value: rep.Q_unitaire },
        { key: 'ext_rdt', label: 'Rendement global extraction', unit: '-', value: rdt },
        { key: 'MES_out', label: 'MES sortie (réel)', unit: 'mg/L', value: conc(reel.out, 'MES') },
      ],
      electricity: { total: electricite_racleur + electricite_extraction, fixed: electricite_racleur, detail: { racleur: electricite_racleur, extraction: electricite_extraction } },
      warnings,
    }
  },
})

function surfaceMiroir(c, V) {
  const bypass = (c.p?.bypass || 0) / 100
  const f = c.choices.dimensionnement === '100' ? 1 : 1 - bypass
  const Qt = f * c.inNominal.Q
  const Qr = f * c.site.Q_retour
  return (c.site.pointe_TP * (Qt - Qr) + Qr) / (CONST.NOMBRE_HEURE_PAR_JOUR * V)
}
