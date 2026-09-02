// Port de F4_Discfilter_III.cls — filtration tertiaire sur disques (Hydrotech)
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, precipitationP, rendementMoteur } from '../core/hypotheses.js'

const H = {
  ratio_ES_lavage: 0.5,
  nombre_pompe_lavage: 2,
  nombre_pompe_retour: 2,
  fonction_limit: 60, // kW
  p_installee: 1.1, // kW installé pour Q_moyen (Rambouillet)
  Q_moyen: 7000, // m³/h
  rdt_moteur_disques: 0.8,
  duree_func: 24,
  rdt_DCO: 0.1,
  rdt_DBO: 0.15,
  rdt_NK: 0.11,
  a_rdt_MES: [0.23, 0.0428, -0.0012, 0.00001], // polynôme en MES entrée (mg/L), filtre 10 µm
  a_ES_conc: [0.075, 0.0455], // g/L
  rdt_machine_pompe_lavage: 0.7,
  rdt_machine_pompe_retour: 0.7,
  dosage_polymere: 1.5,
  rate_MES_FeCl3: (2.8 * 55.845) / 162.04,
  ratio_Psol_P: 0.95,
  ratio_PO4_Psol: 1,
  ratio_FeCl3_Pt: 162.5 / 31,
  hyp_rdt_P_particulaire: 0,
}

export default defineNode({
  id: 'discfilter',
  label: 'Discfilter',
  short: 'Discfilter',
  family: 'tertiaire',
  vba: 'F4_Discfilter_III.cls',
  description: 'Filtre à disques 10 ou 20 µm. Rendement MES fonction de la concentration entrante (10 µm) ; eaux de lavage renvoyées (Q_lavage = 2 × Q eaux sales) ; coagulation FeCl3 optionnelle pour le P.',
  choices: [
    { key: 'filtre', label: 'Seuil de coupure', default: '10', options: [{ value: '10', label: '10 µm' }, { value: '20', label: '20 µm' }] },
  ],
  params: [
    { key: 'rdt_MES', label: 'Rendement MES', unit: '-', group: 'Rendements', default: undefined, hint: 'calculé sur MES entrée si non forcé' },
    { key: 'rdt_DCO', label: 'Rendement DCO', unit: '-', group: 'Rendements', default: H.rdt_DCO },
    { key: 'rdt_DBO', label: 'Rendement DBO5', unit: '-', group: 'Rendements', default: H.rdt_DBO },
    { key: 'rdt_NK', label: 'Rendement NK', unit: '-', group: 'Rendements', default: H.rdt_NK },
    { key: 'rdt_P', label: 'Rendement P', unit: '-', group: 'Rendements', default: undefined, hint: 'calculé si non forcé' },
    { key: 'dosage_FeCl3', label: 'Dosage FeCl3 (pur)', unit: 'mg/L', group: 'Réactifs', default: undefined, hint: 'calculé sur la garantie Pt si non forcé' },
    { key: 'dosage_polymere', label: 'Dosage polymère', unit: 'mg/L', group: 'Réactifs', default: H.dosage_polymere },
    { key: 'ES_concentration_MES', label: 'MES des eaux sales', unit: 'mg/L', group: 'Eaux sales', default: undefined, hint: 'calculée si non forcée' },
    { key: 'pression_refoulement_lavage', label: 'Pression pompes de lavage', unit: 'mCE', group: 'Électricité', default: 80 },
    { key: 'pression_refoulement_retour', label: 'Pression pompes de retour', unit: 'mCE', group: 'Électricité', default: 10 },
    { key: 'rdt_pompe_lavage', label: 'Rendement global pompes de lavage', unit: '-', group: 'Électricité', default: undefined, hint: 'calculé si non forcé' },
    { key: 'rdt_pompe_retour', label: 'Rendement global pompes de retour', unit: '-', group: 'Électricité', default: undefined, hint: 'calculé si non forcé' },
  ],
  compute(ctx) {
    const { p, site, forced, choices } = ctx
    const warnings = []
    if (ctx.upstream.BA_forte) warnings.push('Discfilter derrière une boue activée forte charge : configuration signalée comme incohérente dans le classeur VBA.')

    const pass = (s) => {
      const MESin = conc(s, 'MES')
      let rdt_MES, ES_conc
      if (choices.filtre === '10') {
        rdt_MES = forced.rdt_MES ?? H.a_rdt_MES.reduce((a, c, i) => a + c * Math.pow(MESin, i), 0)
        ES_conc = forced.ES_concentration_MES != null ? forced.ES_concentration_MES / 1000 : H.a_ES_conc.reduce((a, c, i) => a + c * Math.pow(MESin, i), 0)
      } else if (MESin < 30) {
        rdt_MES = forced.rdt_MES ?? 0.55
        ES_conc = forced.ES_concentration_MES != null ? forced.ES_concentration_MES / 1000 : 0.8
      } else {
        rdt_MES = forced.rdt_MES ?? 0.65
        ES_conc = forced.ES_concentration_MES != null ? forced.ES_concentration_MES / 1000 : 1.25
      }
      const { rdt_P, dosage_FeCl3 } = precipitationP({ Q_traite: s.Q, Pt_traite: s.Pt, Q: s.Q, Pt: s.Pt, Pt_garantie: site.Pt_garantie, rdt_P_f: forced.rdt_P, dosage_FeCl3_f: forced.dosage_FeCl3, hyp: H })
      const FeCl3_flux = (dosage_FeCl3 * s.Q) / 1000
      const MES_formes = FeCl3_flux * H.rate_MES_FeCl3
      const polymere_flux = (s.Q / 1000) * p.dosage_polymere
      const MES_ES = s.MES * rdt_MES + MES_formes
      const Q_ES = MES_ES / ES_conc
      const Q_lavage = Q_ES / H.ratio_ES_lavage
      const ES = { origine: 'discfilter', Q: Q_ES, DCO: p.rdt_DCO * s.DCO, DBO: p.rdt_DBO * s.DBO, MES: MES_ES, NK: p.rdt_NK * s.NK, NH4: (s.NH4 * Q_ES) / s.Q, NO3: (s.NO3 * Q_ES) / s.Q, Pt: rdt_P * s.Pt, concentration: ES_conc }
      const out = cloneStream(s)
      out.Q = s.Q + (Q_lavage - Q_ES)
      out.DCO = s.DCO * (1 - p.rdt_DCO)
      out.DBO = s.DBO * (1 - p.rdt_DBO)
      out.MES = s.MES * (1 - rdt_MES)
      out.NK = s.NK * (1 - p.rdt_NK)
      out.NH4 = s.NH4 * (1 - Q_ES / s.Q)
      out.NO3 = s.NO3 * (1 - Q_ES / s.Q)
      out.Pt = s.Pt * (1 - rdt_P)
      return { out, ES, rdt_MES, rdt_P, dosage_FeCl3, FeCl3_flux, polymere_flux, Q_lavage, Q_ES, MESin }
    }
    const nom = pass(ctx.inNominal)
    const reel = pass(ctx.inReel)
    const stockage_Q = ctx.inReel.Q

    // ---- électricité (sur réel)
    const g = CONST.ACCELERATION_PESANTEUR_m_s2, h = CONST.NOMBRE_SECONDE_PAR_HEURE
    let elec_lavage, rdt_lav
    if (forced.rdt_pompe_lavage != null) {
      rdt_lav = forced.rdt_pompe_lavage
      elec_lavage = (reel.Q_lavage / h) * g * (p.pression_refoulement_lavage / rdt_lav)
    } else {
      const Qp = reel.Q_lavage / (H.nombre_pompe_lavage * H.duree_func)
      const Pp = (Qp / h) * g * (p.pression_refoulement_lavage / H.rdt_machine_pompe_lavage)
      const rm = rendementMoteur(Pp)
      rdt_lav = rm * H.rdt_machine_pompe_lavage
      elec_lavage = (Pp / rm) * H.nombre_pompe_lavage * H.duree_func
    }
    const elec_disques = H.duree_func * ((stockage_Q * (H.p_installee / H.Q_moyen)) / H.rdt_moteur_disques)
    let elec_retour, rdt_ret
    if (forced.rdt_pompe_retour != null) {
      rdt_ret = forced.rdt_pompe_retour
      elec_retour = (reel.Q_ES / h) * g * (p.pression_refoulement_retour / rdt_ret)
    } else {
      const Qp = reel.Q_ES / (H.nombre_pompe_retour * H.duree_func)
      const Pp = (Qp / h) * g * (p.pression_refoulement_retour / H.rdt_machine_pompe_retour)
      const rm = rendementMoteur(Pp)
      rdt_ret = rm * H.rdt_machine_pompe_retour
      elec_retour = (Pp / rm) * H.nombre_pompe_retour * H.duree_func
    }

    return {
      outNominal: nom.out,
      outReel: reel.out,
      eauxSales: reel.ES,
      reactifs: { FeCl3_kgj: reel.FeCl3_flux, polymere_kgj: reel.polymere_flux },
      results: [
        { key: 'MESin', label: 'MES entrée (réel)', unit: 'mg/L', value: reel.MESin },
        { key: 'rdt_MES', label: 'Rendement MES (réel)', unit: '-', value: reel.rdt_MES },
        { key: 'dosage_FeCl3', label: 'Dosage FeCl3 (réel)', unit: 'mg/L', value: reel.dosage_FeCl3 },
        { key: 'rdt_P', label: 'Rendement P (réel)', unit: '-', value: reel.rdt_P },
        { key: 'Q_ES', label: 'Eaux sales (réel)', unit: 'm³/j', value: reel.Q_ES },
        { key: 'MES_ES', label: 'MES eaux sales', unit: 'kg/j', value: reel.ES.MES },
        { key: 'Q_lavage', label: 'Eau de lavage (réel)', unit: 'm³/j', value: reel.Q_lavage },
        { key: 'MES_out', label: 'MES sortie (réel)', unit: 'mg/L', value: conc(reel.out, 'MES') },
      ],
      electricity: { total: elec_lavage + elec_disques + elec_retour, fixed: 0, detail: { lavage: elec_lavage, disques: elec_disques, retour: elec_retour } },
      warnings,
    }
  },
})
