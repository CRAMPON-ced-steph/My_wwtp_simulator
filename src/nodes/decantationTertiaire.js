// Port de F3_Decantation_III.cls — décantation tertiaire physico-chimique
// (FeCl3 + polymère, option microsable / recirculation type ACTIFLO).
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, precipitationP, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const H = {
  rate_MES_FeCl3: (2.8 * 55.845) / 162.04, // kg MES / kg FeCl3
  ratio_Psol_P: 0.95,
  ratio_PO4_Psol: 1, // non utilisé dans F3 (dichotomie sur P soluble)
  ratio_FeCl3_Pt: 162.5 / 31,
  hyp_rdt_P_particulaire: 0,
  elec: { coagulation: 1.425, floculation: 1.5, racleur_fixe: 1.5, polymere_prep: 0.88, polymere_pompage: 0.705, FeCl3: 0.35, microsable: 0.3, autre: (4.56 + 31.19) / 32.9 }, // Wh/m³
}
const micro = (c) => c.choices.microsable === 'oui'

export default defineNode({
  id: 'decantation-tertiaire',
  label: 'Décantation tertiaire',
  short: 'Décanteur III',
  family: 'tertiaire',
  vba: 'F3_Decantation_III.cls',
  description: 'Décantation physico-chimique tertiaire. Dosage FeCl3 calculé pour atteindre la garantie Pt (ratio molaire Fe/P fonction de Ptbr) ou forcé ; MES néoformées 2,8 kg/kg Fe.',
  choices: [
    { key: 'lamellaire', label: 'Lamellaire', default: 'oui', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'microsable', label: 'Microsable (ACTIFLO)', default: 'non', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'recirculation', label: 'Recirculation', default: 'non', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
  ],
  params: [
    { key: 'bypass', label: 'Eau by-passée', unit: '%', group: 'Hydraulique', default: 0 },
    { key: 'dosage_FeCl3', label: 'Dosage FeCl3 (produit pur)', unit: 'mg/L', group: 'Réactifs', default: undefined, hint: 'calculé sur la garantie Pt si non forcé (défaut VBA 30)' },
    { key: 'dosage_polymere', label: 'Dosage polymère (pur)', unit: 'mg/L', group: 'Réactifs', default: 0.7 },
    { key: 'rdt_MES', label: 'Rendement MES', unit: '-', group: 'Rendements', default: (c) => (micro(c) ? 0.4 : 0.35) },
    { key: 'rdt_DCO', label: 'Rendement DCO', unit: '-', group: 'Rendements', default: (c) => (micro(c) ? 0.06 : 0.05) },
    { key: 'rdt_DBO', label: 'Rendement DBO5', unit: '-', group: 'Rendements', default: (c) => (micro(c) ? 0.13 : 0.12) },
    { key: 'rdt_NK', label: 'Rendement NK', unit: '-', group: 'Rendements', default: 0.07 },
    { key: 'rdt_P', label: 'Rendement P', unit: '-', group: 'Rendements', default: undefined, hint: 'calculé si non forcé' },
    { key: 'MV_MES', label: 'MV/MES des boues', unit: '-', group: 'Boues', default: 0.4 },
    { key: 'boues_concentration', label: 'Concentration des boues', unit: 'g/L', group: 'Boues', default: (c) => (micro(c) ? 10 : 2) },
    { key: 'recirculation_taux', label: 'Taux de recirculation', unit: '-', group: 'Recirculation', default: 0.06 },
    { key: 'recirculation_nb_pompe', label: 'Nombre de pompes recirculation', unit: 'u', group: 'Recirculation', default: 2 },
    { key: 'recirculation_P_refoulement', label: 'Pression refoulement recirculation', unit: 'mCE', group: 'Recirculation', default: 25 },
    { key: 'recirculation_tps_fct', label: 'Durée de fonctionnement recirculation', unit: 'h/j', group: 'Recirculation', default: 24 },
    { key: 'recirculation_Q_unitaire', label: 'Débit unitaire recirculation', unit: 'm³/h', group: 'Recirculation', default: undefined, hint: 'calculé si non forcé' },
    { key: 'extraction_nb_pompe', label: "Nombre de pompes d'extraction", unit: 'u', group: 'Extraction', default: 2 },
    { key: 'extraction_P_refoulement', label: 'Pression refoulement extraction', unit: 'mCE', group: 'Extraction', default: 15 },
    { key: 'extraction_tps_fct', label: 'Durée de fonctionnement extraction', unit: 'h/j', group: 'Extraction', default: 24 },
    { key: 'extraction_Q_unitaire', label: 'Débit unitaire extraction', unit: 'm³/h', group: 'Extraction', default: undefined, hint: 'calculé si non forcé' },
  ],
  compute(ctx) {
    const { p, site, forced, choices } = ctx
    const warnings = []
    const bypass = (p.bypass || 0) / 100

    const pass = (s, nominal) => {
      const f = nominal ? 1 : 1 - bypass // dimensionnement sur 100 % (dimensionne_100 en dur dans le VBA)
      const Qt = f * s.Q
      const t = { DCO: f * s.DCO, DBO: f * s.DBO, MES: f * s.MES, NK: f * s.NK, Pt: f * s.Pt }
      const { rdt_P, dosage_FeCl3 } = precipitationP({ Q_traite: Qt, Pt_traite: t.Pt, Q: s.Q, Pt: s.Pt, Pt_garantie: site.Pt_garantie, rdt_P_f: forced.rdt_P, dosage_FeCl3_f: forced.dosage_FeCl3, hyp: H })
      const FeCl3_flux = (dosage_FeCl3 * Qt) / 1000
      const MES_formes = FeCl3_flux * H.rate_MES_FeCl3
      const polymere_flux = (p.dosage_polymere * Qt) / 1000
      const boues_MES = t.MES * p.rdt_MES + MES_formes
      const boues_Q = boues_MES / p.boues_concentration
      const out = cloneStream(s)
      out.Q = s.Q - boues_Q
      out.DCO = s.DCO - t.DCO * p.rdt_DCO
      out.DBO = s.DBO - t.DBO * p.rdt_DBO
      out.MES = s.MES - t.MES * p.rdt_MES
      out.NK = s.NK - t.NK * p.rdt_NK
      out.NH4 = (s.NH4 * out.Q) / s.Q
      out.NO3 = (s.NO3 * out.Q) / s.Q
      out.Pt = s.Pt - t.Pt * rdt_P
      return { out, Qt, rdt_P, dosage_FeCl3, FeCl3_flux, polymere_flux, boues_MES, boues_Q, P_precipite: rdt_P * t.Pt }
    }
    const nom = pass(ctx.inNominal, true)
    const reel = pass(ctx.inReel, false)

    // ---- électricité
    const e = H.elec
    const electricite_agitation = ((e.coagulation + e.floculation) * nom.Qt) / 1000
    const electricite_racleur = (e.racleur_fixe * nom.Qt) / 1000
    const electricite_reactifs = ((e.polymere_prep + e.polymere_pompage + e.FeCl3 + (micro(ctx) ? e.microsable : 0)) * reel.Qt) / 1000
    let electricite_recirculation = 0
    let recirc = null
    if (choices.recirculation === 'oui') {
      const Qrec = p.recirculation_taux * reel.Qt
      recirc = repartitionPompes(Qrec, { nb: p.recirculation_nb_pompe, tps: p.recirculation_tps_fct }, { nb: forced.recirculation_nb_pompe, tps: forced.recirculation_tps_fct, Q_unitaire: forced.recirculation_Q_unitaire })
      if (recirc.incoherence) warnings.push('Incohérence sur les pompes de recirculation.')
      const rdt = rendementPompeGlobal(recirc.Q_unitaire, p.recirculation_P_refoulement, 0.53 * 0.92)
      electricite_recirculation = electricitePompage(recirc.Q_unitaire, p.recirculation_P_refoulement, rdt, recirc.nb, recirc.tps)
    }
    const extr = repartitionPompes(reel.boues_Q, { nb: p.extraction_nb_pompe, tps: p.extraction_tps_fct }, { nb: forced.extraction_nb_pompe, tps: forced.extraction_tps_fct, Q_unitaire: forced.extraction_Q_unitaire })
    if (extr.incoherence) warnings.push("Incohérence sur les pompes d'extraction.")
    const rdt_extr = rendementPompeGlobal(extr.Q_unitaire, p.extraction_P_refoulement, 0.7)
    const electricite_extraction = electricitePompage(extr.Q_unitaire, p.extraction_P_refoulement, rdt_extr, extr.nb, extr.tps)
    const electricite_autre = (e.autre * reel.Qt) / 1000
    const total = electricite_racleur + electricite_recirculation + electricite_extraction + electricite_agitation + electricite_reactifs + electricite_autre

    return {
      outNominal: nom.out,
      outReel: reel.out,
      sludge: { origine: 'III_decantation', Q: reel.boues_Q, MES: reel.boues_MES, concentration: p.boues_concentration, MV_MES: p.MV_MES },
      reactifs: { FeCl3_kgj: reel.FeCl3_flux, polymere_kgj: reel.polymere_flux },
      results: [
        { key: 'dosage_FeCl3_nom', label: 'Dosage FeCl3 (nominal)', unit: 'mg/L', value: nom.dosage_FeCl3 },
        { key: 'dosage_FeCl3', label: 'Dosage FeCl3 (réel)', unit: 'mg/L', value: reel.dosage_FeCl3 },
        { key: 'FeCl3_flux', label: 'Consommation FeCl3 pur (réel)', unit: 'kg/j', value: reel.FeCl3_flux },
        { key: 'polymere_flux', label: 'Consommation polymère (réel)', unit: 'kg/j', value: reel.polymere_flux },
        { key: 'rdt_P', label: 'Rendement P (réel)', unit: '-', value: reel.rdt_P },
        { key: 'boues_MES', label: 'Boues extraites (réel)', unit: 'kg MES/j', value: reel.boues_MES },
        { key: 'boues_Q', label: 'Boues extraites (réel)', unit: 'm³/j', value: reel.boues_Q },
        { key: 'Pt_out', label: 'Pt sortie (réel)', unit: 'mg/L', value: conc(reel.out, 'Pt') },
      ],
      electricity: { total, fixed: electricite_racleur + electricite_agitation, detail: { racleur: electricite_racleur, agitation: electricite_agitation, reactifs: electricite_reactifs, recirculation: electricite_recirculation, extraction: electricite_extraction, autre: electricite_autre } },
      warnings,
    }
  },
})
