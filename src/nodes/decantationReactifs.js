// Port de D2_Decantation_reactif.cls — décantation primaire physico-chimique
// (MULTIFLO DUO/TRIO, ACTIFLO) : coagulation FeCl3, floculation polymère,
// option microsable et recirculation.
import { defineNode } from '../core/engine.js'
import { cloneStream, conc } from '../core/stream.js'
import { CONST, HYP, coagulationFloculationRdtPSoluble, ratioMolairePtbr, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const H = {
  ratio_Psol_P: HYP.dec_I_ratio_Psol_P, // 0.7
  ratio_PO4_Psol: HYP.dec_I_ratio_PO4_Psol, // 0.8
  ratio_FeCl3_Pt: 162.5 / 31,
  MV_MES_reference: HYP.dec_I_boues_MV_MES, // 0.68
  MM_P: 31,
  MM_FeO3H3: 55.85 + 3 * 17, // Fe(OH)3
  MM_FePO4: 55.85 + 31 + 64,
  elec: { coagulation: 1.425, floculation: 1.5, polymere_prep: 0.88, polymere_pompage: 0.705, FeCl3: 0.35, microsable: 0.3, divers: 1.0866261398176291 },
  dec_diametre_limite: HYP.dec_I_diametre_limite_m,
}
const micro = (c) => c.choices.microsable === 'oui'

export default defineNode({
  id: 'decantation-reactifs',
  label: 'Décantation avec réactifs',
  short: 'Décanteur I réactifs',
  family: 'primaire',
  vba: 'D2_Decantation_reactif.cls',
  description: "Décantation primaire lestée physico-chimique (MULTIFLO DUO/TRIO, ACTIFLO). Rendement P calculé par dichotomie sur la dose de FeCl3 ; MES néoformées Fe(OH)3 + FePO4 ; rendements MES/DCO/DBO élevés.",
  choices: [
    { key: 'lamellaire', label: 'Lamellaire', default: 'oui', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'microsable', label: 'Microsable (ACTIFLO)', default: 'non', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'recirculation', label: 'Recirculation', default: 'non', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'biostyr_aval', label: 'Biofiltration (Biostyr) à l’aval', default: 'non', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
    { key: 'dimensionnement', label: 'Dimensionnement du poste', default: '100', options: [{ value: '100', label: 'sur 100 % du débit' }, { value: 'bypass', label: 'sur le % non by-passé' }] },
  ],
  params: [
    { key: 'bypass', label: 'Eau by-passée au primaire', unit: '%', group: 'Hydraulique', default: 0 },
    { key: 'dosage_FeCl3', label: 'Dosage FeCl3 (produit pur)', unit: 'mg/L', group: 'Réactifs', default: 30 },
    { key: 'dosage_polymere', label: 'Dosage polymère (pur)', unit: 'mg/L', group: 'Réactifs', default: 0.7 },
    { key: 'nb_ouvrages', label: "Nombre d'ouvrages", unit: '-', group: 'Ouvrage', default: 1 },
    { key: 'V_miroir_limite', label: 'Vitesse au miroir limite (temps de pluie)', unit: 'm/h', group: 'Ouvrage', default: (c) => (c.choices.biostyr_aval === 'oui' ? (micro(c) ? 80 : 30) : micro(c) ? 120 : 40), hint: 'Aquademy / CAM' },
    { key: 'S_miroir', label: 'Surface au miroir', unit: 'm²', group: 'Ouvrage', default: undefined, hint: 'calculée si non forcée' },
    { key: 'hauteur', label: 'Hauteur du décanteur', unit: 'm', group: 'Ouvrage', default: 4 },
    { key: 'rdt_MES', label: 'Rendement MES', unit: '-', group: 'Rendements', default: (c) => (micro(c) ? 0.87 : 0.83) },
    { key: 'rdt_DCO', label: 'Rendement DCO', unit: '-', group: 'Rendements', default: (c) => (micro(c) ? 0.6 : 0.55) },
    { key: 'rdt_DBO', label: 'Rendement DBO5', unit: '-', group: 'Rendements', default: (c) => (micro(c) ? 0.57 : 0.53) },
    { key: 'rdt_NK', label: 'Rendement NK', unit: '-', group: 'Rendements', default: 0.2 },
    { key: 'rdt_P', label: 'Rendement P', unit: '-', group: 'Rendements', default: undefined, hint: 'calculé (dichotomie sur la dose FeCl3) si non forcé' },
    { key: 'MV_MES', label: 'MV/MES des boues', unit: '-', group: 'Boues', default: undefined, hint: 'calculé (0,68 corrigé des MES néoformées) si non forcé' },
    { key: 'boues_concentration', label: 'Concentration des boues', unit: 'g/L', group: 'Boues', default: 30 },
    { key: 'boues_MES', label: 'Boues extraites', unit: 'kg MES/j', group: 'Boues', default: undefined, hint: 'calculé si non forcé' },
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
      const f = nominal && choices.dimensionnement === '100' ? 1 : 1 - bypass
      const Qt = f * s.Q
      const t = { DCO: f * s.DCO, DBO: f * s.DBO, MES: f * s.MES, NK: f * s.NK, Pt: f * s.Pt }
      const Q_retour_t = f * site.Q_retour
      // rendement P particulaire selon Pt entrant
      const hyp_rdt_P_particulaire = (t.Pt / Qt) * 1000 < 10 ? 0.8 : 0.85
      // calcul du P
      let rdt_P, dosage_FeCl3 = p.dosage_FeCl3, P_soluble_removed, rm
      if (forced.rdt_P != null) {
        rdt_P = forced.rdt_P
        P_soluble_removed = rdt_P * t.Pt - (1 - H.ratio_Psol_P) * t.Pt * hyp_rdt_P_particulaire
        if (P_soluble_removed <= 0) {
          P_soluble_removed = 0
          rm = 1
        } else if (forced.dosage_FeCl3 != null) {
          rm = dosage_FeCl3 / (H.ratio_FeCl3_Pt * ((P_soluble_removed / Qt) * 1000))
        } else {
          const P_outlet = ((t.Pt * (1 - rdt_P)) / Qt) * 1000
          rm = ratioMolairePtbr((P_soluble_removed / Qt) * 1000, Math.max(P_outlet, 0.1))
          dosage_FeCl3 = rm * H.ratio_FeCl3_Pt * ((P_soluble_removed / Qt) * 1000)
        }
      } else {
        const rdt_P_soluble = coagulationFloculationRdtPSoluble(Qt, t.Pt, dosage_FeCl3, H.ratio_Psol_P, H.ratio_PO4_Psol, H.ratio_FeCl3_Pt)
        rdt_P = H.ratio_Psol_P * H.ratio_PO4_Psol * rdt_P_soluble + (1 - H.ratio_Psol_P) * hyp_rdt_P_particulaire
        P_soluble_removed = rdt_P_soluble * H.ratio_Psol_P * H.ratio_PO4_Psol * t.Pt
        rm = P_soluble_removed > 0 ? dosage_FeCl3 / (H.ratio_FeCl3_Pt * ((P_soluble_removed / Qt) * 1000)) : 1
      }
      const FeCl3_flux = (dosage_FeCl3 * Qt) / 1000
      const MES_formes = ((rm - 1) * H.MM_FeO3H3 / H.MM_P + H.MM_FePO4 / H.MM_P) * P_soluble_removed
      const polymere_flux = (p.dosage_polymere * Qt) / 1000
      // MV/MES corrigé des MES néoformées
      const inlet_MV = H.MV_MES_reference * s.MES
      const MV_MES = forced.MV_MES ?? inlet_MV / (s.MES + MES_formes)
      // surface au miroir
      const S_miroir = forced.S_miroir ?? (site.pointe_TP * (Qt - Q_retour_t) + Q_retour_t) / (CONST.NOMBRE_HEURE_PAR_JOUR * p.V_miroir_limite)
      // boues et sorties
      let boues_MES = t.MES * p.rdt_MES + MES_formes
      if (!nominal && forced.boues_MES != null) boues_MES = forced.boues_MES
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
      return { out, Qt, S_miroir, rdt_P, dosage_FeCl3, rm, FeCl3_flux, polymere_flux, MES_formes, MV_MES, boues_MES, boues_Q, P_precipite: P_soluble_removed }
    }
    const nom = pass(ctx.inNominal, true)
    const reel = pass(ctx.inReel, false)
    const V_miroir_recalc = (site.pointe_TP * (nom.Qt - site.Q_retour) + site.Q_retour) / (CONST.NOMBRE_HEURE_PAR_JOUR * nom.S_miroir)
    const tps_retention = (nom.S_miroir * p.hauteur) / (reel.Qt / CONST.NOMBRE_HEURE_PAR_JOUR)

    // ---- électricité
    const e = H.elec
    const electricite_agitation = ((e.coagulation + e.floculation) * nom.Qt) / 1000
    const S_unit = nom.S_miroir / p.nb_ouvrages
    const Pw = S_unit < CONST.PI * Math.pow(H.dec_diametre_limite / 2, 2) ? HYP.dec_I_Pw_racleur_inf_Dlim_kW : HYP.dec_I_Pw_racleur_sup_Dlim_kW
    const electricite_racleur = p.nb_ouvrages * Pw * CONST.NOMBRE_HEURE_PAR_JOUR
    const electricite_reactifs = ((e.polymere_prep + e.polymere_pompage + e.FeCl3 + (micro(ctx) ? e.microsable : 0)) * reel.Qt) / 1000
    let electricite_recirculation = 0
    if (choices.recirculation === 'oui') {
      const Qrec = p.recirculation_taux * reel.Qt
      const rep = repartitionPompes(Qrec, { nb: p.recirculation_nb_pompe, tps: p.recirculation_tps_fct }, { nb: forced.recirculation_nb_pompe, tps: forced.recirculation_tps_fct, Q_unitaire: forced.recirculation_Q_unitaire })
      if (rep.incoherence) warnings.push('Incohérence sur les pompes de recirculation.')
      const rdt = rendementPompeGlobal(rep.Q_unitaire, p.recirculation_P_refoulement, 0.53 * 0.92)
      electricite_recirculation = electricitePompage(rep.Q_unitaire, p.recirculation_P_refoulement, rdt, rep.nb, rep.tps)
    }
    const extr = repartitionPompes(reel.boues_Q, { nb: p.extraction_nb_pompe, tps: p.extraction_tps_fct }, { nb: forced.extraction_nb_pompe, tps: forced.extraction_tps_fct, Q_unitaire: forced.extraction_Q_unitaire })
    if (extr.incoherence) warnings.push("Incohérence sur les pompes d'extraction.")
    const rdt_extr = rendementPompeGlobal(extr.Q_unitaire, p.extraction_P_refoulement, 0.7)
    const electricite_extraction = electricitePompage(extr.Q_unitaire, p.extraction_P_refoulement, rdt_extr, extr.nb, extr.tps)
    const electricite_autre = (e.divers * reel.Qt) / 1000
    const total = electricite_racleur + electricite_recirculation + electricite_extraction + electricite_agitation + electricite_reactifs + electricite_autre

    return {
      outNominal: nom.out,
      outReel: reel.out,
      sludge: { origine: 'I_reactif', Q: reel.boues_Q, MES: reel.boues_MES, concentration: p.boues_concentration, MV_MES: reel.MV_MES },
      reactifs: { FeCl3_kgj: reel.FeCl3_flux, polymere_kgj: reel.polymere_flux },
      results: [
        { key: 'S_miroir', label: 'Surface au miroir', unit: 'm²', value: nom.S_miroir },
        { key: 'V_recalc', label: 'Vitesse au miroir limite recalculée', unit: 'm/h', value: V_miroir_recalc },
        { key: 'tps_retention', label: 'Temps de rétention (réel)', unit: 'h', value: tps_retention },
        { key: 'rdt_P', label: 'Rendement P (réel)', unit: '-', value: reel.rdt_P },
        { key: 'rm', label: 'Ratio molaire Fe/P (réel)', unit: '-', value: reel.rm },
        { key: 'FeCl3', label: 'Consommation FeCl3 pur (réel)', unit: 'kg/j', value: reel.FeCl3_flux },
        { key: 'polymere', label: 'Consommation polymère (réel)', unit: 'kg/j', value: reel.polymere_flux },
        { key: 'MES_formes', label: 'MES néoformées (réel)', unit: 'kg/j', value: reel.MES_formes },
        { key: 'MV_MES', label: 'MV/MES des boues (réel)', unit: '-', value: reel.MV_MES },
        { key: 'boues_MES', label: 'Boues extraites (réel)', unit: 'kg MES/j', value: reel.boues_MES },
        { key: 'boues_Q', label: 'Boues extraites (réel)', unit: 'm³/j', value: reel.boues_Q },
        { key: 'Pt_out', label: 'Pt sortie (réel)', unit: 'mg/L', value: conc(reel.out, 'Pt') },
      ],
      electricity: { total, fixed: electricite_racleur + electricite_agitation, detail: { racleur: electricite_racleur, agitation: electricite_agitation, reactifs: electricite_reactifs, recirculation: electricite_recirculation, extraction: electricite_extraction, autre: electricite_autre } },
      warnings,
    }
  },
})
