// Port de H1_Decantation_eaux_sales.cls — décantation physico-chimique des
// eaux sales (eaux de lavage des biofiltres, Discfilter, filtres à sable).
// Le nœud collecte les flux `eauxSales` émis par les procédés placés en amont
// dans la filière, les traite (FeCl3 + polymère) et renvoie le surnageant en
// tête de station (boucle iteration_ES du moteur). Les boues partent en file
// boues. L'eau de la filière traverse ce nœud sans modification.
import { defineNode } from '../core/engine.js'
import { cloneStream } from '../core/stream.js'
import { CONST, HYP, repartitionPompes, rendementPompeGlobal, electricitePompage } from '../core/hypotheses.js'

const H = {
  rate_MES_FeCl3: (2.8 * 55.845) / 162.04, // kg MES néoformées / kg FeCl3
  elec: { coagulation: 1.425, floculation: 1.5, polymere_prep: 0.88, polymere_pompage: 0.705, FeCl3: 0.35, divers: 1.0866261398176291 },
  rate_puissance_fixe: 0.1, // Wh/m³ sur le nominal (râcleur)
  rate_puissance_debit: 1, // Wh/m³ sur le réel
}
const ES_KEYS = ['Q', 'DCO', 'DBO', 'MES', 'NK', 'NH4', 'NO3', 'Pt']

export default defineNode({
  id: 'decantation-eaux-sales',
  label: 'Décantation des eaux sales',
  short: 'Décanteur eaux sales',
  family: 'retours',
  vba: 'H1_Decantation_eaux_sales.cls',
  description: "Décantation physico-chimique des eaux de lavage collectées en amont (biofiltres, Discfilter, filtres à sable). Le surnageant retourne en tête de filière, les boues partent en file boues. À placer après les procédés qui produisent des eaux sales.",
  choices: [
    { key: 'lamellaire', label: 'Lamellaire', default: 'oui', options: [{ value: 'oui', label: 'oui' }, { value: 'non', label: 'non' }] },
  ],
  params: [
    { key: 'dosage_FeCl3', label: 'Dosage FeCl3 (produit pur)', unit: 'mg/L', group: 'Réactifs', default: 30 },
    { key: 'dosage_polymere', label: 'Dosage polymère (pur)', unit: 'mg/L', group: 'Réactifs', default: 0.7 },
    { key: 'rdt_MES', label: 'Rendement MES', unit: '-', group: 'Rendements', default: 0.85 },
    { key: 'rdt_DCO', label: 'Rendement DCO', unit: '-', group: 'Rendements', default: 0.85 },
    { key: 'rdt_DBO', label: 'Rendement DBO5', unit: '-', group: 'Rendements', default: 0.88 },
    { key: 'rdt_NK', label: 'Rendement NK', unit: '-', group: 'Rendements', default: 0.83 },
    { key: 'rdt_P', label: 'Rendement P', unit: '-', group: 'Rendements', default: 0.65 },
    { key: 'MV_MES', label: 'MV/MES des boues', unit: '-', group: 'Boues', default: undefined, hint: 'calculé (bilan MV des eaux sales) si non forcé' },
    { key: 'boues_concentration', label: 'Concentration des boues', unit: 'g/L', group: 'Boues', default: 30 },
    { key: 'extraction_nb_pompe', label: "Nombre de pompes d'extraction", unit: 'u', group: 'Extraction', default: 2 },
    { key: 'extraction_P_refoulement', label: 'Pression refoulement extraction', unit: 'mCE', group: 'Extraction', default: 15 },
    { key: 'extraction_tps_fct', label: 'Durée de fonctionnement extraction', unit: 'h/j', group: 'Extraction', default: 24 },
    { key: 'extraction_Q_unitaire', label: 'Débit unitaire extraction', unit: 'm³/h', group: 'Extraction', default: undefined, hint: 'calculé si non forcé' },
  ],
  compute(ctx) {
    const { p, forced } = ctx
    const warnings = []
    const agreger = (liste) => {
      const t = { Q: 0, DCO: 0, DBO: 0, MES: 0, NK: 0, NH4: 0, NO3: 0, Pt: 0, MV: 0 }
      for (const es of liste) {
        for (const k of ES_KEYS) t[k] += es[k] || 0
        t.MV += (es.MV_MES ?? 0.8) * (es.MES || 0)
      }
      return t
    }
    const traiter = (t) => {
      const FeCl3_flux = (p.dosage_FeCl3 * t.Q) / 1000
      const MES_formes = FeCl3_flux * H.rate_MES_FeCl3
      const MV_MES = forced.MV_MES ?? (t.MES + MES_formes > 0 ? t.MV / (t.MES + MES_formes) : 0.8)
      const boues_MES = (t.MES + MES_formes) * p.rdt_MES
      const boues_Q = boues_MES / p.boues_concentration
      const retour = {
        Q: t.Q - boues_Q,
        DCO: t.DCO * (1 - p.rdt_DCO),
        DBO: t.DBO * (1 - p.rdt_DBO),
        MES: (t.MES + MES_formes) * (1 - p.rdt_MES),
        NK: t.NK * (1 - p.rdt_NK),
        NH4: (t.NH4 * (t.Q - boues_Q)) / (t.Q || 1),
        NO3: (t.NO3 * (t.Q - boues_Q)) / (t.Q || 1),
        Pt: t.Pt * (1 - p.rdt_P),
      }
      return { retour, FeCl3_flux, MES_formes, MV_MES, boues_MES, boues_Q, polymere_flux: (p.dosage_polymere * t.Q) / 1000 }
    }
    const tNom = agreger(ctx.esAmontNominal || [])
    const tReel = agreger(ctx.esAmontReel || [])
    if (tReel.Q === 0) warnings.push("Aucune eau sale collectée : placer ce nœud en aval des procédés qui en produisent (biofiltres, Discfilter, filtre à sable).")
    const nom = traiter(tNom)
    const reel = traiter(tReel)

    // électricité (structure D2/H1)
    const e = H.elec
    const electricite_agitation = ((e.coagulation + e.floculation) * tNom.Q) / 1000
    const electricite_racleur = (H.rate_puissance_fixe * tNom.Q + H.rate_puissance_debit * tReel.Q) / 1000
    const electricite_reactifs = ((e.polymere_prep + e.polymere_pompage + e.FeCl3) * tReel.Q) / 1000
    const extr = repartitionPompes(reel.boues_Q, { nb: p.extraction_nb_pompe, tps: p.extraction_tps_fct }, { nb: forced.extraction_nb_pompe, tps: forced.extraction_tps_fct, Q_unitaire: forced.extraction_Q_unitaire })
    if (reel.boues_Q > 0 && extr.incoherence) warnings.push("Incohérence sur les pompes d'extraction.")
    const rdt_extr = rendementPompeGlobal(extr.Q_unitaire, p.extraction_P_refoulement, 0.7)
    const electricite_extraction = reel.boues_Q > 0 ? electricitePompage(extr.Q_unitaire, p.extraction_P_refoulement, rdt_extr, extr.nb, extr.tps) : 0
    const electricite_autre = (e.divers * tReel.Q) / 1000
    const total = electricite_racleur + electricite_extraction + electricite_agitation + electricite_reactifs + electricite_autre

    return {
      outNominal: cloneStream(ctx.inNominal),
      outReel: cloneStream(ctx.inReel),
      retourTraite: { nominal: nom.retour, reel: reel.retour, esConsommees: (ctx.esAmontReel || []).length },
      sludge: { origine: 'ES_decantees', Q: reel.boues_Q, MES: reel.boues_MES, concentration: p.boues_concentration, MV_MES: reel.MV_MES },
      reactifs: { FeCl3_kgj: reel.FeCl3_flux, polymere_kgj: reel.polymere_flux },
      results: [
        { key: 'Q_ES', label: 'Eaux sales collectées (réel)', unit: 'm³/j', value: tReel.Q },
        { key: 'MES_ES', label: 'MES collectées (réel)', unit: 'kg/j', value: tReel.MES },
        { key: 'MES_formes', label: 'MES néoformées (réel)', unit: 'kg/j', value: reel.MES_formes },
        { key: 'FeCl3', label: 'Consommation FeCl3 pur (réel)', unit: 'kg/j', value: reel.FeCl3_flux },
        { key: 'polymere', label: 'Consommation polymère (réel)', unit: 'kg/j', value: reel.polymere_flux },
        { key: 'boues', label: 'Boues extraites (réel)', unit: 'kg MES/j', value: reel.boues_MES },
        { key: 'boues_Q', label: 'Boues extraites (réel)', unit: 'm³/j', value: reel.boues_Q },
        { key: 'retour_Q', label: 'Surnageant renvoyé en tête (réel)', unit: 'm³/j', value: reel.retour.Q },
        { key: 'retour_MES', label: 'MES renvoyées en tête (réel)', unit: 'kg/j', value: reel.retour.MES },
      ],
      electricity: { total, fixed: electricite_racleur + electricite_agitation, detail: { racleur: electricite_racleur, agitation: electricite_agitation, reactifs: electricite_reactifs, extraction: electricite_extraction, autre: electricite_autre } },
      warnings,
    }
  },
})
