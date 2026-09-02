// ---------------------------------------------------------------------------
// Routines partagées par les procédés de la file boues.
//
// Les classes z_* du classeur répètent trois blocs quasiment à l'identique :
//  - la lecture des entrées (jusqu'à trois, chacune pointant une étape amont et
//    un type de boue, avec une part de flux admise) ;
//  - l'arbre de répartition débit unitaire / nombre de pompes / durée de
//    fonctionnement, avec son contrôle de cohérence à 10 % ;
//  - le rendement global d'une pompe à boues, machine à 0,4 multipliée par un
//    rendement moteur suivant une loi logarithmique.
// Ils sont factorisés ici plutôt que recopiés dans chaque nœud.
// ---------------------------------------------------------------------------
import { ETAPE, P, SOL, NB_TYPES, TYPE } from '../core/sludge.js'
import { CONST, rendementMoteur } from '../core/hypotheses.js'

export const H_POMPE = {
  critere_incoherence: 0.1,
  rdt_pompe: 0.4,
  tps_fct_defaut: 5, // h/j
}

/** correspondance entre une valeur de choix d'entrée et (étape, types de boue) */
export function resoudreEntree(valeur) {
  if (!valeur || valeur === 'aucune') return null
  if (valeur === 'toutes') return { etape: ETAPE.inlet, types: null }
  if (valeur === 'I') return { etape: ETAPE.inlet, types: [TYPE.I] }
  if (valeur === 'II') return { etape: ETAPE.inlet, types: [TYPE.II] }
  if (valeur === 'III') return { etape: ETAPE.inlet, types: [TYPE.III] }
  if (valeur === 'graisses') return { etape: ETAPE.inlet, types: [TYPE.graisses] }
  let m = /^epaississeur_(\d)$/.exec(valeur)
  if (m) return { etape: ETAPE.epaissies[Number(m[1])], types: null }
  if (valeur === 'digerees') return { etape: ETAPE.digerees, types: null }
  if (valeur === 'athos') return { etape: ETAPE.athos, types: null }
  m = /^deshydratees_(\d)$/.exec(valeur)
  if (m) return { etape: ETAPE.deshydratees[Number(m[1])], types: null }
  m = /^chaulees_(\d)$/.exec(valeur)
  if (m) return { etape: ETAPE.chaulees[Number(m[1])], types: null }
  m = /^sechees_(\d)$/.exec(valeur)
  if (m) return { etape: ETAPE.sechees[Number(m[1])], types: null }
  m = /^sechees_bioco_(\d)$/.exec(valeur)
  if (m) return { etape: ETAPE.sechees_bioco[Number(m[1])], types: null }
  m = /^sechees_inos_(\d)$/.exec(valeur)
  if (m) return { etape: ETAPE.sechees_inos[Number(m[1])], types: null }
  return null
}

/** construit la liste des entrées à partir des choix et de leur part admise */
export function entreesDepuisChoix(choices, p, cles) {
  const out = []
  cles.forEach((cle, i) => {
    const e = resoudreEntree(choices[cle])
    if (!e) return
    out.push({ ...e, ratio: p[`ratio_admis_${i + 1}`] ?? 1 })
  })
  return out
}

/**
 * Lit les flux amont d'un procédé et renvoie une ligne par couple
 * (étape, type de boue) effectivement présent. La part admise est déjà
 * appliquée ; `flux_in` porte la fraction de flux consommée, à retrancher de
 * `verif_flux` en amont une fois l'écriture aval faite.
 */
export function lireEntrees(table, soluble, entrees) {
  const out = []
  for (const e of entrees) {
    const types = e.types ?? Array.from({ length: NB_TYPES }, (_, i) => i + 1)
    for (const j of types) {
      const src = table[e.etape][j]
      if (!(src[P.MES] > 0)) continue
      // un flux déjà entièrement consommé par un procédé amont n'est plus
      // disponible : sans ce contrôle, une entrée « toutes origines » le
      // reprendrait une seconde fois
      if (!(src[P.verif_flux] > 1e-9)) continue
      out.push({
        etape: e.etape,
        j,
        src,
        sol: soluble[e.etape][j],
        MES: e.ratio * src[P.MES],
        Q: e.ratio * src[P.Q],
        MV_MES: src[P.MV_MES],
        ratios: {
          DCO: src[P.ratio_DCO_MES],
          DBO: src[P.ratio_DBO_MES],
          NK: src[P.ratio_NK_MES],
          Pt: src[P.ratio_Pt_MES],
        },
        flux_in: src[P.flux_in] * e.ratio,
      })
    }
  }
  return out
}

/**
 * Répartition débit unitaire / nombre de pompes / durée, telle qu'écrite dans
 * les classes de la file boues. Le débit unitaire forcé prime ; à défaut le
 * nombre et la durée sont pris, et le débit s'en déduit. Quand la durée
 * calculée dépasse le plafond, elle y est ramenée et le nombre de pompes est
 * relevé à l'entier supérieur.
 */
export function repartitionPompage(Q, nbDefaut, tpsDefaut, forced, prefixe, tpsMax = H_POMPE.tps_fct_defaut) {
  const nbF = forced[`${prefixe}_pompe_nb`]
  const tpsF = forced[`${prefixe}_tps_fonctionnement`]
  const QuF = forced[`${prefixe}_Q_unitaire`]
  let nb = nbF ?? nbDefaut
  let tps = tpsF ?? tpsDefaut
  let Qu
  const ceilDiv = (a, b) => (b > 0 ? Math.max(1, Math.ceil(a / b - 1e-12)) : 0)
  if (QuF != null) {
    Qu = QuF
    if (tpsF != null) {
      if (nbF == null) nb = ceilDiv(Q, tps * Qu)
    } else {
      tps = nb * Qu > 0 ? Q / (nb * Qu) : 0
      if (tps > tpsMax) {
        tps = tpsMax
        if (nbF == null) {
          nb = ceilDiv(Q, tps * Qu)
          tps = nb * Qu > 0 ? Q / (nb * Qu) : 0
        }
      }
    }
  } else {
    Qu = nb * tps > 0 ? Q / (nb * tps) : 0
  }
  const incoherence = Q > 0 && Math.abs((Qu * nb * tps - Q) / Q) > H_POMPE.critere_incoherence
  return { nb, tps, Qu, incoherence }
}

/** rendement global d'une pompe à boues : machine 0,4 × rendement moteur */
export function rdtPompeBoues(Qu, P_mCE) {
  const puissance = ((Qu / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * P_mCE) / H_POMPE.rdt_pompe
  return H_POMPE.rdt_pompe * rendementMoteur(puissance)
}

/** énergie de pompage d'un poste (kWh/j) */
export function elecPompage(Qu, nb, tps, P_mCE, rdt) {
  if (!(rdt > 0)) return 0
  return ((Qu / CONST.NOMBRE_SECONDE_PAR_HEURE) * CONST.ACCELERATION_PESANTEUR_m_s2 * P_mCE / rdt) * nb * tps
}

/**
 * Écrit une étape aval à partir d'un cumul par type de boue et répartit la
 * pollution soluble entre les boues sortantes et le retour en tête, au prorata
 * des débits — motif commun à la déshydratation et au chaulage.
 */
export function ecrireEtape(table, etape, parType) {
  for (const [jStr, t] of Object.entries(parType)) {
    const j = Number(jStr)
    const dst = table[etape][j]
    dst[P.origine] = t.origine ?? dst[P.origine]
    dst[P.MES] = t.MES
    dst[P.Q] = t.Q
    dst[P.MV_MES] = t.MES > 0 ? (t.MV ?? 0) / t.MES : 0
    dst[P.ratio_DCO_MES] = t.MES > 0 ? (t.DCO ?? 0) / t.MES : 0
    dst[P.ratio_DBO_MES] = t.MES > 0 ? (t.DBO ?? 0) / t.MES : 0
    dst[P.ratio_NK_MES] = t.MES > 0 ? (t.NK ?? 0) / t.MES : 0
    dst[P.ratio_Pt_MES] = t.MES > 0 ? (t.Pt ?? 0) / t.MES : 0
    dst[P.flux_in] = t.flux_in
    dst[P.verif_flux] = t.flux_in
  }
}
