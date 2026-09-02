// ---------------------------------------------------------------------------
// Orchestration des quatre moteurs.
//
// Équivalent de MOD_ProgrammePrincipal.prgm_principal : la file eau alimente la
// file boues, les deux alimentent les utilités, et le bloc transverse agrège le
// tout. L'ordre compte — le bilan électrique a besoin des trois files, et
// l'empreinte CO2 a besoin du bilan électrique.
// ---------------------------------------------------------------------------
import { runChain } from './engine.js'
import { runSludgeChain, apportsDepuisFileEau } from './sludgeEngine.js'
import { runUtilities, contexteDepuisFilieres } from './utilityEngine.js'
import { repartitionElectrique } from '../nodes-transverse/bilan.js'

/**
 * Exécute une simulation complète.
 * filiere = { eau: [...], boues: [...], utilites: [...], transverse: [...] }
 * registres = { eau, boues, utilites, transverse }
 */
export function simuler(filiere, registres, site) {
  const eau = runChain(filiere.eau ?? [], registres.eau, site)
  const boues = runSludgeChain(filiere.boues ?? [], registres.boues, site, apportsDepuisFileEau(eau, registres.eau))
  const contexte = contexteDepuisFilieres(site, eau, boues)
  const utilites = runUtilities(filiere.utilites ?? [], registres.utilites, site, contexte)

  // le bloc transverse voit la répartition électrique des trois files
  contexte.bilan = repartitionElectrique(eau, boues, utilites, registres)
  const transverse = runUtilities(filiere.transverse ?? [], registres.transverse, site, contexte, {
    // le bilan électrique doit être disponible pour l'empreinte CO2
    apresChaqueEtape: (step) => {
      // chaque agrégat alimente le contexte du suivant : le bilan électrique
      // sert à l'empreinte CO2 et aux OPEX, les OPEX servent au ROI
      if (step.bilan) contexte.bilanElectrique = step.bilan
      if (step.opex) contexte.opex = step.opex
    },
  })

  const bilanStep = transverse.steps.find((s) => s.bilan)
  const co2Step = transverse.steps.find((s) => s.co2)
  const opexStep = transverse.steps.find((s) => s.opex)
  const roiStep = transverse.steps.find((s) => s.roi)
  return {
    site,
    eau,
    boues,
    utilites,
    transverse,
    contexte,
    bilan: bilanStep?.bilan ?? null,
    co2: co2Step?.co2 ?? null,
    opex: opexStep?.opex ?? null,
    roi: roiStep?.roi ?? null,
    electricite: {
      eau: eau.electricityTotal,
      boues: boues.electricityTotal,
      utilites: utilites.electricityTotal - utilites.electricityProduite,
      total: eau.electricityTotal + boues.electricityTotal + utilites.electricityTotal - utilites.electricityProduite,
    },
  }
}
