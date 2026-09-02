// Utilités : désodorisation, turbine, gestion des réactifs.
import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { runSludgeChain, apportsDepuisFileEau } from '../src/core/sludgeEngine.js'
import { SLUDGE_REGISTRY } from '../src/nodes-boues/index.js'
import { runUtilities, contexteDepuisFilieres } from '../src/core/utilityEngine.js'
import { UTILITY_REGISTRY } from '../src/nodes-utilites/index.js'
import { DEFAULT_SITE } from '../src/core/stream.js'

function essai(titre, chainEau, chainBoues, chainUtil, site = DEFAULT_SITE) {
  console.log('\n══════ ' + titre)
  const eau = runChain(chainEau, REGISTRY, site)
  const boues = runSludgeChain(chainBoues, SLUDGE_REGISTRY, site, apportsDepuisFileEau(eau, REGISTRY))
  const contexte = contexteDepuisFilieres(site, eau, boues)
  console.log(`Contexte : Q nominal ${contexte.Q_nominal} m³/j, H2S strippé ${contexte.HS_strippe_kgj.toFixed(1)} kg/j, biogaz ${contexte.biogaz_Nm3j.toFixed(0)} Nm³/j`)
  console.log('Réactifs des filières : ' + JSON.stringify(Object.fromEntries(Object.entries(contexte.reactifs).map(([k, v]) => [k, Math.round(v)]))))
  const r = runUtilities(chainUtil, UTILITY_REGISTRY, site, contexte)
  for (const s of r.steps) {
    console.log(`### ${s.label} | élec=${(s.electricity.total || 0).toFixed(0)} kWh/j`)
    for (const x of s.results) console.log('   · ' + x.label + ' : ' + (typeof x.value === 'number' ? x.value.toFixed(2) : x.value) + ' ' + x.unit)
    for (const w of new Set(s.warnings)) console.log('   ⚠ ' + w)
    for (const [k, v] of Object.entries(s.electricity)) if (!['total','fixed','detail','produite'].includes(k)) console.log('   ? ' + k)
    for (const [k, v] of Object.entries(s.results)) if (!Number.isFinite(v.value) && typeof v.value !== 'string') console.log('   ✗ valeur non finie : ' + v.label)
  }
  console.log(`ÉLEC utilités : ${r.electricityTotal.toFixed(0)} kWh/j consommés, ${r.electricityProduite.toFixed(0)} kWh/j produits`)
  const totalStation = eau.electricityTotal + boues.electricityTotal + r.electricityTotal - r.electricityProduite
  console.log(`TOTAL STATION : eau ${eau.electricityTotal.toFixed(0)} + boues ${boues.electricityTotal.toFixed(0)} + utilités ${(r.electricityTotal - r.electricityProduite).toFixed(0)} = ${totalStation.toFixed(0)} kWh/j`)
  console.log(`soit ${(totalStation / eau.outReel.Q).toFixed(3)} kWh/m³`)
  return r
}

const filiereEau = [
  { uid: 'a', nodeId: 'dessablage-deshuilage', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'decantation-simple', choices: {}, forced: {} },
  { uid: 'c', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
]
const filiereBoues = [
  { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
  { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
  { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
]

essai('Désodorisation biologique + turbine + réactifs', filiereEau, filiereBoues, [
  { uid: 'u1', nodeId: 'desodorisation-bio', choices: { couverture: 'partielle' }, forced: {} },
  { uid: 'u2', nodeId: 'turbine-hydraulique', choices: {}, forced: {} },
  { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
])

essai('Désodorisation chimique 3 tours, couverture totale', filiereEau, filiereBoues, [
  { uid: 'u1', nodeId: 'desodorisation-chimique', choices: { couverture: 'totale', tour_acide: 'oui', tour_pH9: 'oui', tour_neutre: 'oui' }, forced: {} },
  { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
])

// eau brute septique : les sulfures chargent l'air vicié en H2S
const siteSeptique = { ...DEFAULT_SITE, HS_nominal_mgL: 5 }
essai('Eau septique : désodorisation chimique 4 tours + réactifs', filiereEau, filiereBoues, [
  { uid: 'u1', nodeId: 'desodorisation-chimique', choices: { couverture: 'totale', tour_acide: 'oui', tour_pH9: 'oui', tour_pH11: 'oui', tour_neutre: 'oui' }, forced: {} },
  { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
], siteSeptique)

// ─────────── HVAC et photovoltaïque
essai('HVAC trois bâtiments + photovoltaïque (climat tempéré)', filiereEau, filiereBoues, [
  { uid: 'h1', nodeId: 'hvac-admin', choices: { climat: 'tempere', chauffage: 'gaz', climatisation: 'groupe_froid' }, forced: {} },
  { uid: 'h2', nodeId: 'hvac-exploitation', choices: { climat: 'tempere' }, forced: {} },
  { uid: 'h3', nodeId: 'hvac-electrique', choices: { climat: 'tempere' }, forced: {} },
  { uid: 'pv', nodeId: 'photovoltaique', choices: { clarte: 'tempere', technologie: 'poly' }, forced: {} },
  { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
])

essai('HVAC avec PAC sur eaux usées + PV méditerranéen', filiereEau, filiereBoues, [
  { uid: 'h1', nodeId: 'hvac-admin', choices: { climat: 'mediterraneen', chauffage: 'pac', climatisation: 'pac' }, forced: {} },
  { uid: 'pv', nodeId: 'photovoltaique', choices: { clarte: 'mediterraneen', technologie: 'mono' }, forced: { latitude: 43.3, surface_installee: 1500 } },
], { ...DEFAULT_SITE })

// ─────────── traitement des sulfures, en amont de la désodorisation
const siteSulfures = { ...DEFAULT_SITE, HS_nominal_mgL: 8 }

essai('Sulfures non traités → désodorisation chimique', filiereEau, filiereBoues, [
  { uid: 's1', nodeId: 'traitement-sulfures', choices: { preventif: 'non', curatif: 'non' }, forced: {} },
  { uid: 'd1', nodeId: 'desodorisation-chimique', choices: { couverture: 'totale' }, forced: {} },
  { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
], siteSulfures)

essai('Sulfures traités en curatif à l\'oxygène → désodorisation allégée', filiereEau, filiereBoues, [
  { uid: 's1', nodeId: 'traitement-sulfures', choices: { preventif: 'oui', curatif: 'O2', oxygenateur: 'oxygenateur' }, forced: {} },
  { uid: 'd1', nodeId: 'desodorisation-chimique', choices: { couverture: 'totale' }, forced: {} },
  { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
], siteSulfures)
