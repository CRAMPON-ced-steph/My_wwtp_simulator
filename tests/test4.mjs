import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
const combos = [
  [['dessablage-deshuilage'],['decantation-reactifs'],['ba-moyenne-charge'],['desinfection-uv']],
]
const chain = ['dessablage-deshuilage','decantation-reactifs','ba-moyenne-charge','decantation-tertiaire','desinfection-uv'].map((n,i)=>({uid:'x'+i,nodeId:n,choices:{},forced:{}}))
const r = runChain(chain, REGISTRY, DEFAULT_SITE)
for (const s of r.steps) {
  console.log('=== '+s.label+' elec='+(s.electricity.total||0).toFixed(0)+'  out: DCO='+conc(s.outReel,'DCO').toFixed(1)+' DBO='+conc(s.outReel,'DBO').toFixed(1)+' MES='+conc(s.outReel,'MES').toFixed(1)+' Pt='+conc(s.outReel,'Pt').toFixed(2))
  for (const w of s.warnings) console.log('   ⚠ '+w)
}
console.log('TOTAL', r.electricityTotal.toFixed(0))
// sanity : chaque nœud porté doit préserver la masse : sortie + boues ≈ entrée pour MES ? (non strict), on vérifie juste pas de NaN/négatif
let bad=0
for (const s of r.steps) for (const k of ['Q','DCO','DBO','MES','NK','NH4','NO3','Pt']) {
  if (!Number.isFinite(s.outReel[k]) || s.outReel[k] < -1e-9) { console.log('NEGATIF/NaN', s.label, k, s.outReel[k]); bad++ }
  if (!Number.isFinite(s.outNominal[k]) || s.outNominal[k] < -1e-9) { console.log('NEG/NaN nominal', s.label, k, s.outNominal[k]); bad++ }
}
console.log(bad? 'ECHEC':'sanity ok')
