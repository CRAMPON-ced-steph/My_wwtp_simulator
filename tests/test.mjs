import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
const chain = ['dessablage-deshuilage','decantation-simple','ba-forte-charge','decantation-tertiaire','discfilter','filtration-sable','desinfection-uv','chloration'].map((nodeId,i)=>({uid:'n'+i,nodeId,choices:{},forced:{}}))
const r = runChain(chain, REGISTRY, DEFAULT_SITE)
for (const s of r.steps) {
  console.log('\n=== '+s.label+'  elec='+s.electricity.total.toFixed(0)+' kWh/j')
  console.log('  out réel: Q='+s.outReel.Q.toFixed(0)+' DCO='+conc(s.outReel,'DCO').toFixed(1)+' DBO='+conc(s.outReel,'DBO').toFixed(1)+' MES='+conc(s.outReel,'MES').toFixed(1)+' NK='+conc(s.outReel,'NK').toFixed(1)+' Pt='+conc(s.outReel,'Pt').toFixed(2))
  for (const x of s.results) console.log('  '+x.label+': '+(typeof x.value==='number'?x.value.toFixed(3):x.value)+' '+x.unit)
  for (const w of s.warnings) console.log('  ⚠ '+w)
}
console.log('\nTOTAL', r.electricityTotal.toFixed(0),'kWh/j')
