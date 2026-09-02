import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
for (const [label, choices] of [['NDNs', {}], ['NDNc', {type:'NDNc'}], ['N media 3.6', {type:'N',diametre_media:'3.6'}], ['C 4.5', {type:'C',diametre_media:'4.5'}]]) {
  const r = runChain([
    {uid:'a',nodeId:'decantation-reactifs',choices:{},forced:{}},
    {uid:'b',nodeId:'biostyr',choices,forced:{}},
  ], REGISTRY, DEFAULT_SITE)
  const s = r.steps[1]
  console.log('### Biostyr '+label+' elec='+(s.electricity.total||0).toFixed(0)+' kWh/j')
  console.log('  out réel: Q='+s.outReel.Q.toFixed(0)+' DCO='+conc(s.outReel,'DCO').toFixed(1)+' DBO='+conc(s.outReel,'DBO').toFixed(1)+' MES='+conc(s.outReel,'MES').toFixed(1)+' NH4='+conc(s.outReel,'NH4').toFixed(1)+' NO3='+conc(s.outReel,'NO3').toFixed(1))
  for (const x of s.results) console.log('   · '+x.label+': '+(typeof x.value==='number'?x.value.toFixed(3):x.value)+' '+x.unit)
  for (const w of s.warnings) console.log('   ⚠ '+w)
  console.log('   élec:', Object.entries(s.electricity.detail).map(([k,v])=>k+'='+v.toFixed(0)).join(' '))
}
