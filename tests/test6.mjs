import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
const r = runChain([
  {uid:'a',nodeId:'decantation-reactifs',choices:{},forced:{}},
  {uid:'b',nodeId:'biostyr',choices:{type:'N',diametre_media:'3.6'},forced:{}},
  {uid:'c',nodeId:'biostyr-pdn',choices:{},forced:{}},
], REGISTRY, DEFAULT_SITE)
for (const s of r.steps.slice(1)) {
  console.log('### '+s.label+' elec='+(s.electricity.total||0).toFixed(0))
  console.log('  out: Q='+s.outReel.Q.toFixed(0)+' DCO='+conc(s.outReel,'DCO').toFixed(1)+' NK='+conc(s.outReel,'NK').toFixed(1)+' NH4='+conc(s.outReel,'NH4').toFixed(1)+' NO3='+conc(s.outReel,'NO3').toFixed(1)+' NGL='+(conc(s.outReel,'NK')+conc(s.outReel,'NO3')).toFixed(1))
  for (const x of s.results) console.log('   · '+x.label+': '+(typeof x.value==='number'?x.value.toFixed(3):x.value)+' '+x.unit)
  for (const w of s.warnings) console.log('   ⚠ '+w)
}
