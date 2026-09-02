import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
const r = runChain([
  {uid:'a',nodeId:'decantation-reactifs',choices:{microsable:'oui'},forced:{}},
  {uid:'b',nodeId:'biostyr',choices:{type:'C',diametre_media:'4.5'},forced:{}},
  {uid:'c',nodeId:'biostyr-nitrifiant-iii',choices:{},forced:{}},
  {uid:'d',nodeId:'biostyr-pdn-iii',choices:{},forced:{}},
], REGISTRY, DEFAULT_SITE)
for (const s of r.steps.slice(1)) {
  console.log('### '+s.label+' elec='+(s.electricity.total||0).toFixed(0))
  console.log('  out: Q='+s.outReel.Q.toFixed(0)+' DCO='+conc(s.outReel,'DCO').toFixed(1)+' DBO='+conc(s.outReel,'DBO').toFixed(1)+' MES='+conc(s.outReel,'MES').toFixed(1)+' NH4='+conc(s.outReel,'NH4').toFixed(1)+' NO3='+conc(s.outReel,'NO3').toFixed(1)+' NGL='+(conc(s.outReel,'NK')+conc(s.outReel,'NO3')).toFixed(1))
  for (const w of s.warnings) console.log('   ⚠ '+w)
}
console.log('TOTAL', r.electricityTotal.toFixed(0),'kWh/j')
