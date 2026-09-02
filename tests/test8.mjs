import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
const r = runChain([
  {uid:'a',nodeId:'decantation-reactifs',choices:{microsable:'oui'},forced:{}},
  {uid:'b',nodeId:'biostyr',choices:{type:'C',diametre_media:'4.5'},forced:{}},
  {uid:'c',nodeId:'biostyr-nitrifiant-iii',choices:{},forced:{}},
  {uid:'d',nodeId:'biostyr-pdn-iii',choices:{},forced:{}},
  {uid:'e',nodeId:'decantation-eaux-sales',choices:{},forced:{}},
], REGISTRY, DEFAULT_SITE)
console.log('RETOUR en tête (réel): Q='+r.retourReel.Q.toFixed(0)+' m³/j, MES='+r.retourReel.MES.toFixed(0)+' kg/j')
for (const s of r.steps) {
  console.log('### '+s.label+' | in Q='+s.inReel.Q.toFixed(0)+' → out Q='+s.outReel.Q.toFixed(0)+' | elec='+(s.electricity.total||0).toFixed(0))
  if (s.nodeId==='decantation-eaux-sales') for (const x of s.results) console.log('   · '+x.label+': '+x.value.toFixed(1)+' '+x.unit)
  for (const w of s.warnings) console.log('   ⚠ '+w)
}
console.log('Sortie finale: Q='+r.outReel.Q.toFixed(0)+' DCO='+conc(r.outReel,'DCO').toFixed(1)+' NGL='+(conc(r.outReel,'NK')+conc(r.outReel,'NO3')).toFixed(1)+' | TOTAL élec '+r.electricityTotal.toFixed(0)+' kWh/j')
