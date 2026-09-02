import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'
const r = runChain([
  {uid:'a',nodeId:'decantation-reactifs',choices:{},forced:{}},
  {uid:'b',nodeId:'ba-faible-charge',choices:{},forced:{}},
], REGISTRY, DEFAULT_SITE)
for (const s of r.steps) {
  console.log('=== '+s.label+' elec='+(s.electricity.total||0).toFixed(0))
  console.log('  out: Q='+s.outReel.Q.toFixed(0)+' DCO='+conc(s.outReel,'DCO').toFixed(1)+' DBO='+conc(s.outReel,'DBO').toFixed(1)+' MES='+conc(s.outReel,'MES').toFixed(1)+' NK='+conc(s.outReel,'NK').toFixed(1)+' NO3='+conc(s.outReel,'NO3').toFixed(1)+' Pt='+conc(s.outReel,'Pt').toFixed(2))
  for (const x of s.results) console.log('   · '+x.label+': '+x.value.toFixed(3)+' '+x.unit)
  for (const w of s.warnings) console.log('   ⚠ '+w)
}
