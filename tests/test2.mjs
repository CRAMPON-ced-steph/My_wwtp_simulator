import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'

const show = (r) => {
  for (const s of r.steps) {
    console.log('\n=== '+s.label+'  elec='+(s.electricity.total||0).toFixed(0)+' kWh/j')
    console.log('  out réel: Q='+s.outReel.Q.toFixed(0)+' | DCO='+conc(s.outReel,'DCO').toFixed(1)+' DBO='+conc(s.outReel,'DBO').toFixed(1)+' MES='+conc(s.outReel,'MES').toFixed(1)+' NK='+conc(s.outReel,'NK').toFixed(1)+' NH4='+conc(s.outReel,'NH4').toFixed(1)+' NO3='+conc(s.outReel,'NO3').toFixed(1)+' Pt='+conc(s.outReel,'Pt').toFixed(2)+' mg/L')
    for (const x of s.results) console.log('   · '+x.label+': '+(typeof x.value==='number'?x.value.toFixed(3):x.value)+' '+x.unit)
    for (const w of s.warnings) console.log('   ⚠ '+w)
    if (s.electricity.detail) console.log('   élec:', Object.entries(s.electricity.detail).map(([k,v])=>k+'='+v.toFixed(0)).join(' '))
  }
  console.log('\nTOTAL', r.electricityTotal.toFixed(0),'kWh/j')
}

console.log('############ Filière 1 : primaire + BA faible charge (plug-flow, co-précipitation)')
show(runChain([
  {uid:'a',nodeId:'dessablage-deshuilage',choices:{},forced:{}},
  {uid:'b',nodeId:'decantation-simple',choices:{},forced:{}},
  {uid:'c',nodeId:'ba-faible-charge',choices:{},forced:{}},
], REGISTRY, DEFAULT_SITE))

console.log('\n\n############ Filière 2 : aération prolongée eau brute (chenal, bio-P)')
show(runChain([
  {uid:'c',nodeId:'aeration-prolongee',choices:{config_bassin:'chenal',dephosphatation:'bio'},forced:{}},
], REGISTRY, DEFAULT_SITE))

console.log('\n\n############ Filière 3 : MBR eau brute')
show(runChain([
  {uid:'c',nodeId:'mbr',choices:{},forced:{}},
], REGISTRY, DEFAULT_SITE))
