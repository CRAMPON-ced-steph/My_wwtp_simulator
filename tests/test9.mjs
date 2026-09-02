// Test du port E6_MBBR : trois configurations types.
import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { DEFAULT_SITE, conc } from '../src/core/stream.js'

function essai(titre, chain, site = DEFAULT_SITE) {
  console.log('\n══════ ' + titre)
  const r = runChain(chain, REGISTRY, site)
  for (const s of r.steps) {
    console.log('### ' + s.label + ' | in Q=' + s.inReel.Q.toFixed(0) + ' → out Q=' + s.outReel.Q.toFixed(0) + ' | élec=' + (s.electricity.total || 0).toFixed(0) + ' kWh/j')
    if (s.nodeId === 'mbbr' || s.nodeId === 'hybas') {
      for (const x of s.results) console.log('   · ' + x.label + ' : ' + (Number.isFinite(x.value) ? x.value.toFixed(2) : String(x.value)) + ' ' + x.unit)
      console.log('   · détail élec : ' + JSON.stringify(Object.fromEntries(Object.entries(s.electricity.detail).map(([k, v]) => [k, Math.round(v)]))))
    }
    for (const w of new Set(s.warnings)) console.log('   ⚠ ' + w)
  }
  const o = r.outReel
  console.log('SORTIE réelle : Q=' + o.Q.toFixed(0) + ' DCO=' + conc(o, 'DCO').toFixed(1) + ' DBO=' + conc(o, 'DBO').toFixed(1) +
    ' MES=' + conc(o, 'MES').toFixed(1) + ' NK=' + conc(o, 'NK').toFixed(1) + ' NH4=' + conc(o, 'NH4').toFixed(1) +
    ' NO3=' + conc(o, 'NO3').toFixed(1) + ' Pt=' + conc(o, 'Pt').toFixed(2))
  const n = r.outNominal
  console.log('SORTIE nominale : Q=' + n.Q.toFixed(0) + ' DBO=' + conc(n, 'DBO').toFixed(1) + ' NH4=' + conc(n, 'NH4').toFixed(1) + ' NO3=' + conc(n, 'NO3').toFixed(1))
  console.log('TOTAL élec ' + r.electricityTotal.toFixed(0) + ' kWh/j soit ' + (r.electricityTotal / (o.Q || 1)).toFixed(3) + ' kWh/m³')
  // contrôle de cohérence
  for (const [k, v] of Object.entries(o)) {
    if (!Number.isFinite(v)) console.log('   ✗ NaN/Inf sur ' + k)
    else if (v < 0) console.log('   ✗ valeur négative sur ' + k + ' = ' + v.toFixed(2))
  }
  return r
}

// 1) MBBR pré-dénit + C + N + dé-ox sur eau brute
essai('MBBR PréDN + C + N + DéOx (eau brute)', [
  { uid: 'a', nodeId: 'dessablage-deshuilage', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'mbbr', choices: {}, forced: {} },
])

// 2) MBBR carbone seul derrière un décanteur primaire
essai('MBBR carbone seul (après décantation primaire)', [
  { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'mbbr', choices: { etape_predenit: 'non', etape_C: 'oui', etape_N: 'non', etape_deox: 'non' }, forced: {} },
])

// 3) MBBR complet avec post-dénitrification au méthanol et ré-oxygénation
essai('MBBR complet PréDN + C + N + DéOx + PostDN + RéOx', [
  { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'mbbr', choices: { etape_postdenit: 'oui', etape_reox: 'oui' }, forced: {} },
])

// 4) MBBR + tertiaire, avec forçage de volumes
essai('MBBR avec volumes forcés + filtration tertiaire', [
  { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'mbbr', choices: { media_N: 'chip_M', agitation_cuve_aeree: 'mecanique' }, forced: { cuve5_volume: 1500, cuve6_volume: 1500 } },
  { uid: 'c', nodeId: 'filtration-sable', choices: {}, forced: {} },
  { uid: 'd', nodeId: 'decantation-eaux-sales', choices: {}, forced: {} },
])

// ─────────── HybAS (E5)
essai('HybAS C/N–H, pré-dénit + dé-ox (eau brute)', [
  { uid: 'a', nodeId: 'dessablage-deshuilage', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'hybas', choices: {}, forced: {} },
])

essai('HybAS C/N–H–H, pré-anoxie + anaérobie bio-P (après primaire)', [
  { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'hybas', choices: { configuration: 'CN_H_H', pre_anoxie: 'oui', anaerobie: 'oui' }, forced: {} },
])

essai('HybAS H–H + post-dénit avec média + ré-ox', [
  { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'hybas', choices: { configuration: 'H_H', post_denit: 'avec_media', reox: 'oui' }, forced: {} },
])

essai('HybAS C/N–H–C/N + post-dénit sans média', [
  { uid: 'a', nodeId: 'decantation-reactifs', choices: {}, forced: {} },
  { uid: 'b', nodeId: 'hybas', choices: { configuration: 'CN_H_CN', post_denit: 'sans_media' }, forced: {} },
])
