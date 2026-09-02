// Bloc transverse : bilan électrique et empreinte CO2 sur une filière complète.
import { simuler } from '../src/core/simulation.js'
import { REGISTRY } from '../src/nodes/index.js'
import { SLUDGE_REGISTRY } from '../src/nodes-boues/index.js'
import { UTILITY_REGISTRY } from '../src/nodes-utilites/index.js'
import { TRANSVERSE_REGISTRY } from '../src/nodes-transverse/index.js'
import { DEFAULT_SITE } from '../src/core/stream.js'

const REGISTRES = { eau: REGISTRY, boues: SLUDGE_REGISTRY, utilites: UTILITY_REGISTRY, transverse: TRANSVERSE_REGISTRY }

const filiere = {
  eau: [
    { uid: 'a', nodeId: 'dessablage-deshuilage', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'c', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  boues: [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 'am', nodeId: 'anita-mox', choices: {}, forced: {} },
  ],
  utilites: [
    { uid: 'u1', nodeId: 'desodorisation-bio', choices: {}, forced: {} },
    { uid: 'u2', nodeId: 'turbine-hydraulique', choices: {}, forced: {} },
    { uid: 'u3', nodeId: 'gestion-reactifs', choices: {}, forced: {} },
  ],
  transverse: [
    { uid: 't1', nodeId: 'bilan-electrique', choices: {}, forced: {} },
    { uid: 't2', nodeId: 'empreinte-co2', choices: { pays: 'france' }, forced: {} },
    { uid: 't3', nodeId: 'gestion-opex', choices: {}, forced: {} },
    { uid: 't4', nodeId: 'retour-investissement', choices: {}, forced: { capex: 12000, capex_reference: 9000, opex_reference: 620 } },
  ],
}

const r = simuler(filiere, REGISTRES, DEFAULT_SITE)

console.log('══════ BILAN ÉLECTRIQUE PAR POSTE')
const largeur = Math.max(...r.bilan.postes.map((p) => p.label.length))
for (const p of r.bilan.postes) {
  const part = r.bilan.consommee > 0 ? p.total / r.bilan.consommee : 0
  const barre = (p.total >= 0 ? '█' : '▒').repeat(Math.max(0, Math.round(Math.abs(part) * 40)))
  console.log(`${p.label.padEnd(largeur)} ${p.total.toFixed(0).padStart(6)} kWh/j  ${(part * 100).toFixed(1).padStart(6)} %  ${barre}`)
}
console.log()
for (const x of r.transverse.steps[0].results) {
  if (x.key.startsWith('poste_')) continue
  console.log('   · ' + x.label + ' : ' + x.value.toFixed(3) + ' ' + x.unit)
}
for (const w of new Set(r.transverse.steps[0].warnings)) console.log('   ⚠ ' + w)

console.log('\n══════ EMPREINTE CO2')
for (const x of r.transverse.steps[1].results) {
  console.log('   · ' + x.label + ' : ' + x.value.toFixed(2) + ' ' + x.unit)
}
for (const w of new Set(r.transverse.steps[1].warnings)) console.log('   ⚠ ' + w)
console.log('\nDétail des réactifs :')
for (const l of r.co2.lignes) {
  console.log(`   ${l.label.padEnd(22)} ${l.tonnes_an.toFixed(1).padStart(7)} t/an  →  ${l.total.toFixed(2).padStart(7)} t CO2/an  (${l.source})`)
}

console.log('\n══════ CONTRÔLES')
const somme = r.bilan.postes.reduce((s, p) => s + p.total, 0)
console.log(`Somme des postes ${somme.toFixed(1)} = total moteurs ${r.electricite.total.toFixed(1)} → écart ${(somme - r.electricite.total).toFixed(3)} kWh/j`)
for (const x of [...r.transverse.steps[0].results, ...r.transverse.steps[1].results]) {
  if (!Number.isFinite(x.value)) console.log('   ✗ valeur non finie : ' + x.label)
}

console.log('\n══════ COÛTS D\'EXPLOITATION')
for (const x of r.transverse.steps[2].results) console.log('   · ' + x.label + ' : ' + x.value.toFixed(2) + ' ' + x.unit)
for (const w of new Set(r.transverse.steps[2].warnings)) console.log('   ⚠ ' + w)
console.log('\nDétail des réactifs :')
for (const l of r.opex.lignesReactifs) {
  console.log(`   ${l.label.padEnd(22)} ${l.tonnes_an.toFixed(1).padStart(7)} t/an × ${String(l.prix).padStart(5)} €/t  =  ${(l.cout_an / 1000).toFixed(1).padStart(6)} k€/an`)
}

console.log('\n══════ RETOUR SUR INVESTISSEMENT')
for (const x of r.transverse.steps[3].results) console.log('   · ' + x.label + ' : ' + x.value.toFixed(2) + ' ' + x.unit)
for (const w of new Set(r.transverse.steps[3].warnings)) console.log('   ⚠ ' + w)
if (r.roi) {
  console.log('\nValeur actuelle nette, année par année (k€) :')
  const max = Math.max(...r.roi.serie.map(Math.abs), 1)
  r.roi.serie.forEach((v, i) => {
    if (i % 2 && i !== r.roi.serie.length - 1) return
    const n = Math.round((Math.abs(v) / max) * 30)
    const barre = v >= 0 ? ' '.repeat(30) + '│' + '█'.repeat(n) : ' '.repeat(30 - n) + '▒'.repeat(n) + '│'
    console.log(`   an ${String(i).padStart(2)}  ${v.toFixed(0).padStart(7)}  ${barre}`)
  })
}
