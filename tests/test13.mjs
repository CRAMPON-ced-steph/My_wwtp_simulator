// Analyse pincement : solveur seul, puis intégration à une filière complète.
import { analyserPincement } from '../src/core/pinch.js'
import { simuler } from '../src/core/simulation.js'
import { REGISTRY } from '../src/nodes/index.js'
import { SLUDGE_REGISTRY } from '../src/nodes-boues/index.js'
import { UTILITY_REGISTRY } from '../src/nodes-utilites/index.js'
import { TRANSVERSE_REGISTRY } from '../src/nodes-transverse/index.js'
import { DEFAULT_SITE } from '../src/core/stream.js'

let echecs = 0
const verifier = (ok, texte) => { console.log(`   ${ok ? '✓' : '✗'} ${texte}`); if (!ok) echecs += 1 }

console.log('══════ CAS ANALYTIQUES — vérifiables à la main')

// (1) Contre-courant parfait : les deux flux ont le même CP et un écart de 50 K
// aux deux extrémités, très au-dessus des 10 K exigés. Récupération totale.
const c1 = analyserPincement([
  { nom: 'H', type: 'chaud', T_in: 200, T_out: 100, CP: 1 },
  { nom: 'C', type: 'froid', T_in: 50, T_out: 150, CP: 1 },
], { dT_contribution_defaut: 5 })
console.log('(1) Contre-courant parfait, 100 kW de part et d\'autre')
verifier(Math.abs(c1.bilan.Qh_min) < 1e-6, `Qh = ${c1.bilan.Qh_min.toFixed(1)} kW, attendu 0`)
verifier(Math.abs(c1.bilan.Qc_min) < 1e-6, `Qc = ${c1.bilan.Qc_min.toFixed(1)} kW, attendu 0`)
verifier(Math.abs(c1.bilan.recuperation - 100) < 1e-6, `récupération = ${c1.bilan.recuperation.toFixed(1)} kW, attendu 100`)

// (2) Même flux chaud, mais le froid monte à 195 °C : la partie haute n'est plus
// couvrable et appelle 45 kW d'utilité chaude. Calcul à la main : le froid
// requiert 145 kW, le chaud n'en fournit que 100 sans violer le ΔT.
const c2 = analyserPincement([
  { nom: 'H', type: 'chaud', T_in: 200, T_out: 100, CP: 1 },
  { nom: 'C', type: 'froid', T_in: 50, T_out: 195, CP: 1 },
], { dT_contribution_defaut: 5 })
console.log('(2) Flux froid plus exigeant, 145 kW à fournir')
verifier(Math.abs(c2.bilan.Qh_min - 45) < 1e-6, `Qh = ${c2.bilan.Qh_min.toFixed(1)} kW, attendu 45`)
verifier(Math.abs(c2.bilan.Qc_min) < 1e-6, `Qc = ${c2.bilan.Qc_min.toFixed(1)} kW, attendu 0`)
verifier(Math.abs(c2.bilan.recuperation - 100) < 1e-6, `récupération = ${c2.bilan.recuperation.toFixed(1)} kW, attendu 100`)

// (3) Contributions ΔT dissymétriques. Chaud 200→100 et froid 100→180, CP = 2
// de part et d'autre. Avec 5 K chacun, l'écart de 20 K au sommet suffit et rien
// n'est requis. Avec 20 K pour le gaz et 2 K pour la boue, l'écart exigé monte à
// 22 K : il en manque 2, soit 2 × CP = 4 kW d'utilité chaude.
const symetrique = analyserPincement([
  { nom: 'Gaz', type: 'chaud', T_in: 200, T_out: 100, CP: 2 },
  { nom: 'Boue', type: 'froid', T_in: 100, T_out: 180, CP: 2 },
], { dT_contribution_defaut: 5 })
const c3 = analyserPincement([
  { nom: 'Gaz', type: 'chaud', T_in: 200, T_out: 100, CP: 2, dT_contribution: 20 },
  { nom: 'Boue', type: 'froid', T_in: 100, T_out: 180, CP: 2, dT_contribution: 2 },
], { dT_contribution_defaut: 5 })
console.log('(3) Contributions dissymétriques, 20 K et 2 K contre 5 K et 5 K')
verifier(Math.abs(symetrique.bilan.Qh_min) < 1e-6, `contributions égales : Qh = ${symetrique.bilan.Qh_min.toFixed(1)} kW, attendu 0`)
verifier(Math.abs(c3.bilan.Qh_min - 4) < 1e-6, `contributions dissymétriques : Qh = ${c3.bilan.Qh_min.toFixed(1)} kW, attendu 4 — l'écart exigé est bien la somme des deux contributions`)

console.log('\n══════ INVARIANTS — jeux de flux aléatoires')
// Trois propriétés doivent tenir quel que soit le jeu de flux :
//   le bilan énergétique global, la positivité de la cascade, et l'absence de
//   croisement des composites — cette dernière étant la preuve que le ΔT est
//   partout respecté.
let alea = 12345
const rnd = () => { alea = (alea * 1103515245 + 12345) % 2147483648; return alea / 2147483648 }
let ok_bilan = 0, ok_cascade = 0, ok_composites = 0, total = 0
for (let essai = 0; essai < 200; essai++) {
  const n = 2 + Math.floor(rnd() * 5)
  const flux = []
  for (let i = 0; i < n; i++) {
    const chaud = rnd() < 0.5
    const a = 20 + rnd() * 250
    const b = 20 + rnd() * 250
    if (Math.abs(a - b) < 5) continue
    flux.push({
      nom: `F${i}`,
      type: chaud ? 'chaud' : 'froid',
      T_in: chaud ? Math.max(a, b) : Math.min(a, b),
      T_out: chaud ? Math.min(a, b) : Math.max(a, b),
      CP: 0.5 + rnd() * 5,
    })
  }
  const r = analyserPincement(flux, { dT_contribution_defaut: 5 })
  if (r.erreur) continue
  total += 1
  // bilan : ce qui entre en utilité chaude moins ce qui sort en froide est
  // exactement le déséquilibre entre besoins et disponibilités
  const bilan = r.bilan.Qh_min - r.bilan.Qc_min - (r.bilan.charge_froide - r.bilan.charge_chaude)
  if (Math.abs(bilan) < 1e-6) ok_bilan += 1
  // la cascade décalée ne doit jamais devenir négative
  if (r.table.cascade.every((q) => q > -1e-9)) ok_cascade += 1
  // les composites ne doivent pas se croiser : à enthalpie égale, la chaude
  // reste au-dessus de la froide
  let croisement = false
  // Une composite peut comporter des segments verticaux, là où aucun flux
  // n'existe dans un intervalle de température : à enthalpie constante, la
  // courbe y prend plusieurs températures. On les ignore et on interpole sur
  // le segment de pente non nulle qui contient le point.
  const interp = (courbe, h) => {
    for (let i = 0; i < courbe.length - 1; i++) {
      const d = courbe[i + 1].H - courbe[i].H
      if (d < 1e-9) continue
      if (h >= courbe[i].H - 1e-9 && h <= courbe[i + 1].H + 1e-9) {
        return courbe[i].T + ((h - courbe[i].H) / d) * (courbe[i + 1].T - courbe[i].T)
      }
    }
    return null
  }
  // le contrôle porte sur les composites décalées, seules cohérentes quand les
  // contributions diffèrent d'un flux à l'autre
  const { chaude, froide } = r.composites
  const H0 = Math.max(chaude[0].H, froide[0].H)
  const H1 = Math.min(chaude[chaude.length - 1].H, froide[froide.length - 1].H)
  // on n'évalue que si le recouvrement est significatif, et en écartant les
  // bornes : aux extrémités exactes les deux courbes se touchent par
  // construction, et un recouvrement quasi nul ne dit rien du ΔT
  const largeur = H1 - H0
  const significatif = largeur > 1e-6 * Math.max(1, Math.abs(H1))
  for (let k = 1; k < 40 && significatif; k++) {
    const h = H0 + (largeur * k) / 40
    const tc = interp(chaude, h)
    const tf = interp(froide, h)
    if (tc != null && tf != null && tc - tf < -1e-6) croisement = true
  }
  if (!croisement) ok_composites += 1
}
console.log(`   ${total} jeux de flux exploitables`)
verifier(ok_bilan === total, `bilan énergétique exact : ${ok_bilan}/${total}`)
verifier(ok_cascade === total, `cascade positive partout : ${ok_cascade}/${total}`)
verifier(ok_composites === total, `composites jamais croisées : ${ok_composites}/${total}`)

console.log('\n══════ CHANGEMENT DE PHASE')
const avecPhase = [
  { nom: 'Condensation vapeur', type: 'chaud', T_in: 120, T_out: 120, charge: 500 },
  { nom: 'Boues à réchauffer', type: 'froid', T_in: 15, T_out: 95, CP: 4 },
]
const b = analyserPincement(avecPhase, { dT_contribution_defaut: 5 })
console.log(`Qh = ${b.bilan.Qh_min.toFixed(1)} kW, Qc = ${b.bilan.Qc_min.toFixed(1)} kW, récupération ${b.bilan.recuperation.toFixed(1)} kW`)
console.log(`Le palier isotherme est bien traité : ${b.flux[0].segments[0].isotherme ? 'oui' : 'non'}`)
verifier(Math.abs(b.bilan.Qh_min - b.bilan.Qc_min - (b.bilan.charge_froide - b.bilan.charge_chaude)) < 1e-6, 'bilan exact malgré le palier')

console.log('\n══════ FILIÈRE COMPLÈTE')
const REGISTRES = { eau: REGISTRY, boues: SLUDGE_REGISTRY, utilites: UTILITY_REGISTRY, transverse: TRANSVERSE_REGISTRY }
const inst = (nodeId, choices = {}, forced = {}) => ({ uid: nodeId, nodeId, choices, forced })
const r = simuler({
  eau: [inst('decantation-simple'), inst('ba-faible-charge')],
  boues: [
    inst('epaississement', { type: 'centrifuge' }),
    inst('exelys', { entree_1: 'epaississeur_1' }),
    inst('deshydratation', { type: 'centrifuge', entree_1: 'digerees', digestion_amont: 'avancee' }),
    inst('sechage-thermique', { technologie: 'indirect', entree_1: 'deshydratees_1' }),
    inst('incineration', { entree_1: 'sechees_1' }),
  ],
  utilites: [inst('hvac-admin', { chauffage: 'recup' }), inst('energie-thermique')],
  transverse: [inst('bilan-electrique'), inst('empreinte-co2')],
}, REGISTRES, DEFAULT_SITE)

const pas = r.utilites.steps.find((s) => s.pincement)
if (!pas) { console.log('Pas d\'analyse : ' + r.utilites.steps.map((s) => s.warnings.join(' ')).join(' | ')) }
else {
  const an = pas.pincement
  console.log('Flux thermiques identifiés :')
  for (const f of an.flux) {
    console.log(`   ${f.chaud ? 'chaud' : 'froid'}  ${f.nom.padEnd(42)} ${f.T_in.toFixed(0).padStart(4)} → ${f.T_out.toFixed(0).padStart(4)} °C   ${f.charge.toFixed(0).padStart(6)} kW`)
  }
  console.log()
  for (const x of pas.results) console.log('   · ' + x.label + ' : ' + x.value.toFixed(2) + ' ' + x.unit)
  for (const w of new Set(pas.warnings)) console.log('   ⚠ ' + w)
  console.log('\nAlertes de conception :')
  for (const al of an.alertes) console.log(`   [${al.niveau}] ${al.texte}`)
}

console.log(echecs === 0 ? '\n✓ toutes les vérifications passent' : `\n✗ ${echecs} vérification(s) en échec`)
if (echecs > 0) process.exitCode = 1
