// File boues : épaississement alimenté par une filière eau réelle.
import { runChain } from '../src/core/engine.js'
import { REGISTRY } from '../src/nodes/index.js'
import { runSludgeChain, apportsDepuisFileEau } from '../src/core/sludgeEngine.js'
import { SLUDGE_REGISTRY } from '../src/nodes-boues/index.js'
import { DEFAULT_SITE } from '../src/core/stream.js'

function essai(titre, chainEau, chainBoues) {
  console.log('\n══════ ' + titre)
  const eau = runChain(chainEau, REGISTRY, DEFAULT_SITE)
  const apports = apportsDepuisFileEau(eau, REGISTRY)
  console.log('Boues produites par la file eau :')
  for (const a of apports) {
    console.log(`   · type ${a.type} (${a.origine}) : ${a.MES.toFixed(0)} kg MES/j, ${a.Q.toFixed(0)} m³/j, MV/MES ${a.MV_MES.toFixed(2)}`)
  }
  const r = runSludgeChain(chainBoues, SLUDGE_REGISTRY, DEFAULT_SITE, apports)
  for (const s of r.steps) {
    console.log(`### ${s.label} | élec=${(s.electricity.total || 0).toFixed(0)} kWh/j`)
    for (const x of s.results) console.log('   · ' + x.label + ' : ' + (Number.isFinite(x.value) ? x.value.toFixed(2) : x.value) + ' ' + x.unit)
    console.log('   · détail élec : ' + JSON.stringify(Object.fromEntries(Object.entries(s.electricity.detail || {}).map(([k, v]) => [k, Math.round(v)]))))
    for (const w of new Set(s.warnings)) console.log('   ⚠ ' + w)
  }
  const e = r.evacuation
  console.log(`ÉVACUATION : ${e.MES.toFixed(0)} kg MES/j, ${e.Q.toFixed(1)} m³/j, siccité ${e.siccite.toFixed(1)} g/L`)
  const t = r.retourResume
  console.log(`RETOUR EN TÊTE : Q=${t.Q.toFixed(0)} m³/j, MES=${t.MES.toFixed(0)}, DCO=${t.DCO.toFixed(0)}, NK=${t.NK.toFixed(1)} kg/j`)
  console.log(`ÉLEC file boues : ${r.electricityTotal.toFixed(0)} kWh/j`)
  // contrôle de bilan matière
  const entree = apports.reduce((s, a) => s + a.MES, 0)
  // les MV détruites en digestion quittent la phase solide sous forme de biogaz
  let gazeux = 0
  for (const st of r.steps) {
    if (!st.results.some((x) => x.key === 'biogaz')) continue
    // certains procédés perdent aussi des MS latéralement (centrat d'Exelys) :
    // ils exposent alors directement les MV parties en biogaz
    const mvRed = st.results.find((x) => x.key === 'MV_reduites_total')?.value
    if (mvRed != null) { gazeux += mvRed; continue }
    const inMES = st.results.find((x) => x.key === 'in_MES')?.value
    const outMES = st.results.find((x) => x.key === 'out_MES')?.value
    if (inMES != null && outMES != null) gazeux += inMES - outMES
  }
  // le chaulage ajoute de la matière sèche : on la retire de l'entrée théorique
  let apporte = 0
  for (const st of r.steps) {
    const g = st.results.find((x) => x.key === 'gain')?.value
    if (g != null) apporte += g
  }
  // l'incinération détruit les MV et sort le minéral en cendres et REFIB
  let incinere = 0
  for (const st of r.steps) {
    const m = st.results.find((x) => x.key === 'in_MES')?.value
    if (st.results.some((x) => x.key === 'cendres') && m != null) incinere += m
  }
  // Athos oxyde les MV et laisse échapper des MS au décanteur, que le classeur
  // ne compte nulle part
  let oxyde = 0
  for (const st of r.steps) {
    const mv = st.results.find((x) => x.key === 'MV_oxydees')?.value
    const ms = st.results.find((x) => x.key === 'MS_non_captees')?.value
    if (mv != null) oxyde += mv + (ms ?? 0)
  }
  // le traitement des retours extrait des boues et laisse un écart de bilan
  // assumé par le classeur
  let retire = 0
  for (const st of r.steps) {
    const b = st.results.find((x) => x.key === 'boues')?.value
    const nb = st.results.find((x) => x.key === 'MES_non_bouclees')?.value
    if (b != null && nb != null) retire += b + nb
  }
  const sortie = e.MES + t.MES + gazeux - apporte + incinere + oxyde + retire
  const detail = (gazeux > 0 ? ` + biogaz ${gazeux.toFixed(0)}` : '')
    + (incinere > 0 ? ` + incinéré ${incinere.toFixed(0)}` : '')
    + (oxyde > 0 ? ` + oxydé/non capté ${oxyde.toFixed(0)}` : '')
    + (retire > 0 ? ` + retours traités ${retire.toFixed(0)}` : '')
    + (apporte > 0 ? ` − chaux ${apporte.toFixed(0)}` : '')
  console.log(`BILAN MES : entrée ${entree.toFixed(0)} = évacué ${e.MES.toFixed(0)} + retour ${t.MES.toFixed(0)}${detail} → écart ${(entree - sortie).toFixed(2)} kg/j`)
  for (const [k, v] of Object.entries(t)) if (!Number.isFinite(v) || v < 0) console.log(`   ✗ retour ${k} = ${v}`)
  return r
}

essai('Primaire + BA faible charge → épaississement gravitaire',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [{ uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} }])

essai('Primaire + MBBR → épaississement centrifuge',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'mbbr', choices: {}, forced: {} },
  ],
  [{ uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} }])

essai('Deux épaississeurs séparés : primaires en gravitaire, secondaires en flottation',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire', entree_1: 'I' }, forced: {} },
    { uid: 'e2', nodeId: 'epaississement', choices: { type: 'flottation', entree_1: 'II' }, forced: {} },
  ])

essai('Primaire + BA faible charge → épaississement gravitaire + digestion mésophile',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
  ])

essai('Épaississeurs séparés + digestion thermophile des deux flux',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire', entree_1: 'I' }, forced: {} },
    { uid: 'e2', nodeId: 'epaississement', choices: { type: 'centrifuge', entree_1: 'II' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: { type: 'thermophile', entree_1: 'epaississeur_1', entree_2: 'epaississeur_2' }, forced: {} },
  ])

essai('Filière complète : épaississement + digestion + déshydratation centrifuge',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
  ])

essai('Sans digestion : épaississement + filtre à plateaux chaux/FeCl3',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'filtre_plateaux', entree_1: 'epaississeur_1', conditionnement: 'chaux_FeCl3', digestion_amont: 'aucune' }, forced: {} },
  ])

essai('Filière complète avec chaulage hygiénisant',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 'c1', nodeId: 'chaulage', choices: { objectif: 'hygiene', type_chaux: 'vive', entree_1: 'deshydratees_1' }, forced: {} },
  ])

essai('Chaulage pour relèvement de siccité après filtre à bandes',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'filtre_bandes', entree_1: 'epaississeur_1', digestion_amont: 'aucune' }, forced: {} },
    { uid: 'c1', nodeId: 'chaulage', choices: { objectif: 'les_deux', type_chaux: 'vive', entree_1: 'deshydratees_1' }, forced: {} },
  ])

essai('Séchage thermique indirect total après digestion + centrifugeuse',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 's1', nodeId: 'sechage-thermique', choices: { technologie: 'indirect', performance: 'total', entree_1: 'deshydratees_1', condensation: 'indirecte' }, forced: {} },
  ])

essai('Séchage thermique CMV partiel, condensation directe',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'filtre_bandes', entree_1: 'epaississeur_1', digestion_amont: 'aucune' }, forced: {} },
    { uid: 's1', nodeId: 'sechage-thermique', choices: { technologie: 'cmv', performance: 'partiel', entree_1: 'deshydratees_1', condensation: 'directe' }, forced: {} },
  ])

essai('Séchage Bioco après déshydratation',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'epaississeur_1', digestion_amont: 'aucune' }, forced: {} },
    { uid: 's1', nodeId: 'sechage-bioco', choices: { entree_1: 'deshydratees_1', condensation: 'indirecte_directe' }, forced: {} },
  ])

essai('Séchage Inos directement sur boues digérées (déshydratation intégrée)',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 's1', nodeId: 'sechage-inos', choices: { entree_1: 'digerees', condensation: 'indirecte' }, forced: {} },
  ])

essai('Incinération de boues séchées (autocombustibles) — voie sèche',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 's1', nodeId: 'sechage-thermique', choices: { technologie: 'indirect', performance: 'total', entree_1: 'deshydratees_1' }, forced: {} },
    { uid: 'i1', nodeId: 'incineration', choices: { entree_1: 'sechees_1', combustible: 'biogaz', traitement_fumees: 'seche', digestion_amont: 'oui' }, forced: {} },
  ])

essai('Incinération directe de boues déshydratées — voie humide + SNCR urée',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'epaississeur_1', digestion_amont: 'aucune' }, forced: {} },
    { uid: 'i1', nodeId: 'incineration', choices: { entree_1: 'deshydratees_1', combustible: 'gaz_naturel', traitement_fumees: 'humide', traitement_NOx: 'uree', digestion_amont: 'non' }, forced: {} },
  ])

essai('Athos sur boues épaissies non digérées',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'at', nodeId: 'athos', choices: { entree_1: 'epaississeur_1', digestion_amont: 'aucune' }, forced: {} },
  ])

essai('Athos sur boues digérées + déshydratation du technosable',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'at', nodeId: 'athos', choices: { entree_1: 'digerees', digestion_amont: 'simple' }, forced: {} },
  ])

essai('Biothelys sur boues épaissies + déshydratation',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'bt', nodeId: 'biothelys', choices: { entree_lysee_1: 'epaississeur_1' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees', digestion_amont: 'avancee' }, forced: {} },
  ])

essai('Biothelys avec by-pass partiel (primaires non lysées)',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge', entree_1: 'II' }, forced: {} },
    { uid: 'e2', nodeId: 'epaississement', choices: { type: 'gravitaire', entree_1: 'I' }, forced: {} },
    { uid: 'bt', nodeId: 'biothelys', choices: { entree_lysee_1: 'epaississeur_1', entree_non_lysee: 'epaississeur_2' }, forced: {} },
  ])

essai('Exelys DLD sur boues épaissies + déshydratation',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'ex', nodeId: 'exelys', choices: { entree_1: 'epaississeur_1' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees', digestion_amont: 'avancee' }, forced: {} },
  ])

essai('Exelys DLD avec by-pass de boues fraîches vers la lyse',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'aeration-prolongee', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire', entree_1: 'I' }, forced: {} },
    { uid: 'e2', nodeId: 'epaississement', choices: { type: 'centrifuge', entree_1: 'II' }, forced: {} },
    { uid: 'ex', nodeId: 'exelys', choices: { entree_1: 'epaississeur_1', entree_non_digeree: 'epaississeur_2' }, forced: {} },
  ])

essai('Digestion + déshydratation + ANITA Mox sur les jus',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 'am', nodeId: 'anita-mox', choices: { jus_traites: 'digestion', type_digestion: 'simple' }, forced: {} },
  ])

essai('Exelys + déshydratation + ANITA Mox (digestion avancée)',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'centrifuge' }, forced: {} },
    { uid: 'ex', nodeId: 'exelys', choices: { entree_1: 'epaississeur_1' }, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees', digestion_amont: 'avancee' }, forced: {} },
    { uid: 'am', nodeId: 'anita-mox', choices: { jus_traites: 'digestion', type_digestion: 'avancee' }, forced: {} },
  ])

essai('Digestion + déshydratation + ANITA Shunt sur les jus',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 'as', nodeId: 'anita-shunt', choices: { jus_traites: 'digestion' }, forced: {} },
  ])

essai('Dégraisseur + Biolix graisses + épaississement des boues',
  [
    { uid: 'a', nodeId: 'dessablage-deshuilage', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'c', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'bx', nodeId: 'biolix', choices: {}, forced: {} },
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire', entree_1: 'toutes' }, forced: {} },
  ])

essai('Digestion + déshydratation + cristallisation MAP sur les jus',
  [
    { uid: 'a', nodeId: 'decantation-simple', choices: {}, forced: {} },
    { uid: 'b', nodeId: 'ba-faible-charge', choices: {}, forced: {} },
  ],
  [
    { uid: 'e1', nodeId: 'epaississement', choices: { type: 'gravitaire' }, forced: {} },
    { uid: 'd1', nodeId: 'digestion', choices: {}, forced: {} },
    { uid: 'h1', nodeId: 'deshydratation', choices: { type: 'centrifuge', entree_1: 'digerees' }, forced: {} },
    { uid: 'mp', nodeId: 'retours-map', choices: { jus_traites: 'digestion' }, forced: {} },
  ])
