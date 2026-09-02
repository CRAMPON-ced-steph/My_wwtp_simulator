import React, { useMemo, useState, useCallback } from 'react'
import { NODE_LIST, REGISTRY } from './nodes/index.js'
import { SLUDGE_NODE_LIST, SLUDGE_REGISTRY } from './nodes-boues/index.js'
import { UTILITY_REGISTRY } from './nodes-utilites/index.js'
import { TRANSVERSE_REGISTRY } from './nodes-transverse/index.js'
import { FAMILIES, runChain } from './core/engine.js'
import { SLUDGE_FAMILIES, runSludgeChain, apportsDepuisFileEau } from './core/sludgeEngine.js'
import { simuler } from './core/simulation.js'
import { DEFAULT_SITE } from './core/stream.js'
import Palette from './components/Palette.jsx'
import Canvas from './components/Canvas.jsx'
import Inspector from './components/Inspector.jsx'
import Dashboard from './components/Dashboard.jsx'
import SludgeCanvas from './components/SludgeCanvas.jsx'
import SludgeInspector from './components/SludgeInspector.jsx'
import { fmt } from './components/format.js'

let uidCounter = 1
const newUid = () => `n${uidCounter++}`

const inst = (nodeId, choices = {}) => ({ uid: newUid(), nodeId, choices, forced: {} })
const DEFAULT_CHAIN = ['dessablage-deshuilage', 'decantation-simple', 'ba-forte-charge', 'discfilter', 'desinfection-uv'].map((nodeId) => inst(nodeId))
// Filière boues, utilités et transverse par défaut : elles alimentent le
// tableau de bord dès l'ouverture, sans configuration préalable.
const DEFAULT_BOUES = [
  inst('epaississement', { type: 'gravitaire' }),
  inst('digestion'),
  inst('deshydratation', { type: 'centrifuge', entree_1: 'digerees' }),
]
const DEFAULT_UTILITES = [inst('desodorisation-bio'), inst('gestion-reactifs')]
const DEFAULT_TRANSVERSE = [inst('bilan-electrique'), inst('empreinte-co2'), inst('gestion-opex'), inst('retour-investissement')]
const REGISTRES = { eau: REGISTRY, boues: SLUDGE_REGISTRY, utilites: UTILITY_REGISTRY, transverse: TRANSVERSE_REGISTRY }

export default function App() {
  const [site, setSite] = useState(DEFAULT_SITE)
  const [chain, setChain] = useState(DEFAULT_CHAIN)
  const [selected, setSelected] = useState('inlet')
  const [mode, setMode] = useState('reel') // 'nominal' | 'reel'
  const [vue, setVue] = useState('filiere') // 'filiere' | 'boues' | 'bilan'

  // filière boues éditable
  const [bouesChain, setBouesChain] = useState(DEFAULT_BOUES)
  const [selectedBoues, setSelectedBoues] = useState(null)

  const sim = useMemo(() => runChain(chain, REGISTRY, site), [chain, site])

  const apports = useMemo(() => apportsDepuisFileEau(sim, REGISTRY), [sim])
  const apportsMES = useMemo(() => apports.reduce((s, a) => s + a.MES, 0), [apports])
  const simBoues = useMemo(() => runSludgeChain(bouesChain, SLUDGE_REGISTRY, site, apports), [bouesChain, site, apports])

  // la simulation complète n'est calculée que si le tableau de bord est ouvert
  const simulation = useMemo(
    () => (vue === 'bilan'
      ? simuler({ eau: chain, boues: bouesChain, utilites: DEFAULT_UTILITES, transverse: DEFAULT_TRANSVERSE }, REGISTRES, site)
      : null),
    [vue, chain, bouesChain, site],
  )

  const insertAt = useCallback((nodeId, index) => {
    setChain((c) => {
      const inst = { uid: newUid(), nodeId, choices: {}, forced: {} }
      const next = [...c]
      next.splice(index, 0, inst)
      setSelected(inst.uid)
      return next
    })
  }, [])
  const moveTo = useCallback((uid, index) => {
    setChain((c) => {
      const from = c.findIndex((x) => x.uid === uid)
      if (from < 0) return c
      const next = [...c]
      const [inst] = next.splice(from, 1)
      const to = index > from ? index - 1 : index
      next.splice(to, 0, inst)
      return next
    })
  }, [])
  const remove = useCallback((uid) => {
    setChain((c) => c.filter((x) => x.uid !== uid))
    setSelected((s) => (s === uid ? 'inlet' : s))
  }, [])
  const updateInst = useCallback((uid, patch) => {
    setChain((c) => c.map((x) => (x.uid === uid ? { ...x, ...patch } : x)))
  }, [])

  const insertBouesAt = useCallback((nodeId, index) => {
    setBouesChain((c) => {
      const i = { uid: newUid(), nodeId, choices: {}, forced: {} }
      const next = [...c]
      next.splice(index, 0, i)
      setSelectedBoues(i.uid)
      return next
    })
  }, [])
  const moveBouesTo = useCallback((uid, index) => {
    setBouesChain((c) => {
      const from = c.findIndex((x) => x.uid === uid)
      if (from < 0) return c
      const next = [...c]
      const [i] = next.splice(from, 1)
      const to = index > from ? index - 1 : index
      next.splice(to, 0, i)
      return next
    })
  }, [])
  const removeBoues = useCallback((uid) => {
    setBouesChain((c) => c.filter((x) => x.uid !== uid))
    setSelectedBoues((s) => (s === uid ? null : s))
  }, [])
  const updateBouesInst = useCallback((uid, patch) => {
    setBouesChain((c) => c.map((x) => (x.uid === uid ? { ...x, ...patch } : x)))
  }, [])

  const selectedInst = chain.find((x) => x.uid === selected)
  const selectedStep = sim.steps.find((x) => x.uid === selected)

  const selectedBouesInst = bouesChain.find((x) => x.uid === selectedBoues)
  const selectedBouesStep = simBoues.steps.find((x) => x.uid === selectedBoues)

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-sub">Simulateur filière eau</span>
        </div>
        <div className="topbar-right">
          <div className="seg">
            <button className={vue === 'filiere' ? 'on' : ''} onClick={() => setVue('filiere')}>Filière eau</button>
            <button className={vue === 'boues' ? 'on' : ''} onClick={() => setVue('boues')}>Filière boues</button>
            <button className={vue === 'bilan' ? 'on' : ''} onClick={() => setVue('bilan')}>Bilan</button>
          </div>
          <div className="seg">
            <button className={mode === 'nominal' ? 'on' : ''} onClick={() => setMode('nominal')}>Nominal</button>
            <button className={mode === 'reel' ? 'on' : ''} onClick={() => setMode('reel')}>Réel</button>
          </div>
          <div className="kpi">
            <span className="kpi-label">Électricité filière</span>
            <span className="kpi-value mono">{fmt(sim.electricityTotal, 0)} <small>kWh/j</small></span>
          </div>
          <div className="kpi">
            <span className="kpi-label">Spécifique</span>
            <span className="kpi-value mono">{fmt(sim.outReel.Q > 0 ? (sim.electricityTotal * 1000) / sim.outReel.Q : 0, 0)} <small>Wh/m³</small></span>
          </div>
        </div>
      </header>
      {vue === 'bilan' ? (
        <Dashboard simulation={simulation} />
      ) : vue === 'boues' ? (
        <div className="workspace">
          <Palette
            nodes={SLUDGE_NODE_LIST}
            families={SLUDGE_FAMILIES}
            onAdd={(nodeId) => insertBouesAt(nodeId, bouesChain.length)}
            dragType="application/x-ocean-boues-node"
          />
          <SludgeCanvas
            chain={bouesChain}
            sim={simBoues}
            apportsMES={apportsMES}
            selected={selectedBoues}
            onSelect={setSelectedBoues}
            onInsert={insertBouesAt}
            onMove={moveBouesTo}
            onRemove={removeBoues}
          />
          <SludgeInspector
            node={selectedBouesInst ? SLUDGE_REGISTRY[selectedBouesInst.nodeId] : null}
            inst={selectedBouesInst}
            step={selectedBouesStep}
            onChange={(patch) => selectedBouesInst && updateBouesInst(selectedBouesInst.uid, patch)}
          />
        </div>
      ) : (
        <div className="workspace">
          <Palette nodes={NODE_LIST} families={FAMILIES} onAdd={(nodeId) => insertAt(nodeId, chain.length)} />
          <Canvas
            chain={chain}
            sim={sim}
            site={site}
            mode={mode}
            selected={selected}
            onSelect={setSelected}
            onInsert={insertAt}
            onMove={moveTo}
            onRemove={remove}
          />
          <Inspector
            site={site}
            onSite={setSite}
            inst={selectedInst}
            step={selectedStep}
            node={selectedInst ? REGISTRY[selectedInst.nodeId] : null}
            mode={mode}
            onChange={(patch) => selectedInst && updateInst(selectedInst.uid, patch)}
            outlet={mode === 'reel' ? sim.outReel : sim.outNominal}
            isInlet={selected === 'inlet'}
          />
        </div>
      )}
    </div>
  )
}
