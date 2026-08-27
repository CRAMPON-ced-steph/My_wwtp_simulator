import React, { useMemo, useState, useCallback } from 'react'
import { NODE_LIST, REGISTRY } from './nodes/index.js'
import { FAMILIES, runChain } from './core/engine.js'
import { DEFAULT_SITE } from './core/stream.js'
import Palette from './components/Palette.jsx'
import Canvas from './components/Canvas.jsx'
import Inspector from './components/Inspector.jsx'
import { fmt } from './components/format.js'

let uidCounter = 1
const newUid = () => `n${uidCounter++}`

const DEFAULT_CHAIN = ['dessablage-deshuilage', 'decantation-simple', 'ba-forte-charge', 'discfilter', 'desinfection-uv'].map((nodeId) => ({ uid: newUid(), nodeId, choices: {}, forced: {} }))

export default function App() {
  const [site, setSite] = useState(DEFAULT_SITE)
  const [chain, setChain] = useState(DEFAULT_CHAIN)
  const [selected, setSelected] = useState('inlet')
  const [mode, setMode] = useState('reel') // 'nominal' | 'reel'

  const sim = useMemo(() => runChain(chain, REGISTRY, site), [chain, site])

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

  const selectedInst = chain.find((x) => x.uid === selected)
  const selectedStep = sim.steps.find((x) => x.uid === selected)

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark">my_wwtp_simulator</span>
          <span className="brand-sub">Simulateur filière eau · port React des modules VBA</span>
        </div>
        <div className="topbar-right">
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
    </div>
  )
}
