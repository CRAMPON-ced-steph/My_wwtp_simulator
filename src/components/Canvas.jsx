import React, { useState } from 'react'
import { conc } from '../core/stream.js'
import { REGISTRY } from '../nodes/index.js'
import { fmt } from './format.js'

const PIPE_KEYS = ['DCO', 'DBO', 'MES', 'NK', 'NH4', 'NO3', 'Pt']

function Pipe({ stream }) {
  return (
    <div className="pipe">
      <div className="pipe-line" />
      <div className="pipe-tag mono">
        <span className="pipe-q">{fmt(stream.Q, 0)} m³/j</span>
        {PIPE_KEYS.map((k) => (
          <span key={k}>
            <b>{k}</b> {fmt(conc(stream, k), 1)}
          </span>
        ))}
      </div>
    </div>
  )
}

function DropZone({ index, onInsert, onMove, wide }) {
  const [over, setOver] = useState(false)
  return (
    <div
      className={`dropzone ${over ? 'over' : ''} ${wide ? 'wide' : ''}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('application/x-ocean-node') || e.dataTransfer.types.includes('application/x-ocean-uid')) {
          e.preventDefault()
          setOver(true)
        }
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => {
        e.preventDefault()
        setOver(false)
        const nodeId = e.dataTransfer.getData('application/x-ocean-node')
        const uid = e.dataTransfer.getData('application/x-ocean-uid')
        if (nodeId) onInsert(nodeId, index)
        else if (uid) onMove(uid, index)
      }}
    >
      {wide ? <span>Déposer un procédé ici</span> : <span>+</span>}
    </div>
  )
}

function NodeCard({ inst, step, selected, onSelect, onRemove, mode }) {
  const node = REGISTRY[inst.nodeId]
  const results = (step?.results || []).slice(0, 4)
  return (
    <div
      className={`node ${selected ? 'selected' : ''} ${node.ported ? '' : 'todo'} fam-${node.family}`}
      draggable
      onDragStart={(e) => {
        e.dataTransfer.setData('application/x-ocean-uid', inst.uid)
        e.dataTransfer.effectAllowed = 'move'
      }}
      onClick={() => onSelect(inst.uid)}
    >
      <div className="node-head">
        <span className="node-title">{node.label}</span>
        <button className="node-x" title="Retirer" onClick={(e) => { e.stopPropagation(); onRemove(inst.uid) }}>×</button>
      </div>
      <div className="node-vba mono">{node.vba}</div>
      {node.ported ? (
        <div className="node-results">
          {results.map((r) => (
            <div className="node-result" key={r.key}>
              <span>{r.label}</span>
              <span className="mono">{fmt(r.value)} {r.unit}</span>
            </div>
          ))}
        </div>
      ) : (
        <div className="node-todo">Calcul non porté — traversée sans modification</div>
      )}
      <div className="node-foot">
        <span className="mono">{fmt(step?.electricity?.total ?? 0, 0)} kWh/j</span>
        {step?.warnings?.length ? <span className="node-warn" title={step.warnings.join('\n')}>⚠ {step.warnings.length}</span> : null}
      </div>
    </div>
  )
}

export default function Canvas({ chain, sim, site, mode, selected, onSelect, onInsert, onMove, onRemove }) {
  const first = mode === 'reel' ? (sim.steps[0]?.inReel ?? sim.outReel) : (sim.steps[0]?.inNominal ?? sim.outNominal)
  return (
    <main className="canvas">
      <div className="chain">
        <div className={`node inlet ${selected === 'inlet' ? 'selected' : ''}`} onClick={() => onSelect('inlet')}>
          <div className="node-head"><span className="node-title">Eau brute</span></div>
          <div className="node-vba mono">Valeurs_générales</div>
          <div className="node-results">
            <div className="node-result"><span>Capacité</span><span className="mono">{fmt(site.Eq_hab, 0)} EH</span></div>
            <div className="node-result"><span>Q nominal</span><span className="mono">{fmt(site.Q_nominal, 0)} m³/j</span></div>
            <div className="node-result"><span>T eau design / exploit.</span><span className="mono">{site.T_eau_design} / {site.T_eau_exploit} °C</span></div>
            <div className="node-result"><span>Charge réelle (Q)</span><span className="mono">{fmt(site.NC_Q * 100, 0)} %</span></div>
          </div>
        </div>
        <Pipe stream={first} />
        <DropZone index={0} onInsert={onInsert} onMove={onMove} />
        {chain.map((inst, i) => {
          const step = sim.steps.find((s) => s.uid === inst.uid)
          const out = step ? (mode === 'reel' ? step.outReel : step.outNominal) : first
          return (
            <React.Fragment key={inst.uid}>
              <NodeCard inst={inst} step={step} selected={selected === inst.uid} onSelect={onSelect} onRemove={onRemove} mode={mode} />
              <Pipe stream={out} />
              <DropZone index={i + 1} onInsert={onInsert} onMove={onMove} />
            </React.Fragment>
          )
        })}
        <div className="node outlet">
          <div className="node-head"><span className="node-title">Rejet</span></div>
          <div className="node-vba mono">{mode === 'reel' ? 'fonctionnement réel' : 'dimensionnement'}</div>
          <div className="node-results">
            {[['DCO', site.DCO_garantie], ['DBO', site.DBO_garantie], ['MES', site.MES_garantie], ['NK', site.NK_garantie], ['Pt', site.Pt_garantie]].map(([k, g]) => {
              const s = mode === 'reel' ? sim.outReel : sim.outNominal
              const v = conc(s, k)
              return (
                <div className={`node-result ${v > g ? 'bad' : 'good'}`} key={k}>
                  <span>{k} <small>(garantie {g})</small></span>
                  <span className="mono">{fmt(v, 1)} mg/L</span>
                </div>
              )
            })}
          </div>
        </div>
        {chain.length === 0 && <DropZone index={0} onInsert={onInsert} onMove={onMove} wide />}
      </div>
    </main>
  )
}
