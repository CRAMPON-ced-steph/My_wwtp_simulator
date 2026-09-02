import React, { useState } from 'react'
import { SLUDGE_REGISTRY } from '../nodes-boues/index.js'
import { fmt } from './format.js'

function DropZone({ index, onInsert, onMove, wide }) {
  const [over, setOver] = useState(false)
  return (
    <div
      className={`dropzone ${over ? 'over' : ''} ${wide ? 'wide' : ''}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('application/x-ocean-boues-node') || e.dataTransfer.types.includes('application/x-ocean-boues-uid')) {
          e.preventDefault()
          setOver(true)
        }
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => {
        e.preventDefault()
        setOver(false)
        const nodeId = e.dataTransfer.getData('application/x-ocean-boues-node')
        const uid = e.dataTransfer.getData('application/x-ocean-boues-uid')
        if (nodeId) onInsert(nodeId, index)
        else if (uid) onMove(uid, index)
      }}
    >
      {wide ? <span>Déposer un procédé ici</span> : <span>+</span>}
    </div>
  )
}

function SludgeNodeCard({ inst, step, selected, onSelect, onRemove }) {
  const node = SLUDGE_REGISTRY[inst.nodeId]
  const results = (step?.results || []).slice(0, 3)
  return (
    <div
      className={`node ${selected ? 'selected' : ''} ${node?.ported ? '' : 'todo'} fam-${node?.family}`}
      draggable
      onDragStart={(e) => {
        e.dataTransfer.setData('application/x-ocean-boues-uid', inst.uid)
        e.dataTransfer.effectAllowed = 'move'
      }}
      onClick={() => onSelect(inst.uid)}
    >
      <div className="node-head">
        <span className="node-title">{node?.label ?? inst.nodeId}</span>
        <button className="node-x" title="Retirer" onClick={(e) => { e.stopPropagation(); onRemove(inst.uid) }}>×</button>
      </div>
      <div className="node-vba mono">{node?.vba}</div>
      {node?.ported ? (
        <div className="node-results">
          {step && (
            <>
              <div className="node-result">
                <span>MES sortie</span>
                <span className="mono">{fmt(step.MES, 1)} kg/j</span>
              </div>
              <div className="node-result">
                <span>Débit sortie</span>
                <span className="mono">{fmt(step.Q, 1)} m³/j</span>
              </div>
            </>
          )}
          {results.map((r) => (
            <div className="node-result" key={r.key}>
              <span>{r.label}</span>
              <span className="mono">{fmt(r.value)} {r.unit}</span>
            </div>
          ))}
        </div>
      ) : (
        <div className="node-todo">Calcul non porté</div>
      )}
      <div className="node-foot">
        <span className="mono">{fmt(step?.electricity?.total ?? 0, 0)} kWh/j</span>
        {step?.warnings?.length ? <span className="node-warn" title={step.warnings.join('\n')}>⚠ {step.warnings.length}</span> : null}
      </div>
    </div>
  )
}

export default function SludgeCanvas({ chain, sim, apportsMES, selected, onSelect, onInsert, onMove, onRemove }) {
  const { evacuation } = sim
  return (
    <main className="canvas">
      <div className="chain">
        <div className="node inlet">
          <div className="node-head"><span className="node-title">Apports filière eau</span></div>
          <div className="node-vba mono">boues primaires + biologiques</div>
          <div className="node-results">
            <div className="node-result"><span>MES totales</span><span className="mono">{fmt(apportsMES, 1)} kg/j</span></div>
          </div>
        </div>
        <DropZone index={0} onInsert={onInsert} onMove={onMove} />
        {chain.map((inst, i) => {
          const step = sim.steps.find((s) => s.uid === inst.uid)
          return (
            <React.Fragment key={inst.uid}>
              <SludgeNodeCard
                inst={inst}
                step={step}
                selected={selected === inst.uid}
                onSelect={onSelect}
                onRemove={onRemove}
              />
              <DropZone index={i + 1} onInsert={onInsert} onMove={onMove} />
            </React.Fragment>
          )
        })}
        <div className="node outlet">
          <div className="node-head"><span className="node-title">Évacuation</span></div>
          <div className="node-vba mono">boues non traitées restantes</div>
          <div className="node-results">
            <div className="node-result"><span>MES évacuées</span><span className="mono">{fmt(evacuation.MES, 1)} kg/j</span></div>
            <div className="node-result"><span>Volume</span><span className="mono">{fmt(evacuation.Q, 1)} m³/j</span></div>
            <div className="node-result"><span>MV/MES</span><span className="mono">{fmt(evacuation.MV_MES * 100, 1)} %</span></div>
            <div className="node-result"><span>Siccité</span><span className="mono">{fmt(evacuation.siccite * 100, 1)} %</span></div>
          </div>
        </div>
        {chain.length === 0 && <DropZone index={0} onInsert={onInsert} onMove={onMove} wide />}
      </div>
    </main>
  )
}
