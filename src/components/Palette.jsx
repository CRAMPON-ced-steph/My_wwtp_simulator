import React from 'react'

export default function Palette({ nodes, families, onAdd, dragType = 'application/x-ocean-node' }) {
  return (
    <aside className="palette">
      <div className="pane-title">Procédés</div>
      <p className="pane-hint">Glisser un procédé sur la filière, ou double-cliquer pour l'ajouter en fin de chaîne.</p>
      {families.map((f) => {
        const items = nodes.filter((n) => n.family === f.id)
        if (!items.length) return null
        return (
          <div className="palette-group" key={f.id}>
            <div className="palette-family">{f.label}</div>
            {items.map((n) => (
              <div
                key={n.id}
                className={`palette-item ${n.ported ? '' : 'todo'}`}
                draggable
                onDragStart={(e) => {
                  e.dataTransfer.setData(dragType, n.id)
                  e.dataTransfer.effectAllowed = 'copy'
                }}
                onDoubleClick={() => onAdd(n.id)}
                title={n.description}
              >
                <span className="palette-item-label">{n.label}</span>
                <span className="palette-item-vba mono">{n.ported ? n.vba : 'à porter'}</span>
              </div>
            ))}
          </div>
        )
      })}
    </aside>
  )
}
