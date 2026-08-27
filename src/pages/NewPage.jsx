import React from 'react'
import { Link } from 'react-router-dom'

export default function NewPage() {
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark">my_wwtp_simulator</span>
          <span className="brand-sub">Nouvelle page</span>
        </div>
        <div className="topbar-right">
          <Link to="/" className="nav-link">← Simulateur</Link>
        </div>
      </header>
      <div style={{ padding: 32 }}>
        {/* Contenu de la nouvelle page */}
      </div>
    </div>
  )
}
