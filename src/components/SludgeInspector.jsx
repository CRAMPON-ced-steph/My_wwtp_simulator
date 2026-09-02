import React from 'react'
import { fmt } from './format.js'

export default function SludgeInspector({ node, inst, step, onChange }) {
  if (!node || !inst) {
    return (
      <aside className="inspector">
        <div className="insp-head">
          <div className="insp-title">Filière boues</div>
          <div className="insp-sub">Sélectionner un procédé pour l'inspecter</div>
        </div>
      </aside>
    )
  }

  const groups = []
  for (const prm of node.params) {
    let g = groups.find((x) => x.name === prm.group)
    if (!g) groups.push((g = { name: prm.group, items: [] }))
    g.items.push(prm)
  }

  const setChoice = (key, value) => onChange({ choices: { ...inst.choices, [key]: value } })
  const setForced = (key, value) => onChange({ forced: { ...inst.forced, [key]: value } })

  return (
    <aside className="inspector">
      <div className="insp-head">
        <div className="insp-title">{node.label}</div>
        <div className="insp-sub mono">{node.vba}</div>
        <p className="insp-desc">{node.description}</p>
        {!node.ported && <div className="banner">Calcul non porté. Source VBA dans <code>vba-source/{node.vba}</code>.</div>}
      </div>

      {step?.warnings?.length ? (
        <div className="block warnings">
          {step.warnings.map((w, i) => <div key={i}>⚠ {w}</div>)}
        </div>
      ) : null}

      {step && (
        <div className="block">
          <div className="block-title">Flux boues en sortie</div>
          <table className="tbl"><tbody>
            <tr><td>MES</td><td className="mono">{fmt(step.MES, 1)} kg/j</td></tr>
            <tr><td>Débit</td><td className="mono">{fmt(step.Q, 1)} m³/j</td></tr>
            {step.Q > 0 && <tr><td>Siccité</td><td className="mono">{fmt(step.MES / step.Q * 100, 1)} %</td></tr>}
          </tbody></table>
        </div>
      )}

      {node.choices.length > 0 && (
        <div className="block">
          <div className="block-title">Choix</div>
          {node.choices.map((c) => (
            <label className="field" key={c.key}>
              <span>{c.label}</span>
              <select value={step?.choices?.[c.key] ?? inst.choices?.[c.key] ?? c.default} onChange={(e) => setChoice(c.key, e.target.value)}>
                {c.options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
            </label>
          ))}
        </div>
      )}

      {groups.length > 0 && (
        <div className="block">
          <div className="block-title">Paramètres <span className="legend"><i>défaut</i> · <b>forcé</b> · effectif</span></div>
          {groups.map((g) => (
            <div key={g.name} className="pgroup">
              <div className="pgroup-name">{g.name}</div>
              <table className="tbl params">
                <tbody>
                  {g.items.map((prm) => {
                    const d = step?.defaults?.[prm.key]
                    const e = step?.p?.[prm.key]
                    const f = inst.forced?.[prm.key] ?? ''
                    return (
                      <tr key={prm.key} className={f !== '' ? 'forced' : ''}>
                        <td>
                          <div className="plabel">{prm.label}</div>
                          {prm.hint && <div className="phint">{prm.hint}</div>}
                        </td>
                        <td className="mono dim">{d == null ? '·' : fmt(d)}</td>
                        <td>
                          <input type="number" step="any" placeholder="forcer" value={f} onChange={(ev) => setForced(prm.key, ev.target.value)} />
                        </td>
                        <td className="mono eff">{e == null ? '·' : fmt(e)} <small>{prm.unit}</small></td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          ))}
        </div>
      )}

      {step?.results?.length ? (
        <div className="block">
          <div className="block-title">Résultats</div>
          <table className="tbl"><tbody>
            {step.results.map((r) => (
              <tr key={r.key}><td>{r.label}</td><td className="mono">{fmt(r.value)} {r.unit}</td></tr>
            ))}
          </tbody></table>
        </div>
      ) : null}

      {step?.electricity && (
        <div className="block">
          <div className="block-title">Électricité</div>
          <table className="tbl"><tbody>
            {Object.entries(step.electricity.detail || {}).map(([k, v]) => (
              <tr key={k}><td>{k}</td><td className="mono">{fmt(v, 1)} kWh/j</td></tr>
            ))}
            <tr className="total"><td>Total</td><td className="mono">{fmt(step.electricity.total, 1)} kWh/j</td></tr>
          </tbody></table>
        </div>
      )}

      {step?.reactifs && Object.keys(step.reactifs).length > 0 && (
        <div className="block">
          <div className="block-title">Réactifs</div>
          <table className="tbl"><tbody>
            {Object.entries(step.reactifs).map(([k, v]) => (
              <tr key={k}><td>{k}</td><td className="mono">{fmt(v, 1)}</td></tr>
            ))}
          </tbody></table>
        </div>
      )}
    </aside>
  )
}
