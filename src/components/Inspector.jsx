import React from 'react'
import { STREAM_KEYS, STREAM_LABELS, conc } from '../core/stream.js'
import { fmt } from './format.js'

const SITE_FIELDS = [
  { group: 'Capacité', fields: [['Eq_hab', 'Équivalents-habitants', 'EH'], ['Q_nominal', 'Débit nominal', 'm³/j'], ['Q_retour', 'Retours en tête', 'm³/j']] },
  { group: 'Eau brute nominale (mg/L)', fields: [['DCO_nominal', 'DCO', 'mg/L'], ['DBO_nominal', 'DBO5', 'mg/L'], ['MES_nominal', 'MES', 'mg/L'], ['NK_nominal', 'NK', 'mg/L'], ['NH4_nominal', 'N-NH4', 'mg/L'], ['NO3_nominal', 'N-NO3', 'mg/L'], ['Pt_nominal', 'Pt', 'mg/L'], ['HS_nominal_mgL', 'Sulfures (S)', 'mg/L']] },
  { group: 'Matières de vidange', fields: [['vidange_Q_nominal', 'Débit', 'm³/j'], ['vidange_MES_mgL_nominal', 'MES', 'mg/L']] },
  { group: 'Garanties de rejet (mg/L)', fields: [['DCO_garantie', 'DCO', 'mg/L'], ['DBO_garantie', 'DBO5', 'mg/L'], ['MES_garantie', 'MES', 'mg/L'], ['NK_garantie', 'NK', 'mg/L'], ['NGL_garantie', 'NGL', 'mg/L'], ['Pt_garantie', 'Pt', 'mg/L']] },
  { group: 'Conditions', fields: [['T_eau_design', 'T eau dimensionnement', '°C'], ['T_eau_exploit', 'T eau exploitation', '°C'], ['altitude', 'Altitude', 'm'], ['pointe_TP', 'Coefficient de pointe temps de pluie', '-']] },
  { group: 'Charge réelle (fraction du nominal)', fields: [['NC_Q', 'Débit', '-'], ['NC_DCO', 'DCO', '-'], ['NC_DBO', 'DBO5', '-'], ['NC_MES', 'MES', '-'], ['NC_NK', 'NK', '-'], ['NC_NH4', 'N-NH4', '-'], ['NC_NO3', 'N-NO3', '-'], ['NC_Pt', 'Pt', '-']] },
]

function StreamTable({ title, stream }) {
  return (
    <div className="block">
      <div className="block-title">{title}</div>
      <table className="tbl">
        <thead><tr><th>Paramètre</th><th>Charge</th><th>Concentration</th></tr></thead>
        <tbody>
          {STREAM_KEYS.map((k) => (
            <tr key={k}>
              <td>{STREAM_LABELS[k].label}</td>
              <td className="mono">{fmt(stream[k])} {STREAM_LABELS[k].unit}</td>
              <td className="mono">{k === 'Q' ? '' : `${fmt(conc(stream, k), 1)} mg/L`}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function SiteEditor({ site, onSite, outlet }) {
  return (
    <>
      <div className="insp-head">
        <div className="insp-title">Eau brute et données générales</div>
        <div className="insp-sub mono">Valeurs_générales · lecture_donnees_generales</div>
      </div>
      <div className="block">
        <label className="field">
          <span>Qualité de rejet visée</span>
          <select value={site.qualite_rejet} onChange={(e) => onSite({ ...site, qualite_rejet: e.target.value })}>
            <option value="C">Carbone seul</option>
            <option value="CN">Carbone + azote</option>
            <option value="CNP">Carbone + azote + phosphore</option>
          </select>
        </label>
      </div>
      {SITE_FIELDS.map((g) => (
        <div className="block" key={g.group}>
          <div className="block-title">{g.group}</div>
          {g.fields.map(([key, label, unit]) => (
            <label className="field" key={key}>
              <span>{label}</span>
              <span className="field-in">
                <input type="number" step="any" value={site[key]} onChange={(e) => onSite({ ...site, [key]: Number(e.target.value) })} />
                <em>{unit}</em>
              </span>
            </label>
          ))}
        </div>
      ))}
      <StreamTable title="Eau traitée en sortie de filière" stream={outlet} />
    </>
  )
}

function NodeEditor({ node, inst, step, mode, onChange }) {
  const groups = []
  for (const prm of node.params) {
    let g = groups.find((x) => x.name === prm.group)
    if (!g) groups.push((g = { name: prm.group, items: [] }))
    g.items.push(prm)
  }
  const setChoice = (key, value) => onChange({ choices: { ...inst.choices, [key]: value } })
  const setForced = (key, value) => onChange({ forced: { ...inst.forced, [key]: value } })
  const stream = step ? (mode === 'reel' ? step.outReel : step.outNominal) : null
  return (
    <>
      <div className="insp-head">
        <div className="insp-title">{node.label}</div>
        <div className="insp-sub mono">{node.vba}</div>
        <p className="insp-desc">{node.description}</p>
        {!node.ported && <div className="banner">Calcul non porté : le nœud laisse l'eau inchangée. Source VBA dans <code>vba-source/{node.vba}</code>.</div>}
      </div>
      {step?.warnings?.length ? (
        <div className="block warnings">
          {step.warnings.map((w, i) => <div key={i}>⚠ {w}</div>)}
        </div>
      ) : null}
      {node.choices.length > 0 && (
        <div className="block">
          <div className="block-title">Choix</div>
          {node.choices.map((c) => (
            <label className="field" key={c.key}>
              <span>{c.label}</span>
              <select value={step?.choices?.[c.key] ?? c.default} onChange={(e) => setChoice(c.key, e.target.value)}>
                {c.options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
            </label>
          ))}
        </div>
      )}
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
      {step?.results?.length ? (
        <div className="block">
          <div className="block-title">Résultats</div>
          <table className="tbl">
            <tbody>
              {step.results.map((r) => (
                <tr key={r.key}><td>{r.label}</td><td className="mono">{fmt(r.value)} {r.unit}</td></tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
      {step?.electricity && (
        <div className="block">
          <div className="block-title">Électricité</div>
          <table className="tbl">
            <tbody>
              {Object.entries(step.electricity.detail || {}).map(([k, v]) => (
                <tr key={k}><td>{k}</td><td className="mono">{fmt(v, 1)} kWh/j</td></tr>
              ))}
              <tr className="total"><td>Total</td><td className="mono">{fmt(step.electricity.total, 1)} kWh/j</td></tr>
              {step.electricity.fixed != null && <tr><td>dont part fixe</td><td className="mono">{fmt(step.electricity.fixed, 1)} kWh/j</td></tr>}
            </tbody>
          </table>
        </div>
      )}
      {step?.sludge && (
        <div className="block">
          <div className="block-title">Boues extraites ({mode === 'reel' ? 'réel' : 'réel'})</div>
          <table className="tbl"><tbody>
            <tr><td>Origine</td><td className="mono">{step.sludge.origine}</td></tr>
            <tr><td>MES</td><td className="mono">{fmt(step.sludge.MES)} kg/j</td></tr>
            <tr><td>Débit</td><td className="mono">{fmt(step.sludge.Q)} m³/j</td></tr>
            <tr><td>Concentration</td><td className="mono">{fmt(step.sludge.concentration)} g/L</td></tr>
            <tr><td>MV/MES</td><td className="mono">{fmt(step.sludge.MV_MES)}</td></tr>
          </tbody></table>
        </div>
      )}
      {step?.eauxSales && (
        <div className="block">
          <div className="block-title">Eaux sales (réel)</div>
          <table className="tbl"><tbody>
            {['Q', 'MES', 'DCO', 'DBO', 'NK', 'NH4', 'NO3', 'Pt'].map((k) => (
              <tr key={k}><td>{k}</td><td className="mono">{fmt(step.eauxSales[k])} {k === 'Q' ? 'm³/j' : 'kg/j'}</td></tr>
            ))}
          </tbody></table>
        </div>
      )}
      {stream && <StreamTable title={`Eau en sortie (${mode === 'reel' ? 'réel' : 'nominal'})`} stream={stream} />}
    </>
  )
}

export default function Inspector({ site, onSite, inst, step, node, mode, onChange, outlet, isInlet }) {
  return (
    <aside className="inspector">
      {isInlet || !inst ? <SiteEditor site={site} onSite={onSite} outlet={outlet} /> : <NodeEditor node={node} inst={inst} step={step} mode={mode} onChange={onChange} />}
    </aside>
  )
}
