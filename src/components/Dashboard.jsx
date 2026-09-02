import { fmt } from './format.js'
import { POSTES } from '../nodes-transverse/bilan.js'

// ---------------------------------------------------------------------------
// Tableau de bord transverse.
//
// Trois vues, dans l'ordre où on les lit en conception :
//   1. les indicateurs de synthèse, ceux qu'on cite dans un rapport ;
//   2. la répartition électrique poste par poste, puis procédé par procédé ;
//   3. l'empreinte CO2, par compartiment et réactif par réactif.
//
// Le parti pris graphique est de rester en SVG et en CSS, sans dépendance :
// des barres horizontales pour les répartitions, un ruban empilé pour les
// parts. Les couleurs distinguent la file eau, la file boues, les utilités et
// la production d'énergie.
// ---------------------------------------------------------------------------

const COULEUR_FILE = {
  eau: 'var(--water)',
  boues: 'var(--sludge)',
  utilites: 'var(--amber)',
}
const COULEUR_OPEX = {
  electricite: 'var(--water)',
  reactifs: 'var(--amber)',
  combustibles: 'var(--sludge)',
  evacuation: 'var(--ink-2)',
  autres: 'var(--muted)',
}
const COULEUR_COMPARTIMENT = {
  electricite: 'var(--water)',
  reactifs: 'var(--amber)',
  transport: 'var(--ink-2)',
  gaz_naturel: 'var(--sludge)',
  fioul: 'var(--danger)',
}
const LIBELLE_COMPARTIMENT = {
  electricite: 'Électricité',
  reactifs: 'Réactifs',
  transport: 'Transport',
  gaz_naturel: 'Gaz naturel',
  fioul: 'Fioul',
}
const fileDuPoste = (id) => POSTES.find((p) => p.id === id)?.file ?? 'utilites'

/** carte d'indicateur, avec sa valeur et son unité */
function Kpi({ label, value, unit, digits, ton }) {
  return (
    <div className={`tb-kpi${ton ? ` tb-kpi-${ton}` : ''}`}>
      <div className="tb-kpi-label">{label}</div>
      <div className="tb-kpi-value mono">
        {fmt(value, digits)} <small>{unit}</small>
      </div>
    </div>
  )
}

/** barre horizontale proportionnelle, avec sa valeur alignée à droite */
function Barre({ label, valeur, maximum, unit, couleur, part, sousTitre }) {
  const largeur = maximum > 0 ? (Math.abs(valeur) / maximum) * 100 : 0
  const negatif = valeur < 0
  return (
    <div className="tb-barre">
      <div className="tb-barre-tete">
        <span className="tb-barre-label">
          {label}
          {sousTitre && <em className="tb-barre-sous">{sousTitre}</em>}
        </span>
        <span className="tb-barre-valeur mono">
          {fmt(valeur, 0)} {unit}
          {part != null && <em className="tb-barre-part">{fmt(part * 100, 1)} %</em>}
        </span>
      </div>
      <div className="tb-barre-piste">
        <div
          className={`tb-barre-jauge${negatif ? ' tb-barre-negative' : ''}`}
          style={{ width: `${Math.min(100, largeur)}%`, background: couleur }}
        />
      </div>
    </div>
  )
}

/** ruban empilé : une bande par segment, proportionnelle à sa part */
function Ruban({ segments, total }) {
  if (!(total > 0)) return null
  return (
    <div className="tb-ruban">
      {segments.map((s) => {
        const part = s.valeur / total
        if (!(part > 0.001)) return null
        return (
          <div
            key={s.id}
            className="tb-ruban-part"
            style={{ width: `${part * 100}%`, background: s.couleur }}
            title={`${s.label} — ${fmt(s.valeur, 0)} (${fmt(part * 100, 1)} %)`}
          >
            {part > 0.08 && <span>{fmt(part * 100, 0)} %</span>}
          </div>
        )
      })}
    </div>
  )
}

/**
 * Courbe de valeur actuelle nette, année par année. Le franchissement de zéro
 * marque le temps de retour actualisé.
 */
function CourbeVAN({ serie, tpsRetour }) {
  if (!serie?.length) return null
  const L = 640
  const H = 150
  const marge = { g: 52, d: 12, h: 12, b: 22 }
  const min = Math.min(0, ...serie)
  const max = Math.max(0, ...serie)
  const etendue = max - min || 1
  const x = (i) => marge.g + (i / (serie.length - 1)) * (L - marge.g - marge.d)
  const y = (v) => marge.h + ((max - v) / etendue) * (H - marge.h - marge.b)
  const zero = y(0)
  const trace = serie.map((v, i) => `${i === 0 ? 'M' : 'L'} ${x(i).toFixed(1)} ${y(v).toFixed(1)}`).join(' ')
  const aire = `${trace} L ${x(serie.length - 1).toFixed(1)} ${zero.toFixed(1)} L ${x(0).toFixed(1)} ${zero.toFixed(1)} Z`
  return (
    <svg className="tb-courbe" viewBox={`0 0 ${L} ${H}`} role="img" aria-label="Valeur actuelle nette par année">
      <path d={aire} fill="var(--water-soft)" opacity="0.75" />
      <line x1={marge.g} x2={L - marge.d} y1={zero} y2={zero} stroke="var(--line)" strokeWidth="1" />
      <path d={trace} fill="none" stroke="var(--water-deep)" strokeWidth="2" />
      {tpsRetour != null && (
        <g>
          <line
            x1={x(tpsRetour)} x2={x(tpsRetour)} y1={marge.h} y2={H - marge.b}
            stroke="var(--ok)" strokeWidth="1.5" strokeDasharray="4 3"
          />
          <text x={x(tpsRetour) + 5} y={marge.h + 11} fontSize="10" fill="var(--ok)">
            retour {fmt(tpsRetour, 1)} an
          </text>
        </g>
      )}
      <text x={4} y={y(max) + 4} fontSize="10" fill="var(--muted)">{fmt(max, 0)} k€</text>
      <text x={4} y={y(min) + 4} fontSize="10" fill="var(--muted)">{fmt(min, 0)} k€</text>
      <text x={marge.g} y={H - 6} fontSize="10" fill="var(--muted)">an 0</text>
      <text x={L - marge.d - 26} y={H - 6} fontSize="10" fill="var(--muted)">an {serie.length - 1}</text>
    </svg>
  )
}

// ---------------------------------------------------------------------------
export default function Dashboard({ simulation }) {
  if (!simulation) {
    return <div className="tb-vide">Lancer une simulation pour voir le bilan.</div>
  }
  const { bilan, co2, opex, roi, electricite, contexte, eau, boues, utilites } = simulation

  if (!bilan) {
    return (
      <div className="tb-vide">
        Ajouter le nœud « Bilan électrique » au bloc transverse pour voir la synthèse.
      </div>
    )
  }

  // ---- répartition par poste, triée par consommation décroissante
  const postes = [...bilan.postes].sort((a, b) => b.total - a.total)
  const maxPoste = Math.max(...postes.map((p) => Math.abs(p.total)), 1)

  // ---- répartition par procédé, tous moteurs confondus
  const procedes = []
  for (const poste of bilan.postes) {
    for (const d of poste.details) {
      if (Math.abs(d.total) < 1e-9) continue
      procedes.push({ ...d, poste: poste.label, file: fileDuPoste(poste.id) })
    }
  }
  procedes.sort((a, b) => b.total - a.total)
  const maxProcede = Math.max(...procedes.map((p) => Math.abs(p.total)), 1)

  const segmentsFile = ['eau', 'boues', 'utilites'].map((file) => ({
    id: file,
    label: { eau: 'File eau', boues: 'File boues', utilites: 'Utilités' }[file],
    couleur: COULEUR_FILE[file],
    valeur: bilan.postes
      .filter((p) => fileDuPoste(p.id) === file)
      .reduce((s, p) => s + Math.max(0, p.total), 0),
  }))

  return (
    <div className="tb">
      {/* ---- indicateurs de synthèse ---- */}
      <section className="tb-section">
        <h2 className="tb-titre">Synthèse</h2>
        <div className="tb-kpis">
          <Kpi label="Consommation" value={bilan.consommee} unit="kWh/j" digits={0} />
          <Kpi label="Rapportée au débit" value={bilan.ratio_Q} unit="kWh/m³" digits={3} />
          <Kpi label="Par équivalent habitant" value={bilan.ratio_EH} unit="kWh/(EH·an)" digits={0} />
          <Kpi label="Production verte" value={bilan.verte} unit="kWh/j" digits={0} ton="vert" />
          <Kpi label="Autosuffisance" value={bilan.autosuffisance * 100} unit="%" digits={1} ton="vert" />
          <Kpi label="Consommation nette" value={bilan.consommee - bilan.verte_consommee} unit="kWh/j" digits={0} />
          {co2 && <Kpi label="Émissions nettes" value={co2.emissions_nettes} unit="t CO2/an" digits={1} ton="ambre" />}
          {co2 && <Kpi label="Par équivalent habitant" value={co2.ratio_EH} unit="kg CO2/(EH·an)" digits={1} ton="ambre" />}
        </div>
      </section>

      {/* ---- répartition électrique ---- */}
      <section className="tb-section">
        <h2 className="tb-titre">
          Répartition électrique <span className="tb-titre-sous">par poste</span>
        </h2>
        <Ruban segments={segmentsFile} total={bilan.consommee} />
        <div className="tb-legende">
          {segmentsFile.map((s) => (
            <span key={s.id} className="tb-legende-item">
              <i style={{ background: s.couleur }} /> {s.label}
              <em className="mono">{fmt(s.valeur, 0)} kWh/j</em>
            </span>
          ))}
        </div>
        <div className="tb-barres">
          {postes.map((p) => (
            <Barre
              key={p.id}
              label={p.label}
              valeur={p.total}
              maximum={maxPoste}
              unit="kWh/j"
              couleur={p.total < 0 ? 'var(--ok)' : COULEUR_FILE[fileDuPoste(p.id)]}
              part={bilan.consommee > 0 ? p.total / bilan.consommee : null}
            />
          ))}
        </div>
      </section>

      {/* ---- part fixe et variable ---- */}
      <section className="tb-section">
        <h2 className="tb-titre">
          Part fixe et part variable
          <span className="tb-titre-sous">
            ce qui tourne en permanence contre ce qui suit la charge
          </span>
        </h2>
        <Ruban
          segments={[
            { id: 'fixe', label: 'Part fixe', couleur: 'var(--ink-2)', valeur: bilan.fixe },
            { id: 'variable', label: 'Part variable', couleur: 'var(--water)', valeur: bilan.variable },
          ]}
          total={bilan.consommee}
        />
        <div className="tb-legende">
          <span className="tb-legende-item">
            <i style={{ background: 'var(--ink-2)' }} /> Part fixe
            <em className="mono">{fmt(bilan.fixe, 0)} kWh/j</em>
          </span>
          <span className="tb-legende-item">
            <i style={{ background: 'var(--water)' }} /> Part variable
            <em className="mono">{fmt(bilan.variable, 0)} kWh/j</em>
          </span>
        </div>
      </section>

      {/* ---- détail procédé par procédé ---- */}
      <section className="tb-section">
        <h2 className="tb-titre">
          Détail <span className="tb-titre-sous">procédé par procédé</span>
        </h2>
        <table className="tb-table">
          <thead>
            <tr>
              <th>Procédé</th>
              <th>Poste</th>
              <th className="tb-num">Fixe</th>
              <th className="tb-num">Variable</th>
              <th className="tb-num">Total</th>
              <th className="tb-num">Part</th>
              <th className="tb-graph">Répartition</th>
            </tr>
          </thead>
          <tbody>
            {procedes.map((d, i) => {
              const part = bilan.consommee > 0 ? d.total / bilan.consommee : 0
              const largeur = (Math.abs(d.total) / maxProcede) * 100
              return (
                <tr key={`${d.label}-${i}`}>
                  <td>{d.label}</td>
                  <td className="tb-discret">{d.poste}</td>
                  <td className="tb-num mono">{fmt(d.fixe, 0)}</td>
                  <td className="tb-num mono">{fmt(d.variable, 0)}</td>
                  <td className="tb-num mono tb-fort">{fmt(d.total, 0)}</td>
                  <td className="tb-num mono tb-discret">{fmt(part * 100, 1)} %</td>
                  <td className="tb-graph">
                    <div className="tb-barre-piste tb-piste-fine">
                      <div
                        className="tb-barre-jauge"
                        style={{
                          width: `${Math.min(100, largeur)}%`,
                          background: d.total < 0 ? 'var(--ok)' : COULEUR_FILE[d.file],
                        }}
                      />
                    </div>
                  </td>
                </tr>
              )
            })}
          </tbody>
          <tfoot>
            <tr>
              <td colSpan={2}>Total</td>
              <td className="tb-num mono">{fmt(bilan.fixe, 0)}</td>
              <td className="tb-num mono">{fmt(bilan.variable, 0)}</td>
              <td className="tb-num mono tb-fort">{fmt(bilan.consommee, 0)}</td>
              <td colSpan={2} />
            </tr>
          </tfoot>
        </table>
      </section>

      {/* ---- empreinte CO2 ---- */}
      {co2 && (
        <section className="tb-section">
          <h2 className="tb-titre">
            Empreinte carbone
            <span className="tb-titre-sous">
              incertitude ± {fmt(co2.incertitude * 100, 0)} %
            </span>
          </h2>
          <div className="tb-kpis tb-kpis-serres">
            <Kpi label="Émissions brutes" value={co2.emissions_brutes} unit="t CO2/an" digits={1} />
            <Kpi label="Réduites par l'autoconsommation" value={-co2.reduites} unit="t CO2/an" digits={1} ton="vert" />
            <Kpi label="Émissions nettes" value={co2.emissions_nettes} unit="t CO2/an" digits={1} ton="ambre" />
            <Kpi label="Rapportées au débit" value={co2.ratio_Q} unit="g CO2/m³" digits={1} />
          </div>
          <Ruban
            segments={co2.repartition.map((r) => ({
              id: r.id,
              label: LIBELLE_COMPARTIMENT[r.id] ?? r.id,
              couleur: COULEUR_COMPARTIMENT[r.id] ?? 'var(--muted)',
              valeur: Math.max(0, r.valeur),
            }))}
            total={co2.repartition.reduce((s, r) => s + Math.max(0, r.valeur), 0)}
          />
          <div className="tb-legende">
            {co2.repartition.map((r) => (
              <span key={r.id} className="tb-legende-item">
                <i style={{ background: COULEUR_COMPARTIMENT[r.id] ?? 'var(--muted)' }} />
                {LIBELLE_COMPARTIMENT[r.id] ?? r.id}
                <em className="mono">{fmt(r.valeur, 1)} t/an</em>
              </span>
            ))}
          </div>

          {co2.lignes.length > 0 && (
            <table className="tb-table tb-table-serree">
              <thead>
                <tr>
                  <th>Réactif</th>
                  <th className="tb-num">Consommation</th>
                  <th className="tb-num">Production</th>
                  <th className="tb-num">Transport</th>
                  <th className="tb-num">Total</th>
                  <th>Source du facteur</th>
                </tr>
              </thead>
              <tbody>
                {co2.lignes.map((l) => (
                  <tr key={l.cle}>
                    <td>{l.label}</td>
                    <td className="tb-num mono">{fmt(l.tonnes_an, 1)} t/an</td>
                    <td className="tb-num mono">{fmt(l.production, 2)}</td>
                    <td className="tb-num mono">{fmt(l.transport, 2)}</td>
                    <td className="tb-num mono tb-fort">{fmt(l.total, 2)}</td>
                    <td className="tb-discret">{l.source}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </section>
      )}

      {/* ---- coûts d'exploitation ---- */}
      {opex && (
        <section className="tb-section">
          <h2 className="tb-titre">
            Coûts d'exploitation
            <span className="tb-titre-sous">{fmt(opex.opex_net_an / 1000, 0)} k€/an</span>
          </h2>
          <div className="tb-kpis tb-kpis-serres">
            <Kpi label="OPEX net" value={opex.opex_net} unit="€/j" digits={0} />
            <Kpi label="Rapporté au débit" value={opex.ratio_Q} unit="€/m³" digits={3} />
            <Kpi label="Par équivalent habitant" value={opex.ratio_EH} unit="€/(EH·an)" digits={1} />
            <Kpi label="Recettes" value={opex.benefices} unit="€/j" digits={0} ton="vert" />
          </div>
          <Ruban
            segments={opex.postes.map((x) => ({
              id: x.id,
              label: x.label,
              couleur: COULEUR_OPEX[x.id] ?? 'var(--muted)',
              valeur: x.valeur,
            }))}
            total={opex.cout_total}
          />
          <div className="tb-legende">
            {opex.postes.map((x) => (
              <span key={x.id} className="tb-legende-item">
                <i style={{ background: COULEUR_OPEX[x.id] ?? 'var(--muted)' }} />
                {x.label}
                <em className="mono">{fmt(x.valeur, 0)} €/j</em>
              </span>
            ))}
          </div>
          {opex.lignesReactifs.length > 0 && (
            <table className="tb-table tb-table-serree">
              <thead>
                <tr>
                  <th>Réactif</th>
                  <th className="tb-num">Consommation</th>
                  <th className="tb-num">Prix</th>
                  <th className="tb-num">Coût</th>
                  <th className="tb-graph">Part</th>
                </tr>
              </thead>
              <tbody>
                {opex.lignesReactifs.map((l) => {
                  const part = opex.cout_total > 0 ? l.cout_j / opex.cout_total : 0
                  return (
                    <tr key={l.cle}>
                      <td>{l.label}</td>
                      <td className="tb-num mono">{fmt(l.tonnes_an, 1)} t/an</td>
                      <td className="tb-num mono tb-discret">{fmt(l.prix, 0)} €/t</td>
                      <td className="tb-num mono tb-fort">{fmt(l.cout_an / 1000, 1)} k€/an</td>
                      <td className="tb-graph">
                        <div className="tb-barre-piste tb-piste-fine">
                          <div className="tb-barre-jauge" style={{ width: `${part * 100}%`, background: 'var(--amber)' }} />
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          )}
        </section>
      )}

      {/* ---- retour sur investissement ---- */}
      {roi && (
        <section className="tb-section">
          <h2 className="tb-titre">
            Retour sur investissement
            <span className="tb-titre-sous">
              comparaison à la référence sur {roi.duree} ans, actualisation {fmt(roi.taux * 100, 1)} %
            </span>
          </h2>
          <div className="tb-kpis tb-kpis-serres">
            <Kpi label="Économie annuelle" value={roi.economie_an} unit="k€/an" digits={0} ton={roi.economie_an > 0 ? 'vert' : null} />
            {roi.tps_retour != null && <Kpi label="Temps de retour simple" value={roi.tps_retour} unit="an" digits={1} />}
            {roi.tps_retour_actualise != null && <Kpi label="Temps de retour actualisé" value={roi.tps_retour_actualise} unit="an" digits={1} />}
            <Kpi label={`VAN à ${roi.duree} ans`} value={roi.van} unit="k€" digits={0} ton={roi.van >= 0 ? 'vert' : 'ambre'} />
          </div>
          <CourbeVAN serie={roi.serie} tpsRetour={roi.tps_retour_actualise} />
          <table className="tb-table tb-table-serree">
            <thead>
              <tr>
                <th />
                <th className="tb-num">Filière simulée</th>
                <th className="tb-num">Référence</th>
                <th className="tb-num">Écart</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Investissement</td>
                <td className="tb-num mono">{fmt(roi.capex, 0)} k€</td>
                <td className="tb-num mono">{fmt(roi.capex_ref, 0)} k€</td>
                <td className="tb-num mono">{fmt(roi.capex - roi.capex_ref, 0)} k€</td>
              </tr>
              <tr>
                <td>Exploitation</td>
                <td className="tb-num mono">{fmt(roi.opex_an, 0)} k€/an</td>
                <td className="tb-num mono">{fmt(roi.opex_ref_an, 0)} k€/an</td>
                <td className="tb-num mono">{fmt(roi.opex_an - roi.opex_ref_an, 0)} k€/an</td>
              </tr>
              <tr>
                <td>Coût complet sur {roi.duree} ans</td>
                <td className="tb-num mono tb-fort">{fmt(roi.capex_x_opex, 0)} k€</td>
                <td className="tb-num mono tb-fort">{fmt(roi.capex_x_opex_ref, 0)} k€</td>
                <td className="tb-num mono tb-fort">{fmt(roi.capex_x_opex - roi.capex_x_opex_ref, 0)} k€</td>
              </tr>
            </tbody>
          </table>
        </section>
      )}

      {/* ---- contexte de la simulation ---- */}
      <section className="tb-section">
        <h2 className="tb-titre">Contexte</h2>
        <table className="tb-table tb-table-serree">
          <tbody>
            <tr>
              <td>Débit traité</td>
              <td className="tb-num mono">{fmt(contexte.Q_reel, 0)} m³/j</td>
              <td>DCO éliminée</td>
              <td className="tb-num mono">{fmt(contexte.DCO_eliminee, 0)} kg/j</td>
            </tr>
            <tr>
              <td>Azote éliminé</td>
              <td className="tb-num mono">{fmt(contexte.N_elimine, 0)} kg/j</td>
              <td>Phosphore éliminé</td>
              <td className="tb-num mono">{fmt(contexte.P_elimine, 1)} kg/j</td>
            </tr>
            <tr>
              <td>Biogaz produit</td>
              <td className="tb-num mono">{fmt(contexte.biogaz_Nm3j, 0)} Nm³/j</td>
              <td>Boues évacuées</td>
              <td className="tb-num mono">{fmt(contexte.boues_evacuees_MES, 0)} kg MES/j</td>
            </tr>
            <tr>
              <td>Sulfures strippés</td>
              <td className="tb-num mono">{fmt(contexte.HS_strippe_kgj, 1)} kg/j</td>
              <td>Besoin thermique</td>
              <td className="tb-num mono">{fmt(contexte.besoin_thermique_kWhj, 0)} kWh/j</td>
            </tr>
          </tbody>
        </table>
      </section>
    </div>
  )
}
