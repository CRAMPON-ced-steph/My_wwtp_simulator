export function fmt(v, digits) {
  if (v == null || Number.isNaN(v) || !Number.isFinite(v)) return '—'
  if (digits != null) return v.toLocaleString('fr-FR', { maximumFractionDigits: digits, minimumFractionDigits: 0 })
  const a = Math.abs(v)
  const d = a >= 1000 ? 0 : a >= 100 ? 1 : a >= 10 ? 2 : a >= 1 ? 2 : 3
  return v.toLocaleString('fr-FR', { maximumFractionDigits: d })
}
