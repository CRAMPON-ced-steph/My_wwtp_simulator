// ---------------------------------------------------------------------------
// Flux d'eau (charges journalières) et contexte de site — équivalent des
// variables ByRef (Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh) passées de procédé
// en procédé dans MOD_ProgrammePrincipal.prgm_principal.
//
// Convention OCEAN conservée :
//   Q   en m³/j
//   DCO, DBO, MES, NK, NH4, NO3, Pt, Sh en kg/j  (charges, pas concentrations)
// ---------------------------------------------------------------------------

export const STREAM_KEYS = ['Q', 'DCO', 'DBO', 'MES', 'NK', 'NH4', 'NO3', 'Pt', 'Sh']

export const STREAM_LABELS = {
  Q: { label: 'Débit', unit: 'm³/j' },
  DCO: { label: 'DCO', unit: 'kg/j' },
  DBO: { label: 'DBO5', unit: 'kg/j' },
  MES: { label: 'MES', unit: 'kg/j' },
  NK: { label: 'NK', unit: 'kg/j' },
  NH4: { label: 'N-NH4', unit: 'kg/j' },
  NO3: { label: 'N-NO3', unit: 'kg/j' },
  Pt: { label: 'Pt', unit: 'kg/j' },
  Sh: { label: 'S-H2S', unit: 'kg/j' },
}

export function makeStream(o = {}) {
  const s = {}
  for (const k of STREAM_KEYS) s[k] = Number(o[k] ?? 0)
  return s
}

export const cloneStream = (s) => makeStream(s)

/** concentration en mg/L d'un paramètre (0 si Q nul) */
export const conc = (s, key) => (s.Q > 0 ? (s[key] * 1000) / s.Q : 0)

/** construit un flux à partir d'un débit (m³/j) et de concentrations (mg/L) */
export function streamFromConc(Q, c) {
  const s = makeStream({ Q })
  for (const k of STREAM_KEYS) if (k !== 'Q') s[k] = ((c[k] ?? 0) * Q) / 1000
  return s
}

export function scaleStream(s, f) {
  const r = makeStream()
  for (const k of STREAM_KEYS) r[k] = s[k] * f
  return r
}

export function addStreams(a, b) {
  const r = makeStream()
  for (const k of STREAM_KEYS) r[k] = a[k] + b[k]
  return r
}

// ---------------------------------------------------------------------------
// Flux de boues / eaux sales — équivalent des vecteurs eaux_sales(i, repere_ES_*)
// et TableauRecapitulatifFluxBoues. Un flux de boues porte en plus une
// concentration et un ratio MV/MES.
// ---------------------------------------------------------------------------
export function makeSludge(o = {}) {
  return {
    origine: o.origine ?? '',
    Q: o.Q ?? 0, // m³/j
    MES: o.MES ?? 0, // kg/j
    concentration: o.concentration ?? 0, // g/L
    MV_MES: o.MV_MES ?? 0, // -
    DCO: o.DCO ?? 0,
    DBO: o.DBO ?? 0,
    NK: o.NK ?? 0,
    NH4: o.NH4 ?? 0,
    NO3: o.NO3 ?? 0,
    Pt: o.Pt ?? 0,
  }
}

// ---------------------------------------------------------------------------
// Contexte de site : les "données générales" lues depuis Valeurs_générales
// (lecture_donnees_generales dans MOD_ProgrammePrincipal). Chaque nœud y accède
// via ctx.site.
// ---------------------------------------------------------------------------
export const DEFAULT_SITE = {
  // capacité
  Eq_hab: 100000, // EH
  // eau brute nominale (concentrations mg/L) — Q_nominal en m³/j
  Q_nominal: 20000,
  DCO_nominal: 750,
  DBO_nominal: 300,
  MES_nominal: 350,
  NK_nominal: 60,
  NH4_nominal: 40,
  NO3_nominal: 0,
  Pt_nominal: 10,
  HS_nominal_mgL: 0,
  // vidanges (matières de vidange), m³/j et mg/L
  vidange_Q_nominal: 0,
  vidange_MES_mgL_nominal: 0,
  // garanties de rejet (mg/L)
  DCO_garantie: 90,
  DBO_garantie: 25,
  MES_garantie: 30,
  NK_garantie: 10,
  NGL_garantie: 15,
  Pt_garantie: 1,
  /** 'C' = traitement carbone seul, 'CN' = carbone + azote, 'CNP' = + phosphore */
  qualite_rejet: 'CN',
  // conditions
  T_eau_design: 12, // °C
  T_eau_exploit: 15, // °C
  altitude: 100, // m
  pointe_TP: 3,
  pointe_TS: 2, // coefficient de pointe temps sec (Biostyr) // coefficient de pointe hydraulique temps de pluie (-)
  // pourcentage de charge réelle (NC_*) : réel = NC × nominal
  NC_Q: 0.8,
  NC_DCO: 0.8,
  NC_DBO: 0.8,
  NC_MES: 0.8,
  NC_NK: 0.8,
  NC_NH4: 0.8,
  NC_NO3: 0.8,
  NC_Pt: 0.8,
  // débit de retours en tête (m³/j) — Q_retour dans OCEAN
  Q_retour: 0,
}

export function nominalStream(site) {
  return streamFromConc(site.Q_nominal + site.Q_retour, {
    DCO: site.DCO_nominal,
    DBO: site.DBO_nominal,
    MES: site.MES_nominal,
    NK: site.NK_nominal,
    NH4: site.NH4_nominal,
    NO3: site.NO3_nominal,
    Pt: site.Pt_nominal,
    Sh: site.HS_nominal_mgL,
  })
}

export function reelStream(site) {
  const n = nominalStream(site)
  return makeStream({
    Q: n.Q * site.NC_Q,
    DCO: n.DCO * site.NC_DCO,
    DBO: n.DBO * site.NC_DBO,
    MES: n.MES * site.NC_MES,
    NK: n.NK * site.NC_NK,
    NH4: n.NH4 * site.NC_NH4,
    NO3: n.NO3 * site.NC_NO3,
    Pt: n.Pt * site.NC_Pt,
    Sh: n.Sh * site.NC_Q,
  })
}
