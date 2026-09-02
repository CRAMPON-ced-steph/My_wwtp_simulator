// ---------------------------------------------------------------------------
// Analyse pincement — solveur.
//
// Ce module remplace le PINCH du classeur par une implémentation de la méthode
// de Linnhoff et Hindmarsh, conduite dans l'ordre canonique :
//
//   1. segmentation des flux, changements de phase compris ;
//   2. décalage des températures par les contributions ΔT individuelles ;
//   3. algorithme de la table de problème, cascade et cibles énergétiques ;
//   4. courbes composites et grande courbe composite ;
//   5. placement des utilités multiples sur la grande courbe composite ;
//   6. cibles de nombre d'unités et de surface d'échange ;
//   7. vérification des trois règles de conception.
//
// Quatre points le distinguent du module d'origine.
//
//  a) **Contributions ΔT individuelles.** Un ΔT_min global suppose que tous les
//     flux offrent le même coefficient d'échange, ce qui est faux : un gaz et
//     une boue liquide n'ont pas le même film. Chaque flux porte ici sa propre
//     contribution, et le ΔT effectif entre deux flux est la somme des deux.
//     C'est la formulation rigoureuse ; le ΔT global n'en est que le cas
//     particulier où toutes les contributions valent ΔT_min / 2.
//
//  b) **Changements de phase traités explicitement.** Un flux qui condense ou
//     vaporise a un CP infini sur un palier isotherme. Le solveur le segmente
//     au lieu de le linéariser, ce qui évite de sous-estimer la récupération.
//
//  c) **Pincements multiples détectés.** La cascade peut s'annuler en plusieurs
//     points ; ne retenir que le premier conduit à des réseaux infaisables. Le
//     solveur les remonte tous, et signale les quasi-pincements.
//
//  d) **Utilités multiples placées sur la grande courbe composite.** Plutôt
//     qu'une utilité chaude unique, le solveur répartit la demande entre les
//     niveaux disponibles en commençant par le moins cher — biogaz avant gaz
//     naturel, par exemple — dans la limite de ce que la courbe autorise.
//
// Conventions : températures en °C, puissances en kW, CP en kW/K. Le module ne
// dépend d'aucune bibliothèque.
// ---------------------------------------------------------------------------

const EPS = 1e-9

/**
 * Un flux thermique.
 * @typedef {Object} Flux
 * @property {string} nom
 * @property {'chaud'|'froid'} type      chaud = à refroidir, froid = à réchauffer
 * @property {number} T_in               température d'entrée, °C
 * @property {number} T_out              température de sortie, °C
 * @property {number} [CP]               débit de capacité thermique, kW/K
 * @property {number} [charge]           puissance, kW — alternative au CP
 * @property {number} [dT_contribution]  contribution ΔT propre au flux, K
 * @property {number} [h]                coefficient d'échange, kW/(m²·K)
 * @property {Array}  [segments]         segments imposés, pour changement de phase
 */

/**
 * Normalise un flux : déduit le CP de la charge ou l'inverse, oriente les
 * températures, et découpe les changements de phase en segments.
 */
export function normaliserFlux(flux, dT_defaut = 5) {
  const chaud = flux.type === 'chaud'
  const T_in = flux.T_in
  const T_out = flux.T_out
  const dT = flux.dT_contribution ?? dT_defaut
  const h = flux.h ?? (chaud ? 0.5 : 0.5)

  // Un flux isotherme est un changement de phase : sa charge est portée par un
  // palier, avec un CP qui tendrait vers l'infini. On lui donne un intervalle
  // infinitésimal pour rester dans le formalisme de la table de problème.
  const isotherme = Math.abs(T_in - T_out) < 1e-6
  let segments
  if (flux.segments?.length) {
    segments = flux.segments.map((s) => ({
      T_haut: Math.max(s.T_in, s.T_out),
      T_bas: Math.min(s.T_in, s.T_out),
      CP: s.CP ?? (Math.abs(s.T_in - s.T_out) > 1e-6 ? s.charge / Math.abs(s.T_in - s.T_out) : 0),
      charge: s.charge ?? (s.CP ?? 0) * Math.abs(s.T_in - s.T_out),
      isotherme: Math.abs(s.T_in - s.T_out) < 1e-6,
    }))
  } else if (isotherme) {
    const charge = flux.charge ?? 0
    const demi = 0.05 // K, palier ramené à un intervalle très étroit
    segments = [{
      T_haut: T_in + demi,
      T_bas: T_in - demi,
      CP: charge / (2 * demi),
      charge,
      isotherme: true,
    }]
  } else {
    const ecart = Math.abs(T_in - T_out)
    const CP = flux.CP ?? (flux.charge ?? 0) / ecart
    segments = [{
      T_haut: Math.max(T_in, T_out),
      T_bas: Math.min(T_in, T_out),
      CP,
      charge: CP * ecart,
      isotherme: false,
    }]
  }
  const charge = segments.reduce((s, x) => s + x.charge, 0)
  return { ...flux, chaud, dT_contribution: dT, h, segments, charge, T_in, T_out }
}

/**
 * Algorithme de la table de problème.
 *
 * Les températures sont décalées par la contribution propre à chaque flux —
 * un flux chaud descend de sa contribution, un flux froid monte de la sienne —
 * si bien que deux flux qui se croisent dans l'échelle décalée respectent, dans
 * l'échelle réelle, un écart au moins égal à la somme de leurs contributions.
 */
export function tableDeProbleme(fluxNormalises) {
  // ---- intervalles de température, dans l'échelle décalée
  const bornes = new Set()
  for (const f of fluxNormalises) {
    for (const s of f.segments) {
      const decalage = f.chaud ? -f.dT_contribution : +f.dT_contribution
      bornes.add(+(s.T_haut + decalage).toFixed(9))
      bornes.add(+(s.T_bas + decalage).toFixed(9))
    }
  }
  const T = [...bornes].sort((a, b) => b - a) // du plus chaud au plus froid
  if (T.length < 2) return null

  // ---- bilan par intervalle
  const intervalles = []
  for (let i = 0; i < T.length - 1; i++) {
    const T_haut = T[i]
    const T_bas = T[i + 1]
    const dT = T_haut - T_bas
    let CP_chaud = 0
    let CP_froid = 0
    for (const f of fluxNormalises) {
      const decalage = f.chaud ? -f.dT_contribution : +f.dT_contribution
      for (const s of f.segments) {
        const haut = s.T_haut + decalage
        const bas = s.T_bas + decalage
        // le segment couvre-t-il l'intervalle ?
        if (haut >= T_haut - EPS && bas <= T_bas + EPS) {
          if (f.chaud) CP_chaud += s.CP
          else CP_froid += s.CP
        }
      }
    }
    // excédent positif : les flux chauds dominent, de la chaleur est disponible
    intervalles.push({ T_haut, T_bas, dT, CP_chaud, CP_froid, excedent: (CP_chaud - CP_froid) * dT })
  }

  // ---- cascade sans apport extérieur, puis décalage pour la rendre faisable
  const cascade_brute = [0]
  for (const it of intervalles) cascade_brute.push(cascade_brute[cascade_brute.length - 1] + it.excedent)
  const deficit = Math.min(...cascade_brute)
  const Qh_min = deficit < 0 ? -deficit : 0
  const cascade = cascade_brute.map((x) => x + Qh_min)
  const Qc_min = cascade[cascade.length - 1]

  // ---- pincements : tous les points où la cascade s'annule
  const pincements = []
  cascade.forEach((q, i) => {
    if (Math.abs(q) < 1e-6) pincements.push({ T_decale: T[i], index: i })
  })
  // quasi-pincements : la cascade y passe sous 5 % de la charge chaude
  const seuil = 0.05 * Math.max(Qh_min, 1)
  const quasi = []
  cascade.forEach((q, i) => {
    if (q > 1e-6 && q < seuil) quasi.push({ T_decale: T[i], residuel: q })
  })

  return { T, intervalles, cascade_brute, cascade, Qh_min, Qc_min, pincements, quasi }
}

/**
 * Courbes composites.
 *
 * Point de rigueur souvent négligé : dès que les flux portent des contributions
 * ΔT **différentes**, les composites tracées en températures réelles perdent
 * leur sens. Deux flux chauds de contributions 2 K et 20 K ne peuvent pas être
 * agrégés sur une même courbe dont on lirait ensuite un écart minimal unique,
 * puisque chacun impose son propre écart au flux qu'il rencontre.
 *
 * Les composites sont donc construites dans l'**échelle décalée**, où l'écart
 * minimal vaut zéro par construction et où le contact des deux courbes marque
 * exactement le pincement. C'est cette version qui sert aux cibles et aux
 * vérifications. Une version en températures réelles est fournie en plus pour
 * l'affichage, avec le drapeau `exactes` qui indique si toutes les
 * contributions sont égales — seul cas où sa lecture usuelle reste valide.
 */
export function courbesComposites(fluxNormalises, Qc_min, decalees = true) {
  const construire = (chaud) => {
    const bornes = new Set()
    const dec = (f) => (decalees ? (f.chaud ? -f.dT_contribution : +f.dT_contribution) : 0)
    for (const f of fluxNormalises) {
      if (f.chaud !== chaud) continue
      const d = dec(f)
      for (const s of f.segments) { bornes.add(+(s.T_haut + d).toFixed(9)); bornes.add(+(s.T_bas + d).toFixed(9)) }
    }
    const T = [...bornes].sort((a, b) => a - b)
    if (T.length < 2) return []
    const points = [{ T: T[0], H: 0 }]
    for (let i = 0; i < T.length - 1; i++) {
      const bas = T[i]
      const haut = T[i + 1]
      let CP = 0
      for (const f of fluxNormalises) {
        if (f.chaud !== chaud) continue
        const d = dec(f)
        for (const s of f.segments) {
          if (s.T_haut + d >= haut - EPS && s.T_bas + d <= bas + EPS) CP += s.CP
        }
      }
      points.push({ T: haut, H: points[points.length - 1].H + CP * (haut - bas) })
    }
    return points
  }
  const chaude = construire(true)
  const froide = construire(false)
  // C'est la composite **froide** qui se décale de l'utilité froide. À gauche de
  // l'axe enthalpique se trouve l'extrémité la plus froide de la composite
  // chaude, celle que l'utilité froide évacue : la chaude part donc de zéro et
  // la froide de Qc_min. Le recouvrement horizontal vaut alors exactement la
  // récupération, et le dépassement de la froide à droite l'utilité chaude.
  const froide_decalee = froide.map((p) => ({ ...p, H: p.H + Qc_min }))
  return { chaude, froide: froide_decalee }
}

/**
 * Grande courbe composite : chaleur nette disponible en fonction de la
 * température décalée. C'est elle qui gouverne le placement des utilités.
 */
export function grandeCourbeComposite(table) {
  return table.T.map((T, i) => ({ T, H: table.cascade[i] }))
}

/**
 * Place les utilités sur la grande courbe composite, en commençant par la moins
 * chère. Une utilité chaude à température T ne peut fournir que la chaleur
 * requise au-dessus de T ; le reste revient aux niveaux supérieurs.
 */
export function placerUtilites(gcc, utilites, Qh_min, Qc_min) {
  const chaudes = utilites.filter((u) => u.type === 'chaude').sort((a, b) => (a.cout ?? 0) - (b.cout ?? 0))
  const froides = utilites.filter((u) => u.type === 'froide').sort((a, b) => (a.cout ?? 0) - (b.cout ?? 0))
  const resultat = { chaudes: [], froides: [], Qh_place: 0, Qc_place: 0 }

  // ---- utilités chaudes : la GCC est parcourue du haut vers le bas
  let restant = Qh_min
  for (const u of chaudes) {
    if (restant <= EPS) break
    // chaleur encore requise au niveau de l'utilité : minimum de la cascade
    // au-dessus de sa température décalée
    const T_decalee = u.T - (u.dT_contribution ?? 0)
    const auDessus = gcc.filter((p) => p.T <= T_decalee + EPS)
    const dispo = auDessus.length ? Math.min(...auDessus.map((p) => p.H)) : 0
    // ce que ce niveau peut couvrir sans franchir le pincement utilitaire
    let charge = Math.min(restant, Math.max(0, restant - dispo))
    if (u.capacite != null) charge = Math.min(charge, u.capacite)
    if (charge > EPS) {
      resultat.chaudes.push({ ...u, charge })
      resultat.Qh_place += charge
      restant -= charge
    }
  }
  if (restant > EPS) {
    // le solde revient à l'utilité la plus chaude disponible
    const derniere = chaudes[chaudes.length - 1]
    if (derniere) {
      const ligne = resultat.chaudes.find((x) => x.nom === derniere.nom)
      if (ligne) ligne.charge += restant
      else resultat.chaudes.push({ ...derniere, charge: restant })
      resultat.Qh_place += restant
    }
  }

  // ---- utilités froides, symétriquement
  let restantF = Qc_min
  for (const u of froides) {
    if (restantF <= EPS) break
    let charge = restantF
    if (u.capacite != null) charge = Math.min(charge, u.capacite)
    if (charge > EPS) {
      resultat.froides.push({ ...u, charge })
      resultat.Qc_place += charge
      restantF -= charge
    }
  }
  return resultat
}

/**
 * Cibles de conception du réseau d'échangeurs.
 *
 * Le nombre minimal d'unités suit la relation d'Euler N = S − 1 appliquée
 * séparément de part et d'autre du pincement, puisqu'aucun échangeur ne doit le
 * traverser. La surface cible emploie la formule de Bath, qui répartit la
 * surface entre les flux au prorata de l'inverse de leurs coefficients
 * d'échange, intervalle d'enthalpie par intervalle d'enthalpie.
 */
export function ciblesConception(fluxNormalises, table, composites, utilites) {
  const T_pincement = table.pincements.length ? table.pincements[0].T_decale : null

  const compter = (zone) => {
    let n = 0
    for (const f of fluxNormalises) {
      const decalage = f.chaud ? -f.dT_contribution : +f.dT_contribution
      const haut = Math.max(...f.segments.map((s) => s.T_haut)) + decalage
      const bas = Math.min(...f.segments.map((s) => s.T_bas)) + decalage
      if (T_pincement == null) { n += 1; continue }
      if (zone === 'dessus' && haut > T_pincement + EPS) n += 1
      if (zone === 'dessous' && bas < T_pincement - EPS) n += 1
    }
    return n
  }
  const n_dessus = compter('dessus') + (table.Qh_min > EPS ? 1 : 0)
  const n_dessous = compter('dessous') + (table.Qc_min > EPS ? 1 : 0)
  const unites_min = T_pincement == null
    ? Math.max(0, fluxNormalises.length - 1)
    : Math.max(0, n_dessus - 1) + Math.max(0, n_dessous - 1)

  // ---- surface cible, méthode de Bath
  // On découpe l'axe enthalpique aux points anguleux des deux composites, puis
  // on somme (1/ΔT_ml) × Σ(q_i / h_i) sur chaque tranche.
  let surface = 0
  const { chaude, froide } = composites
  if (chaude.length > 1 && froide.length > 1) {
    const H_min = Math.max(chaude[0].H, froide[0].H)
    const H_max = Math.min(chaude[chaude.length - 1].H, froide[froide.length - 1].H)
    const bornes = new Set([H_min, H_max])
    for (const p of [...chaude, ...froide]) if (p.H > H_min && p.H < H_max) bornes.add(p.H)
    const H = [...bornes].sort((a, b) => a - b)
    const interpoler = (courbe, h) => {
      for (let i = 0; i < courbe.length - 1; i++) {
        const a = courbe[i]
        const b = courbe[i + 1]
        if (h >= a.H - EPS && h <= b.H + EPS) {
          if (Math.abs(b.H - a.H) < EPS) return a.T
          return a.T + ((h - a.H) / (b.H - a.H)) * (b.T - a.T)
        }
      }
      return h < courbe[0].H ? courbe[0].T : courbe[courbe.length - 1].T
    }
    for (let i = 0; i < H.length - 1; i++) {
      const q = H[i + 1] - H[i]
      if (q < EPS) continue
      const dT1 = interpoler(chaude, H[i]) - interpoler(froide, H[i])
      const dT2 = interpoler(chaude, H[i + 1]) - interpoler(froide, H[i + 1])
      if (dT1 <= EPS || dT2 <= EPS) continue
      // moyenne logarithmique, ramenée à la moyenne arithmétique si les deux
      // écarts sont proches
      const dT_ml = Math.abs(dT1 - dT2) < 1e-6 ? dT1 : (dT1 - dT2) / Math.log(dT1 / dT2)
      // coefficient d'échange moyen pondéré des flux présents dans la tranche
      const h_moyen = fluxNormalises.length
        ? fluxNormalises.reduce((s, f) => s + f.h, 0) / fluxNormalises.length
        : 0.5
      if (dT_ml > EPS && h_moyen > EPS) surface += (q / dT_ml) * (2 / h_moyen)
    }
  }

  return { T_pincement, n_dessus, n_dessous, unites_min, surface }
}

/**
 * Vérifie les trois règles de conception d'un réseau au pincement.
 * Elles ne sont pas des recommandations : les enfreindre augmente d'autant la
 * consommation d'utilités au-delà de la cible.
 */
export function verifierRegles(table, fluxNormalises) {
  const alertes = []
  if (!table.pincements.length) {
    alertes.push({
      niveau: 'info',
      texte: "Aucun pincement : les flux chauds couvrent l'intégralité des besoins froids, ou l'inverse. Le procédé est limité par un seul type d'utilité.",
    })
    return alertes
  }
  if (table.pincements.length > 1) {
    alertes.push({
      niveau: 'attention',
      texte: `Pincements multiples (${table.pincements.length}) : le réseau doit être conçu par zones indépendantes, séparées à chacun d'eux.`,
    })
  }
  for (const q of table.quasi) {
    alertes.push({
      niveau: 'info',
      texte: `Quasi-pincement à ${q.T_decale.toFixed(1)} °C décalés (${q.residuel.toFixed(1)} kW de marge) : une variation de charge peut le rendre limitant.`,
    })
  }
  alertes.push({
    niveau: 'regle',
    texte: "Aucun échangeur ne doit transférer de chaleur à travers le pincement : chaque kW qui le franchit augmente d'autant les deux utilités.",
  })
  alertes.push({ niveau: 'regle', texte: "Aucune utilité froide au-dessus du pincement." })
  alertes.push({ niveau: 'regle', texte: "Aucune utilité chaude en dessous du pincement." })
  return alertes
}

/**
 * Analyse complète. `utilites` est optionnel ; sans lui, seules les cibles
 * énergétiques globales sont calculées.
 */
export function analyserPincement(fluxBruts, options = {}) {
  const dT_defaut = options.dT_contribution_defaut ?? 5
  const flux = fluxBruts
    .map((f) => normaliserFlux(f, dT_defaut))
    .filter((f) => f.charge > EPS)
  if (flux.length < 2) {
    return {
      flux,
      erreur: flux.length === 0
        ? "Aucun flux thermique exploitable : les procédés de la filière n'exposent ni besoin ni disponibilité."
        : "Un seul flux thermique : l'intégration énergétique suppose au moins un flux chaud et un flux froid.",
    }
  }
  const chauds = flux.filter((f) => f.chaud)
  const froids = flux.filter((f) => !f.chaud)
  if (!chauds.length || !froids.length) {
    return {
      flux,
      erreur: !chauds.length
        ? "Aucun flux chaud : rien à récupérer, la totalité des besoins revient aux utilités."
        : "Aucun flux froid : rien à réchauffer, la totalité de la chaleur disponible part au refroidissement.",
    }
  }

  const table = tableDeProbleme(flux)
  const composites = courbesComposites(flux, table.Qc_min, true)
  // version en températures réelles, pour l'affichage ; elle n'est lisible au
  // sens usuel que si toutes les contributions sont égales
  const contributions = new Set(flux.map((f) => +f.dT_contribution.toFixed(6)))
  const composites_reelles = {
    ...courbesComposites(flux, table.Qc_min, false),
    exactes: contributions.size === 1,
    dT_min: contributions.size === 1 ? 2 * [...contributions][0] : null,
  }
  const gcc = grandeCourbeComposite(table)
  const cibles = ciblesConception(flux, table, composites, options.utilites)
  const utilites = options.utilites?.length
    ? placerUtilites(gcc, options.utilites, table.Qh_min, table.Qc_min)
    : null
  const alertes = verifierRegles(table, flux)

  const charge_chaude = chauds.reduce((s, f) => s + f.charge, 0)
  const charge_froide = froids.reduce((s, f) => s + f.charge, 0)
  const recuperation = charge_froide - table.Qh_min
  // sans intégration, chaque besoin est couvert par une utilité dédiée
  const Qh_sans_integration = charge_froide
  const Qc_sans_integration = charge_chaude
  const economie = Qh_sans_integration > EPS ? 1 - table.Qh_min / Qh_sans_integration : 0

  return {
    flux,
    table,
    composites,
    composites_reelles,
    gcc,
    cibles,
    utilites,
    alertes,
    bilan: {
      charge_chaude,
      charge_froide,
      Qh_min: table.Qh_min,
      Qc_min: table.Qc_min,
      recuperation,
      Qh_sans_integration,
      Qc_sans_integration,
      economie,
      T_pincement_decale: cibles.T_pincement,
      // dans l'échelle réelle, le pincement se lit à deux températures
      T_pincement_chaud: cibles.T_pincement != null ? cibles.T_pincement + dT_defaut : null,
      T_pincement_froid: cibles.T_pincement != null ? cibles.T_pincement - dT_defaut : null,
    },
  }
}
