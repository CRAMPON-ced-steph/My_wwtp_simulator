# OCEAN — port React de la filière eau

Port des modules VBA du classeur `140822_OCEAN_CCR.xlsm` (Veolia Water Tech) en nœuds React
autonomes, assemblables par glisser-déposer pour simuler une station.

```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # dist/
```

## Architecture

```
src/core/stream.js       flux d'eau (Q m³/j, charges kg/j), flux de boues, données de site (Valeurs_générales)
src/core/hypotheses.js   onglet Hypothèses + constantes pd_*, ratio() (AA_collection), fonctions communes :
                         facteur K, saturation O2, Patm, rendement moteur, répartition pompes, précipitation P
src/core/engine.js       defineNode(), runChain() : enchaîne les nœuds en double passe
                         (dimensionnement sur l'eau nominale, fonctionnement réel sur l'eau × NC_*)
src/nodes/*.js           un fichier par procédé (voir tableau)
src/nodes/_stub.js       fabrique de nœud "à porter" (traversée sans modification)
src/components/          Palette (glisser-déposer), Canvas (chaîne + tuyaux), Inspector (Vp / Vr / Ve, résultats, élec)
vba-source/              classes VBA découpées par procédé + modules communs + Hypotheses.csv
```

### Contrat d'un nœud

```js
defineNode({
  id, label, family, vba, description,
  choices: [{ key, label, options:[{value,label}], default }],          // boutons btn_* / *_choix_*
  params:  [{ key, label, unit, group, default: (ctx) => number, hint }], // Vp (défaut) — l'utilisateur force Vr
  compute(ctx) => { outNominal, outReel, sludge?, eauxSales?, reactifs?, results, electricity, warnings }
})
// ctx = { site, inNominal, inReel, choices, p (=Ve), forced (=Vr), defaults, upstream }
```

`p[key]` est la valeur effective (forcée si présente, sinon défaut) ; `forced[key]` permet de
reproduire les branches `If xxx_force = True`. `upstream.primaire` reproduit `choix_primaire`.

## État du port

| Nœud | Classe VBA | État |
|---|---|---|
| Dessablage – déshuilage | C2_Dessablage_Deshuilage | porté |
| Décantation simple / lamellaire | D1_decanteur_simple | porté |
| Décantation avec réactifs | D2_Decantation_reactif | squelette |
| BA forte charge | E1_BA_forte_charge | porté |
| BA moyenne charge | E2_BA_moyenne_charge | squelette |
| BA faible charge | E3_BA_faible_charge (+ z_CLS_EauBouesActiveesATVa131) | squelette |
| Aération prolongée | E4_BA_aeration_prolongee | squelette |
| HybAS | E5_HybAS | squelette |
| MBBR | E6_MBBR | squelette |
| MBR | — (absent du classeur, seul `ratio("II_MBR")` existe) | squelette |
| Biostyr | E8_Biostyr | squelette |
| Biostyr PDN | E9_Biostyr_PDN | squelette |
| Biostyr nitrifiant III | F1_Biostyr_N_III | squelette |
| Biostyr PDN III | F2_Biostyr_PDN_III | squelette |
| Décantation tertiaire | F3_Decantation_III | porté |
| Discfilter | F4_Discfilter_III | porté |
| Filtration sable | F5_Filtration_sable | porté |
| Désinfection UV | G1_Desinfection_UV | porté |
| Chloration | G2_Desinfection_Cl | porté |
| Décantation eaux sales | H1_Decantation_eaux_sales | squelette |

Les squelettes déclarent déjà leurs paramètres forçables (cellules `CelVF_*` de l'onglet Calculs)
et laissent l'eau inchangée. Pour porter un module : ouvrir `vba-source/<classe>.cls`, transcrire
`hypotheses` / `attribution_valeur_par_defaut` dans `params`, puis `dimensionnement`,
`fonctionnement_reel` et `calcul_consommation_electrique` dans `compute()`.

## Écarts assumés par rapport au VBA

- **C2** : `graisse_DCO` utilise `ratio("codigestion_graisses","NK_MV")` (0,002) dans le VBA, ce qui
  est manifestement une erreur ; le port utilise `DCO_MV` (2,8).
- **E1** : les `If T<12 … If T<18 …` sans `ElseIf` écrasent la valeur T<12. Le port reproduit ce
  comportement (`VBA_BUG_COMPAT = true` dans `baForteCharge.js`) ; passer à `false` pour 3 / 2,5 / 1,5 g/L.
- **E1** : dans le VBA, `boues_Q` est soustrait de Q avant d'être calculé (valeur de l'appel
  précédent). Le port calcule les boues avant de mettre à jour Q.
- **D1** : le module CCR référence des variables non déclarées (`boues_concentration`,
  `NombreOuvrages`…) ; le port les mappe sur les paramètres du bloc Calculs.
- Les retours d'eaux sales en tête (`eaux_sales`, `Q_retour`) ne sont pas encore rebouclés : les
  nœuds Discfilter / filtre sable exposent leur flux `eauxSales`, à raccorder au nœud
  Décantation eaux sales ou à l'entrée.
- Les consommations spécifiques surpresseurs sont celles codées dans E1 (roots 4,5) et non celles de
  l'onglet Hypothèses (roots 5).
