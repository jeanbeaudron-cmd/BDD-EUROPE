# TECHNICAL.md — Category Review Explorer

> Spec d'implémentation pour Claude Code. Architecture **Option B** :
> DuckDB-WASM dans le navigateur, requêtant le Parquet canonique. Sans backend.

---

## 1. Architecture

```
Stage 1 (Python, hors-ligne)        Stage 2 (navigateur, single-page)
build_category_db.py  ──────────▶   category_db.parquet
   (lit 5 Excel,                         │
    normalise en long)                   ▼
                                    index.html
                                      ├─ DuckDB-WASM (lit le parquet en RAM)
                                      ├─ Query builder (selections → SQL)
                                      ├─ Dashboards pré-faits
                                      └─ Coverage badge / Volume guardrail
```

- ~2,75 M de lignes / ~7 Mo de Parquet → tient sans souci en RAM navigateur.
- **Lancement** : serveur statique local obligatoire (le `fetch` du parquet est bloqué
  en `file://`). `python3 -m http.server` dans le dossier suffit.
- Pinner la version de `@duckdb/duckdb-wasm` (vendor local de préférence) pour ne pas
  dépendre du CDN en prod.

## 2. Schéma de la base (format long)

Une ligne = une mesure. Colonnes produites par `build_category_db.py` :

| Colonne | Type | Rôle |
|---|---|---|
| `country` | str | FR/UK/DE/ES/IT |
| `channel` | str | Total / Hyper |
| `market` | str | YALG / Kefir / Bifidus |
| `level` | str | family…ean (niveau de la ligne) |
| `depth` | int | rang de profondeur (pour comparer les niveaux) |
| `is_mdd` | bool | la ligne est-elle MDD |
| `mdd_detailed` | bool | la MDD est-elle détaillée sous le fabricant **dans cette base** |
| `manufacturer`,`brand`,`segment`,`flavor_en`,`weight`,`multipack`,`bio`,`ean` | str | attributs |
| `variable` | str | code mesure canonique |
| `var_family` | str | value / volume_* / units / price / distr / rotn |
| `agg_rule` | str | **additive / derived / node_only** |
| `year` | int | 2024/2025/2026 |
| `value` | float | la valeur |

Charger en vue SQL nommée `db` (voir §6).

## 3. Les deux régimes d'agrégation (cœur du moteur)

### Régime 1 — nœud nommé : LIRE, ne pas sommer
À `level ∈ {family, category, manufacturer, brand, segment}`, chaque ligne est déjà un
total pré-calculé (une ligne par nœud × pays × circuit × marché × année).

```sql
-- Classement fabricants par valeur (FR / YALG / Total / 2026)
SELECT manufacturer, value AS sales_value
FROM db
WHERE country='FR' AND market='YALG' AND channel='Total'
  AND level='manufacturer' AND variable='sales_value' AND year=2026
ORDER BY value DESC;
```

### Régime 2 — cross-cut d'attribut : SOMMER au niveau de l'attribut
```sql
-- Mix parfums par pays (YALG / Total / 2026)
SELECT country, flavor_en, SUM(value) AS sales_value
FROM db
WHERE market='YALG' AND channel='Total' AND level='flavor'
  AND variable='sales_value' AND year=2026
GROUP BY country, flavor_en;
```
Le `level='flavor'` cible exactement les sous-totaux parfum et évite tout double comptage
avec les niveaux plus fins (item/ean). Idem `level='weight'`/`'multipack'` pour le format,
`level='bio'` pour le bio.

## 4. Coverage badge & MDD (PAR PAYS)

Tout cross-cut (Régime 2) affiche la couverture vs le total marché, **par pays**.

> **Dénominateur robuste = `SUM(value)` au niveau `manufacturer`.** On a prouvé sur les
> 5 pays que le niveau fabricant pave le marché à 100 %. Ne PAS utiliser `category` :
> les libellés `CATEGORY` dupliqués (DE/IT) éclatent ce niveau en plusieurs nœuds et
> faussent le total (on obtenait des couvertures >200 %). Le numérateur, lui, se somme au
> **niveau propre du cross-cut** (`flavor` / `weight` / `bio`).

```sql
WITH ref AS (                       -- total marché robuste
  SELECT country, SUM(value) AS market_total FROM db
  WHERE market='YALG' AND channel='Total' AND level='manufacturer'
    AND variable='sales_value' AND year=2026 GROUP BY country),
flav AS (                           -- numérateur au niveau de l'attribut
  SELECT country, SUM(value) AS detailed FROM db
  WHERE market='YALG' AND channel='Total' AND level='flavor'
    AND variable='sales_value' AND year=2026 GROUP BY country)
SELECT r.country, f.detailed, r.market_total,
       f.detailed / r.market_total AS coverage
FROM ref r JOIN flav f USING(country) ORDER BY coverage;
-- validé : FR 39% · UK 100% · DE 100% · IT 100%  (couverture PARFUM)
```
- `coverage < 1` → badge *« partial — excludes MDD / undetailed »*.
- La couverture est **propre à chaque attribut** : en UK, le parfum tile 100 % alors que
  l'EAN ne couvre que 43 % (plus de produits ont un parfum qu'un UPC). Toujours recalculer
  au niveau du cross-cut affiché.
- Joindre `mdd_detailed` (constant par pays) pour le motif : `false` (FR, ES) → MDD absente
  du cross-cut ; `true` (UK, DE, IT) → l'écart ne vient que du non-détaillé.
- **Ne jamais** agréger la couverture en un seul chiffre Europe : une ligne par pays.

## 5. Total Europe (additif uniquement) + parts recalculées

```sql
-- Lactalis/Danone sur YALG/Total : valeur par pays + Europe + part recalculée
-- Dénominateur = SUM(manufacturer) (total marché robuste, cf. §4)
WITH node AS (
  SELECT country, year, SUM(value) AS v FROM db
  WHERE market='YALG' AND channel='Total' AND level='manufacturer'
    AND manufacturer ILIKE '%DANONE%' AND variable='sales_value'
  GROUP BY country, year),
ref AS (
  SELECT country, year, SUM(value) AS v FROM db
  WHERE market='YALG' AND channel='Total' AND level='manufacturer'
    AND variable='sales_value' GROUP BY country, year)
-- détail pays
SELECT node.country, node.year, node.v AS val, node.v/ref.v AS value_share
FROM node JOIN ref USING(country, year)
UNION ALL
-- Europe : additives sommées, part recalculée (jamais moyennée), qualifier les colonnes
SELECT 'Europe', node.year, SUM(node.v),
       SUM(node.v) / (SELECT SUM(v) FROM ref r2 WHERE r2.year = node.year)
FROM node GROUP BY node.year;
-- validé : Europe Danone 2026 ≈ 86,4 M€, part 3,3%
```

> `value` est un mot réservé en DuckDB → ne pas l'utiliser comme alias (utiliser `val`).

Règles applicatives Europe :
- **additive** (`sales_value`, `sales_units`) → `SUM`.
- **derived** (parts, prix) → recalcul depuis additives (jamais `AVG`).
- **node_only** (DN, DV, ROS, *_yoy, *_chg) → ne pas agréger ; afficher seulement si un
  pays unique / nœud unique est sélectionné.

## 6. Volume guardrail

Les bases volume sont incompatibles → variable distincte par pays :

```js
const VOLUME_BASIS = {            // var code → pays partageant la base
  sales_vol_eq: ['FR'],          // Volume EQ
  sales_vol_kg: ['UK', 'ES'],    // KGS (seule base partagée)
  sales_vol_de: ['DE'],          // Absatz GESAMT
  sales_vol_it: ['IT'],          // V. (ALL)
};
```
Règle UI : si la mesure choisie est un volume, n'autoriser dans le scope que les pays
d'une même base (ex. KGS → cocher UK et/ou ES). En scope « Europe » multi-base, **griser
le volume** et proposer valeur (EUR) ou unités à la place.

## 7. Year-as-dimension (sortie builder)

Le format long rend le pivot trivial — l'année devient une colonne d'affichage à la volée :

```sql
SELECT manufacturer,
  SUM(value) FILTER (WHERE year=2024) AS "2024",
  SUM(value) FILTER (WHERE year=2025) AS "2025",
  SUM(value) FILTER (WHERE year=2026) AS "2026"
FROM db
WHERE country='FR' AND market='YALG' AND channel='Total'
  AND level='manufacturer' AND variable='sales_value'
GROUP BY manufacturer ORDER BY "2026" DESC;
```
Pour le graphe : `year` en abscisse (série temporelle) **ou** le découpage en abscisse et
`year` en séries — c'est le même résultat long, pivoté différemment côté front.

## 8. Contrat du query builder

Une fonction pure `selections → SQL` à implémenter, signature cible :

```js
buildQuery({
  countries,        // ['FR','DE',...]  ('Europe' = rollup additif)
  markets,          // ['YALG']
  channel,          // 'Total' | 'Hyper'
  breakdown,        // {kind:'level', value:'manufacturer'}
                    // | {kind:'attribute', value:'flavor_en', level:'flavor'}
                    // | {kind:'node', level:'manufacturer', match:'LACTALIS'}
  measure,          // 'sales_value' (+ lookup agg_rule, var_family)
  years,            // [2024,2025,2026]
})  →  { sql, needsCoverage:boolean, volumeBlocked:boolean }
```
Le builder applique : régime selon `breakdown.kind`, additivité selon `agg_rule`,
guardrail selon `VOLUME_BASIS`, et déclenche la requête coverage (§4) si
`breakdown.kind==='attribute'`.

## 9. Initialisation DuckDB-WASM

```js
import * as duckdb from '@duckdb/duckdb-wasm';
const bundles = duckdb.getJsDelivrBundles();
const bundle  = await duckdb.selectBundle(bundles);
const worker  = new Worker(bundle.mainWorker);
const db      = new duckdb.AsyncDuckDB(new duckdb.ConsoleLogger(), worker);
await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
const conn = await db.connect();

const buf = new Uint8Array(await (await fetch('category_db.parquet')).arrayBuffer());
await db.registerFileBuffer('category_db.parquet', buf);
await conn.query(`CREATE VIEW db AS SELECT * FROM read_parquet('category_db.parquet')`);
// ensuite : conn.query(sql) → résultats Arrow → .toArray()
```

## 10. Structure de fichiers

```
category-explorer/
  index.html            # UI + DuckDB-WASM + builder (single-file, style SalesIQ)
  category_db.parquet   # sortie Stage 1
  build_category_db.py  # Stage 1
  harmonize_flavors.py  # Stage 2c (proposition LLM-assistée, voir §11)
  serve.sh              # python3 -m http.server 8000
```

## 11. Gaps d'harmonisation connus (à industrialiser)

1. **Parfums composés/rares** non encore mappés (`flavor_en` retombe sur le libellé local
   titre-casé). Process : `harmonize_flavors.py` extrait les `flavor_en` distincts par pays,
   appelle l'API Anthropic pour proposer `local → EN canonique`, **tu valides**, le résultat
   enrichit les `flavor_map` du Stage 1. Même patron réutilisable pour `weight`/`format`.
2. **Fabricants / marques cross-pays NON harmonisés** : un même groupe (ex. Lactalis) peut
   apparaître sous des `manufacturer` différents selon le pays. Pour des rollups groupe
   fiables en Europe, ajouter un `manufacturer_group` map en Stage 1 (même patron que
   `flavor_map`). En v1, le matching fabricant est littéral par pays (`ILIKE`).
3. **Base volume Europe** : pas de base commune (cf. §6) → décision fournisseur, hors code.
```
