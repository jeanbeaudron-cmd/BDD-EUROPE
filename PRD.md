# PRD — Category Review Explorer (Yaourts Europe)

> Interface d'exploration multi-pays des données Nielsen, alimentée par la base
> canonique produite par `build_category_db.py`. Single-page, sans backend,
> en anglais, lancée depuis un serveur statique local.

---

## 1. Objectif

Permettre de naviguer dans une base de revue de catégorie agrégée **5 pays × 3 marchés
(YALG / Kefir / Bifidus) × 2 circuits (Total / Hyper)**, en :

- sommant librement les pays (détail pays **+ total Europe**) ;
- sommant des critères harmonisés (ex. tous les parfums « Plain », tous les formats) ;
- extrayant des tableaux et graphes **configurables**, avec **l'année comme dimension**
  (sélectionner 2024/2025/2026 d'un côté, la variable de l'autre — pas une colonne
  figée par an comme dans les fichiers source) ;
- signalant automatiquement quand une somme est **hors MDD / partielle**.

L'interface ne remplace pas l'analyse : elle fiabilise les chiffres et supprime le
retraitement manuel des 5 fichiers hétérogènes.

## 2. Utilisateur & usage

Utilisateur unique avancé (toi). Deux modes d'usage :

1. **Lecture rapide** : quelques dashboards pré-faits pour la photo d'un marché.
2. **Extraction à la demande** : un builder pour sortir un tableau/graphe précis
   (ex. *« évolution volume du groupe Lactalis sur YALG, 3 ans, Europe + détail pays »*).

## 3. Périmètre données

Source = `category_db.parquet` (format long, voir TECHNICAL.md §2). Couvre :

| Dimension | Valeurs |
|---|---|
| `country` | FR, UK, DE, ES, IT (+ **Europe** = agrégat calculé) |
| `channel` | Total, Hyper |
| `market` | YALG, Kefir, Bifidus |
| `level` | family, category, manufacturer, brand, segment, multipack, weight, bio, flavor, ean |
| attributs | `manufacturer`, `brand`, `segment`, `flavor_en`, `weight`, `multipack`, `bio`, `ean` |
| `variable` | 26 mesures canoniques (value, volume, units, shares, prices, DN/DV, ROS) |
| `year` | 2024, 2025, 2026 |

## 4. Règles métier (NON négociables)

Ces trois règles sont issues de l'analyse des 5 fichiers réels. Toute la logique
d'agrégation en découle.

### 4.1 Deux régimes d'agrégation
- **Nœud nommé** (family, category, manufacturer, brand, segment) → on **lit la ligne
  pré-calculée** de ce niveau. **Jamais** sommer ses enfants. C'est MDD-safe.
- **Cross-cut d'attribut** (flavor, format=weight/multipack, bio) → on **somme les lignes
  au niveau de l'attribut**. Le résultat est partiel (voir 4.2).

### 4.2 Couverture & MDD, calculées PAR PAYS
La part du marché couverte par un cross-cut varie énormément selon le pays. Couverture EAN
constatée (illustratif) :

| Pays | MDD détaillée ? | Couverture EAN (YALG) |
|---|---|---|
| ES | non | 26 % |
| FR | non | 39 % |
| UK | oui | 43 % |
| DE | oui | 100 % |
| IT | oui | 100 % |

→ Tout cross-cut affiche un **badge de couverture par pays** = `somme cross-cut (à son
propre niveau) / total marché`. Le **total marché de référence = somme du niveau
fabricant** (pavage prouvé à 100 % sur les 5 pays ; plus fiable que le niveau catégorie,
dont les libellés sont dupliqués sur certains pays). La couverture est **propre à
l'attribut** : en UK le parfum couvre 100 % alors que l'EAN n'est qu'à 43 %.
En dessous de 100 %, mention *« partial — excludes MDD / undetailed »*.
Ne JAMAIS afficher un seul taux global agrégé : FR/ES excluent la MDD, UK/DE/IT l'incluent
→ un même « Plain » Europe mélange deux traitements. Le warning l'explicite.

### 4.3 Additivité par variable
Chaque variable porte une règle (`agg_rule`) :
- **additive** (`sales_value` en EUR, `sales_units`, volume) → sommable.
- **derived** (parts, prix moyens) → **jamais sommé** ; recalculé depuis les additives.
- **node_only** (DN, DV, ROS, toutes les évolutions/écarts) → n'a de sens qu'à un nœud ;
  affiché uniquement quand un nœud unique est sélectionné, jamais agrégé.

### 4.4 Total Europe — limites
- Sommables en Europe : **valeur (EUR)** et **unités** uniquement.
- **Volume** : 4 bases incompatibles (EQ / KGS / GESAMT / ALL). En vue Europe, le volume
  est **grisé** ; autorisé par pays, et en sous-agrégat **UK+ES** (base KGS commune).
- Parts/prix Europe : **recalculés** depuis les additives sommées (jamais moyennés).

## 5. Fonctionnalités

### 5.1 Dashboards pré-faits (lecture rapide)
- **Market snapshot** : taille (valeur), croissance YoY, PDM MDD, split circuit, top 10
  fabricants. Sélecteurs : pays (ou Europe), marché, année.
- **Manufacturer landscape** : classement fabricants par PDM valeur, évolution 3 ans.
- **Flavor & format mix** : répartition parfums / formats **avec badge couverture**.
- **Distribution & rotation** (node_only) : DN/DV/ROS pour un nœud sélectionné.

### 5.2 Builder (extraction à la demande)
Configuration en 4 blocs :
1. **Scope** : pays (multi-sélection + option Europe), marché(s), circuit.
2. **Découpage** : un niveau hiérarchique (manufacturer/brand/segment) **ou** un attribut
   (flavor/format/bio) **ou** un nœud nommé précis (ex. manufacturer = "LACTALIS").
3. **Mesure** : une variable (filtrée selon additivité + scope ; volume grisé si Europe).
4. **Années** : sélection 2024/2025/2026.

Sorties : **tableau** (export CSV) + **graphe** configurable. Axes : `year` en abscisse
(série temporelle) **ou** le découpage en abscisse, l'autre devient la série. C'est le
besoin de départ — l'année est une dimension, pas trois colonnes.

### 5.3 Composants transverses
- **Coverage badge** (4.2) attaché à tout résultat de cross-cut.
- **Volume guardrail** (4.4) : désactive/avertit selon scope.
- **Harmonisation parfums** : l'interface interroge `flavor_en` (déjà harmonisé en Stage 1).
  Les parfums composés/rares non encore mappés (ex. « Frutti di Bosco », « Mela Cannella »,
  « Macedonia ») passent par un **process LLM-assisté hors interface** : un script propose
  des mappings local→EN, tu valides, le résultat enrichit le `flavor_map` du Stage 1.

## 6. Langue & format
- Interface intégralement en **anglais**.
- Devise d'affichage : **EUR** (valeur canonique cross-pays).
- Pas de localisation multi-langue : la traduction est faite en amont (Stage 1).

## 7. Hors périmètre (v1)
- Édition / écriture de données (lecture seule).
- Réconciliation d'une base volume Europe commune (à arbitrer avec Nielsen).
- Détail EAN pour FR/ES/UK (non fourni à 100 % par la source ; le badge le signale).

## 8. Phases
1. **Stage 1 — fait** : `build_category_db.py` → `category_db.parquet`.
2. **Stage 2a** : moteur DuckDB-WASM + builder + coverage badge (cœur).
3. **Stage 2b** : dashboards pré-faits.
4. **Stage 2c** : process d'harmonisation parfums LLM-assistée.
