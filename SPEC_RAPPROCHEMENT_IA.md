# Spécification — Module de rapprochement commandes ↔ factures avec assistance IA

> **Pour Claude Code.** Ce document spécifie l'ajout d'un module de rapprochement automatique et assisté par IA à l'application `suivi_commandes_factures_marches_FinaàGarder.py`. Il s'inscrit en complément du fonctionnement actuel : **rien ne doit casser**, le rapprochement natif (`factures.code_mouvement = commandes.num_commande`) reste la source primaire et est juste enrichi par une couche de liens manuels/IA.

---

## 1. Contexte métier

L'application sert au suivi des commandes, factures et marchés publics pour la mairie de La Grande-Motte. Les données viennent d'exports Excel du logiciel comptable (CIRIL probable) :
- `commandes_*.xls` : engagements (bons de commande)
- `factures_*.xls` : mouvements comptables (services faits + mandatements)

Le rapprochement actuel se fait par jointure stricte : `factures.code_mouvement == commandes.num_commande`. **Cela fonctionne tant que la comptabilité saisit le n° de bon de commande sur la facture reçue**, mais pose problème dans 4 cas réels rencontrés dans les données 2025 :

### 1.1 Pattern A — Doublons administratifs
La compta engage directement (sans BC) une dépense récurrente (ex. maintenance ascenseurs annuelle ACAF), puis l'utilisateur métier crée par ailleurs un BC pour la même prestation. Le BC ne sera **jamais** facturé car la facture porte un autre `code_mouvement` (engagement direct compta).

**Exemple réel** : commande `25AA01486` ACAF 2376€ "Maintenance annuelle 2025 des ascenseurs" → reste éternellement "Non facturée" alors que les factures `25AA00077/078/079` ACAF couvrent le même périmètre via le marché `2023_17`.

### 1.2 Pattern B — Marchés à bons de commande sans n° BC
Marchés cadres (ex. Bouygues `2020_14G3P`, Dalkia P1/P2/P3) : la compta crée des engagements ad hoc sans associer le BC métier. Les factures portent le bon Marché mais pas le bon `code_mouvement`. Plusieurs factures peuvent cumuler pour couvrir une commande.

**Exemple réel** : commande `25AA00237` Bouygues 36 717€ Marché `2020_14G3P` → 18 factures orphelines du même fournisseur sur le même marché.

### 1.3 Pattern C — Reports d'années antérieures
Les factures 2025 portent souvent un `code_mouvement` 2021–2024 (engagements pluriannuels, REPORTs). Si l'utilisateur n'importe que `commandes_2025.xls`, ces factures restent orphelines à tort.

**Mesure réelle** : 271 factures sur les 1101 du fichier 2025 ont un préfixe `21AA/22AA/23AA/24AA` (~25%).

### 1.4 Pattern D — Vrais oublis fournisseur
La facture n'arrive pas physiquement (oubli du fournisseur, retard postal, litige). Aucune action informatique ne résout ça : il faut **relancer le fournisseur**. L'app doit savoir le **distinguer** des autres patterns.

### 1.5 Mesures sur les données 2025

| Indicateur | Valeur |
|---|---|
| Total commandes 2025 (agrégées) | 214 |
| Commandes "non/partiellement facturées" (statut actuel) | 53 (25 %) |
| Reste à facturer "fantôme" | 508 731 € |
| Factures orphelines (code_mouvement absent de commandes_2025) | 800 |
| ↳ dont vraies à rapprocher (hors fluides EDF/Total/Dalkia abos) | 399 |
| Commandes problématiques avec ≥ 90 jours d'âge (vrais oublis) | 25 |

---

## 2. Vue d'ensemble de la solution

L'architecture s'articule en **4 couches additives**, de la moins intrusive à la plus intelligente :

```
┌─────────────────────────────────────────────────────────────┐
│ COUCHE 4 — UI : Assistant de rapprochement (dialog dédiée)  │
│   Filtres • Liste cmds problématiques • Candidats • IA      │
└─────────────────────────────────────────────────────────────┘
                          ▲
┌─────────────────────────────────────────────────────────────┐
│ COUCHE 3 — IA DeepSeek : analyse sémantique des candidats   │
│   Prompt structuré • JSON strict • Score de confiance       │
└─────────────────────────────────────────────────────────────┘
                          ▲
┌─────────────────────────────────────────────────────────────┐
│ COUCHE 2 — Pré-classification déterministe (sans IA)        │
│   age_jours • candidats_count • diagnostic 1ère intention   │
└─────────────────────────────────────────────────────────────┘
                          ▲
┌─────────────────────────────────────────────────────────────┐
│ COUCHE 1 — Tables SQLite : commande_diagnostic + manual_links│
│   recompute_facturation modifié pour faire l'union          │
└─────────────────────────────────────────────────────────────┘
                          ▲
                  Existant (intact)
```

**Principe directeur** : la jointure native `code_mouvement = num_commande` reste la source de vérité primaire. Les `manual_links` viennent **s'ajouter** au calcul, jamais s'y substituer. On ne modifie **jamais** les données importées.

---

## 3. Périmètre du module `matching_module.py` (nouveau fichier)

### 3.1 Responsabilités

Le module concentre toute la logique de rapprochement et d'IA. Il **n'a pas de dépendance PyQt** — c'est une couche métier pure, testable. Une dialog PyQt séparée (`matching_dialog.py`) consomme ce module.

```
matching_module.py
├── class MatchingEngine
│   ├── diagnose_all_commandes()      # remplit commande_diagnostic
│   ├── find_candidates(commande_id)  # liste factures candidates pré-filtrées
│   ├── deterministic_match(cmd, candidates)  # règles strictes
│   └── ai_match(cmd, candidates)     # appel DeepSeek
├── class DeepSeekClient
│   ├── __init__(api_key, model="deepseek-chat")
│   ├── chat(prompt)  # avec retry/timeout
│   └── parse_json_response(raw)
└── class LinkRepository
    ├── add_link(commande_id, facture_id, montant, source, ...)
    ├── remove_link(link_id)
    ├── list_links_for_commande(commande_id)
    └── total_alloue_for_commande(commande_id)
```

### 3.2 Le module ne fait **PAS** :
- d'UI Qt (laissé à `matching_dialog.py`)
- de modifications sur les tables `commandes` ou `factures` source
- d'appels réseau hors DeepSeek

---

## 4. Schéma SQL — nouvelles tables

À ajouter dans `Database._init_schema()` (avec migration lazy comme le code existant). **Toutes** les `CREATE TABLE` utilisent `IF NOT EXISTS`.

### 4.1 Table `manual_links`

```sql
CREATE TABLE IF NOT EXISTS manual_links (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    commande_id INTEGER NOT NULL,
    facture_id INTEGER NOT NULL,
    montant_alloue REAL NOT NULL,
    confidence INTEGER,                  -- 0..100, NULL si manuel pur
    source TEXT NOT NULL,                -- 'manual' | 'ai_validated' | 'ai_auto' | 'deterministic'
    ai_model TEXT,                       -- ex: 'deepseek-chat'
    ai_reasoning TEXT,                   -- explication IA (max 1000 char)
    created_by TEXT,                     -- 'user' | 'ai'
    created_at TEXT NOT NULL,            -- ISO timestamp
    validated_at TEXT,                   -- date validation manuelle (NULL si auto)
    notes TEXT,
    UNIQUE(commande_id, facture_id),
    FOREIGN KEY (commande_id) REFERENCES commandes(id) ON DELETE CASCADE,
    FOREIGN KEY (facture_id) REFERENCES factures(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_manual_links_cmd ON manual_links(commande_id);
CREATE INDEX IF NOT EXISTS idx_manual_links_fact ON manual_links(facture_id);
```

**Important** : `montant_alloue` permet de répartir une grosse facture entre plusieurs commandes. Contrainte applicative (à vérifier avant insert) : la somme des `montant_alloue` pour une `facture_id` donnée ne doit pas dépasser le `montant_service_fait` de cette facture.

### 4.2 Table `commande_diagnostic`

```sql
CREATE TABLE IF NOT EXISTS commande_diagnostic (
    commande_id INTEGER PRIMARY KEY,
    age_jours INTEGER,
    diagnostic TEXT NOT NULL,            -- voir énumération §4.3
    severite INTEGER,                    -- 0=info, 1=normal, 2=alerte, 3=critique
    candidates_count INTEGER DEFAULT 0,  -- nb factures orphelines même fournisseur
    candidates_same_marche INTEGER DEFAULT 0,
    montant_candidates_total REAL DEFAULT 0,
    last_ai_check_at TEXT,               -- timestamp dernier appel IA (pour cache)
    last_ai_diagnostic TEXT,             -- copie du JSON IA (cache)
    last_diagnostic_at TEXT NOT NULL,
    FOREIGN KEY (commande_id) REFERENCES commandes(id) ON DELETE CASCADE
);
```

### 4.3 Énumération `diagnostic`

| Valeur | Sévérité | Signification | Action UI |
|---|---|---|---|
| `OK` | 0 | Totalement facturée via jointure native | Aucune |
| `OK_RAPPROCHE` | 0 | Totalement facturée grâce à un manual_link | Badge 🔗 |
| `RECENT` | 1 | < 30 jours, normal qu'elle ne soit pas facturée | Aucune |
| `EN_COURS` | 1 | 30-90 jours, facturation en cours probable | Aucune |
| `RAPPROCHEMENT_SUGGERE` | 2 | > 90j, candidats détectés | Badge 🔍 + apparait en priorité dans l'assistant |
| `OUBLI_PROBABLE` | 3 | > 90j, **aucun** candidat → relancer fournisseur | Badge 🚨 |
| `DOUBLON_PROBABLE` | 2 | IA a détecté un doublon administratif | Badge ♻️ + statut spécifique |

### 4.4 Migration des tables existantes

Dans `commandes` : ajouter une colonne pour le nouveau statut doublon :

```sql
-- migration lazy comme le code existant
ALTER TABLE commandes ADD COLUMN statut_metier TEXT;
-- valeurs possibles: NULL (défaut), 'DOUBLON_ADMIN', 'ANNULEE'
```

Une commande avec `statut_metier = 'DOUBLON_ADMIN'` sort du calcul de "reste à facturer" (cf. §5).

---

## 5. Modification de `recompute_facturation`

C'est **la** modification critique. La méthode actuelle (Database.recompute_facturation, lignes ~1166-1225) calcule le total facturé via la jointure native. Il faut **étendre** ce calcul pour additionner aussi les `manual_links`.

### 5.1 Pseudo-code de la nouvelle logique

```python
def recompute_facturation(self):
    cur = self.conn.cursor()

    # 1. Total natif par num_commande (logique actuelle, inchangée)
    cur.execute("""
        SELECT code_mouvement, SUM(montant_service_fait) AS total
        FROM factures
        WHERE code_mouvement IS NOT NULL AND code_mouvement != ''
        GROUP BY code_mouvement
    """)
    totals_natif = {r["code_mouvement"]: (r["total"] or 0.0) for r in cur.fetchall()}

    # 2. NOUVEAU : Total via manual_links par commande_id
    cur.execute("""
        SELECT commande_id, SUM(montant_alloue) AS total
        FROM manual_links
        GROUP BY commande_id
    """)
    totals_manual = {r["commande_id"]: (r["total"] or 0.0) for r in cur.fetchall()}

    # 3. Boucle commandes
    cur.execute("SELECT id, num_commande, montant_ttc, statut, statut_metier FROM commandes")
    for row in cur.fetchall():
        cmd_id = row["id"]
        num = row["num_commande"]
        mt_cmd = row["montant_ttc"] or 0.0

        # NOUVEAU : si la commande est marquée DOUBLON_ADMIN, on la sort du calcul
        if row["statut_metier"] == "DOUBLON_ADMIN":
            cur.execute("""
                UPDATE commandes
                SET montant_facture = 0, reste_a_facturer = 0,
                    statut_facturation = 'Doublon administratif',
                    statut = 'Annulée'
                WHERE id = ?
            """, (cmd_id,))
            continue

        # Total = natif + manual_links
        mt_fact = totals_natif.get(num, 0.0) + totals_manual.get(cmd_id, 0.0)
        reste = mt_cmd - mt_fact

        # Logique statut_facturation : INCHANGÉE
        if mt_cmd <= 0:
            statut_fact = "Non facturée" if mt_fact <= 0 else "Totalement facturée"
        else:
            if mt_fact <= 0:
                statut_fact = "Non facturée"
            elif mt_fact + 1e-6 < mt_cmd:
                statut_fact = "Partiellement facturée"
            else:
                statut_fact = "Totalement facturée"

        # Logique statut commande : INCHANGÉE
        statut_cmd = row["statut"]
        new_statut = "Envoyée" if statut_fact in ("Partiellement facturée", "Totalement facturée") else (statut_cmd or "A suivre")

        cur.execute("""
            UPDATE commandes
            SET montant_facture = ?, reste_a_facturer = ?, statut_facturation = ?, statut = ?
            WHERE id = ?
        """, (mt_fact, reste, statut_fact, new_statut, cmd_id))

    self.conn.commit()
```

### 5.2 Garde-fous critiques

- **Ne jamais** faire le calcul en double : si `code_mouvement = num_commande` (jointure native fonctionne), on ne crée **pas** de manual_link. Dans le code de matching, vérifier en amont qu'on ne propose que des factures dont `code_mouvement != commande.num_commande`.
- L'ajout d'un `manual_link` doit **toujours** déclencher un `recompute_facturation()` derrière (ou au moins recalculer la commande concernée).
- Suppression d'un manual_link : idem, recalcul derrière.

---

## 6. Logique de pré-classification déterministe

La méthode `MatchingEngine.diagnose_all_commandes()` parcourt toutes les commandes et remplit `commande_diagnostic`.

### 6.1 Algorithme

```python
for cmd in toutes_les_commandes:
    # Date d'extraction = date la plus récente parmi factures.last_update et commandes.last_update
    # OU datetime.now() si plus simple
    age_jours = (today - cmd.date_commande).days

    # Calcul du statut effectif (incluant manual_links)
    if cmd.statut_metier == 'DOUBLON_ADMIN':
        diagnostic, severite = 'DOUBLON_PROBABLE', 2
    elif cmd.statut_facturation == 'Totalement facturée':
        # vérifier si c'est via manual_link ou natif
        has_manual = LinkRepository.list_links_for_commande(cmd.id)
        diagnostic = 'OK_RAPPROCHE' if has_manual else 'OK'
        severite = 0
    elif age_jours < 30:
        diagnostic, severite = 'RECENT', 1
    elif age_jours < 90:
        diagnostic, severite = 'EN_COURS', 1
    else:
        # > 90 jours et non totalement facturée : on cherche des candidats
        candidates = find_candidates(cmd.id)
        if len(candidates) == 0:
            diagnostic, severite = 'OUBLI_PROBABLE', 3
        else:
            diagnostic, severite = 'RAPPROCHEMENT_SUGGERE', 2

    upsert commande_diagnostic
```

### 6.2 `find_candidates(commande_id)` — règles de filtrage

Une facture est candidate pour une commande si **toutes** les conditions :

1. La facture n'est **pas** déjà liée à cette commande (ni via `code_mouvement = num_commande`, ni via un `manual_link` existant).
2. La facture a un `montant_service_fait > 0` (on ignore les engagements à 0).
3. **Fournisseur** : `factures.fournisseur` correspond à `commandes.fournisseur` à ≥ 85 % (utiliser `difflib.SequenceMatcher.ratio()` après normalisation upper/strip). La normalisation doit gérer les variations type "BOUYGUES ENERGIES ET SERVICES" vs "BOUYGUES E&S".
4. **Date** : `factures.date_facture >= commandes.date_commande - 30 jours` (tolérance pour les factures saisies en avance).
5. **Marché** (si renseigné des deux côtés) : `factures.marche == commandes.marche` strictement. Si une seule des deux est renseignée, la condition est ignorée (compatible).

### 6.3 Performance

`diagnose_all_commandes()` doit s'exécuter en < 5 secondes sur 1000 commandes / 5000 factures. Optimisations :
- Index SQL sur `factures.fournisseur` et `commandes.fournisseur`.
- Pré-filtrage SQL sur le fournisseur (LIKE ou égalité stricte) avant le calcul de similarité Python.
- Calculer `montant_alloue_total` par facture une seule fois (pas dans la boucle).

---

## 7. Intégration DeepSeek

### 7.1 Configuration (table `config`)

Nouvelles clés dans la table `config` (mécanisme `get_config`/`set_config` existant) :

| Clé | Type | Défaut | Description |
|---|---|---|---|
| `deepseek_api_key` | str | `""` | Clé API (chiffrée idéalement, mais simple stockage texte acceptable v1) |
| `deepseek_model` | str | `"deepseek-chat"` | Modèle (laisser configurable même si on cible chat) |
| `deepseek_endpoint` | str | `"https://api.deepseek.com/v1/chat/completions"` | URL API |
| `ai_auto_threshold` | int | `90` | Seuil d'auto-validation (≥ ce score → rapprochement automatique) |
| `ai_min_threshold` | int | `40` | Seuil minimum d'affichage des suggestions |
| `ai_enabled` | int | `0` | Active/désactive l'IA globalement (0/1) |

L'utilisateur configure tout ça dans `ConfigDialog` (onglet existant ou nouvel onglet "IA").

### 7.2 Classe `DeepSeekClient`

```python
class DeepSeekClient:
    def __init__(self, api_key, model="deepseek-chat", endpoint=None, timeout=30):
        ...

    def chat(self, system_prompt, user_prompt, max_tokens=800, temperature=0.1):
        """Appel synchrone à l'API. Retourne le texte brut de la réponse.
        Raise sur erreur réseau ou HTTP non-200.
        Implémente un retry simple (3 tentatives, backoff exponentiel).
        """

    @staticmethod
    def parse_json_response(raw_text):
        """Extrait le JSON d'une réponse. Tolère les ```json fences,
        les préambules type 'Voici la réponse: ...', etc.
        Raise ValueError si parsing impossible.
        """
```

**Bibliothèque** : utiliser `requests` (déjà probablement présent) ou `urllib.request` (stdlib pure pour minimiser les dépendances).

### 7.3 Prompt unique et stable

```python
SYSTEM_PROMPT = """Tu es un expert en comptabilité publique française (M14/M57).
Tu analyses si une commande peut être rapprochée de factures candidates.
Tu réponds TOUJOURS en JSON strict, sans préambule ni texte hors JSON."""

USER_PROMPT_TEMPLATE = """COMMANDE À RAPPROCHER:
- N°: {num_commande}
- Date: {date_commande}
- Fournisseur: {fournisseur}
- Marché: {marche}
- Service émetteur: {service_emetteur}
- Montant TTC: {montant_ttc}€
- Reste à facturer: {reste}€
- Libellé: {libelle}
- Désignation: {designation}

FACTURES CANDIDATES (factures dont le code_mouvement n'est pas le n° de cette commande,
mais qui pourraient correspondre par fournisseur/marché/sémantique):

{factures_json}

CONSIGNES:
1. Évalue si UNE OU PLUSIEURS factures peuvent ensemble couvrir la commande
2. Méfie-toi des DOUBLONS ADMINISTRATIFS : si la commande ressemble à un engagement déjà
   couvert par un autre code_mouvement (ex. maintenance annuelle facturée mensuellement par
   la compta sans BC), signale-le explicitement
3. Si marché renseigné des deux côtés, ils doivent matcher
4. Date de facture postérieure (ou simultanée à 30j près) à la commande
5. Le montant alloué à chaque facture peut être PARTIEL (ne pas dépasser son montant_sf)
6. Plusieurs factures peuvent cumuler pour atteindre le reste à facturer

Réponds en JSON strict (pas de markdown, pas de fences):
{{
  "diagnostic": "MATCH" | "PARTIEL" | "DOUBLON" | "ORPHELINE" | "INDETERMINE",
  "confidence": <int 0-100>,
  "factures_a_lier": [
    {{
      "code_mouvement": "<str>",
      "montant_alloue": <float>,
      "raison": "<str max 100 char>"
    }}
  ],
  "raisonnement": "<str max 500 char, explication globale>",
  "action_suggeree": "VALIDER_AUTO" | "VALIDER_MANUEL" | "MARQUER_DOUBLON" | "RELANCER_FOURNISSEUR" | "INVESTIGUER"
}}"""
```

### 7.4 Politique de décision après réponse IA

```python
def apply_ai_decision(commande_id, ai_response, auto_threshold, min_threshold):
    confidence = ai_response['confidence']
    diagnostic = ai_response['diagnostic']
    action = ai_response['action_suggeree']

    if diagnostic == 'DOUBLON' and action == 'MARQUER_DOUBLON' and confidence >= auto_threshold:
        # Marquer la commande comme doublon administratif
        db.set_commande_statut_metier(commande_id, 'DOUBLON_ADMIN')
        log("AI auto-marked as duplicate")
        return 'AUTO_DOUBLON'

    if diagnostic in ('MATCH', 'PARTIEL') and confidence >= auto_threshold and action == 'VALIDER_AUTO':
        # Créer les manual_links automatiquement
        for f in ai_response['factures_a_lier']:
            facture_id = db.get_facture_id_by_code_mouvement(f['code_mouvement'])
            link_repo.add_link(
                commande_id=commande_id,
                facture_id=facture_id,
                montant_alloue=f['montant_alloue'],
                confidence=confidence,
                source='ai_auto',
                ai_reasoning=ai_response['raisonnement']
            )
        db.recompute_facturation()
        return 'AUTO_LINKED'

    if confidence >= min_threshold:
        # Sauvegarder la suggestion en cache (commande_diagnostic.last_ai_diagnostic)
        # mais ne rien faire d'automatique
        return 'SUGGESTION'

    return 'IGNORED'
```

### 7.5 Cache des appels IA

L'API DeepSeek a un coût (faible mais non nul) et une latence (1-5s). Pour éviter de réinterroger inutilement :

- Stocker la réponse JSON dans `commande_diagnostic.last_ai_diagnostic`
- Ne pas réinterroger tant que :
  - le `last_ai_check_at` est < 7 jours
  - ET la liste des factures candidates n'a pas changé (calculer un hash trivial des `code_mouvement` candidats triés)

### 7.6 Coût estimé

Sur les volumes mesurés : 25 commandes "vraiment problématiques" × ~1500 tokens prompt + 300 tokens réponse = 45 K tokens. Avec deepseek-chat à environ 0,14 $ / 1M tokens en entrée et 0,28 $ / 1M en sortie, **coût total < 0,02 $ pour traiter toute la base 2025**. Vérifier la grille tarifaire actuelle au moment de l'implémentation.

---

## 8. UI — `matching_dialog.py` (nouveau fichier)

### 8.1 Bouton d'accès

Dans la toolbar principale (`_init_toolbar`), ajouter un bouton **après** le bouton "Configuration" :

```python
act_matching = QAction("🔗 Rapprochement", self)
act_matching.setToolTip("Assistant de rapprochement commandes ↔ factures (avec IA)")
act_matching.triggered.connect(self.open_matching_assistant)
tb1.addAction(act_matching)
# Style violet par exemple
```

### 8.2 Méthode `MainWindow.open_matching_assistant`

```python
def open_matching_assistant(self):
    from matching_dialog import MatchingDialog
    dlg = MatchingDialog(self.db, self)
    dlg.exec_()
    # Au retour : refresh des modèles
    self.cmd_model.refresh()
    self.synth_model.refresh()
    self.refresh_rappels_tab()
```

### 8.3 Layout de la dialog

```
┌─ Assistant de rapprochement ─────────────────────────────────┐
│                                                              │
│ ┌─ Filtres ──────────────────────────────────────────────┐   │
│ │ Diagnostic: [Tous ▼]  Fournisseur: [Tous ▼]  > X jours │   │
│ │ [☑ Oublis probables] [☑ Suggestions] [☐ OK rapprochés] │   │
│ └────────────────────────────────────────────────────────┘   │
│                                                              │
│ ┌─ Commandes (gauche) ─────────┐ ┌─ Détail (droite) ─────┐   │
│ │ Diag │ N° Cmd │ Fourn  │ €   │ │ Commande sélectionnée  │   │
│ │ 🚨   │ 25AA01 │ ACAF   │2376 │ │ [details]              │   │
│ │ 🔍   │ 25AA00 │BOUYGUES│18K  │ │                        │   │
│ │ ♻️   │ 25AA02 │ DALKIA │479  │ │ Factures candidates    │   │
│ │ ...                          │ │ ☐ 25AA00077 ...        │   │
│ │                              │ │ ☐ 25AA00078 ...        │   │
│ │                              │ │                        │   │
│ │                              │ │ [🤖 Demander avis IA]  │   │
│ │                              │ │                        │   │
│ │                              │ │ Réponse IA:            │   │
│ │                              │ │ [zone texte readonly]  │   │
│ │                              │ │                        │   │
│ │                              │ │ [✓ Valider]            │   │
│ │                              │ │ [♻ Marquer doublon]    │   │
│ │                              │ │ [✗ Pas de match]       │   │
│ └──────────────────────────────┘ └────────────────────────┘   │
│                                                              │
│ ┌─ Actions de masse ─────────────────────────────────────┐   │
│ │ [🤖 Analyse IA de toutes les "Suggestions"]            │   │
│ │   → Va appeler l'API pour les N commandes RAPPROCHEMENT│   │
│ │     _SUGGERE et créer les liens auto au-dessus du seuil│   │
│ └────────────────────────────────────────────────────────┘   │
│                                                              │
│ [Fermer]                                                     │
└──────────────────────────────────────────────────────────────┘
```

### 8.4 Comportements attendus

- **Tri par défaut** : sévérité décroissante (critique en haut), puis montant décroissant.
- **Sélection** d'une commande à gauche → met à jour le panneau droit avec ses candidats.
- **Bouton IA** : lance un `QThread` (l'appel API ne doit pas bloquer l'UI). Pendant l'attente, afficher un spinner. Au retour, parser la réponse, pré-cocher les factures que l'IA recommande, afficher le raisonnement.
- **Validation** : crée les `manual_links`, déclenche `recompute_facturation`, met à jour le diagnostic, retire la commande de la liste si elle devient `OK_RAPPROCHE`.
- **"Pas de match"** : enregistre dans `commande_diagnostic.notes` que l'utilisateur a explicitement écarté → ne plus proposer cette commande pendant N jours (configurable).
- **Action de masse** : confirmation avant de lancer (afficher : "L'API va être appelée X fois, coût estimé Y centimes, continuer ?").

### 8.5 Indicateurs visuels

Dans l'onglet **Commandes** existant, ajouter une colonne `Diag` (à gauche du statut) qui affiche l'emoji du diagnostic. Permet de voir d'un coup d'œil les commandes problématiques sans ouvrir l'assistant.

---

## 9. Tests à prévoir

Créer un fichier `test_matching.py` à côté du module avec au minimum :

### 9.1 Tests unitaires `MatchingEngine`
- `test_diagnose_recent_commande` : commande de 5 jours → `RECENT`
- `test_diagnose_oubli_probable` : commande > 90j sans candidat → `OUBLI_PROBABLE`
- `test_diagnose_rapprochement_suggere` : commande > 90j avec candidat → `RAPPROCHEMENT_SUGGERE`
- `test_find_candidates_filter_by_marche` : marché renseigné des deux côtés mais différents → 0 candidat
- `test_find_candidates_fuzzy_fournisseur` : "BOUYGUES E&S" vs "BOUYGUES ENERGIES ET SERVICES" → match

### 9.2 Tests `LinkRepository`
- `test_add_link_normal` : ajout simple
- `test_add_link_unicity` : doublon (commande_id, facture_id) → exception
- `test_montant_alloue_exceeds_facture` : alloué > montant_sf → exception
- `test_remove_link_triggers_recompute` : suppression recalcule bien le statut

### 9.3 Tests `recompute_facturation` modifié
- `test_recompute_with_manual_link` : 1 commande non rapprochée nativement + 1 manual_link couvrant 100% → statut "Totalement facturée"
- `test_recompute_partial_with_manual` : natif 50% + manual 30% → "Partiellement facturée" 80%
- `test_recompute_doublon_admin` : commande avec `statut_metier=DOUBLON_ADMIN` → exclue du reste à facturer

### 9.4 Tests `DeepSeekClient`
- Mock HTTP : tester le parsing de réponses bien formées
- Tester la robustesse : réponse avec ```json fences, préambule, JSON malformé
- Tester le retry sur erreur 5xx

### 9.5 Test d'intégration
Créer un mini-jeu de données reproduisant les 4 patterns (A/B/C/D) et vérifier que chacun produit le bon diagnostic.

---

## 10. Ordre d'implémentation suggéré

Pour livrer en étapes incrémentales testables :

1. **Étape 1 — Fondations SQL** (1-2h)
   - Ajouter les CREATE TABLE dans `Database._init_schema()`
   - Ajouter migration `ALTER TABLE commandes ADD COLUMN statut_metier`
   - Modifier `recompute_facturation` pour intégrer `manual_links` et `statut_metier=DOUBLON_ADMIN`
   - **Test** : insérer manuellement un manual_link dans la base, lancer l'app, vérifier que le statut bascule

2. **Étape 2 — Module matching, partie déterministe** (2-3h)
   - Créer `matching_module.py` avec `MatchingEngine.diagnose_all_commandes` et `find_candidates`
   - Créer `LinkRepository`
   - **Test** : exécuter `diagnose_all_commandes` sur la base réelle, inspecter `commande_diagnostic` à la main

3. **Étape 3 — UI minimale sans IA** (3-4h)
   - Créer `matching_dialog.py` avec layout, listes filtrées, ajout/suppression de liens manuels (sans IA)
   - Bouton dans la toolbar
   - **Test** : faire un rapprochement manuel sur le cas réel `25AA01486 ACAF` ↔ `25AA00079`, vérifier la mise à jour du statut

4. **Étape 4 — Intégration DeepSeek** (3-4h)
   - Créer `DeepSeekClient`
   - Ajouter les clés de config dans `ConfigDialog`
   - Brancher le bouton "🤖 Demander avis IA" sur la dialog
   - **Test** : appeler l'IA sur `25AA01486` (cas doublon ACAF), vérifier que le diagnostic est `DOUBLON` et que le marquage automatique fonctionne

5. **Étape 5 — Action de masse + cache** (2-3h)
   - Implémenter le bouton "Analyse IA de toutes les Suggestions"
   - Cache `last_ai_diagnostic` pour éviter les appels redondants
   - **Test** : lancer sur la base entière, vérifier le coût total et la cohérence

6. **Étape 6 — Polish UI** (1-2h)
   - Colonne Diag dans l'onglet Commandes
   - Badges visuels
   - Filtres avancés dans la dialog

---

## 11. Points d'attention spécifiques

### 11.1 Compatibilité Windows (préférence Fred)
- Pas d'Unicode dans les `print` de console : remplacer les emojis par `[OK]`, `[ERR]`, etc. dans les logs
- Encoding `utf-8` explicite à l'ouverture de tout fichier
- Chemins via `os.path.join` ou `pathlib.Path`
- Tester le module entier sur Windows avant de considérer fini

### 11.2 Pas de régression sur l'existant
- **Tester** avant/après que l'onglet "Suivi marchés" continue de fonctionner (le module `marches_module` lit `factures` directement, donc rien ne change pour lui)
- **Tester** que `import_excel_files_incremental` fonctionne toujours (le hash MD5 ne doit pas être perturbé)
- **Tester** que `export_to_pdf` et `export_to_excel` produisent le même résultat sur les commandes non touchées

### 11.3 Sécurité de la clé API
La clé DeepSeek est sensible. Pour v1 :
- Stockage dans `config` table (texte clair, accepté car la base est locale)
- **Ne jamais** la logger
- **Ne jamais** l'inclure dans les exports/PDF

### 11.4 Mode dégradé sans réseau
Si l'API DeepSeek est inaccessible :
- L'app ne doit **pas** crasher
- L'assistant de rapprochement reste utilisable en mode purement manuel
- Message d'erreur clair dans l'UI : "API DeepSeek inaccessible, mode manuel uniquement"

### 11.5 Logs métier
Toute action de rapprochement (manuelle, IA, auto) doit logger dans `manual_links` avec timestamp et source. Cela constitue une **piste d'audit** essentielle en comptabilité publique. Pas de suppression sans log : si on retire un lien, garder une trace dans une table `manual_links_history` (option si bandwidth, sinon notes).

---

## 12. Diagnostics & métriques à exposer

Une nouvelle action toolbar **🩺 Diagnostic santé** ouvrant une dialog simple :

```
État de la base :
  • 1101 factures importées (4,01 M€)
  • 214 commandes (987 K€)
  • Liens manuels : 0
  • Liens IA validés : 0
  • Liens IA auto : 0

Diagnostic des commandes :
  • OK : 161 (75%)
  • Récentes (<30j) : 18
  • En cours (30-90j) : 10
  • 🔍 Suggestions de rapprochement : 20
  • 🚨 Oublis probables : 22 (60 358 €)
  • ♻️ Doublons admin : 0
  • 🔗 OK via lien manuel : 0

[📤 Exporter rapport CSV]  [Fermer]
```

Permet à Fred de voir d'un coup d'œil l'état du système et de mesurer l'impact de l'IA dans le temps.

---

## 13. Récapitulatif des livrables

| Fichier | Type | Estimé (LoC) |
|---|---|---|
| `matching_module.py` | NOUVEAU | ~600 |
| `matching_dialog.py` | NOUVEAU | ~500 |
| `test_matching.py` | NOUVEAU | ~300 |
| `suivi_commandes_factures_marches_FinaàGarder.py` | MODIFIÉ | ~150 lignes ajoutées/modifiées |

**Total estimé : ~1500 lignes nouvelles/modifiées**, sur ~10-15h de travail Claude Code.

---

## 14. Validation finale (checklist avant merge)

- [ ] L'app démarre sans erreur sur une base existante (migration lazy fonctionne)
- [ ] L'app démarre sans erreur sur une base vierge
- [ ] L'import incrémental fonctionne toujours
- [ ] L'onglet Suivi marchés fonctionne toujours
- [ ] Les rappels emails fonctionnent toujours
- [ ] Les exports PDF/Excel fonctionnent toujours
- [ ] Sans clé API configurée, l'assistant de rapprochement fonctionne en mode manuel
- [ ] Avec clé API, le bouton IA fonctionne et crée les liens
- [ ] Un rapprochement manuel fait basculer le statut de "Non facturée" à "Partiellement/Totalement facturée"
- [ ] Un marquage doublon fait basculer la commande en "Annulée"
- [ ] Le diagnostic santé affiche des chiffres cohérents
- [ ] Les 53 commandes problématiques de la base de test 2025 sont correctement classifiées (au moins 20 en RAPPROCHEMENT_SUGGERE et 25+ en OUBLI_PROBABLE)
- [ ] Tests unitaires verts (au moins 80% de couverture du module matching)

---

## Annexe A — Exemple de session IA réelle

**Input** (commande 25AA01486 ACAF + 3 factures candidates) :

```json
{
  "commande": {
    "num": "25AA01486",
    "date": "2025-04-28",
    "fournisseur": "ACAF ASCENSEURS",
    "marche": "2023_17",
    "service_emetteur": "DSTBATI",
    "montant_ttc": 2376.00,
    "reste": 2376.00,
    "libelle": "Maintenance annuelle 2025 des ascenseurs.",
    "designation": "Prix n° 1 - Palais des sports."
  },
  "factures_candidates": [
    {"code_mouvement": "25AA00077", "libelle": "MAINTENANCE MONTE CHARGE BIBLIO 2025", "marche": "2023_17", "montant_sf": 373.37, "date_sf": "2025-01-13"},
    {"code_mouvement": "25AA00078", "libelle": "MAINTENANCE ASCENSEURS MAIRIE 2025", "marche": "2023_17", "montant_sf": 1393.88, "date_sf": "2025-01-13"},
    {"code_mouvement": "25AA00079", "libelle": "MAINTENANCE ASCENSEUR PALAIS SPORTS 2025", "marche": "2023_17", "montant_sf": 696.94, "date_sf": "2025-01-13"}
  ]
}
```

**Output IA attendu** :

```json
{
  "diagnostic": "DOUBLON",
  "confidence": 92,
  "factures_a_lier": [],
  "raisonnement": "La commande concerne explicitement le 'Palais des sports' (cf. désignation). La facture 25AA00079 'MAINTENANCE ASCENSEUR PALAIS SPORTS 2025' couvre déjà cet équipement via un engagement direct compta antérieur (janvier 2025) sur le même marché 2023_17. La commande d'avril semble être un doublon administratif qui ne sera jamais facturé.",
  "action_suggeree": "MARQUER_DOUBLON"
}
```

**Comportement attendu de l'app** :
- `confidence (92) >= auto_threshold (90)` ET `action == MARQUER_DOUBLON`
- → Marquer automatiquement `commandes.statut_metier = 'DOUBLON_ADMIN'`
- → `recompute_facturation` met `statut_facturation = 'Doublon administratif'`, `statut = 'Annulée'`, `reste_a_facturer = 0`
- → Logger l'action dans une table d'audit ou dans `notes`

---

**Fin du document de spécification.**
