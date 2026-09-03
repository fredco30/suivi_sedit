# suivi_march-
Logiciel de gestion pour suivi des marchés

## Suivi financier par opération

L'export « suivi financier » émet **une ligne par état de bon de commande**,
et non une ligne par écriture SEDIT :

- une ligne par facture mandatée (statut `FACTURÉ`) ;
- une seule ligne d'engagement, et uniquement s'il reste à facturer
  (statut `ENGAGEMENT — reliquat non facturé`), portant le reliquat et non le
  montant initial du BDC.

Pour chaque BDC : `montant_ref = max(montant_bdc_déclaré, Σ factures mandatées)`
et `reliquat = montant_ref - Σ factures`. Un BDC soldé n'a donc pas de ligne
d'engagement, et un report d'exercice ne crée pas de second engagement.

Le classeur produit quatre onglets : `FINANCIER`, `A jour` (une ligne par BDC,
issue du même jeu de données), `Anomalies` (montants de BDC à arbitrer) et
`Lignes neutralisées` (trace des engagements supprimés ou ramenés au reliquat).

La règle d'agrégation est isolée dans `suivi_financier_agg.py`, indépendamment
de la génération Excel.

### Colonnes SEDIT utilisées

| Colonne | Rôle |
|---|---|
| E — `Code mouvement` | N° d'engagement. Renseigné sur **toutes** les lignes ; sert de repli quand `Commande` est vide (près de 40 % des lignes). |
| O — `Montant initial` | Montant **total** du BDC, répété sur chacune de ses lignes. Source du `montant_ref`. |
| AD — `Montant TTC` | Montant de la ligne : facture mandatée, ou part encore engagée. |
| AL / AT — `Facture` / `Mandat` | Une ligne est une réalisation si les deux sont renseignés. |
| AM — `Commande` | N° de BDC, renseigné seulement quand la ligne a été rattachée à une commande. |

### Charger plusieurs exports SEDIT

Les exports SEDIT sont annuels et se recouvrent : le suivi d'une opération
pluriannuelle se lit sur leur **réunion**, pas sur un seul fichier.
`MarchesAnalyzer` accepte donc un fichier, un motif, un répertoire, ou une liste
de ces formes :

```python
MarchesAnalyzer("data_sources/factures*.xls", database=db)
MarchesAnalyzer("data_sources", database=db)
MarchesAnalyzer(["factures_2024.xls", "factures_2025.xls"], database=db)
MarchesAnalyzer("database_sync", database=db)   # lit le cache existant
```

Chaque fichier est synchronisé indépendamment : le cache retient sa provenance,
une ligne présente dans plusieurs exports n'est stockée qu'une fois, et
synchroniser un fichier ne retire jamais les lignes des autres. Un fichier
retiré de la sélection voit ses lignes sortir du cache — sauf celles qu'un autre
export porte encore.

Un `cache_path` permet de désigner le fichier de cache, pour qu'un traitement en
lot ne modifie pas celui de l'application.

### Régénérer les suivis déjà diffusés

Les fichiers produits avant ce correctif comptent deux fois l'enveloppe
consommée. Pour les régénérer :

```bash
python regenerer_suivis.py --tout --sortie exports/
python regenerer_suivis.py 2020_14G3P --sortie exports/
python regenerer_suivis.py --tout --source "data_sources/factures*.xls" --sortie exports/
python regenerer_suivis.py --lister
```

### Tests

```bash
python -m unittest test_suivi_financier.py -v
python -m unittest test_matching.py -v
```
