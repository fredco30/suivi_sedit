# suivi_march-
Logiciel de gestion pour suivi des marchés

## Démarrer l'application

Sous Windows, double-cliquer **`Lancer_suivi_marches.cmd`**. Il repère Python,
installe les dépendances manquantes, puis démarre l'application en journalisant
dans `run_logs\`. En cas d'échec, le journal s'ouvre automatiquement.

```
Lancer_suivi_marches.cmd            lance l'application
Lancer_suivi_marches.cmd --maj      réinstalle les dépendances puis lance
Lancer_suivi_marches.cmd --verif    vérifie seulement, ne lance rien
```

L'installation se déclenche si un module manque **ou** si `requirements.txt` a
changé depuis la dernière installation réussie — une version relevée ou une
dépendance ajoutée qui se trouve déjà présente passerait sinon inaperçue. Une
empreinte du fichier est notée dans `run_logs/requirements_installees.txt`.

La logique est dans `lanceur.py` (utilisable seul : `python lanceur.py --verif`) ;
le `.cmd` ne fait que trouver Python. Si `pip` refuse les versions figées de
`requirements.txt` — `pandas==2.1.4` n'a pas de binaire au-delà de Python 3.12 —
le lanceur réessaie en les traitant comme des minima et le signale.

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

SEDIT tronque les libellés à largeur fixe : le marqueur `(REPORT)` en revient
souvent amputé (`georeferencement obligatoire(REP`). Ces débuts de marqueur sont
retirés en fin de libellé — une parenthèse ouverte n'y apporte rien —, mais une
parenthèse ouverte par une troncature ordinaire est conservée. Les n° de mandat,
que SEDIT remonte comme des nombres, s'écrivent « 1470 » et non « 1470.0 ».

### Un bloc par lot

Le tableau est découpé en **un bloc par marché**, c'est-à-dire par lot : le
marché est l'unité qui porte une enveloppe notifiée, donc la seule contre
laquelle un solde ait un sens. Chaque bloc a son sous-total et son propre
« restant ». Le prestataire et la tranche sont des attributs de l'**écriture**,
pas du bloc — un marché peut être tenu par un groupement d'entreprises et
couvrir plusieurs tranches ; les lire sur la première ligne du marché
réattribuait au mandataire les lignes de ses cotraitants, et faisait dépendre
le résultat de l'ordre de lecture des exports SEDIT.

L'enveloppe d'un lot est lue en base (montant initial, tranches optionnelles et
avenants) ; à défaut, elle est reconstituée depuis SEDIT **toutes tranches
additionnées**. C'est le calcul de l'onglet *Opérations*, si bien que l'export
et le tableau à l'écran ne peuvent plus afficher deux montants différents —
n'en retenir qu'une tranche sous-évaluait 5 opérations de 767 935,72 €.

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

### Enveloppes contractuelles des marchés

Le solde d'un marché se calcule contre son **enveloppe notifiée**, qui figure
dans l'acte d'engagement et dans aucun export SEDIT. Les montants SEDIT sont
ceux *engagés à ce jour* : sur les trois marchés où les deux valeurs sont
connues, l'écart va de 1,06 à 115 fois. Ils ne sont donc pas une approximation
utilisable, et le programme ne les recopie jamais en base.

La saisie se fait en une passe :

```bash
python enveloppes_marches.py exporter --source "data_sources/factures*.xls"
#  ... remplir la colonne « ENVELOPPE CONTRACTUELLE TTC » ...
python enveloppes_marches.py importer enveloppes_a_saisir.xlsx            # simulation
python enveloppes_marches.py importer enveloppes_a_saisir.xlsx --appliquer
python regenerer_suivis.py --tout --source "data_sources/factures*.xls" --sortie exports
```

Le tableau produit liste chaque marché avec son fournisseur, ses opérations, son
engagé et son facturé à ce jour — des repères pour la saisie, jamais l'enveloppe.
L'import est une simulation par défaut ; il repère ses colonnes par leur en-tête,
accepte les montants au format français et n'écrit rien si l'un est illisible.

### Depuis l'interface

Deux boutons couvrent le même besoin sans passer par la ligne de commande :

- **💶 Enveloppes des marchés** (onglet *Suivi des marchés*) ouvre la saisie en
  masse : un marché par ligne, seule la colonne enveloppe est éditable, les
  colonnes « engagé » et « facturé » ne sont que des repères. Le tableau
  s'exporte et se réimporte en `.xlsx` pour une saisie hors application, et rien
  n'est écrit tant que « Enregistrer » n'est pas cliqué.
- **🔁 Tout régénérer** (onglet *Opérations*) régénère toutes les opérations dans
  un dossier au choix, avec barre de progression et annulation.

### Filtres (Commandes, Rappels, Factures, Facturation)

Les filtres **Marché** et **Fournisseur** sont à sélection multiple avec
recherche (`SearchableCheckableComboBox`) : cliquer ouvre une liste à cases à
cocher avec un champ de recherche en tête, nécessaire dès que la centaine de
marchés ou de fournisseurs du dépôt rend une liste plate illisible. L'option
« [Tous] » coche/décoche l'intégralité de la liste, y compris ce qu'une
recherche en cours masque à l'écran.

L'onglet **Facturation** a un filtre **Exercice** fonctionnel : la colonne
existe dans le tableau (`FACTURATION_COLUMNS`) mais n'était filtrable nulle
part.

Chaque groupe de filtres qui doit apparaître/disparaître selon l'onglet actif
(Facturation/Exercice, Marché, les filtres multiples de Commandes, les deux
champs de recherche) est sa **propre `QToolBar`**, montrée/cachée dans son
ensemble plutôt qu'un widget isolé au sein d'une toolbar partagée. Cacher un
widget nu ajouté via `QToolBar.addWidget()` s'est révélé peu fiable : Qt
resynchronise chaque widget sur l'état de sa propre `QAction` dès qu'un widget
*sibling* de la même toolbar change de visibilité, ce qui pouvait laisser
apparaître un filtre masqué par erreur — c'est ce qui produisait un filtre
« Exercice » vide et non fonctionnel sur l'onglet Facturation. Cacher une
toolbar entière n'a pas ce problème. Pour la même raison, l'initialisation des
filtres du premier onglet est différée au premier `showEvent()` plutôt
qu'appelée pendant la construction de la fenêtre.

### Tests

L'interface est testée sans affichage (`QT_QPA_PLATFORM=offscreen`) ;
`test_interface.py` et `test_filtres_ui.py` sont ignorés si PyQt5 n'est pas
installé.


```bash
python -m unittest test_suivi_financier.py -v
python -m unittest test_matching.py -v
python -m unittest test_interface.py -v
python -m unittest test_lanceur.py -v
python -m unittest test_filtres_ui.py -v
```
