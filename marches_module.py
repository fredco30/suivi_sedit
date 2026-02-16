"""
Module de suivi financier des marchés publics
Calcule la vision globale par marché et le détail par tranche
Avec synchronisation SQLite pour optimiser les performances
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Tuple, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from marches_sync import MarchesSync


class MarchesAnalyzer:
    """
    Analyseur des données de marchés publics à partir du fichier Excel.
    Utilise un cache SQLite pour améliorer les performances.

    Colonnes Excel utilisées (index 0-based):
    - O (14): Montant initial
    - AD (29): Montant TTC
    - AH (33): Montant service fait
    - AI (34): Date service fait
    - AM (38): Commande
    - AN (39): Marché
    - AO (40): Tranche
    - AT (45): Mandat
    """

    # Index des colonnes dans le fichier Excel
    COL_MONTANT_INITIAL = 14   # O
    COL_MONTANT_TTC = 29       # AD
    COL_MONTANT_SF = 33        # AH
    COL_DATE_SF = 34           # AI
    COL_COMMANDE = 38          # AM
    COL_MARCHE = 39            # AN
    COL_TRANCHE = 40           # AO
    COL_MANDAT = 45            # AT
    COL_FOURNISSEUR = 8        # I - Nom tiers
    COL_LIBELLE = 13           # N - Libellé
    COL_FACTURE = 36           # AL - Facture

    def __init__(self, excel_path: str, database=None, use_cache: bool = True):
        """
        Initialise l'analyseur avec le chemin du fichier Excel.

        Args:
            excel_path: Chemin vers le fichier Excel des factures
            database: Instance de Database pour accéder aux ajustements manuels (optionnel)
            use_cache: Utiliser le cache SQLite (True par défaut)
        """
        self.excel_path = excel_path
        self.df = None
        self.df_marches = None
        self.db = database
        self.use_cache = use_cache
        self.sync = MarchesSync() if use_cache else None
        self.sync_stats = None

    @staticmethod
    def extract_exercice_from_bdc(num_commande) -> str:
        """Déduit l'exercice à partir du n° de BDC (2 premiers chiffres)."""
        if not num_commande:
            return "Inconnu"
        num_str = str(num_commande).strip()
        if len(num_str) < 2 or not num_str[:2].isdigit():
            return "Inconnu"
        return f"20{num_str[:2]}"

    def get_exercices_for_operation(self, code_operation: str) -> List[str]:
        """Retourne la liste des exercices disponibles pour une opération."""
        operations_data = self.get_vision_operations()
        operation_info = next((op for op in operations_data if op['operation'] == code_operation), None)
        if not operation_info:
            return ["Tous"]

        marches_operation = operation_info['marches']
        exercices = set()
        for marche in marches_operation:
            df_marche = self.df_marches[self.df_marches.iloc[:, self.COL_MARCHE] == marche]
            for _, row in df_marche.iterrows():
                num_commande = row.iloc[self.COL_COMMANDE] if not pd.isna(row.iloc[self.COL_COMMANDE]) else ""
                exercices.add(self.extract_exercice_from_bdc(num_commande))

        exercices_list = sorted(exercices)
        return ["Tous"] + exercices_list

    def load_data(self, force_reload: bool = False):
        """
        Charge les données depuis le fichier Excel avec synchronisation SQLite.

        Args:
            force_reload: Force le rechargement depuis Excel même si le cache est valide

        Returns:
            True si succès, False sinon
        """
        try:
            if self.use_cache and self.sync:
                # Si le chemin est "database_sync", charger uniquement depuis le cache
                # (les données ont déjà été synchronisées depuis la base de données)
                if self.excel_path == "database_sync":
                    print(f"[CACHE] Chargement depuis le cache (source: database)")
                    self.sync_stats = {'status': 'cached', 'message': 'Données depuis database'}
                else:
                    # Vérifier si synchronisation nécessaire
                    needs_sync, reason = self.sync.file_needs_sync(self.excel_path)

                    if needs_sync or force_reload:
                        print(f"[SYNC] Synchronisation necessaire: {reason}")

                        # Charger depuis Excel pour synchroniser
                        df_excel = pd.read_excel(self.excel_path)

                        # Préparer les données pour la synchronisation
                        df_to_sync = df_excel[
                            df_excel.iloc[:, self.COL_MARCHE].notna() &
                            (df_excel.iloc[:, self.COL_MARCHE] != '')
                        ].copy()

                        # Renommer les colonnes pour correspondre au schéma SQLite
                        df_to_sync_renamed = pd.DataFrame({
                            'marche': df_to_sync.iloc[:, self.COL_MARCHE],
                            'fournisseur': df_to_sync.iloc[:, self.COL_FOURNISSEUR],
                            'libelle': df_to_sync.iloc[:, self.COL_LIBELLE],
                            'date_sf': df_to_sync.iloc[:, self.COL_DATE_SF],
                            'num_facture': df_to_sync.iloc[:, self.COL_FACTURE],
                            'montant_initial': df_to_sync.iloc[:, self.COL_MONTANT_INITIAL],
                            'montant_sf': df_to_sync.iloc[:, self.COL_MONTANT_SF],
                            'montant_ttc': df_to_sync.iloc[:, self.COL_MONTANT_TTC],
                            'num_mandat': df_to_sync.iloc[:, self.COL_MANDAT],
                            'tranche': df_to_sync.iloc[:, self.COL_TRANCHE],
                            'commande': df_to_sync.iloc[:, self.COL_COMMANDE],
                        })

                        # Synchroniser vers SQLite
                        self.sync_stats = self.sync.sync_from_excel(
                            self.excel_path,
                            df_to_sync_renamed,
                            force=force_reload
                        )

                        print(f"[OK] Sync terminee: {self.sync_stats['nb_inserted']} inserees, "
                              f"{self.sync_stats['nb_deleted']} supprimees, "
                              f"{self.sync_stats['nb_unchanged']} inchangees "
                              f"(duree: {self.sync_stats['duration']:.2f}s)")

                    else:
                        print(f"[CACHE] Chargement depuis le cache SQLite: {reason}")
                        self.sync_stats = {'status': 'cached', 'message': reason}

                # Charger depuis SQLite (beaucoup plus rapide)
                df_from_cache = self.sync.load_to_dataframe()

                # Reconstituer un DataFrame avec la structure Excel (colonnes par index)
                # Créer un DataFrame vide avec 50 colonnes (suffisant pour couvrir tous les index)
                nb_cols = max([
                    self.COL_MONTANT_INITIAL, self.COL_MONTANT_TTC, self.COL_MONTANT_SF,
                    self.COL_DATE_SF, self.COL_COMMANDE, self.COL_MARCHE, self.COL_TRANCHE,
                    self.COL_MANDAT, self.COL_FOURNISSEUR, self.COL_LIBELLE, self.COL_FACTURE
                ]) + 1

                self.df = pd.DataFrame(index=df_from_cache.index, columns=range(nb_cols))

                # Mapper les colonnes du cache vers les index Excel
                self.df.iloc[:, self.COL_MARCHE] = df_from_cache['marche']
                self.df.iloc[:, self.COL_FOURNISSEUR] = df_from_cache['fournisseur']
                self.df.iloc[:, self.COL_LIBELLE] = df_from_cache['libelle']
                self.df.iloc[:, self.COL_DATE_SF] = df_from_cache['date_sf']
                self.df.iloc[:, self.COL_FACTURE] = df_from_cache['num_facture']
                self.df.iloc[:, self.COL_MONTANT_INITIAL] = df_from_cache['montant_initial']
                self.df.iloc[:, self.COL_MONTANT_SF] = df_from_cache['montant_sf']
                self.df.iloc[:, self.COL_MONTANT_TTC] = df_from_cache['montant_ttc']
                self.df.iloc[:, self.COL_MANDAT] = df_from_cache['num_mandat']
                self.df.iloc[:, self.COL_TRANCHE] = df_from_cache['tranche']
                self.df.iloc[:, self.COL_COMMANDE] = df_from_cache['commande']

                # df_marches pointe déjà vers le DataFrame filtré
                self.df_marches = self.df.copy()

            else:
                # Mode sans cache (chargement direct depuis Excel)
                print("📂 Chargement direct depuis Excel (cache désactivé)")
                self.df = pd.read_excel(self.excel_path)
                self.df_marches = self.df[
                    self.df.iloc[:, self.COL_MARCHE].notna() &
                    (self.df.iloc[:, self.COL_MARCHE] != '')
                ].copy()

            return True

        except Exception as e:
            print(f"[ERREUR] Erreur lors du chargement: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _get_col_value(self, row, col_idx, default=None):
        """Récupère une valeur de colonne de manière sécurisée."""
        try:
            value = row.iloc[col_idx]
            if pd.isna(value) or value == '':
                return default
            return value
        except:
            return default

    def calculate_montant_initial_tranche(self, marche: str, tranche) -> float:
        """
        Calcule le montant initial d'une tranche.

        Priorité :
        1. Pour la TF : Utiliser montant_initial_manuel du marché
        2. Pour les TOs : Utiliser les tranches définies dans la base de données (table tranches)
        3. Sinon, calculer depuis Excel (ancien comportement)

        Règle métier (calcul Excel) :
        - Prendre Montant initial (O)
        - Ignorer les 0,00
        - Dédupliquer les valeurs positives (souvent répétées)
        - Additionner ces valeurs positives distinctes
        """
        # Essayer d'abord de charger depuis la base de données
        if self.db:
            # Normaliser la tranche en valeur numérique
            tranche_num = None
            if not pd.isna(tranche):
                try:
                    tranche_num = float(tranche)
                except (ValueError, TypeError):
                    pass

            # Convertir le numéro de tranche en code (0=TF, 1=TO1, 2=TO2, etc.)
            if pd.isna(tranche) or tranche_num == 0:
                code_tranche = "TF"
                # Pour la TF, utiliser montant_initial_manuel du marché
                marche_record = self.db.get_marche(marche)
                if marche_record and marche_record["montant_initial_manuel"]:
                    montant = float(marche_record["montant_initial_manuel"])
                    print(f"[TRANCHE] {marche} TF: {montant} € (depuis montant_initial_manuel)")
                    return montant
            else:
                # Pour les TOs (TO1, TO2, etc.)
                code_tranche = f"TO{int(tranche_num)}" if tranche_num else str(tranche)

                # Pour les TOs, chercher dans la table tranches
                cur = self.db.conn.cursor()
                cur.execute(
                    "SELECT montant FROM tranches WHERE code_marche = ? AND code_tranche = ?",
                    (marche, code_tranche)
                )
                row = cur.fetchone()
                if row and row["montant"]:
                    print(f"[TRANCHE] {marche} {code_tranche}: {row['montant']} € (depuis DB)")
                    return float(row["montant"])

        # Fallback : calculer depuis Excel (ancien comportement)
        # Filtrer les lignes pour ce marché et cette tranche
        mask = (self.df_marches.iloc[:, self.COL_MARCHE] == marche)

        # Gérer le cas où tranche peut être NaN, None ou une valeur
        if pd.isna(tranche):
            mask &= self.df_marches.iloc[:, self.COL_TRANCHE].isna()
        else:
            mask &= (self.df_marches.iloc[:, self.COL_TRANCHE] == tranche)

        df_tranche = self.df_marches[mask]

        if len(df_tranche) == 0:
            return 0.0

        # Récupérer les montants initiaux
        montants = df_tranche.iloc[:, self.COL_MONTANT_INITIAL].copy()

        # Convertir en float et ignorer les NaN
        montants = pd.to_numeric(montants, errors='coerce').fillna(0.0)

        # Ignorer les 0,00
        montants = montants[montants > 0.01]

        # Dédupliquer les valeurs
        montants_uniques = montants.unique()

        # Additionner
        montant_calcule = float(montants_uniques.sum())
        if montant_calcule > 0:
            print(f"[TRANCHE] {marche} tranche {tranche}: {montant_calcule} € (calculé depuis Excel)")
        return montant_calcule

    def calculate_service_fait_tranche(self, marche: str, tranche) -> float:
        """
        Calcule le service fait cumulé pour une tranche.

        Règle : somme de AH pour ce couple (marché, tranche)
        """
        mask = (self.df_marches.iloc[:, self.COL_MARCHE] == marche)

        if pd.isna(tranche):
            mask &= self.df_marches.iloc[:, self.COL_TRANCHE].isna()
        else:
            mask &= (self.df_marches.iloc[:, self.COL_TRANCHE] == tranche)

        df_tranche = self.df_marches[mask]

        if len(df_tranche) == 0:
            return 0.0

        sf = df_tranche.iloc[:, self.COL_MONTANT_SF].copy()
        sf = pd.to_numeric(sf, errors='coerce').fillna(0.0)

        return float(sf.sum())

    def calculate_paye_tranche(self, marche: str, tranche) -> float:
        """
        Calcule le payé cumulé pour une tranche.

        Règle : somme de AD pour ce couple (marché, tranche) avec AT (mandat) non vide
        """
        mask = (self.df_marches.iloc[:, self.COL_MARCHE] == marche)

        if pd.isna(tranche):
            mask &= self.df_marches.iloc[:, self.COL_TRANCHE].isna()
        else:
            mask &= (self.df_marches.iloc[:, self.COL_TRANCHE] == tranche)

        # Ajouter le filtre sur le mandat (doit être renseigné)
        mask &= self.df_marches.iloc[:, self.COL_MANDAT].notna()
        mask &= (self.df_marches.iloc[:, self.COL_MANDAT] != '')

        df_tranche = self.df_marches[mask]

        if len(df_tranche) == 0:
            return 0.0

        ttc = df_tranche.iloc[:, self.COL_MONTANT_TTC].copy()
        ttc = pd.to_numeric(ttc, errors='coerce').fillna(0.0)

        return float(ttc.sum())

    def get_vision_detaillee(self) -> List[Dict]:
        """
        Retourne la vision détaillée : 1 ligne par couple (marché, tranche).

        Colonnes retournées :
        - marche
        - tranche (affiché comme TF, TO1, TO2... ou la valeur brute)
        - montant_initial_tranche
        - service_fait_tranche
        - paye_tranche
        - pourcent_consomme_tranche
        """
        if self.df_marches is None or len(self.df_marches) == 0:
            return []

        results = []

        # Grouper par (marché, tranche)
        grouped = self.df_marches.groupby(
            [self.df_marches.iloc[:, self.COL_MARCHE],
             self.df_marches.iloc[:, self.COL_TRANCHE]]
        )

        for (marche, tranche), _ in grouped:
            montant_initial = self.calculate_montant_initial_tranche(marche, tranche)
            service_fait = self.calculate_service_fait_tranche(marche, tranche)
            paye = self.calculate_paye_tranche(marche, tranche)

            # Calculer le pourcentage
            pourcent = 0.0
            if montant_initial > 0:
                pourcent = (service_fait / montant_initial) * 100

            # Formater le libellé de la tranche
            if pd.isna(tranche):
                tranche_libelle = "Sans tranche"
            else:
                # Si c'est un nombre, formater en TF, TO1, TO2...
                try:
                    tranche_num = int(float(tranche))
                    if tranche_num == 0:
                        tranche_libelle = "TF"
                    else:
                        tranche_libelle = f"TO{tranche_num}"
                except:
                    tranche_libelle = str(tranche)

            results.append({
                'marche': marche,
                'tranche': tranche,
                'tranche_libelle': tranche_libelle,
                'montant_initial_tranche': montant_initial,
                'service_fait_tranche': service_fait,
                'paye_tranche': paye,
                'pourcent_consomme_tranche': pourcent
            })

        # Trier par marché puis tranche
        results.sort(key=lambda x: (x['marche'], x['tranche'] if not pd.isna(x['tranche']) else -1))

        return results

    def get_vision_globale(self) -> List[Dict]:
        """
        Retourne la vision globale : 1 ligne par marché.

        Colonnes retournées :
        - marche
        - libelle_marche
        - fournisseur
        - montant_initial_marche
        - service_fait_cumule
        - paye_cumule
        - reste_a_realiser
        - reste_a_mandater
        - pourcent_consomme
        """
        if self.df_marches is None or len(self.df_marches) == 0:
            return []

        results = []

        # Grouper par marché
        marches_uniques = self.df_marches.iloc[:, self.COL_MARCHE].unique()

        for marche in marches_uniques:
            df_marche = self.df_marches[self.df_marches.iloc[:, self.COL_MARCHE] == marche]

            # Récupérer le fournisseur (prendre le premier non vide)
            fournisseurs = df_marche.iloc[:, self.COL_FOURNISSEUR].dropna()
            fournisseur = fournisseurs.iloc[0] if len(fournisseurs) > 0 else ""

            # Récupérer le libellé (prendre le premier non vide)
            libelles = df_marche.iloc[:, self.COL_LIBELLE].dropna()
            libelle = libelles.iloc[0] if len(libelles) > 0 else ""

            # Calculer le montant initial du marché
            # Priorité 1: Montant manuel + avenants depuis la BD
            # Priorité 2: Calcul automatique depuis Excel (somme des tranches)
            montant_initial_marche = 0.0
            montant_excel = 0.0
            nb_avenants = 0

            if self.db:
                # Essayer de récupérer le montant depuis la base de données
                montant_bd = self.db.get_montant_total_marche(marche)
                if montant_bd > 0:
                    montant_initial_marche = montant_bd

                    # Récupérer le nombre d'avenants pour affichage
                    avenants = self.db.get_avenants(marche)
                    nb_avenants = len(avenants) if avenants else 0

            # Si pas de montant manuel en BD, calculer depuis Excel
            if montant_initial_marche == 0:
                tranches = df_marche.iloc[:, self.COL_TRANCHE].unique()
                montant_excel = sum(
                    self.calculate_montant_initial_tranche(marche, t)
                    for t in tranches
                )
                montant_initial_marche = montant_excel

            # Service fait cumulé (somme de AH pour tout le marché)
            sf = df_marche.iloc[:, self.COL_MONTANT_SF].copy()
            sf = pd.to_numeric(sf, errors='coerce').fillna(0.0)
            service_fait_cumule = float(sf.sum())

            # Payé cumulé (somme de AD avec mandat non vide)
            df_mandates = df_marche[
                df_marche.iloc[:, self.COL_MANDAT].notna() &
                (df_marche.iloc[:, self.COL_MANDAT] != '')
            ]
            ttc = df_mandates.iloc[:, self.COL_MONTANT_TTC].copy()
            ttc = pd.to_numeric(ttc, errors='coerce').fillna(0.0)
            paye_cumule = float(ttc.sum())

            # Calculs dérivés
            reste_a_realiser = montant_initial_marche - service_fait_cumule
            reste_a_mandater = service_fait_cumule - paye_cumule

            pourcent_consomme = 0.0
            if montant_initial_marche > 0:
                pourcent_consomme = (service_fait_cumule / montant_initial_marche) * 100

            # Extraire le code opération
            code_operation = self.extract_operation(marche)

            results.append({
                'marche': marche,
                'operation': code_operation,
                'libelle_marche': libelle,
                'fournisseur': fournisseur,
                'montant_initial_marche': montant_initial_marche,
                'service_fait_cumule': service_fait_cumule,
                'paye_cumule': paye_cumule,
                'reste_a_realiser': reste_a_realiser,
                'reste_a_mandater': reste_a_mandater,
                'pourcent_consomme': pourcent_consomme,
                'nb_avenants': nb_avenants,
                'montant_excel': montant_excel
            })

        # Trier par marché
        results.sort(key=lambda x: x['marche'])

        return results

    def get_tranches_for_marche(self, marche: str) -> List[Dict]:
        """
        Retourne le détail des tranches pour un marché donné.
        """
        vision_detaillee = self.get_vision_detaillee()
        return [t for t in vision_detaillee if t['marche'] == marche]

    @staticmethod
    def extract_operation(marche: str) -> str:
        """
        Extrait le code opération d'un code marché.

        Règles :
        - Si >= 2 séparateurs ET dernier segment = petit numéro → c'est un LOT
          Exemples: 2024_17_1 → 2024_17 | 2024_1_3 → 2024_1
        - Si 1 seul séparateur → c'est une OPÉRATION INDÉPENDANTE (pas de regroupement)
          Exemples: 2025_12 → 2025_12 | 2023_17 → 2023_17
        - Cas spéciaux restent inchangés
          Exemples: 2020_14G3P → 2020_14G3P
        """
        import re

        if not marche:
            return ""

        normalized = str(marche).strip()

        # Compter les séparateurs (_ et -)
        nb_underscores = normalized.count('_')
        nb_dashes = normalized.count('-')
        total_separators = nb_underscores + nb_dashes

        # Si moins de 2 séparateurs → le marché EST l'opération (pas de regroupement)
        if total_separators < 2:
            return normalized

        # Si >= 2 séparateurs, vérifier si le dernier segment est un numéro de lot
        last_underscore = normalized.rfind('_')
        last_dash = normalized.rfind('-')
        last_sep = max(last_underscore, last_dash)

        if last_sep > 0:
            suffix = normalized[last_sep+1:]

            # Vérifier si c'est un petit numéro (1-2 chiffres) = probable numéro de lot
            if suffix.isdigit() and len(suffix) <= 2:
                # C'est un lot, extraire l'opération
                return normalized[:last_sep]

        # Sinon, le marché est l'opération complète
        return normalized

    def get_vision_operations(self) -> List[Dict]:
        """
        Retourne la vision par opération (regroupement de marchés/lots).

        Retourne :
        - operation : code opération
        - nb_lots : nombre de marchés/lots
        - marchés : liste des codes marchés
        - montant_initial_total
        - service_fait_total
        - paye_total
        - reste_a_realiser
        - reste_a_mandater
        - pourcent_consomme
        - nb_avenants_total
        """
        vision_marches = self.get_vision_globale()

        if not vision_marches:
            return []

        # Regrouper par opération
        operations_dict = {}

        for marche_data in vision_marches:
            operation = self.extract_operation(marche_data['marche'])

            if operation not in operations_dict:
                operations_dict[operation] = {
                    'operation': operation,
                    'marches': [],
                    'libelles': [],
                    'fournisseurs': [],
                    'montant_initial_total': 0.0,
                    'service_fait_total': 0.0,
                    'paye_total': 0.0,
                    'nb_avenants_total': 0
                }

            op = operations_dict[operation]
            op['marches'].append(marche_data['marche'])

            # Ajouter le libellé s'il n'est pas déjà présent
            libelle = marche_data.get('libelle_marche', '')
            if libelle and libelle not in op['libelles']:
                op['libelles'].append(libelle)

            # Ajouter le fournisseur s'il n'est pas déjà présent
            fournisseur = marche_data.get('fournisseur', '')
            if fournisseur and fournisseur not in op['fournisseurs']:
                op['fournisseurs'].append(fournisseur)

            op['montant_initial_total'] += marche_data['montant_initial_marche']
            op['service_fait_total'] += marche_data['service_fait_cumule']
            op['paye_total'] += marche_data['paye_cumule']
            op['nb_avenants_total'] += marche_data['nb_avenants']

        # Calculer les résultats finaux
        results = []
        for operation, data in operations_dict.items():
            nb_lots = len(data['marches'])
            montant_initial = data['montant_initial_total']
            sf_total = data['service_fait_total']
            paye_total = data['paye_total']

            reste_a_realiser = montant_initial - sf_total
            reste_a_mandater = sf_total - paye_total

            pourcent = 0.0
            if montant_initial > 0:
                pourcent = (sf_total / montant_initial) * 100

            # Combiner les libellés et fournisseurs
            libelle_combined = " | ".join(data['libelles']) if data['libelles'] else ""
            fournisseur_combined = " | ".join(data['fournisseurs']) if data['fournisseurs'] else ""

            results.append({
                'operation': operation,
                'nb_lots': nb_lots,
                'marches': data['marches'],
                'libelle': libelle_combined,
                'fournisseur': fournisseur_combined,
                'montant_initial_total': montant_initial,
                'service_fait_total': sf_total,
                'paye_total': paye_total,
                'reste_a_realiser': reste_a_realiser,
                'reste_a_mandater': reste_a_mandater,
                'pourcent_consomme': pourcent,
                'nb_avenants_total': data['nb_avenants_total']
            })

        # Trier par opération
        results.sort(key=lambda x: x['operation'])

        return results

    def get_historique_factures(self, marche: str = None) -> List[Dict]:
        """
        Retourne l'historique détaillé des factures/paiements.

        Args:
            marche: Si spécifié, retourne uniquement l'historique de ce marché.
                   Si None, retourne l'historique complet de tous les marchés.

        Retourne une liste de dictionnaires avec :
        - marche
        - date_sf
        - num_facture
        - libelle
        - montant_sf
        - montant_ttc
        - num_mandat
        - statut (Payé / Service fait / Facturée / En attente)
        """
        if self.df_marches is None or len(self.df_marches) == 0:
            return []

        # Filtrer par marché si spécifié
        if marche:
            df_filtered = self.df_marches[self.df_marches.iloc[:, self.COL_MARCHE] == marche]
        else:
            df_filtered = self.df_marches

        results = []

        for idx, row in df_filtered.iterrows():
            marche_code = row.iloc[self.COL_MARCHE]
            fournisseur = row.iloc[self.COL_FOURNISSEUR]
            date_sf = row.iloc[self.COL_DATE_SF]
            num_facture = row.iloc[self.COL_FACTURE]
            libelle = row.iloc[self.COL_LIBELLE]
            montant_sf = row.iloc[self.COL_MONTANT_SF]
            montant_ttc = row.iloc[self.COL_MONTANT_TTC]
            num_mandat = row.iloc[self.COL_MANDAT]

            # Convertir les valeurs
            try:
                montant_sf = float(montant_sf) if pd.notna(montant_sf) else 0.0
            except:
                montant_sf = 0.0

            try:
                montant_ttc = float(montant_ttc) if pd.notna(montant_ttc) else 0.0
            except:
                montant_ttc = 0.0

            # Formater la date
            date_sf_str = ""
            if pd.notna(date_sf):
                try:
                    if isinstance(date_sf, str):
                        date_sf_str = date_sf
                    else:
                        date_sf_str = date_sf.strftime("%d/%m/%Y")
                except:
                    date_sf_str = str(date_sf)

            # Déterminer le statut
            has_mandat = pd.notna(num_mandat) and str(num_mandat).strip() != ''
            has_facture = pd.notna(num_facture) and str(num_facture).strip() != ''
            has_sf = montant_sf > 0.01

            if has_mandat:
                statut = "✅ Payé"
            elif has_sf:
                statut = "⏳ Service fait"
            elif has_facture:
                statut = "📋 Facturée"
            else:
                statut = "⚠️ En attente"

            results.append({
                'marche': marche_code,
                'fournisseur': str(fournisseur) if pd.notna(fournisseur) else "",
                'date_sf': date_sf_str,
                'num_facture': str(num_facture) if pd.notna(num_facture) else "",
                'libelle': str(libelle) if pd.notna(libelle) else "",
                'montant_sf': montant_sf,
                'montant_ttc': montant_ttc,
                'num_mandat': str(num_mandat) if pd.notna(num_mandat) else "",
                'statut': statut
            })

        # Trier par marché puis par date
        results.sort(key=lambda x: (x['marche'], x['date_sf']), reverse=True)

        return results

    def export_to_excel(self, filepath: str) -> bool:
        """
        Exporte toutes les données dans un fichier Excel avec 5 feuilles :
        1. Opérations - Vision par opération
        2. Marchés - Vision globale par marché
        3. Tranches - Détail par tranche
        4. Avenants - Liste de tous les avenants
        5. Historique - Historique complet des factures/paiements

        Returns:
            bool: True si succès, False sinon
        """
        try:
            wb = Workbook()

            # Styles communs
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="0078D4", end_color="0078D4", fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            border_thin = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            # ===============================
            # FEUILLE 1 : OPÉRATIONS
            # ===============================
            ws_operations = wb.active
            ws_operations.title = "Opérations"

            vision_operations = self.get_vision_operations()

            # En-têtes
            headers_operations = [
                "Opération", "Nb lots", "Marchés", "Libellé", "Fournisseur",
                "Montant initial total", "Avenants", "Service fait total", "Payé total",
                "Reste à réaliser", "Reste à mandater", "% consommé"
            ]

            for col_idx, header in enumerate(headers_operations, 1):
                cell = ws_operations.cell(1, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Données
            for row_idx, op in enumerate(vision_operations, 2):
                ws_operations.cell(row_idx, 1, op.get('operation', ''))
                ws_operations.cell(row_idx, 2, op.get('nb_lots', 0))
                marches_str = ", ".join(op.get('marches', [])) if isinstance(op.get('marches'), list) else str(op.get('marches', ''))
                ws_operations.cell(row_idx, 3, marches_str)
                ws_operations.cell(row_idx, 4, op.get('libelle', ''))
                ws_operations.cell(row_idx, 5, op.get('fournisseur', ''))
                ws_operations.cell(row_idx, 6, op.get('montant_initial_total', 0))
                ws_operations.cell(row_idx, 7, op.get('nb_avenants_total', 0))
                ws_operations.cell(row_idx, 8, op.get('service_fait_total', 0))
                ws_operations.cell(row_idx, 9, op.get('paye_total', 0))
                ws_operations.cell(row_idx, 10, op.get('reste_a_realiser', 0))
                ws_operations.cell(row_idx, 11, op.get('reste_a_mandater', 0))
                ws_operations.cell(row_idx, 12, op.get('pourcent_consomme', 0) / 100)

                # Format numérique
                for col in [6, 8, 9, 10, 11]:
                    ws_operations.cell(row_idx, col).number_format = '#,##0.00 €'
                ws_operations.cell(row_idx, 12).number_format = '0.00%'

            # Ajuster les largeurs
            ws_operations.column_dimensions['A'].width = 20
            ws_operations.column_dimensions['B'].width = 10
            ws_operations.column_dimensions['C'].width = 40
            ws_operations.column_dimensions['D'].width = 40  # Libellé
            ws_operations.column_dimensions['E'].width = 30  # Fournisseur
            for col in ['F', 'G', 'H', 'I', 'J', 'K']:
                ws_operations.column_dimensions[col].width = 18
            ws_operations.column_dimensions['L'].width = 12

            # ===============================
            # FEUILLE 2 : MARCHÉS
            # ===============================
            ws_marches = wb.create_sheet("Marchés")

            vision_globale = self.get_vision_globale()

            # En-têtes
            headers_marches = [
                "Marché", "Opération", "Libellé", "Fournisseur", "Montant initial",
                "Avenants", "Service fait cumulé", "Payé cumulé",
                "Reste à réaliser", "Reste à mandater", "% consommé"
            ]

            for col_idx, header in enumerate(headers_marches, 1):
                cell = ws_marches.cell(1, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Données
            for row_idx, marche in enumerate(vision_globale, 2):
                ws_marches.cell(row_idx, 1, marche.get('marche', ''))
                ws_marches.cell(row_idx, 2, marche.get('operation', ''))
                ws_marches.cell(row_idx, 3, marche.get('libelle_marche', ''))
                ws_marches.cell(row_idx, 4, marche.get('fournisseur', ''))
                ws_marches.cell(row_idx, 5, marche.get('montant_initial_marche', 0))
                ws_marches.cell(row_idx, 6, marche.get('nb_avenants', 0))
                ws_marches.cell(row_idx, 7, marche.get('service_fait_cumule', 0))
                ws_marches.cell(row_idx, 8, marche.get('paye_cumule', 0))
                ws_marches.cell(row_idx, 9, marche.get('reste_a_realiser', 0))
                ws_marches.cell(row_idx, 10, marche.get('reste_a_mandater', 0))
                ws_marches.cell(row_idx, 11, marche.get('pourcent_consomme', 0) / 100)

                # Format numérique
                for col in [5, 7, 8, 9, 10]:
                    ws_marches.cell(row_idx, col).number_format = '#,##0.00 €'
                ws_marches.cell(row_idx, 11).number_format = '0.00%'

            # Ajuster les largeurs
            ws_marches.column_dimensions['A'].width = 20
            ws_marches.column_dimensions['B'].width = 18
            ws_marches.column_dimensions['C'].width = 40
            ws_marches.column_dimensions['D'].width = 30
            for col in ['E', 'F', 'G', 'H', 'I', 'J']:
                ws_marches.column_dimensions[col].width = 18
            ws_marches.column_dimensions['K'].width = 12

            # ===============================
            # FEUILLE 3 : TRANCHES
            # ===============================
            ws_tranches = wb.create_sheet("Tranches")

            vision_detaillee = self.get_vision_detaillee()

            # En-têtes
            headers_tranches = [
                "Marché", "Tranche", "Montant initial tranche",
                "Service fait tranche", "Payé tranche", "% consommé"
            ]

            for col_idx, header in enumerate(headers_tranches, 1):
                cell = ws_tranches.cell(1, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Données
            for row_idx, tranche in enumerate(vision_detaillee, 2):
                ws_tranches.cell(row_idx, 1, tranche.get('marche', ''))
                ws_tranches.cell(row_idx, 2, tranche.get('tranche_libelle', ''))
                ws_tranches.cell(row_idx, 3, tranche.get('montant_initial_tranche', 0))
                ws_tranches.cell(row_idx, 4, tranche.get('service_fait_tranche', 0))
                ws_tranches.cell(row_idx, 5, tranche.get('paye_tranche', 0))
                ws_tranches.cell(row_idx, 6, tranche.get('pourcent_consomme_tranche', 0) / 100)

                # Format numérique
                for col in [3, 4, 5]:
                    ws_tranches.cell(row_idx, col).number_format = '#,##0.00 €'
                ws_tranches.cell(row_idx, 6).number_format = '0.00%'

            # Ajuster les largeurs
            ws_tranches.column_dimensions['A'].width = 20
            ws_tranches.column_dimensions['B'].width = 15
            for col in ['C', 'D', 'E']:
                ws_tranches.column_dimensions[col].width = 18
            ws_tranches.column_dimensions['F'].width = 12

            # ===============================
            # FEUILLE 4 : AVENANTS
            # ===============================
            ws_avenants = wb.create_sheet("Avenants")

            # Récupérer tous les avenants depuis la base de données
            avenants_list = []
            if self.db:
                # Récupérer tous les marchés
                marches_uniques = self.df_marches.iloc[:, self.COL_MARCHE].unique()
                for marche in marches_uniques:
                    avenants = self.db.get_avenants(marche)
                    if avenants:
                        for avenant in avenants:
                            avenants_list.append({
                                'marche': marche,
                                'numero': avenant.get('numero_avenant', ''),
                                'libelle': avenant.get('libelle', ''),
                                'montant': avenant.get('montant', 0),
                                'type': avenant.get('type_modification', ''),
                                'date': avenant.get('date_avenant', ''),
                                'motif': avenant.get('motif', '')
                            })

            # En-têtes
            headers_avenants = [
                "Marché", "N° Avenant", "Libellé", "Montant",
                "Type", "Date", "Motif"
            ]

            for col_idx, header in enumerate(headers_avenants, 1):
                cell = ws_avenants.cell(1, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Données
            for row_idx, avenant in enumerate(avenants_list, 2):
                ws_avenants.cell(row_idx, 1, avenant.get('marche', ''))
                ws_avenants.cell(row_idx, 2, avenant.get('numero', ''))
                ws_avenants.cell(row_idx, 3, avenant.get('libelle', ''))
                ws_avenants.cell(row_idx, 4, avenant.get('montant', 0))
                ws_avenants.cell(row_idx, 5, avenant.get('type', ''))
                ws_avenants.cell(row_idx, 6, avenant.get('date', ''))
                ws_avenants.cell(row_idx, 7, avenant.get('motif', ''))

                # Format numérique
                ws_avenants.cell(row_idx, 4).number_format = '#,##0.00 €'

            # Ajuster les largeurs
            ws_avenants.column_dimensions['A'].width = 20
            ws_avenants.column_dimensions['B'].width = 12
            ws_avenants.column_dimensions['C'].width = 35
            ws_avenants.column_dimensions['D'].width = 15
            ws_avenants.column_dimensions['E'].width = 15
            ws_avenants.column_dimensions['F'].width = 12
            ws_avenants.column_dimensions['G'].width = 40

            # ===============================
            # FEUILLE 5 : HISTORIQUE
            # ===============================
            ws_historique = wb.create_sheet("Historique")

            historique = self.get_historique_factures()

            # En-têtes
            headers_historique = [
                "Marché", "Fournisseur", "Date SF", "N° Facture", "Libellé",
                "Montant SF", "Montant TTC", "N° Mandat", "Statut"
            ]

            for col_idx, header in enumerate(headers_historique, 1):
                cell = ws_historique.cell(1, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Données
            for row_idx, ligne in enumerate(historique, 2):
                ws_historique.cell(row_idx, 1, ligne.get('marche', ''))
                ws_historique.cell(row_idx, 2, ligne.get('fournisseur', ''))
                ws_historique.cell(row_idx, 3, ligne.get('date_sf', ''))
                ws_historique.cell(row_idx, 4, ligne.get('num_facture', ''))
                ws_historique.cell(row_idx, 5, ligne.get('libelle', ''))
                ws_historique.cell(row_idx, 6, ligne.get('montant_sf', 0))
                ws_historique.cell(row_idx, 7, ligne.get('montant_ttc', 0))
                ws_historique.cell(row_idx, 8, ligne.get('num_mandat', ''))
                ws_historique.cell(row_idx, 9, ligne.get('statut', ''))

                # Format numérique
                for col in [6, 7]:
                    ws_historique.cell(row_idx, col).number_format = '#,##0.00 €'

            # Ajuster les largeurs
            ws_historique.column_dimensions['A'].width = 20  # Marché
            ws_historique.column_dimensions['B'].width = 30  # Fournisseur
            ws_historique.column_dimensions['C'].width = 12  # Date SF
            ws_historique.column_dimensions['D'].width = 15  # N° Facture
            ws_historique.column_dimensions['E'].width = 40  # Libellé
            ws_historique.column_dimensions['F'].width = 15  # Montant SF
            ws_historique.column_dimensions['G'].width = 15  # Montant TTC
            ws_historique.column_dimensions['H'].width = 15  # N° Mandat
            ws_historique.column_dimensions['I'].width = 20  # Statut

            # Sauvegarder le fichier
            wb.save(filepath)
            return True

        except Exception as e:
            print(f"Erreur lors de l'export Excel : {e}")
            return False

    def export_suivi_financier_operation(
        self,
        code_operation: str,
        filepath: str,
        exercice_filter: Optional[str] = None,
        special_export: bool = False
    ) -> bool:
        """
        Génère un fichier Excel de suivi financier pour une opération spécifique.
        Format identique au modèle _Suivi_financier_op.xlsx.

        Args:
            code_operation: Code de l'opération (ex: "2024_1", "2024_17")
            filepath: Chemin du fichier Excel à créer
            exercice_filter: Exercice à filtrer (ex: "2024") ou "Tous"/None pour tout exporter
            special_export: Active la logique spécifique (ex: 2020_14G3P)

        Returns:
            True si succès, False sinon
        """
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from datetime import datetime

            wb = Workbook()
            wb.remove(wb.active)  # Supprimer la feuille par défaut

            # Styles
            header_font = Font(bold=True, size=11)
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            border_thin = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            # Filtrer les données pour l'opération
            operations_data = self.get_vision_operations()
            operation_info = next((op for op in operations_data if op['operation'] == code_operation), None)

            if not operation_info:
                print(f"Opération {code_operation} non trouvée")
                return False

            # Récupérer tous les marchés de l'opération
            marches_operation = operation_info['marches']

            # Charger les montants des commandes depuis la base de données
            # pour pouvoir afficher le montant de chaque BDC
            commandes_montants = {}
            if self.db:
                cur = self.db.conn.cursor()
                # Récupérer toutes les commandes pour les marchés de cette opération
                placeholders = ','.join('?' * len(marches_operation))
                query = f"""
                    SELECT num_commande, SUM(montant_ttc) as montant_total
                    FROM commandes
                    WHERE marche IN ({placeholders})
                    GROUP BY num_commande
                """
                cur.execute(query, marches_operation)
                for row in cur.fetchall():
                    if row['num_commande']:
                        commandes_montants[row['num_commande']] = row['montant_total']
                print(f"[DEBUG] Chargé {len(commandes_montants)} montants de commandes")

            # Charger les types de marchés
            marches_types = {}
            if self.db:
                cur = self.db.conn.cursor()
                placeholders = ','.join('?' * len(marches_operation))
                query = f"""
                    SELECT code_marche, type_marche
                    FROM marches
                    WHERE code_marche IN ({placeholders})
                """
                cur.execute(query, marches_operation)
                for row in cur.fetchall():
                    try:
                        type_marche = row['type_marche'] if row['type_marche'] else 'CLASSIQUE'
                    except (KeyError, IndexError):
                        type_marche = 'CLASSIQUE'
                    marches_types[row['code_marche']] = type_marche
                print(f"[DEBUG] Chargé {len(marches_types)} types de marchés")

            is_target_operation = special_export and code_operation == "2020_14G3P"

            # Charger les montants totaux des marchés (plafond pour BDC)
            marches_totaux = {}
            if self.db:
                for marche in marches_operation:
                    marches_totaux[marche] = self.db.get_montant_total_marche(marche)

            # ===============================
            # FEUILLE 1 : FINANCIER
            # ===============================
            ws_financier = wb.create_sheet("FINANCIER")

            # Titre de l'opération
            ws_financier.cell(1, 1, f"SUIVI FINANCIER OPÉRATION {code_operation}")
            ws_financier.cell(1, 1).font = Font(bold=True, size=14)
            ws_financier.merge_cells('A1:N1')

            # Informations opération
            ws_financier.cell(2, 1, f"Libellé: {operation_info.get('libelle', '')}")
            ws_financier.cell(3, 1, f"Fournisseurs: {operation_info.get('fournisseur', '')}")

            # En-têtes (ligne 5)
            headers = [
                "TYPE\nINTERVENTION",
                "DESIGNATION",
                "N° MARCHÉ",
                "NOM PRESTATAIRE",
                "N° BDC ou Tranche (TF ou TO)",
                "MONTANT TTC BDC\nou Tranche",
                "N°FACTURES\nFournisseur",
                "N°MANDAT",
                "Date de\nservice fait",
                "MONTANT FACTURE mandataire HT",
                "MONTANT FACTURE YC mandataire révision TTC",
                "MONTANT FACTURE PAR BDC YC révision-AF-RETGAR TTC",
                "MONTANT FACTURE PAR BDC HORS révision TTC",
                "MONTANT RESTANT SUR BDC Hors rev / engagement TTC"
            ]

            for col_idx, header in enumerate(headers, 1):
                cell = ws_financier.cell(5, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Organiser les données par (prestataire, tranche)
            # Collecter d'abord toutes les données
            from collections import defaultdict

            data_by_prestataire_tranche = defaultdict(list)

            for marche in marches_operation:
                df_marche = self.df_marches[self.df_marches.iloc[:, self.COL_MARCHE] == marche]
                if len(df_marche) == 0:
                    continue

                fournisseur = df_marche.iloc[0, self.COL_FOURNISSEUR] if not df_marche.iloc[0, self.COL_FOURNISSEUR] is pd.NA else ""
                tranche = df_marche.iloc[0, self.COL_TRANCHE] if not df_marche.iloc[0, self.COL_TRANCHE] is pd.NA else ""

                # Récupérer le type de marché
                type_marche = marches_types.get(marche, 'CLASSIQUE')

                # Formater la tranche
                tranche_libelle = ""
                if pd.notna(tranche):
                    try:
                        tranche_num = int(float(tranche))
                        tranche_libelle = "TF" if tranche_num == 0 else f"TO{tranche_num}"
                    except:
                        tranche_libelle = str(tranche)

                montant_initial_tranche = self.calculate_montant_initial_tranche(marche, tranche)

                for idx, row in df_marche.iterrows():
                    # Extraire le N° de commande (qui est dans COL_COMMANDE)
                    num_commande = row.iloc[self.COL_COMMANDE] if not pd.isna(row.iloc[self.COL_COMMANDE]) else ""
                    if is_target_operation:
                        exercice = self.extract_exercice_from_bdc(num_commande)
                        if exercice_filter and exercice_filter != "Tous" and exercice != exercice_filter:
                            continue

                    # Récupérer le montant de la commande depuis le dictionnaire
                    montant_commande = commandes_montants.get(num_commande, 0) if num_commande else 0

                    data_by_prestataire_tranche[(fournisseur, tranche_libelle)].append({
                        'marche': marche,
                        'type_marche': type_marche,
                        'fournisseur': fournisseur,
                        'tranche_libelle': tranche_libelle,
                        'montant_initial_tranche': montant_initial_tranche,
                        'num_commande': num_commande,
                        'montant_commande': montant_commande,
                        'num_facture': row.iloc[self.COL_FACTURE] if not pd.isna(row.iloc[self.COL_FACTURE]) else "",
                        'montant_sf': float(row.iloc[self.COL_MONTANT_SF]) if not pd.isna(row.iloc[self.COL_MONTANT_SF]) else 0,
                        'montant_ttc': float(row.iloc[self.COL_MONTANT_TTC]) if not pd.isna(row.iloc[self.COL_MONTANT_TTC]) else 0,
                        'num_mandat': row.iloc[self.COL_MANDAT] if not pd.isna(row.iloc[self.COL_MANDAT]) else "",
                        'date_sf': row.iloc[self.COL_DATE_SF] if not pd.isna(row.iloc[self.COL_DATE_SF]) else "",
                        'libelle': row.iloc[self.COL_LIBELLE] if not pd.isna(row.iloc[self.COL_LIBELLE]) else ""
                    })

            # Styles pour alternance et sous-totaux
            fill_gray = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
            fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            fill_subtotal = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            font_subtotal = Font(bold=True, size=11)
            border_thick = Border(
                left=Side(style='thick'),
                right=Side(style='thick'),
                top=Side(style='thick'),
                bottom=Side(style='thick')
            )

            # Écrire les données avec alternance et sous-totaux
            row_idx = 6
            current_prestataire = None
            color_index = 0

            restant_courant = None
            log_file = None
            if is_target_operation:
                import os
                log_path = os.path.join("run_logs", "export_2020_14G3P.log")
                log_file = open(log_path, "w", encoding="utf-8")
                log_file.write("ligne;marche;bdc;facture;plafond;montant_ttc;restant\n")

            for (fournisseur, tranche_libelle), factures in sorted(data_by_prestataire_tranche.items()):
                # Changer de couleur si changement de prestataire
                if current_prestataire != fournisseur:
                    current_prestataire = fournisseur
                    color_index += 1

                current_fill = fill_gray if color_index % 2 == 1 else fill_white
                if is_target_operation:
                    factures = sorted(
                        factures,
                        key=lambda x: (x['num_commande'] or "", x['num_facture'] or "")
                    )

                # Accumulateurs pour sous-totaux
                subtotal_montant_initial = 0
                subtotal_montant_bdc_ou_tranche = 0  # Accumule les valeurs réelles de la colonne 7
                subtotal_montant_ht = 0
                subtotal_montant_ttc = 0
                subtotal_par_bdc = 0
                first_row_of_group = row_idx
                previous_bdc = None  # Pour détecter les changements de BDC

                # Écrire chaque facture
                for facture in factures:
                    montant_ttc = facture['montant_ttc']
                    montant_ht = montant_ttc / 1.2 if montant_ttc > 0 else facture['montant_sf'] / 1.2

                    # Données de la ligne
                    ws_financier.cell(row_idx, 1, "TRAVAUX")
                    ws_financier.cell(row_idx, 2, facture['libelle'][:50])
                    ws_financier.cell(row_idx, 3, facture['marche'])
                    ws_financier.cell(row_idx, 4, facture['fournisseur'])
                    # Colonne 5: N° BDC (commande) ou Tranche
                    bdc_ou_tranche = facture['num_commande'] if facture['num_commande'] else facture['tranche_libelle']
                    ws_financier.cell(row_idx, 5, bdc_ou_tranche)
                    # Colonne 6: Montant TTC BDC (si BDC avec montant) ou Montant initial tranche (si tranche) ou Montant TTC facture (fallback)
                    if facture['num_commande'] and facture['montant_commande'] > 0:
                        # BDC avec montant trouvé
                        montant_bdc_ou_tranche = facture['montant_commande']
                    elif not facture['num_commande'] and row_idx == first_row_of_group:
                        # Tranche (première ligne du groupe)
                        montant_bdc_ou_tranche = facture['montant_initial_tranche']
                    elif facture['num_commande'] and facture['montant_commande'] == 0:
                        # BDC sans montant trouvé -> utiliser le montant TTC de la facture (fallback)
                        montant_bdc_ou_tranche = montant_ttc
                    else:
                        montant_bdc_ou_tranche = ""
                    ws_financier.cell(row_idx, 6, montant_bdc_ou_tranche)
                    ws_financier.cell(row_idx, 7, facture['num_facture'])
                    ws_financier.cell(row_idx, 8, facture['num_mandat'])
                    ws_financier.cell(row_idx, 9, facture['date_sf'])  # Date de service fait
                    ws_financier.cell(row_idx, 10, montant_ht)
                    ws_financier.cell(row_idx, 11, montant_ttc)
                    ws_financier.cell(row_idx, 12, montant_ttc)
                    ws_financier.cell(row_idx, 13, montant_ttc)

                    # Calculer le restant
                    type_marche = facture['type_marche']
                    current_bdc = facture['num_commande']

                    if is_target_operation:
                        montant_plafond = marches_totaux.get(facture['marche'], 0) or 0
                        if montant_plafond == 0:
                            montant_plafond = 4175726
                        montant_facture_ttc = montant_ttc
                        if restant_courant is None:
                            restant_courant = montant_plafond - montant_facture_ttc
                            if row_idx == first_row_of_group:
                                subtotal_montant_initial = montant_plafond
                        else:
                            restant_courant -= montant_facture_ttc
                        ws_financier.cell(row_idx, 14, restant_courant)
                        if log_file:
                            log_file.write(
                                f"{row_idx};{facture['marche']};{bdc_ou_tranche};{facture['num_facture']};"
                                f"{montant_plafond};{montant_facture_ttc};{restant_courant}\n"
                            )
                    else:
                        # Déterminer si on doit réinitialiser le restant
                        reset_restant = False
                        if row_idx == first_row_of_group:
                            # Première ligne du groupe : toujours réinitialiser
                            reset_restant = True
                        elif type_marche == 'BDC' and current_bdc and current_bdc != previous_bdc:
                            # Marché à BDC : réinitialiser quand le N° de BDC change
                            reset_restant = True
                        elif type_marche != 'BDC' and current_bdc and current_bdc != previous_bdc:
                            # Marché classique : réinitialiser quand le N° de BDC change
                            reset_restant = True

                        if reset_restant:
                            # Réinitialiser : utiliser le montant de la colonne 7 comme montant de départ
                            montant_depart = float(montant_bdc_ou_tranche) if montant_bdc_ou_tranche != "" else 0
                            restant = montant_depart - montant_ttc
                            ws_financier.cell(row_idx, 14, restant)
                            if row_idx == first_row_of_group:
                                subtotal_montant_initial = montant_depart
                        else:
                            # Continuer : déduire du restant précédent
                            prev_restant = ws_financier.cell(row_idx - 1, 14).value or 0
                            restant = prev_restant - montant_ttc
                            ws_financier.cell(row_idx, 14, restant)

                    # Mettre à jour le BDC précédent
                    previous_bdc = current_bdc

                    # Appliquer la couleur d'alternance
                    for col in range(1, 15):
                        ws_financier.cell(row_idx, col).fill = current_fill

                    # Format numérique
                    for col in [6, 10, 11, 12, 13, 14]:
                        ws_financier.cell(row_idx, col).number_format = '#,##0.00 €'

                    # Accumuler pour sous-totaux
                    subtotal_montant_ht += montant_ht
                    subtotal_montant_ttc += montant_ttc
                    subtotal_par_bdc += montant_ttc
                    # Accumuler le montant BDC ou Tranche (colonne 7) seulement si ce n'est pas une chaîne vide
                    if montant_bdc_ou_tranche != "":
                        subtotal_montant_bdc_ou_tranche += float(montant_bdc_ou_tranche) if montant_bdc_ou_tranche else 0

                    row_idx += 1

                # Insérer ligne de sous-total
                ws_financier.cell(row_idx, 1, "")
                ws_financier.cell(row_idx, 2, f"Sous-total Tranche {tranche_libelle} - {fournisseur}")
                ws_financier.cell(row_idx, 2).font = font_subtotal
                ws_financier.cell(row_idx, 3, "")
                ws_financier.cell(row_idx, 4, "")
                ws_financier.cell(row_idx, 5, "")
                ws_financier.cell(row_idx, 6, subtotal_montant_bdc_ou_tranche)  # Utilise l'accumulateur réel
                ws_financier.cell(row_idx, 7, "")
                ws_financier.cell(row_idx, 8, "")
                ws_financier.cell(row_idx, 9, "")  # Date de service fait (vide pour sous-total)
                ws_financier.cell(row_idx, 10, subtotal_montant_ht)
                ws_financier.cell(row_idx, 11, subtotal_montant_ttc)
                ws_financier.cell(row_idx, 12, subtotal_par_bdc)
                ws_financier.cell(row_idx, 13, subtotal_par_bdc)
                ws_financier.cell(row_idx, 14, subtotal_montant_bdc_ou_tranche - subtotal_par_bdc)  # Utilise l'accumulateur réel

                # Mise en forme du sous-total
                for col in range(1, 15):
                    cell = ws_financier.cell(row_idx, col)
                    cell.fill = fill_subtotal
                    cell.font = font_subtotal
                    cell.border = border_thick

                # Format numérique pour sous-total
                for col in [6, 10, 11, 12, 13, 14]:
                    ws_financier.cell(row_idx, col).number_format = '#,##0.00 €'

                row_idx += 1

            # Ajuster les largeurs
            ws_financier.column_dimensions['A'].width = 12  # Type
            ws_financier.column_dimensions['B'].width = 25  # Désignation
            ws_financier.column_dimensions['C'].width = 20  # N° Marché
            ws_financier.column_dimensions['D'].width = 25  # Nom Prestataire
            ws_financier.column_dimensions['E'].width = 15  # Tranche
            ws_financier.column_dimensions['F'].width = 15  # Montant BDC
            ws_financier.column_dimensions['G'].width = 15  # N° Facture
            ws_financier.column_dimensions['H'].width = 12  # N° Mandat
            ws_financier.column_dimensions['I'].width = 12  # Date de service fait
            for col in ['J', 'K', 'L', 'M', 'N']:
                ws_financier.column_dimensions[col].width = 14

            # ===============================
            # FEUILLE 2 : A JOUR (vue synthétique)
            # ===============================
            ws_ajour = wb.create_sheet("A jour")

            # Titre
            ws_ajour.cell(1, 1, f"OPÉRATION {code_operation}")
            ws_ajour.cell(1, 1).font = Font(bold=True, size=14)

            # En-têtes (ligne 4)
            headers_ajour = [
                "TYPE INTERVENTION",
                "DESIGNATION",
                "N° MARCHÉ",
                "NOM PRESTATAIRE",
                "N° BDC ou Tranche (TF ou TO)",
                "MONTANT TTC BDC ou Tranche",
                "MONTANT FACTURE PAR BDC YC révision TTC",
                "MONTANT RESTANT SUR BDC TTC"
            ]

            for col_idx, header in enumerate(headers_ajour, 1):
                cell = ws_ajour.cell(4, col_idx, header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border_thin

            # Données par tranche
            row_idx = 5
            for marche in marches_operation:
                df_marche = self.df_marches[self.df_marches.iloc[:, self.COL_MARCHE] == marche]
                if len(df_marche) == 0:
                    continue

                tranche = df_marche.iloc[0, self.COL_TRANCHE]
                tranche_libelle = ""
                if pd.notna(tranche):
                    try:
                        tranche_num = int(float(tranche))
                        tranche_libelle = "TF" if tranche_num == 0 else f"TO{tranche_num}"
                    except:
                        tranche_libelle = str(tranche)

                montant_initial = self.calculate_montant_initial_tranche(marche, tranche)
                service_fait = self.calculate_service_fait_tranche(marche, tranche)
                paye = self.calculate_paye_tranche(marche, tranche)
                restant = montant_initial - paye

                fournisseur = df_marche.iloc[0, self.COL_FOURNISSEUR]

                ws_ajour.cell(row_idx, 1, "TRAVAUX")
                ws_ajour.cell(row_idx, 2, "")
                ws_ajour.cell(row_idx, 3, marche)
                ws_ajour.cell(row_idx, 4, fournisseur if pd.notna(fournisseur) else "")
                ws_ajour.cell(row_idx, 5, tranche_libelle)
                ws_ajour.cell(row_idx, 6, montant_initial)
                ws_ajour.cell(row_idx, 7, paye)
                ws_ajour.cell(row_idx, 8, restant)

                # Format numérique
                for col in [6, 7, 8]:
                    ws_ajour.cell(row_idx, col).number_format = '#,##0.00 €'

                row_idx += 1

            # Totaux
            ws_ajour.cell(row_idx, 5, "TOTAL")
            ws_ajour.cell(row_idx, 5).font = Font(bold=True)
            for col in [6, 7, 8]:
                formula = f"=SUM({get_column_letter(col)}5:{get_column_letter(col)}{row_idx-1})"
                cell = ws_ajour.cell(row_idx, col, formula)
                cell.font = Font(bold=True)
                cell.number_format = '#,##0.00 €'

            # Ajuster largeurs
            for col_idx, width in enumerate([12, 25, 20, 25, 15, 18, 18, 18], 1):
                ws_ajour.column_dimensions[get_column_letter(col_idx)].width = width

            if log_file:
                log_file.close()

            # Sauvegarder
            wb.save(filepath)
            print(f"[OK] Export suivi financier operation {code_operation} : {filepath}")
            return True

        except Exception as e:
            print(f"[ERREUR] Erreur lors de l'export suivi financier : {e}")
            import traceback
            traceback.print_exc()
            return False
