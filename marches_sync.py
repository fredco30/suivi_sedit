"""
Module de synchronisation SQLite pour les données de marchés
Gestion intelligente du cache et de la synchronisation incrémentale
"""

import sqlite3
import hashlib
import os
from datetime import datetime
from typing import Optional, Dict, List, Tuple
import pandas as pd


class MarchesSync:
    """Gestionnaire de synchronisation entre fichier Excel et base SQLite."""

    def __init__(self, db_path: str = "marches_cache.db"):
        """
        Initialise le gestionnaire de synchronisation.

        Args:
            db_path: Chemin vers la base de données SQLite
        """
        self.db_path = db_path
        self.conn = None
        self._init_database()

    def _init_database(self):
        """Initialise la structure de la base de données."""
        self.conn = sqlite3.connect(self.db_path)
        cursor = self.conn.cursor()

        # Table des lignes de factures (données brutes)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS lignes_factures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                hash_ligne TEXT UNIQUE NOT NULL,

                -- Colonnes du fichier Excel
                marche TEXT,
                fournisseur TEXT,
                libelle TEXT,
                date_sf TEXT,
                num_facture TEXT,
                montant_initial REAL,
                montant_sf REAL,
                montant_ttc REAL,
                num_mandat TEXT,
                tranche TEXT,
                commande TEXT,

                -- Métadonnées
                date_import TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                date_update TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Index pour améliorer les performances
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_marche ON lignes_factures(marche)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_fournisseur ON lignes_factures(fournisseur)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_hash ON lignes_factures(hash_ligne)")

        # Table de suivi de la synchronisation
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sync_info (
                id INTEGER PRIMARY KEY,
                fichier_path TEXT,
                last_modified REAL,
                file_size INTEGER,
                file_hash TEXT,
                last_sync TIMESTAMP,
                nb_lignes INTEGER,
                sync_status TEXT,
                sync_message TEXT
            )
        """)

        self.conn.commit()

    def _calculate_file_hash(self, filepath: str) -> str:
        """
        Calcule le hash MD5 d'un fichier.

        Args:
            filepath: Chemin vers le fichier

        Returns:
            Hash MD5 du fichier
        """
        hash_md5 = hashlib.md5()
        try:
            with open(filepath, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
        except Exception as e:
            print(f"Erreur lors du calcul du hash: {e}")
            return ""
        return hash_md5.hexdigest()

    def _calculate_row_hash(self, row: pd.Series) -> str:
        """
        Calcule le hash d'une ligne pour détecter les modifications.

        Args:
            row: Ligne pandas à hasher

        Returns:
            Hash MD5 de la ligne
        """
        # Concaténer les valeurs importantes de la ligne
        row_str = "|".join([str(v) for v in row.values])
        return hashlib.md5(row_str.encode()).hexdigest()

    def file_needs_sync(self, filepath: str) -> Tuple[bool, str]:
        """
        Vérifie si le fichier Excel a changé depuis la dernière synchronisation.
        Utilise un hash du contenu du fichier pour détecter les vrais changements.

        Args:
            filepath: Chemin vers le fichier Excel

        Returns:
            Tuple (needs_sync, reason)
        """
        if not os.path.exists(filepath):
            return True, "Fichier introuvable"

        # Récupérer les infos du fichier
        file_stat = os.stat(filepath)
        file_size = file_stat.st_size
        file_mtime = file_stat.st_mtime

        # Vérifier si on a déjà une synchronisation
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT last_modified, file_size, file_hash, nb_lignes FROM sync_info WHERE fichier_path = ? ORDER BY id DESC LIMIT 1",
            (filepath,)
        )
        result = cursor.fetchone()

        if not result:
            return True, "Première synchronisation"

        last_modified, last_size, last_hash, nb_lignes = result

        # Vérifier la taille (rapide)
        if file_size != last_size:
            return True, f"Taille du fichier changée ({last_size} → {file_size} octets)"

        if nb_lignes == 0:
            return True, "Base vide"

        # Si la date a changé, vérifier le hash pour détecter un vrai changement
        if abs(file_mtime - last_modified) > 1:  # Tolérance de 1 seconde
            # Calculer le hash du fichier actuel
            current_hash = self._calculate_file_hash(filepath)

            if current_hash != last_hash:
                return True, f"Contenu du fichier modifié"
            else:
                # Date changée mais contenu identique (fichier ouvert sans modification)
                return False, "Fichier inchangé (date modifiée mais contenu identique)"

        return False, "Fichier inchangé"

    def sync_from_excel(self, filepath: str, df: pd.DataFrame, force: bool = False) -> Dict:
        """
        Synchronise les données depuis un DataFrame pandas vers SQLite.

        Args:
            filepath: Chemin vers le fichier Excel source
            df: DataFrame contenant les données
            force: Forcer la synchronisation même si le fichier n'a pas changé

        Returns:
            Dictionnaire avec les statistiques de synchronisation
        """
        stats = {
            'nb_inserted': 0,
            'nb_updated': 0,
            'nb_unchanged': 0,
            'nb_deleted': 0,
            'duration': 0,
            'status': 'success'
        }

        start_time = datetime.now()

        try:
            # Vérifier si synchronisation nécessaire
            if not force:
                needs_sync, reason = self.file_needs_sync(filepath)
                if not needs_sync:
                    stats['status'] = 'skipped'
                    stats['message'] = reason
                    return stats

            cursor = self.conn.cursor()

            # Récupérer les hash existants
            cursor.execute("SELECT hash_ligne FROM lignes_factures")
            existing_hashes = set(row[0] for row in cursor.fetchall())

            # Préparer les nouvelles données
            new_hashes = set()
            rows_to_insert = []
            rows_to_update = []

            for idx, row in df.iterrows():
                row_hash = self._calculate_row_hash(row)
                new_hashes.add(row_hash)

                # Extraire les données (adapter selon vos colonnes Excel)
                row_data = {
                    'hash_ligne': row_hash,
                    'marche': str(row.get('marche', '')) if pd.notna(row.get('marche')) else None,
                    'fournisseur': str(row.get('fournisseur', '')) if pd.notna(row.get('fournisseur')) else None,
                    'libelle': str(row.get('libelle', '')) if pd.notna(row.get('libelle')) else None,
                    'date_sf': str(row.get('date_sf', '')) if pd.notna(row.get('date_sf')) else None,
                    'num_facture': str(row.get('num_facture', '')) if pd.notna(row.get('num_facture')) else None,
                    'montant_initial': float(row.get('montant_initial', 0)) if pd.notna(row.get('montant_initial')) else 0.0,
                    'montant_sf': float(row.get('montant_sf', 0)) if pd.notna(row.get('montant_sf')) else 0.0,
                    'montant_ttc': float(row.get('montant_ttc', 0)) if pd.notna(row.get('montant_ttc')) else 0.0,
                    'num_mandat': str(row.get('num_mandat', '')) if pd.notna(row.get('num_mandat')) else None,
                    'tranche': str(row.get('tranche', '')) if pd.notna(row.get('tranche')) else None,
                    'commande': str(row.get('commande', '')) if pd.notna(row.get('commande')) else None,
                }

                if row_hash in existing_hashes:
                    stats['nb_unchanged'] += 1
                else:
                    rows_to_insert.append(row_data)

            # Insertion des nouvelles lignes
            if rows_to_insert:
                cursor.executemany("""
                    INSERT OR REPLACE INTO lignes_factures
                    (hash_ligne, marche, fournisseur, libelle, date_sf, num_facture,
                     montant_initial, montant_sf, montant_ttc, num_mandat, tranche, commande)
                    VALUES (:hash_ligne, :marche, :fournisseur, :libelle, :date_sf, :num_facture,
                            :montant_initial, :montant_sf, :montant_ttc, :num_mandat, :tranche, :commande)
                """, rows_to_insert)
                stats['nb_inserted'] = len(rows_to_insert)

            # Supprimer les lignes qui n'existent plus dans le fichier
            hashes_to_delete = existing_hashes - new_hashes
            if hashes_to_delete:
                cursor.execute(
                    f"DELETE FROM lignes_factures WHERE hash_ligne IN ({','.join('?' * len(hashes_to_delete))})",
                    list(hashes_to_delete)
                )
                stats['nb_deleted'] = len(hashes_to_delete)

            # Mettre à jour sync_info
            # Gérer le cas où le filepath est factice (sync depuis database)
            if os.path.exists(filepath):
                file_stat = os.stat(filepath)
                file_mtime = file_stat.st_mtime
                file_size = file_stat.st_size
                file_hash = self._calculate_file_hash(filepath)
            else:
                # Filepath factice (sync depuis database)
                file_mtime = datetime.now().timestamp()
                file_size = 0
                file_hash = "database_sync"

            cursor.execute("""
                INSERT INTO sync_info (fichier_path, last_modified, file_size, file_hash, last_sync, nb_lignes, sync_status, sync_message)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                filepath,
                file_mtime,
                file_size,
                file_hash,
                datetime.now(),
                len(df),
                'success',
                f"Sync OK: {stats['nb_inserted']} insérées, {stats['nb_deleted']} supprimées, {stats['nb_unchanged']} inchangées"
            ))

            self.conn.commit()

        except Exception as e:
            self.conn.rollback()
            stats['status'] = 'error'
            stats['message'] = str(e)
            print(f"Erreur lors de la synchronisation: {e}")

        stats['duration'] = (datetime.now() - start_time).total_seconds()
        return stats

    def load_to_dataframe(self, marche_filter: Optional[str] = None) -> pd.DataFrame:
        """
        Charge les données depuis SQLite vers un DataFrame pandas.

        Args:
            marche_filter: Filtre optionnel sur un marché spécifique

        Returns:
            DataFrame avec les données
        """
        query = "SELECT * FROM lignes_factures"
        params = []

        if marche_filter:
            query += " WHERE marche = ?"
            params.append(marche_filter)

        query += " ORDER BY marche, date_sf"

        return pd.read_sql_query(query, self.conn, params=params)

    def get_sync_status(self, filepath: str) -> Optional[Dict]:
        """
        Récupère les informations de la dernière synchronisation.

        Args:
            filepath: Chemin vers le fichier Excel

        Returns:
            Dictionnaire avec les infos de synchronisation ou None
        """
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT last_sync, nb_lignes, sync_status, sync_message
            FROM sync_info
            WHERE fichier_path = ?
            ORDER BY id DESC
            LIMIT 1
        """, (filepath,))

        result = cursor.fetchone()
        if not result:
            return None

        return {
            'last_sync': result[0],
            'nb_lignes': result[1],
            'sync_status': result[2],
            'sync_message': result[3]
        }

    def clear_cache(self):
        """Vide complètement le cache SQLite."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM lignes_factures")
        cursor.execute("DELETE FROM sync_info")
        self.conn.commit()

    def close(self):
        """Ferme la connexion à la base de données."""
        if self.conn:
            self.conn.close()

    def __del__(self):
        """Destructeur pour fermer la connexion."""
        self.close()
