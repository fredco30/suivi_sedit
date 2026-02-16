#!/usr/bin/env python3
"""
Script de migration pour ajouter les tables manquantes à la base de données.
"""

import sqlite3
import sys

def migrate_database(db_path='suivi_commandes.db'):
    """Ajoute les tables manquantes à la base de données."""

    print(f"Migration de la base de données: {db_path}")
    print("=" * 60)

    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        # Vérifier les tables existantes
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        existing_tables = [row[0] for row in cur.fetchall()]
        print(f"\nTables existantes: {', '.join(existing_tables)}")

        tables_created = []

        # Table Marchés
        if 'marches' not in existing_tables:
            print("\n✓ Création de la table 'marches'...")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS marches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code_marche TEXT UNIQUE NOT NULL,
                    libelle TEXT,
                    fournisseur TEXT,
                    montant_initial_manuel REAL,
                    date_notification TEXT,
                    date_debut TEXT,
                    date_fin_prevue TEXT,
                    notes TEXT,
                    type_marche TEXT DEFAULT 'CLASSIQUE',
                    last_update TEXT
                )
                """
            )
            tables_created.append('marches')
        else:
            print("\n- Table 'marches' existe déjà")
            # Vérifier si la colonne type_marche existe
            cur.execute("PRAGMA table_info(marches)")
            columns = [row[1] for row in cur.fetchall()]
            if 'type_marche' not in columns:
                print("  ✓ Ajout de la colonne 'type_marche'...")
                cur.execute("ALTER TABLE marches ADD COLUMN type_marche TEXT DEFAULT 'CLASSIQUE'")

        # Table Avenants
        if 'avenants' not in existing_tables:
            print("✓ Création de la table 'avenants'...")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS avenants (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code_marche TEXT NOT NULL,
                    numero_avenant INTEGER,
                    libelle TEXT,
                    montant REAL,
                    type_modification TEXT,
                    date_avenant TEXT,
                    motif TEXT,
                    last_update TEXT,
                    FOREIGN KEY (code_marche) REFERENCES marches(code_marche)
                )
                """
            )
            tables_created.append('avenants')
        else:
            print("- Table 'avenants' existe déjà")

        # Table Tranches
        if 'tranches' not in existing_tables:
            print("✓ Création de la table 'tranches'...")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS tranches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code_marche TEXT NOT NULL,
                    code_tranche TEXT NOT NULL,
                    libelle TEXT,
                    montant REAL,
                    ordre INTEGER,
                    last_update TEXT,
                    FOREIGN KEY (code_marche) REFERENCES marches(code_marche),
                    UNIQUE(code_marche, code_tranche)
                )
                """
            )
            tables_created.append('tranches')
        else:
            print("- Table 'tranches' existe déjà")

        # Table Import Tracking
        if 'import_tracking' not in existing_tables:
            print("✓ Création de la table 'import_tracking'...")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS import_tracking (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    filename TEXT UNIQUE NOT NULL,
                    file_path TEXT,
                    file_hash TEXT,
                    file_size INTEGER,
                    last_modified_date TEXT,
                    import_date TEXT,
                    import_type TEXT,
                    records_imported INTEGER,
                    status TEXT,
                    error_message TEXT
                )
                """
            )
            tables_created.append('import_tracking')
        else:
            print("- Table 'import_tracking' existe déjà")

        # Valider les changements
        conn.commit()

        print("\n" + "=" * 60)
        if tables_created:
            print(f"✅ Migration réussie ! Tables créées: {', '.join(tables_created)}")
        else:
            print("✅ Toutes les tables existent déjà. Aucune migration nécessaire.")

        # Afficher les tables finales
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        final_tables = [row[0] for row in cur.fetchall()]
        print(f"\nTables finales: {', '.join(final_tables)}")

        conn.close()
        return True

    except Exception as e:
        print(f"\n❌ Erreur lors de la migration: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    db_path = sys.argv[1] if len(sys.argv) > 1 else 'suivi_commandes.db'
    success = migrate_database(db_path)
    sys.exit(0 if success else 1)
