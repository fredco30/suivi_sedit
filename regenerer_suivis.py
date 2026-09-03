#!/usr/bin/env python3
"""Regeneration en lot des suivis financiers deja diffuses.

Les fichiers produits avant le correctif d'agregation restent faux : ils
comptent deux fois l'enveloppe consommee (une ligne d'engagement *et* une ligne
de facture pour un meme BDC). Cette commande les regenere avec la regle « une
ligne par etat de BDC », onglets « Anomalies » et « Lignes neutralisées »
compris.

Exemples :
    python regenerer_suivis.py --tout
    python regenerer_suivis.py 2020_14G3P --sortie exports/
    python regenerer_suivis.py --tout --source "data_sources/factures*.xls"
    python regenerer_suivis.py --lister
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from typing import List, Optional

from marches_module import MarchesAnalyzer

DB_PAR_DEFAUT = "suivi_commandes.db"
SOURCE_PAR_DEFAUT = "database_sync"
CACHE_APPLICATION = "marches_cache.db"
# Cache dédié : charger des exports SEDIT en lot ne doit pas remplacer le cache
# de travail de l'application.
CACHE_REGENERATION = "regeneration_cache.db"


class BaseSuiviLectureSeule:
    """Acces lecture seule a `suivi_commandes.db` pour l'export hors interface.

    Reproduit les seules methodes de `Database` utilisees par `MarchesAnalyzer`,
    sans dependre de PyQt : la regeneration doit pouvoir tourner en lot.
    """

    def __init__(self, db_path: str = DB_PAR_DEFAUT):
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row

    def get_marche(self, code_marche: str):
        return self.conn.execute(
            "SELECT * FROM marches WHERE code_marche = ?", (code_marche,)
        ).fetchone()

    def get_tranches(self, code_marche: str) -> List[sqlite3.Row]:
        return list(self.conn.execute(
            "SELECT * FROM tranches WHERE code_marche = ? ORDER BY ordre", (code_marche,)
        ))

    def get_avenants(self, code_marche: str) -> List[sqlite3.Row]:
        return list(self.conn.execute(
            "SELECT * FROM avenants WHERE code_marche = ? ORDER BY numero_avenant",
            (code_marche,)
        ))

    def get_montant_total_marche(self, code_marche: str) -> float:
        """Enveloppe initiale : montant de base + tranches optionnelles + avenants."""
        marche = self.get_marche(code_marche)
        if not marche:
            return 0.0

        try:
            type_marche = marche["type_marche"] or "CLASSIQUE"
        except (KeyError, IndexError):
            type_marche = "CLASSIQUE"

        montant = marche["montant_initial_manuel"] or 0.0

        # Marchés CLASSIQUES : montant_initial_manuel porte la TF, la table
        # tranches porte les TOs.
        if type_marche == "CLASSIQUE":
            montant += sum((t["montant"] or 0.0) for t in self.get_tranches(code_marche))

        for avenant in self.get_avenants(code_marche):
            montant_avenant = avenant["montant"] or 0.0
            if avenant["type_modification"] == "Diminution":
                montant_avenant = -montant_avenant
            montant += montant_avenant

        return float(montant)

    def close(self) -> None:
        self.conn.close()


def cache_par_defaut(source) -> str:
    """Cache de l'application en lecture seule, cache dédié dès qu'on charge des exports."""
    sources = [source] if isinstance(source, str) else list(source or [])
    return CACHE_APPLICATION if sources == [SOURCE_PAR_DEFAUT] else CACHE_REGENERATION


def construire_analyzer(source, db_path: str, cache_path: Optional[str] = None) -> MarchesAnalyzer:
    cache_path = cache_path or cache_par_defaut(source)
    print(f"[CACHE] {cache_path}")
    analyzer = MarchesAnalyzer(
        source, database=BaseSuiviLectureSeule(db_path), use_cache=True, cache_path=cache_path
    )
    if not analyzer.load_data():
        raise RuntimeError(f"Chargement des données impossible depuis « {source} »")
    return analyzer


def lister_operations(analyzer: MarchesAnalyzer) -> List[str]:
    return [op["operation"] for op in analyzer.get_vision_operations()]


def regenerer(
    operations: Optional[List[str]] = None,
    sortie: str = ".",
    source=SOURCE_PAR_DEFAUT,
    db_path: str = DB_PAR_DEFAUT,
    exercice: Optional[str] = None,
    cache_path: Optional[str] = None,
) -> int:
    """Regenere les suivis demandes. Retourne le nombre d'echecs."""
    analyzer = construire_analyzer(source, db_path, cache_path)
    cibles = operations or lister_operations(analyzer)

    os.makedirs(sortie, exist_ok=True)
    echecs = 0

    for code_operation in cibles:
        filepath = os.path.join(sortie, f"suivi_financier_{code_operation}.xlsx")
        ok = analyzer.export_suivi_financier_operation(
            code_operation, filepath, exercice_filter=exercice, special_export=False
        )
        if not ok:
            echecs += 1
            print(f"[ECHEC] {code_operation}")

    print(f"\n{len(cibles) - echecs}/{len(cibles)} suivis régénérés dans « {sortie} »")
    return echecs


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("operations", nargs="*", help="Codes d'opération à régénérer")
    parser.add_argument("--tout", action="store_true", help="Régénérer toutes les opérations")
    parser.add_argument("--lister", action="store_true", help="Lister les opérations disponibles")
    parser.add_argument("--sortie", default=".", help="Répertoire de sortie (défaut : .)")
    parser.add_argument("--source", nargs="+", default=[SOURCE_PAR_DEFAUT],
                        help="Exports SEDIT à charger : fichiers, motifs (data_sources/factures*.xls) "
                             "ou répertoires. « database_sync » lit le cache existant.")
    parser.add_argument("--db", default=DB_PAR_DEFAUT, help="Base de suivi des commandes")
    parser.add_argument("--exercice", default=None, help="Filtrer sur un exercice (ex : 2024)")
    parser.add_argument("--cache", default=None,
                        help=f"Cache SQLite à utiliser (défaut : {CACHE_APPLICATION} en lecture "
                             f"du cache existant, sinon {CACHE_REGENERATION})")
    args = parser.parse_args(argv)

    if args.lister:
        analyzer = construire_analyzer(args.source, args.db, args.cache)
        for code_operation in lister_operations(analyzer):
            print(code_operation)
        return 0

    if not args.operations and not args.tout:
        parser.error("préciser des codes d'opération, --tout ou --lister")

    return regenerer(
        operations=args.operations or None,
        sortie=args.sortie,
        source=args.source,
        db_path=args.db,
        exercice=args.exercice,
        cache_path=args.cache,
    )


if __name__ == "__main__":
    sys.exit(main())
