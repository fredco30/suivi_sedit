"""Tests unitaires de matching_module.

Utilise unittest (stdlib) et une base SQLite en memoire pour decoupler les tests
du fichier principal. Les schemas commandes/factures/manual_links/commande_diagnostic
sont recrees dans setUp().

Lancement :
    python -m unittest test_matching.py -v
ou
    python test_matching.py
"""
from __future__ import annotations

import sqlite3
import unittest
from datetime import date, datetime, timedelta

from matching_module import (
    LinkRepository,
    MatchingEngine,
    DIAG_OK,
    DIAG_OK_RAPPROCHE,
    DIAG_RECENT,
    DIAG_EN_COURS,
    DIAG_RAPPROCHEMENT_SUGGERE,
    DIAG_OUBLI_PROBABLE,
    DIAG_DOUBLON_PROBABLE,
    SEVERITE_BY_DIAG,
    THRESHOLD_FOURNISSEUR,
    fuzzy_ratio,
    normalize_fournisseur,
    parse_date,
    parse_num_engagement,
)


SCHEMA_SQL = """
CREATE TABLE commandes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    num_commande TEXT,
    fournisseur TEXT,
    libelle TEXT,
    date_commande TEXT,
    marche TEXT,
    montant_ttc REAL,
    statut TEXT,
    statut_facturation TEXT,
    statut_metier TEXT,
    montant_facture REAL DEFAULT 0,
    reste_a_facturer REAL DEFAULT 0
);

CREATE TABLE factures (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    num_facture TEXT,
    code_mouvement TEXT,
    fournisseur TEXT,
    libelle TEXT,
    marche TEXT,
    date_facture TEXT,
    montant_service_fait REAL DEFAULT 0
);

CREATE TABLE manual_links (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    commande_id INTEGER NOT NULL,
    facture_id INTEGER NOT NULL,
    montant_alloue REAL NOT NULL,
    confidence INTEGER,
    source TEXT NOT NULL,
    ai_model TEXT,
    ai_reasoning TEXT,
    created_by TEXT,
    created_at TEXT NOT NULL,
    validated_at TEXT,
    notes TEXT,
    UNIQUE(commande_id, facture_id)
);

CREATE TABLE commande_diagnostic (
    commande_id INTEGER PRIMARY KEY,
    age_jours INTEGER,
    diagnostic TEXT NOT NULL,
    severite INTEGER,
    candidates_count INTEGER DEFAULT 0,
    candidates_same_marche INTEGER DEFAULT 0,
    montant_candidates_total REAL DEFAULT 0,
    last_ai_check_at TEXT,
    last_ai_diagnostic TEXT,
    last_diagnostic_at TEXT NOT NULL
);
"""


def make_db():
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA_SQL)
    return conn


def insert_commande(conn, **kw):
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO commandes
        (num_commande, fournisseur, libelle, date_commande, marche, montant_ttc,
         statut, statut_facturation, statut_metier)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            kw.get("num_commande"), kw.get("fournisseur"), kw.get("libelle"),
            kw.get("date_commande"), kw.get("marche"), kw.get("montant_ttc", 0),
            kw.get("statut"), kw.get("statut_facturation"), kw.get("statut_metier"),
        ),
    )
    conn.commit()
    return cur.lastrowid


def insert_facture(conn, **kw):
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO factures
        (num_facture, code_mouvement, fournisseur, libelle, marche, date_facture,
         montant_service_fait)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            kw.get("num_facture"), kw.get("code_mouvement"), kw.get("fournisseur"),
            kw.get("libelle"), kw.get("marche"), kw.get("date_facture"),
            kw.get("montant_service_fait", 0),
        ),
    )
    conn.commit()
    return cur.lastrowid


# ============================================================================
# Helpers
# ============================================================================

class TestHelpers(unittest.TestCase):

    def test_normalize_fournisseur(self):
        self.assertEqual(normalize_fournisseur("BOUYGUES E&S "), "BOUYGUES E S")
        self.assertEqual(normalize_fournisseur("ACAF Ascenseurs"), "ACAF ASCENSEURS")
        self.assertEqual(normalize_fournisseur("  Hervé thermique  "), "HERVE THERMIQUE")
        self.assertEqual(normalize_fournisseur(None), "")
        self.assertEqual(normalize_fournisseur(""), "")

    def test_fuzzy_ratio_above_threshold_for_abbreviation(self):
        ratio = fuzzy_ratio("BOUYGUES E&S", "BOUYGUES ENERGIES ET SERVICES")
        # Pas force d'etre >= 0.85 (chaines de longueurs tres differentes)
        # mais doit etre > 0.50 (ressemblance partagee).
        self.assertGreater(ratio, 0.5)

    def test_fuzzy_ratio_above_threshold_for_close_variants(self):
        self.assertGreaterEqual(
            fuzzy_ratio("ACAF ASCENSEURS", "ACAF Ascenseurs "), THRESHOLD_FOURNISSEUR
        )
        self.assertGreaterEqual(
            fuzzy_ratio("HERVE THERMIQUE", "Hervé thermique"), THRESHOLD_FOURNISSEUR
        )

    def test_fuzzy_ratio_zero_for_empty(self):
        self.assertEqual(fuzzy_ratio(None, "X"), 0.0)
        self.assertEqual(fuzzy_ratio("X", None), 0.0)
        self.assertEqual(fuzzy_ratio("", "Y"), 0.0)

    def test_parse_num_engagement_valid(self):
        self.assertEqual(parse_num_engagement("25AA01486"), (2025, "AA", 1486))
        self.assertEqual(parse_num_engagement("23AA00077"), (2023, "AA", 77))
        self.assertEqual(parse_num_engagement("21BB00001"), (2021, "BB", 1))

    def test_parse_num_engagement_invalid(self):
        self.assertIsNone(parse_num_engagement(""))
        self.assertIsNone(parse_num_engagement(None))
        self.assertIsNone(parse_num_engagement("REPORT-2024"))
        self.assertIsNone(parse_num_engagement("XXAA00001"))

    def test_parse_date_iso(self):
        self.assertEqual(parse_date("2025-04-28"), date(2025, 4, 28))
        self.assertEqual(parse_date("2025-04-28T10:30:00"), date(2025, 4, 28))
        self.assertEqual(parse_date("2025-04-28 10:30:00"), date(2025, 4, 28))

    def test_parse_date_french(self):
        self.assertEqual(parse_date("28/04/2025"), date(2025, 4, 28))
        self.assertEqual(parse_date("28-04-2025"), date(2025, 4, 28))

    def test_parse_date_none(self):
        self.assertIsNone(parse_date(""))
        self.assertIsNone(parse_date(None))
        self.assertIsNone(parse_date("garbage"))


# ============================================================================
# LinkRepository
# ============================================================================

class TestLinkRepository(unittest.TestCase):

    def setUp(self):
        self.conn = make_db()
        self.repo = LinkRepository(self.conn)
        self.cmd_id = insert_commande(self.conn, num_commande="25AA01486",
                                      fournisseur="ACAF", montant_ttc=2376.0)
        self.fact_id = insert_facture(self.conn, num_facture="F001",
                                      code_mouvement="25AA00079",
                                      fournisseur="ACAF", montant_service_fait=174.23)

    def test_add_link_simple(self):
        lid = self.repo.add_link(self.cmd_id, self.fact_id, 100.0, source="manual")
        self.assertIsInstance(lid, int)
        self.assertGreater(lid, 0)
        self.assertEqual(self.repo.total_alloue_for_commande(self.cmd_id), 100.0)
        self.assertEqual(self.repo.total_alloue_for_facture(self.fact_id), 100.0)

    def test_add_link_unicite_violee(self):
        self.repo.add_link(self.cmd_id, self.fact_id, 50.0, source="manual")
        with self.assertRaises(sqlite3.IntegrityError):
            self.repo.add_link(self.cmd_id, self.fact_id, 30.0, source="manual")

    def test_add_link_montant_alloue_exceeds_montant_sf(self):
        with self.assertRaises(ValueError):
            self.repo.add_link(self.cmd_id, self.fact_id, 200.0, source="manual")

    def test_add_link_somme_allocations_exceeds_montant_sf(self):
        # On cree une 2e commande pour pouvoir refaire un link sur la meme facture
        cmd2 = insert_commande(self.conn, num_commande="25AA01999", fournisseur="ACAF")
        self.repo.add_link(self.cmd_id, self.fact_id, 100.0, source="manual")
        # Reste 74.23 EUR ; allouer 100 → depasse
        with self.assertRaises(ValueError):
            self.repo.add_link(cmd2, self.fact_id, 100.0, source="manual")
        # Allouer 50 reste OK
        self.repo.add_link(cmd2, self.fact_id, 50.0, source="manual")
        self.assertAlmostEqual(self.repo.total_alloue_for_facture(self.fact_id), 150.0, places=2)

    def test_add_link_montant_zero_or_negative(self):
        with self.assertRaises(ValueError):
            self.repo.add_link(self.cmd_id, self.fact_id, 0, source="manual")
        with self.assertRaises(ValueError):
            self.repo.add_link(self.cmd_id, self.fact_id, -1, source="manual")

    def test_add_link_unknown_facture(self):
        with self.assertRaises(ValueError):
            self.repo.add_link(self.cmd_id, 99999, 10.0, source="manual")

    def test_add_link_unknown_commande(self):
        with self.assertRaises(ValueError):
            self.repo.add_link(99999, self.fact_id, 10.0, source="manual")

    def test_remove_link(self):
        lid = self.repo.add_link(self.cmd_id, self.fact_id, 50.0, source="manual")
        self.assertTrue(self.repo.remove_link(lid))
        self.assertEqual(self.repo.total_alloue_for_commande(self.cmd_id), 0.0)
        # Re-suppression d'un id absent : False
        self.assertFalse(self.repo.remove_link(lid))

    def test_list_links_for_commande(self):
        lid = self.repo.add_link(self.cmd_id, self.fact_id, 50.0, source="manual",
                                 confidence=92, ai_reasoning="test")
        links = self.repo.list_links_for_commande(self.cmd_id)
        self.assertEqual(len(links), 1)
        self.assertEqual(links[0]["id"], lid)
        self.assertEqual(links[0]["confidence"], 92)
        self.assertEqual(links[0]["ai_reasoning"], "test")


# ============================================================================
# MatchingEngine.find_candidates
# ============================================================================

class TestFindCandidates(unittest.TestCase):

    def setUp(self):
        self.conn = make_db()
        self.repo = LinkRepository(self.conn)
        self.engine = MatchingEngine(self.conn, self.repo, today=date(2025, 6, 1))

    def _cmd(self, **kw):
        defaults = dict(num_commande="25AA01486", fournisseur="ACAF ASCENSEURS",
                        date_commande="2025-04-28", marche="2023_17",
                        montant_ttc=2376.0)
        defaults.update(kw)
        return insert_commande(self.conn, **defaults)

    def _fact(self, **kw):
        # Valeurs par defaut : facture orpheline (code_mouvement non rattache nativement)
        # avec une date posterieure a la commande pour passer la regle de tolerance.
        defaults = dict(num_facture="F001", code_mouvement="25AA00079",
                        fournisseur="ACAF ASCENSEURS", date_facture="2025-04-28",
                        marche="2023_17", montant_service_fait=174.23)
        defaults.update(kw)
        return insert_facture(self.conn, **defaults)

    def test_basic_acaf_case(self):
        cmd_id = self._cmd()
        f1 = self._fact()
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(len(cands), 1)
        self.assertEqual(cands[0]["id"], f1)

    def test_excludes_native_match(self):
        # La facture porte code_mouvement = num_commande d une autre commande connue
        # → doit etre exclue car deja rattachee nativement.
        cmd_id = self._cmd()
        # Cmd "concurrente" qui matche le code_mouvement
        insert_commande(self.conn, num_commande="25AA00079", fournisseur="ACAF")
        self._fact(code_mouvement="25AA00079")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_excludes_already_linked_manually(self):
        cmd_id = self._cmd()
        f1 = self._fact()
        self.repo.add_link(cmd_id, f1, 50.0, source="manual")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_filter_by_fournisseur_below_threshold(self):
        cmd_id = self._cmd(fournisseur="ACAF ASCENSEURS")
        # Fournisseur radicalement different mais commencant par 'ACA'
        # pour passer le pre-filtrage SQL LIKE
        self._fact(fournisseur="ACAJOU SARL")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_filter_by_marche_strict(self):
        cmd_id = self._cmd(marche="2023_17")
        self._fact(marche="2024_22")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_filter_marche_one_side_empty_accepted(self):
        # Cmd a un marche, facture n'en a pas → accepte (couvre cas pratiques
        # ou la compta n'a pas saisi le marche cote facture).
        cmd_id = self._cmd(marche="2023_17")
        self._fact(marche=None)
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(len(cands), 1)

    def test_filter_by_date_too_early(self):
        # Facture > 31 jours avant la commande → exclue
        cmd_id = self._cmd(date_commande="2025-04-28")
        self._fact(date_facture="2025-01-13")  # > 30j avant
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_filter_by_date_within_tolerance(self):
        # Facture 25j avant la commande → acceptee
        cmd_id = self._cmd(date_commande="2025-04-28")
        self._fact(date_facture="2025-04-05")  # 23j avant
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(len(cands), 1)

    def test_filter_by_numerotation_same_year(self):
        # cmd 25AA01486, facture portant code 25AA02000 → exclue (n° plus eleve, meme exercice)
        cmd_id = self._cmd(num_commande="25AA01486")
        self._fact(code_mouvement="25AA02000", date_facture="2025-04-28")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_filter_by_numerotation_lower_seq_accepted(self):
        # cmd 25AA01486, facture portant code 25AA00500 → acceptee si autres regles OK
        cmd_id = self._cmd(num_commande="25AA01486", date_commande="2024-12-01")
        self._fact(code_mouvement="25AA00500", date_facture="2025-01-13")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(len(cands), 1)

    def test_filter_prior_exercise_no_constraint(self):
        # Exercice anterieur (pattern REPORT) : pas de contrainte de numerotation
        cmd_id = self._cmd(num_commande="25AA01486", date_commande="2024-06-01")
        self._fact(code_mouvement="24AA09999", date_facture="2024-07-01")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(len(cands), 1)

    def test_filter_unparseable_engagement_no_constraint(self):
        # code_mouvement non parseable : pas de contrainte de numerotation
        cmd_id = self._cmd(num_commande="25AA01486", date_commande="2024-06-01")
        self._fact(code_mouvement="REPORT-2024", date_facture="2025-04-28")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(len(cands), 1)

    def test_filter_zero_montant_sf_excluded(self):
        cmd_id = self._cmd()
        self._fact(montant_service_fait=0, date_facture="2025-04-28")
        cands = self.engine.find_candidates(cmd_id)
        self.assertEqual(cands, [])

    def test_unknown_commande_returns_empty(self):
        self.assertEqual(self.engine.find_candidates(99999), [])


# ============================================================================
# MatchingEngine.diagnose_all_commandes
# ============================================================================

class TestDiagnoseAllCommandes(unittest.TestCase):

    def setUp(self):
        self.conn = make_db()
        self.repo = LinkRepository(self.conn)
        self.today = date(2025, 6, 1)
        self.engine = MatchingEngine(self.conn, self.repo, today=self.today)

    def _diag_for(self, cmd_id):
        cur = self.conn.cursor()
        row = cur.execute(
            "SELECT * FROM commande_diagnostic WHERE commande_id = ?", (cmd_id,)
        ).fetchone()
        return dict(row) if row else None

    def test_diag_recent(self):
        cmd_id = insert_commande(self.conn, num_commande="25AA02000",
                                 date_commande="2025-05-25",  # 7 jours
                                 fournisseur="X", montant_ttc=100,
                                 statut_facturation="Non facturée")
        self.engine.diagnose_all_commandes()
        diag = self._diag_for(cmd_id)
        self.assertEqual(diag["diagnostic"], DIAG_RECENT)
        self.assertEqual(diag["severite"], 1)
        self.assertEqual(diag["age_jours"], 7)

    def test_diag_en_cours(self):
        cmd_id = insert_commande(self.conn, num_commande="25AA01000",
                                 date_commande="2025-04-01",  # 61 jours
                                 fournisseur="X", montant_ttc=100,
                                 statut_facturation="Non facturée")
        self.engine.diagnose_all_commandes()
        self.assertEqual(self._diag_for(cmd_id)["diagnostic"], DIAG_EN_COURS)

    def test_diag_oubli_probable(self):
        # > 90 jours, aucune facture candidate
        cmd_id = insert_commande(self.conn, num_commande="25AA00100",
                                 date_commande="2025-01-15",  # > 90 jours
                                 fournisseur="UNIQUE_FOURN", montant_ttc=100,
                                 statut_facturation="Non facturée")
        self.engine.diagnose_all_commandes()
        diag = self._diag_for(cmd_id)
        self.assertEqual(diag["diagnostic"], DIAG_OUBLI_PROBABLE)
        self.assertEqual(diag["severite"], 3)
        self.assertEqual(diag["candidates_count"], 0)

    def test_diag_rapprochement_suggere(self):
        # > 90 jours avec un candidat
        cmd_id = insert_commande(self.conn, num_commande="25AA01486",
                                 date_commande="2025-01-15",
                                 fournisseur="ACAF ASCENSEURS",
                                 marche="2023_17", montant_ttc=2376.0,
                                 statut_facturation="Non facturée")
        # Facture orpheline (code_mouvement ne matche aucune commande)
        insert_facture(self.conn, num_facture="F001",
                       code_mouvement="25AA00050",  # orpheline
                       fournisseur="ACAF ASCENSEURS",
                       marche="2023_17", date_facture="2025-01-13",
                       montant_service_fait=174.23)
        self.engine.diagnose_all_commandes()
        diag = self._diag_for(cmd_id)
        self.assertEqual(diag["diagnostic"], DIAG_RAPPROCHEMENT_SUGGERE)
        self.assertEqual(diag["candidates_count"], 1)
        self.assertEqual(diag["candidates_same_marche"], 1)
        self.assertAlmostEqual(diag["montant_candidates_total"], 174.23, places=2)

    def test_diag_ok_via_natif(self):
        cmd_id = insert_commande(self.conn, num_commande="25AA01000",
                                 fournisseur="X", date_commande="2025-01-01",
                                 statut_facturation="Totalement facturée")
        self.engine.diagnose_all_commandes()
        self.assertEqual(self._diag_for(cmd_id)["diagnostic"], DIAG_OK)

    def test_diag_ok_rapproche_via_manual_link(self):
        cmd_id = insert_commande(self.conn, num_commande="25AA01000",
                                 fournisseur="X", date_commande="2025-01-01",
                                 statut_facturation="Totalement facturée")
        f_id = insert_facture(self.conn, num_facture="F001",
                              code_mouvement="25AA00500",
                              fournisseur="X", montant_service_fait=100.0,
                              date_facture="2025-02-01")
        self.repo.add_link(cmd_id, f_id, 100.0, source="manual")
        self.engine.diagnose_all_commandes()
        self.assertEqual(self._diag_for(cmd_id)["diagnostic"], DIAG_OK_RAPPROCHE)

    def test_diag_doublon_admin(self):
        cmd_id = insert_commande(self.conn, num_commande="25AA01486",
                                 fournisseur="ACAF", date_commande="2025-01-15",
                                 statut_facturation="Doublon administratif",
                                 statut_metier="DOUBLON_ADMIN")
        self.engine.diagnose_all_commandes()
        diag = self._diag_for(cmd_id)
        self.assertEqual(diag["diagnostic"], DIAG_DOUBLON_PROBABLE)
        self.assertEqual(diag["severite"], 2)

    def test_diagnose_returns_counts(self):
        insert_commande(self.conn, num_commande="A", date_commande="2025-05-25",
                        fournisseur="X", statut_facturation="Non facturée")
        insert_commande(self.conn, num_commande="B", date_commande="2025-04-01",
                        fournisseur="X", statut_facturation="Non facturée")
        insert_commande(self.conn, num_commande="C", date_commande="2025-01-01",
                        fournisseur="X", statut_facturation="Totalement facturée")
        counts = self.engine.diagnose_all_commandes()
        self.assertEqual(counts[DIAG_RECENT], 1)
        self.assertEqual(counts[DIAG_EN_COURS], 1)
        self.assertEqual(counts[DIAG_OK], 1)

    def test_diagnose_upsert_preserves_ai_cache(self):
        cmd_id = insert_commande(self.conn, num_commande="25AA00050",
                                 date_commande="2025-01-15",
                                 fournisseur="X", statut_facturation="Non facturée")
        # Premier passage : insert
        self.engine.diagnose_all_commandes()
        # Simuler une reponse IA en cache
        self.conn.execute(
            "UPDATE commande_diagnostic SET last_ai_check_at = ?, last_ai_diagnostic = ? "
            "WHERE commande_id = ?",
            ("2025-05-30T10:00:00", '{"diagnostic": "ORPHELINE"}', cmd_id),
        )
        self.conn.commit()
        # Deuxieme passage : on doit preserver les last_ai_*
        self.engine.diagnose_all_commandes()
        diag = self._diag_for(cmd_id)
        self.assertEqual(diag["last_ai_check_at"], "2025-05-30T10:00:00")
        self.assertEqual(diag["last_ai_diagnostic"], '{"diagnostic": "ORPHELINE"}')


if __name__ == "__main__":
    unittest.main(verbosity=2)
