"""Tests de l'interface : saisie des enveloppes et regeneration en lot.

Tourne sans affichage grace au greffon Qt « offscreen ». Les tests sont ignores
si PyQt5 n'est pas installe -- la logique metier, elle, est couverte par
test_suivi_financier.py sans dependance graphique.

Lancement :
    python -m unittest test_interface.py -v
"""
from __future__ import annotations

import glob
import os
import shutil
import sqlite3
import tempfile
import unittest
import unittest.mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

SCHEMA_MARCHES = """
CREATE TABLE marches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code_marche TEXT UNIQUE NOT NULL,
    libelle TEXT, fournisseur TEXT, montant_initial_manuel REAL,
    date_notification TEXT, date_debut TEXT, date_fin_prevue TEXT,
    notes TEXT, last_update TEXT, type_marche TEXT DEFAULT 'CLASSIQUE'
);
CREATE TABLE tranches (
    id INTEGER PRIMARY KEY AUTOINCREMENT, code_marche TEXT NOT NULL,
    code_tranche TEXT NOT NULL, libelle TEXT, montant REAL, ordre INTEGER,
    last_update TEXT, UNIQUE(code_marche, code_tranche)
);
CREATE TABLE avenants (
    id INTEGER PRIMARY KEY AUTOINCREMENT, code_marche TEXT NOT NULL,
    numero_avenant INTEGER, libelle TEXT, montant REAL, type_modification TEXT,
    date_avenant TEXT, motif TEXT, last_update TEXT
);
CREATE TABLE commandes (
    id INTEGER PRIMARY KEY AUTOINCREMENT, exercice TEXT, num_commande TEXT,
    fournisseur TEXT, libelle TEXT, date_commande TEXT, marche TEXT,
    montant_ttc REAL, statut TEXT
);
"""


def _module_application():
    """Charge le module principal, dont le nom de fichier n'est pas importable."""
    import importlib.util
    import sys

    if "appmod" not in sys.modules:
        spec = importlib.util.spec_from_file_location(
            "appmod", "suivi_commandes_factures_marches_FinaàGarder.py"
        )
        module = importlib.util.module_from_spec(spec)
        sys.modules["appmod"] = module
        spec.loader.exec_module(module)
    return sys.modules["appmod"]


def _application():
    """Instance QApplication unique du processus de test."""
    from PyQt5.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


class BaseFactice:
    """Le strict nécessaire de `Database` pour le dialogue des enveloppes."""

    def __init__(self, db_path):
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row

    def get_marche(self, code_marche):
        return self.conn.execute(
            "SELECT * FROM marches WHERE code_marche = ?", (code_marche,)
        ).fetchone()

    def get_tranches(self, code_marche):
        return list(self.conn.execute(
            "SELECT * FROM tranches WHERE code_marche = ?", (code_marche,)
        ))

    def get_avenants(self, code_marche):
        return list(self.conn.execute(
            "SELECT * FROM avenants WHERE code_marche = ?", (code_marche,)
        ))

    def get_montant_total_marche(self, code_marche):
        marche = self.get_marche(code_marche)
        return float(marche["montant_initial_manuel"] or 0) if marche else 0.0

    def upsert_marche(self, code_marche, data):
        """Même sémantique que l'application : toutes les colonnes sont réécrites."""
        colonnes = ("libelle", "fournisseur", "type_marche", "montant_initial_manuel",
                    "date_notification", "date_debut", "date_fin_prevue", "notes")
        valeurs = [data.get(c) for c in colonnes]
        if self.get_marche(code_marche) is None:
            self.conn.execute(
                f"INSERT INTO marches (code_marche, {', '.join(colonnes)}) "
                f"VALUES (?, {', '.join('?' * len(colonnes))})",
                [code_marche] + valeurs
            )
        else:
            self.conn.execute(
                f"UPDATE marches SET {', '.join(c + ' = ?' for c in colonnes)} "
                f"WHERE code_marche = ?",
                valeurs + [code_marche]
            )
        self.conn.commit()


class BaseTestInterface(unittest.TestCase):
    """Analyzer sur les exports du dépôt et base jetable."""

    @classmethod
    def setUpClass(cls):
        try:
            from PyQt5.QtWidgets import QApplication  # noqa: F401
        except ImportError:
            raise unittest.SkipTest("PyQt5 indisponible")

        cls.sources = sorted(glob.glob(os.path.join("data_sources", "factures*.xls")))
        if not cls.sources:
            raise unittest.SkipTest("exports SEDIT absents du dépôt")

        cls.app = _application()

        # Les exports sont relus une seule fois : chaque test ne rejoue que la
        # base, pas le chargement des cinq classeurs.
        from marches_module import MarchesAnalyzer

        cls.repertoire_classe = tempfile.mkdtemp()
        cls.analyzer = MarchesAnalyzer(
            cls.sources, use_cache=True,
            cache_path=os.path.join(cls.repertoire_classe, "cache.db"),
        )
        if not cls.analyzer.load_data():
            raise unittest.SkipTest("chargement des exports impossible")

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "repertoire_classe"):
            shutil.rmtree(cls.repertoire_classe, ignore_errors=True)

    def setUp(self):
        self.repertoire = tempfile.mkdtemp()
        self.db_path = os.path.join(self.repertoire, "suivi.db")
        conn = sqlite3.connect(self.db_path)
        conn.executescript(SCHEMA_MARCHES)
        conn.execute(
            "INSERT INTO marches (code_marche, libelle, montant_initial_manuel, "
            "date_notification, notes) VALUES ('2020_14G3P', 'Marché G3P', 4175726.0, "
            "'2025-12-19', 'à conserver')"
        )
        conn.commit()
        conn.close()

        self.db = BaseFactice(self.db_path)
        self.analyzer.db = self.db

        self._museler_les_boites()

    def tearDown(self):
        shutil.rmtree(self.repertoire, ignore_errors=True)

    def _museler_les_boites(self):
        """Neutralise les boîtes de dialogue modales et note ce qu'elles disent."""
        import PyQt5.QtWidgets as W

        self.messages = []
        self.addCleanup(setattr, W.QMessageBox, "information", W.QMessageBox.information)
        self.addCleanup(setattr, W.QMessageBox, "warning", W.QMessageBox.warning)
        self.addCleanup(setattr, W.QMessageBox, "critical", W.QMessageBox.critical)
        self.addCleanup(setattr, W.QMessageBox, "question", W.QMessageBox.question)

        journal = self.messages
        W.QMessageBox.information = staticmethod(
            lambda p, t, msg="", *a, **k: journal.append(("info", t, msg)))
        W.QMessageBox.warning = staticmethod(
            lambda p, t, msg="", *a, **k: journal.append(("warning", t, msg)))
        W.QMessageBox.critical = staticmethod(
            lambda p, t, msg="", *a, **k: journal.append(("critical", t, msg)))
        W.QMessageBox.question = staticmethod(lambda *a, **k: W.QMessageBox.Yes)

    def _titres(self):
        return [titre for _, titre, _ in self.messages]

    def _montants_en_base(self):
        conn = sqlite3.connect(self.db_path)
        try:
            return dict(conn.execute(
                "SELECT code_marche, montant_initial_manuel FROM marches"))
        finally:
            conn.close()


class TestDialogueEnveloppes(BaseTestInterface):
    """Saisie en masse des enveloppes contractuelles."""

    def _dialogue(self):
        from enveloppes_dialog import EnveloppesMarchesDialog
        return EnveloppesMarchesDialog(self.db, self.analyzer)

    def _saisir(self, dialogue, code, texte):
        from enveloppes_dialog import COL_ENVELOPPE
        dialogue.table.item(dialogue._ligne_du_marche(code), COL_ENVELOPPE).setText(texte)

    def test_tous_les_marches_sont_listes(self):
        dialogue = self._dialogue()
        self.assertEqual(dialogue.table.rowCount(), len(dialogue.fiches))
        self.assertGreater(dialogue.table.rowCount(), 1)
        self.assertIsNotNone(dialogue._ligne_du_marche("2020_14G3P"))

    def test_seule_la_colonne_enveloppe_est_editable(self):
        from PyQt5.QtCore import Qt
        from enveloppes_dialog import COL_ENVELOPPE, COL_ENGAGE, COL_FACTURE, COL_MARCHE

        dialogue = self._dialogue()
        self.assertTrue(dialogue.table.item(0, COL_ENVELOPPE).flags() & Qt.ItemIsEditable)
        for colonne in (COL_MARCHE, COL_ENGAGE, COL_FACTURE):
            self.assertFalse(
                dialogue.table.item(0, colonne).flags() & Qt.ItemIsEditable,
                f"colonne {colonne} ne doit pas être éditable",
            )

    def test_saisie_au_format_francais(self):
        from enveloppes_dialog import COL_ENVELOPPE

        dialogue = self._dialogue()
        self._saisir(dialogue, "2020_14G3P", "12 345,67 €")
        ligne = dialogue._ligne_du_marche("2020_14G3P")
        self.assertAlmostEqual(dialogue.fiches[ligne]["montant_initial"], 12345.67, places=2)
        self.assertEqual(dialogue.table.item(ligne, COL_ENVELOPPE).text(), "12 345.67 €")

    def test_montant_illisible_refuse_et_valeur_restauree(self):
        from enveloppes_dialog import COL_ENVELOPPE

        dialogue = self._dialogue()
        ligne = dialogue._ligne_du_marche("2020_14G3P")
        self._saisir(dialogue, "2020_14G3P", "à voir avec le service")

        self.assertIn("Montant illisible", self._titres())
        self.assertAlmostEqual(dialogue.fiches[ligne]["montant_initial"], 4175726.0, places=2)
        self.assertEqual(dialogue.table.item(ligne, COL_ENVELOPPE).text(), "4 175 726.00 €")

    def test_filtre_sur_marche_fournisseur_operation(self):
        dialogue = self._dialogue()
        dialogue._appliquer_filtre("2020_14G3P")
        visibles = [
            ligne for ligne in range(dialogue.table.rowCount())
            if not dialogue.table.isRowHidden(ligne)
        ]
        self.assertEqual(len(visibles), 1)

        dialogue._appliquer_filtre("")
        self.assertFalse(dialogue.table.isRowHidden(0))

    def test_enregistrement_cree_le_marche_absent(self):
        dialogue = self._dialogue()
        code = next(f["marche"] for f in dialogue.fiches if not f.get("montant_initial"))
        self._saisir(dialogue, code, "250000")
        dialogue.save()

        self.assertAlmostEqual(self._montants_en_base()[code], 250000.0, places=2)

    def test_enregistrement_preserve_les_autres_colonnes(self):
        """upsert_marche réécrit tout : le dialogue doit repasser l'existant."""
        dialogue = self._dialogue()
        self._saisir(dialogue, "2020_14G3P", "5000000")
        dialogue.save()

        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        marche = conn.execute(
            "SELECT * FROM marches WHERE code_marche = '2020_14G3P'").fetchone()
        conn.close()

        self.assertAlmostEqual(marche["montant_initial_manuel"], 5000000.0, places=2)
        self.assertEqual(marche["libelle"], "Marché G3P")
        self.assertEqual(marche["date_notification"], "2025-12-19")
        self.assertEqual(marche["notes"], "à conserver")

    def test_enregistrement_sans_modification_ne_touche_pas_la_base(self):
        avant = self._montants_en_base()
        self._dialogue().save()

        self.assertIn("Rien à enregistrer", self._titres())
        self.assertEqual(self._montants_en_base(), avant)

    def test_aller_retour_par_classeur_excel(self):
        import PyQt5.QtWidgets as W

        fichier = os.path.join(self.repertoire, "enveloppes.xlsx")
        code = next(f["marche"] for f in self._dialogue().fiches if not f.get("montant_initial"))

        dialogue = self._dialogue()
        self._saisir(dialogue, code, "175 000,00 €")
        with unittest.mock.patch.object(
            W.QFileDialog, "getSaveFileName", staticmethod(lambda *a, **k: (fichier, ""))
        ):
            dialogue.exporter_excel()
        self.assertTrue(os.path.exists(fichier))

        repris = self._dialogue()
        with unittest.mock.patch.object(
            W.QFileDialog, "getOpenFileName", staticmethod(lambda *a, **k: (fichier, ""))
        ):
            repris.importer_excel()

        ligne = repris._ligne_du_marche(code)
        self.assertAlmostEqual(repris.fiches[ligne]["montant_initial"], 175000.0, places=2)
        # L'import alimente le tableau : c'est « Enregistrer » qui écrit.
        self.assertNotIn(code, self._montants_en_base())

        repris.save()
        self.assertAlmostEqual(self._montants_en_base()[code], 175000.0, places=2)

    def test_enveloppe_saisie_reprise_par_l_export_suivi(self):
        """Une enveloppe saisie change la provenance vue par le générateur."""
        from marches_module import MarchesAnalyzer

        code = next(f["marche"] for f in self._dialogue().fiches if not f.get("montant_initial"))
        operation = next(
            op["operation"] for op in self.analyzer.get_vision_operations()
            if code in op["marches"]
        )

        _, _, _, avant, _ = self.analyzer.collecter_ecritures_operation(operation)
        self.assertNotIn(MarchesAnalyzer.ENVELOPPE_BASE, set(avant.values()))

        dialogue = self._dialogue()
        self._saisir(dialogue, code, "1000000")
        dialogue.save()

        _, _, enveloppes, apres, _ = self.analyzer.collecter_ecritures_operation(operation)
        self.assertIn(MarchesAnalyzer.ENVELOPPE_BASE, set(apres.values()))
        self.assertGreaterEqual(sum(enveloppes.values()), 1000000.0)


class _AnalyzerRestreint:
    """Analyzer réel, restreint à quelques opérations pour la vitesse du test."""

    def __init__(self, analyzer, operations):
        self._analyzer = analyzer
        self._operations = operations

    def get_vision_operations(self):
        return self._operations

    def __getattr__(self, nom):
        return getattr(self._analyzer, nom)


class TestRegenerationDepuisInterface(BaseTestInterface):
    """Bouton « Tout régénérer » : mêmes fichiers que la commande en lot."""

    def _fenetre(self):
        """Un vrai QWidget portant les méthodes de MainWindow que l'on teste.

        Instancier MainWindow ouvrirait la base de production et construirait
        toute l'interface ; on emprunte les deux méthodes telles quelles, sur un
        parent Qt valide (les boîtes de dialogue en exigent un).
        """
        from PyQt5.QtWidgets import QWidget

        module = _module_application()

        class FenetreDeTest(QWidget):
            saisir_enveloppes_marches = module.MainWindow.saisir_enveloppes_marches
            regenerer_tous_les_suivis = module.MainWindow.regenerer_tous_les_suivis

        fenetre = FenetreDeTest()
        fenetre.db = self.db
        fenetre.marches_analyzer = self.analyzer
        self.addCleanup(fenetre.deleteLater)
        return fenetre

    def test_avertit_si_les_donnees_ne_sont_pas_chargees(self):
        fenetre = self._fenetre()
        fenetre.marches_analyzer = None
        fenetre.regenerer_tous_les_suivis()

        self.assertEqual(self._titres(), ["Données non chargées"])

    def test_regenere_toutes_les_operations(self):
        import PyQt5.QtWidgets as W

        destination = os.path.join(self.repertoire, "exports")
        os.makedirs(destination)

        fenetre = self._fenetre()
        # La boucle est la même quel que soit le nombre d'opérations ; on en
        # garde quelques-unes pour que le test reste rapide.
        operations = self.analyzer.get_vision_operations()[:5]
        fenetre.marches_analyzer = _AnalyzerRestreint(self.analyzer, operations)

        with unittest.mock.patch.object(
            W.QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: destination)
        ):
            fenetre.regenerer_tous_les_suivis()

        produits = glob.glob(os.path.join(destination, "*.xlsx"))
        self.assertEqual(len(produits), len(operations))
        self.assertIn("Régénération terminée", self._titres())
        for operation in operations:
            self.assertIn(
                os.path.join(destination, f"suivi_financier_{operation['operation']}.xlsx"),
                produits,
            )

    def test_une_operation_en_echec_est_signalee_sans_interrompre(self):
        import PyQt5.QtWidgets as W

        destination = os.path.join(self.repertoire, "exports")
        os.makedirs(destination)

        operations = self.analyzer.get_vision_operations()[:3]
        restreint = _AnalyzerRestreint(self.analyzer, operations)
        en_echec = operations[1]["operation"]

        export_reel = restreint.export_suivi_financier_operation

        def export_capricieux(code_operation, filepath, **kwargs):
            if code_operation == en_echec:
                raise RuntimeError("export impossible")
            return export_reel(code_operation, filepath, **kwargs)

        restreint.export_suivi_financier_operation = export_capricieux
        fenetre = self._fenetre()
        fenetre.marches_analyzer = restreint

        with unittest.mock.patch.object(
            W.QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: destination)
        ):
            fenetre.regenerer_tous_les_suivis()

        # Les deux autres sont bien produites, et l'échec est nommé.
        self.assertEqual(len(glob.glob(os.path.join(destination, "*.xlsx"))), 2)
        niveaux = [niveau for niveau, _, _ in self.messages]
        self.assertIn("warning", niveaux)
        self.assertIn(en_echec, self.messages[-1][2])

    def test_annuler_le_choix_du_dossier_n_ecrit_rien(self):
        import PyQt5.QtWidgets as W

        fenetre = self._fenetre()
        with unittest.mock.patch.object(
            W.QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: "")
        ):
            fenetre.regenerer_tous_les_suivis()

        self.assertEqual(self.messages, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
