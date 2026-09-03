"""Tests du lanceur.

Portent sur ce que le lanceur decide -- quel script demarrer, ce qui manque,
comment assouplir les versions figees -- sans jamais demarrer l'application.

Lancement :
    python -m unittest test_lanceur.py -v
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

import lanceur


class TestTrouverApplication(unittest.TestCase):
    """Le script principal est repéré par motif : son nom porte un accent."""

    def setUp(self):
        self.repertoire = Path(tempfile.mkdtemp())
        self._racine = lanceur.RACINE
        lanceur.RACINE = self.repertoire
        self.addCleanup(setattr, lanceur, "RACINE", self._racine)
        self.addCleanup(shutil.rmtree, self.repertoire, True)

    def _creer(self, nom, age_secondes=0):
        chemin = self.repertoire / nom
        chemin.write_text("# script", encoding="utf-8")
        if age_secondes:
            ancien = chemin.stat().st_mtime - age_secondes
            os.utime(chemin, (ancien, ancien))
        return chemin

    def test_nom_accentue_trouve(self):
        attendu = self._creer("suivi_commandes_factures_marches_FinaàGarder.py")
        self.assertEqual(lanceur.trouver_application(), attendu)

    def test_le_plus_recent_l_emporte(self):
        self._creer("suivi_commandes_factures_marches_ancien.py", age_secondes=86400)
        recent = self._creer("suivi_commandes_factures_marches_FinaàGarder.py")
        self.assertEqual(lanceur.trouver_application(), recent)

    def test_aucun_script_erreur_explicite(self):
        with self.assertRaises(SystemExit) as contexte:
            lanceur.trouver_application()
        self.assertIn("Aucun script", str(contexte.exception))


class TestDependances(unittest.TestCase):

    def test_module_absent_signale_son_paquet_pip(self):
        original = dict(lanceur.DEPENDANCES)
        lanceur.DEPENDANCES["module_qui_n_existe_pas"] = "paquet-fictif"
        self.addCleanup(lanceur.DEPENDANCES.update, original)
        self.addCleanup(lanceur.DEPENDANCES.pop, "module_qui_n_existe_pas", None)

        self.assertIn("paquet-fictif", lanceur.dependances_manquantes())

    def test_modules_presents_rien_a_signaler(self):
        original = dict(lanceur.DEPENDANCES)
        lanceur.DEPENDANCES.clear()
        lanceur.DEPENDANCES["sys"] = "sys"
        self.addCleanup(lanceur.DEPENDANCES.update, original)

        self.assertEqual(lanceur.dependances_manquantes(), [])


class TestRequirementsAssouplies(unittest.TestCase):
    """Les versions figées deviennent des minima quand pip les refuse.

    pandas 2.1.4 n'a pas de binaire au-delà de Python 3.12 : pip tenterait de
    compiler depuis les sources et échouerait faute de compilateur.
    """

    def setUp(self):
        self.repertoire = Path(tempfile.mkdtemp())
        self._requirements, self._logs = lanceur.REQUIREMENTS, lanceur.DOSSIER_LOGS
        lanceur.REQUIREMENTS = self.repertoire / "requirements.txt"
        lanceur.DOSSIER_LOGS = self.repertoire / "run_logs"
        self.addCleanup(setattr, lanceur, "REQUIREMENTS", self._requirements)
        self.addCleanup(setattr, lanceur, "DOSSIER_LOGS", self._logs)
        self.addCleanup(shutil.rmtree, self.repertoire, True)

    def _assouplir(self, contenu):
        lanceur.REQUIREMENTS.write_text(contenu, encoding="utf-8")
        return lanceur._requirements_assouplies().read_text(encoding="utf-8").splitlines()

    def test_versions_figees_relachees(self):
        lignes = self._assouplir("pandas==2.1.4\nPyQt5==5.15.11\n")
        self.assertEqual(lignes, ["pandas>=2.1.4", "PyQt5>=5.15.11"])

    def test_commentaires_et_lignes_vides_conserves(self):
        lignes = self._assouplir("# outils\n\nxlrd==2.0.2\n")
        self.assertEqual(lignes, ["# outils", "", "xlrd>=2.0.2"])

    def test_contraintes_deja_souples_inchangees(self):
        lignes = self._assouplir("pandas>=2.2.3\nopenpyxl\n")
        self.assertEqual(lignes, ["pandas>=2.2.3", "openpyxl"])

    def test_une_seule_occurrence_remplacee(self):
        # Une ligne avec marqueur d'environnement ne doit pas être abîmée.
        lignes = self._assouplir('pandas==2.1.4; python_version=="3.11"\n')
        self.assertEqual(lignes, ['pandas>=2.1.4; python_version=="3.11"'])


class TestVerificationSeule(unittest.TestCase):

    def test_verif_ne_lance_pas_l_application(self):
        appels = []
        original = lanceur.lancer_application
        lanceur.lancer_application = lambda script: appels.append(script) or 0
        self.addCleanup(setattr, lanceur, "lancer_application", original)

        self.assertEqual(lanceur.main(["--verif"]), 0)
        self.assertEqual(appels, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
