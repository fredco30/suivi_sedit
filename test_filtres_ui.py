"""Tests des filtres de l'interface principale.

Couvre le correctif du bug de visibilite des barres de filtres (un filtre
"Exercice" vide et non fonctionnel apparaissait sur l'onglet Facturation), le
nouveau widget de selection multiple avec recherche (SearchableCheckableComboBox)
et son cablage sur les filtres Marche/Fournisseur, ainsi que le nouveau filtre
Exercice de l'onglet Facturation.

Tourne sans affichage grace au greffon Qt "offscreen". Les tests sont ignores
si PyQt5 est absent, ou si suivi_commandes.db (donnees reelles du depot, lues
mais jamais modifiees) est introuvable.

Lancement :
    python -m unittest test_filtres_ui.py -v
"""
from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

DB_REPO = "suivi_commandes.db"


def _application():
    from PyQt5.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


def _module_application():
    """Charge le module principal, dont le nom de fichier n'est pas importable."""
    import importlib.util
    import sys

    if "appmod_filtres" not in sys.modules:
        spec = importlib.util.spec_from_file_location(
            "appmod_filtres", "suivi_commandes_factures_marches_FinaàGarder.py"
        )
        module = importlib.util.module_from_spec(spec)
        sys.modules["appmod_filtres"] = module
        spec.loader.exec_module(module)
    return sys.modules["appmod_filtres"]


class TestSearchableCheckableComboBox(unittest.TestCase):
    """Le nouveau widget de selection multiple avec recherche, en isolation.

    Une liste plate a cases a cocher devient illisible au-dela d'une
    quinzaine d'items (Marche ~100 valeurs, Fournisseur ~130 valeurs) : ce
    widget ajoute une recherche live sans perdre le multi-choix.
    """

    @classmethod
    def setUpClass(cls):
        try:
            cls.module = _module_application()
        except ImportError:
            raise unittest.SkipTest("PyQt5 indisponible")
        cls.app = _application()

    def _combo(self, items):
        w = self.module.SearchableCheckableComboBox()
        w.addItems(items)
        return w

    def test_tous_les_items_ajoutes(self):
        w = self._combo([f"Item{i}" for i in range(60)])
        # +1 pour la ligne [Tous]
        self.assertEqual(w._model.rowCount(), 61)

    def test_aucune_coche_au_depart(self):
        w = self._combo(["A", "B", "C"])
        self.assertEqual(w.checked_items(), [])
        self.assertEqual(w.currentText(), "Aucun")

    def test_recherche_filtre_la_liste(self):
        w = self._combo(["Paris", "Lyon", "Marseille", "Paris-Nord"])
        w._search.setText("paris")
        visibles = {
            w._proxy.data(w._proxy.index(r, 0)) for r in range(w._proxy.rowCount())
        }
        # [Tous] doit rester joignable meme pendant une recherche.
        self.assertIn("[Tous]", visibles)
        self.assertEqual(visibles - {"[Tous]"}, {"Paris", "Paris-Nord"})

    def _cocher(self, w, texte):
        for r in range(w._proxy.rowCount()):
            if w._proxy.data(w._proxy.index(r, 0)) == texte:
                w._on_item_pressed(w._proxy.index(r, 0))
                return
        raise AssertionError(f"item introuvable dans la vue filtree : {texte}")

    def test_cocher_un_item_trouve_par_recherche(self):
        w = self._combo(["Paris", "Lyon", "Marseille"])
        w._search.setText("lyon")
        self._cocher(w, "Lyon")
        self.assertEqual(w.checked_items(), ["Lyon"])

    def test_selection_persiste_apres_effacement_recherche(self):
        w = self._combo(["Paris", "Lyon", "Marseille"])
        w._search.setText("lyon")
        self._cocher(w, "Lyon")
        w._search.setText("")
        self.assertEqual(w.checked_items(), ["Lyon"])

    def test_tous_coche_tout_y_compris_hors_recherche(self):
        """[Tous] doit cocher l'integralite des items, pas seulement ceux
        filtres par la recherche en cours : un "tout selectionner" qui
        depend du texte tape serait surprenant."""
        w = self._combo(["Paris", "Lyon", "Marseille"])
        w._search.setText("par")  # ne matche que Paris
        self._cocher(w, "[Tous]")
        self.assertEqual(set(w.checked_items()), {"Paris", "Lyon", "Marseille"})

    def test_tous_decoche_tout(self):
        w = self._combo(["Paris", "Lyon"])
        self._cocher(w, "[Tous]")
        self._cocher(w, "[Tous]")
        self.assertEqual(w.checked_items(), [])

    def test_signal_emis_a_chaque_coche(self):
        w = self._combo(["Paris", "Lyon"])
        appels = []
        w.selectionChanged.connect(lambda: appels.append(1))
        self._cocher(w, "Paris")
        self._cocher(w, "Lyon")
        self.assertEqual(len(appels), 2)

    def test_current_text_resume_la_selection(self):
        w = self._combo(["Paris", "Lyon", "Marseille"])
        self._cocher(w, "Paris")
        self.assertEqual(w.currentText(), "Paris")
        self._cocher(w, "Lyon")
        self.assertEqual(w.currentText(), "2 sélectionnés")

    def test_clear_reinitialise_completement(self):
        w = self._combo(["Paris", "Lyon"])
        self._cocher(w, "Paris")
        w.clear()
        self.assertEqual(w._model.rowCount(), 1)  # juste [Tous]
        self.assertEqual(w.checked_items(), [])

    def test_clear_selection_decoche_sans_vider_la_liste(self):
        w = self._combo(["Paris", "Lyon"])
        self._cocher(w, "Paris")
        w.clear_selection()
        self.assertEqual(w.checked_items(), [])
        self.assertEqual(w._model.rowCount(), 3)  # items conserves

    def test_popup_personnalisee_s_affiche_et_se_ferme(self):
        from PyQt5.QtWidgets import QMainWindow

        win = QMainWindow()
        w = self._combo(["Paris", "Lyon"])
        win.setCentralWidget(w)
        win.show()
        w.showPopup()
        self.assertTrue(w._popup.isVisible())
        w.hidePopup()
        self.assertFalse(w._popup.isVisible())


_FENETRE_PARTAGEE = []  # cache module-level : une seule MainWindow pour toutes les classes


class BaseTestFenetre(unittest.TestCase):
    """Fenêtre principale réelle, construite sur les données du dépôt.

    Lecture seule : ces tests ne modifient jamais suivi_commandes.db. Une
    seule MainWindow est construite pour l'ensemble des classes de ce fichier
    (construction coûteuse : migrations SQLite, chargement de ~500 commandes
    et ~2500 factures) ; chaque test la retrouve via tearDown, qui remet les
    filtres à zéro entre deux tests.
    """

    @classmethod
    def setUpClass(cls):
        try:
            module = _module_application()
        except ImportError:
            raise unittest.SkipTest("PyQt5 indisponible")
        if not os.path.exists(DB_REPO):
            raise unittest.SkipTest(f"{DB_REPO} introuvable")

        cls.module = module
        cls.app = _application()
        if not _FENETRE_PARTAGEE:
            win = module.MainWindow(DB_REPO)
            win.show()
            for _ in range(3):
                cls.app.processEvents()
            _FENETRE_PARTAGEE.append(win)
        cls.win = _FENETRE_PARTAGEE[0]

    @classmethod
    def _laisser_respirer(cls, n=3):
        for _ in range(n):
            cls.app.processEvents()

    def tearDown(self):
        # cls.win est partagé par tous les tests de la classe (une seule
        # MainWindow, construite une fois dans setUpClass) : sans ce nettoyage,
        # une sélection laissée cochée par un test fausserait le suivant selon
        # l'ordre alphabétique d'exécution des tests. clear_selection() ne
        # notifie pas le proxy (même convention que clear_multiple_filters
        # dans l'application) : on route donc explicitement, comme le ferait
        # un vrai décochage par l'utilisateur.
        self.win.marche_filter_combo.clear_selection()
        self.win.on_marche_filter_changed()
        self.win.fournisseur_filter_combo.clear_selection()
        self.win.on_fournisseur_filter_changed()
        if self.win.filter2_combo.count():
            self.win.filter2_combo.setCurrentIndex(0)
        self._laisser_respirer()

    def _aller_sur(self, nom_onglet):
        idx = next(
            i for i in range(self.win.tabs.count()) if nom_onglet in self.win.tabs.tabText(i)
        )
        self.win.tabs.setCurrentIndex(idx)
        self._laisser_respirer()
        return idx


class TestVisibiliteDesGroupesDeFiltres(BaseTestFenetre):
    """Le bug racine : un widget caché dans une toolbar partagée par
    d'autres widgets pouvait rester (ou redevenir) visible dès qu'un widget
    voisin changeait de visibilité -- c'est ce qui laissait apparaître un
    filtre "Exercice" vide et non fonctionnel sur l'onglet Facturation. Le
    correctif scinde la toolbar en un groupe par filtre, dont la visibilité
    se pilote au niveau de la toolbar entière.
    """

    ATTENDU = {
        "📋 Commandes": dict(filter2=True, marche=True, multi=True, cmd=True, fact=False),
        "🔔 Rappels": dict(filter2=False, marche=False, multi=False, cmd=False, fact=False),
        "📄 Factures": dict(filter2=True, marche=True, multi=False, cmd=True, fact=True),
        "💰 Facturation": dict(filter2=True, marche=True, multi=False, cmd=False, fact=False),
        "📈 Suivi marchés": dict(filter2=False, marche=False, multi=False, cmd=False, fact=False),
        "📦 Opérations": dict(filter2=False, marche=False, multi=False, cmd=False, fact=False),
        "📜 Historique": dict(filter2=False, marche=False, multi=False, cmd=False, fact=False),
    }

    def _etat(self):
        return dict(
            filter2=self.win.tb_filter2.isVisible(),
            marche=self.win.tb_marche.isVisible(),
            multi=self.win.tb_multi.isVisible(),
            cmd=self.win.tb_search_cmd.isVisible(),
            fact=self.win.tb_search_fact.isVisible(),
        )

    def test_chaque_onglet_affiche_les_bons_groupes(self):
        for nom, attendu in self.ATTENDU.items():
            with self.subTest(onglet=nom):
                self._aller_sur(nom.split(" ", 1)[1] if " " in nom else nom)
                self.assertEqual(self._etat(), attendu, msg=f"onglet {nom}")

    def test_facturation_affiche_exercice_pas_facturation_perimee(self):
        """C'est le bug exact rapporté : le libellé "Exercice" doit
        s'afficher sur Facturation, jamais un vestige de "Facturation" venu
        de l'onglet Commandes visité juste avant."""
        self._aller_sur("Commandes")  # laisse le label sur "Facturation:"
        self._aller_sur("Facturation")
        self.assertIn("Exercice", self.win.filter2_label.text())
        self.assertTrue(self.win.tb_filter2.isVisible())

    def test_visibilite_stable_sur_un_aller_retour_par_onglet_masque(self):
        """Cas concret qui a fait le plus de degats en debogage : revenir sur
        Commandes apres etre passe par un onglet ou la toolbar est masquee
        (Suivi marches) ne doit rien laisser d'incorrectement cache."""
        self._aller_sur("Commandes")
        self._aller_sur("Suivi marchés")
        self._aller_sur("Historique")
        self._aller_sur("Commandes")
        self.assertEqual(self._etat(), self.ATTENDU["📋 Commandes"])

    def test_determinisme_sur_une_sequence_longue(self):
        """La meme sequence d'onglets doit produire le meme resultat a
        chaque execution -- c'est ce qui a permis de distinguer un vrai bug
        d'un artefact de synchronisation Qt pendant le debogage."""
        sequence = [
            "Commandes", "Rappels", "Factures", "Facturation",
            "Suivi marchés", "Historique", "Commandes", "Opérations",
            "Factures", "Commandes",
        ]
        resultats = []
        for nom in sequence:
            self._aller_sur(nom)
            resultats.append(self._etat())

        # Rejoue exactement la même séquence une deuxième fois.
        resultats_bis = []
        for nom in sequence:
            self._aller_sur(nom)
            resultats_bis.append(self._etat())

        self.assertEqual(resultats, resultats_bis)


class TestFiltreExerciceFacturation(BaseTestFenetre):
    """Filtre Exercice de l'onglet Facturation : inexistant jusqu'ici bien
    que la colonne Exercice soit réellement présente dans le tableau."""

    def test_items_exercice_peuples(self):
        self._aller_sur("Facturation")
        items = [self.win.filter2_combo.itemText(i) for i in range(self.win.filter2_combo.count())]
        self.assertEqual(items[0], "Tous")
        exercices_reels = {row["exercice"] for row in self.win.synth_model.rows if row["exercice"]}
        self.assertEqual(set(items[1:]), exercices_reels)

    def test_filtrer_par_exercice_reduit_les_lignes_correctement(self):
        self._aller_sur("Facturation")
        if self.win.filter2_combo.count() <= 1:
            self.skipTest("aucun exercice dans les données du dépôt")

        exercice = self.win.filter2_combo.itemText(1)
        attendu = sum(1 for row in self.win.synth_model.rows if row["exercice"] == exercice)

        self.win.filter2_combo.setCurrentText(exercice)
        self._laisser_respirer()

        self.assertEqual(self.win.synth_proxy.rowCount(), attendu)
        self.assertGreater(attendu, 0)

    def test_repasser_a_tous_restaure_toutes_les_lignes(self):
        self._aller_sur("Facturation")
        if self.win.filter2_combo.count() <= 1:
            self.skipTest("aucun exercice dans les données du dépôt")

        total = len(self.win.synth_model.rows)
        self.win.filter2_combo.setCurrentText(self.win.filter2_combo.itemText(1))
        self._laisser_respirer()
        self.win.filter2_combo.setCurrentText("Tous")
        self._laisser_respirer()

        self.assertEqual(self.win.synth_proxy.rowCount(), total)


class TestFiltreMarcheMultiSelection(BaseTestFenetre):
    """Marché passe d'un choix unique à une sélection multiple avec
    recherche : ~100 marchés en liste plate non filtrable était illisible."""

    def _cocher(self, combo, texte):
        for r in range(combo._proxy.rowCount()):
            if combo._proxy.data(combo._proxy.index(r, 0)) == texte:
                combo._on_item_pressed(combo._proxy.index(r, 0))
                return
        raise AssertionError(f"marché introuvable : {texte}")

    def test_deux_marches_coches_filtrent_par_union(self):
        self._aller_sur("Commandes")
        combo = self.win.marche_filter_combo
        marches = [combo._model.item(i, 0).text() for i in range(1, combo._model.rowCount())]
        if len(marches) < 2:
            self.skipTest("pas assez de marchés dans les données du dépôt")

        cible1, cible2 = marches[0], marches[len(marches) // 2]
        self._cocher(combo, cible1)
        self._cocher(combo, cible2)

        attendu = self.win.db.conn.execute(
            "SELECT COUNT(*) FROM commandes WHERE marche IN (?, ?)", (cible1, cible2)
        ).fetchone()[0]
        self.assertEqual(self.win.cmd_proxy.rowCount(), attendu)
        self.assertGreater(attendu, 0)

        # Décocher les deux : retour à la totalité, sans filtre résiduel.
        self._cocher(combo, cible1)
        self._cocher(combo, cible2)
        self.assertEqual(self.win.cmd_proxy.rowCount(), len(self.win.cmd_model.rows))

    def test_marche_non_reinitialise_par_construction_du_combo(self):
        """Le combo est repeuplé (clear + addItems) à chaque changement
        d'onglet : ça ne doit pas déclencher de filtrage fantôme puisque
        aucun item repeuplé n'est jamais coché automatiquement."""
        self._aller_sur("Commandes")
        total_initial = len(self.win.cmd_model.rows)
        self._aller_sur("Factures")
        self._aller_sur("Commandes")
        self.assertEqual(self.win.cmd_proxy.rowCount(), total_initial)


class TestFiltreFournisseurMultiSelection(BaseTestFenetre):
    """Même bascule multi-sélection que Marché, pour Fournisseur (~130
    valeurs dans les données du dépôt)."""

    def _cocher(self, combo, texte):
        for r in range(combo._proxy.rowCount()):
            if combo._proxy.data(combo._proxy.index(r, 0)) == texte:
                combo._on_item_pressed(combo._proxy.index(r, 0))
                return
        raise AssertionError(f"fournisseur introuvable : {texte}")

    def test_un_fournisseur_coche_filtre_correctement(self):
        self._aller_sur("Facturation")
        combo = self.win.fournisseur_filter_combo
        if combo._model.rowCount() <= 1:
            self.skipTest("pas de fournisseurs dans les données du dépôt")

        cible = combo._model.item(1, 0).text()
        self._cocher(combo, cible)

        attendu = sum(1 for row in self.win.synth_model.rows if row["fournisseur"] == cible)
        self.assertEqual(self.win.synth_proxy.rowCount(), attendu)
        self.assertGreater(attendu, 0)

    def test_fournisseur_survit_au_changement_de_tri(self):
        """Le filtre doit rester actif indépendamment du tri de la table."""
        self._aller_sur("Facturation")
        combo = self.win.fournisseur_filter_combo
        if combo._model.rowCount() <= 1:
            self.skipTest("pas de fournisseurs dans les données du dépôt")

        cible = combo._model.item(1, 0).text()
        self._cocher(combo, cible)
        avant = self.win.synth_proxy.rowCount()

        self.win.synth_proxy.sort(0)
        self._laisser_respirer()

        self.assertEqual(self.win.synth_proxy.rowCount(), avant)


class TestDescriptionFiltresActifs(BaseTestFenetre):
    """_get_active_filters_description() doit refléter les sélections
    multiples, et ne jamais confondre "Aucun coché" avec un filtre actif."""

    def _cocher(self, combo, texte):
        for r in range(combo._proxy.rowCount()):
            if combo._proxy.data(combo._proxy.index(r, 0)) == texte:
                combo._on_item_pressed(combo._proxy.index(r, 0))
                return
        raise AssertionError(f"item introuvable : {texte}")

    def test_aucun_filtre_actif_par_defaut(self):
        self._aller_sur("Commandes")
        self.win.marche_filter_combo.clear_selection()
        self.win.fournisseur_filter_combo.clear_selection()
        self.assertEqual(self.win._get_active_filters_description(), "Aucun filtre appliqué")

    def test_marches_multiples_listes_dans_la_description(self):
        self._aller_sur("Commandes")
        self.win.marche_filter_combo.clear_selection()
        combo = self.win.marche_filter_combo
        marches = [combo._model.item(i, 0).text() for i in range(1, combo._model.rowCount())]
        if len(marches) < 2:
            self.skipTest("pas assez de marchés dans les données du dépôt")

        self._cocher(combo, marches[0])
        self._cocher(combo, marches[1])
        description = self.win._get_active_filters_description()
        self.assertIn(marches[0], description)
        self.assertIn(marches[1], description)

        self._cocher(combo, marches[0])
        self._cocher(combo, marches[1])


if __name__ == "__main__":
    unittest.main(verbosity=2)
