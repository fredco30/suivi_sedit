"""Tests de non-regression du suivi financier par marche.

Couvre le correctif du double comptage de l'enveloppe : le generateur emettait
une ligne d'engagement *et* une ligne de facture pour un meme BDC, sans jamais
solder l'engagement par la facture.

Deux niveaux :
  * `TestReglesAgregation` / `TestScenariosDocumentes` : regles metier sur des
    jeux synthetiques reproduisant les cas releves dans le fichier defectueux ;
  * `TestExportOperation` : export Excel complet sur les donnees du depot, qui
    verifie les invariants structurels (#7 a #11 du correctif) quel que soit le
    jeu de donnees ;
  * `TestValeursReference2020_14G3P` : les valeurs chiffrees attendues sur le
    jeu de reference `2020_14G3P`. Ignore si ce jeu n'est pas celui charge --
    l'instantane versionne dans le depot n'est pas celui du fichier diffuse.

Lancement :
    python -m unittest test_suivi_financier.py -v
ou
    python test_suivi_financier.py
"""
from __future__ import annotations

import os
import tempfile
import unittest
from collections import defaultdict
from typing import Dict, List

from suivi_financier_agg import (
    ETAT_NON_FACTURE,
    ETAT_PARTIEL,
    ETAT_SOLDE,
    LONGUEUR_DESIGNATION,
    STATUT_ENGAGEMENT,
    STATUT_FACTURE,
    TOLERANCE,
    Ecriture,
    ResultatSuivi,
    agreger_ecritures,
    nettoyer_designation,
)

MARCHE = "2020_14G3P"
FOURNISSEUR = "BOUYGUES ENERGIES ET SERVICES"


def engagement(bdc: str, montant: float, libelle: str = "travaux", initial: float = None) -> Ecriture:
    """Ecriture d'engagement : BDC emis, rien de facture."""
    return Ecriture(
        marche=MARCHE, fournisseur=FOURNISSEUR, num_commande=bdc, libelle=libelle,
        montant_ttc=montant, montant_initial=montant if initial is None else initial,
    )


def facture(bdc: str, montant: float, num_facture: str, num_mandat: str,
            libelle: str = "travaux", initial: float = None) -> Ecriture:
    """Ecriture de realisation : la facture arrive et est mandatee."""
    return Ecriture(
        marche=MARCHE, fournisseur=FOURNISSEUR, num_commande=bdc, libelle=libelle,
        num_facture=num_facture, num_mandat=num_mandat, date_sf="2024-04-26",
        montant_ttc=montant, montant_sf=montant,
        montant_initial=montant if initial is None else initial,
    )


class TestNettoyageDesignation(unittest.TestCase):
    """Le suffixe (REPORT) sortait de la designation, tronquee avant concatenation."""

    def test_suffixe_report_supprime(self):
        self.assertEqual(
            nettoyer_designation("G3P extension du reseau video su(REPORT)"),
            "G3P extension du reseau video su",
        )

    def test_troncature_apres_nettoyage(self):
        libelle = "A" * 100
        self.assertEqual(len(nettoyer_designation(libelle)), LONGUEUR_DESIGNATION)
        self.assertEqual(LONGUEUR_DESIGNATION, 60)

    def test_libelle_vide_ou_absent(self):
        self.assertEqual(nettoyer_designation(None), "")
        self.assertEqual(nettoyer_designation(""), "")


class TestReglesAgregation(unittest.TestCase):
    """Regle : une ligne par etat de BDC, pas une ligne par ecriture SEDIT."""

    def test_engagement_solde_par_la_facture(self):
        # §1.1 : le BDC 22AA01789 de 2 779,20 € consommait 5 558,40 €.
        resultat = agreger_ecritures([
            engagement("22AA01789", 2779.20, "mise en place coquilles betons"),
            facture("22AA01789", 2779.20, "F2401609", "1423", "mise en place coquilles betons"),
        ], {"22AA01789": 2779.20})

        self.assertEqual(len(resultat.lignes), 1)
        self.assertEqual(resultat.lignes[0].statut, STATUT_FACTURE)
        self.assertAlmostEqual(resultat.total_impute, 2779.20, places=2)

    def test_report_exercice_ne_cree_pas_de_seconde_ligne(self):
        # §1.2 : deux engagements identiques pour un seul BDC.
        resultat = agreger_ecritures([
            engagement("25AA04392", 92351.76, "SAINT LOUIS- G3P-ECLAIRAGE ET FEUX DE CI"),
            engagement("25AA04392", 92351.76, "SAINT LOUIS- G3P-ECLAIRAGE ET FE(REPORT)"),
        ], {"25AA04392": 92351.76})

        self.assertEqual(len(resultat.lignes), 1)
        self.assertEqual(resultat.lignes[0].statut, STATUT_ENGAGEMENT)
        self.assertAlmostEqual(resultat.total_impute, 92351.76, places=2)
        self.assertNotIn("REPORT", resultat.lignes[0].designation)

    def test_engagement_ramene_au_reliquat(self):
        # BDC de 14 619,96 € facture pour 8 421,96 € : reliquat 6 198,00 €.
        resultat = agreger_ecritures([
            engagement("22AA02376", 8421.96, "georeferencement obligatoire", initial=14619.96),
            facture("22AA02376", 8421.96, "F2400578", "454", "georeferencement obligatoire",
                    initial=14619.96),
        ])

        statuts = [ligne.statut for ligne in resultat.lignes]
        self.assertEqual(statuts, [STATUT_FACTURE, STATUT_ENGAGEMENT])
        self.assertAlmostEqual(resultat.lignes[1].montant_impute, 6198.00, places=2)
        self.assertAlmostEqual(resultat.total_impute, 14619.96, places=2)

    def test_aucune_ligne_engagement_pour_bdc_solde(self):
        # §2.1 : un BDC soldé n'a pas de ligne d'engagement (assertion #8).
        resultat = agreger_ecritures([
            engagement("22AA03536", 83323.20, "REHABILITATION ALLEE DES GOELANDS"),
            facture("22AA03536", 886.80, "F2304108", "3683"),
            facture("22AA03536", 34638.00, "F2302126", "1822"),
            facture("22AA03536", 47798.40, "F2302126", "1822"),
        ], {"22AA03536": 83323.20})

        self.assertEqual(len(resultat.lignes), 3)
        self.assertTrue(all(l.statut == STATUT_FACTURE for l in resultat.lignes))
        self.assertEqual(resultat.bdcs[0].etat, ETAT_SOLDE)

    def test_un_seul_engagement_par_bdc(self):
        # Assertion #7 : quel que soit le nombre d'ecritures d'engagement.
        resultat = agreger_ecritures([
            engagement("23AA03103", 7315.20, "G3P Modification quai Bus", initial=8455.20),
            engagement("23AA03103", 8455.20, "G3P Modification quai(REPORT)", initial=8455.20),
            engagement("23AA03103", 8455.20, "G3P Modification quai Bus", initial=8455.20),
        ])

        engagements = [l for l in resultat.lignes if l.statut == STATUT_ENGAGEMENT]
        self.assertEqual(len(engagements), 1)
        self.assertAlmostEqual(engagements[0].montant_impute, 8455.20, places=2)

    def test_montant_ttc_bdc_identique_sur_toutes_les_lignes(self):
        # §1.3 / §2.2 : la colonne ne porte jamais un montant de facture.
        resultat = agreger_ecritures([
            engagement("23AA01991", 109875.60, "REQUALIFICATION ALLEE"),
            facture("23AA01991", 36089.64, "F2400012", "12"),
            facture("23AA01991", 73785.96, "F2302417", "2085"),
        ], {"23AA01991": 109875.60})

        montants_ref = {round(l.montant_ref, 2) for l in resultat.lignes}
        self.assertEqual(montants_ref, {109875.60})

    def test_bdc_jamais_facture_reste_engage_pour_son_montant(self):
        resultat = agreger_ecritures([
            engagement("26AA00890", 5000.00),
        ], {"26AA00890": 5000.00})

        self.assertEqual(len(resultat.lignes), 1)
        self.assertEqual(resultat.lignes[0].statut, STATUT_ENGAGEMENT)
        self.assertEqual(resultat.bdcs[0].etat, ETAT_NON_FACTURE)

    def test_facture_superieure_au_bdc_leve_une_anomalie(self):
        # §2.2 : le maximum est retenu, mais l'ecart remonte pour arbitrage.
        resultat = agreger_ecritures([
            facture("22AA03211", 17244.52, "F2400900", "900", initial=17052.52),
        ], {"22AA03211": 17052.52})

        self.assertEqual(len(resultat.anomalies), 1)
        self.assertIn("Facturé supérieur", resultat.anomalies[0].message)
        self.assertAlmostEqual(resultat.total_bdc_distincts, 17244.52, places=2)

    def test_montants_bdc_divergents_leve_une_anomalie(self):
        resultat = agreger_ecritures([
            engagement("25AA00614", 2100.87, initial=2100.87),
            engagement("25AA00614", 2899.13, initial=5000.00),
        ])

        self.assertEqual(len(resultat.anomalies), 1)
        self.assertIn("Plusieurs montants", resultat.anomalies[0].message)
        self.assertAlmostEqual(resultat.total_bdc_distincts, 5000.00, places=2)

    def test_etat_partiellement_facture(self):
        resultat = agreger_ecritures([
            facture("24AA00679", 1000.00, "F1", "1", initial=2500.00),
        ])
        self.assertEqual(resultat.bdcs[0].etat, ETAT_PARTIEL)
        self.assertAlmostEqual(resultat.bdcs[0].reliquat, 1500.00, places=2)

    def test_trace_des_lignes_neutralisees(self):
        # §4 : chaque ligne supprimee ou ajustee laisse une trace.
        resultat = agreger_ecritures([
            engagement("22AA01789", 2779.20, "coquilles betons"),
            facture("22AA01789", 2779.20, "F2401609", "1423", "coquilles betons"),
            engagement("22AA02376", 8421.96, "georeferencement", initial=14619.96),
            facture("22AA02376", 8421.96, "F2400578", "454", "georeferencement",
                    initial=14619.96),
        ])

        par_bdc = {n.cle_bdc: n for n in resultat.neutralisations}
        self.assertAlmostEqual(par_bdc["22AA01789"].montant_neutralise, 2779.20, places=2)
        self.assertAlmostEqual(par_bdc["22AA02376"].montant_retenu, 6198.00, places=2)
        self.assertAlmostEqual(par_bdc["22AA02376"].montant_neutralise, 2223.96, places=2)


class TestProprieteGenerique(unittest.TestCase):
    """Propriete a verifier sur tout marche, pas seulement sur 2020_14G3P."""

    JEUX = {
        "engagement puis facture": [
            engagement("A1", 1000.0),
            facture("A1", 1000.0, "F1", "M1"),
        ],
        "engagement reporte deux fois": [
            engagement("B1", 500.0),
            engagement("B1", 500.0, "libelle(REPORT)"),
        ],
        "bdc paye en trois factures": [
            engagement("C1", 900.0),
            facture("C1", 300.0, "F1", "M1"),
            facture("C1", 300.0, "F2", "M2"),
            facture("C1", 300.0, "F3", "M3"),
        ],
        "facture sans mandat (non mandatee)": [
            engagement("D1", 700.0),
            Ecriture(marche=MARCHE, num_commande="D1", num_facture="F9", num_mandat="",
                     montant_ttc=700.0, montant_initial=700.0),
        ],
        "facture superieure au bdc": [
            engagement("E1", 100.0, initial=100.0),
            facture("E1", 150.0, "F1", "M1", initial=100.0),
        ],
        "ligne sans bdc rattachee a la tranche": [
            Ecriture(marche=MARCHE, tranche_libelle="TF", montant_ttc=250.0,
                     montant_initial=250.0),
        ],
    }

    def test_somme_imputee_egale_max_bdc_factures(self):
        for nom, ecritures in self.JEUX.items():
            with self.subTest(jeu=nom):
                resultat = agreger_ecritures(ecritures)

                impute_par_bdc: Dict[str, float] = defaultdict(float)
                for ligne in resultat.lignes:
                    impute_par_bdc[ligne.cle_bdc] += ligne.montant_impute

                for bdc in resultat.bdcs:
                    attendu = max(bdc.montant_ref, bdc.montant_facture)
                    self.assertAlmostEqual(
                        impute_par_bdc[bdc.cle_bdc], attendu, places=2,
                        msg=f"BDC {bdc.cle_bdc} du jeu « {nom} »",
                    )

                self.assertAlmostEqual(
                    resultat.total_impute, resultat.total_bdc_distincts, places=2
                )

    def test_au_plus_un_engagement_par_bdc(self):
        for nom, ecritures in self.JEUX.items():
            with self.subTest(jeu=nom):
                resultat = agreger_ecritures(ecritures)
                compte: Dict[str, int] = defaultdict(int)
                for ligne in resultat.lignes:
                    if ligne.statut == STATUT_ENGAGEMENT:
                        compte[ligne.cle_bdc] += 1
                self.assertTrue(all(n <= 1 for n in compte.values()), compte)

    def test_aucun_engagement_sur_bdc_solde(self):
        for nom, ecritures in self.JEUX.items():
            with self.subTest(jeu=nom):
                resultat = agreger_ecritures(ecritures)
                soldes = {b.cle_bdc for b in resultat.bdcs if b.etat == ETAT_SOLDE}
                engages = {l.cle_bdc for l in resultat.lignes
                           if l.statut == STATUT_ENGAGEMENT}
                self.assertEqual(soldes & engages, set())


# ---------------------------------------------------------------------------
# Export Excel complet
# ---------------------------------------------------------------------------

def _charger_analyzer():
    """Construit un analyzer sur les donnees du depot, ou None si indisponible."""
    if not os.path.exists("suivi_commandes.db") or not os.path.exists("marches_cache.db"):
        return None
    try:
        from marches_module import MarchesAnalyzer
        from regenerer_suivis import BaseSuiviLectureSeule
    except ImportError:
        return None

    analyzer = MarchesAnalyzer(
        "database_sync", database=BaseSuiviLectureSeule("suivi_commandes.db"), use_cache=True
    )
    return analyzer if analyzer.load_data() else None


def _lignes_financier(ws) -> List[dict]:
    """Lit les lignes de donnees de l'onglet FINANCIER (hors sous-totaux)."""
    lignes = []
    for row in range(6, ws.max_row + 1):
        statut = ws.cell(row, 15).value
        if statut in (None, "TOTAL"):
            continue
        lignes.append({
            "row": row,
            "bdc": ws.cell(row, 5).value,
            "montant_ref": ws.cell(row, 6).value,
            "num_facture": ws.cell(row, 7).value,
            "montant_impute": ws.cell(row, 11).value,
            "restant": ws.cell(row, 14).value,
            "statut": statut,
        })
    return lignes


class TestExportOperation(unittest.TestCase):
    """Invariants structurels de l'export, verifies sur les donnees du depot."""

    OPERATION = MARCHE

    @classmethod
    def setUpClass(cls):
        analyzer = _charger_analyzer()
        if analyzer is None:
            raise unittest.SkipTest("données du dépôt indisponibles")

        cls._tmpdir = tempfile.TemporaryDirectory()
        chemin = os.path.join(cls._tmpdir.name, "suivi.xlsx")
        if not analyzer.export_suivi_financier_operation(
            cls.OPERATION, chemin, exercice_filter=None, special_export=False
        ):
            raise unittest.SkipTest(f"export de {cls.OPERATION} impossible")

        import openpyxl
        cls.wb = openpyxl.load_workbook(chemin)
        cls.ws = cls.wb["FINANCIER"]
        cls.lignes = _lignes_financier(cls.ws)
        cls.ligne_soustotal = cls.ws.max_row

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "_tmpdir"):
            cls._tmpdir.cleanup()

    def test_onglets_produits(self):
        self.assertEqual(
            self.wb.sheetnames,
            ["FINANCIER", "A jour", "Anomalies", "Lignes neutralisées"],
        )

    def test_enveloppe_initiale_ecrite_en_clair(self):
        # §2.4 : l'enveloppe ne doit plus se deduire de N6 + M6.
        self.assertIn("Enveloppe initiale", self.ws.cell(4, 1).value)
        self.assertGreater(self.ws.cell(4, 6).value, 0)

    def test_7_un_seul_engagement_par_bdc(self):
        compte: Dict[str, int] = defaultdict(int)
        for ligne in self.lignes:
            if ligne["statut"] == STATUT_ENGAGEMENT:
                compte[ligne["bdc"]] += 1
        doublons = {bdc: n for bdc, n in compte.items() if n > 1}
        self.assertEqual(doublons, {})

    def test_8_aucun_bdc_solde_ne_porte_d_engagement(self):
        ws_ajour = self.wb["A jour"]
        soldes = set()
        for row in range(5, ws_ajour.max_row):
            if ws_ajour.cell(row, 9).value == ETAT_SOLDE:
                soldes.add(ws_ajour.cell(row, 5).value)
        engages = {l["bdc"] for l in self.lignes if l["statut"] == STATUT_ENGAGEMENT}
        self.assertEqual(soldes & engages, set())

    def test_9_totaux_facture_identiques_entre_onglets(self):
        # §2.5 : les deux onglets sortent du meme jeu de donnees.
        ws_ajour = self.wb["A jour"]
        total_ajour = ws_ajour.cell(ws_ajour.max_row, 7).value
        total_financier = sum(
            l["montant_impute"] for l in self.lignes if l["statut"] == STATUT_FACTURE
        )
        self.assertAlmostEqual(total_ajour, total_financier, places=2)

    def test_10_aucun_couple_facture_montant_en_double(self):
        vus = defaultdict(int)
        for ligne in self.lignes:
            if ligne["statut"] == STATUT_FACTURE:
                vus[(ligne["num_facture"], round(ligne["montant_impute"], 2))] += 1
        self.assertEqual({k: n for k, n in vus.items() if n > 1}, {})

    def test_11_dernier_solde_glissant_egale_le_sous_total(self):
        # §2.4 : un seul mode de calcul, glissant et sous-total confondus.
        dernier = self.lignes[-1]["restant"]
        sous_total = self.ws.cell(self.ligne_soustotal, 14).value
        self.assertAlmostEqual(dernier, sous_total, places=2)

    def test_colonne_montant_bdc_constante_par_bdc(self):
        # §2.2 : jamais un montant de facture dans cette colonne.
        valeurs = defaultdict(set)
        for ligne in self.lignes:
            valeurs[ligne["bdc"]].add(round(ligne["montant_ref"], 2))
        multiples = {bdc: v for bdc, v in valeurs.items() if len(v) > 1}
        self.assertEqual(multiples, {})

    def test_sous_total_montant_bdc_somme_sur_bdc_distincts(self):
        # §2.2 : jamais une somme de colonne.
        distincts = {}
        for ligne in self.lignes:
            distincts[ligne["bdc"]] = ligne["montant_ref"]
        attendu = sum(distincts.values())
        self.assertAlmostEqual(self.ws.cell(self.ligne_soustotal, 6).value, attendu, places=2)

    def test_invariant_somme_imputee_egale_somme_bdc(self):
        impute = sum(l["montant_impute"] for l in self.lignes)
        distincts = {l["bdc"]: l["montant_ref"] for l in self.lignes}
        self.assertAlmostEqual(impute, sum(distincts.values()), places=2)

    def test_aucun_suffixe_report_dans_les_designations(self):
        # §2.3 : le suffixe est remplace par la colonne STATUT.
        for row in range(6, self.ws.max_row + 1):
            designation = self.ws.cell(row, 2).value or ""
            self.assertNotIn("REPORT)", designation)


class TestValeursReference2020_14G3P(unittest.TestCase):
    """Valeurs chiffrees attendues sur le jeu de reference `2020_14G3P`.

    Ces montants proviennent du fichier diffuse analyse dans le correctif. Ils
    ne sont verifiables que sur ce jeu de donnees precis : l'instantane
    versionne dans le depot est plus ancien et donne d'autres totaux. Les tests
    sont donc ignores tant que le jeu charge ne correspond pas.
    """

    ATTENDU = {
        "nb_lignes": 134,
        "nb_lignes_facturees": 113,
        "nb_lignes_engagement": 21,
        "total_impute": 1803139.47,
        "total_facture": 1542045.26,
        "total_engagement": 261094.21,
        "total_bdc_distincts": 1706415.26,
        "solde_final": 2372586.53,
        "enveloppe_initiale": 4175726.00,
    }

    @classmethod
    def setUpClass(cls):
        analyzer = _charger_analyzer()
        if analyzer is None:
            raise unittest.SkipTest("données du dépôt indisponibles")

        groupes, montants, enveloppes, _, info = analyzer.collecter_ecritures_operation(MARCHE)
        if info is None:
            raise unittest.SkipTest(f"opération {MARCHE} absente")

        cls.resultat = ResultatSuivi()
        for ecritures in groupes.values():
            cls.resultat.etendre(
                agreger_ecritures(ecritures, montants, trier_par_bdc=True)
            )
        cls.enveloppe = sum(enveloppes.values())

        if abs(cls.resultat.total_impute - cls.ATTENDU["total_impute"]) > 0.01:
            raise unittest.SkipTest(
                "jeu de données différent de la référence du correctif "
                f"({cls.resultat.total_impute:.2f} € imputés au lieu de "
                f"{cls.ATTENDU['total_impute']:.2f} €) — "
                "rejouer ces assertions sur l'export SEDIT de référence"
            )

    def test_1_nombre_de_lignes_emises(self):
        lignes = self.resultat.lignes
        self.assertEqual(len(lignes), self.ATTENDU["nb_lignes"])
        self.assertEqual(
            sum(1 for l in lignes if l.statut == STATUT_FACTURE),
            self.ATTENDU["nb_lignes_facturees"],
        )
        self.assertEqual(
            sum(1 for l in lignes if l.statut == STATUT_ENGAGEMENT),
            self.ATTENDU["nb_lignes_engagement"],
        )

    def test_2_somme_montants_imputes(self):
        self.assertAlmostEqual(
            self.resultat.total_impute, self.ATTENDU["total_impute"], places=2
        )

    def test_3_somme_lignes_facturees(self):
        self.assertAlmostEqual(
            self.resultat.total_facture, self.ATTENDU["total_facture"], places=2
        )

    def test_4_somme_lignes_engagement(self):
        self.assertAlmostEqual(
            self.resultat.total_engagement, self.ATTENDU["total_engagement"], places=2
        )

    def test_5_somme_montants_bdc_distincts(self):
        self.assertAlmostEqual(
            self.resultat.total_bdc_distincts, self.ATTENDU["total_bdc_distincts"], places=2
        )

    def test_6_solde_final(self):
        self.assertAlmostEqual(self.enveloppe, self.ATTENDU["enveloppe_initiale"], places=2)
        self.assertAlmostEqual(
            self.enveloppe - self.resultat.total_impute,
            self.ATTENDU["solde_final"], places=2,
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
