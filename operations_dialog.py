"""Dialogue d'arbitrage du rattachement des marchés à leur opération."""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QDialogButtonBox, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout,
)

from operations_marches import operations_a_arbitrer, racine_codification

COL_MARCHE, COL_REGLE, COL_OPERATION, COL_TITULAIRE = range(4)

EN_TETES = [
    "N° MARCHÉ",
    "Opération déduite\n(règle de codification)",
    "OPÉRATION RETENUE",
    "Titulaire(s)",
]

COULEUR_A_ARBITRER = QColor("#FFF2CC")
COULEUR_REGLE = QColor("#F5F5F5")
COULEUR_SAISI = QColor("#E0F0FF")


class CorrespondanceOperationsDialog(QDialog):
    """Rattache à la main les marchés que la codification ne permet pas de lire.

    La règle automatique retire du code marché un dernier segment numérique de
    un ou deux chiffres et y voit un n° de lot. Elle ne peut rien deviner
    au-delà : `MC157_01` et `MC157_02` deviennent deux opérations, `2019_06P1`
    à `P3R` quatre, `2020_14G1` à `GO` sept. Savoir si `P2` est une phase de
    `2019_06` ou une affaire distincte relève du dossier, pas du programme.

    Seule la colonne « OPÉRATION RETENUE » est modifiable. La laisser vide,
    c'est s'en remettre à la règle : la table de rattachement est vide par
    défaut, et rien ne change tant que rien n'y est saisi.
    """

    def __init__(self, db, analyzer, parent=None):
        super().__init__(parent)
        self.db = db
        self.analyzer = analyzer
        self.fiches = []
        self._chargement = False

        self.setWindowTitle("Rattachement des marchés aux opérations")
        self.resize(1000, 700)

        self._init_ui()
        self._charger_marches()

    # ------------------------------------------------------------------ UI

    def _init_ui(self):
        layout = QVBoxLayout(self)

        explication = QLabel(
            "<b>Un marché appartient à une opération.</b> Le programme le déduit du code "
            "(<i>2024_17_3</i> → <i>2024_17</i>), mais certaines codifications lui échappent : "
            "<i>MC157_01</i> et <i>MC157_02</i>, <i>2019_06P1</i> à <i>P3R</i>, "
            "<i>2020_14G1</i> à <i>GO</i>…<br>"
            "Saisissez l'opération dans la troisième colonne pour trancher ; "
            "laissez vide pour vous en remettre à la règle. "
            "Les lignes <span style='background:#FFF2CC'>surlignées</span> sont celles "
            "dont le code ressemble à celui d'un voisin — à vérifier en priorité."
        )
        explication.setWordWrap(True)
        layout.addWidget(explication)

        barre = QHBoxLayout()
        barre.addWidget(QLabel("🔍 Rechercher :"))
        self.recherche = QLineEdit()
        self.recherche.setPlaceholderText("Marché, opération ou titulaire…")
        self.recherche.setClearButtonEnabled(True)
        self.recherche.textChanged.connect(self._appliquer_filtre)
        barre.addWidget(self.recherche)

        self.case_a_arbitrer = QCheckBox("N'afficher que les cas à arbitrer")
        self.case_a_arbitrer.toggled.connect(lambda _: self._appliquer_filtre(self.recherche.text()))
        barre.addWidget(self.case_a_arbitrer)
        barre.addStretch()
        layout.addLayout(barre)

        self.table = QTableWidget(0, len(EN_TETES))
        self.table.setHorizontalHeaderLabels(EN_TETES)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        entete = self.table.horizontalHeader()
        entete.setSectionResizeMode(COL_MARCHE, QHeaderView.ResizeToContents)
        entete.setSectionResizeMode(COL_REGLE, QHeaderView.ResizeToContents)
        entete.setSectionResizeMode(COL_OPERATION, QHeaderView.ResizeToContents)
        entete.setSectionResizeMode(COL_TITULAIRE, QHeaderView.Stretch)
        self.table.itemChanged.connect(self._on_item_change)
        layout.addWidget(self.table)

        self.etat = QLabel()
        self.etat.setStyleSheet("color: #666;")
        layout.addWidget(self.etat)

        boutons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel, self)
        boutons.button(QDialogButtonBox.Save).setText("Enregistrer")
        boutons.button(QDialogButtonBox.Cancel).setText("Annuler")
        boutons.accepted.connect(self.save)
        boutons.rejected.connect(self.reject)

        bouton_vider = QPushButton("Effacer tous les rattachements")
        bouton_vider.setToolTip(
            "Revenir à la règle automatique pour tous les marchés."
        )
        bouton_vider.clicked.connect(self._effacer_rattachements)
        boutons.addButton(bouton_vider, QDialogButtonBox.ResetRole)

        layout.addWidget(boutons)

    # ------------------------------------------------------- Chargement

    def _charger_marches(self):
        """Dresse la liste des marchés vus dans les écritures, avec leur opération."""
        self._chargement = True
        try:
            rattachements = dict(self.analyzer.rattachements_manuels())
            fiches = []
            for marche in self.analyzer.get_vision_globale():
                code = str(marche.get("marche") or "").strip()
                if not code:
                    continue
                fiches.append({
                    "marche": code,
                    "regle": self.analyzer.extract_operation(code),
                    "saisi": rattachements.get(code, ""),
                    "titulaire": marche.get("fournisseur", ""),
                })
            fiches.sort(key=lambda fiche: fiche["marche"])
            self.fiches = fiches

            # Les souches portant plusieurs opérations : c'est là que la
            # codification prête à discussion.
            retenues = {fiche["saisi"] or fiche["regle"] for fiche in fiches}
            self.souches_a_arbitrer = set(operations_a_arbitrer(retenues))

            self.table.setRowCount(len(fiches))
            for ligne, fiche in enumerate(fiches):
                self._remplir_ligne(ligne, fiche)
        finally:
            self._chargement = False
        self._rafraichir_etat()

    def _remplir_ligne(self, ligne, fiche):
        a_arbitrer = self._a_arbitrer(fiche)

        marche = QTableWidgetItem(fiche["marche"])
        marche.setFlags(marche.flags() & ~Qt.ItemIsEditable)
        marche.setFont(QFont("", -1, QFont.Bold))
        self.table.setItem(ligne, COL_MARCHE, marche)

        regle = QTableWidgetItem(fiche["regle"])
        regle.setFlags(regle.flags() & ~Qt.ItemIsEditable)
        regle.setBackground(QBrush(COULEUR_REGLE))
        regle.setToolTip("Ce que la règle de codification déduit du code marché.")
        self.table.setItem(ligne, COL_REGLE, regle)

        operation = QTableWidgetItem(fiche["saisi"])
        operation.setBackground(QBrush(
            COULEUR_SAISI if fiche["saisi"]
            else (COULEUR_A_ARBITRER if a_arbitrer else COULEUR_REGLE)
        ))
        operation.setToolTip(
            "Vide : la règle décide.\n"
            "Renseigné : ce marché rejoint cette opération, à l'écran comme à l'export."
        )
        self.table.setItem(ligne, COL_OPERATION, operation)

        titulaire = QTableWidgetItem(str(fiche["titulaire"] or ""))
        titulaire.setFlags(titulaire.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(ligne, COL_TITULAIRE, titulaire)

        if a_arbitrer:
            marche.setBackground(QBrush(COULEUR_A_ARBITRER))
            marche.setToolTip(
                "D'autres marchés portent un code très proche mais tombent dans une "
                "autre opération : à confirmer ou à regrouper."
            )

    def _a_arbitrer(self, fiche) -> bool:
        retenue = fiche["saisi"] or fiche["regle"]
        return racine_codification(retenue) in self.souches_a_arbitrer

    def _ligne_du_marche(self, code_marche):
        for ligne, fiche in enumerate(self.fiches):
            if fiche["marche"] == code_marche:
                return ligne
        return None

    # ---------------------------------------------------------- Édition

    def _on_item_change(self, item):
        if self._chargement or item.column() != COL_OPERATION:
            return
        ligne = item.row()
        if not 0 <= ligne < len(self.fiches):
            return
        self.fiches[ligne]["saisi"] = item.text().strip()
        self._chargement = True
        try:
            item.setBackground(QBrush(
                COULEUR_SAISI if self.fiches[ligne]["saisi"]
                else (COULEUR_A_ARBITRER if self._a_arbitrer(self.fiches[ligne])
                      else COULEUR_REGLE)
            ))
        finally:
            self._chargement = False
        self._rafraichir_etat()

    def _effacer_rattachements(self):
        if not any(fiche["saisi"] for fiche in self.fiches):
            return
        reponse = QMessageBox.question(
            self,
            "Effacer les rattachements",
            "Revenir à la règle automatique pour tous les marchés ?\n\n"
            "Les regroupements saisis seront perdus à l'enregistrement.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reponse != QMessageBox.Yes:
            return
        self._chargement = True
        try:
            for ligne, fiche in enumerate(self.fiches):
                fiche["saisi"] = ""
                item = self.table.item(ligne, COL_OPERATION)
                if item is not None:
                    item.setText("")
                    item.setBackground(QBrush(
                        COULEUR_A_ARBITRER if self._a_arbitrer(fiche) else COULEUR_REGLE
                    ))
        finally:
            self._chargement = False
        self._rafraichir_etat()

    def _appliquer_filtre(self, texte):
        texte = (texte or "").strip().lower()
        seulement_arbitrage = self.case_a_arbitrer.isChecked()
        for ligne, fiche in enumerate(self.fiches):
            champs = " ".join([
                fiche["marche"], fiche["regle"], fiche["saisi"],
                str(fiche["titulaire"] or ""),
            ]).lower()
            visible = (not texte or texte in champs) and (
                not seulement_arbitrage or self._a_arbitrer(fiche)
            )
            self.table.setRowHidden(ligne, not visible)

    def _rafraichir_etat(self):
        saisis = sum(1 for fiche in self.fiches if fiche["saisi"])
        arbitrer = sum(1 for fiche in self.fiches if self._a_arbitrer(fiche))
        operations = {fiche["saisi"] or fiche["regle"] for fiche in self.fiches}
        self.etat.setText(
            f"{len(self.fiches)} marchés → {len(operations)} opérations · "
            f"{saisis} rattachement(s) saisi(s) · "
            f"{arbitrer} marché(s) à vérifier"
        )

    # ------------------------------------------------------ Enregistrement

    def rattachements(self):
        """Ce qui sera écrit : les marchés dont l'opération a été saisie."""
        return {
            fiche["marche"]: fiche["saisi"]
            for fiche in self.fiches
            if fiche["saisi"]
        }

    def save(self):
        """Écrit les rattachements en base, et efface ceux qu'on a vidés."""
        ecrivain = getattr(self.db, "set_operation_marche", None)
        if ecrivain is None:
            QMessageBox.critical(
                self,
                "Enregistrement impossible",
                "Cette base ne gère pas le rattachement des marchés aux opérations.",
            )
            return

        try:
            for fiche in self.fiches:
                # Un champ vidé efface le rattachement : l'appel est fait
                # aussi pour eux, sans quoi un arbitrage annulé survivrait.
                ecrivain(fiche["marche"], fiche["saisi"])
        except Exception as erreur:
            QMessageBox.critical(
                self, "Enregistrement impossible", f"Écriture en base refusée :\n\n{erreur}"
            )
            return

        self.analyzer.invalider_vision()
        self.accept()
