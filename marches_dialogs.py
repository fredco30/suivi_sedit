"""
Dialogue d'édition d'un marché avec gestion des avenants
"""

from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel,
    QLineEdit, QTextEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QDialogButtonBox, QDoubleSpinBox, QDateEdit, QComboBox, QSpinBox,
    QMessageBox, QHeaderView, QGroupBox
)


class EditMarcheDialog(QDialog):
    """Dialogue pour éditer les informations d'un marché et gérer ses avenants."""

    def __init__(self, db, code_marche, marche_data=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.code_marche = code_marche
        self.marche_data = marche_data or {}

        self.setWindowTitle(f"Édition du marché : {code_marche}")
        self.resize(900, 700)

        # Charger les données existantes depuis la BD
        self.marche_bd = self.db.get_marche(code_marche)
        self.avenants = list(self.db.get_avenants(code_marche)) if self.marche_bd else []
        self.tranches = list(self.db.get_tranches(code_marche)) if self.marche_bd else []

        self._init_ui()
        self._load_data()

    def _init_ui(self):
        """Initialise l'interface utilisateur."""
        layout = QVBoxLayout(self)

        # === INFORMATIONS GÉNÉRALES ===
        group_info = QGroupBox("Informations générales")
        form_info = QFormLayout()

        # Code marché (lecture seule)
        self.edit_code = QLineEdit(self.code_marche)
        self.edit_code.setReadOnly(True)
        form_info.addRow("Code marché:", self.edit_code)

        # Libellé
        self.edit_libelle = QLineEdit()
        self.edit_libelle.setPlaceholderText("Libellé du marché")
        form_info.addRow("Libellé:", self.edit_libelle)

        # Fournisseur (lecture seule, vient d'Excel)
        self.edit_fournisseur = QLineEdit()
        self.edit_fournisseur.setReadOnly(True)
        self.edit_fournisseur.setStyleSheet("background-color: #f0f0f0;")
        form_info.addRow("Fournisseur:", self.edit_fournisseur)

        # Type de marché
        self.combo_type = QComboBox()
        self.combo_type.addItem("Marché classique (montant initial + tranches)", "CLASSIQUE")
        self.combo_type.addItem("Marché à bons de commande", "BDC")
        self.combo_type.currentIndexChanged.connect(self._on_type_changed)
        form_info.addRow("Type de marché:", self.combo_type)

        # Montant initial manuel
        self.spin_montant = QDoubleSpinBox()
        self.spin_montant.setRange(0, 999999999)
        self.spin_montant.setDecimals(2)
        self.spin_montant.setSuffix(" €")
        self.spin_montant.setGroupSeparatorShown(True)
        form_info.addRow("Montant initial:", self.spin_montant)

        # Montant Excel (info)
        self.label_montant_excel = QLabel()
        self.label_montant_excel.setStyleSheet("color: #666; font-style: italic;")
        form_info.addRow("Montant Excel:", self.label_montant_excel)

        group_info.setLayout(form_info)
        layout.addWidget(group_info)

        # === DATES ===
        group_dates = QGroupBox("Dates du marché")
        form_dates = QFormLayout()

        self.date_notification = QDateEdit()
        self.date_notification.setCalendarPopup(True)
        self.date_notification.setDate(QDate.currentDate())
        self.date_notification.setDisplayFormat("dd/MM/yyyy")
        form_dates.addRow("Date notification:", self.date_notification)

        self.date_debut = QDateEdit()
        self.date_debut.setCalendarPopup(True)
        self.date_debut.setDate(QDate.currentDate())
        self.date_debut.setDisplayFormat("dd/MM/yyyy")
        form_dates.addRow("Date début:", self.date_debut)

        self.date_fin = QDateEdit()
        self.date_fin.setCalendarPopup(True)
        self.date_fin.setDate(QDate.currentDate().addYears(1))
        self.date_fin.setDisplayFormat("dd/MM/yyyy")
        form_dates.addRow("Date fin prévue:", self.date_fin)

        group_dates.setLayout(form_dates)
        layout.addWidget(group_dates)

        # === TRANCHES ===
        # Cette section n'est visible que pour les marchés classiques
        self.group_tranches = QGroupBox("Tranches (TF, TO1, TO2...)")
        layout_tranches = QVBoxLayout()

        # Boutons actions tranches
        btn_layout_tranches = QHBoxLayout()
        self.btn_add_tranche = QPushButton("➕ Ajouter une tranche")
        self.btn_add_tranche.clicked.connect(self.add_tranche)
        self.btn_edit_tranche = QPushButton("✏️ Modifier")
        self.btn_edit_tranche.clicked.connect(self.edit_tranche)
        self.btn_delete_tranche = QPushButton("🗑️ Supprimer")
        self.btn_delete_tranche.clicked.connect(self.delete_tranche)

        btn_layout_tranches.addWidget(self.btn_add_tranche)
        btn_layout_tranches.addWidget(self.btn_edit_tranche)
        btn_layout_tranches.addWidget(self.btn_delete_tranche)
        btn_layout_tranches.addStretch()

        layout_tranches.addLayout(btn_layout_tranches)

        # Table des tranches
        self.table_tranches = QTableWidget()
        self.table_tranches.setColumnCount(5)
        self.table_tranches.setHorizontalHeaderLabels([
            "Code", "Libellé", "Montant", "Ordre", "ID"
        ])
        self.table_tranches.setColumnHidden(4, True)  # Cacher la colonne ID
        self.table_tranches.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_tranches.setSelectionMode(QTableWidget.SingleSelection)
        self.table_tranches.horizontalHeader().setStretchLastSection(False)
        self.table_tranches.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_tranches.doubleClicked.connect(self.edit_tranche)

        layout_tranches.addWidget(self.table_tranches)

        # Total des tranches
        self.label_total_tranches = QLabel()
        self.label_total_tranches.setStyleSheet("font-weight: bold; color: #0078d4;")
        layout_tranches.addWidget(self.label_total_tranches)

        self.group_tranches.setLayout(layout_tranches)
        layout.addWidget(self.group_tranches)

        # === AVENANTS ===
        group_avenants = QGroupBox("Avenants")
        layout_avenants = QVBoxLayout()

        # Boutons actions avenants
        btn_layout = QHBoxLayout()
        btn_add = QPushButton("➕ Ajouter un avenant")
        btn_add.clicked.connect(self.add_avenant)
        btn_edit = QPushButton("✏️ Modifier")
        btn_edit.clicked.connect(self.edit_avenant)
        btn_delete = QPushButton("🗑️ Supprimer")
        btn_delete.clicked.connect(self.delete_avenant)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_delete)
        btn_layout.addStretch()

        layout_avenants.addLayout(btn_layout)

        # Table des avenants
        self.table_avenants = QTableWidget()
        self.table_avenants.setColumnCount(6)
        self.table_avenants.setHorizontalHeaderLabels([
            "N°", "Libellé", "Montant", "Type", "Date", "ID"
        ])
        self.table_avenants.setColumnHidden(5, True)  # Cacher la colonne ID
        self.table_avenants.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_avenants.setSelectionMode(QTableWidget.SingleSelection)
        self.table_avenants.horizontalHeader().setStretchLastSection(False)
        self.table_avenants.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_avenants.doubleClicked.connect(self.edit_avenant)

        layout_avenants.addWidget(self.table_avenants)

        # Total des avenants
        self.label_total_avenants = QLabel()
        self.label_total_avenants.setStyleSheet("font-weight: bold; color: #0078d4;")
        layout_avenants.addWidget(self.label_total_avenants)

        group_avenants.setLayout(layout_avenants)
        layout.addWidget(group_avenants)

        # === NOTES ===
        group_notes = QGroupBox("Notes")
        layout_notes = QVBoxLayout()

        self.edit_notes = QTextEdit()
        self.edit_notes.setMaximumHeight(100)
        self.edit_notes.setPlaceholderText("Notes et remarques sur le marché...")
        layout_notes.addWidget(self.edit_notes)

        group_notes.setLayout(layout_notes)
        layout.addWidget(group_notes)

        # === BOUTONS ===
        buttons = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_type_changed(self):
        """Appelé quand le type de marché change."""
        type_marche = self.combo_type.currentData()

        # Pour les marchés à BDC, le montant initial n'a pas de sens
        # (il est calculé depuis les BDC eux-mêmes)
        # Et la section tranches n'est pas pertinente
        if type_marche == "BDC":
            self.spin_montant.setEnabled(False)
            self.spin_montant.setValue(0)
            self.spin_montant.setStyleSheet("background-color: #f0f0f0;")
            self.group_tranches.setVisible(False)
        else:
            self.spin_montant.setEnabled(True)
            self.spin_montant.setStyleSheet("")
            self.group_tranches.setVisible(True)

    def _load_data(self):
        """Charge les données dans le formulaire."""
        # Fournisseur depuis Excel
        fournisseur = self.marche_data.get('fournisseur', '')
        self.edit_fournisseur.setText(fournisseur)

        # Montant Excel
        montant_excel = self.marche_data.get('montant_excel', 0)
        if montant_excel > 0:
            self.label_montant_excel.setText(f"{montant_excel:,.2f} €".replace(",", " "))
        else:
            self.label_montant_excel.setText("Non calculable depuis Excel")

        if self.marche_bd:
            # Charger depuis la BD
            self.edit_libelle.setText(self.marche_bd["libelle"] or "")
            self.spin_montant.setValue(self.marche_bd["montant_initial_manuel"] or 0)
            self.edit_notes.setPlainText(self.marche_bd["notes"] or "")

            # Type de marché
            try:
                type_marche = self.marche_bd["type_marche"] if self.marche_bd["type_marche"] else "CLASSIQUE"
            except (KeyError, IndexError):
                type_marche = "CLASSIQUE"
            index = self.combo_type.findData(type_marche)
            if index >= 0:
                self.combo_type.setCurrentIndex(index)

            # Dates
            if self.marche_bd["date_notification"]:
                try:
                    date = QDate.fromString(self.marche_bd["date_notification"], "yyyy-MM-dd")
                    self.date_notification.setDate(date)
                except:
                    pass

            if self.marche_bd["date_debut"]:
                try:
                    date = QDate.fromString(self.marche_bd["date_debut"], "yyyy-MM-dd")
                    self.date_debut.setDate(date)
                except:
                    pass

            if self.marche_bd["date_fin_prevue"]:
                try:
                    date = QDate.fromString(self.marche_bd["date_fin_prevue"], "yyyy-MM-dd")
                    self.date_fin.setDate(date)
                except:
                    pass
        else:
            # Nouvelle saisie : pré-remplir avec les données Excel
            libelle = self.marche_data.get('libelle_marche', '')
            self.edit_libelle.setText(libelle)

        # Charger les avenants
        self.refresh_avenants_table()

        # Charger les tranches
        self.refresh_tranches_table()

    def refresh_avenants_table(self):
        """Rafraîchit la table des avenants."""
        self.avenants = list(self.db.get_avenants(self.code_marche))
        self.table_avenants.setRowCount(len(self.avenants))

        total_augmentation = 0
        total_diminution = 0

        for row, avenant in enumerate(self.avenants):
            # N°
            self.table_avenants.setItem(row, 0, QTableWidgetItem(str(avenant["numero_avenant"] or "")))

            # Libellé
            self.table_avenants.setItem(row, 1, QTableWidgetItem(avenant["libelle"] or ""))

            # Montant
            montant = avenant["montant"] or 0
            self.table_avenants.setItem(row, 2, QTableWidgetItem(f"{montant:,.2f} €".replace(",", " ")))

            # Type
            type_mod = avenant["type_modification"] or "Augmentation"
            self.table_avenants.setItem(row, 3, QTableWidgetItem(type_mod))

            # Date
            self.table_avenants.setItem(row, 4, QTableWidgetItem(avenant["date_avenant"] or ""))

            # ID (caché)
            self.table_avenants.setItem(row, 5, QTableWidgetItem(str(avenant["id"])))

            # Calcul du total
            if type_mod == "Diminution":
                total_diminution += montant
            else:
                total_augmentation += montant

        # Afficher le total
        total_net = total_augmentation - total_diminution
        signe = "+" if total_net >= 0 else ""
        self.label_total_avenants.setText(
            f"Total avenants : {signe}{total_net:,.2f} € "
            f"(+{total_augmentation:,.2f} € / -{total_diminution:,.2f} €)".replace(",", " ")
        )

    def add_avenant(self):
        """Ajoute un nouvel avenant."""
        dialog = EditAvenantDialog(self.code_marche, None, self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            self.db.add_avenant(self.code_marche, data)
            self.refresh_avenants_table()

    def edit_avenant(self):
        """Modifie l'avenant sélectionné."""
        selected = self.table_avenants.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Aucune sélection", "Veuillez sélectionner un avenant à modifier.")
            return

        avenant_id = int(self.table_avenants.item(selected, 5).text())
        avenant = next((a for a in self.avenants if a["id"] == avenant_id), None)

        if avenant:
            dialog = EditAvenantDialog(self.code_marche, avenant, self)
            if dialog.exec_() == QDialog.Accepted:
                data = dialog.get_data()
                self.db.update_avenant(avenant_id, data)
                self.refresh_avenants_table()

    def delete_avenant(self):
        """Supprime l'avenant sélectionné."""
        selected = self.table_avenants.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Aucune sélection", "Veuillez sélectionner un avenant à supprimer.")
            return

        reply = QMessageBox.question(
            self,
            "Confirmer la suppression",
            "Êtes-vous sûr de vouloir supprimer cet avenant ?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            avenant_id = int(self.table_avenants.item(selected, 5).text())
            self.db.delete_avenant(avenant_id)
            self.refresh_avenants_table()

    def add_tranche(self):
        """Ajoute une nouvelle tranche."""
        dialog = EditTrancheDialog(self.code_marche, None, self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            self.db.add_tranche(self.code_marche, data)
            self.refresh_tranches_table()

    def edit_tranche(self):
        """Modifie la tranche sélectionnée."""
        selected = self.table_tranches.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Aucune sélection", "Veuillez sélectionner une tranche à modifier.")
            return

        tranche_id = int(self.table_tranches.item(selected, 4).text())
        tranche = next((t for t in self.tranches if t["id"] == tranche_id), None)

        if tranche:
            dialog = EditTrancheDialog(self.code_marche, tranche, self)
            if dialog.exec_() == QDialog.Accepted:
                data = dialog.get_data()
                self.db.update_tranche(tranche_id, data)
                self.refresh_tranches_table()

    def delete_tranche(self):
        """Supprime la tranche sélectionnée."""
        selected = self.table_tranches.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Aucune sélection", "Veuillez sélectionner une tranche à supprimer.")
            return

        reply = QMessageBox.question(
            self,
            "Confirmer la suppression",
            "Êtes-vous sûr de vouloir supprimer cette tranche ?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            tranche_id = int(self.table_tranches.item(selected, 4).text())
            self.db.delete_tranche(tranche_id)
            self.refresh_tranches_table()

    def refresh_tranches_table(self):
        """Rafraîchit la table des tranches."""
        self.tranches = list(self.db.get_tranches(self.code_marche))
        self.table_tranches.setRowCount(len(self.tranches))

        total_tranches = 0

        for row, tranche in enumerate(self.tranches):
            # Code
            self.table_tranches.setItem(row, 0, QTableWidgetItem(tranche["code_tranche"] or ""))

            # Libellé
            self.table_tranches.setItem(row, 1, QTableWidgetItem(tranche["libelle"] or ""))

            # Montant
            montant = tranche["montant"] or 0
            self.table_tranches.setItem(row, 2, QTableWidgetItem(f"{montant:,.2f} €".replace(",", " ")))
            total_tranches += montant

            # Ordre
            self.table_tranches.setItem(row, 3, QTableWidgetItem(str(tranche["ordre"] or 0)))

            # ID (caché)
            self.table_tranches.setItem(row, 4, QTableWidgetItem(str(tranche["id"])))

        # Afficher le total
        self.label_total_tranches.setText(f"Total tranches : {total_tranches:,.2f} €".replace(",", " "))

    def save(self):
        """Sauvegarde les modifications."""
        data = {
            "libelle": self.edit_libelle.text().strip(),
            "fournisseur": self.edit_fournisseur.text().strip(),
            "type_marche": self.combo_type.currentData(),
            "montant_initial_manuel": self.spin_montant.value(),
            "date_notification": self.date_notification.date().toString("yyyy-MM-dd"),
            "date_debut": self.date_debut.date().toString("yyyy-MM-dd"),
            "date_fin_prevue": self.date_fin.date().toString("yyyy-MM-dd"),
            "notes": self.edit_notes.toPlainText().strip()
        }

        self.db.upsert_marche(self.code_marche, data)
        self.accept()


class EditAvenantDialog(QDialog):
    """Dialogue pour éditer un avenant."""

    def __init__(self, code_marche, avenant=None, parent=None):
        super().__init__(parent)
        self.code_marche = code_marche
        self.avenant = avenant

        title = "Modifier l'avenant" if avenant else "Ajouter un avenant"
        self.setWindowTitle(title)
        self.resize(500, 400)

        self._init_ui()
        if avenant:
            self._load_data()

    def _init_ui(self):
        """Initialise l'interface."""
        layout = QVBoxLayout(self)
        form = QFormLayout()

        # Numéro d'avenant
        self.spin_numero = QSpinBox()
        self.spin_numero.setRange(1, 99)
        form.addRow("Numéro d'avenant:", self.spin_numero)

        # Libellé
        self.edit_libelle = QLineEdit()
        self.edit_libelle.setPlaceholderText("Ex: Avenant n°1 - Travaux supplémentaires")
        form.addRow("Libellé:", self.edit_libelle)

        # Montant
        self.spin_montant = QDoubleSpinBox()
        self.spin_montant.setRange(0, 999999999)
        self.spin_montant.setDecimals(2)
        self.spin_montant.setSuffix(" €")
        self.spin_montant.setGroupSeparatorShown(True)
        form.addRow("Montant:", self.spin_montant)

        # Type de modification
        self.combo_type = QComboBox()
        self.combo_type.addItems(["Augmentation", "Diminution"])
        form.addRow("Type:", self.combo_type)

        # Date
        self.date_avenant = QDateEdit()
        self.date_avenant.setCalendarPopup(True)
        self.date_avenant.setDate(QDate.currentDate())
        self.date_avenant.setDisplayFormat("dd/MM/yyyy")
        form.addRow("Date de l'avenant:", self.date_avenant)

        # Motif
        self.edit_motif = QTextEdit()
        self.edit_motif.setMaximumHeight(100)
        self.edit_motif.setPlaceholderText("Motif de l'avenant...")
        form.addRow("Motif:", self.edit_motif)

        layout.addLayout(form)

        # Boutons
        buttons = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _load_data(self):
        """Charge les données de l'avenant."""
        self.spin_numero.setValue(self.avenant["numero_avenant"] or 1)
        self.edit_libelle.setText(self.avenant["libelle"] or "")
        self.spin_montant.setValue(self.avenant["montant"] or 0)
        self.combo_type.setCurrentText(self.avenant["type_modification"] or "Augmentation")
        self.edit_motif.setPlainText(self.avenant["motif"] or "")

        if self.avenant["date_avenant"]:
            try:
                date = QDate.fromString(self.avenant["date_avenant"], "yyyy-MM-dd")
                self.date_avenant.setDate(date)
            except:
                pass

    def get_data(self):
        """Retourne les données saisies."""
        return {
            "numero_avenant": self.spin_numero.value(),
            "libelle": self.edit_libelle.text().strip(),
            "montant": self.spin_montant.value(),
            "type_modification": self.combo_type.currentText(),
            "date_avenant": self.date_avenant.date().toString("yyyy-MM-dd"),
            "motif": self.edit_motif.toPlainText().strip()
        }


class EditTrancheDialog(QDialog):
    """Dialogue pour éditer une tranche."""

    def __init__(self, code_marche, tranche=None, parent=None):
        super().__init__(parent)
        self.code_marche = code_marche
        self.tranche = tranche

        title = "Modifier la tranche" if tranche else "Ajouter une tranche"
        self.setWindowTitle(title)
        self.resize(450, 300)

        self._init_ui()
        if tranche:
            self._load_data()

    def _init_ui(self):
        """Initialise l'interface."""
        layout = QVBoxLayout(self)
        form = QFormLayout()

        # Code tranche
        self.edit_code = QLineEdit()
        self.edit_code.setPlaceholderText("Ex: TF, TO1, TO2...")
        form.addRow("Code tranche:", self.edit_code)

        # Libellé
        self.edit_libelle = QLineEdit()
        self.edit_libelle.setPlaceholderText("Ex: Tranche ferme, Tranche optionnelle 1...")
        form.addRow("Libellé:", self.edit_libelle)

        # Montant
        self.spin_montant = QDoubleSpinBox()
        self.spin_montant.setRange(0, 999999999)
        self.spin_montant.setDecimals(2)
        self.spin_montant.setSuffix(" €")
        self.spin_montant.setGroupSeparatorShown(True)
        form.addRow("Montant:", self.spin_montant)

        # Ordre
        self.spin_ordre = QSpinBox()
        self.spin_ordre.setRange(0, 99)
        self.spin_ordre.setToolTip("Ordre d'affichage (0=TF, 1=TO1, 2=TO2, etc.)")
        form.addRow("Ordre:", self.spin_ordre)

        layout.addLayout(form)

        # Boutons
        buttons = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _load_data(self):
        """Charge les données de la tranche."""
        self.edit_code.setText(self.tranche["code_tranche"] or "")
        self.edit_libelle.setText(self.tranche["libelle"] or "")
        self.spin_montant.setValue(self.tranche["montant"] or 0)
        self.spin_ordre.setValue(self.tranche["ordre"] or 0)

    def get_data(self):
        """Retourne les données saisies."""
        return {
            "code_tranche": self.edit_code.text().strip(),
            "libelle": self.edit_libelle.text().strip(),
            "montant": self.spin_montant.value(),
            "ordre": self.spin_ordre.value()
        }
