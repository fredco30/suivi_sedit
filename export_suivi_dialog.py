"""Choix des opérations et des options de l'export « suivi financier ».

Un seul écran remplace les trois boutons d'autrefois — « Exporter suivi
financier » (une opération, sans choix d'exercice), « Export 2020_14G3P » (une
opération écrite en dur, avec choix d'exercice et présentation de contrôle) et
« Tout régénérer » (toutes les opérations, sans aucune option). Les mêmes
options sont désormais offertes quelle que soit la portée : le tri par n° de
bon de commande et le journal de contrôle n'étaient atteignables que pour une
opération sur soixante-dix-sept.
"""

from typing import List, NamedTuple, Optional, Sequence

from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QLabel,
    QRadioButton,
    QVBoxLayout,
)

PORTEE_SELECTION = "selection"
PORTEE_FILTREES = "filtrees"
PORTEE_TOUTES = "toutes"

EXERCICE_TOUS = "Tous"


def _designer(operations: Sequence[str], si_vide: str) -> str:
    """Nomme l'opération quand il n'y en a qu'une, les compte au-delà."""
    if not operations:
        return si_vide
    if len(operations) == 1:
        return operations[0]
    return f"{len(operations)} opérations"


class ChoixExport(NamedTuple):
    """Ce que l'utilisateur a demandé : quoi exporter, et comment."""

    operations: List[str]
    exercice: Optional[str]
    trier_par_bdc: bool
    journal: bool

    @property
    def multiple(self) -> bool:
        return len(self.operations) > 1


class ExportSuiviFinancierDialog(QDialog):
    """Portée de l'export, exercice, et options de présentation."""

    def __init__(
        self,
        selection: Sequence[str],
        filtrees: Sequence[str],
        toutes: Sequence[str],
        exercices: Sequence[str],
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Exporter le suivi financier")
        self.setMinimumWidth(460)

        self._portees = {
            PORTEE_SELECTION: list(selection),
            PORTEE_FILTREES: list(filtrees),
            PORTEE_TOUTES: list(toutes),
        }

        disposition = QVBoxLayout(self)

        groupe_portee = QGroupBox("Opérations à exporter")
        choix_portee = QVBoxLayout(groupe_portee)
        self._boutons = QButtonGroup(self)

        def _ajouter(cle: str, intitule: str, actif: bool = True) -> QRadioButton:
            bouton = QRadioButton(intitule)
            bouton.setEnabled(actif)
            bouton.setProperty("portee", cle)
            self._boutons.addButton(bouton)
            choix_portee.addWidget(bouton)
            return bouton

        selectionnees = self._portees[PORTEE_SELECTION]
        self._bouton_selection = _ajouter(
            PORTEE_SELECTION,
            "La sélection — " + _designer(selectionnees, "aucune ligne sélectionnée"),
            actif=bool(selectionnees),
        )

        # Le filtre du bandeau ne restreignait pas « Tout régénérer » : taper un
        # code puis régénérer produisait les soixante-dix-sept fichiers.
        operations_filtrees = self._portees[PORTEE_FILTREES]
        filtre_actif = 0 < len(operations_filtrees) < len(self._portees[PORTEE_TOUTES])
        self._bouton_filtrees = _ajouter(
            PORTEE_FILTREES,
            "Le filtre en cours — " + _designer(operations_filtrees, "aucune opération"),
            actif=filtre_actif,
        )
        self._bouton_filtrees.setVisible(filtre_actif)

        self._bouton_toutes = _ajouter(
            PORTEE_TOUTES, f"Toutes les {len(self._portees[PORTEE_TOUTES])} opérations"
        )

        if selectionnees:
            self._bouton_selection.setChecked(True)
        elif filtre_actif:
            self._bouton_filtrees.setChecked(True)
        else:
            self._bouton_toutes.setChecked(True)

        disposition.addWidget(groupe_portee)

        groupe_options = QGroupBox("Options")
        formulaire = QFormLayout(groupe_options)

        self._exercice = QComboBox()
        self._exercice.addItems(list(exercices) or [EXERCICE_TOUS])
        self._exercice.setCurrentIndex(0)
        formulaire.addRow("Exercice :", self._exercice)

        self._trier = QCheckBox("Trier les lignes par n° de bon de commande")
        self._journal = QCheckBox("Écrire un journal de contrôle dans run_logs/")
        formulaire.addRow(self._trier)
        formulaire.addRow(self._journal)

        disposition.addWidget(groupe_options)

        rappel = QLabel(
            "Une seule opération : le fichier est nommé à l'enregistrement.\n"
            "Plusieurs : un dossier est demandé, un fichier par opération."
        )
        rappel.setStyleSheet("color: #666;")
        rappel.setWordWrap(True)
        disposition.addWidget(rappel)

        boutons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        boutons.button(QDialogButtonBox.Ok).setText("Exporter")
        boutons.button(QDialogButtonBox.Cancel).setText("Annuler")
        boutons.accepted.connect(self.accept)
        boutons.rejected.connect(self.reject)
        disposition.addWidget(boutons)

    def portee(self) -> str:
        bouton = self._boutons.checkedButton()
        return bouton.property("portee") if bouton else PORTEE_TOUTES

    def choix(self) -> ChoixExport:
        """Traduit l'écran en consigne d'export."""
        exercice = self._exercice.currentText()
        return ChoixExport(
            operations=list(self._portees[self.portee()]),
            exercice=None if exercice == EXERCICE_TOUS else exercice,
            trier_par_bdc=self._trier.isChecked(),
            journal=self._journal.isChecked(),
        )
