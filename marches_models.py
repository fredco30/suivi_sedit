"""
Modèles PyQt5 pour le suivi des marchés publics
"""

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel
from PyQt5.QtGui import QBrush, QColor
from datetime import datetime, date
from typing import List, Dict


# ============== COLONNES POUR LA VISION GLOBALE ==============

MARCHES_GLOBAUX_COLUMNS = [
    ("marche", "Marché"),
    ("operation", "Opération"),
    ("libelle_marche", "Libellé\nmarché"),
    ("fournisseur", "Fournisseur"),
    ("montant_initial_marche", "Montant initial\nmarché"),
    ("nb_avenants", "Avenants"),
    ("service_fait_cumule", "Service fait\ncumulé"),
    ("paye_cumule", "Payé\ncumulé"),
    ("reste_a_realiser", "Reste à\nréaliser"),
    ("reste_a_mandater", "Reste à\nmandater"),
    ("pourcent_consomme", "% consommé"),
]


# ============== COLONNES POUR LE DÉTAIL PAR TRANCHES ==============

MARCHES_TRANCHES_COLUMNS = [
    ("marche", "Marché"),
    ("tranche_libelle", "Tranche"),
    ("montant_initial_tranche", "Montant initial\ntranche"),
    ("service_fait_tranche", "Service fait\ntranche"),
    ("paye_tranche", "Payé\ntranche"),
    ("pourcent_consomme_tranche", "% consommé"),
]


# ============== COLONNES POUR LA VISION OPÉRATIONS ==============

OPERATIONS_COLUMNS = [
    ("operation", "Opération"),
    ("nb_lots", "Nb lots"),
    ("marches", "Marchés"),
    ("libelle", "Libellé"),
    ("fournisseur", "Fournisseur"),
    ("montant_initial_total", "Montant initial\ntotal"),
    ("nb_avenants_total", "Avenants"),
    ("service_fait_total", "Service fait\ntotal"),
    ("paye_total", "Payé\ntotal"),
    ("reste_a_realiser", "Reste à\nréaliser"),
    ("reste_a_mandater", "Reste à\nmandater"),
    ("pourcent_consomme", "% consommé"),
]


# ============== COLONNES POUR L'HISTORIQUE DES FACTURES ==============

HISTORIQUE_COLUMNS = [
    ("marche", "Marché"),
    ("fournisseur", "Fournisseur"),
    ("date_sf", "Date SF"),
    ("num_facture", "N° Facture"),
    ("libelle", "Libellé"),
    ("montant_sf", "Montant SF"),
    ("montant_ttc", "Montant TTC"),
    ("num_mandat", "N° Mandat"),
    ("statut", "Statut"),
]


# ============== MODÈLE POUR LA VISION GLOBALE ==============

class MarchesGlobauxTableModel(QAbstractTableModel):
    """Modèle pour la vision globale des marchés (1 ligne = 1 marché)."""

    def __init__(self, rows=None):
        super().__init__()
        self.rows = rows or []

    def set_data(self, rows: List[Dict]):
        """Met à jour les données du modèle."""
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def refresh(self, rows: List[Dict]):
        """Alias de set_data pour compatibilité."""
        self.set_data(rows)

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(MARCHES_GLOBAUX_COLUMNS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row_data = self.rows[index.row()]
        key, _ = MARCHES_GLOBAUX_COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            value = row_data.get(key)

            # Formatage des montants
            if key in ("montant_initial_marche", "service_fait_cumule", "paye_cumule",
                      "reste_a_realiser", "reste_a_mandater"):
                if value is not None:
                    try:
                        return f"{float(value):,.2f} €".replace(",", " ")
                    except:
                        return str(value)

            # Formatage du pourcentage
            if key == "pourcent_consomme":
                if value is not None:
                    try:
                        return f"{float(value):.1f} %"
                    except:
                        return str(value)

            # Formatage des avenants
            if key == "nb_avenants":
                nb = row_data.get("nb_avenants", 0)
                if nb > 0:
                    return f"✓ {nb}"
                return ""

            # Texte standard
            return str(value) if value is not None else ""

        if role == Qt.ToolTipRole:
            # Tooltip pour le montant initial
            if key == "montant_initial_marche":
                montant_excel = row_data.get("montant_excel", 0)
                nb_avenants = row_data.get("nb_avenants", 0)
                if nb_avenants > 0:
                    return f"Montant ajusté manuellement avec {nb_avenants} avenant(s)"
                elif montant_excel > 0:
                    return "Montant calculé automatiquement depuis Excel"
                else:
                    return "Aucun montant initial renseigné"

        if role == Qt.TextAlignmentRole:
            # Alignement à droite pour les montants et pourcentages
            if key in ("montant_initial_marche", "service_fait_cumule", "paye_cumule",
                      "reste_a_realiser", "reste_a_mandater", "pourcent_consomme"):
                return Qt.AlignRight | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.BackgroundRole:
            # Coloration selon le pourcentage consommé
            pourcent = row_data.get("pourcent_consomme", 0)
            if pourcent >= 100:
                return QBrush(QColor("#ffe0e0"))  # Rouge si dépassé
            elif pourcent >= 90:
                return QBrush(QColor("#fff3cd"))  # Orange si proche
            elif pourcent >= 50:
                return QBrush(QColor("#fff9e6"))  # Jaune léger si en cours
            else:
                return QBrush(QColor("#e0ffe0"))  # Vert si peu consommé

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return MARCHES_GLOBAUX_COLUMNS[section][1]
        return section + 1


# ============== MODÈLE POUR LE DÉTAIL PAR TRANCHES ==============

class MarchesTranchesTableModel(QAbstractTableModel):
    """Modèle pour le détail par tranches (1 ligne = 1 couple marché/tranche)."""

    def __init__(self, rows=None):
        super().__init__()
        self.rows = rows or []

    def set_data(self, rows: List[Dict]):
        """Met à jour les données du modèle."""
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def refresh(self, rows: List[Dict]):
        """Alias de set_data pour compatibilité."""
        self.set_data(rows)

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(MARCHES_TRANCHES_COLUMNS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row_data = self.rows[index.row()]
        key, _ = MARCHES_TRANCHES_COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            value = row_data.get(key)

            # Formatage des montants
            if key in ("montant_initial_tranche", "service_fait_tranche", "paye_tranche"):
                if value is not None:
                    try:
                        return f"{float(value):,.2f} €".replace(",", " ")
                    except:
                        return str(value)

            # Formatage du pourcentage
            if key == "pourcent_consomme_tranche":
                if value is not None:
                    try:
                        return f"{float(value):.1f} %"
                    except:
                        return str(value)

            # Texte standard
            return str(value) if value is not None else ""

        if role == Qt.TextAlignmentRole:
            # Alignement à droite pour les montants et pourcentages
            if key in ("montant_initial_tranche", "service_fait_tranche", "paye_tranche",
                      "pourcent_consomme_tranche"):
                return Qt.AlignRight | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.BackgroundRole:
            # Coloration selon le pourcentage consommé de la tranche
            pourcent = row_data.get("pourcent_consomme_tranche", 0)
            if pourcent >= 100:
                return QBrush(QColor("#ffe0e0"))  # Rouge si dépassé
            elif pourcent >= 90:
                return QBrush(QColor("#fff3cd"))  # Orange si proche
            elif pourcent >= 50:
                return QBrush(QColor("#fff9e6"))  # Jaune léger si en cours
            else:
                return QBrush(QColor("#e0ffe0"))  # Vert si peu consommé

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return MARCHES_TRANCHES_COLUMNS[section][1]
        return section + 1


# ============== PROXY POUR FILTRES ET TRI ==============

class MarchesGlobauxProxy(QSortFilterProxyModel):
    """Proxy pour le tri et le filtrage de la vision globale."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.filter_marche = ""
        self.filter_fournisseur = ""

    def setMarcheFilter(self, text):
        self.filter_marche = text or ""
        self.invalidateFilter()

    def setFournisseurFilter(self, text):
        self.filter_fournisseur = text or ""
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        src = self.sourceModel()
        if not src or source_row >= len(src.rows):
            return True

        row = src.rows[source_row]

        # Filtre par marché
        if self.filter_marche:
            marche = str(row.get("marche", "")).lower()
            if self.filter_marche.lower() not in marche:
                return False

        # Filtre par fournisseur
        if self.filter_fournisseur:
            fournisseur = str(row.get("fournisseur", "")).lower()
            if self.filter_fournisseur.lower() not in fournisseur:
                return False

        return True

    def lessThan(self, left, right):
        """Tri personnalisé pour chaque colonne."""
        src = self.sourceModel()
        col = left.column()
        key, _ = MARCHES_GLOBAUX_COLUMNS[col]

        lv = src.rows[left.row()].get(key)
        rv = src.rows[right.row()].get(key)

        # Gestion des valeurs None
        if lv is None:
            lv = ""
        if rv is None:
            rv = ""

        # Tri numérique pour les montants et pourcentages
        if key in ("montant_initial_marche", "service_fait_cumule", "paye_cumule",
                  "reste_a_realiser", "reste_a_mandater", "pourcent_consomme"):
            try:
                lv = float(lv) if lv != "" else 0.0
                rv = float(rv) if rv != "" else 0.0
            except:
                pass

        return lv < rv


class MarchesTranchesProxy(QSortFilterProxyModel):
    """Proxy pour le tri et le filtrage de la vision détaillée."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.filter_marche = ""

    def setMarcheFilter(self, text):
        self.filter_marche = text or ""
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        src = self.sourceModel()
        if not src or source_row >= len(src.rows):
            return True

        row = src.rows[source_row]

        # Filtre par marché
        if self.filter_marche:
            marche = str(row.get("marche", "")).lower()
            if marche != self.filter_marche.lower():
                return False

        return True

    def lessThan(self, left, right):
        """Tri personnalisé pour chaque colonne."""
        src = self.sourceModel()
        col = left.column()
        key, _ = MARCHES_TRANCHES_COLUMNS[col]

        lv = src.rows[left.row()].get(key)
        rv = src.rows[right.row()].get(key)

        # Gestion des valeurs None
        if lv is None:
            lv = ""
        if rv is None:
            rv = ""

        # Tri numérique pour les montants et pourcentages
        if key in ("montant_initial_tranche", "service_fait_tranche", "paye_tranche",
                  "pourcent_consomme_tranche"):
            try:
                lv = float(lv) if lv != "" else 0.0
                rv = float(rv) if rv != "" else 0.0
            except:
                pass

        # Tri spécial pour les tranches (TF < TO1 < TO2...)
        if key == "tranche":
            # Gérer le tri numérique des tranches
            try:
                lv = float(lv) if lv != "" else -1
                rv = float(rv) if rv != "" else -1
            except:
                pass

        return lv < rv


# ============== MODÈLE POUR LA VISION OPÉRATIONS ==============

class OperationsTableModel(QAbstractTableModel):
    """Modèle pour la vision par opérations (1 ligne = 1 opération regroupant plusieurs lots)."""

    def __init__(self, rows=None):
        super().__init__()
        self.rows = rows or []

    def set_data(self, rows: List[Dict]):
        """Met à jour les données du modèle."""
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def refresh(self, rows: List[Dict]):
        """Alias de set_data pour compatibilité."""
        self.set_data(rows)

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(OPERATIONS_COLUMNS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row_data = self.rows[index.row()]
        key, _ = OPERATIONS_COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            value = row_data.get(key)

            # Formatage des montants
            if key in ("montant_initial_total", "service_fait_total", "paye_total",
                      "reste_a_realiser", "reste_a_mandater"):
                if value is not None:
                    try:
                        return f"{float(value):,.2f} €".replace(",", " ")
                    except:
                        return str(value)

            # Formatage du pourcentage
            if key == "pourcent_consomme":
                if value is not None:
                    try:
                        return f"{float(value):.1f} %"
                    except:
                        return str(value)

            # Formatage du nombre de lots
            if key == "nb_lots":
                return str(value) if value else "1"

            # Formatage des marchés (liste)
            if key == "marches":
                if isinstance(value, list):
                    return ", ".join(value)
                return str(value) if value else ""

            # Formatage des avenants
            if key == "nb_avenants_total":
                nb = row_data.get("nb_avenants_total", 0)
                if nb > 0:
                    return f"✓ {nb}"
                return ""

            # Texte standard
            return str(value) if value is not None else ""

        if role == Qt.ToolTipRole:
            # Tooltip pour afficher tous les marchés de l'opération
            if key == "marches":
                marches = row_data.get("marches", [])
                if isinstance(marches, list) and len(marches) > 1:
                    return "Lots: " + ", ".join(marches)

        if role == Qt.TextAlignmentRole:
            # Alignement à droite pour les montants et pourcentages
            if key in ("montant_initial_total", "service_fait_total", "paye_total",
                      "reste_a_realiser", "reste_a_mandater", "pourcent_consomme"):
                return Qt.AlignRight | Qt.AlignVCenter
            if key in ("nb_lots", "nb_avenants_total"):
                return Qt.AlignCenter | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.BackgroundRole:
            # Coloration selon le pourcentage consommé
            pourcent = row_data.get("pourcent_consomme", 0)
            if pourcent >= 100:
                return QBrush(QColor("#ffe0e0"))  # Rouge si dépassé
            elif pourcent >= 90:
                return QBrush(QColor("#fff3cd"))  # Orange si proche
            elif pourcent >= 50:
                return QBrush(QColor("#fff9e6"))  # Jaune léger si en cours
            else:
                return QBrush(QColor("#e0ffe0"))  # Vert si peu consommé

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return OPERATIONS_COLUMNS[section][1]
        return section + 1


# ============== MODÈLE POUR L'HISTORIQUE DES FACTURES ==============

class HistoriqueTableModel(QAbstractTableModel):
    """Modèle pour l'historique des factures et paiements."""

    def __init__(self, rows=None):
        super().__init__()
        self.rows = rows or []

    def set_data(self, rows: List[Dict]):
        """Met à jour les données du modèle."""
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def refresh(self, rows: List[Dict]):
        """Alias de set_data pour compatibilité."""
        self.set_data(rows)

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(HISTORIQUE_COLUMNS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row_data = self.rows[index.row()]
        key, _ = HISTORIQUE_COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            value = row_data.get(key)

            # Formatage des montants
            if key in ("montant_sf", "montant_ttc"):
                if value is not None:
                    try:
                        return f"{float(value):,.2f} €".replace(",", " ")
                    except:
                        return str(value)

            # Formatage des dates
            if key == "date_sf":
                if value:
                    # Convertir format Excel si nécessaire
                    try:
                        if isinstance(value, str):
                            # Essayer de parser différents formats
                            for fmt in ["%Y-%m-%d", "%d/%m/%Y"]:
                                try:
                                    dt = datetime.strptime(value, fmt)
                                    return dt.strftime("%d/%m/%Y")
                                except:
                                    pass
                        return str(value)
                    except:
                        return str(value)

            # Texte standard
            return str(value) if value is not None else ""

        if role == Qt.TextAlignmentRole:
            # Alignement à droite pour les montants
            if key in ("montant_sf", "montant_ttc"):
                return Qt.AlignRight | Qt.AlignVCenter
            # Centré pour les numéros et dates
            if key in ("num_facture", "num_mandat", "date_sf"):
                return Qt.AlignCenter | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.BackgroundRole:
            # Coloration selon le statut
            statut = row_data.get("statut", "")
            if "✅" in statut or "Payé" in statut:
                return QBrush(QColor("#d4edda"))  # Vert pour payé
            elif "📋" in statut or "Facturée" in statut:
                return QBrush(QColor("#fff3cd"))  # Orange pour facturé
            elif "⏳" in statut or "Service fait" in statut:
                return QBrush(QColor("#d1ecf1"))  # Bleu pour service fait
            elif "⚠️" in statut or "attente" in statut:
                return QBrush(QColor("#f8d7da"))  # Rouge pour attente

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return HISTORIQUE_COLUMNS[section][1]
        return section + 1


# ============== PROXY POUR OPÉRATIONS ==============

class OperationsProxy(QSortFilterProxyModel):
    """Proxy pour le tri et le filtrage de la vision opérations."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.filter_operation = ""

    def setOperationFilter(self, text):
        self.filter_operation = text or ""
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        src = self.sourceModel()
        if not src or source_row >= len(src.rows):
            return True

        row = src.rows[source_row]

        # Filtre par opération
        if self.filter_operation:
            operation = str(row.get("operation", "")).lower()
            if self.filter_operation.lower() not in operation:
                return False

        return True

    def lessThan(self, left, right):
        """Tri personnalisé pour chaque colonne."""
        src = self.sourceModel()
        col = left.column()
        key, _ = OPERATIONS_COLUMNS[col]

        lv = src.rows[left.row()].get(key)
        rv = src.rows[right.row()].get(key)

        # Gestion des valeurs None
        if lv is None:
            lv = ""
        if rv is None:
            rv = ""

        # Tri numérique pour les montants et pourcentages
        if key in ("montant_initial_total", "service_fait_total", "paye_total",
                  "reste_a_realiser", "reste_a_mandater", "pourcent_consomme"):
            try:
                lv = float(lv) if lv != "" else 0.0
                rv = float(rv) if rv != "" else 0.0
            except:
                pass

        # Tri numérique pour nb_lots
        if key == "nb_lots":
            try:
                lv = int(lv) if lv != "" else 0
                rv = int(rv) if rv != "" else 0
            except:
                pass

        return lv < rv


# ============== PROXY POUR HISTORIQUE ==============

class HistoriqueProxy(QSortFilterProxyModel):
    """Proxy pour le tri et le filtrage de l'historique."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.filter_marche = ""
        # Filtres par plage de dates
        self.filter_date_from = None  # date ou None
        self.filter_date_to = None    # date ou None

    def setMarcheFilter(self, text):
        self.filter_marche = text or ""
        self.invalidateFilter()

    def setDateFilter(self, date_from, date_to):
        """Filtre par plage de dates (date_sf). date_from/date_to: date ou None."""
        self.filter_date_from = date_from
        self.filter_date_to = date_to
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        src = self.sourceModel()
        if not src or source_row >= len(src.rows):
            return True

        row = src.rows[source_row]

        # Filtre par marché
        if self.filter_marche:
            marche = str(row.get("marche", "")).lower()
            if self.filter_marche.lower() not in marche:
                return False

        # Filtre par plage de dates (date_sf au format dd/mm/yyyy)
        if self.filter_date_from or self.filter_date_to:
            date_str = row.get("date_sf", "") or ""
            if date_str:
                try:
                    # Essayer dd/mm/yyyy puis yyyy-mm-dd
                    row_date = None
                    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                        try:
                            row_date = datetime.strptime(date_str[:10], fmt).date()
                            break
                        except ValueError:
                            continue
                    if row_date:
                        if self.filter_date_from and row_date < self.filter_date_from:
                            return False
                        if self.filter_date_to and row_date > self.filter_date_to:
                            return False
                except (ValueError, TypeError):
                    pass
            else:
                if self.filter_date_from or self.filter_date_to:
                    return False

        return True

    def lessThan(self, left, right):
        """Tri personnalisé pour chaque colonne."""
        src = self.sourceModel()
        col = left.column()
        key, _ = HISTORIQUE_COLUMNS[col]

        lv = src.rows[left.row()].get(key)
        rv = src.rows[right.row()].get(key)

        # Gestion des valeurs None
        if lv is None:
            lv = ""
        if rv is None:
            rv = ""

        # Tri numérique pour les montants
        if key in ("montant_sf", "montant_ttc"):
            try:
                lv = float(lv) if lv != "" else 0.0
                rv = float(rv) if rv != "" else 0.0
            except:
                pass

        # Tri par date
        if key == "date_sf":
            try:
                if isinstance(lv, str):
                    for fmt in ["%Y-%m-%d", "%d/%m/%Y"]:
                        try:
                            lv = datetime.strptime(lv, fmt)
                            break
                        except:
                            pass
                if isinstance(rv, str):
                    for fmt in ["%Y-%m-%d", "%d/%m/%Y"]:
                        try:
                            rv = datetime.strptime(rv, fmt)
                            break
                        except:
                            pass
            except:
                pass

        return lv < rv
