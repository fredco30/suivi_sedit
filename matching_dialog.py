"""Assistant de rapprochement commandes <-> factures (UI Qt).

Etape 3 : interface utilisateur sans IA. Le bouton "Demander avis IA" est
present mais desactive ; il sera active a l'etape 4.

L'UI s'appuie integralement sur matching_module (LinkRepository + MatchingEngine)
et sur Database.recompute_facturation() pour la coherence des totaux.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from matching_module import (
    DIAG_DOUBLON_PROBABLE,
    DIAG_EN_COURS,
    DIAG_OK,
    DIAG_OK_RAPPROCHE,
    DIAG_OUBLI_PROBABLE,
    DIAG_RAPPROCHEMENT_SUGGERE,
    DIAG_RECENT,
    LinkRepository,
    MatchingEngine,
    SEVERITE_BY_DIAG,
)


# ============================================================================
# Constantes UI
# ============================================================================

DIAG_LABEL = {
    DIAG_OK: "OK",
    DIAG_OK_RAPPROCHE: "OK (rapproche)",
    DIAG_RECENT: "Recent",
    DIAG_EN_COURS: "En cours",
    DIAG_RAPPROCHEMENT_SUGGERE: "Rapprochement suggere",
    DIAG_OUBLI_PROBABLE: "Oubli probable",
    DIAG_DOUBLON_PROBABLE: "Doublon probable",
}

DIAG_EMOJI = {
    DIAG_OK: "🟢",
    DIAG_OK_RAPPROCHE: "🟢",
    DIAG_RECENT: "⚪",
    DIAG_EN_COURS: "🔵",
    DIAG_RAPPROCHEMENT_SUGGERE: "🟡",
    DIAG_OUBLI_PROBABLE: "🔴",
    DIAG_DOUBLON_PROBABLE: "🟠",
}

# Diagnostics actionnables (coches par defaut a l'ouverture)
DIAG_DEFAULT_CHECKED = {
    DIAG_RAPPROCHEMENT_SUGGERE,
    DIAG_OUBLI_PROBABLE,
    DIAG_DOUBLON_PROBABLE,
}

# Ordre d'affichage : severite descendante puis nom
DIAG_DISPLAY_ORDER = [
    DIAG_OUBLI_PROBABLE,
    DIAG_RAPPROCHEMENT_SUGGERE,
    DIAG_DOUBLON_PROBABLE,
    DIAG_EN_COURS,
    DIAG_RECENT,
    DIAG_OK_RAPPROCHE,
    DIAG_OK,
]

DEFAULT_DISMISS_DAYS = 30


def _fmt_eur(v) -> str:
    try:
        return f"{float(v or 0):,.2f} €".replace(",", " ").replace(".", ",").replace(" ", " ")
    except (TypeError, ValueError):
        return "—"


def _fmt_date(s: Optional[str]) -> str:
    if not s:
        return "—"
    return s[:10]


# ============================================================================
# MatchingDialog
# ============================================================================

class MatchingDialog(QDialog):
    """Assistant de rapprochement avec pre-classification deterministe."""

    # Colonnes de la table commandes (gauche)
    COL_DIAG = 0
    COL_NUM = 1
    COL_FOURN = 2
    COL_TTC = 3
    COL_RESTE = 4
    COL_AGE = 5
    COL_MARCHE = 6
    COL_CMD_ID = 7  # cachee, sert a retrouver l'id depuis la selection

    LEFT_HEADERS = ["Diag", "N° Cmd", "Fournisseur", "TTC", "Reste a facturer",
                    "Age (j)", "Marche", "_id"]

    # Colonnes de la table candidats (droite)
    CC_CHECK = 0
    CC_NUM = 1
    CC_CODE_MVT = 2
    CC_DATE = 3
    CC_FOURN = 4
    CC_MARCHE = 5
    CC_MSF = 6
    CC_ALLOUE = 7
    CC_LIBELLE = 8
    CC_FACT_ID = 9  # cachee

    CAND_HEADERS = ["✓", "N° facture", "Code mvt", "Date", "Fournisseur",
                    "Marche", "Service fait", "A allouer", "Libelle", "_id"]

    LK_NUM = 0
    LK_SOURCE = 1
    LK_DATE = 2
    LK_MONTANT = 3
    LK_ACTION = 4
    LK_LINK_ID = 5  # cachee

    LINK_HEADERS = ["N° facture", "Source", "Cree le", "Montant alloue",
                    "Action", "_id"]

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.link_repo = LinkRepository(self.db.conn)
        self.engine = MatchingEngine(self.db.conn, self.link_repo)
        self.current_cmd_id: Optional[int] = None
        # Cache des donnees commandes affichees, pour filtrage sans re-requeter
        self._all_rows: list[dict] = []

        self.setWindowTitle("Assistant de rapprochement commandes ↔ factures")
        self.resize(1400, 850)
        self._restore_geometry()

        self._build_ui()
        # Premier remplissage
        self.refresh_all()

    # ------------------------------------------------------------------
    # Construction de l'UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        # 1. Bandeau de filtres
        outer.addWidget(self._build_filters())

        # 2. Splitter principal
        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 4)
        splitter.setSizes([520, 880])
        outer.addWidget(splitter, 1)

        # 3. Status bar + Fermer
        bottom = QHBoxLayout()
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #555;")
        bottom.addWidget(self.status_label, 1)
        btn_close = QPushButton("Fermer")
        btn_close.clicked.connect(self.accept)
        bottom.addWidget(btn_close)
        outer.addLayout(bottom)

    def _build_filters(self) -> QWidget:
        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(10)

        # Checkboxes de diagnostic
        layout.addWidget(QLabel("Diagnostics :"))
        self.diag_checks = {}
        for diag in DIAG_DISPLAY_ORDER:
            cb = QCheckBox(f"{DIAG_EMOJI[diag]} {DIAG_LABEL[diag]}")
            cb.setChecked(diag in DIAG_DEFAULT_CHECKED)
            cb.stateChanged.connect(self._apply_filters)
            self.diag_checks[diag] = cb
            layout.addWidget(cb)

        layout.addSpacing(12)

        # Combo fournisseur
        layout.addWidget(QLabel("Fournisseur :"))
        self.fourn_combo = QComboBox()
        self.fourn_combo.setMinimumWidth(220)
        self.fourn_combo.addItem("[Tous]", "")
        self.fourn_combo.currentIndexChanged.connect(self._apply_filters)
        layout.addWidget(self.fourn_combo)

        # Age min
        layout.addWidget(QLabel("Age min (j) :"))
        self.age_min_spin = QSpinBox()
        self.age_min_spin.setRange(0, 3650)
        self.age_min_spin.setValue(0)
        self.age_min_spin.setMaximumWidth(80)
        self.age_min_spin.valueChanged.connect(self._apply_filters)
        layout.addWidget(self.age_min_spin)

        # Afficher ecartees
        self.show_dismissed_cb = QCheckBox("Afficher ecartees")
        self.show_dismissed_cb.setChecked(False)
        self.show_dismissed_cb.stateChanged.connect(self._apply_filters)
        layout.addWidget(self.show_dismissed_cb)

        layout.addStretch(1)

        # Bouton Refresh
        btn_refresh = QPushButton("🔄 Recalculer")
        btn_refresh.setToolTip("Relance le diagnostic complet sur toutes les commandes")
        btn_refresh.clicked.connect(self.refresh_all)
        layout.addWidget(btn_refresh)

        return frame

    def _build_left_panel(self) -> QWidget:
        wrap = QWidget()
        v = QVBoxLayout(wrap)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(4)

        self.counts_label = QLabel("")
        self.counts_label.setStyleSheet("font-weight: bold;")
        v.addWidget(self.counts_label)

        self.cmd_table = QTableWidget(0, len(self.LEFT_HEADERS))
        self.cmd_table.setHorizontalHeaderLabels(self.LEFT_HEADERS)
        self.cmd_table.setColumnHidden(self.COL_CMD_ID, True)
        self.cmd_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.cmd_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.cmd_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.cmd_table.setAlternatingRowColors(True)
        self.cmd_table.verticalHeader().setVisible(False)
        # Sorting desactive : on respecte l'ordre Python (severite desc, montant desc)
        # qui matche la spec 8.4. Reactiver permettrait au header click de casser
        # cet ordre (le tri Qt sur l'emoji ne reflete pas la severite).
        self.cmd_table.setSortingEnabled(False)
        h = self.cmd_table.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.Interactive)
        h.setStretchLastSection(False)
        h.resizeSection(self.COL_DIAG, 40)
        h.resizeSection(self.COL_NUM, 100)
        h.resizeSection(self.COL_FOURN, 200)
        h.resizeSection(self.COL_TTC, 110)
        h.resizeSection(self.COL_RESTE, 110)
        h.resizeSection(self.COL_AGE, 60)
        h.setSectionResizeMode(self.COL_MARCHE, QHeaderView.Stretch)
        self.cmd_table.itemSelectionChanged.connect(self._on_commande_selected)
        v.addWidget(self.cmd_table, 1)

        return wrap

    def _build_right_panel(self) -> QWidget:
        wrap = QWidget()
        v = QVBoxLayout(wrap)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(6)

        # Bloc details
        gb_details = QGroupBox("Details commande")
        d_layout = QVBoxLayout(gb_details)
        self.details_label = QLabel("Sélectionnez une commande à gauche.")
        self.details_label.setWordWrap(True)
        self.details_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        d_layout.addWidget(self.details_label)
        v.addWidget(gb_details)

        # Bloc candidats
        gb_cands = QGroupBox("Factures candidates")
        c_layout = QVBoxLayout(gb_cands)
        self.cand_table = QTableWidget(0, len(self.CAND_HEADERS))
        self.cand_table.setHorizontalHeaderLabels(self.CAND_HEADERS)
        self.cand_table.setColumnHidden(self.CC_FACT_ID, True)
        self.cand_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.cand_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.cand_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.cand_table.setAlternatingRowColors(True)
        self.cand_table.verticalHeader().setVisible(False)
        ch = self.cand_table.horizontalHeader()
        ch.setSectionResizeMode(QHeaderView.Interactive)
        ch.resizeSection(self.CC_CHECK, 30)
        ch.resizeSection(self.CC_NUM, 110)
        ch.resizeSection(self.CC_CODE_MVT, 100)
        ch.resizeSection(self.CC_DATE, 90)
        ch.resizeSection(self.CC_FOURN, 180)
        ch.resizeSection(self.CC_MARCHE, 100)
        ch.resizeSection(self.CC_MSF, 110)
        ch.resizeSection(self.CC_ALLOUE, 110)
        ch.setSectionResizeMode(self.CC_LIBELLE, QHeaderView.Stretch)
        c_layout.addWidget(self.cand_table)
        v.addWidget(gb_cands, 1)

        # Bloc liens existants
        gb_links = QGroupBox("Liens manuels existants")
        l_layout = QVBoxLayout(gb_links)
        self.link_table = QTableWidget(0, len(self.LINK_HEADERS))
        self.link_table.setHorizontalHeaderLabels(self.LINK_HEADERS)
        self.link_table.setColumnHidden(self.LK_LINK_ID, True)
        self.link_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.link_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.link_table.setAlternatingRowColors(True)
        self.link_table.verticalHeader().setVisible(False)
        self.link_table.setMaximumHeight(140)
        lh = self.link_table.horizontalHeader()
        lh.setSectionResizeMode(QHeaderView.Interactive)
        lh.resizeSection(self.LK_NUM, 110)
        lh.resizeSection(self.LK_SOURCE, 90)
        lh.resizeSection(self.LK_DATE, 140)
        lh.resizeSection(self.LK_MONTANT, 110)
        lh.setSectionResizeMode(self.LK_ACTION, QHeaderView.Stretch)
        l_layout.addWidget(self.link_table)
        v.addWidget(gb_links)

        # Boutons d'action
        actions_row = QHBoxLayout()
        self.btn_validate = QPushButton("✓ Valider sélection")
        self.btn_validate.setToolTip(
            "Cree des liens manuels pour les factures cochees, avec le montant indique"
        )
        self.btn_validate.clicked.connect(self._on_validate_clicked)

        self.btn_doublon = QPushButton("⚠ Marquer doublon admin")
        self.btn_doublon.setToolTip(
            "Marque cette commande comme un doublon administratif (statut_metier=DOUBLON_ADMIN)"
        )
        self.btn_doublon.clicked.connect(self._on_mark_doublon_clicked)

        self.btn_no_match = QPushButton(f"⨯ Pas de match ({DEFAULT_DISMISS_DAYS}j)")
        self.btn_no_match.setToolTip(
            f"Ecarte cette commande de l'assistant pendant {DEFAULT_DISMISS_DAYS} jours"
        )
        self.btn_no_match.clicked.connect(self._on_no_match_clicked)

        self.btn_ai = QPushButton("🤖 Demander avis IA")
        self.btn_ai.setEnabled(False)
        self.btn_ai.setToolTip(
            "Configurer la cle API DeepSeek (etape 4) pour activer l'assistance IA"
        )

        for btn in (self.btn_validate, self.btn_doublon, self.btn_no_match, self.btn_ai):
            btn.setEnabled(False)
            actions_row.addWidget(btn)
        actions_row.addStretch(1)
        v.addLayout(actions_row)

        return wrap

    # ------------------------------------------------------------------
    # Chargement et rafraichissement des donnees
    # ------------------------------------------------------------------

    def refresh_all(self):
        """Recalcule diagnostic + montants, puis recharge la table de gauche."""
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            # Recompute facturation puis diagnostic, dans cet ordre.
            self.db.recompute_facturation()
            counts = self.engine.diagnose_all_commandes()
            self._load_rows()
            self._populate_fournisseurs_combo()
            self._apply_filters()
            self._update_counts_label(counts)
            self.status_label.setText("Diagnostic recalcule.")
        finally:
            QApplication.restoreOverrideCursor()

    def _load_rows(self):
        """Charge en memoire toutes les lignes commandes + diagnostic."""
        cur = self.db.conn.cursor()
        cur.execute(
            """
            SELECT c.id AS cmd_id,
                   c.num_commande, c.fournisseur, c.libelle,
                   c.date_commande, c.marche, c.service_emetteur,
                   c.montant_ttc, c.reste_a_facturer,
                   c.statut_facturation, c.statut_metier,
                   d.diagnostic, d.severite, d.age_jours,
                   d.candidates_count, d.candidates_same_marche,
                   d.montant_candidates_total, d.dismissed_until, d.notes
            FROM commandes c
            LEFT JOIN commande_diagnostic d ON d.commande_id = c.id
            """
        )
        rows = [dict(r) for r in cur.fetchall()]
        # Tri en Python : severite desc, puis montant desc
        rows.sort(
            key=lambda r: (
                -(r.get("severite") if r.get("severite") is not None else -1),
                -(r.get("montant_ttc") or 0.0),
            )
        )
        self._all_rows = rows

    def _populate_fournisseurs_combo(self):
        prev = self.fourn_combo.currentData()
        self.fourn_combo.blockSignals(True)
        self.fourn_combo.clear()
        self.fourn_combo.addItem("[Tous]", "")
        seen = set()
        for r in self._all_rows:
            f = (r.get("fournisseur") or "").strip()
            if f and f not in seen:
                seen.add(f)
        for f in sorted(seen):
            self.fourn_combo.addItem(f, f)
        if prev:
            idx = self.fourn_combo.findData(prev)
            if idx >= 0:
                self.fourn_combo.setCurrentIndex(idx)
        self.fourn_combo.blockSignals(False)

    def _update_counts_label(self, counts: dict):
        total = sum(counts.values())
        bits = [f"{total} commandes"]
        for diag in DIAG_DISPLAY_ORDER:
            n = counts.get(diag, 0)
            if n:
                bits.append(f"{DIAG_EMOJI[diag]} {DIAG_LABEL[diag]} : {n}")
        self.counts_label.setText("   |   ".join(bits))

    # ------------------------------------------------------------------
    # Filtres et remplissage table gauche
    # ------------------------------------------------------------------

    def _apply_filters(self):
        active_diags = {d for d, cb in self.diag_checks.items() if cb.isChecked()}
        fourn_filter = self.fourn_combo.currentData() or ""
        age_min = self.age_min_spin.value()
        show_dismissed = self.show_dismissed_cb.isChecked()
        today_iso = datetime.now().date().isoformat()

        filtered = []
        for r in self._all_rows:
            diag = r.get("diagnostic")
            if diag is None:
                # Commande sans diagnostic (cas extreme : table vide). On l'inclut
                # uniquement si tous les diags actionnables sont coches (filet).
                if not active_diags:
                    filtered.append(r)
                continue
            if diag not in active_diags:
                continue
            if fourn_filter and (r.get("fournisseur") or "").strip() != fourn_filter:
                continue
            age = r.get("age_jours")
            if age is not None and age < age_min:
                continue
            if not show_dismissed:
                du = r.get("dismissed_until")
                if du and du > today_iso:
                    continue
            filtered.append(r)

        self._fill_left_table(filtered)
        self.status_label.setText(
            f"{len(filtered)} commande(s) affichee(s) sur {len(self._all_rows)}."
        )

    def _fill_left_table(self, rows: list[dict]):
        self.cmd_table.setRowCount(0)
        self.cmd_table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            diag = r.get("diagnostic") or ""
            emoji = DIAG_EMOJI.get(diag, "")
            it_diag = QTableWidgetItem(emoji)
            it_diag.setToolTip(DIAG_LABEL.get(diag, diag))
            it_diag.setTextAlignment(Qt.AlignCenter)
            sev = r.get("severite") if r.get("severite") is not None else -1
            it_diag.setData(Qt.UserRole, sev)
            self.cmd_table.setItem(i, self.COL_DIAG, it_diag)

            self.cmd_table.setItem(i, self.COL_NUM, QTableWidgetItem(r.get("num_commande") or ""))
            self.cmd_table.setItem(i, self.COL_FOURN, QTableWidgetItem(r.get("fournisseur") or ""))

            ttc = float(r.get("montant_ttc") or 0.0)
            it_ttc = QTableWidgetItem(_fmt_eur(ttc))
            it_ttc.setData(Qt.UserRole, ttc)
            it_ttc.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.cmd_table.setItem(i, self.COL_TTC, it_ttc)

            reste = float(r.get("reste_a_facturer") or 0.0)
            it_reste = QTableWidgetItem(_fmt_eur(reste))
            it_reste.setData(Qt.UserRole, reste)
            it_reste.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.cmd_table.setItem(i, self.COL_RESTE, it_reste)

            age = r.get("age_jours")
            it_age = QTableWidgetItem("" if age is None else str(age))
            it_age.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.cmd_table.setItem(i, self.COL_AGE, it_age)

            self.cmd_table.setItem(i, self.COL_MARCHE, QTableWidgetItem(r.get("marche") or ""))
            self.cmd_table.setItem(i, self.COL_CMD_ID, QTableWidgetItem(str(r["cmd_id"])))

            # Coloration legere selon severite
            if sev >= 3:
                color = QColor(255, 230, 230)
            elif sev == 2:
                color = QColor(255, 244, 220)
            elif sev == 1:
                color = QColor(240, 248, 255)
            else:
                color = None
            if color is not None:
                for c in range(self.cmd_table.columnCount() - 1):  # exclut colonne id cachee
                    item = self.cmd_table.item(i, c)
                    if item is not None:
                        item.setBackground(QBrush(color))

    # ------------------------------------------------------------------
    # Selection commande -> panneau droit
    # ------------------------------------------------------------------

    def _on_commande_selected(self):
        rows = self.cmd_table.selectionModel().selectedRows()
        if not rows:
            self.current_cmd_id = None
            self._clear_right_panel()
            return
        row = rows[0].row()
        id_item = self.cmd_table.item(row, self.COL_CMD_ID)
        if id_item is None:
            return
        try:
            cmd_id = int(id_item.text())
        except ValueError:
            return
        self.current_cmd_id = cmd_id
        self._refresh_right_panel(cmd_id)

    def _clear_right_panel(self):
        self.details_label.setText("Sélectionnez une commande à gauche.")
        self.cand_table.setRowCount(0)
        self.link_table.setRowCount(0)
        for btn in (self.btn_validate, self.btn_doublon, self.btn_no_match):
            btn.setEnabled(False)

    def _refresh_right_panel(self, cmd_id: int):
        cur = self.db.conn.cursor()
        cur.execute("SELECT * FROM commandes WHERE id = ?", (cmd_id,))
        cmd = cur.fetchone()
        if cmd is None:
            self._clear_right_panel()
            return
        cmd_d = dict(cmd)

        # Bloc details : HTML simple
        already_alloue = self.link_repo.total_alloue_for_commande(cmd_id)
        details_lines = [
            f"<b>{cmd_d.get('num_commande') or '—'}</b> &nbsp; "
            f"{cmd_d.get('fournisseur') or '—'}",
            f"Libelle : {cmd_d.get('libelle') or '—'}",
            f"Marche : {cmd_d.get('marche') or '—'}"
            f" &nbsp;|&nbsp; Service : {cmd_d.get('service_emetteur') or '—'}",
            f"Date : {_fmt_date(cmd_d.get('date_commande'))}"
            f" &nbsp;|&nbsp; TTC : {_fmt_eur(cmd_d.get('montant_ttc'))}"
            f" &nbsp;|&nbsp; Reste : {_fmt_eur(cmd_d.get('reste_a_facturer'))}"
            f" &nbsp;|&nbsp; Statut fact. : {cmd_d.get('statut_facturation') or '—'}",
        ]
        if cmd_d.get("statut_metier"):
            details_lines.append(
                f"<b>Statut metier :</b> {cmd_d['statut_metier']}"
            )
        if already_alloue > 0:
            details_lines.append(
                f"Deja alloue via liens manuels : {_fmt_eur(already_alloue)}"
            )
        self.details_label.setText("<br>".join(details_lines))

        # Candidats
        candidates = self.engine.find_candidates(cmd_id)
        self._fill_candidates_table(candidates, cmd_d)

        # Liens existants
        links = self.link_repo.list_links_for_commande(cmd_id)
        self._fill_links_table(links)

        # Activer les boutons. Validate seulement s'il y a au moins un candidat.
        has_candidates = bool(candidates)
        is_doublon = (cmd_d.get("statut_metier") == "DOUBLON_ADMIN")
        self.btn_validate.setEnabled(has_candidates and not is_doublon)
        self.btn_doublon.setEnabled(not is_doublon)
        self.btn_no_match.setEnabled(True)

    def _fill_candidates_table(self, candidates: list, cmd_d: dict):
        self.cand_table.setRowCount(0)
        if not candidates:
            return
        reste = float(cmd_d.get("reste_a_facturer") or 0.0)
        if reste <= 0:
            reste = float(cmd_d.get("montant_ttc") or 0.0)

        self.cand_table.setRowCount(len(candidates))
        for i, c in enumerate(candidates):
            # Checkbox
            chk = QTableWidgetItem("")
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            chk.setCheckState(Qt.Unchecked)
            chk.setTextAlignment(Qt.AlignCenter)
            self.cand_table.setItem(i, self.CC_CHECK, chk)

            self.cand_table.setItem(i, self.CC_NUM,
                                    QTableWidgetItem(c.get("num_facture") or ""))
            self.cand_table.setItem(i, self.CC_CODE_MVT,
                                    QTableWidgetItem(c.get("code_mouvement") or ""))
            self.cand_table.setItem(i, self.CC_DATE,
                                    QTableWidgetItem(_fmt_date(c.get("date_facture"))))
            self.cand_table.setItem(i, self.CC_FOURN,
                                    QTableWidgetItem(c.get("fournisseur") or ""))
            self.cand_table.setItem(i, self.CC_MARCHE,
                                    QTableWidgetItem(c.get("marche") or ""))

            msf = float(c.get("montant_service_fait") or 0.0)
            it_msf = QTableWidgetItem(_fmt_eur(msf))
            it_msf.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.cand_table.setItem(i, self.CC_MSF, it_msf)

            # SpinBox montant a allouer : default = min(MSF, reste)
            spin = QDoubleSpinBox()
            spin.setRange(0.0, max(msf, 1e9))
            spin.setDecimals(2)
            spin.setSingleStep(10.0)
            spin.setSuffix(" €")
            default_val = min(msf, reste) if reste > 0 else msf
            # On ne peut pas excéder MSF (sinon add_link refusera)
            default_val = max(0.0, min(default_val, msf))
            spin.setValue(default_val)
            self.cand_table.setCellWidget(i, self.CC_ALLOUE, spin)

            self.cand_table.setItem(i, self.CC_LIBELLE,
                                    QTableWidgetItem((c.get("libelle") or "")[:120]))
            self.cand_table.setItem(i, self.CC_FACT_ID,
                                    QTableWidgetItem(str(c.get("id") or "")))

    def _fill_links_table(self, links: list):
        self.link_table.setRowCount(0)
        if not links:
            return
        cur = self.db.conn.cursor()
        self.link_table.setRowCount(len(links))
        for i, lk in enumerate(links):
            cur.execute("SELECT num_facture FROM factures WHERE id = ?", (lk["facture_id"],))
            r = cur.fetchone()
            num_fact = r["num_facture"] if r else f"id={lk['facture_id']}"

            self.link_table.setItem(i, self.LK_NUM, QTableWidgetItem(num_fact or ""))
            self.link_table.setItem(i, self.LK_SOURCE, QTableWidgetItem(lk.get("source") or ""))
            self.link_table.setItem(i, self.LK_DATE, QTableWidgetItem(lk.get("created_at") or ""))
            it_m = QTableWidgetItem(_fmt_eur(lk.get("montant_alloue")))
            it_m.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.link_table.setItem(i, self.LK_MONTANT, it_m)

            btn = QPushButton("Supprimer")
            link_id = lk["id"]
            btn.clicked.connect(lambda _=False, lid=link_id: self._on_remove_link_clicked(lid))
            self.link_table.setCellWidget(i, self.LK_ACTION, btn)

            self.link_table.setItem(i, self.LK_LINK_ID, QTableWidgetItem(str(link_id)))

    # ------------------------------------------------------------------
    # Actions utilisateur
    # ------------------------------------------------------------------

    def _on_validate_clicked(self):
        if self.current_cmd_id is None:
            return
        # Recolte des lignes cochees
        to_link = []
        for i in range(self.cand_table.rowCount()):
            chk = self.cand_table.item(i, self.CC_CHECK)
            if chk is None or chk.checkState() != Qt.Checked:
                continue
            id_item = self.cand_table.item(i, self.CC_FACT_ID)
            spin = self.cand_table.cellWidget(i, self.CC_ALLOUE)
            if id_item is None or spin is None:
                continue
            try:
                fact_id = int(id_item.text())
            except ValueError:
                continue
            montant = float(spin.value())
            if montant <= 0:
                continue
            to_link.append((fact_id, montant))

        if not to_link:
            QMessageBox.information(
                self,
                "Aucune sélection",
                "Cochez au moins une facture candidate et indiquez un montant > 0.",
            )
            return

        cmd_id = self.current_cmd_id
        ok_count = 0
        errors = []
        for fact_id, montant in to_link:
            try:
                self.link_repo.add_link(
                    commande_id=cmd_id,
                    facture_id=fact_id,
                    montant_alloue=montant,
                    source="manual",
                    created_by="user",
                )
                ok_count += 1
            except Exception as e:
                errors.append(f"facture id={fact_id} : {e}")

        # Recompute global puis re-diagnose, puis refresh UI
        self._recompute_after_action()

        if errors:
            QMessageBox.warning(
                self,
                "Liens partiellement crees",
                f"{ok_count} lien(s) cree(s), {len(errors)} en erreur :\n\n" + "\n".join(errors),
            )
        else:
            self.status_label.setText(f"{ok_count} lien(s) manuel(s) cree(s).")

    def _on_mark_doublon_clicked(self):
        if self.current_cmd_id is None:
            return
        ans = QMessageBox.question(
            self,
            "Marquer doublon administratif",
            "Confirmer le marquage de cette commande comme DOUBLON_ADMIN ?\n\n"
            "Le statut deviendra 'Annulée' et le reste a facturer sera mis a 0.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ans != QMessageBox.Yes:
            return
        cur = self.db.conn.cursor()
        cur.execute(
            "UPDATE commandes SET statut_metier = 'DOUBLON_ADMIN' WHERE id = ?",
            (self.current_cmd_id,),
        )
        self.db.conn.commit()
        self._recompute_after_action()
        self.status_label.setText("Commande marquee comme doublon administratif.")

    def _on_no_match_clicked(self):
        if self.current_cmd_id is None:
            return
        until = (datetime.now().date() + timedelta(days=DEFAULT_DISMISS_DAYS)).isoformat()
        cur = self.db.conn.cursor()
        cur.execute(
            "SELECT 1 FROM commande_diagnostic WHERE commande_id = ?",
            (self.current_cmd_id,),
        )
        existing = cur.fetchone()
        now_iso = datetime.now().isoformat(timespec="seconds")
        if existing:
            cur.execute(
                "UPDATE commande_diagnostic SET dismissed_until = ?, notes = ? "
                "WHERE commande_id = ?",
                (until, "Ecartee par l'utilisateur (Pas de match)", self.current_cmd_id),
            )
        else:
            # Ligne de diagnostic pas encore creee : on en cree une minimale.
            cur.execute(
                """
                INSERT INTO commande_diagnostic
                (commande_id, diagnostic, severite, dismissed_until, notes,
                 last_diagnostic_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    self.current_cmd_id,
                    DIAG_OUBLI_PROBABLE,
                    SEVERITE_BY_DIAG[DIAG_OUBLI_PROBABLE],
                    until,
                    "Ecartee par l'utilisateur (Pas de match)",
                    now_iso,
                ),
            )
        self.db.conn.commit()
        # Pas besoin de recompute_facturation : aucune modif financiere.
        # Mais on refresh la liste pour qu'elle disparaisse de la vue.
        self._load_rows()
        self._apply_filters()
        # Le panneau droit reste affiche pour info, mais on perd la selection
        self.current_cmd_id = None
        self._clear_right_panel()
        self.status_label.setText(
            f"Commande ecartee jusqu'au {until} (decochez 'Afficher ecartees' pour la voir)."
        )

    def _on_remove_link_clicked(self, link_id: int):
        ans = QMessageBox.question(
            self,
            "Supprimer le lien",
            f"Confirmer la suppression du lien manuel #{link_id} ?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ans != QMessageBox.Yes:
            return
        ok = self.link_repo.remove_link(link_id)
        if not ok:
            QMessageBox.warning(self, "Erreur", "Lien introuvable ou deja supprime.")
            return
        self._recompute_after_action()
        self.status_label.setText(f"Lien #{link_id} supprime.")

    def _recompute_after_action(self):
        """Apres une action utilisateur : recompute global + re-diagnose + refresh UI."""
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.db.recompute_facturation()
            counts = self.engine.diagnose_all_commandes()
            self._load_rows()
            self._update_counts_label(counts)
            self._apply_filters()
            # Re-selectionner la commande courante si toujours visible
            if self.current_cmd_id is not None:
                self._reselect_current_cmd()
                self._refresh_right_panel(self.current_cmd_id)
        finally:
            QApplication.restoreOverrideCursor()

    def _reselect_current_cmd(self):
        if self.current_cmd_id is None:
            return
        for i in range(self.cmd_table.rowCount()):
            it = self.cmd_table.item(i, self.COL_CMD_ID)
            if it is not None and it.text() == str(self.current_cmd_id):
                self.cmd_table.selectRow(i)
                return

    # ------------------------------------------------------------------
    # Persistance geometrie
    # ------------------------------------------------------------------

    def _restore_geometry(self):
        try:
            settings = QSettings("suivi_sedit", "matching_dialog")
            geom = settings.value("geometry")
            if geom:
                self.restoreGeometry(geom)
        except Exception:
            pass

    def closeEvent(self, event):
        try:
            settings = QSettings("suivi_sedit", "matching_dialog")
            settings.setValue("geometry", self.saveGeometry())
        except Exception:
            pass
        super().closeEvent(event)
