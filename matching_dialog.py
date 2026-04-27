"""Assistant de rapprochement commandes <-> factures (UI Qt).

Etape 3 : interface utilisateur sans IA. Le bouton "Demander avis IA" est
present mais desactive ; il sera active a l'etape 4.

L'UI s'appuie integralement sur matching_module (LinkRepository + MatchingEngine)
et sur Database.recompute_facturation() pour la coherence des totaux.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

from PyQt5.QtCore import Qt, QSettings, QThread, pyqtSignal
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
    QProgressDialog,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from matching_module import (
    ACTION_AUTO_DOUBLON,
    ACTION_AUTO_LINKED,
    ACTION_IGNORED,
    ACTION_SUGGESTION,
    DIAG_DOUBLON_PROBABLE,
    DIAG_EN_COURS,
    DIAG_OK,
    DIAG_OK_RAPPROCHE,
    DIAG_OUBLI_PROBABLE,
    DIAG_RAPPROCHEMENT_SUGGERE,
    DIAG_RECENT,
    DeepSeekClient,
    LinkRepository,
    MatchingEngine,
    SEVERITE_BY_DIAG,
    apply_ai_decision,
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
        # Etat IA
        self._last_ai_response: Optional[dict] = None
        self._ai_worker: Optional["AiWorker"] = None

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

        # Marche
        layout.addWidget(QLabel("Marche :"))
        self.marche_combo = QComboBox()
        self.marche_combo.setMinimumWidth(140)
        self.marche_combo.addItem("[Tous]", "")
        self.marche_combo.currentIndexChanged.connect(self._apply_filters)
        layout.addWidget(self.marche_combo)

        # Montant TTC min
        layout.addWidget(QLabel("Montant min :"))
        self.montant_min_spin = QSpinBox()
        self.montant_min_spin.setRange(0, 10_000_000)
        self.montant_min_spin.setValue(0)
        self.montant_min_spin.setSingleStep(100)
        self.montant_min_spin.setSuffix(" €")
        self.montant_min_spin.setMaximumWidth(110)
        self.montant_min_spin.valueChanged.connect(self._apply_filters)
        layout.addWidget(self.montant_min_spin)

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

        self.btn_batch_ai = QPushButton("🤖 Analyser en lot")
        self.btn_batch_ai.setToolTip(
            "Analyse IA de toutes les commandes RAPPROCHEMENT_SUGGERE non ecartees"
        )
        self.btn_batch_ai.clicked.connect(self._on_batch_ai_clicked)
        layout.addWidget(self.btn_batch_ai)

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

        # Bloc Avis IA (visible mais inactif si pas de cle API configuree)
        self.gb_ai = QGroupBox("Avis IA")
        ai_layout = QVBoxLayout(self.gb_ai)
        ai_header = QHBoxLayout()
        self.ai_diag_label = QLabel("Pas encore consulte.")
        self.ai_diag_label.setStyleSheet("font-weight: bold;")
        self.ai_cache_label = QLabel("")
        self.ai_cache_label.setStyleSheet("color: #888; font-style: italic;")
        ai_header.addWidget(self.ai_diag_label)
        ai_header.addWidget(self.ai_cache_label)
        ai_header.addStretch(1)
        self.btn_ai_apply = QPushButton("Appliquer la suggestion auto")
        self.btn_ai_apply.setEnabled(False)
        self.btn_ai_apply.clicked.connect(self._on_apply_ai_clicked)
        self.btn_ai_refresh = QPushButton("Forcer un nouvel appel IA")
        self.btn_ai_refresh.setEnabled(False)
        self.btn_ai_refresh.clicked.connect(lambda: self._on_ai_clicked(use_cache=False))
        ai_header.addWidget(self.btn_ai_refresh)
        ai_header.addWidget(self.btn_ai_apply)
        ai_layout.addLayout(ai_header)
        self.ai_reasoning_text = QTextEdit()
        self.ai_reasoning_text.setReadOnly(True)
        self.ai_reasoning_text.setMaximumHeight(100)
        self.ai_reasoning_text.setPlaceholderText(
            "Cliquez sur 'Demander avis IA' pour analyser cette commande."
        )
        ai_layout.addWidget(self.ai_reasoning_text)
        v.addWidget(self.gb_ai)

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
        self.btn_ai.clicked.connect(lambda: self._on_ai_clicked(use_cache=True))

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
            self._update_batch_button_state()
            self.status_label.setText("Diagnostic recalcule.")
        finally:
            QApplication.restoreOverrideCursor()

    def _update_batch_button_state(self):
        ai_ready = self._is_ai_configured()
        n_eligible = len(self.engine.list_cmd_ids_for_batch(
            diagnostics=[DIAG_RAPPROCHEMENT_SUGGERE]
        ))
        self.btn_batch_ai.setEnabled(ai_ready and n_eligible > 0)
        if not ai_ready:
            self.btn_batch_ai.setToolTip(
                "Configurer la cle API DeepSeek dans Configuration / Onglet IA"
            )
        elif n_eligible == 0:
            self.btn_batch_ai.setToolTip(
                "Aucune commande RAPPROCHEMENT_SUGGERE en attente d'analyse"
            )
        else:
            self.btn_batch_ai.setToolTip(
                f"Analyse IA de {n_eligible} commande(s) RAPPROCHEMENT_SUGGERE"
            )

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

        # Combo marche
        prev_m = self.marche_combo.currentData()
        self.marche_combo.blockSignals(True)
        self.marche_combo.clear()
        self.marche_combo.addItem("[Tous]", "")
        marches = set()
        for r in self._all_rows:
            m = (r.get("marche") or "").strip()
            if m:
                marches.add(m)
        for m in sorted(marches):
            self.marche_combo.addItem(m, m)
        if prev_m:
            idx = self.marche_combo.findData(prev_m)
            if idx >= 0:
                self.marche_combo.setCurrentIndex(idx)
        self.marche_combo.blockSignals(False)

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
        marche_filter = self.marche_combo.currentData() or ""
        age_min = self.age_min_spin.value()
        montant_min = float(self.montant_min_spin.value())
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
            if marche_filter and (r.get("marche") or "").strip() != marche_filter:
                continue
            age = r.get("age_jours")
            if age is not None and age < age_min:
                continue
            if montant_min > 0:
                ttc = float(r.get("montant_ttc") or 0.0)
                if ttc < montant_min:
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
        for btn in (self.btn_validate, self.btn_doublon, self.btn_no_match,
                    self.btn_ai, self.btn_ai_apply, self.btn_ai_refresh):
            btn.setEnabled(False)
        self._clear_ai_panel()

    def _clear_ai_panel(self):
        self._last_ai_response = None
        self.ai_diag_label.setText("Pas encore consulte.")
        self.ai_diag_label.setStyleSheet("font-weight: bold;")
        self.ai_cache_label.setText("")
        self.ai_reasoning_text.clear()
        self.btn_ai_apply.setEnabled(False)
        self.btn_ai_apply.setText("Appliquer la suggestion auto")

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

        # Bouton IA : enabled si configure et candidats presents
        ai_ready = self._is_ai_configured()
        self.btn_ai.setEnabled(ai_ready and has_candidates)
        if not ai_ready:
            self.btn_ai.setToolTip(
                "Configurer la cle API DeepSeek dans Configuration / Onglet IA"
            )
        elif not has_candidates:
            self.btn_ai.setToolTip("Aucun candidat a analyser")
        else:
            self.btn_ai.setToolTip(
                "Lance une analyse IA des factures candidates pour cette commande"
            )

        # Charger le cache IA si present
        self._load_cached_ai_response(cmd_id)

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
    # IA : configuration, appel asynchrone, application de la decision
    # ------------------------------------------------------------------

    def _is_ai_configured(self) -> bool:
        api_key = self.db.get_config("deepseek_api_key", "") or ""
        enabled = self.db.get_config("ai_enabled", "0") == "1"
        return bool(api_key.strip()) and enabled

    def _build_ai_client(self) -> Optional[DeepSeekClient]:
        api_key = (self.db.get_config("deepseek_api_key", "") or "").strip()
        if not api_key:
            return None
        model = self.db.get_config("deepseek_model", "deepseek-chat") or "deepseek-chat"
        endpoint = (self.db.get_config(
            "deepseek_endpoint",
            "https://api.deepseek.com/v1/chat/completions",
        ) or "https://api.deepseek.com/v1/chat/completions")
        return DeepSeekClient(api_key=api_key, model=model, endpoint=endpoint, timeout=30)

    def _load_cached_ai_response(self, cmd_id: int):
        cur = self.db.conn.cursor()
        cur.execute(
            "SELECT last_ai_diagnostic, last_ai_check_at FROM commande_diagnostic "
            "WHERE commande_id = ?",
            (cmd_id,),
        )
        row = cur.fetchone()
        if not row or not row["last_ai_diagnostic"]:
            return
        try:
            import json as _json
            data = _json.loads(row["last_ai_diagnostic"])
            data["_from_cache"] = True
        except Exception:
            return
        self._display_ai_response(data, cached_at=row["last_ai_check_at"])

    def _on_ai_clicked(self, use_cache: bool = True):
        if self.current_cmd_id is None:
            return
        if self._ai_worker is not None and self._ai_worker.isRunning():
            QMessageBox.information(
                self, "IA en cours",
                "Un appel IA est deja en cours, attendez sa fin."
            )
            return
        client = self._build_ai_client()
        if client is None:
            QMessageBox.warning(
                self, "IA non configuree",
                "Configurer la cle API DeepSeek dans Configuration / Onglet IA."
            )
            return
        candidates = self.engine.find_candidates(self.current_cmd_id)
        if not candidates:
            QMessageBox.information(
                self, "Aucun candidat",
                "Pas de facture candidate a analyser."
            )
            return

        # UI : feedback debut
        self._set_ai_busy(True)
        self.status_label.setText("Appel IA en cours...")
        self.ai_diag_label.setText("Analyse en cours...")
        self.ai_diag_label.setStyleSheet("font-weight: bold; color: #555;")
        self.ai_reasoning_text.setPlainText("")

        worker = AiWorker(self.engine, self.current_cmd_id,
                          candidates, client, use_cache)
        worker.finished_with_response.connect(self._on_ai_response)
        worker.finished_with_error.connect(self._on_ai_error)
        # Auto-cleanup
        worker.finished.connect(lambda: setattr(self, "_ai_worker", None))
        self._ai_worker = worker
        worker.start()

    def _on_ai_response(self, response: dict):
        self._set_ai_busy(False)
        self._display_ai_response(response)
        cached = response.get("_from_cache", False)
        self.status_label.setText(
            "Reponse IA recuperee depuis le cache."
            if cached else "Reponse IA recue."
        )

    def _on_ai_error(self, message: str):
        self._set_ai_busy(False)
        self.ai_diag_label.setText("Erreur IA")
        self.ai_diag_label.setStyleSheet("font-weight: bold; color: #b00020;")
        self.ai_reasoning_text.setPlainText(message)
        self.status_label.setText("Erreur lors de l'appel IA.")
        QMessageBox.critical(self, "Erreur IA",
                             f"L'appel a l'API DeepSeek a echoue :\n\n{message}")

    def _set_ai_busy(self, busy: bool):
        self.btn_ai.setEnabled(not busy)
        self.btn_ai_refresh.setEnabled(not busy and self._last_ai_response is not None)
        self.btn_ai_apply.setEnabled(not busy and self._can_apply_ai_decision())

    def _can_apply_ai_decision(self) -> bool:
        if not self._last_ai_response:
            return False
        try:
            auto_thr = int(self.db.get_config("ai_auto_threshold", "90") or 90)
        except (TypeError, ValueError):
            auto_thr = 90
        confidence = int(self._last_ai_response.get("confidence", 0))
        action = self._last_ai_response.get("action_suggeree")
        diag = self._last_ai_response.get("diagnostic")
        if confidence < auto_thr:
            return False
        if diag == "DOUBLON" and action == "MARQUER_DOUBLON":
            return True
        if diag in ("MATCH", "PARTIEL") and action == "VALIDER_AUTO":
            return True
        return False

    def _display_ai_response(self, response: dict, cached_at: Optional[str] = None):
        self._last_ai_response = response
        diag = response.get("diagnostic", "?")
        confidence = response.get("confidence", 0)
        action = response.get("action_suggeree", "?")
        reasoning = response.get("raisonnement", "")

        # Colorer selon diagnostic + confiance
        color = "#555"
        if diag == "DOUBLON":
            color = "#b25400"
        elif diag == "MATCH":
            color = "#1b7e2c"
        elif diag == "PARTIEL":
            color = "#996800"
        elif diag == "ORPHELINE":
            color = "#b00020"
        self.ai_diag_label.setStyleSheet(f"font-weight: bold; color: {color};")
        self.ai_diag_label.setText(
            f"{diag} — confiance {confidence}% — action suggeree : {action}"
        )
        if response.get("_from_cache") or cached_at:
            stamp = cached_at or "cache"
            self.ai_cache_label.setText(f"(cache : {stamp[:19]})")
        else:
            self.ai_cache_label.setText("")

        self.ai_reasoning_text.setPlainText(reasoning)

        # Pre-cocher les factures suggerees + remplir leurs montants
        suggested = response.get("factures_a_lier") or []
        self._apply_ai_suggestions_to_candidates(suggested)

        # Bouton applique uniquement si confidence >= auto_threshold + action coherente
        self.btn_ai_apply.setEnabled(self._can_apply_ai_decision())
        self.btn_ai_apply.setText(self._apply_button_label(diag, action))
        self.btn_ai_refresh.setEnabled(True)

    def _apply_button_label(self, diag: str, action: str) -> str:
        if action == "MARQUER_DOUBLON":
            return "Marquer doublon (auto)"
        if action == "VALIDER_AUTO":
            return "Creer les liens (auto)"
        return "Appliquer la suggestion auto"

    def _apply_ai_suggestions_to_candidates(self, suggested: list):
        if not suggested:
            return
        # Map code_mouvement -> (montant_alloue suggere)
        by_code = {}
        for s in suggested:
            code = (s.get("code_mouvement") or "").strip()
            if not code:
                continue
            try:
                m = float(s.get("montant_alloue") or 0.0)
            except (TypeError, ValueError):
                continue
            by_code[code] = m

        for i in range(self.cand_table.rowCount()):
            code_item = self.cand_table.item(i, self.CC_CODE_MVT)
            if code_item is None:
                continue
            code = code_item.text().strip()
            if code in by_code:
                chk = self.cand_table.item(i, self.CC_CHECK)
                if chk is not None:
                    chk.setCheckState(Qt.Checked)
                spin = self.cand_table.cellWidget(i, self.CC_ALLOUE)
                if spin is not None and by_code[code] > 0:
                    # Cap au max du spinbox
                    spin.setValue(min(by_code[code], spin.maximum()))

    def _on_apply_ai_clicked(self):
        if self.current_cmd_id is None or not self._last_ai_response:
            return
        try:
            auto_thr = int(self.db.get_config("ai_auto_threshold", "90") or 90)
            min_thr = int(self.db.get_config("ai_min_threshold", "40") or 40)
        except (TypeError, ValueError):
            auto_thr, min_thr = 90, 40

        ans = QMessageBox.question(
            self, "Confirmer",
            f"Appliquer la decision IA ?\n\n"
            f"Diagnostic : {self._last_ai_response.get('diagnostic')}\n"
            f"Confiance : {self._last_ai_response.get('confidence')}%\n"
            f"Action : {self._last_ai_response.get('action_suggeree')}",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if ans != QMessageBox.Yes:
            return

        try:
            result = apply_ai_decision(
                self.db.conn, self.link_repo,
                self.current_cmd_id, self._last_ai_response,
                auto_threshold=auto_thr, min_threshold=min_thr,
            )
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Echec : {e}")
            return

        msg = {
            ACTION_AUTO_DOUBLON: "Commande marquee comme doublon administratif (auto IA).",
            ACTION_AUTO_LINKED: "Liens IA crees automatiquement.",
            ACTION_SUGGESTION: "Confiance insuffisante pour auto-validation, suggestion enregistree.",
            ACTION_IGNORED: "Decision IA ignoree (sous le seuil minimum).",
        }.get(result, f"Resultat : {result}")
        self.status_label.setText(msg)
        # Recompute global + refresh
        self._recompute_after_action()

    # ------------------------------------------------------------------
    # IA en lot (etape 5)
    # ------------------------------------------------------------------

    def _on_batch_ai_clicked(self):
        if self._ai_worker is not None and self._ai_worker.isRunning():
            QMessageBox.information(
                self, "IA en cours",
                "Un appel IA est deja en cours, attendez sa fin."
            )
            return
        client = self._build_ai_client()
        if client is None:
            QMessageBox.warning(
                self, "IA non configuree",
                "Configurer la cle API DeepSeek dans Configuration / Onglet IA."
            )
            return
        cmd_ids = self.engine.list_cmd_ids_for_batch(
            diagnostics=[DIAG_RAPPROCHEMENT_SUGGERE]
        )
        if not cmd_ids:
            QMessageBox.information(
                self, "Rien a analyser",
                "Aucune commande RAPPROCHEMENT_SUGGERE non ecartee."
            )
            return

        # Estimation cout et duree (~0.1ct/cmd, ~3s/cmd selon spec 7.6)
        n = len(cmd_ids)
        cost_cents = max(1, round(n * 0.1, 1))
        eta_s = n * 3
        ans = QMessageBox.question(
            self, "Analyser en lot",
            f"Lancer l'analyse IA pour {n} commande(s) RAPPROCHEMENT_SUGGERE ?\n\n"
            f"- Cout estime : ~{cost_cents} centimes\n"
            f"- Duree estimee : ~{eta_s}s ({n} appels en serie)\n"
            f"- Les decisions a confiance >= seuil auto seront appliquees automatiquement.\n"
            f"- Les commandes deja analysees recemment utiliseront le cache (gratuit).\n\n"
            "Continuer ?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if ans != QMessageBox.Yes:
            return

        # Seuils
        try:
            auto_thr = int(self.db.get_config("ai_auto_threshold", "90") or 90)
            min_thr = int(self.db.get_config("ai_min_threshold", "40") or 40)
        except (TypeError, ValueError):
            auto_thr, min_thr = 90, 40

        # Progress dialog avec cancel
        progress = QProgressDialog(
            "Analyse IA en lot...", "Annuler", 0, n, self
        )
        progress.setWindowTitle("Batch IA")
        progress.setWindowModality(Qt.WindowModal)
        progress.setAutoClose(True)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        cancel_state = {"cancel": False}
        progress.canceled.connect(lambda: cancel_state.update({"cancel": True}))

        worker = BatchAiWorker(
            self.engine, cmd_ids, client,
            auto_threshold=auto_thr, min_threshold=min_thr,
            cancel_check=lambda: cancel_state["cancel"],
        )

        def _on_progress(i, total, msg):
            progress.setMaximum(total)
            progress.setValue(i)
            progress.setLabelText(f"({i}/{total}) {msg}")

        worker.progress.connect(_on_progress)
        worker.finished_with_summary.connect(
            lambda summary: self._on_batch_finished(summary, progress)
        )
        worker.finished_with_error.connect(
            lambda msg: self._on_batch_error(msg, progress)
        )
        worker.finished.connect(lambda: setattr(self, "_ai_worker", None))

        self._ai_worker = worker
        self.btn_batch_ai.setEnabled(False)
        self.status_label.setText(f"Batch IA en cours sur {n} commande(s)...")
        worker.start()

    def _on_batch_finished(self, summary: dict, progress: QProgressDialog):
        progress.setValue(progress.maximum())
        # Recompute global + refresh
        self._recompute_after_action()
        self._update_batch_button_state()

        cancelled = " (interrompu)" if summary.get("cancelled") else ""
        details = (
            f"<b>Batch IA termine{cancelled}</b><br><br>"
            f"Total : {summary['total']}<br>"
            f"Traitees : {summary['processed']}<br>"
            f"&nbsp;&nbsp;Appels API : {summary['ai_called']}<br>"
            f"&nbsp;&nbsp;Cache : {summary['from_cache']}<br>"
            f"<br>"
            f"Resultats :<br>"
            f"&nbsp;&nbsp;Auto-doublons : {summary['auto_doublon']}<br>"
            f"&nbsp;&nbsp;Auto-liens : {summary['auto_linked']}<br>"
            f"&nbsp;&nbsp;Suggestions (sous seuil auto) : {summary['suggestions']}<br>"
            f"&nbsp;&nbsp;Ignorees : {summary['ignored']}<br>"
            f"&nbsp;&nbsp;Sans candidat : {summary['skipped_no_candidate']}<br>"
        )
        if summary.get("errors"):
            n_err = len(summary["errors"])
            details += f"<br><b>{n_err} erreur(s)</b><br>"
            for cmd_id, msg in summary["errors"][:5]:
                details += f"&nbsp;&nbsp;cmd #{cmd_id} : {msg[:100]}<br>"
            if n_err > 5:
                details += f"&nbsp;&nbsp;... et {n_err - 5} autre(s)<br>"

        self.status_label.setText(
            f"Batch IA : {summary['auto_doublon']} doublons, "
            f"{summary['auto_linked']} liens, "
            f"{summary['suggestions']} suggestions."
        )
        QMessageBox.information(self, "Batch IA termine", details)

    def _on_batch_error(self, msg: str, progress: QProgressDialog):
        progress.cancel()
        self._update_batch_button_state()
        self.status_label.setText("Erreur batch IA.")
        QMessageBox.critical(self, "Erreur batch IA",
                             f"L'analyse en lot a echoue :\n\n{msg}")

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
        # Stopper proprement un worker IA en cours
        if self._ai_worker is not None and self._ai_worker.isRunning():
            self._ai_worker.wait(2000)
        try:
            settings = QSettings("suivi_sedit", "matching_dialog")
            settings.setValue("geometry", self.saveGeometry())
        except Exception:
            pass
        super().closeEvent(event)


# ============================================================================
# AiWorker : appel asynchrone DeepSeek dans un QThread (UI ne bloque pas)
# ============================================================================

class AiWorker(QThread):
    """Encapsule un appel a MatchingEngine.ai_match dans un thread dedie.

    Emet finished_with_response(dict) en cas de succes, ou
    finished_with_error(str) en cas d'erreur.
    """
    finished_with_response = pyqtSignal(dict)
    finished_with_error = pyqtSignal(str)

    def __init__(self, engine: MatchingEngine, commande_id: int,
                 candidates: list, client: DeepSeekClient,
                 use_cache: bool = True, parent=None):
        super().__init__(parent)
        self.engine = engine
        self.commande_id = commande_id
        self.candidates = candidates
        self.client = client
        self.use_cache = use_cache

    def run(self):
        try:
            response = self.engine.ai_match(
                self.commande_id, self.candidates, self.client,
                use_cache=self.use_cache,
            )
            self.finished_with_response.emit(response)
        except Exception as e:
            self.finished_with_error.emit(f"{type(e).__name__}: {e}")


# ============================================================================
# BatchAiWorker : analyse en lot de N commandes (etape 5)
# ============================================================================

class BatchAiWorker(QThread):
    """Encapsule batch_ai_match dans un thread pour ne pas bloquer l'UI.

    Emet :
    - progress(i, total, message) apres chaque commande
    - finished_with_summary(summary_dict) en fin de batch
    - finished_with_error(message) en cas d'erreur fatale (rare,
      les erreurs par-commande sont collectees dans summary['errors'])
    """
    progress = pyqtSignal(int, int, str)
    finished_with_summary = pyqtSignal(dict)
    finished_with_error = pyqtSignal(str)

    def __init__(self, engine: MatchingEngine, cmd_ids: list,
                 client: DeepSeekClient,
                 auto_threshold: int = 90,
                 min_threshold: int = 40,
                 apply_decisions: bool = True,
                 use_cache: bool = True,
                 cancel_check=None,
                 parent=None):
        super().__init__(parent)
        self.engine = engine
        self.cmd_ids = cmd_ids
        self.client = client
        self.auto_threshold = auto_threshold
        self.min_threshold = min_threshold
        self.apply_decisions = apply_decisions
        self.use_cache = use_cache
        self._external_cancel = cancel_check

    def run(self):
        try:
            summary = self.engine.batch_ai_match(
                self.cmd_ids, self.client,
                auto_threshold=self.auto_threshold,
                min_threshold=self.min_threshold,
                apply_decisions=self.apply_decisions,
                use_cache=self.use_cache,
                progress_callback=lambda i, n, msg: self.progress.emit(i, n, msg),
                cancel_check=self._external_cancel,
            )
            self.finished_with_summary.emit(summary)
        except Exception as e:
            self.finished_with_error.emit(f"{type(e).__name__}: {e}")
