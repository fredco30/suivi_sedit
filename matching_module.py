"""Module de rapprochement commandes <-> factures.

Couche metier pure : aucune dependance PyQt, testable en isolation.
Utilise sqlite3 + stdlib uniquement.

Etape 2 : pre-classification deterministe (sans IA).
- LinkRepository : CRUD sur la table manual_links.
- MatchingEngine.find_candidates : pre-filtrage des factures candidates pour une commande.
- MatchingEngine.diagnose_all_commandes : remplit la table commande_diagnostic.

Les classes DeepSeekClient et MatchingEngine.ai_match sont reservees a l'etape 4.
"""
from __future__ import annotations

import re
import sqlite3
from datetime import datetime, date, timedelta
from difflib import SequenceMatcher
from typing import Optional


# ============================================================================
# Constantes : enumeration des diagnostics et seuils
# ============================================================================

DIAG_OK = "OK"
DIAG_OK_RAPPROCHE = "OK_RAPPROCHE"
DIAG_RECENT = "RECENT"
DIAG_EN_COURS = "EN_COURS"
DIAG_RAPPROCHEMENT_SUGGERE = "RAPPROCHEMENT_SUGGERE"
DIAG_OUBLI_PROBABLE = "OUBLI_PROBABLE"
DIAG_DOUBLON_PROBABLE = "DOUBLON_PROBABLE"

SEVERITE_BY_DIAG = {
    DIAG_OK: 0,
    DIAG_OK_RAPPROCHE: 0,
    DIAG_RECENT: 1,
    DIAG_EN_COURS: 1,
    DIAG_RAPPROCHEMENT_SUGGERE: 2,
    DIAG_DOUBLON_PROBABLE: 2,
    DIAG_OUBLI_PROBABLE: 3,
}

THRESHOLD_FOURNISSEUR = 0.85
AGE_RECENT_DAYS = 30
AGE_EN_COURS_DAYS = 90
DATE_TOLERANCE_DAYS = 30


# ============================================================================
# Helpers (fonctions pures)
# ============================================================================

_PUNCT_RE = re.compile(r"[^A-Z0-9 ]+")
_SPACES_RE = re.compile(r"\s+")
_NUM_ENGAGEMENT_RE = re.compile(r"^(\d{2})([A-Z]+)(\d+)$")


def normalize_fournisseur(s: Optional[str]) -> str:
    """Normalise un nom de fournisseur pour comparaison fuzzy.

    Upper-case, retire les accents grossierement (via remplacement ASCII-friendly),
    retire la ponctuation et les espaces multiples.
    """
    if not s:
        return ""
    out = s.upper().strip()
    # Remplacements simples des accents les plus courants
    for src, dst in (("É", "E"), ("È", "E"), ("Ê", "E"), ("Ë", "E"),
                     ("À", "A"), ("Â", "A"), ("Ä", "A"),
                     ("Î", "I"), ("Ï", "I"),
                     ("Ô", "O"), ("Ö", "O"),
                     ("Ù", "U"), ("Û", "U"), ("Ü", "U"),
                     ("Ç", "C")):
        out = out.replace(src, dst)
    out = _PUNCT_RE.sub(" ", out)
    out = _SPACES_RE.sub(" ", out).strip()
    return out


def fuzzy_ratio(a: Optional[str], b: Optional[str]) -> float:
    """Ratio de similarite entre deux noms de fournisseur, apres normalisation.
    Retourne 0.0 si l'un des deux est vide.
    """
    na, nb = normalize_fournisseur(a), normalize_fournisseur(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def parse_num_engagement(code: Optional[str]) -> Optional[tuple]:
    """Parse un identifiant type '25AA01486' -> (2025, 'AA', 1486).

    Renvoie None si le format ne correspond pas (cas REPORT, vide, etc.).
    L'exercice 2-chiffres est interprete comme 2000+yy (les donnees couvrent 2020-2099).
    """
    if not code:
        return None
    m = _NUM_ENGAGEMENT_RE.match(code.strip().upper())
    if not m:
        return None
    yy, kind, seq = m.groups()
    return (2000 + int(yy), kind, int(seq))


def parse_date(s: Optional[str]) -> Optional[date]:
    """Parse une date au format ISO (YYYY-MM-DD[Thh:mm:ss]) ou DD/MM/YYYY.
    Renvoie None si parsing impossible.
    """
    if not s:
        return None
    s = s.strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # Fallback : prendre les 10 premiers caracteres comme YYYY-MM-DD
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


# ============================================================================
# LinkRepository : CRUD sur manual_links
# ============================================================================

class LinkRepository:
    """Couche d'acces a la table manual_links.

    Ne declenche pas recompute_facturation : c'est l'appelant (UI ou test
    d'integration) qui doit le faire apres add/remove.
    """

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn
        # On exige le row_factory : tous les fetchall renvoient des dict-like.
        self.conn.row_factory = sqlite3.Row

    # --------- Lecture ---------

    def list_links_for_commande(self, commande_id: int) -> list:
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM manual_links WHERE commande_id = ? ORDER BY created_at",
            (commande_id,),
        )
        return [dict(r) for r in cur.fetchall()]

    def list_links_for_facture(self, facture_id: int) -> list:
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM manual_links WHERE facture_id = ? ORDER BY created_at",
            (facture_id,),
        )
        return [dict(r) for r in cur.fetchall()]

    def total_alloue_for_commande(self, commande_id: int) -> float:
        cur = self.conn.cursor()
        cur.execute(
            "SELECT COALESCE(SUM(montant_alloue), 0) AS total FROM manual_links WHERE commande_id = ?",
            (commande_id,),
        )
        return float(cur.fetchone()["total"] or 0.0)

    def total_alloue_for_facture(self, facture_id: int) -> float:
        cur = self.conn.cursor()
        cur.execute(
            "SELECT COALESCE(SUM(montant_alloue), 0) AS total FROM manual_links WHERE facture_id = ?",
            (facture_id,),
        )
        return float(cur.fetchone()["total"] or 0.0)

    # --------- Ecriture ---------

    def add_link(
        self,
        commande_id: int,
        facture_id: int,
        montant_alloue: float,
        source: str,
        confidence: Optional[int] = None,
        ai_model: Optional[str] = None,
        ai_reasoning: Optional[str] = None,
        created_by: str = "user",
        notes: Optional[str] = None,
        validated_at: Optional[str] = None,
    ) -> int:
        """Insere un lien manuel commande <-> facture.

        Verifications applicatives avant insert :
        - montant_alloue > 0
        - montant_alloue + somme deja allouee sur cette facture <= montant_service_fait
        Sinon raise ValueError. Le doublon (commande_id, facture_id) sera detecte
        par la contrainte UNIQUE et remontera comme sqlite3.IntegrityError.
        """
        if montant_alloue <= 0:
            raise ValueError(f"montant_alloue doit etre > 0 (recu {montant_alloue})")

        cur = self.conn.cursor()
        cur.execute(
            "SELECT montant_service_fait FROM factures WHERE id = ?",
            (facture_id,),
        )
        row = cur.fetchone()
        if row is None:
            raise ValueError(f"facture_id {facture_id} introuvable")
        montant_sf = float(row["montant_service_fait"] or 0.0)

        deja_alloue = self.total_alloue_for_facture(facture_id)
        if deja_alloue + montant_alloue > montant_sf + 1e-6:
            raise ValueError(
                f"Allocation cumulee {deja_alloue + montant_alloue:.2f} EUR depasserait "
                f"montant_service_fait {montant_sf:.2f} EUR sur facture id={facture_id}"
            )

        # Verifier que la commande existe
        cur.execute("SELECT 1 FROM commandes WHERE id = ?", (commande_id,))
        if cur.fetchone() is None:
            raise ValueError(f"commande_id {commande_id} introuvable")

        cur.execute(
            """
            INSERT INTO manual_links
            (commande_id, facture_id, montant_alloue, confidence, source,
             ai_model, ai_reasoning, created_by, created_at, validated_at, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                commande_id, facture_id, float(montant_alloue),
                confidence, source, ai_model, ai_reasoning,
                created_by, datetime.now().isoformat(timespec="seconds"),
                validated_at, notes,
            ),
        )
        self.conn.commit()
        return cur.lastrowid

    def remove_link(self, link_id: int) -> bool:
        cur = self.conn.cursor()
        cur.execute("DELETE FROM manual_links WHERE id = ?", (link_id,))
        self.conn.commit()
        return cur.rowcount > 0


# ============================================================================
# MatchingEngine : pre-classification deterministe
# ============================================================================

class MatchingEngine:
    """Pre-classifie les commandes et trouve les factures candidates.

    Parametres :
    - conn : connexion sqlite3 ouverte sur la base.
    - link_repo : instance de LinkRepository (peut etre cree par l'appelant).
    - today : date de reference (par defaut, datetime.now().date()) ; surchargeable
      pour les tests deterministes.
    """

    def __init__(self, conn: sqlite3.Connection, link_repo: LinkRepository,
                 today: Optional[date] = None):
        self.conn = conn
        self.conn.row_factory = sqlite3.Row
        self.link_repo = link_repo
        self.today = today or datetime.now().date()

    # ------------------------------------------------------------------
    # find_candidates
    # ------------------------------------------------------------------

    def find_candidates(self, commande_id: int) -> list:
        """Retourne la liste des factures candidates pour une commande donnee.

        Regles (cf. specification 6.2) :
        1. montant_service_fait > 0
        2. La facture n'est pas deja rattachee nativement a une commande
           (code_mouvement n'est le num_commande d'aucune commande connue).
        3. La facture n'est pas deja liee manuellement a la commande cible.
        4. fuzzy_ratio(facture.fournisseur, commande.fournisseur) >= 0.85
        5. facture.date_facture >= commande.date_commande - 30j
        6. Si commande.marche et facture.marche tous deux non vides : strict equal
        7. Numerotation : si meme exercice (parsable), num_seq(code_mouvement) <=
           num_seq(num_commande). Exercice anterieur : pas de contrainte. Format
           inconnu : pas de contrainte (souple).
        """
        cur = self.conn.cursor()
        cur.execute(
            "SELECT id, num_commande, fournisseur, marche, date_commande, montant_ttc "
            "FROM commandes WHERE id = ?",
            (commande_id,),
        )
        cmd = cur.fetchone()
        if cmd is None:
            return []
        cmd_dict = dict(cmd)

        date_cmd = parse_date(cmd_dict.get("date_commande"))
        seuil_date = (date_cmd - timedelta(days=DATE_TOLERANCE_DAYS)) if date_cmd else None
        cmd_engagement = parse_num_engagement(cmd_dict.get("num_commande"))
        cmd_marche = (cmd_dict.get("marche") or "").strip() or None
        cmd_fourn_norm = normalize_fournisseur(cmd_dict.get("fournisseur"))
        if not cmd_fourn_norm:
            return []

        # Pre-filtrage SQL grossier sur fournisseur (LIKE) pour ne pas iterer
        # sur des milliers de factures sans rapport. On utilise le premier token
        # significatif (>= 3 caracteres) du fournisseur normalise.
        first_token = next((t for t in cmd_fourn_norm.split() if len(t) >= 3), cmd_fourn_norm)
        like_pat = f"%{first_token}%"

        # Pre-charger les num_commande connues pour exclure les factures rattachees
        # nativement (regle 2).
        cur.execute("SELECT DISTINCT num_commande FROM commandes WHERE num_commande IS NOT NULL AND num_commande != ''")
        nums_connus = {r["num_commande"] for r in cur.fetchall()}

        # Pre-charger les facture_id deja liees a cette commande (regle 3).
        already_linked = {
            link["facture_id"] for link in self.link_repo.list_links_for_commande(commande_id)
        }

        cur.execute(
            """
            SELECT id, num_facture, code_mouvement, fournisseur, marche,
                   date_facture, montant_service_fait, libelle
            FROM factures
            WHERE montant_service_fait IS NOT NULL
              AND montant_service_fait > 0
              AND UPPER(fournisseur) LIKE ?
            """,
            (like_pat,),
        )
        rows = [dict(r) for r in cur.fetchall()]

        candidates = []
        for f in rows:
            if f["id"] in already_linked:
                continue

            code_mvt = (f.get("code_mouvement") or "").strip()
            # Regle 2 : facture deja rattachee nativement a une commande connue
            if code_mvt and code_mvt in nums_connus:
                continue

            # Regle 4 : fuzzy fournisseur
            if fuzzy_ratio(cmd_dict.get("fournisseur"), f.get("fournisseur")) < THRESHOLD_FOURNISSEUR:
                continue

            # Regle 5 : date facture >= date commande - 30j
            if seuil_date is not None:
                date_fact = parse_date(f.get("date_facture"))
                if date_fact is not None and date_fact < seuil_date:
                    continue

            # Regle 6 : marche strict si tous deux renseignes
            f_marche = (f.get("marche") or "").strip() or None
            if cmd_marche and f_marche and cmd_marche != f_marche:
                continue

            # Regle 7 : numerotation chronologique si meme exercice
            f_engagement = parse_num_engagement(code_mvt)
            if cmd_engagement and f_engagement:
                if f_engagement[0] == cmd_engagement[0]:  # meme exercice
                    if f_engagement[2] > cmd_engagement[2]:
                        continue
                # exercices differents : pas de contrainte (cas REPORT)

            candidates.append(f)

        return candidates

    # ------------------------------------------------------------------
    # diagnose_all_commandes
    # ------------------------------------------------------------------

    def diagnose_all_commandes(self) -> dict:
        """Parcourt toutes les commandes et remplit la table commande_diagnostic.

        Renvoie un dict {diagnostic: count} pour log/UI.
        Ne touche pas aux colonnes last_ai_check_at / last_ai_diagnostic (cache IA, etape 4).
        Reutilise UPSERT pour preserver le cache IA si commande_diagnostic deja peuplee.
        """
        cur = self.conn.cursor()
        cur.execute(
            "SELECT id, num_commande, date_commande, statut_facturation, statut_metier "
            "FROM commandes"
        )
        commandes = [dict(r) for r in cur.fetchall()]

        counts = {d: 0 for d in SEVERITE_BY_DIAG}
        now_iso = datetime.now().isoformat(timespec="seconds")

        for cmd in commandes:
            cmd_id = cmd["id"]
            statut_metier = cmd.get("statut_metier")
            statut_fact = cmd.get("statut_facturation")

            date_cmd = parse_date(cmd.get("date_commande"))
            age_jours = (self.today - date_cmd).days if date_cmd else None

            candidates_count = 0
            candidates_same_marche = 0
            montant_candidates_total = 0.0

            # Determination du diagnostic
            if statut_metier == "DOUBLON_ADMIN":
                diag = DIAG_DOUBLON_PROBABLE
            elif statut_fact == "Totalement facturée":
                has_manual = bool(self.link_repo.list_links_for_commande(cmd_id))
                diag = DIAG_OK_RAPPROCHE if has_manual else DIAG_OK
            elif age_jours is None:
                # Pas de date : on traite comme RECENT par defaut (severite 1)
                diag = DIAG_RECENT
            elif age_jours < AGE_RECENT_DAYS:
                diag = DIAG_RECENT
            elif age_jours < AGE_EN_COURS_DAYS:
                diag = DIAG_EN_COURS
            else:
                # > 90j et non totalement facturee : on cherche des candidats
                cands = self.find_candidates(cmd_id)
                candidates_count = len(cands)
                cmd_marche = (cur.execute("SELECT marche FROM commandes WHERE id = ?", (cmd_id,))
                              .fetchone()["marche"] or "").strip()
                if cmd_marche:
                    candidates_same_marche = sum(
                        1 for c in cands if (c.get("marche") or "").strip() == cmd_marche
                    )
                montant_candidates_total = sum(
                    float(c.get("montant_service_fait") or 0.0) for c in cands
                )
                diag = DIAG_RAPPROCHEMENT_SUGGERE if candidates_count > 0 else DIAG_OUBLI_PROBABLE

            severite = SEVERITE_BY_DIAG[diag]
            counts[diag] = counts.get(diag, 0) + 1

            # Upsert : on preserve last_ai_* si la ligne existe deja.
            cur.execute("SELECT 1 FROM commande_diagnostic WHERE commande_id = ?", (cmd_id,))
            existing = cur.fetchone()
            if existing:
                cur.execute(
                    """
                    UPDATE commande_diagnostic
                    SET age_jours = ?, diagnostic = ?, severite = ?,
                        candidates_count = ?, candidates_same_marche = ?,
                        montant_candidates_total = ?, last_diagnostic_at = ?
                    WHERE commande_id = ?
                    """,
                    (age_jours, diag, severite, candidates_count, candidates_same_marche,
                     montant_candidates_total, now_iso, cmd_id),
                )
            else:
                cur.execute(
                    """
                    INSERT INTO commande_diagnostic
                    (commande_id, age_jours, diagnostic, severite,
                     candidates_count, candidates_same_marche, montant_candidates_total,
                     last_diagnostic_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (cmd_id, age_jours, diag, severite, candidates_count,
                     candidates_same_marche, montant_candidates_total, now_iso),
                )

        self.conn.commit()
        return counts
