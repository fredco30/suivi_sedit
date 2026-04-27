"""Module de rapprochement commandes <-> factures.

Couche metier pure : aucune dependance PyQt, testable en isolation.
Utilise sqlite3 + stdlib uniquement (urllib pour DeepSeek, pas de requests).

Etape 2 : pre-classification deterministe (sans IA).
- LinkRepository : CRUD sur la table manual_links.
- MatchingEngine.find_candidates : pre-filtrage des factures candidates pour une commande.
- MatchingEngine.diagnose_all_commandes : remplit la table commande_diagnostic.

Etape 4 : assistance IA via DeepSeek.
- DeepSeekClient : POST sync avec retry, parse_json_response tolerant.
- MatchingEngine.ai_match : appel IA avec cache (commande_diagnostic.last_ai_*).
- apply_ai_decision : applique la decision IA selon les seuils (auto / suggestion).
"""
from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import time
import urllib.error
import urllib.request
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

    # ------------------------------------------------------------------
    # ai_match (etape 4)
    # ------------------------------------------------------------------

    def ai_match(self, commande_id: int, candidates: list,
                 client: "DeepSeekClient",
                 use_cache: bool = True,
                 max_age_days: int = 7) -> dict:
        """Interroge l'IA pour analyser une commande et ses candidats.

        Cache : si commande_diagnostic.last_ai_diagnostic existe avec un
        last_ai_check_at < max_age_days ET le hash candidats identique au
        last_ai_candidates_hash, retourne le cache sans rappeler l'API.

        Renvoie le dict parse de la reponse IA (cf. spec 7.3 pour la structure).
        Le dict contient en plus une cle "_from_cache" : bool, et "_raw" : str.
        """
        cur = self.conn.cursor()

        cur.execute(
            "SELECT id, num_commande, fournisseur, marche, service_emetteur, "
            "montant_ttc, reste_a_facturer, libelle, date_commande "
            "FROM commandes WHERE id = ?",
            (commande_id,),
        )
        cmd_row = cur.fetchone()
        if cmd_row is None:
            raise ValueError(f"commande_id {commande_id} introuvable")
        cmd = dict(cmd_row)

        cand_hash = compute_candidates_hash(candidates)

        if use_cache:
            cur.execute(
                "SELECT last_ai_check_at, last_ai_diagnostic, last_ai_candidates_hash "
                "FROM commande_diagnostic WHERE commande_id = ?",
                (commande_id,),
            )
            existing = cur.fetchone()
            if existing and existing["last_ai_diagnostic"] and existing["last_ai_check_at"]:
                if existing["last_ai_candidates_hash"] == cand_hash:
                    cached_iso = existing["last_ai_check_at"]
                    cached_dt = parse_date(cached_iso)
                    if cached_dt is not None and (self.today - cached_dt).days <= max_age_days:
                        try:
                            data = json.loads(existing["last_ai_diagnostic"])
                            data["_from_cache"] = True
                            data["_raw"] = existing["last_ai_diagnostic"]
                            return data
                        except json.JSONDecodeError:
                            # Cache corrompu : on ignore et on rappelle l'API
                            pass

        user_prompt = build_ai_user_prompt(cmd, candidates)
        raw = client.chat(
            system_prompt=AI_SYSTEM_PROMPT,
            user_prompt=user_prompt,
            max_tokens=900,
            temperature=0.1,
        )
        parsed = DeepSeekClient.parse_json_response(raw)
        validate_ai_response(parsed)

        # Persist cache (UPSERT : on prend soin de ne pas ecraser les autres champs)
        now_iso = datetime.now().isoformat(timespec="seconds")
        payload = json.dumps(parsed, ensure_ascii=False)
        cur.execute(
            "SELECT 1 FROM commande_diagnostic WHERE commande_id = ?", (commande_id,)
        )
        if cur.fetchone():
            cur.execute(
                "UPDATE commande_diagnostic SET last_ai_check_at = ?, "
                "last_ai_diagnostic = ?, last_ai_candidates_hash = ? "
                "WHERE commande_id = ?",
                (now_iso, payload, cand_hash, commande_id),
            )
        else:
            cur.execute(
                """
                INSERT INTO commande_diagnostic
                (commande_id, diagnostic, severite, last_diagnostic_at,
                 last_ai_check_at, last_ai_diagnostic, last_ai_candidates_hash)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (commande_id, DIAG_RAPPROCHEMENT_SUGGERE,
                 SEVERITE_BY_DIAG[DIAG_RAPPROCHEMENT_SUGGERE],
                 now_iso, now_iso, payload, cand_hash),
            )
        self.conn.commit()

        parsed["_from_cache"] = False
        parsed["_raw"] = raw
        return parsed


# ============================================================================
# DeepSeekClient + helpers IA (etape 4)
# ============================================================================

AI_SYSTEM_PROMPT = (
    "Tu es un expert en comptabilite publique francaise (M14/M57). "
    "Tu analyses si une commande peut etre rapprochee de factures candidates. "
    "Tu reponds TOUJOURS en JSON strict, sans preambule ni texte hors JSON."
)

AI_USER_PROMPT_TEMPLATE = """COMMANDE A RAPPROCHER:
- N°: {num_commande}
- Date: {date_commande}
- Fournisseur: {fournisseur}
- Marche: {marche}
- Service emetteur: {service_emetteur}
- Montant TTC: {montant_ttc} EUR
- Reste a facturer: {reste} EUR
- Libelle: {libelle}

FACTURES CANDIDATES (factures dont le code_mouvement n'est pas le n° de cette commande,
mais qui pourraient correspondre par fournisseur/marche/semantique):

{factures_json}

CONSIGNES:
1. Evalue si UNE OU PLUSIEURS factures peuvent ensemble couvrir la commande
2. Mefie-toi des DOUBLONS ADMINISTRATIFS : si la commande ressemble a un engagement deja
   couvert par un autre code_mouvement (ex. maintenance annuelle facturee mensuellement par
   la compta sans BC), signale-le explicitement
3. Si marche renseigne des deux cotes, ils doivent matcher
4. Date de facture posterieure (ou simultanee a 30j pres) a la commande
5. Le montant alloue a chaque facture peut etre PARTIEL (ne pas depasser son montant_sf)
6. Plusieurs factures peuvent cumuler pour atteindre le reste a facturer

Reponds en JSON strict (pas de markdown, pas de fences):
{{
  "diagnostic": "MATCH" | "PARTIEL" | "DOUBLON" | "ORPHELINE" | "INDETERMINE",
  "confidence": <int 0-100>,
  "factures_a_lier": [
    {{
      "code_mouvement": "<str>",
      "montant_alloue": <float>,
      "raison": "<str max 100 char>"
    }}
  ],
  "raisonnement": "<str max 500 char, explication globale>",
  "action_suggeree": "VALIDER_AUTO" | "VALIDER_MANUEL" | "MARQUER_DOUBLON" | "RELANCER_FOURNISSEUR" | "INVESTIGUER"
}}"""


AI_VALID_DIAGNOSTICS = {"MATCH", "PARTIEL", "DOUBLON", "ORPHELINE", "INDETERMINE"}
AI_VALID_ACTIONS = {"VALIDER_AUTO", "VALIDER_MANUEL", "MARQUER_DOUBLON",
                    "RELANCER_FOURNISSEUR", "INVESTIGUER"}


def compute_candidates_hash(candidates: list) -> str:
    """Hash stable d'une liste de candidats (par code_mouvement trie).

    Utilise pour invalider le cache IA si les candidats changent.
    """
    keys = sorted(
        (c.get("code_mouvement") or "").strip()
        for c in candidates
    )
    payload = "|".join(keys)
    return hashlib.md5(payload.encode("utf-8")).hexdigest()[:16]


def build_ai_user_prompt(cmd: dict, candidates: list) -> str:
    """Construit le prompt utilisateur a partir d'une commande et de ses candidats."""
    factures_payload = []
    for c in candidates:
        factures_payload.append({
            "id": c.get("id"),
            "num_facture": c.get("num_facture"),
            "code_mouvement": c.get("code_mouvement"),
            "fournisseur": c.get("fournisseur"),
            "marche": c.get("marche"),
            "date_facture": c.get("date_facture"),
            "montant_sf": float(c.get("montant_service_fait") or 0.0),
            "libelle": c.get("libelle"),
        })
    factures_json = json.dumps(factures_payload, ensure_ascii=False, indent=2)

    return AI_USER_PROMPT_TEMPLATE.format(
        num_commande=cmd.get("num_commande") or "—",
        date_commande=cmd.get("date_commande") or "—",
        fournisseur=cmd.get("fournisseur") or "—",
        marche=cmd.get("marche") or "—",
        service_emetteur=cmd.get("service_emetteur") or "—",
        montant_ttc=f"{float(cmd.get('montant_ttc') or 0.0):.2f}",
        reste=f"{float(cmd.get('reste_a_facturer') or 0.0):.2f}",
        libelle=cmd.get("libelle") or "—",
        factures_json=factures_json,
    )


def validate_ai_response(parsed: dict) -> None:
    """Verifie la structure du JSON IA. Raise ValueError si invalide."""
    if not isinstance(parsed, dict):
        raise ValueError(f"Reponse IA non-dict : {type(parsed).__name__}")
    for key in ("diagnostic", "confidence", "factures_a_lier",
                "raisonnement", "action_suggeree"):
        if key not in parsed:
            raise ValueError(f"Cle manquante dans la reponse IA : {key}")
    if parsed["diagnostic"] not in AI_VALID_DIAGNOSTICS:
        raise ValueError(
            f"diagnostic IA invalide : {parsed['diagnostic']!r} "
            f"(attendu {AI_VALID_DIAGNOSTICS})"
        )
    if parsed["action_suggeree"] not in AI_VALID_ACTIONS:
        raise ValueError(
            f"action_suggeree invalide : {parsed['action_suggeree']!r}"
        )
    try:
        conf = int(parsed["confidence"])
    except (TypeError, ValueError):
        raise ValueError(f"confidence non-entiere : {parsed['confidence']!r}")
    if not 0 <= conf <= 100:
        raise ValueError(f"confidence hors plage 0-100 : {conf}")
    if not isinstance(parsed["factures_a_lier"], list):
        raise ValueError("factures_a_lier doit etre une liste")


class DeepSeekClient:
    """Client minimal pour l'API DeepSeek (compatible OpenAI chat/completions).

    Implementation stdlib uniquement (urllib.request) pour eviter d'ajouter
    'requests' aux dependances. Retry 3 tentatives avec backoff exponentiel
    sur erreurs reseau et 5xx ; fail immediat sur 4xx.
    """

    def __init__(self, api_key: str,
                 model: str = "deepseek-chat",
                 endpoint: str = "https://api.deepseek.com/v1/chat/completions",
                 timeout: int = 30):
        if not api_key:
            raise ValueError("api_key vide")
        self.api_key = api_key
        self.model = model
        self.endpoint = endpoint
        self.timeout = timeout

    def chat(self, system_prompt: str, user_prompt: str,
             max_tokens: int = 800, temperature: float = 0.1) -> str:
        """Appel synchrone, retourne le texte brut de la reponse (content)."""
        body = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "max_tokens": max_tokens,
            "temperature": temperature,
            "response_format": {"type": "json_object"},
        }
        data = json.dumps(body).encode("utf-8")
        last_exc = None
        for attempt in range(3):
            req = urllib.request.Request(
                self.endpoint,
                data=data,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                method="POST",
            )
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    raw = resp.read().decode("utf-8")
                payload = json.loads(raw)
                # Format chat/completions standard
                choices = payload.get("choices") or []
                if not choices:
                    raise ValueError(f"Pas de 'choices' dans la reponse : {raw[:200]}")
                content = choices[0].get("message", {}).get("content")
                if content is None:
                    raise ValueError(f"Pas de 'message.content' : {raw[:200]}")
                return content
            except urllib.error.HTTPError as e:
                # 4xx : pas de retry (cle invalide, prompt trop long, etc.)
                if 400 <= e.code < 500:
                    body_err = ""
                    try:
                        body_err = e.read().decode("utf-8", errors="replace")[:300]
                    except Exception:
                        pass
                    raise RuntimeError(
                        f"HTTP {e.code} {e.reason} : {body_err}"
                    ) from e
                last_exc = e
            except (urllib.error.URLError, TimeoutError, OSError) as e:
                last_exc = e
            # Backoff exponentiel : 1s, 2s, 4s
            if attempt < 2:
                time.sleep(2 ** attempt)
        raise RuntimeError(
            f"Echec de l'appel DeepSeek apres 3 tentatives : {last_exc}"
        ) from last_exc

    @staticmethod
    def parse_json_response(raw_text: str) -> dict:
        """Extrait le premier JSON object d'une reponse texte.

        Tolere :
        - JSON brut : {"a": 1}
        - Code fences markdown : ```json\\n{...}\\n```
        - Preambule : "Voici la reponse :\\n{...}"

        Raise ValueError si aucun JSON ne peut etre extrait.
        """
        if not isinstance(raw_text, str) or not raw_text.strip():
            raise ValueError("Reponse IA vide")
        text = raw_text.strip()

        # Premier essai : json.loads tel quel
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            pass

        # Deuxieme essai : retirer les fences markdown
        fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```",
                                text, re.DOTALL | re.IGNORECASE)
        if fence_match:
            try:
                data = json.loads(fence_match.group(1))
                if isinstance(data, dict):
                    return data
            except json.JSONDecodeError:
                pass

        # Troisieme essai : trouver le premier { ... } equilibre
        start = text.find("{")
        if start >= 0:
            depth = 0
            in_string = False
            escape = False
            for i in range(start, len(text)):
                ch = text[i]
                if escape:
                    escape = False
                    continue
                if ch == "\\":
                    escape = True
                    continue
                if ch == '"':
                    in_string = not in_string
                    continue
                if in_string:
                    continue
                if ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        candidate = text[start:i + 1]
                        try:
                            data = json.loads(candidate)
                            if isinstance(data, dict):
                                return data
                        except json.JSONDecodeError:
                            pass
                        break
        raise ValueError(f"Aucun JSON extractible : {text[:200]!r}")


# ============================================================================
# apply_ai_decision : politique de decision sur la reponse IA (spec 7.4)
# ============================================================================

ACTION_AUTO_DOUBLON = "AUTO_DOUBLON"
ACTION_AUTO_LINKED = "AUTO_LINKED"
ACTION_SUGGESTION = "SUGGESTION"
ACTION_IGNORED = "IGNORED"


def apply_ai_decision(conn: sqlite3.Connection,
                      link_repo: LinkRepository,
                      commande_id: int,
                      ai_response: dict,
                      auto_threshold: int = 90,
                      min_threshold: int = 40) -> str:
    """Applique la decision IA selon les seuils configures.

    Renvoie une des constantes ACTION_* indiquant ce qui a ete fait.
    Le commit SQL est gere ici. recompute_facturation reste a la charge de
    l'appelant (UI), pour eviter une dependance vers la classe Database.
    """
    if not isinstance(ai_response, dict):
        raise ValueError("ai_response doit etre un dict")
    confidence = int(ai_response.get("confidence", 0))
    diagnostic = ai_response.get("diagnostic")
    action = ai_response.get("action_suggeree")
    raisonnement = ai_response.get("raisonnement", "")[:1000]

    cur = conn.cursor()

    # Cas 1 : doublon administratif auto-marque
    if (diagnostic == "DOUBLON" and action == "MARQUER_DOUBLON"
            and confidence >= auto_threshold):
        cur.execute(
            "UPDATE commandes SET statut_metier = 'DOUBLON_ADMIN' WHERE id = ?",
            (commande_id,),
        )
        conn.commit()
        return ACTION_AUTO_DOUBLON

    # Cas 2 : rapprochement automatique
    if (diagnostic in ("MATCH", "PARTIEL") and action == "VALIDER_AUTO"
            and confidence >= auto_threshold):
        for f in ai_response.get("factures_a_lier", []):
            code_mvt = (f.get("code_mouvement") or "").strip()
            if not code_mvt:
                continue
            cur.execute(
                "SELECT id FROM factures WHERE code_mouvement = ? "
                "ORDER BY date_facture DESC LIMIT 1",
                (code_mvt,),
            )
            row = cur.fetchone()
            if row is None:
                # On ne peut pas resoudre le code -> on saute silencieusement
                continue
            facture_id = row["id"]
            try:
                montant = float(f.get("montant_alloue") or 0.0)
            except (TypeError, ValueError):
                continue
            if montant <= 0:
                continue
            try:
                link_repo.add_link(
                    commande_id=commande_id,
                    facture_id=facture_id,
                    montant_alloue=montant,
                    source="ai_auto",
                    confidence=confidence,
                    ai_model=None,
                    ai_reasoning=raisonnement,
                    created_by="ai",
                )
            except (sqlite3.IntegrityError, ValueError):
                # Lien deja existant ou contrainte -> on ignore et continue
                continue
        return ACTION_AUTO_LINKED

    # Cas 3 : suggestion en cache (le diag est deja persiste par ai_match)
    if confidence >= min_threshold:
        return ACTION_SUGGESTION

    return ACTION_IGNORED
