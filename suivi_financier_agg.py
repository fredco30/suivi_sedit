"""Agregation des ecritures SEDIT pour le suivi financier par marche.

Ce module isole la regle metier du generateur Excel : il transforme une liste
d'ecritures brutes (une ligne par mouvement SEDIT) en une liste de lignes de
suivi (une ligne par etat de bon de commande).

Le defaut corrige ici est un double comptage de l'enveloppe : le generateur
emettait une ligne d'engagement ET une ligne de facture pour un meme BDC, sans
jamais solder l'engagement par la facture. Le report d'exercice ajoutait une
seconde ligne d'engagement identique.

Regle d'agregation, pour chaque BDC :

    montant_ref = max(montant_bdc_declare, somme des factures mandatees)
    facture     = somme des factures mandatees
    reliquat    = montant_ref - facture

Le tableau emet alors :
  1. une ligne par facture mandatee, telle quelle ;
  2. une seule ligne d'engagement, uniquement si reliquat > 0, portant le
     montant du reliquat et non le montant initial du BDC.

Invariant garanti (verifie par les tests) :
    somme des montants imputes == somme sur BDC distincts de montant_ref
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence

# Tolerance de comparaison sur les montants (en euros).
# Les montants viennent d'exports flottants : 0,005 € evite qu'une erreur
# d'arrondi ne fasse apparaitre un reliquat fantome.
TOLERANCE = 0.005

STATUT_FACTURE = "FACTURÉ"
STATUT_ENGAGEMENT = "ENGAGEMENT — reliquat non facturé"

# Longueur de troncature de la designation. La troncature se fait sur le
# libelle nettoye, jamais avant concatenation d'un suffixe (cf. `(REPORT)`).
LONGUEUR_DESIGNATION = 60

# Motifs de neutralisation, traces dans l'onglet « Lignes neutralisées ».
MOTIF_SOLDE = "BDC soldé par ses factures — engagement supprimé"
MOTIF_REPORT = "Report d'exercice — engagement déjà porté par l'exercice précédent"
MOTIF_DOUBLON = "Engagement déjà porté par une autre ligne du même BDC"
MOTIF_AJUSTEMENT = "Engagement conservé, ramené au reliquat non facturé"

# Etats d'un BDC, repris dans l'onglet « A jour ».
ETAT_SOLDE = "SOLDÉ"
ETAT_PARTIEL = "PARTIELLEMENT FACTURÉ"
ETAT_NON_FACTURE = "NON FACTURÉ"


@dataclass
class Ecriture:
    """Une ecriture SEDIT brute, telle que lue dans le fichier des factures."""

    marche: str = ""
    fournisseur: str = ""
    tranche_libelle: str = ""
    num_commande: str = ""
    libelle: str = ""
    num_facture: str = ""
    num_mandat: str = ""
    date_sf: str = ""
    montant_ttc: float = 0.0
    montant_sf: float = 0.0
    montant_initial: float = 0.0
    type_marche: str = "CLASSIQUE"

    @property
    def cle_bdc(self) -> str:
        """Cle de regroupement : le n° de BDC, ou la tranche a defaut."""
        return self.num_commande or self.tranche_libelle

    @property
    def est_facture(self) -> bool:
        """Une ecriture est une realisation si elle porte facture ET mandat."""
        return bool(str(self.num_facture).strip()) and bool(str(self.num_mandat).strip())


@dataclass
class LigneSuivi:
    """Une ligne emise dans l'onglet FINANCIER."""

    marche: str = ""
    fournisseur: str = ""
    tranche_libelle: str = ""
    cle_bdc: str = ""
    designation: str = ""
    montant_ref: float = 0.0
    num_facture: str = ""
    num_mandat: str = ""
    date_sf: str = ""
    montant_impute: float = 0.0
    statut: str = STATUT_FACTURE

    @property
    def montant_ht(self) -> float:
        """Montant HT reconstitue depuis le TTC (TVA 20 %)."""
        return self.montant_impute / 1.2


@dataclass
class Neutralisation:
    """Trace d'une ligne d'origine supprimee ou ramenee a son reliquat."""

    marche: str = ""
    cle_bdc: str = ""
    designation: str = ""
    montant_origine: float = 0.0
    montant_retenu: float = 0.0
    motif: str = ""

    @property
    def montant_neutralise(self) -> float:
        return self.montant_origine - self.montant_retenu


@dataclass
class Anomalie:
    """Avertissement demandant un arbitrage humain."""

    marche: str = ""
    cle_bdc: str = ""
    message: str = ""
    valeurs: str = ""
    valeur_retenue: float = 0.0


@dataclass
class EtatBdc:
    """Synthese d'un BDC, alimente l'onglet « A jour »."""

    marche: str = ""
    fournisseur: str = ""
    tranche_libelle: str = ""
    cle_bdc: str = ""
    designation: str = ""
    montant_ref: float = 0.0
    montant_facture: float = 0.0
    reliquat: float = 0.0
    etat: str = ETAT_NON_FACTURE


@dataclass
class ResultatSuivi:
    """Resultat complet de l'agregation, pour un groupe (prestataire, tranche)."""

    lignes: List[LigneSuivi] = field(default_factory=list)
    bdcs: List[EtatBdc] = field(default_factory=list)
    neutralisations: List[Neutralisation] = field(default_factory=list)
    anomalies: List[Anomalie] = field(default_factory=list)

    @property
    def total_impute(self) -> float:
        return sum(ligne.montant_impute for ligne in self.lignes)

    @property
    def total_facture(self) -> float:
        return sum(l.montant_impute for l in self.lignes if l.statut == STATUT_FACTURE)

    @property
    def total_engagement(self) -> float:
        return sum(l.montant_impute for l in self.lignes if l.statut == STATUT_ENGAGEMENT)

    @property
    def total_bdc_distincts(self) -> float:
        """Somme des montants de reference sur BDC distincts.

        Jamais une somme de colonne : un BDC ne compte qu'une fois, quel que
        soit son nombre de lignes.
        """
        return sum(bdc.montant_ref for bdc in self.bdcs)

    def etendre(self, autre: "ResultatSuivi") -> None:
        self.lignes.extend(autre.lignes)
        self.bdcs.extend(autre.bdcs)
        self.neutralisations.extend(autre.neutralisations)
        self.anomalies.extend(autre.anomalies)


def nettoyer_designation(libelle, longueur: int = LONGUEUR_DESIGNATION) -> str:
    """Nettoie un libelle SEDIT et le tronque *apres* nettoyage.

    Le suffixe `(REPORT)` marquait un report d'exercice ; il est desormais porte
    par la colonne STATUT. Il etait de surcroit concatene *apres* troncature,
    ce qui produisait des libelles ampute du type
    `G3P extension du reseau video su(REPORT)`.
    """
    texte = "" if libelle is None else str(libelle)
    texte = texte.replace("(REPORT)", " ").replace("(REPORT", " ")
    return " ".join(texte.split())[:longueur]


def _est_report(libelle) -> bool:
    return "(REPORT" in ("" if libelle is None else str(libelle))


def _montant_reference(
    cle: str,
    ecritures: Sequence[Ecriture],
    montant_declare: Optional[float],
    montant_facture: float,
    anomalies: List[Anomalie],
) -> float:
    """Determine le montant de reference d'un BDC et signale les divergences.

    Ordre des sources : montant declare (base commandes) puis, a defaut, le
    montant initial porte par les ecritures, puis le plus gros engagement.
    Quand plusieurs sources divergent, on retient le maximum ET on leve un
    avertissement : c'est un arbitrage humain, pas une decision du programme.
    """
    marche = ecritures[0].marche if ecritures else ""

    candidats: List[float] = []
    if montant_declare and montant_declare > TOLERANCE:
        candidats.append(float(montant_declare))

    montants_initiaux = sorted({
        round(float(e.montant_initial), 2)
        for e in ecritures
        if e.montant_initial and float(e.montant_initial) > TOLERANCE
    })
    candidats.extend(montants_initiaux)

    if not candidats:
        # Aucune source declarative : l'engagement le plus eleve fait foi.
        engagements = [float(e.montant_ttc) for e in ecritures if not e.est_facture]
        if engagements:
            candidats.append(max(engagements))

    distincts = sorted({round(v, 2) for v in candidats})
    if len(distincts) > 1:
        anomalies.append(Anomalie(
            marche=marche,
            cle_bdc=cle,
            message="Plusieurs montants de BDC remontés par la source : maximum retenu",
            valeurs=" / ".join(f"{v:,.2f}".replace(",", " ") for v in distincts),
            valeur_retenue=distincts[-1],
        ))

    montant_ref = distincts[-1] if distincts else 0.0

    if montant_facture - montant_ref > TOLERANCE:
        # Un BDC sans aucun montant declare n'est pas un cas d'arbitrage :
        # le facture fait foi. On n'avertit que si une source existait bien.
        if montant_ref > TOLERANCE:
            anomalies.append(Anomalie(
                marche=marche,
                cle_bdc=cle,
                message="Facturé supérieur au montant du BDC : montant facturé retenu",
                valeurs=f"BDC {montant_ref:.2f} / facturé {montant_facture:.2f}",
                valeur_retenue=montant_facture,
            ))
        montant_ref = montant_facture

    return montant_ref


def _etat_bdc(montant_facture: float, reliquat: float) -> str:
    if reliquat <= TOLERANCE:
        return ETAT_SOLDE
    if montant_facture > TOLERANCE:
        return ETAT_PARTIEL
    return ETAT_NON_FACTURE


def agreger_ecritures(
    ecritures: Iterable[Ecriture],
    montants_declares: Optional[Dict[str, float]] = None,
    trier_par_bdc: bool = False,
) -> ResultatSuivi:
    """Agrege des ecritures SEDIT en lignes de suivi : une ligne par etat.

    Args:
        ecritures: ecritures brutes, une par mouvement SEDIT.
        montants_declares: montant de BDC declare en base, par n° de BDC.
        trier_par_bdc: trie les BDC par n° plutot que par ordre d'apparition.

    Returns:
        Un `ResultatSuivi` portant les lignes emises, la synthese par BDC, les
        lignes neutralisees et les anomalies a arbitrer.
    """
    montants_declares = montants_declares or {}
    resultat = ResultatSuivi()

    # Regroupement par BDC en conservant l'ordre d'apparition.
    groupes: Dict[str, List[Ecriture]] = {}
    for ecriture in ecritures:
        groupes.setdefault(ecriture.cle_bdc, []).append(ecriture)

    cles = sorted(groupes) if trier_par_bdc else list(groupes)

    for cle in cles:
        lignes_bdc = groupes[cle]
        factures = [e for e in lignes_bdc if e.est_facture]
        engagements = [e for e in lignes_bdc if not e.est_facture]

        montant_facture = sum(float(e.montant_ttc) for e in factures)
        montant_ref = _montant_reference(
            cle, lignes_bdc, montants_declares.get(cle), montant_facture, resultat.anomalies
        )
        reliquat = round(montant_ref - montant_facture, 2)
        if abs(reliquat) <= TOLERANCE:
            reliquat = 0.0

        reference = lignes_bdc[0]

        # 1. Une ligne par facture mandatee : elles portent la realite comptable.
        if trier_par_bdc:
            factures = sorted(factures, key=lambda e: str(e.num_facture))
        for ecriture in factures:
            resultat.lignes.append(LigneSuivi(
                marche=ecriture.marche,
                fournisseur=ecriture.fournisseur,
                tranche_libelle=ecriture.tranche_libelle,
                cle_bdc=cle,
                designation=nettoyer_designation(ecriture.libelle),
                montant_ref=montant_ref,
                num_facture=str(ecriture.num_facture),
                num_mandat=str(ecriture.num_mandat),
                date_sf=ecriture.date_sf,
                montant_impute=float(ecriture.montant_ttc),
                statut=STATUT_FACTURE,
            ))

        # 2. Une seule ligne d'engagement, et seulement s'il reste a facturer.
        #    Le report d'exercice ne cree pas de ligne : il prolonge l'engagement.
        conservee = engagements[0] if engagements else None
        if reliquat > 0:
            porteuse = conservee or reference
            resultat.lignes.append(LigneSuivi(
                marche=porteuse.marche,
                fournisseur=porteuse.fournisseur,
                tranche_libelle=porteuse.tranche_libelle,
                cle_bdc=cle,
                designation=nettoyer_designation(porteuse.libelle),
                montant_ref=montant_ref,
                num_facture="",
                num_mandat="",
                date_sf="",
                montant_impute=reliquat,
                statut=STATUT_ENGAGEMENT,
            ))

        # 3. Trace des lignes d'origine supprimees ou ramenees a leur reliquat.
        for rang, ecriture in enumerate(engagements):
            montant_origine = float(ecriture.montant_ttc)
            est_conservee = reliquat > 0 and ecriture is conservee
            montant_retenu = reliquat if est_conservee else 0.0

            if est_conservee:
                if abs(montant_origine - reliquat) <= TOLERANCE:
                    continue  # ligne inchangee : rien a tracer
                motif = MOTIF_AJUSTEMENT
            elif _est_report(ecriture.libelle):
                motif = MOTIF_REPORT
            elif rang > 0:
                motif = MOTIF_DOUBLON
            else:
                motif = MOTIF_SOLDE

            resultat.neutralisations.append(Neutralisation(
                marche=ecriture.marche,
                cle_bdc=cle,
                designation=nettoyer_designation(ecriture.libelle),
                montant_origine=montant_origine,
                montant_retenu=montant_retenu,
                motif=motif,
            ))

        resultat.bdcs.append(EtatBdc(
            marche=reference.marche,
            fournisseur=reference.fournisseur,
            tranche_libelle=reference.tranche_libelle,
            cle_bdc=cle,
            designation=nettoyer_designation(reference.libelle),
            montant_ref=montant_ref,
            montant_facture=montant_facture,
            reliquat=reliquat,
            etat=_etat_bdc(montant_facture, reliquat),
        ))

    return resultat
