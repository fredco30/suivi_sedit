#!/usr/bin/env python3
"""Lanceur de l'application de suivi des marches.

Verifie les dependances, les installe si besoin, puis demarre l'application en
journalisant sa sortie dans run_logs/.

Appele par « Lancer_suivi_marches.cmd ». Le travail est fait ici plutot qu'en
batch parce que le script principal porte un accent dans son nom, ce que cmd.exe
gere mal selon la page de code active.

    python lanceur.py             # verifie les dependances puis lance
    python lanceur.py --maj       # reinstalle les dependances, puis lance
    python lanceur.py --verif     # verifie seulement, ne lance rien

L'installation se declenche si un module manque, ou si requirements.txt a change
depuis la derniere installation reussie -- une version relevee ou une dependance
ajoutee qui se trouve deja presente passerait sinon inapercue.
"""
from __future__ import annotations

import argparse
import glob
import hashlib
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
REQUIREMENTS = RACINE / "requirements.txt"
DOSSIER_LOGS = RACINE / "run_logs"
# Empreinte du requirements.txt tel qu'installé la dernière fois : c'est elle
# qui permet de repérer qu'il a changé alors que tout s'importe encore.
MARQUEUR = DOSSIER_LOGS / "requirements_installees.txt"
MOTIF_APPLICATION = "suivi_commandes_factures_marches_*.py"

# Modules dont l'absence empêche l'application de démarrer, et le paquet pip
# qui les fournit (les deux noms diffèrent parfois).
DEPENDANCES = {
    "PyQt5": "PyQt5",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",
}


def afficher_titre(texte: str) -> None:
    print()
    print("=" * 70)
    print(texte)
    print("=" * 70)


def trouver_application() -> Path:
    """Repère le script principal, sans avoir à écrire son nom accentué."""
    candidats = [Path(chemin) for chemin in glob.glob(str(RACINE / MOTIF_APPLICATION))]
    if not candidats:
        raise SystemExit(
            f"[ERREUR] Aucun script « {MOTIF_APPLICATION} » dans {RACINE}.\n"
            "         Le lanceur doit être placé dans le dossier de l'application."
        )
    if len(candidats) == 1:
        return candidats[0]

    # Plusieurs versions du script cohabitent : la plus récemment modifiée.
    candidats.sort(key=lambda chemin: chemin.stat().st_mtime, reverse=True)
    print("[INFO] Plusieurs scripts trouvés, le plus récent est retenu :")
    for chemin in candidats:
        horodatage = datetime.fromtimestamp(chemin.stat().st_mtime)
        marque = "  <--" if chemin is candidats[0] else ""
        print(f"       {chemin.name}  ({horodatage:%d/%m/%Y %H:%M}){marque}")
    return candidats[0]


def _empreinte_requirements() -> str:
    return hashlib.sha256(REQUIREMENTS.read_bytes()).hexdigest()


def requirements_ont_change() -> bool:
    """Le fichier a-t-il changé depuis la dernière installation réussie ?

    Vérifier que les modules s'importent ne suffit pas : une version relevée ou
    une dépendance ajoutée qui se trouve déjà présente passerait inaperçue. On
    compare donc l'empreinte du fichier à celle notée au dernier succès.
    """
    if not REQUIREMENTS.exists():
        return False
    if not MARQUEUR.exists():
        return True
    return MARQUEUR.read_text(encoding="utf-8").strip() != _empreinte_requirements()


def _noter_installation() -> None:
    """Retient l'empreinte du requirements.txt qui vient d'être installé."""
    try:
        MARQUEUR.parent.mkdir(exist_ok=True)
        MARQUEUR.write_text(_empreinte_requirements() + "\n", encoding="utf-8")
    except OSError:
        pass  # Sans marqueur, on réinstallera au prochain lancement : sans gravité.


def dependances_manquantes() -> list:
    manquantes = []
    for module, paquet in DEPENDANCES.items():
        try:
            __import__(module)
        except ImportError:
            manquantes.append(paquet)
    return manquantes


def _pip(*arguments) -> int:
    commande = [sys.executable, "-m", "pip", *arguments]
    print(f"[PIP] {' '.join(commande[2:])}")
    return subprocess.call(commande)


def _requirements_assouplies() -> Path:
    """Copie de requirements.txt dont les versions figées deviennent des minima.

    Certaines versions épinglées n'ont pas de binaire pour les Python récents
    (pandas 2.1.4 s'arrête à Python 3.12) : pip tenterait alors de compiler
    depuis les sources et échouerait faute de compilateur.
    """
    lignes = []
    for ligne in REQUIREMENTS.read_text(encoding="utf-8").splitlines():
        nettoyee = ligne.strip()
        if nettoyee and not nettoyee.startswith("#") and "==" in nettoyee:
            nettoyee = nettoyee.replace("==", ">=", 1)
        lignes.append(nettoyee)

    assouplies = DOSSIER_LOGS / "requirements_assouplies.txt"
    assouplies.parent.mkdir(exist_ok=True)
    assouplies.write_text("\n".join(lignes) + "\n", encoding="utf-8")
    return assouplies


def installer_dependances() -> bool:
    """Installe les dépendances, en assouplissant les versions si nécessaire."""
    if not REQUIREMENTS.exists():
        print(f"[ERREUR] {REQUIREMENTS.name} introuvable.")
        return False

    afficher_titre("INSTALLATION DES DÉPENDANCES")
    if _pip("install", "-r", str(REQUIREMENTS)) == 0:
        _noter_installation()
        return True

    print()
    print("[INFO] Installation refusée avec les versions figées.")
    print(f"       Python {sys.version_info.major}.{sys.version_info.minor} n'a "
          "peut-être pas de binaire pour l'une d'elles.")
    print("       Nouvelle tentative en les traitant comme des versions minimales.")

    if _pip("install", "-r", str(_requirements_assouplies())) == 0:
        _noter_installation()
        print()
        print("[OK] Installation réussie avec des versions plus récentes.")
        print("     Signalez-le : requirements.txt gagnerait à être mis à jour.")
        return True

    print()
    print("[ERREUR] Installation impossible. Pistes :")
    print("   - vérifier la connexion réseau (ou le proxy de l'entreprise) ;")
    print("   - lancer une invite de commandes en administrateur ;")
    print(f"   - installer à la main : {sys.executable} -m pip install "
          "PyQt5 pandas openpyxl xlrd")
    return False


def lancer_application(script: Path) -> int:
    """Démarre l'application en journalisant sa sortie."""
    DOSSIER_LOGS.mkdir(exist_ok=True)
    journal = DOSSIER_LOGS / f"AE_RUN_{datetime.now():%Y%m%d_%H%M%S}.log"

    afficher_titre(f"DÉMARRAGE — {script.name}")
    print(f"Journal : {journal}")
    print()

    with open(journal, "w", encoding="utf-8", errors="replace") as sortie:
        sortie.write(f"Lancement : {datetime.now():%d/%m/%Y %H:%M:%S}\n")
        sortie.write(f"Script    : {script}\n")
        sortie.write(f"Python    : {sys.version}\n")
        sortie.write(f"Exécutable: {sys.executable}\n")
        sortie.write("=" * 70 + "\n\n")
        sortie.flush()

        processus = subprocess.Popen(
            [sys.executable, "-X", "dev", "-u", str(script)],
            cwd=str(RACINE), stdout=sortie, stderr=subprocess.STDOUT,
        )
        code = processus.wait()

    if code == 0:
        print("[OK] Application fermée normalement.")
        return 0

    print(f"[ERREUR] L'application s'est arrêtée (code {code}).")
    print(f"         Détail dans {journal}")
    _ouvrir(journal)
    return code


def _ouvrir(chemin: Path) -> None:
    """Ouvre un fichier avec l'application par défaut du système."""
    try:
        if os.name == "nt":
            os.startfile(str(chemin))  # noqa: S606 - Windows uniquement
        elif sys.platform == "darwin":
            subprocess.call(["open", str(chemin)])
        else:
            subprocess.call(["xdg-open", str(chemin)])
    except Exception:
        pass  # L'ouverture du journal est un confort, pas une obligation.


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--maj", action="store_true",
                        help="Réinstaller les dépendances avant de lancer")
    parser.add_argument("--verif", action="store_true",
                        help="Vérifier les dépendances sans lancer l'application")
    arguments = parser.parse_args(argv)

    afficher_titre("SUIVI DES MARCHÉS — LANCEUR")
    print(f"Dossier : {RACINE}")
    print(f"Python  : {sys.version.split()[0]}  ({sys.executable})")

    manquantes = dependances_manquantes()
    if manquantes:
        print(f"Manquant: {', '.join(manquantes)}")
    else:
        print("Modules : tous présents")

    change = requirements_ont_change()
    if change:
        print("Requis  : requirements.txt a changé depuis la dernière installation")

    if arguments.maj or manquantes or change:
        if not installer_dependances():
            return 1
        encore = dependances_manquantes()
        if encore:
            print(f"\n[ERREUR] Toujours manquant après installation : {', '.join(encore)}")
            return 1
        print("\n[OK] Toutes les dépendances sont installées.")

    if arguments.verif:
        print("\nVérification terminée, l'application n'a pas été lancée (--verif).")
        return 0

    return lancer_application(trouver_application())


if __name__ == "__main__":
    sys.exit(main())
