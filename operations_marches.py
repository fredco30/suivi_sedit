"""Rattachement des marchés à leur opération, quand la codification s'y refuse.

`MarchesAnalyzer.extract_operation` retire du code marché un dernier segment
numérique de un ou deux chiffres, et y voit un n° de lot. La règle tient pour
`2024_17_3`, mais elle ne peut rien deviner ailleurs :

- `MC157_01` et `MC157_02` n'ont qu'un séparateur : chacun devient sa propre
  opération, alors que `MC109_1_02` et `MC109_1_03` se regroupent bien ;
- `MC152_3_2A` et `MC152_3_2B` ont un suffixe non numérique : deux opérations ;
- `2019_06P1` à `2019_06P3R`, `2020_14G1` à `2020_14GO` : sept ou quatre
  opérations là où il n'y a vraisemblablement qu'une affaire.

Aucun programme ne peut trancher : `P2` est-il une phase de `2019_06` ou une
opération à part ? Ce module ne tranche donc pas — il **rassemble les cas
douteux** pour qu'ils soient arbitrés une fois, et le rattachement saisi vaut
ensuite pour l'écran comme pour le traitement en lot.
"""

from typing import Dict, Iterable, List, Sequence


def racine_codification(code: str) -> str:
    """Souche à laquelle rattacher un code, pour comparer les voisins.

    Deux familles de codes cohabitent dans les exports :

    - `ANNÉE_NUMÉRO…` — la souche est l'année et le numéro, sans ce qui suit :
      `2019_06P1` et `2019_06P3R` donnent tous deux `2019_06`, tandis que
      `2019_08` reste distinct ;
    - `MCnnn_…` — la souche est le premier segment : `MC157_01` et `MC157_02`
      donnent `MC157`, mais `MC109` et `MC113` restent distincts.

    Ce n'est qu'un critère de **rapprochement visuel** : il ne décide rien, il
    met côte à côte des codes qui méritent un coup d'œil.
    """
    code = str(code or "").strip()
    if not code:
        return ""

    segments = code.replace("-", "_").split("_")
    tete = segments[0]

    if not (len(tete) == 4 and tete.isdigit()):
        return tete

    if len(segments) < 2:
        return tete

    # Année : on garde les chiffres de tête du deuxième segment, et rien d'autre.
    chiffres = ""
    for caractere in segments[1]:
        if not caractere.isdigit():
            break
        chiffres += caractere
    return f"{tete}_{chiffres}" if chiffres else tete


def operations_a_arbitrer(operations: Iterable[str]) -> Dict[str, List[str]]:
    """Souches portant plusieurs opérations — les regroupements à confirmer.

    Returns:
        {souche: [codes opération, triés]} pour les seules souches qui en
        comptent au moins deux. Une souche à une seule opération n'appelle
        aucun arbitrage.
    """
    par_racine: Dict[str, List[str]] = {}
    for operation in operations:
        code = str(operation or "").strip()
        if not code:
            continue
        par_racine.setdefault(racine_codification(code), []).append(code)

    return {
        racine: sorted(set(codes))
        for racine, codes in par_racine.items()
        if len(set(codes)) > 1
    }


def lot_isole(code_marche: str, code_operation: str, marches_operation: Sequence[str]) -> bool:
    """Vrai quand l'opération n'a qu'un lot alors que le code en annonce d'autres.

    `2020_24` n'apparaît que par son lot `2020_24_7` : la colonne « Nb lots »
    affiche 1, ce qui se lit « opération à lot unique » alors qu'il faut lire
    « un seul lot facturé ». Les autres lots existent peut-être, sans écriture
    dans les exports SEDIT.
    """
    marche = str(code_marche or "").strip()
    operation = str(code_operation or "").strip()
    return (
        len(marches_operation) == 1
        and bool(marche)
        and bool(operation)
        and marche != operation
        and marche.startswith(operation)
    )
