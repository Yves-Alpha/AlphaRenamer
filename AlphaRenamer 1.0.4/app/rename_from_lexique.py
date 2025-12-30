#!/usr/bin/env python3
"""
Renommage de fichiers a partir d'un lexique Excel (NOCLI/NOMCLI).
Fonctions reutilisables en GUI + CLI.
"""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, Iterable, Tuple

import pandas as pd


class LexiqueError(Exception):
    """Erreur fonctionnelle sur le lexique."""


def normalize_text(txt: str) -> str:
    """Supprime accents, contenu entre parentheses et nettoie les espaces."""
    if not isinstance(txt, str):
        txt = str(txt)

    txt = re.sub(r"\s*\([^)]*\)", "", txt)
    nfkd_form = unicodedata.normalize("NFD", txt)
    without_accents = "".join(c for c in nfkd_form if not unicodedata.combining(c))
    without_accents = re.sub(r"[^\w\s\.-]", " ", without_accents)
    without_accents = re.sub(r"\s+", " ", without_accents)
    return without_accents.strip()


def load_lexique(path: Path) -> Dict[str, str]:
    """Charge le lexique et retourne un dict code -> nom normalise."""
    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise LexiqueError(f"Lecture du lexique impossible ({path}) : {e}") from e

    cols = {c.upper(): c for c in df.columns}
    if "NOCLI" not in cols or "NOMCLI" not in cols:
        raise LexiqueError(f"Colonnes NOCLI/NOMCLI introuvables. Colonnes detectees: {list(df.columns)}")

    code_col = cols["NOCLI"]
    nom_col = cols["NOMCLI"]

    mapping: Dict[str, str] = {}
    for _, row in df.iterrows():
        code = str(row[code_col]).strip()
        if not code or code.lower() == "nan":
            continue
        nom = str(row[nom_col]).strip()
        if not nom or nom.lower() == "nan":
            continue

        nom_clean = normalize_text(nom)
        nom_clean = re.sub(r"\s+", " ", nom_clean).strip().upper()
        mapping[code] = nom_clean

    if not mapping:
        raise LexiqueError("Le lexique ne contient aucune ligne exploitable (NOCLI/NOMCLI).")

    return mapping


def find_year_token(tokens: Iterable[str]):
    """Retourne l'index du dernier token qui ressemble a une annee (4 chiffres)."""
    year_idx = None
    for i, t in enumerate(tokens):
        if re.fullmatch(r"\d{4}", t):
            year = int(t)
            if 1900 <= year <= 2100:
                year_idx = i
    return year_idx


def rename_file(path: Path, lexique_mapping: Dict[str, str], *, dry_run: bool = False) -> str:
    """Renomme un fichier selon le lexique. Retourne 'renamed' ou 'skipped'."""
    dirpath = path.parent
    filename = path.name
    base = path.stem
    ext = path.suffix

    normalized_base = normalize_text(base)
    tokens = [t for t in re.split(r"_+", normalized_base) if t]
    if not tokens:
        return "skipped"

    codes = set(lexique_mapping.keys())
    code_idx = None
    code_value = None
    for i, t in enumerate(tokens):
        if t in codes:
            code_idx = i
            code_value = t
            break

    if code_idx is not None:
        year_idx = find_year_token(tokens)
        prefix_tokens = tokens[:code_idx]
        suffix_tokens = []
        year_token = None
        if year_idx is not None:
            year_token = tokens[year_idx]
            suffix_tokens = tokens[year_idx + 1 :]
        client_tokens = [code_value, lexique_mapping.get(code_value, tokens[code_idx])]
        new_tokens = prefix_tokens + client_tokens
        if year_token:
            new_tokens.append(year_token)
        new_tokens += suffix_tokens
    else:
        new_tokens = tokens

    new_base = "_".join(new_tokens)
    new_base = re.sub(r"_+", "_", new_base).strip("_")
    new_name = new_base + ext

    if new_name == filename:
        return "skipped"

    new_path = dirpath / new_name
    if new_path.exists():
        i = 2
        while True:
            candidate = dirpath / f"{new_base}_v{i}{ext}"
            if not candidate.exists():
                new_path = candidate
                new_name = candidate.name
                break
            i += 1

    if dry_run:
        return "renamed"

    path.rename(new_path)
    return "renamed"


def rename_with_lexique(
    folder: Path,
    lexique_path: Path,
    *,
    dry_run: bool = False,
    allowed_ext: Iterable[str] = ("pdf",),
    skip_dirs: Iterable[str] = ("_ERREURS",),
) -> Tuple[int, int, int]:
    """Renomme les fichiers du dossier selon le lexique. Retourne (renamed, skipped, errors)."""
    mapping = load_lexique(lexique_path)
    folder = folder.resolve()
    allowed = {e.lower().lstrip(".") for e in allowed_ext}
    skip_set = {s.lower() for s in skip_dirs}

    renamed = 0
    skipped = 0
    errors = 0

    for f in folder.rglob("*"):
        if f.is_dir():
            continue
        rel_parts = [p.lower() for p in f.relative_to(folder).parts]
        if any(part in skip_set for part in rel_parts):
            continue
        if f.suffix.lower().lstrip(".") not in allowed:
            continue

        try:
            status = rename_file(f, mapping, dry_run=dry_run)
            if status == "renamed":
                renamed += 1
            else:
                skipped += 1
        except Exception:
            errors += 1

    return renamed, skipped, errors


def main():
    parser = argparse.ArgumentParser(
        description="Renommage de fichiers a partir d'un lexique (NOCLI/NOMCLI) avec normalisation."
    )
    parser.add_argument("-d", "--dossier", required=True, help="Dossier contenant les fichiers a renommer.")
    parser.add_argument("-l", "--lexique", required=True, help="Fichier Excel du lexique (ex: LEXIQUE.xlsx).")
    parser.add_argument("--dry-run", action="store_true", help="Simulation uniquement (aucun renommage).")
    args = parser.parse_args()

    folder = Path(args.dossier).expanduser().resolve()
    lex_path = Path(args.lexique).expanduser().resolve()

    if not folder.is_dir():
        print(f"Le dossier '{folder}' n'existe pas ou n'est pas un dossier.", file=sys.stderr)
        sys.exit(1)
    if not lex_path.is_file():
        print(f"Le fichier lexique '{lex_path}' n'existe pas.", file=sys.stderr)
        sys.exit(1)

    try:
        renamed, skipped, errors = rename_with_lexique(folder, lex_path, dry_run=args.dry_run)
    except LexiqueError as e:
        print(e, file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Erreur inattendue: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Renommage termine. Renommes: {renamed}, inchanges: {skipped}, erreurs: {errors}")


if __name__ == "__main__":
    main()
