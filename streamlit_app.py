"""Streamlit front-end pour AlphaRenamer (renommage basé sur un lexique Excel)."""

from __future__ import annotations

import io
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import pandas as pd  # force import pour s'assurer que deps sont installées
import streamlit as st

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.rename_from_lexique import LexiqueError, rename_with_lexique


def rezip_folder(src: Path, dest_zip: Path) -> Path:
    """Crée une archive ZIP du contenu de src."""
    with zipfile.ZipFile(dest_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for f in src.rglob("*"):
            if f.is_file():
                arcname = f.relative_to(src)
                zf.write(f, arcname)
    return dest_zip


def main():
    st.set_page_config(page_title="Renommage de PDF", page_icon="🧾", layout="centered")
    st.title("🧾 Renommage de PDF")
    st.markdown(
        "Guidé pour la compta :\n"
        "1) Exporter les PDFs Word (publipostage) en ZIP.\n"
        "2) Charger le **lexique Excel** (colonnes `NOCLI` et `NOMCLI`).\n"
        "3) Charger le **ZIP de PDFs**.\n"
        "4) Lancer le renommage, puis télécharger le ZIP renommé."
    )

    with st.expander("Infos rapides"):
        st.markdown(
            "- Le lexique doit contenir les colonnes **NOCLI** (code client) et **NOMCLI** (nom magasin).\n"
            "- Les PDFs sont lus et renommés en fonction du code trouvé dans le texte (OCR déjà géré dans le backend Word → PDF).\n"
            "- Option **Simulation** : permet de vérifier les correspondances sans écrire de fichiers."
        )

    lex_file = st.file_uploader("Étape 1 — Lexique Excel (NOCLI / NOMCLI)", type=["xlsx"], help="Fichier Excel du tableau clients")
    zip_file = st.file_uploader("Étape 2 — ZIP des PDFs à renommer", type=["zip"], help="ZIP contenant les PDFs générés depuis Word")
    dry_run = st.checkbox("Étape 3 — Simulation (dry-run) : ne pas écrire les fichiers", value=False)

    if not (lex_file and zip_file):
        st.info("Charge le lexique et le ZIP pour lancer le renommage.")
        return

    if st.button("Étape 4 — 🚀 Lancer le renommage"):
        with st.spinner("Traitement en cours..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmpdir_path = Path(tmpdir)
                    # Sauvegarde des uploads
                    lex_path = tmpdir_path / lex_file.name
                    lex_path.write_bytes(lex_file.getvalue())

                    zip_path = tmpdir_path / zip_file.name
                    zip_path.write_bytes(zip_file.getvalue())

                    # Extraction du ZIP
                    extract_dir = tmpdir_path / "extracted"
                    extract_dir.mkdir(parents=True, exist_ok=True)
                    with zipfile.ZipFile(zip_path, "r") as zf:
                        zf.extractall(extract_dir)

                    renamed, skipped, errors = rename_with_lexique(
                        extract_dir, lex_path, dry_run=dry_run, allowed_ext=("pdf",)
                    )

                    st.success(f"Terminé : renommés {renamed} | inchangés {skipped} | erreurs {errors}")

                    if not dry_run:
                        out_zip = tmpdir_path / "renamed.zip"
                        rezip_folder(extract_dir, out_zip)
                        st.download_button(
                            "⬇️ Télécharger le ZIP renommé",
                            data=out_zip.read_bytes(),
                            file_name="AlphaRenamer.zip",
                            mime="application/zip",
                        )
                    else:
                        st.info("Simulation uniquement (dry-run) : aucun fichier écrit. Décoche l’option pour produire le ZIP.")
            except LexiqueError as e:
                st.error(f"Erreur lexique : {e}")
            except Exception as e:
                st.error(f"Erreur inattendue : {type(e).__name__} — {e}")


if __name__ == "__main__":
    main()
