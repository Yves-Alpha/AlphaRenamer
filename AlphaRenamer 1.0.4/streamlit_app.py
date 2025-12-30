"""Streamlit front-end pour AlphaRenamer (renommage basé sur un lexique Excel)."""

from __future__ import annotations

import io
import shutil
import tempfile
import zipfile
from pathlib import Path

import pandas as pd  # force import pour s'assurer que deps sont installées
import streamlit as st

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
    st.set_page_config(page_title="AlphaRenamer Cloud", page_icon="🅰️", layout="centered")
    st.title("🅰️ AlphaRenamer — Cloud")
    st.markdown(
        "Renomme des fichiers PDF à partir d’un lexique Excel (colonnes NOCLI/NOMCLI). "
        "Charge un ZIP de fichiers, applique le lexique, puis télécharge le ZIP renommé."
    )

    lex_file = st.file_uploader("Lexique Excel (NOCLI / NOMCLI)", type=["xlsx"])
    zip_file = st.file_uploader("ZIP contenant les fichiers à renommer (PDF)", type=["zip"])
    dry_run = st.checkbox("Simulation (dry-run) — ne pas écrire les fichiers", value=False)

    if not (lex_file and zip_file):
        st.info("Charge le lexique et le ZIP pour lancer le renommage.")
        return

    if st.button("🚀 Lancer le renommage"):
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
