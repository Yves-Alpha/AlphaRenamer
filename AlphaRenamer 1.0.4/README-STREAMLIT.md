# AlphaRenamer (Streamlit Cloud)

## Fichiers à pousser
- `streamlit_app.py`
- dossier `app/` (contient `rename_from_lexique.py`)
- `requirements.txt`
- (optionnel) `.gitignore`

## Déploiement Streamlit Cloud
1. Repo GitHub : inclure les fichiers ci-dessus (pas besoin des `.app`, `dist/`, `venv/`).
2. Sur Streamlit Cloud : New app → choisir le repo/branche → fichier principal `streamlit_app.py`.
3. `requirements.txt` installe streamlit, pandas, openpyxl (pour lire l’Excel). Aucun binaire système particulier.

## Usage (UI Cloud)
1. Charger le lexique Excel (colonnes `NOCLI` et `NOMCLI`).
2. Charger un ZIP contenant les fichiers à renommer (PDF).
3. Option "Simulation (dry-run)" si besoin.
4. Lancer : un ZIP renommé est proposé en téléchargement (sauf en dry-run).

## Test local rapide
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```
