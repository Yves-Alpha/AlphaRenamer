# Package portable v1.0.4

Prerequis :
- macOS avec Python 3.11/3.12 installe (idealement depuis python.org pour Tk).
- Acces reseau pour `pip` (ou un cache/miroir via `PIP_INDEX_URL`).

Construction :
1. `cd /Users/yvesnowak/Documents/TEST APP/AlphaRenamer/AlphaRenamer\ 1.0.4`
2. `./build_portable.sh` (ou `./build_portable.sh --clean` pour recreer le venv)

Ce que fait le script :
- recopie le code source le plus recent dans `PDF-Renommage.app/Contents/Resources/app/`
- construit/rafraichit un venv portable sous `PDF-Renommage.app/Contents/Resources/app/venv`
- installe les dependances listees dans `Installer/requirements.txt`
- produit `dist/PDF-Renommage-portable-1.0.4.zip` (archive prete a distribuer)
- inclut le renommage via lexique (LEXIQUE.xlsx) : bouton dedie + option "renommer apres traitement"

Utilisation du package portable :
- dezippe l'archive ou tu veux, puis lance `PDF-Renommage.app`
- l'app utilise d'abord le venv embarque (`./Contents/Resources/app/venv`), puis tombe sur la venv utilisateur si besoin

Notes :
- tu peux supprimer `./Contents/Resources/app/venv` avant d'archiver pour forcer une reconstruction propre
- les logs restent dans `~/Library/Logs/PDF-Renommage/launcher.log`
