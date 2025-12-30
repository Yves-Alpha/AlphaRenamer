#!/usr/bin/env bash
set -euxo pipefail

# ============================================================
# Installeur hors-ligne pour PDF-Renommage (v1.0.4)
# - Crée/Met à jour une venv UTILISATEUR
# - Vérifie Tkinter
# - Installe les dépendances (requirements.txt)
# - Vérifie/installe docx2pdf et valide son import
# - Log complet: ~/Library/Logs/PDF-Renommage/install.log
# ============================================================

APP_NAME="PDF-Renommage"
USER_BASE="$HOME/Library/Application Support/$APP_NAME"
VENV="$USER_BASE/venv"
LOGDIR="$HOME/Library/Logs/$APP_NAME"
LOGFILE="$LOGDIR/install.log"

mkdir -p "$USER_BASE" "$LOGDIR"
exec >>"$LOGFILE" 2>&1

echo "=== $(date) : Installation démarrée ==="
echo "ENV: PATH=$PATH"
echo "ENV: HOME=$HOME"
echo "APP_NAME=$APP_NAME"
echo "USER_BASE=$USER_BASE"
echo "VENV=$VENV"

say_dialog() { /usr/bin/osascript -e 'display dialog "'$1'" buttons {"OK"} default button 1 with icon note'; }
warn_dialog(){ /usr/bin/osascript -e 'display dialog "'$1'" buttons {"OK"} default button 1 with icon caution'; }

# 1) Sélection d'un python3 (python.org > Homebrew > PATH)
pick_bootstrap_python() {
  local CANDIDATES=(
    "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3"
    "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3"
    "/opt/homebrew/bin/python3"
    "/usr/local/bin/python3"
  )
  for p in "${CANDIDATES[@]}"; do
    if [ -x "$p" ]; then echo "$p"; return 0; fi
  done
  if command -v python3 >/dev/null 2>&1; then command -v python3; return 0; fi
  return 1
}

PY=""
if ! PY="$(pick_bootstrap_python)"; then
  warn_dialog "Python 3 est introuvable sur ce Mac.\nInstalle Python depuis python.org (3.12 recommandé) puis relance l’installeur."
  echo "ERREUR: python3 introuvable"
  exit 1
fi

echo "Bootstrap Python: $PY"
"$PY" -V || true

# 2) Créer / mettre à jour la venv
if [ -d "$VENV" ]; then
  say_dialog "Une installation existe déjà et va être mise à jour." || true
fi
"$PY" -m venv "$VENV" || {
  warn_dialog "Échec de création de la venv dans:\n$VENV\n\nConsulte le log pour le détail: $LOGFILE"
  echo "ERREUR: venv creation failed"
  exit 1
}

"$VENV/bin/python" -m pip install --upgrade pip setuptools wheel
"$VENV/bin/python" -c 'import sys; print("Python:", sys.version)'
"$VENV/bin/pip" --version

# 3) Vérifier Tkinter sur CE Python
echo "Check Tkinter…"
if ! "$VENV/bin/python" - <<'PY' >/dev/null 2>&1
import tkinter as tk  # noqa
PY
then
  warn_dialog "Ce Python ne dispose pas de Tkinter (interface graphique).\nInstalle Python 3 depuis python.org (inclut Tk), puis relance l’installeur."
  echo "ERREUR: Tkinter absent"
  exit 1
fi
echo "Tkinter OK"

# 4) Installer les dépendances
REQ_SOURCE_DIR="$(cd "$(dirname "$0")" && pwd)"
REQ="$REQ_SOURCE_DIR/requirements.txt"
echo "REQ_SOURCE_DIR=$REQ_SOURCE_DIR"
echo "REQ=$REQ"
if [ ! -f "$REQ" ]; then
  warn_dialog "requirements.txt introuvable à côté de l’installeur.\nPlace ce fichier à côté de PDF-Renommage-installer.command."
  echo "ERREUR: requirements.txt manquant"
  exit 1
fi

echo "Installation des dépendances…"
PIP_OPTS=(--disable-pip-version-check)
[ -n "${PIP_INDEX_URL:-}" ] && PIP_OPTS+=("--index-url" "$PIP_INDEX_URL")
[ -n "${PIP_EXTRA_INDEX_URL:-}" ] && PIP_OPTS+=("--extra-index-url" "$PIP_EXTRA_INDEX_URL")
"$VENV/bin/pip" install -r "$REQ" "${PIP_OPTS[@]}"
echo "Dépendances installées."

# 5) Sanity import (modules critiques)
echo "Sanity import de base…"
if ! "$VENV/bin/python" - <<'PY' >/dev/null 2>&1
import pdfminer.high_level, pypdf, rapidfuzz, docx2pdf, pandas, openpyxl  # noqa
PY
then
  warn_dialog "Échec des imports après installation (voir le log):\n$LOGFILE"
  echo "ERREUR: imports KO"
  exit 1
fi
echo "OK: imports de base validés"

# 6) Vérifier docx2pdf et installer si besoin
DOCX2PDF_OK=0
if "$VENV/bin/python" - <<'PY' >/dev/null 2>&1
import docx2pdf  # noqa
PY
then
  DOCX2PDF_OK=1
else
  echo "docx2pdf absent: tentative d’installation ciblée…"
  if "$VENV/bin/pip" install docx2pdf "${PIP_OPTS[@]}"; then
    if "$VENV/bin/python" - <<'PY' >/dev/null 2>&1
import docx2pdf  # noqa
PY
    then
      DOCX2PDF_OK=1
    fi
  fi
fi

if [ "$DOCX2PDF_OK" -eq 1 ]; then
  echo "docx2pdf OK"
else
  warn_dialog "docx2pdf n’a pas pu être installé.\nLa conversion Word→PDF pourra échouer si aucune autre méthode n’est disponible (LibreOffice ou Microsoft Word).\nTu peux relancer l’installeur après avoir installé LibreOffice ou Word."
  echo "AVERTISSEMENT: docx2pdf indisponible"
fi

# 7) Détection des moteurs disponibles (information utilisateur)
WORD_PRESENT=0
if /usr/bin/osascript -e 'id of application "Microsoft Word"' >/dev/null 2>&1; then
  WORD_PRESENT=1
fi
LIBRE_PRESENT=0
if [ -x "/Applications/LibreOffice.app/Contents/MacOS/soffice" ] || command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1; then
  LIBRE_PRESENT=1
fi

echo "WORD_PRESENT=$WORD_PRESENT"
echo "LIBRE_PRESENT=$LIBRE_PRESENT"

if [ "$DOCX2PDF_OK" -eq 0 ] && [ "$WORD_PRESENT" -eq 0 ] && [ "$LIBRE_PRESENT" -eq 0 ]; then
  warn_dialog "Aucun moteur de conversion Word→PDF détecté.\nInstalle Microsoft Word (recommandé) ou LibreOffice, puis relance l’app."
fi

say_dialog "Installation terminée.\nLa venv est prête dans:\n$VENV\n\nMoteurs détectés:\n- docx2pdf: $([ "$DOCX2PDF_OK" -eq 1 ] && echo OK || echo Non)\n- Microsoft Word: $([ "$WORD_PRESENT" -eq 1 ] && echo Oui || echo Non)\n- LibreOffice: $([ "$LIBRE_PRESENT" -eq 1 ] && echo Oui || echo Non)\n\nTu peux lancer l’application maintenant."
echo "=== $(date) : Installation terminée ==="
