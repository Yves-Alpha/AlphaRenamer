#!/usr/bin/env bash
set -euo pipefail

# Genere un bundle portable de PDF-Renommage (v1.0.4) avec venv embarquee,
# puis cree une archive .zip dans dist/.
# Usage:
#   ./build_portable.sh          # construit/rafraichit le venv et produit l'archive
#   ./build_portable.sh --clean  # recree le venv avant de packager

VERSION="1.0.4"
APP_NAME="PDF-Renommage"
ROOT="$(cd "$(dirname "$0")" && pwd)"
APP_BUNDLE="$ROOT/${APP_NAME}.app"
RES_DIR="$APP_BUNDLE/Contents/Resources"
APP_DIR="$RES_DIR/app"
SRC_MAIN="$ROOT/app/alpha_renamer_gui.py"
REQ_SRC="$ROOT/Installer/requirements.txt"
REQ_DST="$APP_DIR/requirements.txt"
VENV_DIR="$APP_DIR/venv"
DIST_DIR="$ROOT/dist"
ARCHIVE="$DIST_DIR/${APP_NAME}-portable-${VERSION}.zip"

pick_python() {
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

if [ ! -d "$APP_BUNDLE" ]; then
  echo "ERROR: Bundle introuvable: $APP_BUNDLE" >&2
  exit 1
fi
if [ ! -f "$SRC_MAIN" ]; then
  echo "ERROR: Script principal introuvable: $SRC_MAIN" >&2
  exit 1
fi
if [ ! -f "$REQ_SRC" ]; then
  echo "ERROR: requirements.txt introuvable: $REQ_SRC" >&2
  exit 1
fi

PY=""
if ! PY="$(pick_python)"; then
  echo "ERROR: python3 introuvable (installe Python 3.12 depuis python.org)" >&2
  exit 1
fi
echo "Python utilise: $PY"
"$PY" -V || true

CLEAN=0
if [ "${1:-}" = "--clean" ]; then
  CLEAN=1
fi

echo "== Sync des sources dans le bundle =="
rsync -a --delete --exclude "venv" "$ROOT/app/" "$APP_DIR/"
cp "$REQ_SRC" "$REQ_DST"

if [ "$CLEAN" -eq 1 ] && [ -d "$VENV_DIR" ]; then
  echo "Suppression du venv portable existant ($VENV_DIR)..."
  rm -rf "$VENV_DIR"
fi

if [ ! -d "$VENV_DIR" ]; then
  echo "Creation du venv portable..."
  "$PY" -m venv "$VENV_DIR"
fi

echo "Mise a jour pip/setuptools/wheel..."
"$VENV_DIR/bin/python" -m pip install --upgrade pip setuptools wheel

echo "Installation des dependances..."
PIP_OPTS=(--disable-pip-version-check)
[ -n "${PIP_INDEX_URL:-}" ] && PIP_OPTS+=("--index-url" "$PIP_INDEX_URL")
[ -n "${PIP_EXTRA_INDEX_URL:-}" ] && PIP_OPTS+=("--extra-index-url" "$PIP_EXTRA_INDEX_URL")
"$VENV_DIR/bin/pip" install -r "$REQ_SRC" "${PIP_OPTS[@]}"

echo "Sanity import..."
"$VENV_DIR/bin/python" - <<'PY'
import tkinter as tk  # noqa: F401
import pdfminer.high_level, pypdf, rapidfuzz, docx2pdf, pandas, openpyxl  # noqa: F401
print("OK: imports de base valides")
PY

echo "Nettoyage des __pycache__..."
find "$VENV_DIR" -name "__pycache__" -type d -prune -exec rm -rf {} +

mkdir -p "$DIST_DIR"
if [ -f "$ARCHIVE" ]; then
  rm -f "$ARCHIVE"
fi

echo "Creation de l'archive portable..."
(cd "$ROOT" && /usr/bin/ditto -c -k --sequesterRsrc --keepParent "$APP_NAME.app" "$ARCHIVE")

echo "Portable pret : $ARCHIVE"
