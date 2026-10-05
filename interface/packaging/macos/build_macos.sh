#!/usr/bin/env bash
# Construit « PubMed Search.app » et l'image disque .dmg pour l'architecture du Mac courant :
#   Mac Apple Silicon (arm64)  → dist/PubMedSearch-<version>-macOS-AppleSilicon.dmg
#   Mac Intel (x86_64)         → dist/PubMedSearch-<version>-macOS-Intel.dmg
# PyInstaller ne sait pas construire pour macOS depuis Linux ou Windows : ce script tourne
# sur un Mac (ou sur les machines macOS de GitHub Actions, voir .github/workflows/macos.yml).
# Prérequis : Python 3.10+ (python.org ou Homebrew). Utilisation : ./packaging/macos/build_macos.sh
set -euo pipefail
set -x   # affiche chaque commande dans le journal (diagnostic)
cd "$(dirname "$0")/../.."          # dossier interface/

PY="${PYBIN:-python3}"
VERSION=$("$PY" packaging/version.py)
case "$(uname -m)" in
  arm64)  ARCH=AppleSilicon ;;
  x86_64) ARCH=Intel ;;
  *) echo "Architecture non prise en charge : $(uname -m)"; exit 1 ;;
esac
echo "PubMed Search $VERSION — macOS $ARCH"

"$PY" -m venv .venv-build
# shellcheck disable=SC1091
. .venv-build/bin/activate
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements-gui.txt
cp -f ../script/pubmed_search.py resources/pubmed_search.py 2>/dev/null || true

rm -rf build dist
pyinstaller --noconfirm --clean pubmed_gui.spec

APP="dist/PubMed Search.app"
[ -d "$APP" ] || { echo "Application absente : $APP"; exit 1; }
# PyInstaller signe l'application « ad hoc » (sans certificat Apple), ce qu'exige Apple
# Silicon pour exécuter le code ; Gatekeeper demandera tout de même une confirmation au
# premier lancement. On vérifie la signature et on la refait si besoin.
codesign --verify --deep --strict "$APP" 2>/dev/null || codesign --force --deep --sign - "$APP"

echo "Vérification : le Python intégré exécute le script…"
"$APP/Contents/MacOS/PubMedSearch" --run-script resources/pubmed_search.py --help | grep -q -- "--output" \
  || { echo "Échec de la vérification du script intégré"; exit 1; }
lipo -archs "$APP/Contents/MacOS/PubMedSearch"

DMG="dist/PubMedSearch-$VERSION-macOS-$ARCH.dmg"
STAGE=$(mktemp -d)
cp -R "$APP" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
# hdiutil échoue parfois (« Resource busy ») sur les machines de GitHub : jusqu'à 5 essais.
ok=0
for i in 1 2 3 4 5; do
  if hdiutil create -volname "PubMed Search $VERSION" -srcfolder "$STAGE" -ov -format UDZO "$DMG"; then
    ok=1; break
  fi
  echo "hdiutil : essai $i échoué, nouvel essai dans 10 s…"; sleep 10
done
[ "$ok" = 1 ] || { echo "Impossible de créer l'image disque"; exit 1; }
rm -rf "$STAGE"
echo "Terminé : $DMG ($(du -h "$DMG" | cut -f1))"
