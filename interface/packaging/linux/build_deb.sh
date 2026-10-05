#!/usr/bin/env bash
# Construit pubmed-search-gui_<version>_amd64.deb (à lancer depuis la racine du projet)
set -euo pipefail
cd "$(dirname "$0")/../.."
VERSION=$(python3 packaging/version.py)
PKG=pubmed-search-gui
ARCH=amd64
# Python « portable » (python-build-standalone, compilé pour glibc 2.17) : le paquet ne
# dépend pas de la version de glibc de la machine de construction.
PYBIN=${PYBIN:-}
if [ -z "$PYBIN" ]; then
  python3 -m pip install -q uv
  python3 -m uv python install 3.11 >/dev/null
  PYBIN=$(python3 -m uv python find 3.11 --managed-python 2>/dev/null || python3 -m uv python find 3.11)
fi
rm -rf .venv-build && "$PYBIN" -m venv .venv-build
.venv-build/bin/pip install -q -r requirements-gui.txt
# version pure Python (plus légère) de charset-normalizer, dépendance de requests
.venv-build/bin/pip install -q --force-reinstall --no-deps --no-binary charset-normalizer charset-normalizer
rm -rf build dist
.venv-build/bin/python -m PyInstaller --noconfirm pubmed_gui.spec
GLIBC_MIN=$(python3 packaging/linux/glibc_min.py "dist/$PKG")
echo "glibc minimale requise : $GLIBC_MIN"

ROOT=build/deb/${PKG}_${VERSION}_${ARCH}
rm -rf "$ROOT"
mkdir -p "$ROOT/DEBIAN" "$ROOT/opt" "$ROOT/usr/bin" "$ROOT/usr/share/applications" \
         "$ROOT/usr/share/icons/hicolor/512x512/apps" "$ROOT/usr/share/doc/$PKG"
cp -r "dist/$PKG" "$ROOT/opt/$PKG"
ln -s "/opt/$PKG/$PKG" "$ROOT/usr/bin/$PKG"
cp resources/icon.png "$ROOT/usr/share/icons/hicolor/512x512/apps/$PKG.png"
cp resources/fonts/OFL-*.txt "$ROOT/usr/share/doc/$PKG/"
cat > "$ROOT/usr/share/applications/$PKG.desktop" <<DESK
[Desktop Entry]
Type=Application
Name=PubMed Search
GenericName=Recherche PubMed et PDF
Comment=Rechercher dans PubMed et télécharger les PDF en libre accès
Exec=$PKG %u
MimeType=x-scheme-handler/pubmedsearch;
Icon=$PKG
Terminal=false
Categories=Science;Education;Office;
Keywords=PubMed;PMC;PDF;bibliographie;
StartupWMClass=$PKG
DESK
SIZE=$(du -sk "$ROOT" | cut -f1)
cat > "$ROOT/DEBIAN/control" <<CTRL
Package: $PKG
Version: $VERSION
Section: science
Priority: optional
Architecture: $ARCH
Installed-Size: $SIZE
Depends: libc6 (>= $GLIBC_MIN), libstdc++6, libgcc-s1 | libgcc1, zlib1g, libglib2.0-0, libfreetype6, libfontconfig1, libdbus-1-3, libegl1, libgl1, libx11-6, libx11-xcb1, libxkbcommon0, libxkbcommon-x11-0, libxcb1, libxcb-cursor0, libxcb-icccm4, libxcb-image0, libxcb-keysyms1, libxcb-randr0, libxcb-render0, libxcb-render-util0, libxcb-shape0, libxcb-shm0, libxcb-sync1, libxcb-util1, libxcb-xfixes0, libxcb-xinerama0, libxcb-xkb1
Maintainer: Serge Sawadogo <sergesawadogo@gmail.com>
Description: Interface graphique pour pubmed_search.py
 Recherche PubMed (tri pertinence / plus récents), téléchargement des PDF
 en libre accès (PMC S3, Unpaywall, Europe PMC, HAL, CORE, éditeurs),
 export Excel et import manuel des PDF. Python et ses bibliothèques sont
 intégrés ; le script pubmed_search.py reste un fichier externe modifiable.
CTRL
cat > "$ROOT/DEBIAN/postinst" <<'POST'
#!/bin/sh
set -e
command -v update-desktop-database >/dev/null 2>&1 && update-desktop-database -q || true
command -v gtk-update-icon-cache >/dev/null 2>&1 && gtk-update-icon-cache -q /usr/share/icons/hicolor || true
exit 0
POST
chmod 755 "$ROOT/DEBIAN/postinst"
mkdir -p dist
dpkg-deb --root-owner-group -Zxz -z9 -Sextreme --build "$ROOT" "dist/${PKG}_${VERSION}_${ARCH}.deb"
echo "Paquet : dist/${PKG}_${VERSION}_${ARCH}.deb"
