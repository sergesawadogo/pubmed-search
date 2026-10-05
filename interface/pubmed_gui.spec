# -*- mode: python ; coding: utf-8 -*-
# Construction : pyinstaller pubmed_gui.spec            (dossier, pour .deb / installateur)
#                PUBMED_ONEFILE=1 pyinstaller pubmed_gui.spec   (un seul .exe portable)
import os, sys
ONEFILE = os.environ.get("PUBMED_ONEFILE") == "1"
WIN = sys.platform.startswith("win")
NAME = "PubMedSearch" if WIN else "pubmed-search-gui"

a = Analysis(
    ["pubmed_gui.py"],
    pathex=["."],
    datas=[("resources", "resources")],
    hiddenimports=["runtime_deps"],
    excludes=["tkinter", "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.Qt3DCore",
              "PySide6.QtQuick", "PySide6.QtQml", "PySide6.QtMultimedia", "PySide6.QtCharts",
              "PySide6.QtPdf", "PySide6.QtNetwork", "matplotlib", "numpy", "PIL", "lxml",
              "cryptography"],
    noarchive=False,
)
# Allège le paquet : bibliothèques Qt inutiles à l'interface (Quick/QML/PDF tirées par
# certains greffons) et traductions autres que le français.
DROP = ("Qt6Quick", "Qt6Qml", "Qt6Pdf", "qpdf", "Qt6VirtualKeyboard")  # garder Qt6Svg : icônes SVG
def keep(entry):
    dest = entry[0].replace("\\", "/")
    if any(d in dest for d in DROP):
        return False
    if "/translations/" in dest:   # l'interface ne charge pas de traductions Qt
        return False
    return True
a.binaries = [b for b in a.binaries if keep(b)]
if sys.platform.startswith("linux"):
    # Greffons Qt inutiles sur un poste de bureau X11/XWayland (paquet < 30 Mo)
    LINUX_DROP = ("qt6wayland", "/wayland-", "qwayland", "qt6eglfs", "/egldeviceintegrations/",
                  "/generic/", "qeglfs", "qlinuxfb", "qvnc", "qvkkhrdisplay", "qminimal",
                  "qt6network", "qtga", "qtiff", "qwbmp", "qicns", "qwebp", "qgif", "qjpeg",
                  "qt6wlshell", "qgtk3", "qxcb-egl-integration", "/iconengines/")
    a.binaries = [b for b in a.binaries
                  if not any(d in b[0].replace("\\", "/").lower() for d in LINUX_DROP)]
if sys.platform.startswith("linux"):
    # Ne pas embarquer les bibliothèques système de la machine de construction (X11, GTK,
    # fontconfig, glib…) : elles exigeraient sa version de glibc. Le .deb les déclare en
    # dépendances et utilise celles du système cible. On garde Python et les roues PyPI.
    SYS = ("/usr/lib/", "/lib/", "/lib64/", "/usr/lib64/")
    a.binaries = [b for b in a.binaries if not str(b[1]).startswith(SYS)]
a.datas = [d for d in a.datas if keep(d)]
if not WIN:
    a.datas = [d for d in a.datas if not d[0].replace("\\", "/").endswith(("icon.ico", "IBMPlexMono-500.ttf"))]
pyz = PYZ(a.pure)
icon = "resources/icon.ico" if WIN else None
if ONEFILE:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=NAME, console=False, icon=icon,
              upx=False, runtime_tmpdir=None)
else:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=NAME, console=False, icon=icon, upx=False,
              strip=not WIN)
    coll = COLLECT(exe, a.binaries, a.datas, name=NAME, upx=False, strip=not WIN)
