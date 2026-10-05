"""Affiche APP_VERSION de pubmed_gui.py (utilisé par les scripts de construction)."""
import os, re
src = open(os.path.join(os.path.dirname(__file__), "..", "pubmed_gui.py"), encoding="utf-8").read()
print(re.search(r'APP_VERSION = "([^"]+)"', src).group(1))
