"""Version de l'application (lue dans pubmed_gui.py) ; exécuté seul, l'affiche."""
import os
import re
_src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "pubmed_gui.py"),
            encoding="utf-8").read()
VERSION = re.search(r'APP_VERSION = "([^"]+)"', _src).group(1)
if __name__ == "__main__":
    print(VERSION)
