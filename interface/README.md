# PubMed Search — interface graphique

Interface de bureau pour `pubmed_search.py`. Elle ne contient pas la logique de recherche :
elle **charge le script choisi** (fichier externe que vous pouvez faire évoluer) et l'exécute
avec le Python intégré à l'application. Rien d'autre à installer pour l'utilisateur.

## Installation

| Système | Fichier | Installation |
|---|---|---|
| Ubuntu / Debian | `pubmed-search-gui_1.2.0_amd64.deb` | `sudo apt install ./pubmed-search-gui_1.2.0_amd64.deb` |
| Windows 10 / 11 (64 bits) | `PubMedSearch-Setup-1.2.0.exe` | double-clic (droits administrateur) |

L'installateur Windows n'est pas signé numériquement : au premier lancement, Windows
SmartScreen affiche « Windows a protégé votre ordinateur » → *Informations complémentaires*
→ *Exécuter quand même*. Désinstallation : Paramètres → Applications → PubMed Search.

Le .deb embarque un Python « portable » (python-build-standalone) et n'emporte pas les
bibliothèques système de la machine de construction : il demande seulement glibc ≥ 2.28
(Ubuntu 20.04+, Debian 10+, Linux Mint 20+). Les bibliothèques X11/Qt nécessaires sont
installées automatiquement par `apt` comme dépendances.

## Utilisation

- **Recherche** : requête PubMed (colorée : champs `[ti]` en rose, opérateurs en violet),
  tri Pertinence / Plus récents / Les deux, nombre maximum d'articles, passes supplémentaires,
  dossier de destination (bouton *Nouveau dossier…* pour le créer où vous voulez),
  nom du fichier Excel. *Options avancées* : programmer le lancement (`--time-set`).
- **Filtres** (volets dépliables sous la requête) : *Régions cibles*, *Niveau de développement*,
  *Type d'articles*. Les cases d'un même volet sont combinées par OR, les volets entre eux et
  avec la requête par AND ; la requête finale envoyée à PubMed s'affiche en dessous.
  Les filtres sont définis dans `resources/filtres.json` (copie modifiable via
  Réglages → *Modifier les filtres…*).
- **File d'attente** : *Ajouter à la file* (onglet Recherche) prépare plusieurs recherches ;
  l'onglet *File d'attente* les exécute l'une après l'autre, tout de suite ou à une date/heure
  programmée (l'application doit rester ouverte). La file est conservée entre deux sessions.
- **Import manuel** : ouvre par lots de 10 les liens « À ouvrir manuellement » d'error.txt,
  puis importe les PDF enregistrés (`--import-pdf`).
- **Réglages** : script utilisé, email, clés API (NCBI, CORE, Elsevier, OpenAlex),
  mémorisation des clés, interpréteur Python.
- **Journal** : sortie du script en direct, compteurs (PDF obtenus, sans PDF, restants,
  taux de réussite), bouton *Arrêter* (le script enregistre ses résultats partiels),
  ouverture du dossier, de l'Excel et d'error.txt.

Les clés API sont passées au script par variables d'environnement (`NCBI_API_KEY`,
`CORE_API_KEY`, `ELSEVIER_API_KEY`, `OPENALEX_API_KEY`) : elles n'apparaissent ni dans le
journal ni dans error.txt. Si vous cochez « Mémoriser », elles sont écrites en clair dans
le fichier de réglages (`~/.config/pubmed-search-gui/settings.json`, ou
`%LOCALAPPDATA%\pubmed-search-gui\settings.json` sous Windows).

## Faire évoluer le script

Réglages → *Choisir…* → votre `pubmed_search.py`. Au chargement, l'interface lit
`--help` du script et désactive les champs dont l'option n'existe pas dans cette version.
Les nouvelles options (autres que celles prévues) ne sont pas affichées : il faut alors
modifier `pubmed_gui.py`.

Si une nouvelle version du script importe une bibliothèque absente du Python intégré
(intégrées : requests, openpyxl, pypdf + bibliothèque standard) : soit choisir
« Autre interpréteur Python » dans Réglages, soit ajouter l'import dans
`runtime_deps.py` puis reconstruire.

## Construire les paquets

**Linux (.deb)** : `./packaging/linux/build_deb.sh` → `dist/pubmed-search-gui_<version>_amd64.deb`

**Windows (.exe) depuis Linux** (méthode utilisée pour la 1.2.0) : Python Windows
(python-build-standalone) exécuté sous Wine pour PyInstaller, puis
`makensis -DVERSION=1.2.0 -DSRCDIR=../../dist-win/PubMedSearch packaging/windows/installer.nsi`
(NSIS, installateur 64 bits pour Windows 10/11).

**Windows (.exe)**, sur un PC Windows avec Python 3.10+ :
double-clic sur `packaging\windows\build_windows.bat` →
`dist\portable\PubMedSearch.exe` et, si [Inno Setup 6](https://jrsoftware.org/isdl.php)
est installé, `dist\PubMedSearch-Setup-<version>.exe`.

**GitHub Actions** : poussez ce dossier dans un dépôt GitHub, puis onglet *Actions* →
« Construire les paquets » → *Run workflow* (ou poussez une étiquette `v1.0.0`).
Les paquets .deb et .exe sont dans les *artifacts* de l'exécution.

## Développement

```
pip install -r requirements-gui.txt
python pubmed_gui.py
```

Polices : Atkinson Hyperlegible Next et IBM Plex Mono (SIL Open Font License, fichiers
dans `resources/fonts`).
