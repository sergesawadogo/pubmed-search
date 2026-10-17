## PubMed Search 1.5.0 (script pubmed_search.py 1.8.0)

**Nouveautés de la file d'attente**
- **Journal par recherche** : un clic sur une recherche de la file affiche son journal, ses
  compteurs et ses boutons (dossier, Excel, error.txt) dans le panneau de droite ; un bandeau
  rappelle quelle recherche est affichée et ramène au journal en direct. Les journaux sont
  conservés entre deux sessions.
- **Importer une liste de recherches (Excel)** : *Créer le modèle Excel…* produit un fichier à
  remplir (une ligne par recherche, une colonne à cocher par filtre, y compris vos filtres
  personnalisés, listes déroulantes pour le tri et les options) ; *Importer une liste
  (Excel)…* ajoute toutes les recherches à la file et signale les lignes à corriger.

**Rappel 1.4.0** : identification d'un dossier de PDF quelconque (`--identify-pdf`), copies
renommées, Excel des références et bibliographie Word au style Vancouver.

**Fichiers**
- `PubMedSearch-Setup-1.5.0.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.5.0_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.5.0_amd64.deb`).
- `PubMedSearch-1.5.0-macOS-AppleSilicon.dmg` et `PubMedSearch-1.5.0-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés). Construits par GitHub Actions ; ils apparaissent ici une
  quinzaine de minutes après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.

Mise à jour : installez la nouvelle version par-dessus l'ancienne ; réglages, clés et file
d'attente sont conservés.
