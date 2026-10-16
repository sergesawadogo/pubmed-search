## PubMed Search 1.4.0 (script pubmed_search.py 1.8.0)

**Nouveauté : identifier un dossier de PDF et créer sa bibliographie**
- Onglet **Import manuel** → *Identifier un dossier de PDF (sans recherche)*, ou en ligne de
  commande `--identify-pdf DOSSIER --output Nom`.
- Chaque PDF est retrouvé dans PubMed par le PMID ou le DOI présent dans son nom, ses
  propriétés ou sa première page, sinon par son titre ; hors PubMed, par Crossref. Le titre
  trouvé est comparé au texte du PDF (protection contre le DOI d'une référence citée).
- Résultats dans un nouveau dossier, les originaux restant intacts : copies renommées
  `PMID-Auteur_Initiales(Année).pdf` (`NoPMID-…` hors PubMed), Excel des références avec
  liens PubMed, DOI et PDF, bibliographie Word au style Vancouver (numérotée, ordre
  alphabétique), liste des PDF à vérifier. Les identifications incertaines ne sont jamais
  renommées : elles sont copiées telles quelles dans `A_verifier/`.

**Rappel 1.3.1** : copie bloquée détectée au lancement, journal `demarrage.log`, ID converter
PMC remplacé par E-utilities en cas de refus, CORE limité à 5 min (`--core-budget`).

**Fichiers**
- `PubMedSearch-Setup-1.4.0.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.4.0_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.4.0_amd64.deb`).
- `PubMedSearch-1.4.0-macOS-AppleSilicon.dmg` et `PubMedSearch-1.4.0-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés). Construits par GitHub Actions ; ils apparaissent ici une
  quinzaine de minutes après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.

Mise à jour : installez la nouvelle version par-dessus l'ancienne ; réglages, clés et file
d'attente sont conservés.
