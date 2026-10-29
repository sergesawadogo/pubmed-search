## PubMed Search 1.9.0 (script pubmed_search.py 1.10.0)

**Préparer pour une IA locale (XML JATS)** : nouvelle case sous « Préparer aussi pour Gemini
Notebook ». Pour chaque article présent dans PMC, un dossier `JATS_<nom>/PMID-Auteur(Année)/` :
- le XML JATS d'origine (PMC Cloud Service, sinon Europe PMC), complet et sans perte ;
- un Markdown lisible par une IA : texte, **tableaux complets** (HTML si cellules fusionnées),
  légendes et liens vers les figures, formules, références ;
- un CSV par tableau et les images des figures ; les suppléments restent en lien (volume).
Excel : deux colonnes (lien vers le Markdown, contenu) ; error.txt : articles sans XML et raison.

**Mots-clés et termes MeSH équivalents** : chaque recherche crée `<nom>_mots-cles.txt` et un
onglet « Mots-clés » dans l'Excel, une ligne par groupe d'équivalents
(ex. `Extracellular Traps = NETosis = NETs = Neutrophil Extracellular Traps`). Sources : MeSH et
synonymes officiels, mots-clés d'auteurs rattachés par PubMed, sigles définis dans les résumés.
File d'attente → **Mots-clés des requêtes** : même fichier pour les termes d'une liste de
recherches, enregistré à côté de la liste Excel.

**File d'attente** : les recherches importées d'une liste Excel reprennent les options « Gemini
Notebook » et « IA locale » cochées dans l'onglet Recherche (auparavant, la conversion Gemini
n'était jamais faite pour elles) ; « Modifier dans Recherche » restaure aussi ces deux options.

**Fichiers**
- `PubMedSearch-Setup-1.9.0.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.9.0_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.9.0_amd64.deb`).
- `PubMedSearch-1.9.0-macOS-AppleSilicon.dmg` et `PubMedSearch-1.9.0-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés), construits par GitHub Actions après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.
