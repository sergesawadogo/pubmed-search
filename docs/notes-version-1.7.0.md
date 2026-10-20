## PubMed Search 1.7.0 (script pubmed_search.py 1.9.1)

**Onglet PDF (ex-« Import manuel ») réorganisé en deux sous-onglets bien distincts**
- **Importer dans une recherche** (« Importer les PDF d'une recherche PubMed Search ») : pour
  compléter une recherche avec les PDF téléchargés à la main ; ils sont ajoutés à son dossier
  PDF_ et à son Excel.
- **Renommer des PDF** (« Renommer des PDF au format PubMed Search ») : pour des PDF venus
  d'ailleurs ; nouveau dossier avec copies renommées, Excel et bibliographie Word.
- Chaque page commence par un encadré « À utiliser quand… ». Garde-fous : choisir le dossier
  d'une recherche dans *Renommer* propose d'aller à *Importer* ; un dossier de PDF qui ressemble
  à un dossier de résultats est signalé dans *Importer*.

**Nouvel onglet Partager**
- Archive zip du dossier d'une recherche, contenu au choix (Excel, PDF, JSON, error.txt et
  listes de PMID, bibliographie Word) avec la taille affichée.
- Message prérempli (requête, nombre d'articles et de PDF) ouvert dans **Thunderbird** ou
  **Outlook** avec la pièce jointe ; sinon, archive prête et dossier ouvert.
- Au-delà de la limite par mail (15 Mo par défaut) : **plusieurs archives autonomes** (un mail
  par archive, chacune s'ouvre seule) ou **une archive + lien SwissTransfer**.
- Archives créées en arrière-plan (l'interface reste utilisable), avec contrôle de l'espace disque.

**Fichiers**
- `PubMedSearch-Setup-1.7.0.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.7.0_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.7.0_amd64.deb`).
- `PubMedSearch-1.7.0-macOS-AppleSilicon.dmg` et `PubMedSearch-1.7.0-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés), construits par GitHub Actions une quinzaine de minutes
  après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.
