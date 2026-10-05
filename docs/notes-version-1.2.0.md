## PubMed Search 1.2.0

**Nouveautés**
- Extension Firefox « Envoyer vers PubMed Search » : bouton sur pubmed.ncbi.nlm.nih.gov,
  filtres de PubMed convertis en requête, menu contextuel sur une sélection.
- L'application reçoit les liens `pubmedsearch://` et ne s'ouvre qu'une fois (une demande
  envoyée alors qu'elle est ouverte arrive dans la fenêtre existante).
- Filtres Régions cibles / Niveau de développement / Type d'articles, file d'attente
  programmable (depuis la 1.1.0).

**Fichiers**
- `PubMedSearch-Setup-1.2.0.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.2.0_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.2.0_amd64.deb`).
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox non signée (installation
  temporaire via about:debugging, ou permanente sur Developer Edition / Nightly / ESR).
- `SHA256SUMS.txt` : empreintes SHA-256.

Notice complète : voir le README du dépôt.
