## PubMed Search 1.3.1 (script pubmed_search.py 1.7.3)

Version corrective de la 1.3.0, avec une nouvelle icône (une loupe sur une double hélice d'ADN)
pour reconnaître d'un coup d'œil la version installée.

**Corrections**
- **L'application ne s'ouvrait plus** quand une ancienne copie restait lancée sans fenêtre
  (par exemple après une mise à jour) : chaque nouveau lancement se fermait sans rien dire.
  Désormais, si la copie déjà ouverte ne répond pas dans les 5 secondes, PubMed Search
  propose de la fermer puis s'ouvre normalement.
- Une fenêtre enregistrée hors de l'écran (écran débranché, autre résolution) est recentrée.
- **Journal de démarrage** `demarrage.log` (dossier des réglages) : version, système,
  environnement graphique, erreurs Python et Qt, et pile d'appels si la fenêtre n'apparaît
  pas dans les 45 secondes. Sous Linux, le lanceur y ajoute aussi les messages d'erreur
  quand l'application est lancée depuis le menu ou le dock.
- **Script 1.7.3** : quand le site PMC refuse l'ID converter (HTTP 403), les PMCID sont
  cherchés par E-utilities (ELink) au lieu d'afficher une erreur ; CORE est limité à
  5 minutes par recherche (`--core-budget`, -1 = sans limite) : son quota gratuit pouvait
  occuper l'essentiel d'une recherche pour 1 ou 2 PDF ; le taux de réussite ne compte plus
  les PDF d'articles non signalés gratuits.
- macOS : nouvelle tentative automatique quand la création du .dmg échoue sur les machines
  de GitHub (« Resource busy »), et publication d'une architecture même si l'autre échoue.

**Rappel des nouveautés de la 1.3.0** : passe JSON de texte intégral (`--getjson`), clés
Springer Nature, choix des sources (`--sources`), filtre Burkina Faso, éditeur de filtres
intégré, couleurs du journal, jeton institutionnel Elsevier, section « CONSEILS » d'error.txt.

**Fichiers**
- `PubMedSearch-Setup-1.3.1.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.3.1_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.3.1_amd64.deb`).
- `PubMedSearch-1.3.1-macOS-AppleSilicon.dmg` et `PubMedSearch-1.3.1-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés). Construits par GitHub Actions ; ils apparaissent ici une
  quinzaine de minutes après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.

Mise à jour : installez la nouvelle version par-dessus l'ancienne ; réglages, clés mémorisées,
file d'attente et filtres personnalisés sont conservés.
