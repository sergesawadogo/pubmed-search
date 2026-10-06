## PubMed Search 1.3.0 (script pubmed_search.py 1.7.2)

**Nouveautés**
- **Passe JSON** (case « Télécharger les JSON », option `--getjson`) : pour les articles restés
  sans PDF, récupération du texte intégral structuré depuis NCBI BioC (PMC Open Access et
  manuscrits d'auteurs), Europe PMC, Springer Nature Open Access API et API Elsevier ; résumé
  via Springer Nature Meta API en dernier recours. Le PDF reste toujours prioritaire. Les JSON
  (schéma commun : titre, résumé, sections, figures, tableaux, références, licence) sont rangés
  dans `JSON_<nom>/` et listés dans l'Excel (3 nouvelles colonnes) et dans error.txt.
- **Springer Nature** : clés Open Access API et Meta API dans Réglages. La Meta API fournit
  aussi le lien PDF officiel des articles Springer/BMC/Nature en libre accès.
- **Sources utilisées** (menu déroulant à cases, option `--sources`) : Elsevier, Springer Nature
  Open Access, Springer Nature Meta, Unpaywall, CORE, OpenAlex — toutes cochées par défaut.
- **Filtre Burkina Faso** : appellations politiques, centres de recherche, hôpitaux, régions
  sanitaires et chefs-lieux (y compris les 17 régions renommées en 2025), et case **ALL** qui
  coche tout le volet. Ajouté automatiquement à vos filtres existants.
- **Éditeur de filtres intégré** (Réglages → Modifier les filtres…) : volets et cases dans une
  liste, requête colorée, vérification des parenthèses et guillemets, application immédiate
  sans redémarrage.
- **Journal** : PDF obtenus en vert, sans PDF en magenta, JSON en cyan, compteur « JSON
  obtenus » ; fond de la zone Journal au choix (Réglages → Apparence du journal).

- **Elsevier** : jeton institutionnel facultatif (Réglages, ou `--elsevier-insttoken`) ; message
  clair et conseil dans error.txt quand la clé est valide mais sans droits sur le texte intégral.
- **error.txt** : section « CONSEILS » (clé manquante, articles PMC hors Open Access à importer
  à la main…), reprise en fin de journal.

**Limites connues**
- API Elsevier : sans réseau d'une institution abonnée ni jeton institutionnel, le texte intégral
  est refusé, y compris pour certains articles gratuits (Lancet).
- Karger ne propose pas d'API publique de texte intégral (accès TDM sur contrat, par FTP) :
  ses articles libres restent récupérés par les voies habituelles (PMC, Unpaywall, éditeur).
- Forfait gratuit Springer Nature : quota journalier ; en cas de dépassement (HTTP 429) ou de
  clé refusée, l'API concernée est désactivée pour le reste de la recherche.

**Fichiers**
- `PubMedSearch-Setup-1.3.0.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.3.0_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.3.0_amd64.deb`).
- `PubMedSearch-1.3.0-macOS-AppleSilicon.dmg` et `PubMedSearch-1.3.0-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés). Construits par GitHub Actions ; ils apparaissent ici une
  quinzaine de minutes après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.

Mise à jour : installez simplement la nouvelle version par-dessus l'ancienne ; vos réglages,
clés mémorisées, file d'attente et filtres personnalisés sont conservés.
