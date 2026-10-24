## PubMed Search 1.8.2 (script pubmed_search.py 1.9.2)

**Panne d'un service (ex. Europe PMC en HTTP 503) : la recherche ne s'enlise plus**
- Après 3 erreurs serveur de suite (500, 502, 503, 504 ou pas de réponse), le service est **mis en
  pause** (2 min, puis 5, puis 10) au lieu d'être réinterrogé pour chaque lot et chaque article ;
  un seul essai de reprise à la fin de chaque pause. Les autres sources (PMC S3, Unpaywall,
  OpenAlex, HAL, éditeurs, Elsevier…) continuent normalement. Services concernés : Europe PMC,
  OpenAlex, Unpaywall, HAL, Crossref, NCBI BioC, ID converter PMC (pas les E-utilities NCBI,
  indispensables à la recherche).
- Exemple observé : 873 articles, Europe PMC en 503 → 36 requêtes et environ 3 min sans résultat
  avant 1.9.2 ; désormais 6 requêtes au plus, puis passage à l'étape suivante.
- Messages d'erreur courts (« HTTP 503 (service indisponible) ») au lieu de l'URL complète de
  plusieurs milliers de caractères.
- Progression affichée à chaque étape de l'enrichissement (ID converter, Europe PMC, LinkOut,
  OpenAlex) : l'écran ne semble plus figé.
- Statistiques et error.txt : ligne « Services mis en pause » et conseil de relancer plus tard.

**Correction** : quand une recherche était arrêtée avant le téléchargement, error.txt affichait
des conseils trompeurs (« N articles Elsevier non obtenus malgré la clé ») alors qu'aucun
téléchargement n'avait été tenté. Il indique maintenant que la recherche a été interrompue avant
la passe 1.

**Fichiers**
- `PubMedSearch-Setup-1.8.2.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.8.2_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.8.2_amd64.deb`).
- `PubMedSearch-1.8.2-macOS-AppleSilicon.dmg` et `PubMedSearch-1.8.2-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés), construits par GitHub Actions après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.
