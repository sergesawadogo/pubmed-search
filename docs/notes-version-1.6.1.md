## PubMed Search 1.6.1 (script pubmed_search.py 1.9.0)

**Correction** : la 1.6.0 affichait « Erreur inattendue : Stat.show() missing 1 required positional
argument » quand une recherche avec texte intégral JSON (par exemple importée d'une liste Excel)
atteignait sa passe JSON. Corrigé, ainsi qu'un message d'erreur possible à la fermeture.

**Nouveauté : évaluer une recherche avant de la lancer**
- Bouton **Évaluer** (onglet Recherche) : sans rien télécharger, PubMed indique le nombre de
  résultats, de gratuits (free full text), dont dans PMC, et de payants, au total et parmi les
  articles qui seront traités (maximum et tri choisis). Termes non reconnus, requête telle que
  PubMed l'a comprise et conseils d'affinage sont affichés. Ensuite : **Lancer la recherche**,
  **Ajouter à la file** ou **Affiner la requête**.
- **Évaluer la file** (onglet File d'attente) : même évaluation pour chaque recherche en attente,
  colonne *Évaluation* dans le tableau, et Excel d'évaluation enregistré dans le dossier du
  fichier Excel de la liste importée.
- Script : option `--evaluate` (ligne finale `EVAL_JSON` lisible par d'autres programmes).

**Rappel 1.5.0** : journal par recherche de la file, import d'une liste de recherches depuis Excel.

**Fichiers**
- `PubMedSearch-Setup-1.6.1.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.6.1_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.6.1_amd64.deb`).
- `PubMedSearch-1.6.1-macOS-AppleSilicon.dmg` et `PubMedSearch-1.6.1-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés), construits par GitHub Actions une quinzaine de minutes
  après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.

Mise à jour : installez la nouvelle version par-dessus l'ancienne ; réglages, clés et file
d'attente sont conservés.
