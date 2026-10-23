## PubMed Search 1.8.1 (script pubmed_search.py 1.9.1)

**Gemini Notebook : conversion des recherches déjà faites, en une fois**

Onglet PDF → « Pour Gemini Notebook », le champ accepte désormais :
- le dossier d'une recherche (comme avant) ;
- son dossier `JSON_…` ou `PDF_…` choisi directement (auparavant : rien n'était trouvé) ;
- un dossier de JSON isolés, déplacés hors de leur recherche ;
- un dossier contenant **plusieurs recherches** (exploré sur 3 niveaux) : un dossier
  `GeminiNotebook_…` est créé dans chaque recherche qui a des JSON. Case « Ignorer les recherches
  déjà converties » (cochée par défaut). La liste des recherches trouvées s'affiche au survol.

**Correction** : avec l'option automatique et une file d'attente, si une recherche se terminait
pendant la conversion de la précédente, sa conversion était silencieusement abandonnée. Elle est
maintenant mise en attente puis traitée.

Les JSON d'origine ne sont jamais modifiés. Les auteurs viennent de l'Excel de la recherche
(cherché dans le dossier puis dans le dossier parent) ; sans Excel, les PDF sont créés sans eux.

---

### Rappel : nouveautés de la 1.8.0 (incluses ici, non publiée séparément)

**Nouvelle identité et interface modernisée**
- Nouveau logo (anticorps, loupe, couleurs du Burkina Faso) pour l'application, l'en-tête et
  les installateurs ; version simplifiée aux petites tailles (16-32 px).
- Couleurs tirées du logo, avec un sens fixe : vert = actions et PDF obtenus, ambre =
  avertissements et articles sans PDF, rouge = erreurs. Contrastes du texte vérifiés (≥ 4,5:1).
- **Apparence sombre** : l'application suit le réglage clair/sombre du système.
- Anneau de focus visible sur tous les contrôles (navigation au clavier), coins arrondis harmonisés.
- **Barre de menus** (Fichier, Recherche, Affichage, Aide) avec raccourcis : Ctrl/⌘+E évaluer,
  Ctrl/⌘+Entrée lancer, Ctrl/⌘+Maj+Entrée ajouter à la file, Ctrl/⌘+. arrêter, Ctrl/⌘+1…4 onglets,
  Ctrl/⌘+, réglages ; sur Mac, Réglages et À propos rejoignent le menu de l'application.

**File d'attente : « Modifier dans Recherche »** (remplace « Reprendre dans Recherche ») : la
recherche reste à sa place dans la file ; un bandeau signale la modification en cours ;
« Enregistrer dans la file » la met à jour. Elle n'est pas lancée tant que la modification est
ouverte ; impossible pendant son exécution.

**Gemini Notebook (ex-NotebookLM)** : onglet PDF → « Pour Gemini Notebook » : les JSON de texte
intégral deviennent des PDF à texte sélectionnable, réunis avec les PDF téléchargés dans un dossier
prêt à importer, en lots si nécessaire (50 sources par notebook gratuit). Option automatique en
fin de recherche.

**Fichiers**
- `PubMedSearch-Setup-1.8.1.exe` : Windows 10 et 11, 64 bits (non signé : SmartScreen →
  Informations complémentaires → Exécuter quand même).
- `pubmed-search-gui_1.8.1_amd64.deb` : Ubuntu 20.04+, Debian 10+, Linux Mint 20+
  (`sudo apt install ./pubmed-search-gui_1.8.1_amd64.deb`).
- `PubMedSearch-1.8.1-macOS-AppleSilicon.dmg` et `PubMedSearch-1.8.1-macOS-Intel.dmg` :
  macOS 12+ (non signés ni notariés), construits par GitHub Actions après la publication.
- `envoyer-vers-pubmed-search-1.0.0.xpi` : extension Firefox (inchangée).
- `SHA256SUMS.txt` et `SHA256SUMS-macOS.txt` : empreintes SHA-256.
