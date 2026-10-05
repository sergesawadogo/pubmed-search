# Extension Firefox « Envoyer vers PubMed Search »

Envoie la recherche affichée sur pubmed.ncbi.nlm.nih.gov (requête + filtres de la barre
latérale convertis en syntaxe PubMed + tri) à l'application de bureau PubMed Search,
via des liens `pubmedsearch://search?q=…&action=fill|queue|run&sort=…&max=…`.

- Bouton « Envoyer vers PubMed Search » sous la barre de recherche de PubMed.
- Icône de la barre d'outils : requête modifiable, tri, nombre max, trois actions
  (remplir le formulaire, ajouter à la file, lancer).
- Menu contextuel sur un texte sélectionné (n'importe quel site).

Construction : `zip -r -FS ../envoyer-vers-pubmed-search-1.0.0.xpi * -x README.md`
Signature Mozilla (facultative) : `web-ext sign --channel=unlisted --api-key=… --api-secret=…`
