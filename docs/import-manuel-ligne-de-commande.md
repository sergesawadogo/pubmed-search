# Compléter les PDF manquants en ligne de commande (`--import-pdf`)

Équivalent, sans l'interface, de l'onglet **Import manuel** de l'application.
Prérequis : `pip install requests openpyxl pypdf`.

## 1. Extraire les liens à ouvrir
```bash
cd <DOSSIER_RESULTATS>
grep "À ouvrir manuellement" error.txt | sed 's/.*manuellement : //' > liens.txt
wc -l liens.txt
```

## 2. Préparer un dossier de réception réservé à ces PDF
```bash
mkdir -p ~/Téléchargements/pdf_import
```

## 3. Ouvrir les liens par lots de 10 et enregistrer les PDF
```bash
sed -n '1,10p'  liens.txt | xargs -n1 xdg-open
sed -n '11,20p' liens.txt | xargs -n1 xdg-open
```
- PMC : patienter quelques secondes pendant la vérification, puis le PDF s'affiche.
- Page d'article au lieu du PDF : cliquer « PDF » ou « Download PDF ».
- Accès payant : passer, l'article n'est pas libre chez cet éditeur.

## 4. Importer
Fermer le fichier Excel des résultats, puis :
```bash
python3 script/pubmed_search.py --import-pdf ~/Téléchargements/pdf_import --results <DOSSIER_RESULTATS>
```
Ajouter `--move` pour déplacer les PDF au lieu de les copier. Reconnaissance : PMID dans le nom
du fichier, sinon DOI ou titre lus dans les 3 premières pages. Un PDF non reconnu (scan sans
texte) : le renommer `<PMID>.pdf` et relancer la commande.
