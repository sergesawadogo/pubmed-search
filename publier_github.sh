#!/usr/bin/env bash
# Publie ce dossier sur GitHub : crée le dépôt public (s'il n'existe pas), pousse le code et
# la notice, puis crée la Release avec les installateurs du dossier release/.
#
# Prérequis : git, curl, python3 ; un jeton GitHub « classic » avec la portée public_repo
#   (github.com → Settings → Developer settings → Personal access tokens → Tokens (classic)).
# Utilisation :  ./publier_github.sh            (le jeton est demandé, il n'est pas affiché)
#                GITHUB_TOKEN=xxx ./publier_github.sh
set -euo pipefail
cd "$(dirname "$0")"

OWNER="sergesawadogo"
REPO="pubmed-search"
TAG="v1.2.0"
TITLE="PubMed Search 1.2.0"
NOTES="docs/notes-version-1.2.0.md"
API="${GH_API:-https://api.github.com}"          # variables GH_* : tests uniquement
UPLOADS="${GH_UPLOADS:-https://uploads.github.com}"
GITURL="${GH_GIT:-https://github.com/$OWNER/$REPO.git}"
ASSETS=(release/PubMedSearch-Setup-1.2.0.exe release/pubmed-search-gui_1.2.0_amd64.deb
        release/envoyer-vers-pubmed-search-1.0.0.xpi release/SHA256SUMS.txt)

for cmd in git curl python3; do
  command -v "$cmd" >/dev/null || { echo "Commande manquante : $cmd (sudo apt install $cmd)"; exit 1; }
done
for f in "${ASSETS[@]}"; do
  [ -f "$f" ] || { echo "Fichier absent : $f — placez les installateurs dans le dossier release/."; exit 1; }
done
echo "Vérification des empreintes…"
(cd release && sha256sum -c --quiet SHA256SUMS.txt) || { echo "Empreintes incorrectes : fichier abîmé ou d'une autre version."; exit 1; }

TOKEN="${GITHUB_TOKEN:-}"
if [ -z "$TOKEN" ]; then
  read -rsp "Jeton GitHub (public_repo) : " TOKEN; echo
fi
AUTH=(-H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github+json" -H "X-GitHub-Api-Version: 2022-11-28")
json() { python3 -c "import sys,json; d=json.load(sys.stdin); print(d$1 if isinstance(d,dict) else '')" 2>/dev/null || true; }

LOGIN=$(curl -fsS "${AUTH[@]}" "$API/user" | json "['login']") || true
[ -n "$LOGIN" ] || { echo "Jeton refusé par GitHub."; exit 1; }
[ "$LOGIN" = "$OWNER" ] || { echo "Ce jeton appartient à « $LOGIN », pas à « $OWNER »."; exit 1; }

# 1. Dépôt public
CODE=$(curl -s -o /dev/null -w '%{http_code}' "${AUTH[@]}" "$API/repos/$OWNER/$REPO")
if [ "$CODE" = "404" ]; then
  echo "Création du dépôt public $OWNER/$REPO…"
  curl -fsS "${AUTH[@]}" -X POST "$API/user/repos" -d "$(python3 - <<'PY'
import json
print(json.dumps({"name": "pubmed-search", "private": False, "has_issues": True, "has_wiki": False,
  "description": "Recherche PubMed et téléchargement des PDF en libre accès (Windows, Linux, extension Firefox)"}))
PY
)" >/dev/null
else
  echo "Le dépôt $OWNER/$REPO existe déjà : mise à jour."
fi

# 2. Code et notice (le jeton n'est écrit dans aucun fichier)
B64=$(printf 'x-access-token:%s' "$TOKEN" | base64 | tr -d '\n')
GITAUTH=(-c "http.https://github.com/.extraheader=AUTHORIZATION: basic $B64")
[ -d .git ] || git init -q -b main
git config user.name >/dev/null || git config user.name "$OWNER"
git config user.email >/dev/null || git config user.email "$OWNER@users.noreply.github.com"
git add -A
git diff --cached --quiet || git commit -q -m "PubMed Search ${TAG#v}"
git remote get-url origin >/dev/null 2>&1 || git remote add origin "$GITURL"
echo "Envoi du code…"
git "${GITAUTH[@]}" push -q -u origin main

# 3. Release + fichiers
REL=$(curl -s "${AUTH[@]}" "$API/repos/$OWNER/$REPO/releases/tags/$TAG")
RID=$(echo "$REL" | json "['id']")
if [ -z "$RID" ]; then
  echo "Création de la Release $TAG…"
  BODY=$(python3 -c "import json,sys; print(json.dumps({'tag_name':'$TAG','target_commitish':'main','name':'$TITLE','body':open('$NOTES',encoding='utf-8').read(),'make_latest':'true'}))")
  REL=$(curl -fsS "${AUTH[@]}" -X POST "$API/repos/$OWNER/$REPO/releases" -d "$BODY")
  RID=$(echo "$REL" | json "['id']")
fi
for f in "${ASSETS[@]}"; do
  name=$(basename "$f")
  OLD=$(curl -s "${AUTH[@]}" "$API/repos/$OWNER/$REPO/releases/$RID/assets" |
        python3 -c "import sys,json; print(next((a['id'] for a in json.load(sys.stdin) if a['name']=='$name'),''))")
  [ -n "$OLD" ] && curl -fsS "${AUTH[@]}" -X DELETE "$API/repos/$OWNER/$REPO/releases/assets/$OLD" >/dev/null
  echo "Envoi de $name…"
  curl -fsS "${AUTH[@]}" -H "Content-Type: application/octet-stream" --data-binary @"$f" \
       "$UPLOADS/repos/$OWNER/$REPO/releases/$RID/assets?name=$name" >/dev/null
done

echo
echo "Publié. Liens à partager :"
echo "  Page du projet et notice : https://github.com/$OWNER/$REPO"
echo "  Téléchargements          : https://github.com/$OWNER/$REPO/releases/latest"
for f in "${ASSETS[@]}"; do
  echo "  $(basename "$f") : https://github.com/$OWNER/$REPO/releases/download/$TAG/$(basename "$f")"
done
