#!/usr/bin/env bash
# Publie ce dossier sur GitHub : crée le dépôt public (s'il n'existe pas), pousse le code et
# la notice, puis crée la Release avec les installateurs du dossier release/.
# Les versions macOS (.dmg) ne peuvent pas être construites ici : l'envoi du code déclenche
# leur construction sur les Mac de GitHub (fichier .github/workflows/macos.yml), qui les
# ajoute à la Release une quinzaine de minutes plus tard.
#
# Prérequis : git, curl, python3 ; un jeton GitHub « classic » avec les portées
#   public_repo ET workflow (github.com → Settings → Developer settings →
#   Personal access tokens → Tokens (classic) → Generate new token (classic)).
#   Sans « workflow », GitHub refuse l'envoi du fichier qui construit la version macOS.
# Utilisation :  ./publier_github.sh            (le jeton est demandé, il n'est pas affiché)
#                GITHUB_TOKEN=xxx ./publier_github.sh
set -euo pipefail
cd "$(dirname "$0")"

OWNER="sergesawadogo"
REPO="pubmed-search"
VERSION=$(python3 interface/packaging/version.py)   # lue dans interface/pubmed_gui.py
TAG="v$VERSION"
TITLE="PubMed Search $VERSION"
NOTES="docs/notes-version-$VERSION.md"
API="${GH_API:-https://api.github.com}"          # variables GH_* : tests uniquement
UPLOADS="${GH_UPLOADS:-https://uploads.github.com}"
GITURL="${GH_GIT:-https://github.com/$OWNER/$REPO.git}"
ASSETS=(release/PubMedSearch-Setup-$VERSION.exe release/pubmed-search-gui_${VERSION}_amd64.deb
        release/envoyer-vers-pubmed-search-1.0.0.xpi release/SHA256SUMS.txt)

[ -f "$NOTES" ] || { echo "Notes de version absentes : $NOTES"; exit 1; }
for cmd in git curl python3; do
  command -v "$cmd" >/dev/null || { echo "Commande manquante : $cmd (sudo apt install $cmd)"; exit 1; }
done
# Fichiers envoyés en morceaux (.part00, .part01…) : reconstitution automatique
for f in "${ASSETS[@]}"; do
  if [ ! -f "$f" ] && ls "$f".part?? >/dev/null 2>&1; then
    echo "Reconstitution de $f à partir de ses morceaux…"
    cat "$f".part?? > "$f"
  fi
done
for f in "${ASSETS[@]}"; do
  [ -f "$f" ] || { echo "Fichier absent : $f — placez les installateurs dans le dossier release/."; exit 1; }
done
echo "Vérification des empreintes…"
(cd release && sha256sum -c --quiet SHA256SUMS.txt) || { echo "Empreintes incorrectes : fichier abîmé ou d'une autre version."; exit 1; }

# Fichier de construction macOS : copie de référence dans interface/packaging/macos/
WF=.github/workflows/macos.yml
WFSRC=interface/packaging/macos/github-workflow-macos.yml
if [ -f "$WFSRC" ] && ! cmp -s "$WFSRC" "$WF" 2>/dev/null; then
  mkdir -p .github/workflows && cp "$WFSRC" "$WF"
fi

TOKEN="${GITHUB_TOKEN:-}"
if [ -z "$TOKEN" ]; then
  read -rsp "Jeton GitHub (public_repo + workflow) : " TOKEN; echo
fi
AUTH=(-H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github+json" -H "X-GitHub-Api-Version: 2022-11-28")
json() { python3 -c "import sys,json; d=json.load(sys.stdin); print(d$1 if isinstance(d,dict) else '')" 2>/dev/null || true; }

LOGIN=$(curl -fsS "${AUTH[@]}" "$API/user" | json "['login']") || true
[ -n "$LOGIN" ] || { echo "Jeton refusé par GitHub."; exit 1; }
[ "$LOGIN" = "$OWNER" ] || { echo "Ce jeton appartient à « $LOGIN », pas à « $OWNER »."; exit 1; }
SCOPES=$(curl -sS -o /dev/null -D - "${AUTH[@]}" "$API/user" | tr -d '\r' |
         awk -F': ' 'tolower($1)=="x-oauth-scopes"{print $2}')
if [ -n "$SCOPES" ]; then
  case ", $SCOPES," in
    *", workflow,"*) ;;
    *) echo "Ce jeton n'a pas la portée « workflow » (portées : $SCOPES)."
       echo "Créez un jeton classic en cochant public_repo ET workflow, puis relancez."; exit 1 ;;
  esac
fi

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
else
  echo "Mise à jour du texte de la Release $TAG…"
  BODY=$(python3 -c "import json; print(json.dumps({'name':'$TITLE','body':open('$NOTES',encoding='utf-8').read()}))")
  curl -fsS "${AUTH[@]}" -X PATCH "$API/repos/$OWNER/$REPO/releases/$RID" -d "$BODY" >/dev/null
fi
LIST=$(curl -s "${AUTH[@]}" "$API/repos/$OWNER/$REPO/releases/$RID/assets?per_page=100")
for f in "${ASSETS[@]}"; do
  name=$(basename "$f")
  size=$(wc -c < "$f" | tr -d ' ')
  read -r OLD OLDSIZE < <(echo "$LIST" | python3 -c "import sys,json; a=next((a for a in json.load(sys.stdin) if a['name']=='$name'),None); print(a['id'], a['size']) if a else print('', '')")
  if [ -n "$OLD" ] && [ "$OLDSIZE" = "$size" ]; then
    echo "$name : déjà en ligne, identique (taille) — conservé."; continue
  fi
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
echo
echo "Version macOS : construction en cours sur GitHub (10 à 20 minutes). Suivi :"
echo "  https://github.com/$OWNER/$REPO/actions"
echo "Les fichiers PubMedSearch-${TAG#v}-macOS-AppleSilicon.dmg et …-Intel.dmg apparaîtront"
echo "ensuite sur la page des téléchargements. En cas d'échec (croix rouge), envoyez-moi le"
echo "journal de l'étape en erreur."
