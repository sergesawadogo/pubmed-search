/* Sur pubmed.ncbi.nlm.nih.gov : bouton « Envoyer vers PubMed Search » + relais pour la fenêtre de l'extension */
"use strict";

function openApp(url) {
  // Le lien pubmedsearch:// est confié à Firefox, qui demande quelle application l'ouvre
  // (la première fois) puis le transmet à PubMed Search. La page PubMed reste affichée.
  window.location.href = url;
}

browser.runtime.onMessage.addListener((msg) => {
  if (msg && msg.type === "pms-open" && typeof msg.url === "string" && msg.url.startsWith("pubmedsearch://")) {
    openApp(msg.url);
    return Promise.resolve(true);
  }
  if (msg && msg.type === "pms-parse") {
    return Promise.resolve(PMS.parsePubmedUrl(location.href));
  }
  return undefined;
});

async function sendCurrentSearch(btn) {
  const info = PMS.parsePubmedUrl(location.href);
  if (!info) {
    btn.textContent = "Lancez d'abord une recherche PubMed";
    setTimeout(() => { btn.textContent = "Envoyer vers PubMed Search"; }, 2500);
    return;
  }
  const s = await PMS.getSettings();
  const sort = s.sortMode === "page" ? info.sort : s.sortMode;
  openApp(PMS.buildAppUrl(info.query, { action: s.action, sort, max: s.max }));
  btn.textContent = "Envoyée à PubMed Search";
  setTimeout(() => { btn.textContent = "Envoyer vers PubMed Search"; }, 2500);
}

function injectButton() {
  if (document.getElementById("pms-send")) return;
  const info = PMS.parsePubmedUrl(location.href);
  if (!info) return;                       // page d'accueil sans recherche : pas de bouton
  const btn = document.createElement("button");
  btn.id = "pms-send";
  btn.type = "button";
  btn.textContent = "Envoyer vers PubMed Search";
  btn.title = "Requête envoyée : " + info.query +
    (info.ignored.length ? "\nFiltres non convertis : " + info.ignored.join(", ") : "");
  btn.addEventListener("click", () => sendCurrentSearch(btn));
  const anchor = document.querySelector("#search-form") || document.querySelector("form.search-form") ||
                 document.querySelector("main") || document.body;
  const wrap = document.createElement("div");
  wrap.id = "pms-wrap";
  wrap.appendChild(btn);
  if (anchor === document.body) document.body.prepend(wrap);
  else anchor.insertAdjacentElement("afterend", wrap);
}

injectButton();
// PubMed modifie l'adresse sans recharger la page (pagination, filtres) : on surveille l'URL.
let lastHref = location.href;
setInterval(() => {
  if (location.href !== lastHref) {
    lastHref = location.href;
    const old = document.getElementById("pms-wrap");
    if (old) old.remove();
    injectButton();
  }
}, 1000);
