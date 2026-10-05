"use strict";

const $ = (id) => document.getElementById(id);
let tabId = null;

async function init() {
  const s = await PMS.getSettings();
  $("max").value = s.max;
  const [tab] = await browser.tabs.query({ active: true, currentWindow: true });
  tabId = tab ? tab.id : null;
  const info = tab ? PMS.parsePubmedUrl(tab.url || "") : null;
  if (info) {
    $("query").value = info.query;
    $("sort").value = s.sortMode === "page" ? info.sort : s.sortMode;
    const parts = [];
    if (info.converted.length) parts.push(`${info.converted.length} filtre(s) de PubMed ajouté(s) à la requête.`);
    if (info.ignored.length) parts.push(`Non convertis, à ajouter à la main si besoin : ${info.ignored.join(", ")}.`);
    $("note").textContent = parts.join(" ");
  } else {
    $("sort").value = s.sortMode === "page" ? "best" : s.sortMode;
    $("note").textContent = "Aucune recherche PubMed dans cet onglet : saisissez une requête.";
  }
  for (const action of ["fill", "queue", "run"]) $(action).addEventListener("click", () => send(action));
  $("options").addEventListener("click", (e) => { e.preventDefault(); browser.runtime.openOptionsPage(); });
  $("query").focus();
}

async function send(action) {
  const query = $("query").value.trim().replace(/\s+/g, " ");
  if (!query) {
    $("status").className = "hint err";
    $("status").textContent = "Saisissez une requête PubMed.";
    return;
  }
  const max = Math.max(1, Math.min(10000, parseInt($("max").value, 10) || 100));
  const url = PMS.buildAppUrl(query, { action, sort: $("sort").value, max });
  try {
    if (tabId === null) throw new Error("aucun onglet");
    await browser.runtime.sendMessage({ type: "pms-open-tab", tabId, url });
    $("status").className = "hint ok";
    $("status").textContent = { fill: "Formulaire rempli dans PubMed Search.",
      queue: "Recherche ajoutée à la file de PubMed Search.",
      run: "Recherche lancée dans PubMed Search." }[action];
    setTimeout(() => window.close(), 1200);
  } catch (e) {
    $("status").className = "hint err";
    $("status").textContent = "Envoi impossible depuis cet onglet (page interne de Firefox ?). Ouvrez un onglet PubMed et réessayez.";
  }
}

init();
