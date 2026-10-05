/* Menu contextuel : envoyer un texte sélectionné (n'importe quelle page) comme requête */
"use strict";

function createMenus() {
  browser.contextMenus.removeAll().then(() => {
    browser.contextMenus.create({
      id: "pms-selection",
      title: "Envoyer « %s » vers PubMed Search",
      contexts: ["selection"]
    });
  });
}
browser.runtime.onInstalled.addListener(createMenus);
browser.runtime.onStartup.addListener(createMenus);

async function openInTab(tabId, url) {
  try {
    await browser.tabs.sendMessage(tabId, { type: "pms-open", url });
  } catch (e) {
    // Onglet hors PubMed : activeTab (clic sur le menu ou l'icône) autorise ce petit script.
    await browser.tabs.executeScript(tabId, { code: `window.location.href = ${JSON.stringify(url)};` });
  }
}

browser.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId !== "pms-selection" || !info.selectionText) return;
  const s = await PMS.getSettings();
  const sort = s.sortMode === "page" ? "best" : s.sortMode;
  const url = PMS.buildAppUrl(info.selectionText.trim(), { action: s.action, sort, max: s.max });
  openInTab(tab.id, url);
});

browser.runtime.onMessage.addListener((msg) => {
  if (msg && msg.type === "pms-open-tab") return openInTab(msg.tabId, msg.url);
  return undefined;
});
