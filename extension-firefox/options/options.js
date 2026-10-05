"use strict";
async function load() {
  const s = await PMS.getSettings();
  document.querySelector(`input[name=action][value=${s.action}]`).checked = true;
  document.querySelector(`input[name=sortMode][value=${s.sortMode}]`).checked = true;
  document.getElementById("max").value = s.max;
}
async function save() {
  const max = Math.max(1, Math.min(10000, parseInt(document.getElementById("max").value, 10) || 100));
  await browser.storage.local.set({
    action: document.querySelector("input[name=action]:checked").value,
    sortMode: document.querySelector("input[name=sortMode]:checked").value,
    max
  });
  const el = document.getElementById("saved");
  el.textContent = "Réglages enregistrés.";
  setTimeout(() => { el.textContent = ""; }, 1500);
}
document.addEventListener("change", save);
load();
