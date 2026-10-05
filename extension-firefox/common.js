/* Fonctions partagées : lecture d'une URL PubMed et construction du lien pubmedsearch:// */
"use strict";

const PMS = {};

/* Filtres de la barre latérale de PubMed -> syntaxe de requête équivalente */
PMS.FILTERS = {
  "pubt.review": "review[pt]",
  "pubt.systematicreview": "systematic review[pt]",
  "pubt.meta-analysis": "meta-analysis[pt]",
  "pubt.randomizedcontrolledtrial": "randomized controlled trial[pt]",
  "pubt.clinicaltrial": "clinical trial[pt]",
  "pubt.casereports": "case reports[pt]",
  "pubt.booksdocs": "books and documents[filter]",
  "simsearch1.fha": "hasabstract",
  "simsearch2.ffrft": "free full text[sb]",
  "simsearch3.fft": "full text[sb]",
  "hum_ani.humans": "humans[mh]",
  "hum_ani.animal": "animals[mh] NOT humans[mh]",
  "lang.english": "english[la]",
  "lang.french": "french[la]",
  "lang.spanish": "spanish[la]",
  "lang.german": "german[la]",
  "lang.portuguese": "portuguese[la]",
  "sex.female": "female[mh]",
  "sex.male": "male[mh]",
  "other.excludepreprints": "NOT preprint[pt]",
  "datesearch.y_1": "\"last 1 year\"[dp]",
  "datesearch.y_5": "\"last 5 years\"[dp]",
  "datesearch.y_10": "\"last 10 years\"[dp]"
};

/* Analyse une URL pubmed.ncbi.nlm.nih.gov : requête, tri, filtres convertis/ignorés */
PMS.parsePubmedUrl = function (href) {
  let url;
  try { url = new URL(href); } catch (e) { return null; }
  if (!/(^|\.)pubmed\.ncbi\.nlm\.nih\.gov$/.test(url.hostname)) return null;
  const p = url.searchParams;
  let term = (p.get("term") || "").trim();
  const converted = [], ignored = [];
  for (const f of p.getAll("filter")) {
    const m = /^years\.(\d{4})-(\d{4})$/.exec(f);
    if (m) { converted.push(`("${m[1]}"[dp] : "${m[2]}"[dp])`); continue; }
    if (PMS.FILTERS[f]) converted.push(PMS.FILTERS[f]); else ignored.push(f);
  }
  // Page d'un article : /12345678/
  const art = /^\/(\d{5,9})\/?$/.exec(url.pathname);
  if (!term && art) term = `${art[1]}[pmid]`;
  if (!term) return null;
  let query = term;
  if (converted.length) {
    const pos = converted.filter(c => !c.startsWith("NOT "));
    const neg = converted.filter(c => c.startsWith("NOT "));
    query = [`(${term})`].concat(pos.map(c => (/ OR | NOT /.test(c) ? `(${c})` : c))).join(" AND ");
    if (neg.length) query += " " + neg.join(" ");
  }
  const sort = p.get("sort") === "date" || p.get("sort") === "pubdate" ? "recent" : "best";
  return { term, query, sort, converted, ignored, isArticle: !!art };
};

/* Lien compris par l'application PubMed Search */
PMS.buildAppUrl = function (query, opts) {
  const params = new URLSearchParams();
  params.set("q", query);
  params.set("action", (opts && opts.action) || "fill");
  if (opts && opts.sort) params.set("sort", opts.sort);
  if (opts && opts.max) params.set("max", String(opts.max));
  params.set("source", "firefox");
  return "pubmedsearch://search?" + params.toString();
};

PMS.DEFAULTS = { action: "fill", sortMode: "page", max: 100 };

PMS.getSettings = async function () {
  try {
    const s = await browser.storage.local.get(PMS.DEFAULTS);
    return Object.assign({}, PMS.DEFAULTS, s);
  } catch (e) {
    return Object.assign({}, PMS.DEFAULTS);
  }
};
