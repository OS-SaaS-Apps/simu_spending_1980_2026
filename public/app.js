"use strict";

// Emplacement des fichiers Excel (modifiable : toute URL accessible en CORS).
const DATA_URLS = { y80: "data/budget_1980.xlsx", y26: "data/budget_2026.xlsx" };

const SERIES = {
  r80: { key: "--c80", label: "1980" },
  t80: { key: "--ctr", label: "1980 transposé en 2026" },
  y26: { key: "--c26", label: "2026" },
};
const FIAB_ICON = { "élevée": "●●●", "moyenne": "●●○", "faible": "●○○" };

const state = { side: "dep", unit: "capita", money: "const", sort: "v26", sSide: "dep", sUnit: "capita", gSide: "dep", gUnit: "capita", choix: {} };
try { Object.assign(state, JSON.parse(localStorage.getItem("b8026") || "{}")); } catch (e) { /* stockage indisponible */ }

let D = null; // données calculées

// ---------- Formatage ----------
const nf0 = new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 0 });
const nf1 = new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 1, minimumFractionDigits: 1 });
const eur = (v) => `${nf0.format(Math.round(v / 10) * 10)} €`;
const eurExact = (v) => `${nf0.format(v)} €`;
const md = (v) => `${Math.abs(v) < 100 ? nf1.format(v) : nf0.format(v)} Md€`;
const pct = (v) => `${nf1.format(v)} %`;
const signed = (s, v) => (v > 0 ? "+" : v < 0 ? "−" : "") + s(Math.abs(v));
const parFr = (s) => `${s} / Français`;
const UNIT_FR = `<span class="u">/ Français</span>`;
const UNIT_EL = `<span class="u">/ élève</span>`;
// école : le nombre d'élèves baisse quand la population augmente, le montant par élève est donc plus parlant
const NOTE_EL = `<small class="rule">Barres en € par élève, à leur propre échelle.</small>`;
const parEl = (v) => `${eur(v)} / élève`;
// largeur d'une barre horizontale : la place de l'étiquette est réservée pour garder les proportions
const barW = (ratio, capita) => `calc((100% - ${capita ? 125 : 70}px) * ${ratio})`;
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// ---------- Lecture des Excel ----------
async function loadWorkbook(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`${url} : ${res.status}`);
  return XLSX.read(await res.arrayBuffer(), { type: "array" });
}

function rows(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`Feuille « ${name} » introuvable`);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}

function readYear(wb) {
  const params = {};
  for (const r of rows(wb, "Paramètres").slice(1)) if (r[0]) params[r[0]] = r[1];
  const toEur = params.franc ? 1 / params.franc : 1; // 1980 : saisie en francs
  const readCats = (sheet) => {
    const all = rows(wb, sheet);
    const head = all[0].map((h) => String(h || ""));
    const col = (re) => head.findIndex((h) => re.test(h));
    const iVal = col(/^Montant/), iFiab = col(/^Fiabilit/), iSrc = col(/^Source/), iIdx = col(/^Indexation/);
    return all.slice(1).filter((r) => r[0] && r[0] !== "TOTAL").map((r) => ({
      id: r[0], label: r[1], icon: r[2] || "", desc: r[3] || "",
      cur: (Number(r[iVal]) || 0) * toEur, // Md€ courants
      fiab: r[iFiab] || "", src: r[iSrc] || "", idx: (iIdx >= 0 && r[iIdx]) || "pib",
    }));
  };
  const method = rows(wb, "Méthode").map((r) => r[0]).filter(Boolean);
  const sources = rows(wb, "Sources").slice(1).filter((r) => r[0]);
  const ind = {};
  if (wb.Sheets["Indicateurs"]) for (const r of rows(wb, "Indicateurs").slice(1)) {
    if (r[0]) ind[r[0]] = { label: r[1], val: r[2] === null || r[2] === "" ? null : Number(r[2]), unit: r[3] || "", note: r[4] || "", fiab: r[5] || "" };
  }
  return {
    params, dep: readCats("Dépenses"), rec: readCats("Recettes"), method, sources, ind,
    nationals: params.pop - params.etrangers, // millions
    coef: Number(params.coef) || 1,
    pib: (params.pib || 0) * toEur,
  };
}

function compute(a, b) {
  const build = (side) => {
    const byId = new Map(b[side].map((c) => [c.id, c]));
    return a[side].map((c80) => {
      const c26 = byId.get(c80.id) || { cur: 0, fiab: "", src: "", idx: "pib" };
      const r80c = c80.cur * a.coef; // Md€ 2026
      const cap80 = (r80c * 1e3) / a.nationals; // € 2026 par Français
      return {
        id: c80.id, label: c80.label, icon: c80.icon, desc: c80.desc || c26.desc,
        fiab80: c80.fiab, src80: c80.src, fiab26: c26.fiab, src26: c26.src, idx: c26.idx,
        r80cur: c80.cur, r80: r80c, t80: (cap80 * b.nationals) / 1e3, y26: c26.cur,
        cap80cur: (c80.cur * 1e3) / a.nationals, cap80, cap26: (c26.cur * 1e3) / b.nationals,
      };
    });
  };
  const dep = build("dep"), rec = build("rec");
  const sum = (arr, k) => arr.reduce((s, c) => s + c[k], 0);
  const tot = (k) => ({ dep: sum(dep, k), rec: sum(rec, k) });
  return {
    a, b, dep, rec,
    totals: {
      r80cur: tot("r80cur"), r80: tot("r80"), t80: tot("t80"), y26: tot("y26"),
      cap80: tot("cap80"), cap26: tot("cap26"), cap80cur: tot("cap80cur"),
    },
  };
}

// ---------- Scénario « 2026 à la sauce 1980 » ----------
// Chaque poste 1980 (en € constants) est appliqué à la démographie de 2026 : même montant par Français,
// par bénéficiaire ou par cotisant (voir la colonne « Indexation » du fichier 2026).
// Montant moyen par bénéficiaire affiché sous les barres, pour les postes qui ont des bénéficiaires dénombrés.
const AVG_SAUCE = {
  retraites: { k: "retraites_nb", lbl: "Par retraité", per: "mois", div: 12 },
  moins20: { k: "moins20", lbl: "Par jeune de moins de 20 ans", per: "an", div: 1 },
  chomeurs: { k: "chomeurs", lbl: "Par chômeur", per: "an", div: 1 },
  eleves: { k: "eleves", lbl: "Par élève", per: "an", div: 1 },
  etudiants: { k: "etudiants", lbl: "Par étudiant", per: "an", div: 1 },
};

function computeSauce() {
  const a = D.a, b = D.b;
  const N = b.nationals / a.nationals;
  const ind = (y, k) => (y.ind[k] && y.ind[k].val) || null;
  const m = (v) => (v < 1 ? v.toLocaleString("fr-FR", { minimumFractionDigits: 2 }) : nf1.format(v));
  const perFr = (c) => ({ s: c.r80 * N, rule: `Même montant par Français qu'en 1980 : ${eur(c.cap80)} par an.` });
  const apply = (c) => {
    const k = c.idx || "francais";
    if (k === "reel") return { s: c.y26, rule: "Montant réel de 2026 : contrainte héritée, pas un choix politique." };
    if (k === "retraites") {
      const N80 = ind(a, "retraites_nb"), N26 = ind(b, "retraites_nb"), C80 = ind(a, "cotisants_nb"), C26 = ind(b, "cotisants_nb");
      if (!N80 || !N26 || !C80 || !C26) return perFr(c);
      // même effort par cotisant qu'en 1980 : la pension suit la richesse par cotisant et le nombre de cotisants par retraité
      const prod = (b.pib / C26) / ((a.pib * a.coef) / C80);
      const p80 = (c.r80 * 1e3) / N80 / 12, r80 = C80 / N80, r26 = C26 / N26, p = p80 * prod * (r26 / r80);
      return { s: c.r80 * (C26 / C80) * prod,
        rule: `Même effort par cotisant qu'en 1980 : la pension de 1980 (${eur(p80)} par mois) suit la richesse par cotisant ` +
          `(×${prod.toLocaleString("fr-FR", { maximumFractionDigits: 2 })}) et le nombre de cotisants par retraité ` +
          `(${nf1.format(r80)} → ${nf1.format(r26)}), soit ${eur(p)} par mois au lieu de ${eur((c.y26 * 1e3) / N26 / 12)}.` };
    }
    if (k !== "francais" && k !== "pib") {
      const d80 = ind(a, k), d26 = ind(b, k), lbl = (b.ind[k] || a.ind[k] || {}).label || k;
      if (d80 && d26) return { s: c.r80 * (d26 / d80),
        rule: `Même montant par bénéficiaire qu'en 1980 (${eur((c.r80 * 1e3) / d80)} par an). ${lbl} : ${m(d80)} M → ${m(d26)} M.` };
    }
    return perFr(c);
  };
  for (const c of D.dep) {
    const r = apply(c);
    const P = b.params;
    // école : l'école inclusive n'existait pas en 1980, la surcotisation retraite fausse le montant par élève de 2026
    const inclusive = (c.id === "scolaire" && Number(P.ecole_inclusive)) || 0;
    const surcot = c.id === "scolaire" && P.cas_education && P.taux_cas ? P.cas_education * (1 - P.taux_cas_1980 / P.taux_cas) : 0;
    if (inclusive) { r.s += inclusive; r.rule = `${r.rule} S'y ajoutent ${md(inclusive)} pour les AESH et l'école inclusive.`; }
    if (c.id === "scolaire") r.rule = `${r.rule} Pas d'éducation prioritaire : les ZEP, créées en 1981, et les REP sont supprimées.`;
    if (surcot) r.rule = `${r.rule} Barre 2026 hors ${md(surcot)} de surcotisation retraite.`;
    c.s26 = r.s; c.capS = (r.s * 1e3) / b.nationals; c.rule = r.rule; c.yBar = c.y26 - surcot;
    const av = AVG_SAUCE[c.idx], n = av && ind(b, av.k);
    c.avg = n ? { lbl: av.lbl, per: av.per, n, div: av.div, surcot, s: (c.s26 * 1e3) / n / av.div, y: (c.yBar * 1e3) / n / av.div,
      yLbl: surcot ? "réel hors surcotisation retraite" : "réel" } : null;
  }
  // le surplus de pensions est financé par les cotisations : le poste retraites ne creuse pas le solde
  const ret = D.dep.find((c) => c.id === "retraites");
  // les recettes reposent sur les revenus du travail : même montant par personne en emploi qu'en 1980
  const C80 = ind(a, "cotisants_nb"), C26 = ind(b, "cotisants_nb");
  const perJob = (c) => (C80 && C26
    ? { s: c.r80 * (C26 / C80), rule: `Même montant par personne en emploi qu'en 1980 : ${eur((c.r80 * 1e3) / C80)} par an. Emplois : ${nf1.format(C80)} M → ${nf1.format(C26)} M.` }
    : perFr(c));
  const extra = ret ? ret.s26 - perJob(ret).s : 0;
  for (const c of D.rec) {
    const r = perJob(c);
    if (c.id === "cotisations") { r.s += extra; r.rule = `${r.rule} S'y ajoutent ${md(extra)} de cotisations pour financer les pensions.`; }
    c.s26 = r.s; c.capS = (r.s * 1e3) / b.nationals; c.rule = r.rule; c.yBar = c.y26;
  }
  D.sauce = { dep: D.dep.reduce((x, c) => x + c.s26, 0), rec: D.rec.reduce((x, c) => x + c.s26, 0) };
}

function renderSauce() {
  const b = D.b, S = D.sauce, t = D.totals, capita = state.sUnit === "capita";
  const f = (v) => (capita ? eur((v * 1e3) / b.nationals) : md(v));
  const kpi = (lbl, v1, v2, note) => `<div class="skpi"><div class="lbl">${lbl}</div>
    <div class="pair"><div><small><i style="--k:var(--csa)"></i>Sauce 1980</small><b>${v1}</b></div><div><small><i style="--k:var(--c26)"></i>2026 réel</small><b>${v2}</b></div></div>
    <p>${note}</p></div>`;
  const sol = S.rec - S.dep, sol26 = t.y26.rec - t.y26.dep, lbl = (s) => (capita ? `${s} par Français` : s);
  document.getElementById("sauce-kpis").innerHTML =
    kpi(lbl("Dépenses"), f(S.dep), f(t.y26.dep), `Écart : ${signed(f, S.dep - t.y26.dep)}`) +
    kpi(lbl("Recettes"), f(S.rec), f(t.y26.rec), `Écart : ${signed(f, S.rec - t.y26.rec)}`) +
    kpi(lbl("Solde"), signed(f, sol), signed(f, sol26), `${signed(pct, (sol / b.pib) * 100)} du PIB, contre ${signed(pct, (sol26 / b.pib) * 100)}`);

  const list = D[state.sSide].slice().sort((x, y) => y.yBar - x.yBar);
  const val = (c, k) => (capita ? (c[k] * 1e3) / b.nationals : c[k]);
  const max = Math.max(...list.flatMap((c) => [val(c, "s26"), val(c, "yBar")]), 1e-9);
  document.getElementById("sauce-chart").innerHTML = list.map((c) => {
    const g = c.s26 - c.yBar; // Md€, mis en forme par f
    const el = capita && c.id === "scolaire" && c.avg ? { s26: c.avg.s, yBar: c.avg.y } : null;
    return `<div role="listitem" class="srow">
      <span class="cat"><span class="ico" aria-hidden="true">${c.icon}</span><span>${esc(c.label)}<small class="rule">${esc(c.rule)}</small>${el ? NOTE_EL : ""}</span></span>
      <span class="bars">${[["s26", "--csa"], ["yBar", "--c26"]].map(([k, col]) =>
        `<span class="bar-line"><span class="bar" style="--k:var(${col});width:${barW(val(c, k) / max, capita)}"></span><span class="bar-val">${el ? parEl(el[k]) : capita ? parFr(f(c[k])) : f(c[k])}</span></span>`).join("")}${c.avg && !el && state.sSide === "dep"
        ? `<span class="avg">${c.avg.lbl}&nbsp;: <b>${eur(c.avg.s)}/${c.avg.per}</b> sauce 1980, ${eur(c.avg.y)}/${c.avg.per} ${c.avg.yLbl}</span>` : ""}</span>
      <span class="delta"><div class="main">${el ? `${signed(eur, el.s26 - el.yBar)}${UNIT_EL}` : `${signed(f, g)}${capita ? UNIT_FR : ""}`}</div><div class="ratio">sauce 1980 vs 2026 réel</div></span>
    </div>`;
  }).join("");
}

// ---------- Jeu : votre budget ----------
// Chaque curseur part de la sauce 1980 et couvre au moins ±50 %, ainsi que le montant réel de 2026.
const THUMB = 18; // largeur du curseur en px, voir .gslider dans styles.css
const amountOf = (c) => state.choix[c.id] ?? c.s26;
const gameCap = () => state.gUnit === "capita";
const gameFmt = (v) => (gameCap() ? eur((v * 1e3) / D.b.nationals) : md(v));
const gameEl = (c) => (gameCap() && c.id === "scolaire" && c.avg ? c.avg.n : null);
const gameBarVal = (c, v) => (gameEl(c) ? parEl((v * 1e3) / gameEl(c)) : gameCap() ? parFr(gameFmt(v)) : gameFmt(v));
const onTrack = (p) => `calc(${THUMB / 2}px + (100% - ${THUMB}px) * ${p})`;

function setupGame() {
  D.gMax = {};
  for (const side of ["dep", "rec"]) {
    for (const c of D[side]) {
      const span = Math.max(c.s26 * 0.5, Math.abs(c.s26 - c.y26) * 1.2);
      c.step = span < 10 ? 0.1 : span < 50 ? 0.5 : 1;
      c.tMin = -Math.floor(Math.min(span, c.s26) / c.step);
      c.tMax = Math.ceil(span / c.step);
    }
    D.gMax[side] = Math.max(...D[side].map((c) => c.s26 + c.tMax * c.step));
  }
}

const trackPos = (c, v) => {
  const t = Math.min(c.tMax, Math.max(c.tMin, (v - c.s26) / c.step));
  return (t - c.tMin) / (c.tMax - c.tMin);
};

function gameAvg(c) {
  const a = c.avg, v = (x) => (x * 1e3) / a.n / a.div;
  return `${a.lbl}&nbsp;: <b>${eur(v(amountOf(c)))}/${a.per}</b>, contre ${eur(v(c.y26))} en 2026 réel` +
    `${a.surcot ? ` (${eur(a.y)} hors surcotisation retraite)` : ""} et ${eur(a.s)} sauce 1980`;
}

function renderGameKpis() {
  const b = D.b, S = D.sauce, sum = (arr) => arr.reduce((x, c) => x + amountOf(c), 0);
  const dep = sum(D.dep), rec = sum(D.rec), sol = rec - dep, lbl = (s) => (gameCap() ? `${s} par Français` : s);
  const kpi = (lbl, val, note, cls = "") => `<div><div class="lbl">${lbl}</div><b class="${cls}">${val}</b><small>${note}</small></div>`;
  document.getElementById("game-kpis").innerHTML =
    kpi(lbl("Dépenses"), gameFmt(dep), `${signed(gameFmt, dep - S.dep)} vs sauce 1980`) +
    kpi(lbl("Recettes"), gameFmt(rec), `${signed(gameFmt, rec - S.rec)} vs sauce 1980`) +
    kpi(lbl(sol < 0 ? "Déficit" : "Excédent"), signed(gameFmt, sol), `${signed(pct, (sol / b.pib) * 100)} du PIB, sauce 1980 : ${signed(gameFmt, S.rec - S.dep)}`, sol < 0 ? "neg" : "");
  document.getElementById("game-result").innerHTML = resultChart("Votre budget", "g");
}

// Barres verticales en SVG : contrairement aux fonds CSS, elles s'impriment dans le rapport PDF.
function resultChart(name, uid) {
  const sum = (arr) => arr.reduce((x, c) => x + amountOf(c), 0), t = D.totals.y26;
  const groups = [
    { lbl: "2026 réel", k: "--c26", rec: t.rec, dep: t.dep, cx: 115 },
    { lbl: name, k: "--cje", rec: sum(D.rec), dep: sum(D.dep), cx: 345 },
  ];
  const BASE = 236, H = 196, W = 66, max = Math.max(...groups.flatMap((g) => [g.rec, g.dep]));
  const y = (v) => BASE - (v / max) * H;
  const bar = (x, v, k, op) => `<path style="fill:var(${k});fill-opacity:${op}" d="M${x} ${BASE}V${y(v) + 4}Q${x} ${y(v)} ${x + 4} ${y(v)}H${x + W - 4}Q${x + W} ${y(v)} ${x + W} ${y(v) + 4}V${BASE}Z"/>`;
  const svg = groups.map((g, i) => {
    const sol = g.rec - g.dep, xR = g.cx - W - 3, xD = g.cx + 3, top = y(Math.max(g.rec, g.dep)), id = `hatch-${uid}-${i}`;
    const col = sol < 0 ? "--bad" : "--good", xGap = sol < 0 ? xR : xD;
    const vals = [[xR, g.rec, "Recettes"], [xD, g.dep, "Dépenses"]];
    return `<pattern id="${id}" width="5" height="5" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="2" height="5" style="fill:var(${col})"/></pattern>
      ${bar(xR, g.rec, g.k, 0.55)}${bar(xD, g.dep, g.k, 1)}
      <rect x="${xGap}" y="${top}" width="${W}" height="${y(Math.min(g.rec, g.dep)) - top}" fill="url(#${id})"/>
      ${vals.map(([x, v, l]) => `<text x="${x + W / 2}" y="${top - 6}" class="v">${md(v)}</text><text x="${x + W / 2}" y="${BASE + 16}">${l}</text>`).join("")}
      <text x="${g.cx}" y="${BASE + 36}" class="v">${esc(g.lbl)}</text>
      <text x="${g.cx}" y="${BASE + 54}"${sol < 0 ? ' class="neg"' : ""}>${sol < 0 ? "Déficit" : "Excédent"} ${signed(md, sol)} (${signed(pct, (sol / D.b.pib) * 100)} du PIB)</text>`;
  }).join("");
  const alt = groups.map((g) => `${g.lbl} : recettes ${md(g.rec)}, dépenses ${md(g.dep)}`).join(". ");
  return `<svg class="rchart" viewBox="0 0 460 296" role="img" aria-label="${esc(alt)}">
    <line x1="20" x2="210" y1="${BASE}" y2="${BASE}"/><line x1="250" x2="440" y1="${BASE}" y2="${BASE}"/>
    <text x="230" y="130" class="arrow">→</text>${svg}</svg>`;
}

function updateGameRow(c, row) {
  const v = amountOf(c), g = v - c.s26, input = row.querySelector("input");
  row.querySelector(".bar").style.width = barW(v / D.gMax[state.gSide], gameCap());
  row.querySelector(".bar-val").textContent = gameBarVal(c, v);
  row.querySelector(".delta .main").innerHTML = gameEl(c) ? `${signed(eur, (g * 1e3) / gameEl(c))}${UNIT_EL}` : `${signed(gameFmt, g)}${gameCap() ? UNIT_FR : ""}`;
  row.querySelector(".delta .ratio").textContent = `${c.s26 > 0 ? `${signed(pct, (g / c.s26) * 100)} ` : ""}vs sauce 1980`;
  if (c.avg) row.querySelector(".avg").innerHTML = gameAvg(c);
  input.value = Math.round((v - c.s26) / c.step);
  input.setAttribute("aria-valuetext", gameBarVal(c, v));
  input.style.setProperty("--a", onTrack(Math.min(trackPos(c, v), trackPos(c, c.s26))));
  input.style.setProperty("--b", onTrack(Math.max(trackPos(c, v), trackPos(c, c.s26))));
}

function renderGame() {
  const cats = D[state.gSide], list = cats.slice().sort((x, y) => y.y26 - x.y26);
  const line = (c, v, col) => `<span class="bar-line"><span class="bar" style="--k:var(${col});width:${barW(v / D.gMax[state.gSide], gameCap())}"></span><span class="bar-val">${gameBarVal(c, v)}</span></span>`;
  const mark = (c, v, col) => `<i class="mk" aria-hidden="true" style="--k:var(${col});left:${onTrack(trackPos(c, v))}"></i>`;
  const host = document.getElementById("game-chart");
  host.innerHTML = list.map((c) => `<div role="listitem" class="grow" data-id="${esc(c.id)}">
      <span class="cat"><span class="ico" aria-hidden="true">${c.icon}</span><span>${esc(c.label)}${gameEl(c) ? NOTE_EL : ""}</span></span>
      <span class="bars">${line(c, amountOf(c), "--cje")}${line(c, c.s26, "--csa")}${line(c, c.y26, "--c26")}
        <span class="gslider">${mark(c, c.s26, "--csa")}${mark(c, c.y26, "--c26")}<input type="range" min="${c.tMin}" max="${c.tMax}" step="1" aria-label="${esc(c.label)} : votre budget"></span>
        ${c.avg ? `<span class="avg"></span>` : ""}</span>
      <span class="delta"><div class="main"></div><div class="ratio"></div></span>
    </div>`).join("");
  host.querySelectorAll(".grow").forEach((row) => {
    const c = cats.find((x) => x.id === row.dataset.id);
    updateGameRow(c, row);
    const input = row.querySelector("input");
    input.addEventListener("input", () => {
      state.choix[c.id] = c.s26 + Number(input.value) * c.step;
      updateGameRow(c, row); renderGameKpis();
    });
    input.addEventListener("change", save);
  });
  renderGameKpis();
}

function presetGame(fn) { state.choix = fn(); save(); renderGame(); }
document.getElementById("game-reset").addEventListener("click", () => presetGame(() => ({})));
document.getElementById("game-2026").addEventListener("click", () => presetGame(() => Object.fromEntries([...D.dep, ...D.rec].map((c) => [c.id, c.y26]))));

// ---------- Rapport du joueur ----------
function reportData() {
  const row = (c, side) => { const v = amountOf(c); return { side, label: c.label, v, s: c.s26, y: c.y26, g: v - c.s26 }; };
  const byV = (x, y) => y.v - x.v;
  const dep = D.dep.map((c) => row(c, "Dépense")).sort(byV), rec = D.rec.map((c) => row(c, "Recette")).sort(byV);
  const tot = (arr, k) => arr.reduce((x, r) => x + r[k], 0);
  const sum = (k) => ({ dep: tot(dep, k), rec: tot(rec, k), sol: tot(rec, k) - tot(dep, k) });
  return { dep, rec, choix: [...dep, ...rec].filter((r) => Math.abs(r.g) > 1e-9), v: sum("v"), s: sum("s"), y: sum("y"), pib: D.b.pib, nat: D.b.nationals };
}

const REPORT_NOTE = "Montants en euros de 2026. Point de départ : le budget de 1980 appliqué à la France de 2026 (« sauce 1980 »). Les montants de 1980 sont en partie estimés.";

function printReport() {
  const R = reportData(), today = new Date().toLocaleDateString("fr-FR", { dateStyle: "long" });
  const cap = (v) => eur((v * 1e3) / R.nat);
  const part = (r) => (r.s > 0 ? signed(pct, (r.g / r.s) * 100) : "");
  const table = (title, rows, total) => `<h2>${title}</h2><table>
    <thead><tr><th>Poste</th><th>Mon budget</th><th>Par Français</th><th>Sauce 1980</th><th>Écart</th><th>2026 réel</th></tr></thead>
    <tbody>${rows.map((r) => `<tr><td>${esc(r.label)}</td><td>${md(r.v)}</td><td>${cap(r.v)}</td><td>${md(r.s)}</td><td>${signed(md, r.g)}</td><td>${md(r.y)}</td></tr>`).join("")}</tbody>
    <tfoot><tr><td>Total</td><td>${md(total.v)}</td><td>${cap(total.v)}</td><td>${md(total.s)}</td><td>${signed(md, total.v - total.s)}</td><td>${md(total.y)}</td></tr></tfoot>
  </table>`;
  const line = (lbl, k, fn) => `<tr><td>${lbl}</td>${[R.v, R.s, R.y].map((x) => `<td>${fn(x[k])}</td>`).join("")}</tr>`;
  const solPct = (x) => signed(pct, (x.sol / R.pib) * 100);
  document.getElementById("report").innerHTML = `
    <h1>Mon budget pour la France</h1>
    <p>Établi le ${today}. ${REPORT_NOTE}</p>
    <h2>Résumé</h2>
    <table><thead><tr><th></th><th>Mon budget</th><th>Sauce 1980</th><th>2026 réel</th></tr></thead><tbody>
      ${line("Dépenses", "dep", md)}${line("Recettes", "rec", md)}${line("Solde", "sol", (v) => signed(md, v))}
      <tr><td>Solde en % du PIB</td><td>${solPct(R.v)}</td><td>${solPct(R.s)}</td><td>${solPct(R.y)}</td></tr>
    </tbody></table>
    <h2>De 2026 à mon budget</h2>
    ${resultChart("Mon budget", "r")}
    <h2>Mes choix (${R.choix.length})</h2>
    ${R.choix.length ? `<table><thead><tr><th>Poste</th><th>Type</th><th>Sauce 1980</th><th>Mon budget</th><th>Écart</th><th>Écart en %</th><th>2026 réel</th></tr></thead>
      <tbody>${R.choix.map((r) => `<tr><td>${esc(r.label)}</td><td>${r.side}</td><td>${md(r.s)}</td><td>${md(r.v)}</td><td>${signed(md, r.g)}</td><td>${part(r)}</td><td>${md(r.y)}</td></tr>`).join("")}</tbody></table>`
      : "<p>Aucun poste modifié : ce budget est celui de la sauce 1980.</p>"}
    ${table("Dépenses", R.dep, { v: R.v.dep, s: R.s.dep, y: R.y.dep })}
    ${table("Recettes", R.rec, { v: R.v.rec, s: R.s.rec, y: R.y.rec })}`;
  const title = document.title;
  document.title = "Mon budget";
  document.body.classList.add("printing");
  window.addEventListener("afterprint", () => { document.body.classList.remove("printing"); document.title = title; }, { once: true });
  window.print();
}

function excelReport() {
  const R = reportData(), n1 = (v) => Math.round(v * 10) / 10, cap = (v) => Math.round((v * 1e3) / R.nat);
  const sheet = (aoa, cols) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws["!cols"] = cols.map((wch) => ({ wch })); return ws; };
  const head = ["Poste", "Mon budget (Md€)", "Mon budget (€ par Français)", "Sauce 1980 (Md€)", "Écart (Md€)", "2026 réel (Md€)"];
  const rows = (arr) => arr.map((r) => [r.label, n1(r.v), cap(r.v), n1(r.s), n1(r.g), n1(r.y)]);
  const total = (k) => ["Total", n1(R.v[k]), cap(R.v[k]), n1(R.s[k]), n1(R.v[k] - R.s[k]), n1(R.y[k])];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet([
    ["Mon budget pour la France"], [REPORT_NOTE], [],
    ["", "Mon budget", "Sauce 1980", "2026 réel"],
    ...["dep", "rec", "sol"].map((k, i) => [["Dépenses (Md€)", "Recettes (Md€)", "Solde (Md€)"][i], n1(R.v[k]), n1(R.s[k]), n1(R.y[k])]),
    ["Solde (% du PIB)", ...[R.v, R.s, R.y].map((x) => n1((x.sol / R.pib) * 100))],
  ], [24, 12, 12, 12]), "Résumé");
  XLSX.utils.book_append_sheet(wb, sheet([["Poste", "Type", "Sauce 1980 (Md€)", "Mon budget (Md€)", "Écart (Md€)", "2026 réel (Md€)"],
    ...R.choix.map((r) => [r.label, r.side, n1(r.s), n1(r.v), n1(r.g), n1(r.y)])], [38, 10, 16, 16, 12, 14]), "Mes choix");
  XLSX.utils.book_append_sheet(wb, sheet([head, ...rows(R.dep), total("dep")], [38, 16, 24, 16, 12, 14]), "Dépenses");
  XLSX.utils.book_append_sheet(wb, sheet([head, ...rows(R.rec), total("rec")], [38, 16, 24, 16, 12, 14]), "Recettes");
  XLSX.writeFile(wb, "mon_budget.xlsx");
}
document.getElementById("game-pdf").addEventListener("click", printReport);
document.getElementById("game-xlsx").addEventListener("click", excelReport);

// ---------- Rendu : en-tête ----------
function renderLede() {
  const t = D.totals;
  const def80 = t.r80cur.rec - t.r80cur.dep, def26 = t.y26.rec - t.y26.dep;
  const ratio = t.cap26.dep / t.cap80.dep;
  document.getElementById("lede").innerHTML =
    `En 1980, l'ensemble des administrations publiques dépensait <strong>${eur(t.cap80.dep)} par Français</strong> ` +
    `(en euros d'aujourd'hui), avec un budget quasiment à l'équilibre (déficit de ${pct(-def80 / D.a.pib * 100)} du PIB). ` +
    `En 2026, c'est <strong>${eur(t.cap26.dep)}</strong>, soit ${nf1.format(ratio)} fois plus, ` +
    `avec un déficit de <strong>${eur(-(t.cap26.rec - t.cap26.dep))} par Français</strong> (${pct(-def26 / D.b.pib * 100)} du PIB).`;
}

function card(o, maxCap) {
  const solde = o.rec - o.dep, capSolde = o.capRec - o.capDep;
  const bad = !o.balanced; // quasi-équilibre : déficit inférieur à 1 % du PIB
  const w = (v) => `${Math.max(0, (v / maxCap) * 100)}%`;
  return `
  <article class="card" style="--key:var(${o.key})">
    <div class="card-head"><h3><span class="dot"></span>${o.title}</h3><span class="tag">${o.tag}</span></div>
    <p class="desc">${o.desc}</p>
    <div class="kpis">
      <div class="kpi"><div class="lbl">Dépenses par Français</div><div class="val">${eur(o.capDep)}</div><div class="tot">${md(o.dep)} au total</div></div>
      <div class="kpi"><div class="lbl">Recettes par Français</div><div class="val">${eur(o.capRec)}</div><div class="tot">${md(o.rec)} au total</div></div>
    </div>
    <div class="balance" aria-hidden="true">
      <div class="row"><span>Dépenses</span><div class="track"><div class="fill" style="width:${w(o.capDep)}"></div></div></div>
      <div class="row"><span>Recettes</span><div class="track"><div class="fill rec" style="width:${w(o.capRec)}"></div>${
        capSolde < 0 ? `<div class="gap" style="left:${w(o.capRec)};width:${w(-capSolde)}" title="Déficit"></div>` : ""}</div></div>
    </div>
    <div class="solde ${bad ? "bad" : "good"}">
      <span class="ico" aria-hidden="true">${bad ? "▼" : "≈"}</span>
      <div><div class="lbl">${bad ? "Déficit par Français" : "Quasi-équilibre : solde par Français"}</div><div class="val">${signed(eur, capSolde)}</div></div>
      <div class="pct">${signed(md, solde)}<br>${o.pibPct !== null ? `${signed(pct, o.pibPct)} du PIB` : ""}</div>
    </div>
    ${o.extra ? `<p class="note" style="margin:0">${o.extra}</p>` : ""}
  </article>`;
}

function renderCards() {
  const t = D.totals, a = D.a, b = D.b;
  const maxCap = Math.max(t.cap26.dep, t.cap26.rec, t.cap80.dep, t.cap80.rec);
  const cards = [
    { key: "--c80", title: "1980", tag: "en € de 2026",
      desc: `Le budget réel de 1980, corrigé de l'inflation (${nf1.format(a.nationals)} millions de Français).`,
      dep: t.r80.dep, rec: t.r80.rec, capDep: t.cap80.dep, capRec: t.cap80.rec,
      pibPct: ((t.r80cur.rec - t.r80cur.dep) / a.pib) * 100, balanced: (t.r80cur.rec - t.r80cur.dep) / a.pib > -0.01,
      extra: `Soit ${md(t.r80cur.dep * a.params.franc)}`.replace("Md€", "milliards de francs") + ` de dépenses à l'époque.` },
    { key: "--c26", title: "2026", tag: "prévision annuelle",
      desc: `Lois de finances et de financement de la Sécurité sociale 2026 (${nf1.format(b.nationals)} millions de Français).`,
      dep: t.y26.dep, rec: t.y26.rec, capDep: t.cap26.dep, capRec: t.cap26.rec,
      pibPct: ((t.y26.rec - t.y26.dep) / b.pib) * 100, balanced: (t.y26.rec - t.y26.dep) / b.pib > -0.01,
      extra: `Déficit prévu sur l'année entière, toutes administrations confondues${
        b.params.solde_etat ? `, dont ${signed(md, b.params.solde_etat)} pour l'État seul` : ""}.` },
    { key: "--ctr", title: "1980 transposé en 2026", tag: "en € de 2026",
      desc: `Les dépenses et recettes de 1980 par Français, appliquées aux ${nf1.format(b.nationals)} millions de Français de 2026.`,
      dep: t.t80.dep, rec: t.t80.rec, capDep: t.cap80.dep, capRec: t.cap80.rec, pibPct: null,
      balanced: (t.r80cur.rec - t.r80cur.dep) / a.pib > -0.01 },
  ];
  document.getElementById("cards").innerHTML = cards.map((c) => card(c, maxCap)).join("");
  document.getElementById("cards-note").textContent =
    `Écart entre le budget 2026 et le budget 1980 transposé : ${signed(md, t.y26.dep - t.t80.dep)} de dépenses, ` +
    `${signed(md, t.y26.rec - t.t80.rec)} de recettes.`;
}

// ---------- Rendu : graphique par catégorie ----------
function seriesFor() {
  if (state.money === "cur") return ["r80cur", "y26"];
  return state.unit === "capita" ? ["cap80", "y26"] : ["r80", "t80", "y26"];
}

function valueOf(c, s) {
  const capita = state.unit === "capita";
  switch (s) {
    case "r80cur": return capita ? c.cap80cur : c.r80cur;
    case "cap80": return capita ? c.cap80 : c.t80;
    case "r80": return c.r80;
    case "t80": return c.t80;
    case "y26": return capita ? c.cap26 : c.y26;
  }
  return 0;
}
const fmtBare = (v) => (state.unit === "capita" ? eur(v) : md(v));
const fmtVal = (v) => (state.unit === "capita" ? parFr(eur(v)) : md(v));

// école, en mode Par Français : montant par élève de la série s, null pour les autres postes
function eleveVal(c, s) {
  const e80 = D.a.ind.eleves?.val, e26 = D.b.ind.eleves?.val;
  if (state.unit !== "capita" || c.id !== "scolaire" || !e80 || !e26) return null;
  return s === "y26" ? (c.y26 * 1e3) / e26 : ((s === "r80cur" ? c.r80cur : c.r80) * 1e3) / e80;
}
const rowVal = (c, s) => { const el = eleveVal(c, s); return el === null ? fmtVal(valueOf(c, s)) : parEl(el); };

function seriesMeta(s) {
  const cur = state.money === "cur";
  if (s === "r80cur") return { k: "--c80", label: "1980, en € de l'époque" };
  if (s === "cap80") return { k: "--c80", label: "1980 (en € de 2026)" };
  if (s === "r80") return { k: "--c80", label: "1980 réel (en € de 2026)" };
  if (s === "t80") return { k: "--ctr", label: "1980 transposé en 2026" };
  return { k: "--c26", label: cur ? "2026" : "2026" };
}

function gapOf(c) { // 2026 vs 1980 transposé, dans l'unité courante (toujours en € 2026)
  return state.unit === "capita" ? c.cap26 - c.cap80 : c.y26 - c.t80;
}

function renderChart() {
  const list = D[state.side].slice();
  const ser = seriesFor();
  if (state.sort === "v26") list.sort((x, y) => valueOf(y, "y26") - valueOf(x, "y26"));
  else if (state.sort === "gap") list.sort((x, y) => gapOf(y) - gapOf(x));
  else list.sort((x, y) => x.label.localeCompare(y.label, "fr"));
  const max = Math.max(...list.flatMap((c) => ser.map((s) => valueOf(c, s))), 1e-9);

  document.getElementById("legend").innerHTML =
    ser.map((s) => { const m = seriesMeta(s); return `<span><i style="--k:var(${m.k})"></i>${m.label}</span>`; }).join("") +
    (state.money === "cur"
      ? `<span class="hint">Montants non corrigés de l'inflation : à ne pas comparer directement.</span>`
      : state.unit === "capita" ? `<span class="hint">Par Français, 1980 et « 1980 transposé » sont identiques.</span>` : "");

  document.getElementById("detail-sub").textContent = state.side === "dep"
    ? "Les dépenses publiques réparties en 18 grandes politiques. Cliquez sur une ligne pour le détail."
    : "Les recettes publiques réparties en 8 grandes familles. Cliquez sur une ligne pour le détail.";

  document.getElementById("chart").innerHTML = list.map((c) => {
    const g = gapOf(c);
    const base = state.unit === "capita" ? c.cap80 : c.t80;
    const e26 = eleveVal(c, "y26"), e80 = eleveVal(c, "cap80");
    const ratio = e26 !== null ? e26 / e80 : base > 0 ? (state.unit === "capita" ? c.cap26 : c.y26) / base : null;
    const delta = state.money === "cur"
      ? `<div class="ratio">—</div>`
      : `<div class="main">${e26 !== null ? `${signed(eur, e26 - e80)}${UNIT_EL}` : `${signed(fmtBare, g)}${state.unit === "capita" ? UNIT_FR : ""}`}</div><div class="ratio">${ratio ? `×${nf1.format(ratio)} vs 1980` : ""}</div>`;
    return `<div role="listitem"><button class="row-btn" data-id="${esc(c.id)}" aria-label="${esc(c.label)} : voir le détail">
      <span class="cat"><span class="ico" aria-hidden="true">${c.icon}</span><span>${esc(c.label)}${e26 !== null ? NOTE_EL : ""}</span></span>
      <span class="bars">${ser.map((s) => {
        const m = seriesMeta(s), el = eleveVal(c, s);
        const len = el === null ? valueOf(c, s) / max : (el / e26) * (valueOf(c, "y26") / max);
        return `<span class="bar-line"><span class="bar" style="--k:var(${m.k});width:${barW(len, state.unit === "capita")}"></span><span class="bar-val">${rowVal(c, s)}</span></span>`;
      }).join("")}</span>
      <span class="delta">${delta}</span>
    </button></div>`;
  }).join("");
  document.querySelectorAll(".row-btn").forEach((el) => {
    el.addEventListener("click", () => openDetail(el.dataset.id));
    el.addEventListener("mousemove", (e) => showTip(e, el.dataset.id));
    el.addEventListener("mouseleave", hideTip);
  });
  if (!document.getElementById("table-wrap").hidden) renderTable();
}

// ---------- Infobulle ----------
const tip = document.getElementById("tip");
function showTip(e, id) {
  if (window.matchMedia("(hover: none)").matches) return;
  const c = D[state.side].find((x) => x.id === id);
  const ser = seriesFor();
  tip.innerHTML = `<div class="t">${c.icon} ${esc(c.label)}</div>` +
    ser.map((s) => { const m = seriesMeta(s); return `<div class="r"><span><i style="--k:var(${m.k})"></i>${m.label}</span><b>${rowVal(c, s)}</b></div>`; }).join("") +
    `<div class="h">Cliquez pour le détail et les sources</div>`;
  tip.hidden = false;
  const w = tip.offsetWidth, h = tip.offsetHeight;
  let x = e.clientX + 16, y = e.clientY + 16;
  if (x + w > innerWidth - 8) x = e.clientX - w - 16;
  if (y + h > innerHeight - 8) y = e.clientY - h - 16;
  tip.style.left = `${x}px`; tip.style.top = `${y}px`;
}
function hideTip() { tip.hidden = true; }

// ---------- Détail (dialogue) ----------
function openDetail(id) {
  hideTip();
  const c = D[state.side].find((x) => x.id === id);
  const total80 = D.totals.r80cur[state.side];
  const total26 = D.totals.y26[state.side];
  document.getElementById("dlg-body").innerHTML = `
    <h3 id="dlg-title"><span aria-hidden="true">${c.icon}</span>${esc(c.label)}</h3>
    <p class="d">${esc(c.desc)}</p>
    <div class="tbl-scroll"><table>
      <thead><tr><th></th><th>Total</th><th>Par Français</th><th>Part du budget</th></tr></thead>
      <tbody>
        <tr><td><span class="dot" style="--key:var(--c80);display:inline-block"></span> 1980 (€ de l'époque)</td><td>${md(c.r80cur)}</td><td>${eurExact(Math.round(c.cap80cur))}</td><td rowspan="2">${pct((c.r80cur / total80) * 100)}</td></tr>
        <tr><td>1980 (€ de 2026)</td><td>${md(c.r80)}</td><td>${eur(c.cap80)}</td></tr>
        <tr><td><span class="dot" style="--key:var(--ctr);display:inline-block"></span> 1980 transposé en 2026</td><td>${md(c.t80)}</td><td>${eur(c.cap80)}</td><td>${pct((c.r80cur / total80) * 100)}</td></tr>
        <tr><td><span class="dot" style="--key:var(--c26);display:inline-block"></span> 2026</td><td>${md(c.y26)}</td><td>${eur(c.cap26)}</td><td>${pct((c.y26 / total26) * 100)}</td></tr>
      </tbody>
      <tfoot><tr><td>Écart 2026 − 1980 transposé</td><td>${signed(md, c.y26 - c.t80)}</td><td>${signed(eur, c.cap26 - c.cap80)}</td><td></td></tr></tfoot>
    </table></div>
    <p class="srcnote"><span class="fiab" title="Fiabilité de la donnée 1980">1980 · ${FIAB_ICON[c.fiab80] || ""} fiabilité ${esc(c.fiab80)}</span> ${esc(c.src80)}</p>
    <p class="srcnote"><span class="fiab" title="Fiabilité de la donnée 2026">2026 · ${FIAB_ICON[c.fiab26] || ""} fiabilité ${esc(c.fiab26)}</span> ${esc(c.src26)}</p>`;
  document.getElementById("dlg").showModal();
}
document.getElementById("dlg").addEventListener("click", (e) => { if (e.target.id === "dlg") e.target.close(); });

// ---------- Tableau ----------
function renderTable() {
  const list = D[state.side];
  const t = (k) => list.reduce((s, c) => s + c[k], 0);
  const cols = [
    ["1980 (Md€ de l'époque)", "r80cur", md], ["1980 (Md€ 2026)", "r80", md], ["1980 transposé (Md€)", "t80", md],
    ["2026 (Md€)", "y26", md], ["1980 par Français (€ 2026)", "cap80", eur], ["2026 par Français", "cap26", eur],
  ];
  document.getElementById("table-wrap").innerHTML = `<div class="tbl-scroll"><table>
    <thead><tr><th>${state.side === "dep" ? "Dépense" : "Recette"}</th>${cols.map((c) => `<th>${c[0]}</th>`).join("")}<th>Écart par Français</th></tr></thead>
    <tbody>${list.map((c) => `<tr><td>${c.icon} ${esc(c.label)}</td>${cols.map((k) => `<td>${k[2](c[k[1]])}</td>`).join("")}<td>${signed(eur, c.cap26 - c.cap80)}</td></tr>`).join("")}</tbody>
    <tfoot><tr><td>Total</td>${cols.map((k) => `<td>${k[2](t(k[1]))}</td>`).join("")}<td>${signed(eur, t("cap26") - t("cap80"))}</td></tr></tfoot>
  </table></div>`;
}
document.getElementById("btn-table").addEventListener("click", (e) => {
  const w = document.getElementById("table-wrap");
  w.hidden = !w.hidden;
  e.target.setAttribute("aria-expanded", String(!w.hidden));
  e.target.textContent = w.hidden ? "Voir sous forme de tableau" : "Masquer le tableau";
  if (!w.hidden) renderTable();
});

// ---------- Contexte ----------
function renderContext() {
  const a = D.a, b = D.b, t = D.totals;
  const depPib80 = (t.r80cur.dep / a.pib) * 100, depPib26 = (t.y26.dep / b.pib) * 100;
  const pibCap80 = (a.pib * a.coef * 1e3) / a.nationals, pibCap26 = (b.pib * 1e3) / b.nationals;
  const tiles = [
    ["Dépenses publiques", pct(depPib80), pct(depPib26), "en % du PIB",
      `À part de PIB égale, le budget 1980 transposé représenterait ${md((depPib80 / 100) * b.pib)} de dépenses en 2026.`],
    ["Recettes publiques", pct((t.r80cur.rec / a.pib) * 100), pct((t.y26.rec / b.pib) * 100), "en % du PIB", "Impôts, cotisations et autres recettes."],
    ["Dette publique", pct(a.params.dette_pib), pct(b.params.dette_pib), "en % du PIB",
      "Les intérêts de la dette coûtent aujourd'hui plus que la défense."],
    ["Richesse produite par Français", eur(pibCap80), eur(pibCap26), "PIB par Français, € de 2026",
      `Elle a été multipliée par ${nf1.format(pibCap26 / pibCap80)} : les Français de 2026 sont plus riches, ce qui permet un budget plus élevé.`],
    ["Nombre de Français", `${nf1.format(a.nationals)} M`, `${nf1.format(b.nationals)} M`, "personnes de nationalité française",
      `Population totale : ${nf1.format(a.params.pop)} M en 1980, ${nf1.format(b.params.pop)} M en 2026.`],
  ];
  document.getElementById("ctx").innerHTML = tiles.map(([l, v1, v2, u, p]) => `
    <div class="ctx"><div class="lbl">${l}</div>
      <div class="pair"><div><small>1980</small><b>${v1}</b></div><span class="arrow">→</span><div><small>2026</small><b>${v2}</b></div></div>
      <p>${u}. ${p}</p></div>`).join("");
}

// ---------- Indicateurs transversaux ----------
function renderIndicators() {
  const a = D.a, b = D.b;
  const keys = [...new Set([...Object.keys(b.ind), ...Object.keys(a.ind)])];
  const num = (v, unit) => (v === null || v === undefined || Number.isNaN(v) ? null
    : unit.startsWith("%") ? (v < 0 ? signed(pct, v) : pct(v)) : unit === "millions" ? `${v < 1 ? v.toLocaleString("fr-FR", { minimumFractionDigits: 2 }) : nf1.format(v)} M` : nf1.format(v));
  const tile = (label, v1, v2, unit, note, est) => ({ label, v1, v2, unit, note, est });
  const tiles = keys.map((k) => {
    const i80 = a.ind[k] || {}, i26 = b.ind[k] || {}, unit = i26.unit || i80.unit || "";
    const note = [i80.val !== null && i80.note ? `1980 : ${i80.note}` : "", i26.note ? `2026 : ${i26.note}` : ""].filter(Boolean).join(" ");
    return { k, ...tile(i26.label || i80.label, num(i80.val, unit), num(i26.val, unit), unit, note, i80.fiab === "faible") };
  });
  // Indicateurs calculés à partir des deux feuilles
  const v = (y, k) => (y.ind[k] ? y.ind[k].val : null);
  const ag80 = (v(a, "agents_fp") || 0) + (v(a, "agents_secu") || 0), ag26 = (v(b, "agents_fp") || 0) + (v(b, "agents_secu") || 0);
  if (ag80 && ag26) {
    tiles.unshift(tile("Agents publics pour 1 000 Français", nf0.format((ag80 / a.nationals) * 1e3), nf0.format((ag26 / b.nationals) * 1e3),
      "fonctions publiques et Sécurité sociale", `Total : ${nf1.format(ag80)} M en 1980, ${nf1.format(ag26)} M en 2026.`, a.ind.agents_secu?.fiab === "faible"));
  }
  const r = D.dep.find((c) => c.id === "retraites"), n80 = v(a, "retraites_nb"), n26 = v(b, "retraites_nb");
  if (r && n80 && n26) {
    const m80 = (r.r80 * 1e3) / n80 / 12, m26 = (r.y26 * 1e3) / n26 / 12;
    tiles.splice(tiles.findIndex((t) => t.k === "retraites_nb") + 1, 0,
      tile("Dépense de retraite par retraité", eur(m80), eur(m26), "par mois, € de 2026",
        `Pensions, réversion et minimum vieillesse divisés par le nombre de retraités : +${nf0.format((m26 / m80 - 1) * 100)} % en euros constants, ` +
        `quand le nombre de retraités a été multiplié par ${nf1.format(n26 / n80)}.`, a.ind.retraites_nb?.fiab === "faible"));
  }
  const c80 = v(a, "cotisants_nb"), c26 = v(b, "cotisants_nb");
  if (c80 && c26 && n80 && n26) {
    tiles.splice(tiles.findIndex((t) => t.k === "cotisants_nb") + 1, 0,
      tile("Cotisants par retraité", nf1.format(c80 / n80), nf1.format(c26 / n26), "ratio",
        "Moins de cotisants pour financer chaque retraite.", a.ind.cotisants_nb?.fiab === "faible"));
  }
  if (a.params.etrangers && b.params.etrangers) {
    tiles.splice(tiles.findIndex((t) => t.k === "immigres"), 0,
      tile("Étrangers (non Français)", `${nf1.format(a.params.etrangers)} M`, `${nf1.format(b.params.etrangers)} M`, "millions",
        `${pct((a.params.etrangers / a.params.pop) * 100)} de la population en 1980, ${pct((b.params.etrangers / b.params.pop) * 100)} en 2026. ` +
        "Les montants « par Français » incluent les dépenses pour les étrangers résidents.", false));
  }
  const dette = D.dep.find((c) => c.id === "dette");
  if (dette) tiles.push(tile("Intérêts de la dette", pct((dette.r80cur / D.totals.r80cur.rec) * 100), pct((dette.y26 / D.totals.y26.rec) * 100),
    "en % des recettes publiques", "Part des impôts et cotisations absorbée par la charge de la dette.", false));
  document.getElementById("ind").innerHTML = tiles.map((t) => `
    <div class="ctx"><div class="lbl">${esc(t.label)}</div>
      <div class="pair"><div><small>1980</small><b${t.v1 ? "" : ' class="na"'}>${t.v1 || "n.d."}</b></div>
        <span class="arrow">→</span><div><small>2026</small><b${t.v2 ? "" : ' class="na"'}>${t.v2 || "n.d."}</b></div></div>
      ${t.est ? `<span class="est">1980 : estimation</span>` : ""}
      <p>${esc(t.unit)}. ${esc(t.note)}</p></div>`).join("");
}

// ---------- Sources ----------
function renderSources() {
  const seen = new Set();
  const src = [...D.a.sources, ...D.b.sources].filter((r) => { const k = r[0] + r[1]; if (seen.has(k)) return false; seen.add(k); return true; });
  document.getElementById("method").innerHTML =
    `<p><strong>1980.</strong> ${D.a.method.map(esc).join(" ")}</p>` +
    `<p><strong>2026.</strong> ${D.b.method.map(esc).join(" ")}</p>` +
    `<ul>${src.map((r) => `<li><strong>${esc(r[0])}</strong> : ${r[2] ? `<a href="${esc(r[2])}" target="_blank" rel="noopener">${esc(r[1])}</a>` : esc(r[1])}</li>`).join("")}</ul>`;
  document.getElementById("dl80").href = DATA_URLS.y80;
  document.getElementById("dl26").href = DATA_URLS.y26;
}

// ---------- Contrôles ----------
function syncControls() {
  document.querySelectorAll(".seg").forEach((g) => {
    g.querySelectorAll("button").forEach((b) => b.setAttribute("aria-checked", String(state[g.dataset.key] === b.dataset.v)));
  });
  document.getElementById("sort").value = state.sort;
}
function save() { try { localStorage.setItem("b8026", JSON.stringify(state)); } catch (e) { /* ignoré */ } }

document.querySelectorAll(".seg").forEach((g) => {
  g.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    state[g.dataset.key] = b.dataset.v;
    syncControls(); save();
    if (g.dataset.key === "sSide" || g.dataset.key === "sUnit") renderSauce();
    else if (g.dataset.key === "gSide" || g.dataset.key === "gUnit") renderGame();
    else renderChart();
  });
  g.addEventListener("keydown", (e) => {
    if (!["ArrowLeft", "ArrowRight"].includes(e.key)) return;
    const btns = [...g.querySelectorAll("button")];
    const i = btns.findIndex((b) => b.dataset.v === state[g.dataset.key]);
    const n = btns[(i + (e.key === "ArrowRight" ? 1 : btns.length - 1)) % btns.length];
    n.focus(); n.click();
  });
});
document.getElementById("sort").addEventListener("change", (e) => { state.sort = e.target.value; save(); renderChart(); });

// ---------- Démarrage ----------
(async function init() {
  syncControls();
  try {
    const [w80, w26] = await Promise.all([loadWorkbook(DATA_URLS.y80), loadWorkbook(DATA_URLS.y26)]);
    D = compute(readYear(w80), readYear(w26));
    computeSauce();
    setupGame();
    renderLede(); renderCards(); renderChart(); renderSauce(); renderGame(); renderContext(); renderIndicators(); renderSources();
  } catch (err) {
    console.error(err);
    document.getElementById("lede").textContent = `Impossible de charger les fichiers Excel (${err.message}).`;
  }
})();
