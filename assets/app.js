/* Planten Databank – client-side (GitHub Pages)
   Update: foto's via kolom 'foto_ids' (pipe-separated) + carousel in drawer.
   - Foto's via Cloudflare R2 CDN
   - Klik op Nederlandse naam opent drawer met foto's
   - Fyto: 3 extra kolommen in tabel wanneer Fyto-laag actief
   - Layer matching: Latijnse naam (primair) + Nederlandse naam (fallback)
*/
"use strict";

const \$ = (sel) => document.querySelector(sel);

// ─── CONSTANTEN ───────────────────────────────────────────────────────────────
const IMAGE_BASE_URL = "https://pub-bb204453b9b642598d8514f7ac4f68be.r2.dev";

// ─── STATE ────────────────────────────────────────────────────────────────────
const STATE = {
  config: null,
  plants: [],
  loaded: { typologies: new Map(), layers: new Map() },
  options: {
    bobo:   { groups: [], codesByGroup: new Map() },
    region: { districts: [] }
  },
  selected: {
    typology:     null,
    subtype:      null,
    soilType:     "ALLE",
    soilMoisture: "ALLE",
    acidity:      "ALLE",
    spread:       "ALLE",
    layers: {
      klimaat:   false,
      amber:     false,
      regionaal: false,
      bobo:      false,
      fyto:      false
    },
    district:    "",
    regionMode:  "paint",
    boboGroup:   "",
    boboCode:    "ALLE",
    facets:      new Map()
  },
  table: { extraCols: [] }
};

// ─── HULPFUNCTIES ─────────────────────────────────────────────────────────────
function norm(s) {
  return String(s ?? "").replace(/\u00a0/g, " ").trim().replace(/\s+/g, " ");
}
function keyLatin(s)  { return norm(s).toLowerCase(); }
function keyDutch(s)  { return "nl:" + norm(s).toLowerCase(); }

function splitList(s) {
  const t = norm(s);
  if (!t) return [];
  return t.split(/[,|]/).map(x => norm(x).toLowerCase()).filter(Boolean);
}
function splitPipesRaw(s) {
  const t = norm(s);
  if (!t) return [];
  return t.split("|").map(x => norm(x)).filter(Boolean);
}
function uniqSorted(arr) {
  return Array.from(new Set(arr)).sort((a, b) => a.localeCompare(b));
}

// ─── XLSX HELPERS ─────────────────────────────────────────────────────────────
async function fetchXlsx(path) {
  const res = await fetch(path);
  if (!res.ok) throw new Error(`Fetch failed: \${path} (\${res.status})`);
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, { type: "array" });
}
function firstSheetName(wb)        { return wb.SheetNames[0]; }
function sheetToJson(wb, sheetName) {
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
}

// ─── CONFIG ───────────────────────────────────────────────────────────────────
async function loadConfig() {
  const res = await fetch("data/config.json");
  if (!res.ok) throw new Error("Missing data/config.json");
  STATE.config = await res.json();
}

// ─── TYPOLOGIE OPTIES ─────────────────────────────────────────────────────────
function typologyOptions() {
  if (STATE.config?.typologies) {
    return Object.keys(STATE.config.typologies);
  }
  return [
    "1.Bomen",
    "2.Haagplanten",
    "3.wadi",
    "4.bloemenweide",
    "5.inheemse_planten",
    "6.gevelgroen",
    "7.grasland_weide",
    "9.fruit_groenten_kruiden"
  ];
}
function subtypeOptions(typology) {
  const node = STATE.config?.typologies?.[typology];
  const subs  = node?.subtypes ? Object.keys(node.subtypes) : ["Alle"];
  return subs.length ? subs : ["Alle"];
}
function filesForTypology(typology, subtype) {
  const node = STATE.config?.typologies?.[typology];
  if (!node) return [];
  const subtypeKeys = Object.keys(node.subtypes || {});
  const sub = node.subtypes?.[subtype]
           ?? node.subtypes?.[subtypeKeys[0]]
           ?? [];
  return (sub ?? []).map(norm).filter(Boolean).map(n => `data/typologies/\${n}.xlsx`);
}

// ─── LAYER HELPERS ────────────────────────────────────────────────────────────
function districtToFilename(d) {
  const cleaned = norm(d)
    .replaceAll("/", "_")
    .replaceAll(" ", " ")
    .replaceAll(".", "");
  return cleaned.split(" ").join("_") + ".xlsx";
}
function layerFiles(layerKey, typology) {
  if (layerKey === "regionaal") {
    if (!STATE.selected.district) return [];
    return [`data/layers/\${districtToFilename(STATE.selected.district)}`];
  }
  if (layerKey === "bobo") {
    if (!STATE.selected.boboGroup) return [];
    return [`data/layers/BOBO_\${STATE.selected.boboGroup}.xlsx`];
  }
  if (layerKey === "fyto") {
    const names = (STATE.config?.layers?.fytoremediatie?.[typology] ?? [])
      .map(norm).filter(Boolean);
    return names.map(n => `data/layers/\${n}.xlsx`);
  }
  const keyMap = { klimaat: "klimaatbomenlijst", amber: "amberlijst" };
  const names  = (STATE.config?.layers?.[keyMap[layerKey]] ?? [])
    .map(norm).filter(Boolean);
  return names.map(n => `data/layers/\${n}.xlsx`);
}

// ─── ROW PARSERS ──────────────────────────────────────────────────────────────
function rowLatin(row) {
  return norm(
    row["Latijnse naam"]  ||
    row["Latijnse naam "] ||
    row["latijnse naam"]  ||
    row["LatijnseNaam"]   ||
    row["latin"]          ||
    ""
  );
}
function rowDutch(row) {
  return norm(
    row["Nederlandse naam"]  ||
    row["Nederlandse naam "] ||
    row["nederlandse naam"]  ||
    row["NederlandseNaam"]   ||
    row["dutch"]             ||
    ""
  );
}

const BASE_FILTER_KEYS = new Set([
  "kenmerk_bodemtype",
  "kenmerk_bodemvochtigheid",
  "kenmerk_zuurtegraad",
  "kenmerk_verspreiding"
]);

function plantFromRow(row) {
  const latin = rowLatin(row);
  const dutch = rowDutch(row);
  const traits = {
    bodemtype:    splitList(row["kenmerk_bodemtype"]),
    bodemvocht:   splitList(row["kenmerk_bodemvochtigheid"]),
    zuur:         splitList(row["kenmerk_zuurtegraad"]),
    verspreiding: splitList(row["kenmerk_verspreiding"]),
  };
  const dynamic = new Map();
  for (const [k, v] of Object.entries(row)) {
    const kk = norm(k);
    if (!kk.toLowerCase().startsWith("kenmerk_")) continue;
    const kLower = kk.toLowerCase();
    if (BASE_FILTER_KEYS.has(kLower)) continue;
    const vals = splitList(v);
    if (vals.length) dynamic.set(kLower, vals);
  }
  const fotoIds = splitPipesRaw(
    row["foto_ids"] || row["foto_id"] || row["foto"] || ""
  );
  return {
    latin, dutch,
    traits, dynamic,
    fotoIds,
    layers: { klimaat: false, amber: false, regionaal: false, bobo: false, fyto: false },
    boboCode: null,
    fytoRow:  null,
    raw: row
  };
}

// ─── DATA LADEN ───────────────────────────────────────────────────────────────
async function loadTypologyPlants(typology, subtype) {
  const files = filesForTypology(typology, subtype);
  if (!files.length) return [];
  const all = [];
  for (const file of files) {
    if (STATE.loaded.typologies.has(file)) {
      all.push(...STATE.loaded.typologies.get(file));
      continue;
    }
    try {
      const wb     = await fetchXlsx(file);
      const rows   = sheetToJson(wb, firstSheetName(wb));
      const plants = rows.map(plantFromRow).filter(p => p.latin);
      STATE.loaded.typologies.set(file, plants);
      all.push(...plants);
    } catch (e) {
      console.warn(`Kon bestand niet laden: \${file}`, e);
    }
  }
  // Dedupliceer op Latijnse naam
  const merged = new Map();
  for (const p of all) {
    const k = keyLatin(p.latin);
    if (!merged.has(k)) {
      merged.set(k, p);
    } else {
      const ex = merged.get(k);
      if (!ex.dutch && p.dutch) ex.dutch = p.dutch;
      ex.traits.bodemtype    = uniqSorted([...ex.traits.bodemtype,    ...p.traits.bodemtype]);
      ex.traits.bodemvocht   = uniqSorted([...ex.traits.bodemvocht,   ...p.traits.bodemvocht]);
      ex.traits.zuur         = uniqSorted([...ex.traits.zuur,         ...p.traits.zuur]);
      ex.traits.verspreiding = uniqSorted([...ex.traits.verspreiding, ...p.traits.verspreiding]);
      for (const [dk, vals] of p.dynamic.entries()) {
        const cur = ex.dynamic.get(dk) || [];
        ex.dynamic.set(dk, uniqSorted([...cur, ...vals]));
      }
      ex.fotoIds = Array.from(new Set([...(ex.fotoIds || []), ...(p.fotoIds || [])]));
    }
  }
  return Array.from(merged.values());
}

async function loadLayerIndex(files) {
  const out = new Map();
  for (const file of files) {
    if (STATE.loaded.layers.has(file)) {
      for (const [k, v] of STATE.loaded.layers.get(file).entries()) out.set(k, v);
      continue;
    }
    try {
      const wb   = await fetchXlsx(file);
      const rows = sheetToJson(wb, firstSheetName(wb));
      const map  = new Map();
      for (const r of rows) {
        const latin = rowLatin(r);
        const dutch = rowDutch(r);
        if (latin) map.set(keyLatin(latin), r);
        if (dutch) map.set(keyDutch(dutch), r);
      }
      STATE.loaded.layers.set(file, map);
      for (const [k, v] of map.entries()) out.set(k, v);
    } catch (e) {
      console.warn(`Layer bestand niet gevonden: \${file}`, e);
    }
  }
  return out;
}

// ─── SELECT HELPERS ───────────────────────────────────────────────────────────
function fillSelect(selectEl, values, includeAll = true) {
  if (!selectEl) return;
  const prev = selectEl.value;
  selectEl.innerHTML = "";
  const add = (v, label) => {
    const opt = document.createElement("option");
    opt.value       = v;
    opt.textContent = label ?? v;
    selectEl.appendChild(opt);
  };
  if (includeAll) add("ALLE", "Alle");
  for (const v of values) add(v, v);
  const cand = prev || (includeAll ? "ALLE" : (values[0] ?? ""));
  if (Array.from(selectEl.options).some(o => o.value === cand)) {
    selectEl.value = cand;
  }
}

// ─── FILTER OPTIES BEREKENEN ──────────────────────────────────────────────────
function computeBaseFilterOptions(plants) {
  const soils = [], moist = [], acid = [], spread = [];
  for (const p of plants) {
    soils.push(...p.traits.bodemtype);
    moist.push(...p.traits.bodemvocht);
    acid.push(...p.traits.zuur);
    spread.push(...p.traits.verspreiding);
  }
  return {
    soil:   uniqSorted(soils),
    moist:  uniqSorted(moist),
    acid:   uniqSorted(acid),
    spread: uniqSorted(spread)
  };
}
function computeFacetOptions(plants) {
  const counts = new Map();
  for (const p of plants) {
    for (const [k, vals] of p.dynamic.entries()) {
      if (!counts.has(k)) counts.set(k, new Map());
      const m = counts.get(k);
      for (const v of vals) m.set(v, (m.get(v) || 0) + 1);
    }
  }
  return counts;
}

// ─── FILTER MATCHING ──────────────────────────────────────────────────────────
function matchesBaseFilters(p) {
  const st = STATE.selected.soilType.toLowerCase();
  const sm = STATE.selected.soilMoisture.toLowerCase();
  const ac = STATE.selected.acidity.toLowerCase();
  const sp = STATE.selected.spread.toLowerCase();
  return (
    (st === "alle" || p.traits.bodemtype.includes(st))  &&
    (sm === "alle" || p.traits.bodemvocht.includes(sm)) &&
    (ac === "alle" || p.traits.zuur.includes(ac))       &&
    (sp === "alle" || p.traits.verspreiding.includes(sp))
  );
}
function matchesFacets(p) {
  for (const [k, set] of STATE.selected.facets.entries()) {
    if (!set || set.size === 0) continue;
    const vals = p.dynamic.get(k) || [];
    if (!vals.some(v => set.has(v))) return false;
  }
  return true;
}
function matchesRegionMode(p) {
  if (!STATE.selected.layers.regionaal) return true;
  if (STATE.selected.regionMode !== "filter") return true;
  return p.layers.regionaal === true;
}
function matchesAll(p) {
  return matchesBaseFilters(p) && matchesFacets(p) && matchesRegionMode(p);
}

// ─── LAGEN TOEPASSEN ──────────────────────────────────────────────────────────
function badgesForPlant(p) {
  const b = [];
  if (p.layers.klimaat)   b.push("Klimaat");
  if (p.layers.amber)     b.push("AMBER");
  if (p.layers.regionaal) b.push("Regionaal");
  if (p.layers.bobo)      b.push("BOBO");
  if (p.layers.fyto)      b.push("Fyto");
  return b;
}
function pickAny(row, keys) {
  for (const k of keys) {
    if (row && row[k] !== undefined && row[k] !== null &&
        String(row[k]).trim() !== "") {
      return norm(row[k]);
    }
  }
  return "";
}
function matchLayerRow(layerIndex, plant) {
  const k1 = keyLatin(plant.latin);
  if (layerIndex.has(k1)) return layerIndex.get(k1);
  if (plant.dutch) {
    const k2 = keyDutch(plant.dutch);
    if (layerIndex.has(k2)) return layerIndex.get(k2);
  }
  return null;
}

function buildFytoColumns() {
  STATE.table.extraCols = [
    {
      label: "Comments on phytoremedial effectiveness",
      getter: (p) => pickAny(p.fytoRow, ["Comments on phytoremedial effectiveness"])
    },
    {
      label: "Continent-Country-City-Site",
      getter: (p) => pickAny(p.fytoRow, ["Continent-Country-City-Site"])
    },
    {
      label: "Reference (author, year, doi)",
      getter: (p) => pickAny(p.fytoRow, [
        "Reference (author, year, doi)",
        "Reference (author, year)",
        "Reference"
      ])
    }
  ];
}
function clearExtraColumns() { STATE.table.extraCols = []; }

async function applyLayers() {
  for (const p of STATE.plants) {
    p.layers  = { klimaat: false, amber: false, regionaal: false, bobo: false, fyto: false };
    p.fytoRow = null;
    p.boboCode = null;
  }
  clearExtraColumns();
  const typ = STATE.selected.typology;

  if (STATE.selected.layers.klimaat) {
    const idx = await loadLayerIndex(layerFiles("klimaat", typ));
    for (const p of STATE.plants) {
      if (matchLayerRow(idx, p)) p.layers.klimaat = true;
    }
  }
  if (STATE.selected.layers.amber) {
    const idx = await loadLayerIndex(layerFiles("amber", typ));
    for (const p of STATE.plants) {
      if (matchLayerRow(idx, p)) p.layers.amber = true;
    }
  }
  if (STATE.selected.layers.regionaal) {
    const idx = await loadLayerIndex(layerFiles("regionaal", typ));
    for (const p of STATE.plants) {
      if (matchLayerRow(idx, p)) p.layers.regionaal = true;
    }
  }
  if (STATE.selected.layers.bobo) {
    const idx    = await loadLayerIndex(layerFiles("bobo", typ));
    const wanted = STATE.selected.boboCode;
    for (const p of STATE.plants) {
      const row = matchLayerRow(idx, p);
      if (!row) continue;
      const code = pickAny(row, [
        "code", "Code", "bodemcode", "Bodemcode",
        "BOBO", "bobo_code", "BOBO code", "bobo"
      ]);
      if (wanted && wanted !== "ALLE") {
        if (code && norm(code).toLowerCase() === norm(wanted).toLowerCase()) {
          p.layers.bobo = true;
          p.boboCode    = code || wanted;
        }
      } else {
        p.layers.bobo = true;
        p.boboCode    = code || null;
      }
    }
  }
  if (STATE.selected.layers.fyto) {
    const idx = await loadLayerIndex(layerFiles("fyto", typ));
    for (const p of STATE.plants) {
      const row = matchLayerRow(idx, p);
      if (!row) continue;
      p.layers.fyto = true;
      p.fytoRow     = row;
    }
    buildFytoColumns();
  }
}

// ─── TABEL ────────────────────────────────────────────────────────────────────
function rebuildTableHeader() {
  const thead = \$("#results thead");
  if (!thead) return;
  const base = ["Latijnse naam", "Nederlandse naam", "Bodemtype", "Vocht", "pH", "Lagen"];
  const cols = [...base, ...STATE.table.extraCols.map(c => c.label)];
  thead.innerHTML = "";
  const tr = document.createElement("tr");
  for (const label of cols) {
    const th = document.createElement("th");
    th.textContent = label;
    tr.appendChild(th);
  }
  thead.appendChild(tr);
}

function render(plants) {
  rebuildTableHeader();
  const tbody   = \$("#results tbody");
  tbody.innerHTML = "";
  const filtered = plants.filter(matchesAll);
  \$("#resultsMeta").textContent = `\${filtered.length} resultaten`;

  for (const p of filtered) {
    const tr = document.createElement("tr");

    // Latijnse naam
    const tdLatin = document.createElement("td");
    tdLatin.textContent = p.latin;
    tr.appendChild(tdLatin);

    // Nederlandse naam (klikbaar → drawer)
    const tdDutch = document.createElement("td");
    tdDutch.textContent = p.dutch || "—";
    tdDutch.className   = "dutchCell";
    tdDutch.addEventListener("click", (e) => {
      e.stopPropagation();
      openDrawer(p);
    });
    tr.appendChild(tdDutch);

    // Bodem / Vocht / pH
    const tdSoil = document.createElement("td");
    tdSoil.textContent = (p.traits.bodemtype  || []).join(", ") || "—";
    tr.appendChild(tdSoil);

    const tdMoist = document.createElement("td");
    tdMoist.textContent = (p.traits.bodemvocht || []).join(", ") || "—";
    tr.appendChild(tdMoist);

    const tdPh = document.createElement("td");
    tdPh.textContent = (p.traits.zuur || []).join(", ") || "—";
    tr.appendChild(tdPh);

    // Lagen badges
    const tdBadges = document.createElement("td");
    const wrap     = document.createElement("div");
    wrap.className = "badges";
    for (const t of badgesForPlant(p)) {
      const s = document.createElement("span");
      s.className   = "badge ok";
      s.textContent = t;
      wrap.appendChild(s);
    }
    if (p.layers.bobo && p.boboCode) {
      const s = document.createElement("span");
      s.className   = "badge";
      s.textContent = p.boboCode;
      wrap.appendChild(s);
    }
    tdBadges.appendChild(wrap);
    tr.appendChild(tdBadges);

    // Extra kolommen (bv. fyto)
    for (const c of STATE.table.extraCols) {
      const td = document.createElement("td");
      const v  = c.getter(p);
      td.textContent = v ? String(v) : "—";
      tr.appendChild(td);
    }

    tr.addEventListener("click", () => openDrawer(p));
    tbody.appendChild(tr);
  }
}

// ─── FOTO / CAROUSEL ──────────────────────────────────────────────────────────
function urlForFotoId(id, ext) {
  return `\${IMAGE_BASE_URL}/\${id}.\${ext}`;
}
function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload  = () => resolve(url);
    img.onerror = () => reject(new Error("not found"));
    img.src = url;
  });
}
async function resolveFotoUrls(fotoIds, max = 80) {
  const out = [];
  for (const id of fotoIds) {
    let okUrl = null;
    try {
      okUrl = await loadImage(urlForFotoId(id, "jpg"));
    } catch {
      try {
        okUrl = await loadImage(urlForFotoId(id, "png"));
      } catch {
        okUrl = null;
      }
    }
    if (okUrl) {
      out.push(okUrl);
      if (out.length >= max) break;
    }
  }
  return out;
}

// ─── DRAWER ───────────────────────────────────────────────────────────────────
function openDrawer(plant) {
  const drawer = \$("#detailDrawer");
  if (!drawer) return;

  \$("#drawerTitle").textContent  = plant.latin;
  \$("#drawerSub").textContent    = plant.dutch || "";
  \$("#drawerSoil").textContent   = (plant.traits.bodemtype    || []).join(", ") || "—";
  \$("#drawerMoist").textContent  = (plant.traits.bodemvocht   || []).join(", ") || "—";
  \$("#drawerPh").textContent     = (plant.traits.zuur         || []).join(", ") || "—";
  \$("#drawerSpread").textContent = (plant.traits.verspreiding || []).join(", ") || "—";

  // Fyto sectie
  const fytoBox = \$("#drawerFyto");
  if (STATE.selected.layers.fyto && plant.fytoRow) {
    fytoBox.style.display = "block";
    \$("#fytoComments").textContent = pickAny(plant.fytoRow,
      ["Comments on phytoremedial effectiveness"]) || "—";
    \$("#fytoSite").textContent = pickAny(plant.fytoRow,
      ["Continent-Country-City-Site"]) || "—";
    \$("#fytoRef").textContent  = pickAny(plant.fytoRow,
      ["Reference (author, year, doi)", "Reference (author, year)", "Reference"]) || "—";
  } else {
    fytoBox.style.display = "none";
  }

  // Foto carousel
  const imgHost = \$("#drawerImages");
  imgHost.innerHTML = `<div class="hint">Foto's laden…</div>`;

  const ids = plant.fotoIds || [];
  if (!ids.length) {
    imgHost.innerHTML = `<div class="hint">Geen foto_ids voor deze plant.</div>`;
  } else {
    resolveFotoUrls(ids, 80).then(urls => {
      if (\$("#drawerTitle").textContent !== plant.latin) return;
      if (!urls.length) {
        imgHost.innerHTML = `<div class="hint">Geen foto's gevonden voor deze plant.</div>`;
        return;
      }
      let idx = 0;
      function renderCarousel() {
        imgHost.innerHTML = `
          <div class="carousel">
            <img class="carouselMain"
                 src="\${urls[idx]}"
                 alt="foto \${idx + 1} van \${urls.length}">
            <button class="carouselBtn prev" aria-label="vorige foto">&#8249;</button>
            <button class="carouselBtn next" aria-label="volgende foto">&#8250;</button>
            <div class="carouselCounter">\${idx + 1} / \${urls.length}</div>
          </div>
        `;
        imgHost.querySelector(".carouselBtn.prev").addEventListener("click", (e) => {
          e.stopPropagation();
          idx = (idx - 1 + urls.length) % urls.length;
          renderCarousel();
        });
        imgHost.querySelector(".carouselBtn.next").addEventListener("click", (e) => {
          e.stopPropagation();
          idx = (idx + 1) % urls.length;
          renderCarousel();
        });
      }
      renderCarousel();
    });
  }

  drawer.classList.add("open");
  drawer.setAttribute("aria-hidden", "false");
}

function closeDrawer() {
  const drawer = \$("#detailDrawer");
  if (!drawer) return;
  drawer.classList.remove("open");
  drawer.setAttribute("aria-hidden", "true");
  const imgHost = \$("#drawerImages");
  if (imgHost) imgHost.innerHTML = "";
}

// ─── CSV EXPORT ───────────────────────────────────────────────────────────────
function toCsv(rows) {
  const esc = (v) => `"\${String(v ?? "").replaceAll('"', '""')}"`;
  const headers = [
    "Latijnse naam", "Nederlandse naam", "Bodemtype", "Vocht",
    "pH", "Verspreiding", "Lagen",
    ...STATE.table.extraCols.map(c => c.label)
  ];
  return [
    headers.map(esc).join(","),
    ...rows.map(p => {
      const layers = badgesForPlant(p).join("|");
      const base   = [
        p.latin,
        p.dutch,
        (p.traits.bodemtype    || []).join("|"),
        (p.traits.bodemvocht   || []).join("|"),
        (p.traits.zuur         || []).join("|"),
        (p.traits.verspreiding || []).join("|"),
        layers
      ];
      const extras = STATE.table.extraCols.map(c => c.getter(p) ?? "");
      return [...base, ...extras].map(esc).join(",");
    })
  ].join("\n");
}

function downloadText(filename, text) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 250);
}

// ─── EXTRA FILTERS UI ─────────────────────────────────────────────────────────
function toggleFacetPanel(id) {
  const body = document.querySelector(`#facetBody_\${id}`);
  if (body) body.classList.toggle("open");
}
function humanizeFacetKey(k) {
  return norm(k).replace(/^kenmerk_/i, "").replaceAll("_", " ");
}

function renderExtraFilters(plants) {
  const host = \$("#extraFiltersList");
  if (!host) return;
  host.innerHTML = "";
  const counts = computeFacetOptions(plants);
  const keys   = uniqSorted(Array.from(counts.keys()));
  if (!keys.length) {
    host.innerHTML = `<div class="hint">Geen extra kenmerken beschikbaar.</div>`;
    return;
  }
  keys.forEach((k, idx) => {
    const map     = counts.get(k);
    const entries = Array.from(map.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

    const facet  = document.createElement("div");
    facet.className = "facet";

    const header = document.createElement("div");
    header.className = "facetHeader";
    header.innerHTML = `
      <div>
        <div class="facetTitle">\${humanizeFacetKey(k)}</div>
        <div class="facetMeta">\${entries.length} opties</div>
      </div>
      <div>&#9662;</div>
    `;
    header.addEventListener("click", () => toggleFacetPanel(idx));

    const body = document.createElement("div");
    body.className = "facetBody";
    body.id        = `facetBody_\${idx}`;

    const opts = document.createElement("div");
    opts.className = "facetOptions";

    const selectedSet = STATE.selected.facets.get(k) || new Set();
    for (const [val, count] of entries) {
      const safeId = `facet_\${idx}_\${val.replace(/[^a-z0-9]+/g, "_")}`;
      const wrap   = document.createElement("label");
      wrap.className = "opt";
      wrap.innerHTML = `
        <input type="checkbox" id="\${safeId}">
        <span>\${val} <span style="opacity:.7">(\${count})</span></span>
      `;
      const cb = wrap.querySelector("input");
      cb.checked = selectedSet.has(val);
      cb.addEventListener("change", () => {
        let set = STATE.selected.facets.get(k);
        if (!set) { set = new Set(); STATE.selected.facets.set(k, set); }
        if (cb.checked) set.add(val); else set.delete(val);
        if (set.size === 0) STATE.selected.facets.delete(k);
        render(STATE.plants);
      });
      opts.appendChild(wrap);
    }
    body.appendChild(opts);
    facet.appendChild(header);
    facet.appendChild(body);
    host.appendChild(facet);
  });
}

// ─── OPTIES LADEN (BOBO + REGIO) ──────────────────────────────────────────────
async function loadBoboOptions() {
  try {
    const wb    = await fetchXlsx("data/layers/bobo_bodemcodes_volledige_lijst.xlsx");
    const sheet = wb.Sheets["alle_codes"] ? "alle_codes" : firstSheetName(wb);
    const rows  = sheetToJson(wb, sheet);
    const groups = uniqSorted(rows.map(r => norm(r["groep"])).filter(Boolean));
    const codesBy = new Map();
    for (const g of groups) {
      const codes = rows
        .filter(r => norm(r["groep"]) === g)
        .map(r => norm(r["code"]))
        .filter(Boolean);
      codesBy.set(g, uniqSorted(codes));
    }
    STATE.options.bobo.groups        = groups;
    STATE.options.bobo.codesByGroup  = codesBy;
  } catch (e) {
    console.warn("BOBO opties niet geladen:", e);
    STATE.options.bobo.groups       = [];
    STATE.options.bobo.codesByGroup = new Map();
  }
}

async function loadRegionOptions() {
  try {
    const wb   = await fetchXlsx("data/layers/List_options_regionale_soortenlijst.xlsx");
    const rows = sheetToJson(wb, firstSheetName(wb));
    const districts = rows
      .map(r => norm(r["Districten"] || r["districten"]))
      .filter(Boolean);
    STATE.options.region.districts = uniqSorted(districts);
  } catch (e) {
    console.warn("Regionale opties niet geladen:", e);
    STATE.options.region.districts = [];
  }
}

// ─── HOOFD LAAD FUNCTIE ───────────────────────────────────────────────────────
async function loadTypologyAndRender() {
  const meta = \$("#resultsMeta");
  if (meta) meta.textContent = "Laden…";
  STATE.selected.facets = new Map();
  try {
    STATE.plants = await loadTypologyPlants(
      STATE.selected.typology,
      STATE.selected.subtype
    );
  } catch (e) {
    console.error("Fout bij laden planten:", e);
    if (meta) meta.textContent = "Fout bij laden data.";
    return;
  }
  const opts = computeBaseFilterOptions(STATE.plants);
  fillSelect(\$("#soilType"),     opts.soil);
  fillSelect(\$("#soilMoisture"), opts.moist);
  fillSelect(\$("#acidity"),      opts.acid);
  fillSelect(\$("#spread"),       opts.spread);
  await applyLayers();
  renderExtraFilters(STATE.plants);
  render(STATE.plants);
}

// ─── EVENT LISTENERS ──────────────────────────────────────────────────────────
function wire() {
  \$("#drawerClose")?.addEventListener("click", closeDrawer);
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeDrawer();
  });

  // Typology
  \$("#typology").addEventListener("change", async (e) => {
    STATE.selected.typology = e.target.value;
    const subs   = subtypeOptions(STATE.selected.typology);
    const subSel = \$("#subtype");
    subSel.innerHTML = "";
    for (const s of subs) {
      const opt = document.createElement("option");
      opt.value = s; opt.textContent = s;
      subSel.appendChild(opt);
    }
    STATE.selected.subtype = subs[0];
    subSel.value = STATE.selected.subtype;
    await loadTypologyAndRender();
  });

  // Subtype
  \$("#subtype").addEventListener("change", async (e) => {
    STATE.selected.subtype = e.target.value;
    await loadTypologyAndRender();
  });

  // Basis filters
  \$("#soilType").addEventListener("change", (e) => {
    STATE.selected.soilType = e.target.value;
    render(STATE.plants);
  });
  \$("#soilMoisture").addEventListener("change", (e) => {
    STATE.selected.soilMoisture = e.target.value;
    render(STATE.plants);
  });
  \$("#acidity").addEventListener("change", (e) => {
    STATE.selected.acidity = e.target.value;
    render(STATE.plants);
  });
  \$("#spread").addEventListener("change", (e) => {
    STATE.selected.spread = e.target.value;
    render(STATE.plants);
  });

  // Lagen
  \$("#layer_klimaat").addEventListener("change", async (e) => {
    STATE.selected.layers.klimaat = e.target.checked;
    await applyLayers();
    render(STATE.plants);
  });
  \$("#layer_amber").addEventListener("change", async (e) => {
    STATE.selected.layers.amber = e.target.checked;
    await applyLayers();
    render(STATE.plants);
  });

  // Regionaal
  \$("#layer_regionaal").addEventListener("change", async (e) => {
    STATE.selected.layers.regionaal = e.target.checked;
    \$("#districtWrap").style.display = e.target.checked ? "block" : "none";
    await applyLayers();
    render(STATE.plants);
  });
  \$("#district").addEventListener("change", async (e) => {
    STATE.selected.district = e.target.value;
    await applyLayers();
    render(STATE.plants);
  });
  const paint = \$("#regionModePaint");
  const filt  = \$("#regionModeFilter");
  if (paint && filt) {
    paint.addEventListener("change", () => {
      if (paint.checked) { STATE.selected.regionMode = "paint"; render(STATE.plants); }
    });
    filt.addEventListener("change", () => {
      if (filt.checked)  { STATE.selected.regionMode = "filter"; render(STATE.plants); }
    });
  }

  // BOBO
  \$("#layer_bobo").addEventListener("change", async (e) => {
    STATE.selected.layers.bobo = e.target.checked;
    \$("#boboWrap").style.display     = e.target.checked ? "block" : "none";
    \$("#boboCodeWrap").style.display = e.target.checked ? "block" : "none";
    await applyLayers();
    render(STATE.plants);
  });
  \$("#boboGroup").addEventListener("change", async (e) => {
    STATE.selected.boboGroup = e.target.value;
    const codes = STATE.options.bobo.codesByGroup.get(STATE.selected.boboGroup) || [];
    fillSelect(\$("#boboCode"), codes, true);
    STATE.selected.boboCode = \$("#boboCode").value || "ALLE";
    await applyLayers();
    render(STATE.plants);
  });
  \$("#boboCode").addEventListener("change", async (e) => {
    STATE.selected.boboCode = e.target.value;
    await applyLayers();
    render(STATE.plants);
  });

  // Fyto
  \$("#layer_fyto").addEventListener("change", async (e) => {
    STATE.selected.layers.fyto = e.target.checked;
    await applyLayers();
    render(STATE.plants);
  });

  // Export
  \$("#exportCsv").addEventListener("click", () => {
    const filtered = STATE.plants.filter(matchesAll);
    const filename = `planten_\${STATE.selected.typology}_\${STATE.selected.subtype}.csv`;
    downloadText(filename, toCsv(filtered));
  });
}

// ─── INIT ─────────────────────────────────────────────────────────────────────
async function init() {
  try {
    await loadConfig();
  } catch (e) {
    console.error("Config laden mislukt:", e);
  }
  await loadBoboOptions();
  await loadRegionOptions();

  // Typologie dropdown
  const typSel = \$("#typology");
  typSel.innerHTML = "";
  for (const t of typologyOptions()) {
    const opt = document.createElement("option");
    opt.value = t; opt.textContent = t;
    typSel.appendChild(opt);
  }
  STATE.selected.typology = typologyOptions()[0];
  typSel.value = STATE.selected.typology;

  // Subtype dropdown
  const subSel = \$("#subtype");
  subSel.innerHTML = "";
  const subs = subtypeOptions(STATE.selected.typology);
  for (const s of subs) {
    const opt = document.createElement("option");
    opt.value = s; opt.textContent = s;
    subSel.appendChild(opt);
  }
  STATE.selected.subtype = subs[0];
  subSel.value = STATE.selected.subtype;

  // District dropdown
  const districtSel = \$("#district");
  districtSel.innerHTML = "";
  const opt0 = document.createElement("option");
  opt0.value = ""; opt0.textContent = "Kies district…";
  districtSel.appendChild(opt0);
  for (const d of STATE.options.region.districts) {
    const opt = document.createElement("option");
    opt.value = d; opt.textContent = d;
    districtSel.appendChild(opt);
  }
  STATE.selected.district = "";

  // BOBO dropdowns
  const bg   = \$("#boboGroup");
  bg.innerHTML = "";
  const optG = document.createElement("option");
  optG.value = ""; optG.textContent = "Kies groep…";
  bg.appendChild(optG);
  for (const g of STATE.options.bobo.groups) {
    const opt = document.createElement("option");
    opt.value = g; opt.textContent = g;
    bg.appendChild(opt);
  }
  STATE.selected.boboGroup = "";
  fillSelect(\$("#boboCode"), [], true);
  STATE.selected.boboCode = "ALLE";

  wire();
  await loadTypologyAndRender();
}

window.addEventListener("DOMContentLoaded", init);
