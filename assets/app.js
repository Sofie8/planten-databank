/* Planten Databank – app.js (stable)
   - Dropdowns + filters + layers (folders)
   - BOBO per bodemcode file: data/layers/bobo_regio/BOBO_<code>.xlsx
   - FYTO summary per typology file with sheets (bodemwater_PFAS, lucht_metalen, ...)
   - Drawer: photos + Ecoflora all columns + FYTO summary + FYTO detail (Meer details + search + CSV exports)
*/

"use strict";
const $ = (sel) => document.querySelector(sel);

function setMeta(msg){ const el = $("#resultsMeta"); if(el) el.textContent = msg; }
function logErr(err){ console.error(err); setMeta("⚠️ " + (err?.message || err)); }

const FYTO_POLLUTANTS = ["PFAS", "metalen", "organische"];
const FYTO_POLLUTANT_LABELS = {
  PFAS: "PFAS",
  metalen: "zware metalen",
  organische: "dioxine- en polychloorbifenylen (PCBs)"
};
const FYTO_MEDIA = ["bodemwater", "lucht"];
const FYTO_MEDIA_LABELS = {
  bodemwater: "bodem en water",
  lucht: "lucht"
};

const IMAGE_BASE_URL = "https://pub-bb204453b9b642598d8514f7ac4f68be.r2.dev/images";

const STATE = {
  config: null,
  plants: [],
  loaded: {
    typology: new Map(),
    layerIndex: new Map(),
    fytoDetail: new Map(),
  },
  selected: {
    typology: null,
    subtype: null,
    soilType: "ALLE",
    soilMoisture: "ALLE",
    acidity: "ALLE",
    spread: "ALLE",
    globalSearch: "",
    layers: { klimaat:false, amber:false, regionaal:false, bobo:false, fyto:false },
    district: "",
    regionMode: "paint",
    boboCode: "ALLE",
    fytoPollutant: "PFAS",
    fytoMedium: "bodemwater",
  },
  table: { extraCols: [] },
  ui: { currentPlantKey: "" }
};

function norm(s){ return String(s ?? "").replace(/\u00a0/g," ").trim().replace(/\s+/g," "); }
function keyLatin(s){ return norm(s).toLowerCase(); }
function keyDutch(s){ return "nl:" + norm(s).toLowerCase(); }
function splitList(s){
  const t = norm(s); if(!t) return [];
  return t.split(/[,|]/).map(x=>norm(x).toLowerCase()).filter(Boolean);
}
function splitPipesRaw(s){
  const t = norm(s); if(!t) return [];
  return t.split("|").map(x=>norm(x)).filter(Boolean);
}
function uniqSorted(arr){ return Array.from(new Set(arr)).sort((a,b)=>a.localeCompare(b)); }

async function fetchXlsx(path){
  const res = await fetch(path);
  if(!res.ok) throw new Error(`Fetch failed: ${path} (${res.status})`);
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, {type:"array"});
}
function sheetToJson(wb, sheetName){
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {defval:""});
}
function pickBestSheetName(wb){
  const wanted = new Set(["latijnse naam","nederlandse naam","kenmerk_bodemtype","comments on phytoremedial effectiveness","reference (author, year, doi)","extractor / excluder"]);
  for(const name of wb.SheetNames){
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if(!rows || !rows.length) continue;
    const header = rows[0].map(h=>norm(h).toLowerCase());
    if(header.some(h=>wanted.has(h))) return name;
  }
  return wb.SheetNames[0];
}
function pickFytoSheetName(wb){
  const pol = (STATE.selected.fytoPollutant || "PFAS").toString();
  const med = (STATE.selected.fytoMedium || "bodemwater").toString();
  const candidates = [
    `${med}_${pol}`,
    `${med}_${pol.toLowerCase()}`,
    `${med.toLowerCase()}_${pol}`,
    `${med.toLowerCase()}_${pol.toLowerCase()}`
  ];
  for(const cand of candidates){
    const hit = wb.SheetNames.find(n => n.toLowerCase() === cand.toLowerCase());
    if(hit) return hit;
  }
  return pickBestSheetName(wb);
}

function typologyFileToPath(name){ return `data/typologies/${name}.xlsx`; }
function layerFileToPath(name){ return `data/layers/${name}.xlsx`; }

async function loadConfig(){
  const res = await fetch("data/config.json");
  if(!res.ok) throw new Error("Missing data/config.json");
  STATE.config = await res.json();
}

function typologyOptions(){
  const cfg = STATE.config?.typologies ? Object.keys(STATE.config.typologies) : [];
  return cfg.length ? cfg : ["1.Bomen","2.Haagplanten","3.wadi","4.bloemenweide","5.inheemse_planten","6.gevelgroen","7.grasland_weide","9.fruit_groenten_kruiden"];
}
function subtypeOptions(typ){
  const node = STATE.config?.typologies?.[typ];
  const subs = node?.subtypes ? Object.keys(node.subtypes) : ["Alle"];
  return subs.length ? subs : ["Alle"];
}
function filesForTypology(typ, subtype){
  const node = STATE.config?.typologies?.[typ];
  if(!node) return [];
  const keys = Object.keys(node.subtypes || {});
  const chosen = keys.includes(subtype) ? subtype : (keys[0] || "Alle");
  const refs = node.subtypes?.[chosen] ?? [];
  return (refs ?? []).map(norm).filter(Boolean).map(typologyFileToPath);
}

function districtToFilename(d){
  const cleaned = norm(d).replaceAll("/", "_").replaceAll(".", "");
  return cleaned.split(" ").join("_");
}
function layerFiles(layerKey, typ){
  if(layerKey==="regionaal"){
    if(!STATE.selected.district) return [];
    return [layerFileToPath(`regionale_soortenlijst/${districtToFilename(STATE.selected.district)}`)];
  }
  if(layerKey==="bobo"){
    const code = (STATE.selected.boboCode || "").trim();
    if(!code || code==="ALLE") return [];
    return [layerFileToPath(`bobo_regio/BOBO_${code}`)];
  }
  if(layerKey==="fyto"){
    const refs = (STATE.config?.layers?.fytoremediatie?.[typ] ?? []).map(norm).filter(Boolean);
    return refs.map(layerFileToPath);
  }
  const keyMap = { klimaat:"klimaatbomenlijst", amber:"amberlijst" };
  const refs = (STATE.config?.layers?.[keyMap[layerKey]] ?? []).map(norm).filter(Boolean);
  return refs.map(layerFileToPath);
}

function rowLatin(row){
  return norm(row["Latijnse naam"] || row["Latijnse naam "] || row["latijnse naam"] || row["LatijnseNaam"] || row["latin"] || row["Soort"] || "");
}
function rowDutch(row){
  return norm(row["Nederlandse naam"] || row["Nederlandse naam "] || row["nederlandse naam"] || row["NederlandseNaam"] || row["dutch"] || "");
}
const BASE_FILTER_KEYS = new Set(["kenmerk_bodemtype","kenmerk_bodemvochtigheid","kenmerk_zuurtegraad","kenmerk_verspreiding"]);

function plantFromRow(row){
  const latin=rowLatin(row);
  const dutch=rowDutch(row);
  const traits={
    bodemtype: splitList(row["kenmerk_bodemtype"]),
    bodemvocht: splitList(row["kenmerk_bodemvochtigheid"]),
    zuur: splitList(row["kenmerk_zuurtegraad"]),
    verspreiding: splitList(row["kenmerk_verspreiding"]),
  };
  const dynamic=new Map();
  for(const [k,v] of Object.entries(row)){
    const kk=norm(k).toLowerCase();
    if(!kk.startsWith("kenmerk_")) continue;
    if(BASE_FILTER_KEYS.has(kk)) continue;
    const vals=splitList(v);
    if(vals.length) dynamic.set(kk, vals);
  }
  const fotoIds=splitPipesRaw(row["foto_ids"] || row["foto_id"] || row["foto"] || "");
  return {
    latin, dutch,
    traits, dynamic,
    fotoIds,
    raw: row,
    layers:{klimaat:false, amber:false, regionaal:false, bobo:false, fyto:false},
    boboCode:null,
    fytoRow:null,
  };
}

async function loadTypologyPlants(typ, subtype){
  const files = filesForTypology(typ, subtype);
  const all=[];
  for(const file of files){
    if(STATE.loaded.typology.has(file)){
      all.push(...STATE.loaded.typology.get(file));
      continue;
    }
    const wb = await fetchXlsx(file);
    const sheet = pickBestSheetName(wb);
    const rows = sheetToJson(wb, sheet);
    const plants = rows.map(plantFromRow).filter(p=>p.latin);
    STATE.loaded.typology.set(file, plants);
    all.push(...plants);
  }
  const merged=new Map();
  for(const p of all){
    const k=keyLatin(p.latin);
    if(!merged.has(k)) merged.set(k,p);
    else{
      const ex=merged.get(k);
      if(!ex.dutch && p.dutch) ex.dutch=p.dutch;
      ex.traits.bodemtype=uniqSorted([...ex.traits.bodemtype,...p.traits.bodemtype]);
      ex.traits.bodemvocht=uniqSorted([...ex.traits.bodemvocht,...p.traits.bodemvocht]);
      ex.traits.zuur=uniqSorted([...ex.traits.zuur,...p.traits.zuur]);
      ex.traits.verspreiding=uniqSorted([...ex.traits.verspreiding,...p.traits.verspreiding]);
      ex.fotoIds = Array.from(new Set([...(ex.fotoIds||[]), ...(p.fotoIds||[])]));
      for(const [k2,v2] of Object.entries(p.raw||{})){
        if(ex.raw[k2]===undefined || ex.raw[k2]==="") ex.raw[k2]=v2;
      }
    }
  }
  return Array.from(merged.values());
}

async function loadLayerIndexOne(filePath, sheetName){
  const key = `${filePath}|${sheetName||"AUTO"}`;
  if(STATE.loaded.layerIndex.has(key)) return STATE.loaded.layerIndex.get(key);
  const wb = await fetchXlsx(filePath);
  const sheet = sheetName || pickBestSheetName(wb);
  const rows = sheetToJson(wb, sheet);
  const idx=new Map();
  for(const r of rows){
    const latin=rowLatin(r);
    const dutch=rowDutch(r);
    if(latin) idx.set(keyLatin(latin), r);
    if(dutch) idx.set(keyDutch(dutch), r);
  }
  STATE.loaded.layerIndex.set(key, idx);
  return idx;
}
function matchLayerRow(idx, plant){
  const k1=keyLatin(plant.latin);
  if(idx.has(k1)) return idx.get(k1);
  if(plant.dutch){
    const k2=keyDutch(plant.dutch);
    if(idx.has(k2)) return idx.get(k2);
  }
  return null;
}

function computeBaseFilterOptions(plants){
  const soils=[], moist=[], acid=[], spread=[];
  for(const p of plants){
    soils.push(...p.traits.bodemtype);
    moist.push(...p.traits.bodemvocht);
    acid.push(...p.traits.zuur);
    spread.push(...p.traits.verspreiding);
  }
  return { soil: uniqSorted(soils), moist: uniqSorted(moist), acid: uniqSorted(acid), spread: uniqSorted(spread) };
}
function fillSelect(selectEl, values, includeAll=true, allLabel="Alle"){
  if(!selectEl) return;
  const prev=selectEl.value;
  selectEl.innerHTML="";
  const add=(v,t)=>{ const o=document.createElement("option"); o.value=v; o.textContent=t ?? v; selectEl.appendChild(o); };
  if(includeAll) add("ALLE", allLabel);
  for(const v of values) add(v,v);
  const cand = prev || (includeAll ? "ALLE" : (values[0]||""));
  if(Array.from(selectEl.options).some(o=>o.value===cand)) selectEl.value=cand;
}
function fillSelectWithLabels(selectEl, values, labelsMap){
  if(!selectEl) return;
  const prev=selectEl.value;
  selectEl.innerHTML="";
  for(const v of values){
    const o=document.createElement("option");
    o.value=v;
    o.textContent = (labelsMap && labelsMap[v]) ? labelsMap[v] : v;
    selectEl.appendChild(o);
  }
  if(prev && Array.from(selectEl.options).some(o=>o.value===prev)) selectEl.value=prev;
}

function matchesAll(p){
  const st=STATE.selected.soilType.toLowerCase();
  const sm=STATE.selected.soilMoisture.toLowerCase();
  const ac=STATE.selected.acidity.toLowerCase();
  const sp=STATE.selected.spread.toLowerCase();
  if(!((st==="alle")||p.traits.bodemtype.includes(st))) return false;
  if(!((sm==="alle")||p.traits.bodemvocht.includes(sm))) return false;
  if(!((ac==="alle")||p.traits.zuur.includes(ac))) return false;
  if(!((sp==="alle")||p.traits.verspreiding.includes(sp))) return false;
  if(STATE.selected.layers.regionaal && STATE.selected.regionMode==="filter" && !p.layers.regionaal) return false;
  const q=(STATE.selected.globalSearch||"").trim().toLowerCase();
  if(q){
    const blob = `${p.latin} ${p.dutch||""} ${Object.values(p.raw||{}).map(v=>norm(v)).join(" ")}`.toLowerCase();
    if(!blob.includes(q)) return false;
  }
  return true;
}
function badgesForPlant(p){
  const b=[];
  if(p.layers.klimaat) b.push("Klimaat");
  if(p.layers.amber) b.push("AMBER");
  if(p.layers.regionaal) b.push("Regionaal");
  if(p.layers.bobo) b.push("BOBO");
  if(p.layers.fyto) b.push("Fyto");
  return b;
}
function pickAny(row, keys){
  for(const k of keys){
    if(row && row[k]!==undefined && row[k]!==null && String(row[k]).trim()!=="") return norm(row[k]);
  }
  return "";
}

function buildFytoColumns(){
  STATE.table.extraCols = [
    { label:"Comments on phytoremedial effectiveness", getter:(p)=>pickAny(p.fytoRow,["Comments on phytoremedial effectiveness","Comments"]) },
    { label:"Continent-Country-City-Site", getter:(p)=>pickAny(p.fytoRow,["Continent-Country-City-Site","Site"]) },
    { label:"Reference (author, year, doi)", getter:(p)=>pickAny(p.fytoRow,["Reference (author, year, doi)","Reference (author, year)","Reference"]) },
    { label:"Extractor / Excluder", getter:(p)=>pickAny(p.fytoRow,["Extractor / Excluder","extractor_excluder","Extractor/Excluder"]) },
  ];
}
function clearExtraCols(){ STATE.table.extraCols=[]; }

async function applyLayers(){
  for(const p of STATE.plants){
    p.layers={klimaat:false, amber:false, regionaal:false, bobo:false, fyto:false};
    p.boboCode=null;
    p.fytoRow=null;
  }
  clearExtraCols();
  const typ=STATE.selected.typology;

  if(STATE.selected.layers.klimaat){
    for(const file of layerFiles("klimaat", typ)){
      const idx = await loadLayerIndexOne(file);
      for(const p of STATE.plants) if(matchLayerRow(idx,p)) p.layers.klimaat=true;
    }
  }
  if(STATE.selected.layers.amber){
    for(const file of layerFiles("amber", typ)){
      const idx = await loadLayerIndexOne(file);
      for(const p of STATE.plants) if(matchLayerRow(idx,p)) p.layers.amber=true;
    }
  }
  if(STATE.selected.layers.regionaal){
    for(const file of layerFiles("regionaal", typ)){
      const idx = await loadLayerIndexOne(file);
      for(const p of STATE.plants) if(matchLayerRow(idx,p)) p.layers.regionaal=true;
    }
  }
  if(STATE.selected.layers.bobo){
    const code=(STATE.selected.boboCode||"").trim();
    if(code && code!=="ALLE"){
      for(const file of layerFiles("bobo", typ)){
        const idx = await loadLayerIndexOne(file);
        for(const p of STATE.plants){
          if(matchLayerRow(idx,p)){ p.layers.bobo=true; p.boboCode=code; }
        }
      }
    }
  }
  if(STATE.selected.layers.fyto){
    const files = layerFiles("fyto", typ);
    const idxAll=new Map();
    for(const file of files){
      const wb = await fetchXlsx(file);
      const sheet = pickFytoSheetName(wb);
      const rows = sheetToJson(wb, sheet);
      for(const r of rows){
        const latin=rowLatin(r);
        const dutch=rowDutch(r);
        if(latin) idxAll.set(keyLatin(latin), r);
        if(dutch) idxAll.set(keyDutch(dutch), r);
      }
    }
    for(const p of STATE.plants){
      const row=matchLayerRow(idxAll,p);
      if(!row) continue;
      p.layers.fyto=true;
      p.fytoRow=row;
    }
    buildFytoColumns();
  }
}

function rebuildTableHeader(){
  const thead=$("#results thead"); if(!thead) return;
  const base=["Latijnse naam","Nederlandse naam","Bodemtype","Vocht","pH","Lagen"];
  const cols=[...base, ...STATE.table.extraCols.map(c=>c.label)];
  thead.innerHTML = `<tr>${cols.map(h=>`<th>${h}</th>`).join("")}</tr>`;
}
function render(){
  rebuildTableHeader();
  const tbody=$("#results tbody"); if(!tbody) return;
  tbody.innerHTML="";
  const filtered = STATE.plants.filter(matchesAll);
  setMeta(`${filtered.length} resultaten`);
  for(const p of filtered){
    const tr=document.createElement("tr");
    tr.innerHTML = `
      <td>${p.latin}</td>
      <td class="dutchCell">${p.dutch||"—"}</td>
      <td>${(p.traits.bodemtype||[]).join(", ")||"—"}</td>
      <td>${(p.traits.bodemvocht||[]).join(", ")||"—"}</td>
      <td>${(p.traits.zuur||[]).join(", ")||"—"}</td>
      <td></td>
      ${STATE.table.extraCols.map(()=>"<td></td>").join("")}
    `;
    // badges + extra cols
    const badgeTd = tr.children[5];
    const wrap=document.createElement("div"); wrap.className="badges";
    for(const b of badgesForPlant(p)){
      const s=document.createElement("span"); s.className="badge ok"; s.textContent=b; wrap.appendChild(s);
    }
    if(p.layers.bobo && p.boboCode){
      const s=document.createElement("span"); s.className="badge"; s.textContent=p.boboCode; wrap.appendChild(s);
    }
    badgeTd.appendChild(wrap);

    for(let i=0;i<STATE.table.extraCols.length;i++){
      const td = tr.children[6+i];
      const v = STATE.table.extraCols[i].getter(p);
      td.textContent = v || "—";
    }

    tr.addEventListener("click", ()=>openDrawer(p));
    tr.querySelector(".dutchCell")?.addEventListener("click",(e)=>{ e.stopPropagation(); openDrawer(p); });
    tbody.appendChild(tr);
  }
}

// CSV utilities
function csvEscape(v){ return `"${String(v??"").replaceAll('"','""')}"`; }
function toCsv(rows){
  const headers=["Latijnse naam","Nederlandse naam","Bodemtype","Vocht","pH","Verspreiding","Lagen", ...STATE.table.extraCols.map(c=>c.label)];
  return [
    headers.map(csvEscape).join(","),
    ...rows.map(p=>{
      const layers=badgesForPlant(p).join("|");
      const base=[p.latin,p.dutch,(p.traits.bodemtype||[]).join("|"),(p.traits.bodemvocht||[]).join("|"),(p.traits.zuur||[]).join("|"),(p.traits.verspreiding||[]).join("|"),layers];
      const extras=STATE.table.extraCols.map(c=>c.getter(p) ?? "");
      return [...base,...extras].map(csvEscape).join(",");
    })
  ].join("\n");
}
function rowsToCsv(rows){
  if(!rows || !rows.length) return "";
  const headers = Object.keys(rows[0]);
  return [headers.map(csvEscape).join(","), ...rows.map(r=>headers.map(h=>csvEscape(r[h])).join(","))].join("\n");
}
function downloadText(filename,text){
  const blob=new Blob([text],{type:"text/plain;charset=utf-8"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url; a.download=filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),250);
}

// Drawer: images
function urlForFotoId(id, ext){ return `${IMAGE_BASE_URL}/${id}.${ext}`; }
function loadImage(url){
  return new Promise((resolve,reject)=>{
    const img=new Image();
    img.onload=()=>resolve(url);
    img.onerror=()=>reject(new Error("notfound"));
    img.src=url;
  });
}
async function resolveFotoUrls(ids, max=60){
  const out=[];
  for(const id of ids){
    let ok=null;
    try{ ok=await loadImage(urlForFotoId(id,"jpg")); }
    catch{ try{ ok=await loadImage(urlForFotoId(id,"png")); }catch{ ok=null; } }
    if(ok){ out.push(ok);
      if(out.length>=max) break; }
  }
  return out;
}

// FYTO detail helpers
function fytoBaseNameForTypology(typ){
  const arr = STATE.config?.layers?.fytoremediatie?.[typ] || [];
  const first = Array.isArray(arr) ? (arr[0] || "") : "";
  const ref = norm(first);
  return (ref.split("/").pop() || "Fytoremediatie");
}
function buildFytoDetailPath(typ){
  const base = fytoBaseNameForTypology(typ);
  const med = (STATE.selected.fytoMedium || "bodemwater").toString();
  const pol = (STATE.selected.fytoPollutant || "PFAS").toString();
  return `data/layers/fytoremediatie/detail/${base}_${med}_${pol}_detail.xlsx`;
}
function stringifyRow(row){ return Object.values(row).map(v=>norm(v)).join(" ").toLowerCase(); }

function renderEcofloraDetails(plant){
  const tbody = $("#ecoTable tbody");
  if(!tbody) return;
  tbody.innerHTML = "";
  const entries = Object.entries(plant.raw || {})
    .filter(([k,v])=>norm(k)!=="" && norm(v)!=="")
    .sort((a,b)=>a[0].localeCompare(b[0]));
  for(const [k,v] of entries){
    const tr=document.createElement("tr");
    const tdK=document.createElement("td"); tdK.textContent=norm(k); tr.appendChild(tdK);
    const tdV=document.createElement("td"); tdV.textContent=norm(v); tr.appendChild(tdV);
    tbody.appendChild(tr);
  }
}

async function loadPlantFytoDetailRows(plant){
  const path = buildFytoDetailPath(STATE.selected.typology);
  let all = STATE.loaded.fytoDetail.get(path);
  if(!all){
    const wb = await fetchXlsx(path);
    const sheet = pickBestSheetName(wb);
    all = sheetToJson(wb, sheet);
    STATE.loaded.fytoDetail.set(path, all);
  }
  const latin = keyLatin(plant.latin);
  const dutch = keyLatin(plant.dutch || "");
  return all.filter(r=>{
    const rLat = keyLatin(rowLatin(r));
    const rDu = keyLatin(rowDutch(r));
    return (rLat && rLat===latin) || (rDu && rDu===dutch);
  });
}

async function renderFytoDetail(plant){
  const detailBox=$("#fytoDetailBox");
  const dtHead=$("#fytoDetailTable thead");
  const dtBody=$("#fytoDetailTable tbody");
  const meta=$("#fytoDetailMeta");
  const q = ($("#fytoDetailSearch")?.value || "").trim().toLowerCase();
  if(!detailBox || !dtHead || !dtBody) return {all:[], filtered:[]};

  const allRows = await loadPlantFytoDetailRows(plant);
  const filtered = q ? allRows.filter(r=>stringifyRow(r).includes(q)) : allRows;

  if(meta){
    meta.textContent = `${filtered.length} / ${allRows.length} studies (${buildFytoDetailPath(STATE.selected.typology).split("/").pop()})`;
  }

  dtHead.innerHTML=""; dtBody.innerHTML="";
  if(filtered.length===0){
    dtBody.innerHTML = `<tr><td style="padding:10px">Geen detailrecords gevonden.</td></tr>`;
    return {all:allRows, filtered};
  }
  const headers = Object.keys(filtered[0]);
  dtHead.innerHTML = `<tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr>`;
  for(const r of filtered){
    const tr=document.createElement("tr");
    for(const h of headers){
      const td=document.createElement("td");
      td.textContent = norm(r[h]);
      tr.appendChild(td);
    }
    dtBody.appendChild(tr);
  }
  return {all:allRows, filtered};
}

function openDrawer(plant){
  const drawer=$("#detailDrawer"); if(!drawer) return;
  STATE.ui.currentPlantKey = keyLatin(plant.latin);

  $("#drawerTitle").textContent = plant.latin;
  $("#drawerSub").textContent = plant.dutch || "";
  $("#drawerSoil").textContent = (plant.traits.bodemtype||[]).join(", ") || "—";
  $("#drawerMoist").textContent = (plant.traits.bodemvocht||[]).join(", ") || "—";
  $("#drawerPh").textContent = (plant.traits.zuur||[]).join(", ") || "—";
  $("#drawerSpread").textContent = (plant.traits.verspreiding||[]).join(", ") || "—";

  // Ecoflora details table
  renderEcofloraDetails(plant);
  const ecoWrap=$("#ecoTableWrap");
  if(ecoWrap) ecoWrap.style.display="none";
  $("#ecoToggleBtn")?.addEventListener("click", ()=>{
    if(!ecoWrap) return;
    ecoWrap.style.display = ecoWrap.style.display==="none" ? "block" : "none";
  });

  // FYTO summary fields
  const fytoBox=$("#drawerFyto");
  if(STATE.selected.layers.fyto && plant.fytoRow){
    fytoBox.style.display="block";
    $("#fytoComments").textContent = pickAny(plant.fytoRow,["Comments on phytoremedial effectiveness","Comments"]) || "—";
    $("#fytoSite").textContent = pickAny(plant.fytoRow,["Continent-Country-City-Site","Site"]) || "—";
    $("#fytoRef").textContent = pickAny(plant.fytoRow,["Reference (author, year, doi)","Reference (author, year)","Reference"]) || "—";
  }else{
    fytoBox.style.display="none";
  }

  // Reset detail box
  const detailBox=$("#fytoDetailBox"); if(detailBox) detailBox.style.display="none";
  const meta=$("#fytoDetailMeta"); if(meta) meta.textContent="";
  const dtHead=$("#fytoDetailTable thead"); if(dtHead) dtHead.innerHTML="";
  const dtBody=$("#fytoDetailTable tbody"); if(dtBody) dtBody.innerHTML="";
  const search=$("#fytoDetailSearch"); if(search) search.value="";

  // Wire "Meer details"
  const moreBtn=$("#fytoMoreBtn");
  if(moreBtn){
    moreBtn.onclick = async ()=>{
      if(!detailBox) return;
      const open = detailBox.style.display==="none";
      detailBox.style.display = open ? "block" : "none";
      if(open){
        try{ await renderFytoDetail(plant); }
        catch(e){ if(meta) meta.textContent="Detailbestand niet gevonden of niet leesbaar."; }
      }
    };
  }
  $("#fytoDetailSearch")?.addEventListener("input", async ()=>{
    if(detailBox && detailBox.style.display!=="none"){
      try{ await renderFytoDetail(plant); }catch(e){}
    }
  });
  $("#fytoDetailCsvBtn")?.addEventListener("click", async ()=>{
    try{
      const {filtered} = await renderFytoDetail(plant);
      downloadText(`fyto_detail_${plant.latin}.csv`, rowsToCsv(filtered));
    }catch(e){ if(meta) meta.textContent="Kan detail CSV niet exporteren."; }
  });
  $("#fytoExtendedCsvBtn")?.addEventListener("click", async ()=>{
    try{
      const {all} = await renderFytoDetail(plant);
      const summary = Object.assign({}, plant.raw || {});
      summary["Latijnse naam"]=plant.latin;
      summary["Nederlandse naam"]=plant.dutch||"";
      summary["FYTO_pollutant"]=STATE.selected.fytoPollutant;
      summary["FYTO_medium"]=STATE.selected.fytoMedium;
      summary["FYTO_comments"]=$("#fytoComments")?.textContent || "";
      summary["FYTO_site"]=$("#fytoSite")?.textContent || "";
      summary["FYTO_reference"]=$("#fytoRef")?.textContent || "";
      const combo = rowsToCsv([summary]) + "\n\nFYTO_DETAIL\n" + rowsToCsv(all);
      downloadText(`extended_${plant.latin}.csv`, combo);
    }catch(e){ if(meta) meta.textContent="Kan extended CSV niet exporteren."; }
  });

  // Images
  const imgHost=$("#drawerImages");
  if(imgHost) imgHost.innerHTML = `<div class="hint">Foto’s laden…</div>`;
  const ids = plant.fotoIds || [];
  if(imgHost){
    if(!ids.length){
      imgHost.innerHTML = `<div class="hint">Geen foto_ids voor deze plant.</div>`;
    }else{
      resolveFotoUrls(ids, 60).then(urls=>{
        if(STATE.ui.currentPlantKey !== keyLatin(plant.latin)) return;
        if(!urls.length){ imgHost.innerHTML = `<div class="hint">Geen foto's gevonden.</div>`; return; }
        let idx=0;
        const renderCarousel=()=>{
          imgHost.innerHTML = `
            <div class="carousel">
              <img class="carouselMain" src="${urls[idx]}" alt="foto ${idx+1}">
              <button class="carouselBtn prev" id="carouselPrev" aria-label="vorige">‹</button>
              <button class="carouselBtn next" id="carouselNext" aria-label="volgende">›</button>
              <div class="carouselCounter">${idx+1} / ${urls.length}</div>
            </div>`;
          $("#carouselPrev").addEventListener("click",(e)=>{ e.stopPropagation(); idx=(idx-1+urls.length)%urls.length; renderCarousel(); });
          $("#carouselNext").addEventListener("click",(e)=>{ e.stopPropagation(); idx=(idx+1)%urls.length; renderCarousel(); });
        };
        renderCarousel();
      });
    }
  }

  drawer.classList.add("open");
  drawer.setAttribute("aria-hidden","false");
}
function closeDrawer(){
  const drawer=$("#detailDrawer"); if(!drawer) return;
  drawer.classList.remove("open");
  drawer.setAttribute("aria-hidden","true");
  const imgHost=$("#drawerImages"); if(imgHost) imgHost.innerHTML="";
}

async function loadTypologyAndRender(){
  setMeta("Laden…");
  STATE.plants = await loadTypologyPlants(STATE.selected.typology, STATE.selected.subtype);
  const opts = computeBaseFilterOptions(STATE.plants);
  fillSelect($("#soilType"), opts.soil, true);
  fillSelect($("#soilMoisture"), opts.moist, true);
  fillSelect($("#acidity"), opts.acid, true);
  fillSelect($("#spread"), opts.spread, true);
  await applyLayers();
  render();
}

function wire(){
  $("#drawerClose")?.addEventListener("click", closeDrawer);
  window.addEventListener("keydown",(e)=>{ if(e.key==="Escape") closeDrawer(); });

  $("#typology")?.addEventListener("change", async (e)=>{
    STATE.selected.typology=e.target.value;
    const subs=subtypeOptions(STATE.selected.typology);
    const subSel=$("#subtype");
    subSel.innerHTML="";
    for(const s of subs){ const o=document.createElement("option"); o.value=s; o.textContent=s; subSel.appendChild(o); }
    STATE.selected.subtype=subs[0]; subSel.value=STATE.selected.subtype;
    await loadTypologyAndRender();
  });
  $("#subtype")?.addEventListener("change", async (e)=>{ STATE.selected.subtype=e.target.value; await loadTypologyAndRender(); });

  $("#soilType")?.addEventListener("change",(e)=>{ STATE.selected.soilType=e.target.value; render(); });
  $("#soilMoisture")?.addEventListener("change",(e)=>{ STATE.selected.soilMoisture=e.target.value; render(); });
  $("#acidity")?.addEventListener("change",(e)=>{ STATE.selected.acidity=e.target.value; render(); });
  $("#spread")?.addEventListener("change",(e)=>{ STATE.selected.spread=e.target.value; render(); });
  $("#globalSearch")?.addEventListener("input",(e)=>{ STATE.selected.globalSearch=e.target.value||""; render(); });

  $("#layer_klimaat")?.addEventListener("change", async (e)=>{ STATE.selected.layers.klimaat=e.target.checked; await applyLayers(); render(); });
  $("#layer_amber")?.addEventListener("change", async (e)=>{ STATE.selected.layers.amber=e.target.checked; await applyLayers(); render(); });

  $("#layer_regionaal")?.addEventListener("change", async (e)=>{
    STATE.selected.layers.regionaal=e.target.checked;
    const wrap=$("#districtWrap"); if(wrap) wrap.style.display = e.target.checked ? "block" : "none";
    await applyLayers(); render();
  });
  $("#district")?.addEventListener("change", async (e)=>{ STATE.selected.district=e.target.value; await applyLayers(); render(); });
  $("#regionModePaint")?.addEventListener("change", ()=>{ if($("#regionModePaint")?.checked){ STATE.selected.regionMode="paint"; render(); } });
  $("#regionModeFilter")?.addEventListener("change", ()=>{ if($("#regionModeFilter")?.checked){ STATE.selected.regionMode="filter"; render(); } });

  $("#layer_bobo")?.addEventListener("change", async (e)=>{
    STATE.selected.layers.bobo=e.target.checked;
    const wrap=$("#boboCodeWrap"); if(wrap) wrap.style.display = e.target.checked ? "block" : "none";
    await applyLayers(); render();
  });
  $("#boboCode")?.addEventListener("change", async (e)=>{ STATE.selected.boboCode=e.target.value; await applyLayers(); render(); });

  $("#layer_fyto")?.addEventListener("change", async (e)=>{
    STATE.selected.layers.fyto=e.target.checked;
    const wrap=$("#fytoWrap"); if(wrap) wrap.style.display = e.target.checked ? "grid" : "none";
    await applyLayers(); render();
  });
  $("#fytoPollutant")?.addEventListener("change", async (e)=>{ STATE.selected.fytoPollutant=e.target.value; if(STATE.selected.layers.fyto){ await applyLayers(); render(); } });
  $("#fytoMedium")?.addEventListener("change", async (e)=>{ STATE.selected.fytoMedium=e.target.value; if(STATE.selected.layers.fyto){ await applyLayers(); render(); } });

  $("#exportCsv")?.addEventListener("click", ()=>{
    const filtered=STATE.plants.filter(matchesAll);
    downloadText(`planten_${STATE.selected.typology}_${STATE.selected.subtype}.csv`, toCsv(filtered));
  });
}

async function init(){
  try{
    if(!window.XLSX){ setMeta("⚠️ XLSX library ontbreekt (check index.html script tag)."); return; }
    setMeta("Init…");
    await loadConfig();

    const typSel=$("#typology"); typSel.innerHTML="";
    const typs=typologyOptions();
    for(const t of typs){ const o=document.createElement("option"); o.value=t; o.textContent=t; typSel.appendChild(o); }
    STATE.selected.typology=typs[0]; typSel.value=STATE.selected.typology;

    const subSel=$("#subtype"); subSel.innerHTML="";
    const subs=subtypeOptions(STATE.selected.typology);
    for(const s of subs){ const o=document.createElement("option"); o.value=s; o.textContent=s; subSel.appendChild(o); }
    STATE.selected.subtype=subs[0]; subSel.value=STATE.selected.subtype;

    fillSelectWithLabels($("#fytoPollutant"), FYTO_POLLUTANTS, FYTO_POLLUTANT_LABELS);
    fillSelectWithLabels($("#fytoMedium"), FYTO_MEDIA, FYTO_MEDIA_LABELS);
    if($("#fytoPollutant")) $("#fytoPollutant").value = STATE.selected.fytoPollutant;
    if($("#fytoMedium")) $("#fytoMedium").value = STATE.selected.fytoMedium;

    wire();
    await loadTypologyAndRender();
  }catch(e){ logErr(e); }
}
window.addEventListener("DOMContentLoaded", init);
