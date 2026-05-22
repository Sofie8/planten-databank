/* PlantenDatabank – client-side mockup (GitHub Pages ready)
   v8:
   - Basisfilters: bodemtype, bodemvochtigheid, zuurtegraad, verspreiding
   - Extra filters: auto-detect alle kolommen die beginnen met kenmerk_ (behalve basisfilters) en toon ze als checkbox-facets
   - Regionaal: labelen (paint) vs alleen tonen (hard filter)
   - BOBO: 2-level selector (groep -> code), paintlaag (optioneel filter op code als kolom aanwezig)
*/

const $ = (sel) => document.querySelector(sel);

const STATE = {
  config: null,
  plants: [],
  loaded: { typologies: new Map(), layers: new Map() },
  options: {
    bobo: { groups: [], codesByGroup: new Map() },
    region: { districts: [] }
  },
  selected: {
    typology: null, subtype: null,
    soilType: "ALLE", soilMoisture: "ALLE", acidity: "ALLE", spread: "ALLE",
    layers: { klimaat:false, amber:false, regionaal:false, bobo:false, fyto:false },
    district: "",
    regionMode: "paint", // paint | filter
    boboGroup: "",
    boboCode: "ALLE",
    facets: new Map() // key -> Set(values)
  }
};

function norm(s){ return String(s ?? "").replace(/\u00a0/g," ").trim().replace(/\s+/g," "); }
function keyLatin(s){ return norm(s).toLowerCase(); }
function splitList(s){
  const t = norm(s);
  if(!t) return [];
  return t.split(/[,|]/).map(x=>norm(x).toLowerCase()).filter(Boolean);
}
function uniqSorted(arr){ return Array.from(new Set(arr)).sort((a,b)=>a.localeCompare(b)); }

async function fetchXlsx(path){
  const res = await fetch(path);
  if(!res.ok) throw new Error(`Fetch failed: ${path}`);
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, {type:"array"});
}
function firstSheetName(wb){ return wb.SheetNames[0]; }
function sheetToJson(wb, sheetName){
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, {defval:""});
}

async function loadConfig(){
  const res = await fetch("data/config.json");
  if(!res.ok) throw new Error("Missing data/config.json");
  STATE.config = await res.json();
}

function typologyOptions(){
  return ["1.Bomen","2.Haagplanten","3.wadi","4.bloemenweide","5.inheemse_planten","6.gevelgroen","7.grasland_weide","9.fruit_groenten_kruiden"];
}
function subtypeOptions(typology){
  const node = STATE.config?.typologies?.[typology];
  const subs = node?.subtypes ? Object.keys(node.subtypes) : ["Alle"];
  return subs.length ? subs : ["Alle"];
}
function filesForTypology(typology, subtype){
  const node = STATE.config?.typologies?.[typology];
  if(!node) return [];
  const sub = node.subtypes?.[subtype] ?? node.subtypes?.[Object.keys(node.subtypes||{})[0]] ?? [];
  return (sub ?? []).map(norm).filter(Boolean).map(n=>`data/typologies/${n}.xlsx`);
}

function districtToFilename(d){
  const cleaned = norm(d).replaceAll("/", "_").replaceAll("  ", " ").replaceAll(".", "");
  return cleaned.split(" ").join("_") + ".xlsx";
}

function layerFiles(layerKey, typology){
  if(layerKey==="regionaal"){
    if(!STATE.selected.district) return [];
    return [`data/layers/${districtToFilename(STATE.selected.district)}`];
  }
  if(layerKey==="bobo"){
    if(!STATE.selected.boboGroup) return [];
    return [`data/layers/BOBO_${STATE.selected.boboGroup}.xlsx`];
  }
  if(layerKey==="fyto"){
    const names = (STATE.config?.layers?.fytoremediatie?.[typology] ?? []).map(norm).filter(Boolean);
    return names.map(n=>`data/layers/${n}.xlsx`);
  }
  const names = (STATE.config?.layers?.[{
    klimaat:"klimaatbomenlijst",
    amber:"amberlijst"
  }[layerKey]] ?? []).map(norm).filter(Boolean);
  return names.map(n=>`data/layers/${n}.xlsx`);
}

function rowLatin(row){
  return norm(row["Latijnse naam"] || row["Latijnse naam "] || row["latijnse naam"] || row["LatijnseNaam"] || row["latin"] || "");
}
function rowDutch(row){
  return norm(row["Nederlandse naam"] || row["Nederlandse naam "] || row["nederlandse naam"] || row["NederlandseNaam"] || row["dutch"] || "");
}

const BASE_FILTER_KEYS = new Set([
  "kenmerk_bodemtype",
  "kenmerk_bodemvochtigheid",
  "kenmerk_zuurtegraad",
  "kenmerk_verspreiding"
]);

function plantFromRow(row){
  const latin = rowLatin(row);
  const dutch = rowDutch(row);
  const traits = {
    bodemtype: splitList(row["kenmerk_bodemtype"]),
    bodemvocht: splitList(row["kenmerk_bodemvochtigheid"]),
    zuur: splitList(row["kenmerk_zuurtegraad"]),
    verspreiding: splitList(row["kenmerk_verspreiding"]),
  };

  const dynamic = new Map();
  for(const [k,v] of Object.entries(row)){
    const kk = norm(k);
    if(!kk.toLowerCase().startsWith("kenmerk_")) continue;
    const kLower = kk.toLowerCase();
    if(BASE_FILTER_KEYS.has(kLower)) continue;
    const vals = splitList(v);
    if(vals.length) dynamic.set(kLower, vals);
  }

  return {
    latin, dutch,
    traits,
    dynamic,
    buy:{ url: norm(row["url"]), prijs: norm(row["prijs"]) },
    layers:{ klimaat:false, amber:false, regionaal:false, bobo:false, fyto:false },
    boboCode: null,
    fyto:null,
    raw: row
  };
}

async function loadTypologyPlants(typology, subtype){
  const files = filesForTypology(typology, subtype);
  if(files.length===0) return [];
  const all = [];

  for(const file of files){
    if(STATE.loaded.typologies.has(file)){
      all.push(...STATE.loaded.typologies.get(file));
      continue;
    }
    try{
      const wb = await fetchXlsx(file);
      const rows = sheetToJson(wb, firstSheetName(wb));
      const plants = rows.map(plantFromRow).filter(p=>p.latin);
      STATE.loaded.typologies.set(file, plants);
      all.push(...plants);
    }catch(e){
      console.warn("Could not load typology file:", file, e);
    }
  }

  const merged = new Map();
  for(const p of all){
    const k = keyLatin(p.latin);
    if(!merged.has(k)) merged.set(k, p);
    else{
      const ex = merged.get(k);
      if(!ex.dutch && p.dutch) ex.dutch = p.dutch;

      ex.traits.bodemtype = uniqSorted([...ex.traits.bodemtype, ...p.traits.bodemtype]);
      ex.traits.bodemvocht = uniqSorted([...ex.traits.bodemvocht, ...p.traits.bodemvocht]);
      ex.traits.zuur = uniqSorted([...ex.traits.zuur, ...p.traits.zuur]);
      ex.traits.verspreiding = uniqSorted([...ex.traits.verspreiding, ...p.traits.verspreiding]);

      for(const [dk, vals] of p.dynamic.entries()){
        const cur = ex.dynamic.get(dk) || [];
        ex.dynamic.set(dk, uniqSorted([...cur, ...vals]));
      }

      if(!ex.buy.url && p.buy.url) ex.buy.url = p.buy.url;
      if(!ex.buy.prijs && p.buy.prijs) ex.buy.prijs = p.buy.prijs;
    }
  }
  return Array.from(merged.values());
}

async function loadLayerSet(files){
  const out = new Map();
  for(const file of files){
    if(STATE.loaded.layers.has(file)){
      const cached = STATE.loaded.layers.get(file);
      for(const [k,v] of cached.entries()) out.set(k,v);
      continue;
    }
    try{
      const wb = await fetchXlsx(file);
      const rows = sheetToJson(wb, firstSheetName(wb));
      const map = new Map();
      for(const r of rows){
        const latin = rowLatin(r);
        if(!latin) continue;
        map.set(keyLatin(latin), r);
      }
      STATE.loaded.layers.set(file, map);
      for(const [k,v] of map.entries()) out.set(k,v);
    }catch(e){
      console.warn("Could not load layer file:", file, e);
    }
  }
  return out;
}

function fillSelect(selectEl, values, includeAll=true){
  const prev = selectEl.value;
  selectEl.innerHTML = "";
  const add = (v, label) => {
    const opt=document.createElement("option");
    opt.value=v;
    opt.textContent=label??v;
    selectEl.appendChild(opt);
  };
  if(includeAll) add("ALLE","Alle");
  for(const v of values) add(v,v);
  const cand = prev || (includeAll ? "ALLE" : (values[0] ?? ""));
  if(Array.from(selectEl.options).some(o=>o.value===cand)) selectEl.value=cand;
}

function computeBaseFilterOptions(plants){
  const soils=[], moist=[], acid=[], spread=[];
  for(const p of plants){
    soils.push(...p.traits.bodemtype);
    moist.push(...p.traits.bodemvocht);
    acid.push(...p.traits.zuur);
    spread.push(...p.traits.verspreiding);
  }
  return {
    soil: uniqSorted(soils),
    moist: uniqSorted(moist),
    acid: uniqSorted(acid),
    spread: uniqSorted(spread)
  };
}

function computeFacetOptions(plants){
  // key -> Map(value -> count)
  const counts = new Map();
  for(const p of plants){
    for(const [k, vals] of p.dynamic.entries()){
      if(!counts.has(k)) counts.set(k, new Map());
      const m = counts.get(k);
      for(const v of vals) m.set(v, (m.get(v)||0)+1);
    }
  }
  return counts;
}

function matchesBaseFilters(p){
  const st=STATE.selected.soilType.toLowerCase();
  const sm=STATE.selected.soilMoisture.toLowerCase();
  const ac=STATE.selected.acidity.toLowerCase();
  const sp=STATE.selected.spread.toLowerCase();

  const okSoil = (st==="alle") || p.traits.bodemtype.includes(st);
  const okMoist = (sm==="alle") || p.traits.bodemvocht.includes(sm);
  const okAc = (ac==="alle") || p.traits.zuur.includes(ac);
  const okSp = (sp==="alle") || p.traits.verspreiding.includes(sp);

  return okSoil && okMoist && okAc && okSp;
}

function matchesFacets(p){
  for(const [k, wantedSet] of STATE.selected.facets.entries()){
    if(!wantedSet || wantedSet.size===0) continue;
    const vals = p.dynamic.get(k) || [];
    const ok = vals.some(v=>wantedSet.has(v));
    if(!ok) return false;
  }
  return true;
}

function matchesRegionMode(p){
  if(!STATE.selected.layers.regionaal) return true;
  if(STATE.selected.regionMode !== "filter") return true;
  return p.layers.regionaal === true;
}

function matchesAll(p){
  return matchesBaseFilters(p) && matchesFacets(p) && matchesRegionMode(p);
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

function render(plants){
  const tbody=$("#results tbody");
  tbody.innerHTML="";
  const filtered=plants.filter(matchesAll);
  $("#resultsMeta").textContent =
    `${filtered.length} van ${plants.length} planten (typologie: ${STATE.selected.typology} • subtype: ${STATE.selected.subtype})`;

  for(const p of filtered){
    const tr=document.createElement("tr");
    tr.innerHTML = `<td>${p.latin}</td>
      <td>${p.dutch||"—"}</td>
      <td>${(p.traits.bodemtype||[]).join(", ")||"—"}</td>
      <td>${(p.traits.bodemvocht||[]).join(", ")||"—"}</td>
      <td>${(p.traits.zuur||[]).join(", ")||"—"}</td>
      <td></td>`;

    const td=tr.querySelector("td:last-child");
    const wrap=document.createElement("div");
    wrap.className="badges";

    for(const t of badgesForPlant(p)){
      const s=document.createElement("span");
      s.className="badge ok";
      s.textContent=t;
      wrap.appendChild(s);
    }
    if(p.layers.bobo && p.boboCode){
      const s=document.createElement("span");
      s.className="badge";
      s.textContent=`${p.boboCode}`;
      wrap.appendChild(s);
    }
    if(p.buy?.url){
      const s=document.createElement("span");
      s.className="badge";
      s.textContent = p.buy.prijs ? `Ecoflora ${p.buy.prijs}` : "Ecoflora";
      wrap.appendChild(s);
    }
    if(p.layers.fyto && p.fyto?.note){
      const s=document.createElement("span");
      s.className="badge";
      s.textContent="Fyto-notitie";
      s.title=p.fyto.note;
      wrap.appendChild(s);
    }
    td.appendChild(wrap);
    tbody.appendChild(tr);
  }
}

function toCsv(rows){
  const esc=(v)=>`"${String(v??"").replaceAll('"','""')}"`;
  return [["Latijnse naam","Nederlandse naam","Bodemtype","Vocht","pH","Verspreiding","Lagen"].map(esc).join(","),
    ...rows.map(p=>{
      const layers=badgesForPlant(p).join("|");
      return [
        p.latin,p.dutch,
        (p.traits.bodemtype||[]).join("|"),
        (p.traits.bodemvocht||[]).join("|"),
        (p.traits.zuur||[]).join("|"),
        (p.traits.verspreiding||[]).join("|"),
        layers
      ].map(esc).join(",");
    })
  ].join("\n");
}

function downloadText(filename,text){
  const blob=new Blob([text],{type:"text/plain;charset=utf-8"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;
  a.download=filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),250);
}

function pickAny(row, keys){
  for(const k of keys){
    if(row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== "") return norm(row[k]);
  }
  return "";
}

async function applyLayers(){
  for(const p of STATE.plants){
    p.layers={klimaat:false,amber:false,regionaal:false,bobo:false,fyto:false};
    p.fyto=null;
    p.boboCode=null;
  }
  const typ=STATE.selected.typology;

  if(STATE.selected.layers.klimaat){
    const map=await loadLayerSet(layerFiles("klimaat",typ));
    for(const p of STATE.plants) if(map.has(keyLatin(p.latin))) p.layers.klimaat=true;
  }
  if(STATE.selected.layers.amber){
    const map=await loadLayerSet(layerFiles("amber",typ));
    for(const p of STATE.plants) if(map.has(keyLatin(p.latin))) p.layers.amber=true;
  }
  if(STATE.selected.layers.regionaal){
    const map=await loadLayerSet(layerFiles("regionaal",typ));
    for(const p of STATE.plants) if(map.has(keyLatin(p.latin))) p.layers.regionaal=true;
  }
  if(STATE.selected.layers.bobo){
    const map=await loadLayerSet(layerFiles("bobo",typ));
    const wanted = STATE.selected.boboCode;
    for(const p of STATE.plants){
      const k = keyLatin(p.latin);
      if(!map.has(k)) continue;
      const row = map.get(k);
      const code = pickAny(row, ["code","Code","bodemcode","Bodemcode","BOBO","bobo_code","BOBO code","bobo"]);
      if(wanted && wanted !== "ALLE"){
        if(code && norm(code).toLowerCase() === norm(wanted).toLowerCase()){
          p.layers.bobo = true;
          p.boboCode = code || wanted;
        }
      }else{
        p.layers.bobo = true;
        p.boboCode = code || null;
      }
    }
  }
  if(STATE.selected.layers.fyto){
    const map=await loadLayerSet(layerFiles("fyto",typ));
    for(const p of STATE.plants){
      const k=keyLatin(p.latin);
      if(map.has(k)){
        p.layers.fyto=true;
        const r=map.get(k);
        const note = r["Comments on phytoremedial effectiveness"] || r["comment"] || r["AI comment (berm Genk, droog zand)"] || "";
        p.fyto={note:norm(note)};
      }
    }
  }
}

function toggleFacetPanel(id){
  const body = document.querySelector(`#facetBody_${id}`);
  if(body) body.classList.toggle("open");
}
function humanizeFacetKey(k){
  return norm(k).replace(/^kenmerk_/i, "").replaceAll("_", " ");
}

function renderExtraFilters(plants){
  const host = $("#extraFiltersList");
  host.innerHTML = "";

  const counts = computeFacetOptions(plants);
  const keys = uniqSorted(Array.from(counts.keys()));
  if(keys.length===0){
    host.innerHTML = `<div class="hint">Geen extra filters gevonden voor deze typologie/subtype.</div>`;
    return;
  }

  keys.forEach((k, idx)=>{
    const map = counts.get(k);
    const entries = Array.from(map.entries()).sort((a,b)=>b[1]-a[1] || a[0].localeCompare(b[0]));
    const facet = document.createElement("div");
    facet.className = "facet";

    const header = document.createElement("div");
    header.className = "facetHeader";
    header.innerHTML = `<div>
        <div class="facetTitle">${humanizeFacetKey(k)}</div>
        <div class="facetMeta">${entries.length} opties</div>
      </div>
      <div>▾</div>`;
    header.addEventListener("click", ()=>toggleFacetPanel(idx));

    const body = document.createElement("div");
    body.className = "facetBody";
    body.id = `facetBody_${idx}`;

    const opts = document.createElement("div");
    opts.className = "facetOptions";

    const selectedSet = STATE.selected.facets.get(k) || new Set();

    for(const [val, count] of entries){
      const id = `facet_${idx}_${val.replace(/[^a-z0-9]+/g,"_")}`;
      const wrap = document.createElement("label");
      wrap.className = "opt";
      wrap.innerHTML = `<input type="checkbox" id="${id}"><span>${val} <span style="opacity:.7">(${count})</span></span>`;
      const cb = wrap.querySelector("input");
      cb.checked = selectedSet.has(val);
      cb.addEventListener("change", ()=>{
        let set = STATE.selected.facets.get(k);
        if(!set){ set = new Set(); STATE.selected.facets.set(k, set); }
        if(cb.checked) set.add(val); else set.delete(val);
        if(set.size===0) STATE.selected.facets.delete(k);
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

async function loadTypologyAndRender(){
  $("#resultsMeta").textContent="Laden…";
  STATE.selected.facets = new Map(); // reset extra filters on typology/subtype change

  STATE.plants = await loadTypologyPlants(STATE.selected.typology, STATE.selected.subtype);

  const opts=computeBaseFilterOptions(STATE.plants);
  fillSelect($("#soilType"), opts.soil);
  fillSelect($("#soilMoisture"), opts.moist);
  fillSelect($("#acidity"), opts.acid);
  fillSelect($("#spread"), opts.spread);

  await applyLayers();
  renderExtraFilters(STATE.plants);
  render(STATE.plants);
}

async function loadBoboOptions(){
  try{
    const wb = await fetchXlsx("data/layers/bobo_bodemcodes_volledige_lijst.xlsx");
    const sheet = wb.Sheets["alle_codes"] ? "alle_codes" : firstSheetName(wb);
    const rows = sheetToJson(wb, sheet);
    const groups = uniqSorted(rows.map(r=>norm(r["groep"])).filter(Boolean));
    const codesBy = new Map();
    for(const g of groups){
      const codes = rows.filter(r=>norm(r["groep"])===g).map(r=>norm(r["code"])).filter(Boolean);
      codesBy.set(g, uniqSorted(codes));
    }
    STATE.options.bobo.groups = groups;
    STATE.options.bobo.codesByGroup = codesBy;
  }catch(e){
    console.warn("BOBO options not found.", e);
    STATE.options.bobo.groups = [];
    STATE.options.bobo.codesByGroup = new Map();
  }
}

async function loadRegionOptions(){
  try{
    const wb = await fetchXlsx("data/layers/List_options_regionale_soortenlijst.xlsx");
    const rows = sheetToJson(wb, firstSheetName(wb));
    const districts = rows.map(r=>norm(r["Districten"] || r["districten"])).filter(Boolean);
    STATE.options.region.districts = uniqSorted(districts);
  }catch(e){
    console.warn("Region options not found.", e);
    STATE.options.region.districts = [];
  }
}

function wire(){
  $("#typology").addEventListener("change", async (e)=>{
    STATE.selected.typology = e.target.value;
    const subs=subtypeOptions(STATE.selected.typology);
    const subSel=$("#subtype");
    subSel.innerHTML="";
    for(const s of subs){
      const opt=document.createElement("option");
      opt.value=s;
      opt.textContent=s;
      subSel.appendChild(opt);
    }
    STATE.selected.subtype=subs[0];
    subSel.value=STATE.selected.subtype;
    await loadTypologyAndRender();
  });

  $("#subtype").addEventListener("change", async (e)=>{
    STATE.selected.subtype=e.target.value;
    await loadTypologyAndRender();
  });

  $("#soilType").addEventListener("change",(e)=>{ STATE.selected.soilType=e.target.value; render(STATE.plants); });
  $("#soilMoisture").addEventListener("change",(e)=>{ STATE.selected.soilMoisture=e.target.value; render(STATE.plants); });
  $("#acidity").addEventListener("change",(e)=>{ STATE.selected.acidity=e.target.value; render(STATE.plants); });
  $("#spread").addEventListener("change",(e)=>{ STATE.selected.spread=e.target.value; render(STATE.plants); });

  $("#layer_klimaat").addEventListener("change", async (e)=>{ STATE.selected.layers.klimaat=e.target.checked; await applyLayers(); render(STATE.plants); });
  $("#layer_amber").addEventListener("change", async (e)=>{ STATE.selected.layers.amber=e.target.checked; await applyLayers(); render(STATE.plants); });

  $("#layer_regionaal").addEventListener("change", async (e)=>{
    STATE.selected.layers.regionaal=e.target.checked;
    $("#districtWrap").style.display = e.target.checked ? "block" : "none";
    await applyLayers();
    render(STATE.plants);
  });

  $("#district").addEventListener("change", async (e)=>{
    STATE.selected.district=e.target.value;
    await applyLayers();
    render(STATE.plants);
  });

  const paint = $("#regionModePaint");
  const filt = $("#regionModeFilter");
  if(paint && filt){
    paint.addEventListener("change", ()=>{
      if(paint.checked){ STATE.selected.regionMode="paint"; render(STATE.plants); }
    });
    filt.addEventListener("change", ()=>{
      if(filt.checked){ STATE.selected.regionMode="filter"; render(STATE.plants); }
    });
  }

  $("#layer_bobo").addEventListener("change", async (e)=>{
    STATE.selected.layers.bobo=e.target.checked;
    $("#boboWrap").style.display = e.target.checked ? "block" : "none";
    $("#boboCodeWrap").style.display = e.target.checked ? "block" : "none";
    await applyLayers(); render(STATE.plants);
  });

  $("#boboGroup").addEventListener("change", async (e)=>{
    STATE.selected.boboGroup = e.target.value;
    const codes = STATE.options.bobo.codesByGroup.get(STATE.selected.boboGroup) || [];
    fillSelect($("#boboCode"), codes, true);
    STATE.selected.boboCode = $("#boboCode").value || "ALLE";
    await applyLayers();
    render(STATE.plants);
  });

  $("#boboCode").addEventListener("change", async (e)=>{
    STATE.selected.boboCode = e.target.value;
    await applyLayers();
    render(STATE.plants);
  });

  $("#layer_fyto").addEventListener("change", async (e)=>{
    STATE.selected.layers.fyto=e.target.checked;
    await applyLayers();
    render(STATE.plants);
  });

  $("#exportCsv").addEventListener("click", ()=>{
    const filtered = STATE.plants.filter(matchesAll);
    downloadText(`planten_${STATE.selected.typology}_${STATE.selected.subtype}.csv`, toCsv(filtered));
  });
}

async function init(){
  await loadConfig();
  await loadBoboOptions();
  await loadRegionOptions();

  const typSel=$("#typology");
  typSel.innerHTML="";
  for(const t of typologyOptions()){
    const opt=document.createElement("option");
    opt.value=t; opt.textContent=t;
    typSel.appendChild(opt);
  }
  STATE.selected.typology = typologyOptions()[0];
  typSel.value = STATE.selected.typology;

  const subSel=$("#subtype");
  subSel.innerHTML="";
  const subs=subtypeOptions(STATE.selected.typology);
  for(const s of subs){
    const opt=document.createElement("option");
    opt.value=s; opt.textContent=s;
    subSel.appendChild(opt);
  }
  STATE.selected.subtype=subs[0];
  subSel.value=STATE.selected.subtype;

  const districtSel=$("#district");
  districtSel.innerHTML="";
  const opt0=document.createElement("option");
  opt0.value=""; opt0.textContent="Kies district…";
  districtSel.appendChild(opt0);
  for(const d of STATE.options.region.districts){
    const opt=document.createElement("option");
    opt.value=d; opt.textContent=d;
    districtSel.appendChild(opt);
  }
  STATE.selected.district = "";

  const bg=$("#boboGroup");
  bg.innerHTML="";
  const optG=document.createElement("option");
  optG.value=""; optG.textContent="Kies groep…";
  bg.appendChild(optG);
  for(const g of STATE.options.bobo.groups){
    const opt=document.createElement("option");
    opt.value=g; opt.textContent=g;
    bg.appendChild(opt);
  }
  STATE.selected.boboGroup="";
  fillSelect($("#boboCode"), [], true);
  STATE.selected.boboCode="ALLE";

  wire();
  await loadTypologyAndRender();
}

window.addEventListener("DOMContentLoaded", init);
