/* =========================
   Storage / Utilities
========================= */
const LS_KEY = "CERT_MVP_V1";

function uid(prefix="id"){
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}

function todayISO(){
  const d = new Date();
  return d.toISOString().slice(0,10);
}

function daysBetween(aISO, bISO){
  if(!aISO || !bISO) return null;
  const a = new Date(aISO);
  const b = new Date(bISO);
  return Math.floor((b - a) / (1000*60*60*24));
}

function clampStr(s, n=40){
  if(!s) return "";
  return s.length > n ? s.slice(0,n-1) + "…" : s;
}

function saveState(state){
  localStorage.setItem(LS_KEY, JSON.stringify(state));
}
function loadState(){
  const raw = localStorage.getItem(LS_KEY);
  if(raw){
    try { return JSON.parse(raw); } catch(e){}
  }
  // default seed
  return {
    settings: { dueDays: 60 },
    models: [],            // {model_id, product_code, model_name, gas_type}
    doctypes: [],          // {doctype_id, category, name, org, defaultRenewMonths}
    requirements: {},      // key: `${model_id}__${doctype_id}` => {status, reason, updatedAt}
    documents: [],         // {document_id, doctype_id, title, issuer, issued, expiry, renewMonths, scope, plant, memo, file:{name,dataUrl}?}
    documentModelMap: []   // {document_id, model_id}
  };
}

let state = loadState();

/* =========================
   DOM Helpers
========================= */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function renderTable(el, cols, rows){
  const thead = `<thead><tr>${cols.map(c=>`<th>${c.label}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(r=>{
    return `<tr>${cols.map(c=>`<td>${c.render(r)}</td>`).join("")}</tr>`;
  }).join("")}</tbody>`;
  el.innerHTML = thead + tbody;
}

function downloadBlob(filename, blob){
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href), 500);
}

function badgeForStatus(s){
  if(s==="VALID") return `<span class="badge b-valid">유효</span>`;
  if(s==="DUE") return `<span class="badge b-due">임박</span>`;
  if(s==="EXPIRED") return `<span class="badge b-expired">만료</span>`;
  return "";
}

function docStatus(doc){
  if(!doc.expiry) return "VALID";
  const dueDays = state.settings.dueDays ?? 60;
  const d = daysBetween(todayISO(), doc.expiry);
  if(d < 0) return "EXPIRED";
  if(d <= dueDays) return "DUE";
  return "VALID";
}

/* =========================
   Tabs
========================= */
function initTabs(){
  $$(".nav__item").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      $$(".nav__item").forEach(b=>b.classList.remove("is-active"));
      btn.classList.add("is-active");
      const tabId = btn.dataset.tab;
      $$(".tab").forEach(t=>t.classList.remove("is-active"));
      $("#"+tabId).classList.add("is-active");
      // refresh on enter
      refreshAll();
    });
  });
}

/* =========================
   Models: Import / CRUD
========================= */
function normalizeHeader(h){
  if(!h) return "";
  const s = String(h).trim().toLowerCase();
  // map korean & english variants
  if(["제품코드","product_code","productcode","코드"].includes(s)) return "product_code";
  if(["모델명","model_name","modelname","모델"].includes(s)) return "model_name";
  if(["가스구분","gas_type","gastype","가스"].includes(s)) return "gas_type";
  return s;
}

async function importModelsFromExcel(file){
  if(!file) throw new Error("파일이 없다.");

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" }); // first row as headers

  // header normalize
  const mapped = json.map(row=>{
    const out = {};
    Object.keys(row).forEach(k=>{
      out[normalizeHeader(k)] = row[k];
    });
    return out;
  });

  let inserted=0, updated=0, skipped=0;
  mapped.forEach(r=>{
    const product_code = String(r.product_code || "").trim();
    const model_name = String(r.model_name || "").trim();
    const gas_type = String(r.gas_type || "").trim();
    if(!product_code || !model_name){
      skipped++;
      return;
    }
    const idx = state.models.findIndex(m=>m.product_code === product_code);
    if(idx >= 0){
      state.models[idx] = { ...state.models[idx], model_name, gas_type };
      updated++;
    }else{
      state.models.push({
        model_id: uid("m"),
        product_code,
        model_name,
        gas_type
      });
      inserted++;
    }
  });

  saveState(state);
  return {inserted, updated, skipped, total: mapped.length};
}

function addModelManual(){
  const product_code = prompt("제품코드 입력");
  if(!product_code) return;
  const model_name = prompt("모델명 입력");
  if(!model_name) return;
  const gas_type = prompt("가스구분 입력(선택)");
  const exists = state.models.some(m=>m.product_code===product_code.trim());
  if(exists){
    alert("이미 존재하는 제품코드이다.");
    return;
  }
  state.models.push({
    model_id: uid("m"),
    product_code: product_code.trim(),
    model_name: model_name.trim(),
    gas_type: (gas_type||"").trim()
  });
  saveState(state);
  refreshAll();
}

function deleteModel(model_id){
  if(!confirm("이 모델을 삭제하나? (연결된 매트릭스/매핑도 함께 정리 권장)")) return;
  state.models = state.models.filter(m=>m.model_id!==model_id);
  // remove requirements for this model
  Object.keys(state.requirements).forEach(k=>{
    if(k.startsWith(model_id+"__")) delete state.requirements[k];
  });
  // remove doc mappings
  state.documentModelMap = state.documentModelMap.filter(x=>x.model_id!==model_id);
  saveState(state);
  refreshAll();
}

/* =========================
   DocTypes CRUD
========================= */
function addDocType(){
  const category = $("#dtCategory").value;
  const name = $("#dtName").value.trim();
  const org = $("#dtOrg").value.trim();
  const renew = $("#dtRenew").value ? Number($("#dtRenew").value) : null;
  if(!name){
    alert("문서종류명을 입력해야 한다.");
    return;
  }
  const dup = state.doctypes.some(d=>d.name.toLowerCase()===name.toLowerCase() && d.category===category);
  if(dup){
    alert("동일한 DocType이 이미 존재한다.");
    return;
  }
  state.doctypes.push({
    doctype_id: uid("dt"),
    category,
    name,
    org,
    defaultRenewMonths: renew
  });
  $("#dtName").value="";
  $("#dtOrg").value="";
  $("#dtRenew").value="";
  saveState(state);
  refreshAll();
}

function deleteDocType(doctype_id){
  if(!confirm("DocType을 삭제하나? 연결된 매트릭스/문서에도 영향이 있다.")) return;
  state.doctypes = state.doctypes.filter(d=>d.doctype_id!==doctype_id);
  // requirements cleanup
  Object.keys(state.requirements).forEach(k=>{
    const [, dt] = k.split("__");
    if(dt===doctype_id) delete state.requirements[k];
  });
  // documents cleanup (or keep but broken) -> remove related docs
  const docIds = state.documents.filter(d=>d.doctype_id===doctype_id).map(d=>d.document_id);
  state.documents = state.documents.filter(d=>d.doctype_id!==doctype_id);
  state.documentModelMap = state.documentModelMap.filter(x=>!docIds.includes(x.document_id));
  saveState(state);
  refreshAll();
}

/* =========================
   Requirements Matrix
========================= */
const REQ_STATUSES = ["REQUIRED","NOT_REQUIRED","NOT_POSSIBLE","TBD"];

function setRequirement(model_id, doctype_id, status){
  const key = `${model_id}__${doctype_id}`;
  state.requirements[key] = {
    status,
    updatedAt: new Date().toISOString()
  };
  saveState(state);
}

function bulkFillTBDForFilteredModels(){
  const keyword = ($("#mxModelFilter").value||"").trim().toLowerCase();
  const dtId = $("#mxDocTypeFilter").value;
  const models = state.models.filter(m=>{
    if(!keyword) return true;
    return (m.product_code+" "+m.model_name+" "+(m.gas_type||"")).toLowerCase().includes(keyword);
  });
  if(models.length===0){ alert("대상 모델이 없다."); return; }
  if(!dtId){ alert("DocType을 선택해야 한다."); return; }

  models.forEach(m=>{
    const k = `${m.model_id}__${dtId}`;
    if(!state.requirements[k]){
      state.requirements[k] = { status:"TBD", updatedAt:new Date().toISOString() };
    }
  });
  saveState(state);
  refreshAll();
}

/* =========================
   Documents / Mapping / File
========================= */
let pickedModelIds = new Set();

function updateScopeUI(){
  const scope = $("#docScope").value;
  $("#plantBox").classList.toggle("hidden", scope !== "PLANT");
}

function renderModelPicker(){
  const keyword = ($("#docModelSearch").value||"").trim().toLowerCase();
  const list = $("#docModelPickList");
  const chips = $("#docModelChips");

  const filtered = state.models
    .filter(m=>{
      if(!keyword) return true;
      return (m.product_code+" "+m.model_name+" "+(m.gas_type||"")).toLowerCase().includes(keyword);
    })
    .slice(0, 50);

  list.innerHTML = filtered.map(m=>{
    const picked = pickedModelIds.has(m.model_id);
    return `<div class="pick__item" data-mid="${m.model_id}">
      ${picked ? "✅ " : ""}<b>${m.product_code}</b> — ${clampStr(m.model_name, 30)} <span class="muted">(${m.gas_type||"-"})</span>
    </div>`;
  }).join("");

  list.querySelectorAll(".pick__item").forEach(el=>{
    el.addEventListener("click", ()=>{
      const mid = el.dataset.mid;
      if(pickedModelIds.has(mid)) pickedModelIds.delete(mid);
      else pickedModelIds.add(mid);
      renderModelPicker();
    });
  });

  const picked = state.models.filter(m=>pickedModelIds.has(m.model_id));
  chips.innerHTML = picked.map(m=>`
    <span class="chip">
      ${m.product_code}
      <button data-mid="${m.model_id}">×</button>
    </span>
  `).join("");
  chips.querySelectorAll("button").forEach(b=>{
    b.addEventListener("click", ()=>{
      pickedModelIds.delete(b.dataset.mid);
      renderModelPicker();
    });
  });
}

function readFileAsDataUrl(file){
  return new Promise((resolve, reject)=>{
    if(!file) return resolve(null);
    const reader = new FileReader();
    reader.onload = ()=> resolve(String(reader.result));
    reader.onerror = ()=> reject(new Error("파일 읽기 실패"));
    reader.readAsDataURL(file);
  });
}

async function addDocument(){
  const doctype_id = $("#docTypeSelect").value;
  const title = $("#docTitle").value.trim();
  const issuer = $("#docIssuer").value.trim();
  const issued = $("#docIssued").value || null;
  const expiry = $("#docExpiry").value || null;
  const renewMonths = $("#docRenew").value ? Number($("#docRenew").value) : null;
  const scope = $("#docScope").value;
  const plant = ($("#docPlant").value||"").trim() || null;
  const memo = ($("#docMemo").value||"").trim() || null;

  if(!doctype_id){ alert("DocType을 선택해야 한다."); return; }
  if(!title){ alert("문서명을 입력해야 한다."); return; }
  if(scope==="PLANT" && !plant){ alert("공장 범위면 공장 값을 입력해야 한다."); return; }
  if(scope==="MODEL" && pickedModelIds.size===0){
    alert("모델(품목) 범위면 적용 모델을 1개 이상 선택해야 한다.");
    return;
  }

  const file = $("#docFile").files?.[0] || null;
  const dataUrl = await readFileAsDataUrl(file);

  const document_id = uid("doc");
  state.documents.push({
    document_id,
    doctype_id,
    title,
    issuer,
    issued,
    expiry,
    renewMonths,
    scope,
    plant,
    memo,
    file: file ? { name: file.name, dataUrl } : null,
    createdAt: new Date().toISOString()
  });

  // mapping
  if(pickedModelIds.size>0){
    pickedModelIds.forEach(mid=>{
      state.documentModelMap.push({ document_id, model_id: mid });
    });
  }

  saveState(state);

  // reset form
  $("#docTitle").value="";
  $("#docIssuer").value="";
  $("#docIssued").value="";
  $("#docExpiry").value="";
  $("#docRenew").value="";
  $("#docMemo").value="";
  $("#docPlant").value="";
  $("#docFile").value="";
  pickedModelIds = new Set();
  renderModelPicker();
  refreshAll();
}

function deleteDocument(document_id){
  if(!confirm("문서를 삭제하나?")) return;
  state.documents = state.documents.filter(d=>d.document_id!==document_id);
  state.documentModelMap = state.documentModelMap.filter(x=>x.document_id!==document_id);
  saveState(state);
  refreshAll();
}

function openDocumentFile(document_id){
  const doc = state.documents.find(d=>d.document_id===document_id);
  if(!doc || !doc.file || !doc.file.dataUrl){
    alert("첨부 파일이 없다.");
    return;
  }
  const a = document.createElement("a");
  a.href = doc.file.dataUrl;
  a.download = doc.file.name || "document";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/* =========================
   Renderers
========================= */
function renderModels(){
  const el = $("#modelTable");
  const keyword = ($("#modelFilter").value||"").trim().toLowerCase();
  const rows = state.models
    .filter(m=>{
      if(!keyword) return true;
      return (m.product_code+" "+m.model_name+" "+(m.gas_type||"")).toLowerCase().includes(keyword);
    })
    .sort((a,b)=>a.product_code.localeCompare(b.product_code));

  const cols = [
    {label:"제품코드", render:r=>`<b>${r.product_code}</b>`},
    {label:"모델명", render:r=>clampStr(r.model_name, 28)},
    {label:"가스구분", render:r=>r.gas_type||"-"},
    {label:"", render:r=>`<button class="btn" data-del="${r.model_id}">삭제</button>`}
  ];
  renderTable(el, cols, rows);
  el.querySelectorAll("button[data-del]").forEach(b=>{
    b.addEventListener("click", ()=>deleteModel(b.dataset.del));
  });
}

function renderDocTypes(){
  const el = $("#dtTable");
  const keyword = ($("#dtFilter").value||"").trim().toLowerCase();
  const rows = state.doctypes
    .filter(d=>{
      if(!keyword) return true;
      return (d.name+" "+(d.org||"")+" "+d.category).toLowerCase().includes(keyword);
    })
    .sort((a,b)=>a.category.localeCompare(b.category) || a.name.localeCompare(b.name));
  const cols = [
    {label:"구분", render:r=>r.category},
    {label:"DocType", render:r=>`<b>${r.name}</b>`},
    {label:"기관", render:r=>r.org||"-"},
    {label:"기본갱신(개월)", render:r=>r.defaultRenewMonths ?? "-"},
    {label:"", render:r=>`<button class="btn" data-del="${r.doctype_id}">삭제</button>`},
  ];
  renderTable(el, cols, rows);
  el.querySelectorAll("button[data-del]").forEach(b=>{
    b.addEventListener("click", ()=>deleteDocType(b.dataset.del));
  });

  // selects
  const mxSel = $("#mxDocTypeFilter");
  const docSel = $("#docTypeSelect");

  const opts = state.doctypes.map(d=>`<option value="${d.doctype_id}">[${d.category}] ${d.name}</option>`).join("");
  mxSel.innerHTML = `<option value="">모든 DocType</option>` + opts;
  docSel.innerHTML = `<option value="">선택</option>` + opts;
}

function renderMatrix(){
  const el = $("#matrixTable");

  const modelKw = ($("#mxModelFilter").value||"").trim().toLowerCase();
  const dtId = $("#mxDocTypeFilter").value;

  const models = state.models
    .filter(m=>{
      if(!modelKw) return true;
      return (m.product_code+" "+m.model_name+" "+(m.gas_type||"")).toLowerCase().includes(modelKw);
    })
    .slice(0, 80); // MVP: 성능 제한

  const doctypes = state.doctypes
    .filter(d=>!dtId || d.doctype_id===dtId)
    .slice(0, 30);

  if(models.length===0 || doctypes.length===0){
    el.innerHTML = "<thead><tr><th>표시할 데이터가 없다</th></tr></thead>";
    return;
  }

  // table header
  const thead = `<thead><tr>
    <th>제품코드</th><th>모델명</th><th>가스</th>
    ${doctypes.map(d=>`<th>${clampStr(d.name, 18)}</th>`).join("")}
  </tr></thead>`;

  // body
  const tbodyRows = models.map(m=>{
    const tds = doctypes.map(d=>{
      const key = `${m.model_id}__${d.doctype_id}`;
      const cur = state.requirements[key]?.status || "";
      const sel = `<select data-mid="${m.model_id}" data-dt="${d.doctype_id}" class="input" style="padding:6px;border-radius:8px;">
        <option value="">(미설정)</option>
        ${REQ_STATUSES.map(s=>`<option value="${s}" ${cur===s?"selected":""}>${s}</option>`).join("")}
      </select>`;
      return `<td>${sel}</td>`;
    }).join("");

    return `<tr>
      <td><b>${m.product_code}</b></td>
      <td>${clampStr(m.model_name, 22)}</td>
      <td>${m.gas_type||"-"}</td>
      ${tds}
    </tr>`;
  }).join("");

  el.innerHTML = thead + `<tbody>${tbodyRows}</tbody>`;

  el.querySelectorAll("select[data-mid]").forEach(s=>{
    s.addEventListener("change", ()=>{
      const mid = s.dataset.mid;
      const dt = s.dataset.dt;
      const val = s.value;
      if(!val){
        delete state.requirements[`${mid}__${dt}`];
      }else{
        setRequirement(mid, dt, val);
      }
      saveState(state);
    });
  });
}

function renderDocuments(){
  const el = $("#docTable");
  const keyword = ($("#docFilter").value||"").trim().toLowerCase();

  const rows = state.documents
    .map(d=>{
      const dt = state.doctypes.find(x=>x.doctype_id===d.doctype_id);
      const mappedModels = state.documentModelMap
        .filter(x=>x.document_id===d.document_id)
        .map(x=>state.models.find(m=>m.model_id===x.model_id))
        .filter(Boolean);
      const modelText = mappedModels.slice(0,3).map(m=>m.product_code).join(", ") + (mappedModels.length>3 ? ` 외 ${mappedModels.length-3}건` : "");
      return {
        ...d,
        _doctypeName: dt ? dt.name : "(삭제된 DocType)",
        _status: docStatus(d),
        _modelText: modelText || (d.scope==="MODEL" ? "-" : "(전사/공장 범위)")
      };
    })
    .filter(r=>{
      if(!keyword) return true;
      return (r.title+" "+r._doctypeName+" "+(r.issuer||"")).toLowerCase().includes(keyword);
    })
    .sort((a,b)=>(a.expiry||"9999-99-99").localeCompare(b.expiry||"9999-99-99"));

  const cols = [
    {label:"상태", render:r=>badgeForStatus(r._status)},
    {label:"DocType", render:r=>clampStr(r._doctypeName, 18)},
    {label:"문서명", render:r=>`<b title="${r.title}">${clampStr(r.title, 28)}</b>`},
    {label:"범위", render:r=>r.scope + (r.plant ? `(${r.plant})` : "")},
    {label:"만료일", render:r=>r.expiry||"-"},
    {label:"적용모델", render:r=>clampStr(r._modelText, 24)},
    {label:"파일", render:r=>r.file ? `<button class="btn" data-open="${r.document_id}">다운로드</button>` : "-"},
    {label:"", render:r=>`<button class="btn" data-del="${r.document_id}">삭제</button>`},
  ];
  renderTable(el, cols, rows);

  el.querySelectorAll("button[data-del]").forEach(b=>{
    b.addEventListener("click", ()=>deleteDocument(b.dataset.del));
  });
  el.querySelectorAll("button[data-open]").forEach(b=>{
    b.addEventListener("click", ()=>openDocumentFile(b.dataset.open));
  });
}

function renderSearch(){
  const el = $("#searchTable");
  const kw = ($("#qKeyword").value||"").trim().toLowerCase();
  const qModel = ($("#qModel").value||"").trim().toLowerCase();
  const qStatus = $("#qStatus").value;

  const rows = state.documents.map(d=>{
    const dt = state.doctypes.find(x=>x.doctype_id===d.doctype_id);
    const mapped = state.documentModelMap
      .filter(x=>x.document_id===d.document_id)
      .map(x=>state.models.find(m=>m.model_id===x.model_id))
      .filter(Boolean);
    const st = docStatus(d);
    return {
      ...d,
      _doctypeName: dt ? dt.name : "(삭제된 DocType)",
      _status: st,
      _models: mapped
    };
  }).filter(r=>{
    if(qStatus && r._status!==qStatus) return false;
    if(kw){
      const t = (r.title+" "+r._doctypeName+" "+(r.issuer||"")).toLowerCase();
      if(!t.includes(kw)) return false;
    }
    if(qModel){
      const mt = r._models.map(m=>`${m.product_code} ${m.model_name}`).join(" ").toLowerCase();
      const allow = mt.includes(qModel) || (r.scope!=="MODEL" && (r.title||"").toLowerCase().includes(qModel));
      if(!allow) return false;
    }
    return true;
  });

  const cols = [
    {label:"상태", render:r=>badgeForStatus(r._status)},
    {label:"DocType", render:r=>clampStr(r._doctypeName, 18)},
    {label:"문서명", render:r=>clampStr(r.title, 30)},
    {label:"기관", render:r=>r.issuer||"-"},
    {label:"만료일", render:r=>r.expiry||"-"},
    {label:"적용모델", render:r=>{
      if(r.scope!=="MODEL") return "(전사/공장)";
      const codes = r._models.map(m=>m.product_code).slice(0,4);
      return codes.join(", ") + (r._models.length>4 ? ` 외 ${r._models.length-4}건` : "");
    }},
  ];
  renderTable(el, cols, rows);
}

function renderExpiry(){
  const el = $("#expiryTable");
  const rows = state.documents
    .map(d=>{
      const dt = state.doctypes.find(x=>x.doctype_id===d.doctype_id);
      const st = docStatus(d);
      return {
        ...d,
        _doctypeName: dt ? dt.name : "(삭제된 DocType)",
        _status: st,
        _left: d.expiry ? daysBetween(todayISO(), d.expiry) : null
      };
    })
    .filter(r=>r._status==="DUE" || r._status==="EXPIRED")
    .sort((a,b)=>(a._left ?? 999999) - (b._left ?? 999999));

  const cols = [
    {label:"상태", render:r=>badgeForStatus(r._status)},
    {label:"D-일", render:r=>r._left===null ? "-" : r._left},
    {label:"DocType", render:r=>clampStr(r._doctypeName, 18)},
    {label:"문서명", render:r=>clampStr(r.title, 34)},
    {label:"만료일", render:r=>r.expiry||"-"},
    {label:"범위", render:r=>r.scope + (r.plant?`(${r.plant})`:"")},
  ];
  renderTable(el, cols, rows);

  // store summary for mail
  state._lastExpiryRows = rows;
}

function refreshAll(){
  // settings
  $("#dueDays").value = state.settings.dueDays ?? 60;

  renderModels();
  renderDocTypes();
  renderMatrix();

  updateScopeUI();
  renderModelPicker();
  renderDocuments();
  renderExpiry();
}

/* =========================
   Export / Reset / Template
========================= */
function exportJson(){
  const blob = new Blob([JSON.stringify(state, null, 2)], { type:"application/json" });
  downloadBlob("cert_manager_export.json", blob);
}

function resetAll(){
  if(!confirm("모든 데이터를 삭제하고 초기화하나?")) return;
  localStorage.removeItem(LS_KEY);
  state = loadState();
  saveState(state);
  refreshAll();
}

function downloadExcelTemplate(){
  // create workbook
  const ws = XLSX.utils.aoa_to_sheet([
    ["제품코드","모델명","가스구분"],
    ["P001","RANGE-ABC","LNG"],
    ["P002","RANGE-XYZ","LPG"]
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "models");
  const out = XLSX.write(wb, { bookType:"xlsx", type:"array" });
  downloadBlob("model_master_template.xlsx", new Blob([out], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
}

/* =========================
   Mail helpers (manual)
========================= */
function buildExpirySummaryText(){
  const rows = state.documents
    .map(d=>{
      const dt = state.doctypes.find(x=>x.doctype_id===d.doctype_id);
      const st = docStatus(d);
      if(st!=="DUE" && st!=="EXPIRED") return null;
      const left = d.expiry ? daysBetween(todayISO(), d.expiry) : null;
      const dname = dt ? dt.name : "(삭제된 DocType)";
      return `- [${st}] D${left!==null?left:"-"} | ${dname} | ${d.title} | 만료:${d.expiry||"-"} | 범위:${d.scope}${d.plant?`(${d.plant})`:""}`;
    })
    .filter(Boolean);

  const head = `만료/임박 목록 (기준 ${state.settings.dueDays}일, ${todayISO()})`;
  return head + "\n" + (rows.length ? rows.join("\n") : "- 대상 없음");
}

function buildMailto(){
  const emails = ($("#manualEmails").value||"").trim();
  if(!emails){ alert("수신자 이메일을 입력해야 한다."); return; }
  const subject = encodeURIComponent(`[인증서관리] 만료/임박 알림 (${todayISO()})`);
  const body = encodeURIComponent(buildExpirySummaryText());
  const url = `mailto:${encodeURIComponent(emails)}?subject=${subject}&body=${body}`;
  $("#mailtoOut").innerHTML = `<a href="${url}">${clampStr(url, 120)}</a>`;
}

async function copySummary(){
  const txt = buildExpirySummaryText();
  await navigator.clipboard.writeText(txt);
  alert("요약을 클립보드에 복사했다.");
}

/* =========================
   Event bindings
========================= */
function bindEvents(){
  // header
  $("#btnExportJson").addEventListener("click", exportJson);
  $("#btnReset").addEventListener("click", resetAll);

  // models
  $("#btnImportModels").addEventListener("click", async ()=>{
    const file = $("#modelFile").files?.[0];
    try{
      const res = await importModelsFromExcel(file);
      $("#modelImportResult").textContent =
        `총 ${res.total}행 처리: 신규 ${res.inserted}, 업데이트 ${res.updated}, 스킵 ${res.skipped}`;
      refreshAll();
    }catch(e){
      alert(e.message || "업로드 실패");
    }
  });
  $("#btnDownloadTemplate").addEventListener("click", downloadExcelTemplate);
  $("#btnAddModel").addEventListener("click", addModelManual);
  $("#modelFilter").addEventListener("input", renderModels);

  // doctypes
  $("#btnAddDocType").addEventListener("click", addDocType);
  $("#dtFilter").addEventListener("input", renderDocTypes);

  // matrix
  $("#mxModelFilter").addEventListener("input", renderMatrix);
  $("#mxDocTypeFilter").addEventListener("change", renderMatrix);
  $("#btnMatrixBulkFill").addEventListener("click", bulkFillTBDForFilteredModels);

  // documents
  $("#docScope").addEventListener("change", ()=>{
    updateScopeUI();
  });
  $("#docModelSearch").addEventListener("input", renderModelPicker);
  $("#btnAddDocument").addEventListener("click", addDocument);
  $("#docFilter").addEventListener("input", renderDocuments);

  // search
  $("#btnSearch").addEventListener("click", renderSearch);

  // expiry
  $("#btnRecalcDue").addEventListener("click", ()=>{
    const v = Number($("#dueDays").value);
    if(Number.isNaN(v) || v<0){ alert("0 이상의 숫자여야 한다."); return; }
    state.settings.dueDays = v;
    saveState(state);
    renderExpiry();
  });
  $("#btnBuildMailto").addEventListener("click", buildMailto);
  $("#btnCopySummary").addEventListener("click", copySummary);
}

/* =========================
   Init
========================= */
function init(){
  initTabs();
  bindEvents();
  // seed: if empty, add a couple of doctypes for convenience
  if(state.doctypes.length===0){
    state.doctypes.push(
      { doctype_id: uid("dt"), category:"CERT", name:"ISO 9001", org:"", defaultRenewMonths:36 },
      { doctype_id: uid("dt"), category:"CERT", name:"KS B 8114", org:"", defaultRenewMonths:null },
      { doctype_id: uid("dt"), category:"TEST_REPORT", name:"연소성능 시험성적서", org:"", defaultRenewMonths:null },
    );
    saveState(state);
  }
  refreshAll();
}
init();
