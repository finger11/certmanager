const LS_KEY = "CERT_SYS_V1";

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function uid(prefix="id"){
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}
function todayISO(){ return new Date().toISOString().slice(0,10); }
function daysBetween(aISO, bISO){
  if(!aISO || !bISO) return null;
  const a = new Date(aISO), b = new Date(bISO);
  return Math.floor((b - a) / (1000*60*60*24));
}
function saveState(s){ localStorage.setItem(LS_KEY, JSON.stringify(s)); }
function loadState(){
  const raw = localStorage.getItem(LS_KEY);
  if(raw){
    try{ return JSON.parse(raw); } catch {}
  }
  return {
    models: [],        // {model_id, product_code, model_name, gas_type}
    certs: [],         // {cert_id, cert_no, type, issuer, issued, valid_from, valid_to, memo, file:{name,dataUrl}? , created_at}
    certModelMap: []   // {cert_id, model_id}  (N:1 / N:M 모두 커버)
  };
}
let state = loadState();

// ===== modal helpers
function openModal(id){ $("#"+id).classList.add("is-open"); }
function closeModal(id){ $("#"+id).classList.remove("is-open"); }
function bindModalClose(){
  $$("[data-close]").forEach(b=>{
    b.addEventListener("click", ()=> closeModal(b.dataset.close));
  });
  // backdrop click close
  $$(".modalBack").forEach(back=>{
    back.addEventListener("click",(e)=>{
      if(e.target === back) back.classList.remove("is-open");
    });
  });
}

// ===== Excel import (models)
function normalizeHeader(h){
  const s = String(h||"").trim().toLowerCase();
  if(["제품코드","product_code","productcode","코드"].includes(s)) return "product_code";
  if(["모델명","model_name","modelname","모델"].includes(s)) return "model_name";
  if(["가스구분","gas_type","gastype","가스"].includes(s)) return "gas_type";
  return s;
}
async function importModelsFromExcel(file){
  if(!file) throw new Error("엑셀 파일이 없다.");
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(ws, { defval:"" });

  const mapped = json.map(row=>{
    const out = {};
    Object.keys(row).forEach(k=>{
      out[normalizeHeader(k)] = row[k];
    });
    return out;
  });

  let inserted=0, updated=0, skipped=0;
  mapped.forEach(r=>{
    const product_code = String(r.product_code||"").trim();
    const model_name = String(r.model_name||"").trim();
    const gas_type = String(r.gas_type||"").trim();
    if(!product_code || !model_name){ skipped++; return; }

    const idx = state.models.findIndex(m=>m.product_code===product_code);
    if(idx>=0){
      state.models[idx] = { ...state.models[idx], model_name, gas_type };
      updated++;
    }else{
      state.models.push({ model_id: uid("m"), product_code, model_name, gas_type });
      inserted++;
    }
  });

  saveState(state);
  return { total:mapped.length, inserted, updated, skipped };
}

// template
function downloadModelTemplate(){
  const ws = XLSX.utils.aoa_to_sheet([
    ["제품코드","모델명","가스구분"],
    ["158060002","RC620-22KF","LNG"],
    ["158070002","RC620-27KF","LNG"]
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "models");
  const out = XLSX.write(wb, { bookType:"xlsx", type:"array" });
  const blob = new Blob([out], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "model_master_template.xlsx";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href), 500);
}

// ===== basic render helpers
function badgeForStatus(st){
  if(st==="VALID") return `<span class="badge b-valid">유효</span>`;
  if(st==="DUE") return `<span class="badge b-due">임박</span>`;
  if(st==="EXPIRED") return `<span class="badge b-expired">만료</span>`;
  return "";
}
function calcStatus(cert){
  if(!cert.valid_to) return "VALID";
  const d = daysBetween(todayISO(), cert.valid_to);
  if(d < 0) return "EXPIRED";
  if(d <= 60) return "DUE";
  return "VALID";
}
function getModelsByCert(cert_id){
  const mids = state.certModelMap.filter(x=>x.cert_id===cert_id).map(x=>x.model_id);
  return mids.map(mid=>state.models.find(m=>m.model_id===mid)).filter(Boolean);
}

// ===== render Model Master table
function renderModels(){
  const kw = ($("#modelFilter").value||"").trim().toLowerCase();
  const rows = state.models
    .filter(m=>{
      if(!kw) return true;
      return (m.product_code+" "+m.model_name+" "+(m.gas_type||"")).toLowerCase().includes(kw);
    })
    .sort((a,b)=>a.product_code.localeCompare(b.product_code));

  const el = $("#modelTable");
  el.innerHTML = `
    <thead><tr>
      <th>제품코드</th><th>모델명</th><th>가스구분</th><th></th>
    </tr></thead>
    <tbody>
      ${rows.map(m=>`
        <tr>
          <td><b>${m.product_code}</b></td>
          <td>${m.model_name}</td>
          <td>${m.gas_type||"-"}</td>
          <td><button class="btn" data-del-model="${m.model_id}">삭제</button></td>
        </tr>
      `).join("")}
    </tbody>
  `;
  el.querySelectorAll("[data-del-model]").forEach(b=>{
    b.addEventListener("click", ()=>{
      const id = b.dataset.delModel;
      if(!confirm("모델을 삭제하나? (연결된 적용대상 매핑도 삭제됨)")) return;
      state.models = state.models.filter(m=>m.model_id!==id);
      state.certModelMap = state.certModelMap.filter(x=>x.model_id!==id);
      saveState(state);
      renderModels();
      renderCertList();
      renderTargetPicker(); // 선택창도 갱신
    });
  });
}

// manual add
function addModelManual(){
  const product_code = prompt("제품코드 입력");
  if(!product_code) return;
  const model_name = prompt("모델명 입력");
  if(!model_name) return;
  const gas_type = prompt("가스구분(선택)");
  if(state.models.some(m=>m.product_code===product_code.trim())){
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
  renderModels();
  renderTargetPicker();
}

// ===== Cert registration state (picked targets)
let pickedModelIds = new Set();
function updatePickedCount(){
  $("#pickedCount").textContent = `선택 ${pickedModelIds.size}건`;
}

// ===== target picker render
function renderTargetPicker(){
  const kw = ($("#tpKeyword").value||"").trim().toLowerCase();
  const rows = state.models
    .filter(m=>{
      if(!kw) return true;
      return (m.product_code+" "+m.model_name+" "+(m.gas_type||"")).toLowerCase().includes(kw);
    })
    .sort((a,b)=>a.product_code.localeCompare(b.product_code));

  const el = $("#targetTable");
  el.innerHTML = `
    <thead><tr>
      <th>제품코드</th><th>모델명</th><th>가스구분</th><th>적용구분</th>
    </tr></thead>
    <tbody>
      ${rows.map(m=>{
        const checked = pickedModelIds.has(m.model_id) ? "checked" : "";
        return `
          <tr>
            <td><b>${m.product_code}</b></td>
            <td>${m.model_name}</td>
            <td>${m.gas_type||"-"}</td>
            <td style="text-align:center">
              <input type="checkbox" data-pick-mid="${m.model_id}" ${checked}/>
            </td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  el.querySelectorAll("[data-pick-mid]").forEach(chk=>{
    chk.addEventListener("change", ()=>{
      const mid = chk.dataset.pickMid;
      if(chk.checked) pickedModelIds.add(mid);
      else pickedModelIds.delete(mid);
      updatePickedCount();
    });
  });
}

// apply picked targets
function applyPickedTargets(){
  updatePickedCount();
  closeModal("modalTargetPicker");
}

// ===== file to dataURL
function readFileAsDataUrl(file){
  return new Promise((resolve,reject)=>{
    if(!file) return resolve(null);
    const reader = new FileReader();
    reader.onload = ()=> resolve(String(reader.result));
    reader.onerror = ()=> reject(new Error("파일 읽기 실패"));
    reader.readAsDataURL(file);
  });
}

// ===== register cert
async function saveCert(){
  if(pickedModelIds.size===0){
    alert("적용대상을 1개 이상 선택해야 한다.");
    return;
  }

  const cert_no = $("#regCertNo").value.trim();
  const type = $("#regType").value.trim();
  const issuer = $("#regIssuer").value.trim();
  const issued = $("#regIssued").value || null;
  const valid_from = $("#regValidFrom").value || null;

  const noExpiry = $("#regNoExpiry").checked;
  const valid_to = noExpiry ? null : ($("#regValidTo").value || null);

  if(!cert_no || !type || !issuer){
    alert("성적서(인증서)번호 / 종류 / 발급기관은 필수이다.");
    return;
  }
  if(!noExpiry && (!valid_from || !valid_to)){
    alert("유효기간 없음이 아니라면 From/To를 입력해야 한다.");
    return;
  }

  const file = $("#regFile").files?.[0] || null;
  const dataUrl = await readFileAsDataUrl(file);

  const cert_id = uid("c");
  state.certs.push({
    cert_id,
    cert_no,
    type,
    issuer,
    issued,
    valid_from,
    valid_to,
    memo: ($("#regMemo").value||"").trim() || null,
    file: file ? { name:file.name, dataUrl } : null,
    created_at: new Date().toISOString()
  });

  pickedModelIds.forEach(mid=>{
    state.certModelMap.push({ cert_id, model_id: mid });
  });

  saveState(state);

  // reset registration form
  $("#regCertNo").value="";
  $("#regType").value="";
  $("#regIssuer").value="";
  $("#regIssued").value="";
  $("#regValidFrom").value="";
  $("#regValidTo").value="";
  $("#regNoExpiry").checked=false;
  $("#regFile").value="";
  $("#regFilePath").value="";
  $("#regMemo").value="";
  pickedModelIds = new Set();
  updatePickedCount();

  closeModal("modalRegMaster");
  renderCertList();
}

// ===== file download
function downloadFile(cert_id){
  const cert = state.certs.find(c=>c.cert_id===cert_id);
  if(!cert?.file?.dataUrl){
    alert("첨부파일이 없다.");
    return;
  }
  const a = document.createElement("a");
  a.href = cert.file.dataUrl;
  a.download = cert.file.name || "file";
  document.body.appendChild(a); a.click(); a.remove();
}

// ===== view targets popup
function openTargetView(cert_id){
  const models = getModelsByCert(cert_id);
  const el = $("#targetViewTable");
  el.innerHTML = `
    <thead><tr>
      <th>제품코드</th><th>모델명</th><th>가스구분</th>
    </tr></thead>
    <tbody>
      ${models.map(m=>`
        <tr>
          <td><b>${m.product_code}</b></td>
          <td>${m.model_name}</td>
          <td>${m.gas_type||"-"}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
  openModal("modalTargetView");
}

// ===== search & list
function renderCertList(){
  const fModelName = ($("#fModelName").value||"").trim().toLowerCase();
  const fIssuer = ($("#fIssuer").value||"").trim().toLowerCase();
  const fType = ($("#fType").value||"").trim().toLowerCase();
  const fGas = ($("#fGas").value||"").trim().toLowerCase();
  const fFrom = $("#fValidFrom").value || null;
  const fTo = $("#fValidTo").value || null;

  const rows = state.certs
    .map(c=>{
      const models = getModelsByCert(c.cert_id);
      const gasSet = Array.from(new Set(models.map(m=>m.gas_type||"").filter(Boolean)));
      const status = calcStatus(c);
      return { ...c, _models: models, _gasSet: gasSet, _status: status };
    })
    .filter(r=>{
      if(fIssuer && !(r.issuer||"").toLowerCase().includes(fIssuer)) return false;
      if(fType && !(r.type||"").toLowerCase().includes(fType)) return false;

      if(fModelName){
        const t = r._models.map(m=>m.model_name).join(" ").toLowerCase();
        if(!t.includes(fModelName)) return false;
      }
      if(fGas){
        const g = r._gasSet.join(" ").toLowerCase();
        if(!g.includes(fGas)) return false;
      }

      // 기간 필터: 유효기간(To)이 존재할 때만 비교
      if((fFrom || fTo) && r.valid_from && r.valid_to){
        if(fFrom && r.valid_to < fFrom) return false;
        if(fTo && r.valid_from > fTo) return false;
      }
      return true;
    })
    .sort((a,b)=>(a.valid_to||"9999-99-99").localeCompare(b.valid_to||"9999-99-99"));

  const el = $("#certTable");
  el.innerHTML = `
    <thead><tr>
      <th>적용대상</th>
      <th>성적서(인증서)번호</th>
      <th>발급기관</th>
      <th>종류</th>
      <th>발급일자</th>
      <th>유효기간(From)</th>
      <th>유효기간(To)</th>
      <th>상태</th>
      <th>파일다운로드</th>
    </tr></thead>
    <tbody>
      ${rows.map(r=>`
        <tr>
          <td><button class="btn" data-view-target="${r.cert_id}">확인하기</button></td>
          <td><b>${r.cert_no}</b></td>
          <td>${r.issuer}</td>
          <td>${r.type}</td>
          <td>${r.issued || "-"}</td>
          <td>${r.valid_from || "-"}</td>
          <td>${r.valid_to || "-"}</td>
          <td>${badgeForStatus(r._status)}</td>
          <td>
            ${r.file ? `<button class="btn" data-dl="${r.cert_id}">⬇</button>` : "-"}
          </td>
        </tr>
      `).join("")}
    </tbody>
  `;

  el.querySelectorAll("[data-view-target]").forEach(b=>{
    b.addEventListener("click", ()=> openTargetView(b.dataset.viewTarget));
  });
  el.querySelectorAll("[data-dl]").forEach(b=>{
    b.addEventListener("click", ()=> downloadFile(b.dataset.dl));
  });
}

// ===== status badge
function badgeForStatus(st){
  if(st==="VALID") return `<span class="badge b-valid">유효</span>`;
  if(st==="DUE") return `<span class="badge b-due">임박</span>`;
  if(st==="EXPIRED") return `<span class="badge b-expired">만료</span>`;
  return "";
}

// ===== export/reset
function exportJson(){
  const blob = new Blob([JSON.stringify(state,null,2)], { type:"application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "cert_system_export.json";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href), 500);
}
function resetAll(){
  if(!confirm("모든 데이터를 삭제하고 초기화하나?")) return;
  localStorage.removeItem(LS_KEY);
  state = loadState();
  saveState(state);
  pickedModelIds = new Set();
  updatePickedCount();
  renderModels();
  renderTargetPicker();
  renderCertList();
}

// ===== bindings
function bindEvents(){
  bindModalClose();

  // open modals
  $("#btnOpenModelMaster").addEventListener("click", ()=>{
    openModal("modalModelMaster");
    renderModels();
  });

  $("#btnOpenRegMaster").addEventListener("click", ()=>{
    if(state.models.length===0){
      alert("먼저 모델 Master를 업로드/등록해야 한다.");
      return;
    }
    openModal("modalRegMaster");
    updatePickedCount();
  });

  // 모델 업로드/템플릿/수기
  $("#btnImportModels").addEventListener("click", async ()=>{
    try{
      const file = $("#modelFile").files?.[0];
      const res = await importModelsFromExcel(file);
      $("#modelImportResult").textContent =
        `총 ${res.total}행 처리: 신규 ${res.inserted}, 업데이트 ${res.updated}, 스킵 ${res.skipped}`;
      renderModels();
      renderTargetPicker();
    }catch(e){
      alert(e.message || "업로드 실패");
    }
  });
  $("#btnDownloadModelTemplate").addEventListener("click", downloadModelTemplate);
  $("#modelFilter").addEventListener("input", renderModels);
  $("#btnAddModelManual").addEventListener("click", addModelManual);

  // 등록: 파일명 표시
  $("#regFile").addEventListener("change", ()=>{
    const f = $("#regFile").files?.[0];
    $("#regFilePath").value = f ? f.name : "";
  });
  $("#regNoExpiry").addEventListener("change", ()=>{
    const on = $("#regNoExpiry").checked;
    $("#regValidFrom").disabled = on;
    $("#regValidTo").disabled = on;
  });

  // 적용대상 지정 팝업
  $("#btnOpenTargetPicker").addEventListener("click", ()=>{
    openModal("modalTargetPicker");
    renderTargetPicker();
  });
  $("#btnTpSearch").addEventListener("click", renderTargetPicker);
  $("#tpKeyword").addEventListener("input", renderTargetPicker);

  $("#btnTpSelectAll").addEventListener("click", ()=>{
    state.models.forEach(m=>pickedModelIds.add(m.model_id));
    renderTargetPicker();
    updatePickedCount();
  });
  $("#btnTpClear").addEventListener("click", ()=>{
    pickedModelIds = new Set();
    renderTargetPicker();
    updatePickedCount();
  });
  $("#btnTpApply").addEventListener("click", applyPickedTargets);

  // 저장
  $("#btnSaveCert").addEventListener("click", saveCert);

  // 검색
  $("#btnSearch").addEventListener("click", renderCertList);

  // export/reset
  $("#btnExport").addEventListener("click", exportJson);
  $("#btnReset").addEventListener("click", resetAll);
}

// ===== init
function init(){
  // seed: 예시 모델 몇 개(원하면 삭제 가능)
  if(state.models.length===0){
    state.models.push(
      { model_id: uid("m"), product_code:"158060002", model_name:"RC620-22KF", gas_type:"LNG" },
      { model_id: uid("m"), product_code:"158070002", model_name:"RC620-27KF", gas_type:"LNG" },
      { model_id: uid("m"), product_code:"158080002", model_name:"RC620-30KF", gas_type:"LNG" }
    );
    saveState(state);
  }

  bindEvents();
  renderModels();
  renderTargetPicker();
  renderCertList();
  updatePickedCount();
}
init();
