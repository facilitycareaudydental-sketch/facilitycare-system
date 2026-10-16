var Za=Object.defineProperty;var pt=(e,t)=>()=>(e&&(t=e(e=0)),t);var ct=(e,t)=>{for(var i in t)Za(e,i,{get:t[i],enumerable:!0})};var Rt={};ct(Rt,{API:()=>Ot,CLIENT_SIDE_MAX_ROWS:()=>Ae,IS_DEVELOPMENT:()=>Ne,apiFetch:()=>M,clearToken:()=>qe,getToken:()=>Ke,getUser:()=>Z,setToken:()=>ut,setUser:()=>Oe});function Ke(){return localStorage.getItem("fm_token")}function ut(e){localStorage.setItem("fm_token",e)}function qe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function Z(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Oe(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function M(e,t={}){let i=Ke(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let p=`cb=${Date.now()}`,h=e.includes("?")?"&":"?",g=`${Ot}${e}${h}${p}`,u=await fetch(g,{...t,headers:a}),r;try{let f=await u.text();try{r=JSON.parse(f)}catch{r={error:`Server Error (${u.status}): ${f.substring(0,80)}...`}}}catch{r={error:"Gagal membaca respon dari server"}}return u.status===401&&(qe(),window.location.hash="#/login"),{ok:u.ok,status:u.status,data:r}}catch(p){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${p.message})`}}}}var Ne,en,Ot,Ae,V=pt(()=>{Ne=!1,en="https://fm-operations-api.facilitycare-audydental.workers.dev",Ot=en,Ae=1e4});var ht={};ct(ht,{confirmDialog:()=>W,createModal:()=>J});function J({title:e,content:t,onConfirm:i,onCancel:a,confirmText:p="Simpan",cancelText:h="Batal",size:g="md",confirmClass:u="btn-primary"}){let r={sm:"400px",md:"560px",lg:"720px",xl:"900px"},f=document.createElement("div");f.className="modal-overlay",f.innerHTML=`
    <div class="modal" style="max-width:${r[g]||r.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${h}</button>
        ${i?`<button class="btn ${u} modal-confirm">${p}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&f.querySelector(".modal-body").appendChild(t);let D=()=>{f.classList.remove("show"),setTimeout(()=>f.remove(),250)};return f.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),D()}),f.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),D()}),i&&f.querySelector(".modal-confirm").addEventListener("click",()=>i(f,D)),f.addEventListener("click",F=>{F.target===f&&(a&&a(),D())}),document.body.appendChild(f),requestAnimationFrame(()=>f.classList.add("show")),{overlay:f,close:D}}function W(e,t,i="Konfirmasi"){return J({title:i,content:`<p>${e}</p>`,onConfirm:(a,p)=>{t(),p()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var le=pt(()=>{});var he={};ct(he,{downloadExcel:()=>Y,parseExcel:()=>Ue,renderExcelButtons:()=>an});function Ue(e){return new Promise((t,i)=>{let a=new FileReader;a.onload=p=>{try{let h=new Uint8Array(p.target.result),g=XLSX.read(h,{type:"array"}),u=g.SheetNames[0],r=g.Sheets[u];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${g.SheetNames.join(", ")}`),console.log(`Sheet Used: ${u}`);let f=XLSX.utils.decode_range(r["!ref"]||"A1:A1"),D=f.e.r-f.s.r+1,F=f.e.c-f.s.c+1;console.log(`Total Rows (including empty): ${D}`),console.log(`Total Columns: ${F}`);let S=[];for(let C=f.s.c;C<=f.e.c;++C){let w=r[XLSX.utils.encode_cell({c:C,r:f.s.r})];w&&w.v&&S.push(w.v)}console.log(`Headers Found: ${S.join(", ")}`),console.log("---------------------------");let E=XLSX.utils.sheet_to_json(r,{defval:""});Object.defineProperty(E,"__worksheet",{value:r,enumerable:!1}),Object.defineProperty(E,"__headers",{value:S,enumerable:!1}),t(E)}catch(h){i(h)}},a.onerror=p=>i(p),a.readAsArrayBuffer(e)})}function Y(e,t){try{let i=XLSX.utils.json_to_sheet(e),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function an(e){return`
    <div class="excel-actions-dropdown" style="position:relative; display:inline-block;">
      <button class="btn" id="btn-aksi-${e}" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('excel-menu-${e}').classList.toggle('show-aksi-menu')">
        \u22EE Aksi
      </button>
      <div id="excel-menu-${e}" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:180px; z-index:999; padding:8px 0;">
        <button class="dropdown-item" id="btn-export-${e}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
          \u{1F4E5} Export Excel
        </button>
        <button class="dropdown-item" id="btn-template-${e}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
          \u{1F4C4} Download Template
        </button>
        <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${e}">
          \u{1F4E4} Import Excel
          <input type="file" id="input-import-${e}" accept=".xlsx, .xls, .csv" style="display:none;">
        </label>
      </div>
    </div>
    <style>
      .show-aksi-menu { display: flex !important; }
      .dropdown-item:hover { background-color: #F8FAFC !important; color: #2563EB !important; }
    </style>
  `}var ie=pt(()=>{});V();var gt={},We=null;function te(e,t){gt[e]=t}function $e(e){window.location.hash=e}function Ht(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...a]=t.split("?"),p=gt[i];if(!p){for(let[g,u]of Object.entries(gt))if(g.endsWith("/*")&&i.startsWith(g.slice(0,-2))){p=u;break}}We&&(We(),We=null);let h=document.getElementById("main-content");if(h&&(h.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),p){let g=new URLSearchParams(a.join("?")),u=i.split("/").filter(Boolean),r=await p({path:i,params:g,segments:u,main:h});r&&(We=r)}else{let g=h||document.getElementById("app");g&&(g.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var Re;function tn(){return Re||(Re=document.createElement("div"),Re.id="toast-container",document.body.appendChild(Re)),Re}function mt(e,t="info",i=3500){let a=tn(),p=document.createElement("div");p.className=`toast toast-${t}`;let h={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};p.innerHTML=`<span class="toast-icon">${h[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,a.appendChild(p),requestAnimationFrame(()=>p.classList.add("show")),setTimeout(()=>{p.classList.remove("show"),setTimeout(()=>p.remove(),350)},i)}var K=e=>mt(e,"success"),N=e=>mt(e,"error");var bt=e=>mt(e,"info");le();V();V();le();function Xe({columns:e,data:t,onEdit:i,onDelete:a,onView:p,actions:h=[],emptyText:g="Tidak ada data",bulkSelect:u=null}){let r=document.createElement("div");if(r.className="table-wrapper",!t||t.length===0)return r.innerHTML=`<div class="empty-state"><p>${g}</p></div>`,r;let f=document.createElement("table");f.className="data-table";let D=document.createElement("thead"),F=document.createElement("tr");if(u){let E=document.createElement("th");E.style.width="40px",E.style.textAlign="center";let C=document.createElement("input");C.type="checkbox",C.id="select-all-checkbox",C.title="Pilih semua",C.addEventListener("change",()=>{t.forEach(w=>{C.checked?u.selectedIds.add(w.id):u.selectedIds.delete(w.id)}),r.querySelectorAll(".row-checkbox").forEach(w=>w.checked=C.checked),u.onToggle()}),E.appendChild(C),F.appendChild(E)}if(e.forEach(E=>{let C=document.createElement("th");C.textContent=E.label,E.width&&(C.style.width=E.width),F.appendChild(C)}),i||a||p||h.length>0){let E=document.createElement("th");E.textContent="Aksi",E.style.width="120px",F.appendChild(E)}D.appendChild(F),f.appendChild(D);let S=document.createElement("tbody");return t.forEach(E=>{let C=document.createElement("tr");if(u){let w=document.createElement("td");w.style.textAlign="center",w.style.width="40px";let y=document.createElement("input");y.type="checkbox",y.className="row-checkbox",y.checked=u.selectedIds.has(E.id),y.addEventListener("change",()=>{if(y.checked)u.selectedIds.add(E.id);else{u.selectedIds.delete(E.id);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1)}u.onToggle()}),w.appendChild(y),C.appendChild(w)}if(e.forEach(w=>{let y=document.createElement("td");if(w.render){let b=w.render(E[w.key],E);b instanceof HTMLElement?y.appendChild(b):y.innerHTML=b||""}else y.textContent=E[w.key]!==null&&E[w.key]!==void 0&&E[w.key]!==""?E[w.key]:"";w.nowrap&&(y.style.whiteSpace="nowrap"),C.appendChild(y)}),i||a||p||h.length>0){let w=document.createElement("td");w.className="actions-cell";let y=document.createElement("div");if(y.className="btn-group",p){let b=document.createElement("button");b.className="btn btn-xs btn-ghost",b.innerHTML="\u{1F441}",b.title="Lihat",b.addEventListener("click",()=>p(E)),y.appendChild(b)}if(i){let b=document.createElement("button");b.className="btn btn-xs btn-secondary",b.innerHTML="\u270F\uFE0F",b.title="Edit",b.addEventListener("click",()=>i(E)),y.appendChild(b)}h.forEach(b=>{let n=document.createElement("button");n.className=`btn btn-xs ${b.class||"btn-ghost"}`,n.innerHTML=b.icon||b.label,n.title=b.label,n.addEventListener("click",()=>b.handler(E)),y.appendChild(n)}),w.appendChild(y),C.appendChild(w)}S.appendChild(C)}),f.appendChild(S),r.appendChild(f),r}function Ze({page:e,pages:t,total:i,limit:a,onPage:p}){if(t<=1)return null;let h=document.createElement("div");h.className="pagination";let g=document.createElement("span");g.className="pagination-info",g.textContent=`Total: ${i} data`,h.appendChild(g);let u=document.createElement("div");u.className="pagination-btns";let r=(F,S,E=!1,C=!1)=>{let w=document.createElement("button");w.className=`btn btn-sm ${C?"btn-primary":"btn-ghost"} pagination-btn`,w.textContent=F,w.disabled=E,w.addEventListener("click",()=>p(S)),u.appendChild(w)};r("\xAB",1,e===1),r("\u2039",e-1,e===1);let f=Math.max(1,e-2),D=Math.min(t,e+2);for(let F=f;F<=D;F++)r(F,F,!1,F===e);return r("\u203A",e+1,e===t),r("\xBB",t,e===t),h.appendChild(u),h}le();function He(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${He(t.fields)}</div>`;let i=t.required?"required":"",a=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",p="";switch(t.type){case"textarea":p=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let g=(t.options||[]).map(E=>{let C=typeof E=="object"?E.value:E,w=typeof E=="object"?E.label:E,y=t.value==C?"selected":"";return`<option value="${C}" ${y}>${w}</option>`}).join("");p=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${g}</select>`;break;case"combobox":let u=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,r=(t.options||[]).map(E=>{let C=typeof E=="object"?E.value:E,w=typeof E=="object"?E.label||E.value||"":E||"";return(w==="undefined"||w==="[object Object]"||w==="null")&&(w=""),w?`<option value="${w}"></option>`:""}).join(""),f=t.value||"";if(t.value){let E=(t.options||[]).find(C=>(typeof C=="object"?C.value:C)==t.value);if(E){let C=typeof E=="object"?E.label||E.value||"":E||"";C&&C!=="undefined"&&C!=="[object Object]"&&C!=="null"&&(f=C)}}p=`
          <input type="text" name="${t.name}" list="${u}" class="form-control" value="${f}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${u}">${r}</datalist>
        `;break;case"checkbox":p=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let D=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";p=`<input type="date" name="${t.name}" class="form-control" value="${D}" ${i}>`;break;case"month":let F=t.value?String(t.value).slice(0,7):"";p=`<input type="month" name="${t.name}" class="form-control" value="${F}" ${i}>`;break;case"number":let S=t.readonly?"readonly":"";p=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${S}>`;break;case"email":p=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":p=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:p=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let h=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${a}${p}${h}</div>`}).join("")}function et(e){let t={},i=new FormData(e);for(let[a,p]of i.entries())t[a]=p===""?null:p;return e.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(t[a.name]=null)}),t}function tt(e,t){t&&Object.entries(t).forEach(([i,a])=>{let p=e.querySelector(`[name="${i}"]`);p&&(p.hasAttribute("list")||(p.type==="checkbox"?p.checked=!!a:p.type==="date"&&a&&window.parseFlexibleDate?p.value=window.parseFlexibleDate(a):p.type==="month"&&a?p.value=String(a).slice(0,7):p.value=a??""))})}ie();var Ce={},at={on(e,t){Ce[e]||(Ce[e]=new Set),Ce[e].add(t)},off(e,t){Ce[e]&&Ce[e].delete(t)},emit(e,t){Ce[e]&&Ce[e].forEach(i=>{try{i(t)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(Ce).forEach(e=>delete Ce[e])}},nn=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function be(e){if(!e){at.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");at.emit("data:changed",{module:t,relevant:nn.has(t)})}function se({container:e,title:t,icon:i,apiPath:a,columns:p,formFields:h,filterFields:g,defaultFilters:u={},itemLabel:r="Data",canCreate:f=!0,canEdit:D=!0,canDelete:F=!0,onBeforeSubmit:S,onAfterLoad:E,onDataLoaded:C,extraActions:w=[],initialSearch:y="",exportOptions:b=null,bulkDelete:n=!1,paginationMode:o="server"}){let s=Z();s&&s.role==="viewer"&&(f=!1,D=!1,F=!1,n=!1,b=null),s&&s.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(f=!1,D=!1,F=!1,n=!1,b=null),s&&s.role==="input_lembur"&&(a==="/api/overtime"?(F=!1,n=!1,b=null):(f=!1,D=!1,F=!1,n=!1,b=null));let d=1,m={...u};y&&(m.search=y);let l=new Set;e.innerHTML=`
    ${n?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; justify-content:space-between; background:#2563EB; padding:12px 16px; border-radius:8px; margin-bottom:16px; box-shadow:0 4px 6px -1px rgba(37,99,235,0.2);">
      <button id="btn-bulk-cancel" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6L6 18M6 6l12 12"></path></svg>
      </button>
      <span id="bulk-count" style="font-weight:500; font-size:0.95rem; color:white;">0 item dipilih</span>
      <button id="btn-bulk-delete" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
      </button>
    </div>`:""}

    <div class="page-header">
      <h1 class="page-title">${i} ${t}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${f?`<button class="btn btn-primary" id="btn-create">+ Tambah ${r}</button>`:""}
        ${b?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${b.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${b.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${b.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${b.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
              </label>

            </div>
          </div>
          <style>
            .show-aksi-menu { display: flex !important; }
            .dropdown-item:hover { background-color: #F8FAFC !important; color: #2563EB !important; }
          </style>
        `:""}
      </div>
    </div>
    

    ${g&&g.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${g.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${m.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${g.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let B=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${B}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${m[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${g.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let B=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${B}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${m[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function c(){let x=document.getElementById("bulk-toolbar");if(!x)return;let B=document.getElementById("bulk-count"),L=document.getElementById("btn-bulk-delete"),H=document.getElementById("btn-bulk-cancel");B.textContent=`${l.size} item dipilih`,l.size>0?(x.style.display="flex",L.disabled=!1,H.disabled=!1):(x.style.display="none",L.disabled=!0,H.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let j=document.querySelectorAll(".row-checkbox");if(j.length>0){let q=[...j].every(Q=>Q.checked),R=[...j].some(Q=>Q.checked);O.checked=q,O.indeterminate=R&&!q}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{l.clear(),document.querySelectorAll(".row-checkbox").forEach(B=>B.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),c()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(l.size===0)return;let x=[...l],B=document.createElement("div");B.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",B.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${r}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(B),B.querySelector("#bulk-cancel-btn").addEventListener("click",()=>B.remove()),B.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let L=B.querySelector("#bulk-confirm-btn");L.disabled=!0,L.textContent="Menghapus...";let H=await M(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});B.remove(),H.ok?(K(`${x.length} ${r} berhasil dihapus.`),l.clear(),c(),be(a),_()):N(H.data?.error||"Gagal menghapus data.")})});let v=document.getElementById("filter-search"),k;v?.addEventListener("input",x=>{clearTimeout(k),k=setTimeout(()=>{m.search=x.target.value,d=1,l.clear(),c(),_()},400)}),g?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",B=>{m[x.name]=B.target.value;let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value=B.target.value),d=1,l.clear(),c(),_()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",B=>{m[x.name]=B.target.value;let L=document.getElementById(`filter-${x.name}`);L&&(L.value=B.target.value),d=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{m={},v&&(v.value=""),g?.forEach(x=>{let B=document.getElementById(`filter-${x.name}`);B&&(B.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),d=1,l.clear(),c(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{m={},v&&(v.value=""),g?.forEach(x=>{let B=document.getElementById(`filter-${x.name}`);B&&(B.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),d=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>z(null)),b&&document.addEventListener("click",function(x){let B=document.getElementById("aksi-menu-main"),L=document.getElementById("btn-aksi-main");B&&L&&!L.contains(x.target)&&!B.contains(x.target)&&B.classList.remove("show-aksi-menu")});let T=document.getElementById("btn-mobile-filter"),A=document.getElementById("filter-options-wrapper"),P=document.getElementById("btn-close-filter-sheet");if(T&&A&&(T.addEventListener("click",x=>{x.preventDefault(),A.classList.add("sheet-open")}),P&&P.addEventListener("click",x=>{x.preventDefault(),A.classList.remove("sheet-open")})),b){document.getElementById(`btn-export-${b.moduleName}`)?.addEventListener("click",async B=>{let L=B.target,H=L.innerHTML;L.innerHTML="\u23F3 Loading...",L.disabled=!0;try{await b.onExport()}catch{N("Gagal export data")}finally{L.innerHTML=H,L.disabled=!1}}),document.getElementById(`btn-template-${b.moduleName}`)?.addEventListener("click",()=>{b.onTemplate()});let x=document.getElementById(`input-import-${b.moduleName}`);x?.addEventListener("change",async B=>{let L=B.target.files[0];if(!L)return;let H=document.getElementById(`label-import-${b.moduleName}`),O=H?H.querySelector(".import-text"):null,j=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),H&&(H.style.pointerEvents="none"),x.disabled=!0;try{let q=await Ue(L);if(q.length===0)throw new Error("File kosong atau format salah");await b.onImport(q),K("Import berhasil!"),be(a),_()}catch(q){N(q.message||"Gagal import data")}finally{O&&(O.innerText=j),H&&(H.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){c();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let B=o==="client",L=B?1:d,H=B?Ae:20,O=new URLSearchParams({page:L,limit:H,...Object.fromEntries(Object.entries(m).filter(([,G])=>G))}),j=await M(`${a}?${O}`);if(!j.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${j.data?.error||"Error"}</p></div>`;return}let q=j.data?.data||j.data||[],R=j.data?.pagination,Q=q.length;if(B){q=C(q);let G=q.length,ae=20,X=Math.ceil(G/ae);d>X&&X>0&&(d=X);let U=(d-1)*ae,ue=d*ae;q=q.slice(U,ue),R={page:d,limit:ae,total:G,pages:X}}!1,E&&E(q);let ne=Xe({columns:p,data:q,onEdit:D?G=>z(G):null,actions:w.map(G=>({...G,handler:ae=>G.handler(ae,_)})),emptyText:`Tidak ada ${String(r||"").toLowerCase()}`,bulkSelect:n?{selectedIds:l,onToggle:c}:null});x.innerHTML="",x.appendChild(ne);let oe=document.getElementById("pagination-container");if(oe&&(oe.innerHTML="",R&&R.pages>1)){let G=Ze({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:ae=>{d=ae,_()}});G&&oe.appendChild(G)}}function I(x){let B=typeof h=="function"?h(x):h;return He(B)}function z(x){let B=!!x,L=document.createElement("form");if(L.noValidate=!0,L.innerHTML=I(x),B){let O=typeof h=="function"?h(x):h;tt(L,x)}let{close:H}=J({title:B?`Edit ${r}`:`Tambah ${r}`,content:L,size:"lg",confirmText:B?"Simpan Perubahan":`Tambah ${r}`,onConfirm:async(O,j)=>{if(!L.reportValidity())return;let q=O.querySelector(".modal-confirm");q.disabled=!0,q.textContent="Menyimpan...";let R=et(L),Q=typeof h=="function"?h(x):h,ne=async X=>{for(let U of X)if(U.type==="row")await ne(U.fields);else if(U.type==="combobox"&&R[U.name]){let ue=R[U.name],xe=(U.options||[]).find(de=>{let me=String(typeof de=="object"?de.value:de),dt=String(typeof de=="object"?de.label:de);return me===ue||dt===ue});if(xe)R[U.name]=typeof xe=="object"?xe.value:xe;else if(U.createApi){let de={};de[U.createApi.field]=ue,U.createApi.extra&&Object.assign(de,U.createApi.extra);let me=await M(U.createApi.path,{method:"POST",body:JSON.stringify(de)});if(me.ok&&me.data?.id)R[U.name]=me.data.id;else if(me.ok&&!me.data?.id)R[U.name]=ue;else throw new Error(`Gagal membuat master data: ${me.data?.error||"Unknown error"}`)}}};try{await ne(Q)}catch(X){N(X.message),q.disabled=!1,q.textContent=B?"Simpan Perubahan":`Tambah ${r}`;return}S&&(R=await S(R,x));let oe=B?"PUT":"POST",G=B?`${a}/${x.id}`:a,ae=await M(G,{method:oe,body:JSON.stringify(R)});ae.ok?(K(B?`${r} berhasil diperbarui.`:`${r} berhasil ditambahkan.`),j(),be(a),_()):(N(ae.data?.error||"Gagal menyimpan data."),q.disabled=!1,q.textContent=B?"Simpan Perubahan":`Tambah ${r}`)}})}function $(x){W(`Hapus ${r} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let B=await M(`${a}/${x.id}`,{method:"DELETE"});B.ok?(K(`${r} berhasil dihapus.`),be(a),_()):N(B.data?.error||"Gagal menghapus.")},`Hapus ${r}`)}return _(),_}V();V();var De=null,nt=null;async function _e(e=!1){if(De&&!e)return console.log("Employees Raw (Cache Hit)",De.slice(0,5)),De;let t=await M(`/api/employees?limit=${Ae}&status=Aktif`);return De=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",De.slice(0,5)),De}async function ge(e=!1){let i=(await _e(e)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function ee(e=!1){return nt&&!e||(nt=((await M("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),nt}function pe(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function ft(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function xt(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function yt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function ye(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}ie();V();le();V();ie();V();le();V();var Fe={};function Le(e){if(Fe[e]){try{Fe[e].destroy()}catch{}delete Fe[e]}}function rn(){Object.keys(Fe).forEach(Le)}var fe=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},Te=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function it(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let a=Math.max(0,Math.round(fe(t))),p=e.textContent?e.textContent.replace(/[^\d]/g,""):"",h=p!==""&&parseInt(p)||0;if(a===h){e.textContent=a.toLocaleString("id-ID"),e._animFrame=null;return}let g=Date.now(),u=()=>{let r=Math.min((Date.now()-g)/i,1),f=1-Math.pow(1-r,3);e.textContent=Math.round(h+f*(a-h)).toLocaleString("id-ID"),r<1?e._animFrame=requestAnimationFrame(u):(e.textContent=a.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(u)}var on={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},sn=e=>{let t=Te(e,"\u2014");return`<span class="status-pill ${on[t]||"pill-neutral"}">${t}</span>`};var ke={family:"Inter",size:11},Ee="#94A3B8",Pe="#F1F5F9",vt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ln=()=>window.innerWidth<768;function rt(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ln()?"bottom":"top",labels:{font:ke,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ke,titleFont:{...ke,weight:"700"}}},scales:{x:{grid:{color:Pe},ticks:{font:ke,color:Ee,maxRotation:0}},y:{grid:{color:Pe},ticks:{font:ke,color:Ee},beginAtZero:!0}},...e}}var dn=()=>Array(6).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none;padding:16px">
    <div style="display:flex; gap:16px; align-items:flex-start">
      <div class="skeleton" style="width:48px;height:48px;border-radius:12px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:40px;height:24px;margin-bottom:6px"></div>
        <div class="skeleton skeleton-text" style="width:80px;height:12px;margin-bottom:4px"></div>
        <div class="skeleton skeleton-text" style="width:100px;height:10px"></div>
      </div>
    </div>
    <div style="display:flex; align-items:flex-end; gap:8px; margin-top:16px">
      <div class="skeleton" style="flex:1;height:24px;border-radius:4px"></div>
      <div class="skeleton skeleton-text" style="width:30px;height:12px"></div>
    </div>
  </div>`).join(""),pn=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ut(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ve(e,t,i=8e3){try{let a=new AbortController,p=setTimeout(()=>a.abort(),i),h=await M(e,{signal:a.signal}).catch(()=>null);if(clearTimeout(p),!h||!h.ok)return t;let g=h.data;return g?g.data!==void 0?g.data??t:g:t}catch{return t}}function cn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let p=document.getElementById(a);p&&(p.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let p=document.getElementById(a);if(p&&p.style.display==="none"){p.style.display="block";let h=p.parentElement;if(h&&!h.querySelector(".chart-empty")){let g=document.createElement("div");g.className="chart-empty",g.textContent="Belum ada data",p.style.display="none",h.appendChild(g)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Yt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Vt({}),["table-contracts","table-issues"].forEach(a=>{let p=document.getElementById(a);p&&p.querySelector(".skeleton")&&(p.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Jt(e){rn(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${dn()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${pn()}</div>

      <!-- Charts Row -->
      <div class="charts-row" style="grid-template-columns: 5fr 3fr 5fr;">
        <!-- Jadwal Kegiatan Chart -->
        <div class="chart-card">
          <div class="chart-card-header" style="align-items:flex-start">
            <div>
              <div class="chart-card-title">Jadwal Kegiatan</div>
            </div>
            <select id="filter-jadwal-year" class="btn-ghost" style="padding:4px;font-size:0.7rem;border:1px solid var(--border);border-radius:4px;cursor:pointer;color:var(--primary)">
              <option value="2026">2026</option>
              <option value="2027">2027</option>
            </select>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative;margin-top:10px">
            <div id="skel-jadwal" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-jadwal" style="display:none"></canvas>
          </div>
          <div id="jadwal-legend" style="display:flex;justify-content:center;gap:6px;margin-top:10px;font-size:0.55rem;font-weight:600;color:var(--text-2);flex-wrap:nowrap;white-space:nowrap;">
            <div style="display:flex;align-items:center;gap:3px"><div style="width:10px;height:8px;border-radius:2px;background:#3B82F6"></div> Inspeksi</div>
            <div style="display:flex;align-items:center;gap:3px"><div style="width:10px;height:8px;border-radius:2px;background:#10B981"></div> General Cleaning</div>
            <div style="display:flex;align-items:center;gap:3px"><div style="width:10px;height:8px;border-radius:2px;background:#F59E0B"></div> Deep Cleaning</div>
            <div style="display:flex;align-items:center;gap:3px"><div style="width:10px;height:8px;border-radius:2px;background:#EF4444"></div> Fogging</div>
          </div>
        </div>
        <div class="chart-card" style="display:flex; flex-direction:column;">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="flex:1; display:flex; flex-direction:column; justify-content:center;">
            <div style="display:flex; align-items:center; justify-content:center; gap:0px;">
              <div class="chart-canvas-wrap" style="width:110px;height:110px;position:relative">
                <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
                <canvas id="chart-donut" style="display:none"></canvas>
              </div>
              <div id="donut-legend" class="donut-legend" style="width:65px; margin-left:8px"></div>
            </div>
          </div>
          <div style="text-align:center; font-size:0.75rem; color:var(--text-3); margin-top:16px">
            Periode: 22 Juni - 22 Juli 2026
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header" style="align-items:flex-start">
            <div class="chart-card-title" style="font-size:0.85rem">Trend Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:10px;font-size:0.6rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:6px"><div style="width:10px;height:10px;border-radius:50%;background:#EF4444"></div> Open</div>
               <div style="display:flex;align-items:center;gap:6px"><div style="width:10px;height:10px;border-radius:50%;background:#10B981"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative;margin-top:10px">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Charts Row 2 -->
      <div class="charts-row" style="grid-template-columns: 1fr; margin-top:16px;">
        <div class="chart-card">
          <div class="chart-card-header" style="align-items:flex-start">
            <div>
              <a href="#/reports/inspection" class="chart-card-title" style="text-decoration:none; display:inline-block">Rata-rata Skor Inspeksi per Cabang <span style="font-size:0.8rem; color:var(--primary); font-weight:600; margin-left:8px">Lihat Laporan &rarr;</span></a>
              <div class="chart-card-subtitle" style="font-size:0.65rem">Skor rata-rata SCM & Cleaning</div>
            </div>
            <select id="filter-insp-month" class="btn-ghost" style="padding:4px;font-size:0.7rem;border:1px solid var(--border);border-radius:4px;cursor:pointer">
              <option value="">Pilih Bulan</option>
              <option value="01">Januari</option>
              <option value="02">Februari</option>
              <option value="03">Maret</option>
              <option value="04">April</option>
              <option value="05">Mei</option>
              <option value="06">Juni</option>
              <option value="07">Juli</option>
              <option value="08">Agustus</option>
              <option value="09">September</option>
              <option value="10">Oktober</option>
              <option value="11">November</option>
              <option value="12">Desember</option>
            </select>
          </div>
          <div class="chart-canvas-wrap" style="height:350px;position:relative;margin-top:10px">
            <div id="skel-insp" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-insp" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Bottom Row -->
      <div class="bottom-row" style="margin-top:24px;">
        <!-- Jadwal Hari Ini -->
        <div class="chart-card">
          <div class="chart-card-header" style="flex-wrap: wrap; gap: 8px;">
            <div class="chart-card-title">Jadwal Hari Ini <span style="font-size:0.75rem; font-weight:normal; color:var(--text-3); margin-left:6px">${t}</span></div>
            <a href="#/calendar" class="chart-link">Lihat Kalender</a>
          </div>
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Ut(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Ut(3)}</div>
        </div>
        <!-- Kontrak Akan Habis -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Kontrak Akan Habis</div>
            <a href="#/contracts" class="chart-link">Lihat Data</a>
          </div>
          <div class="chart-canvas-wrap" style="height:160px;position:relative;margin-top:10px">
            <div id="skel-contract-mini" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-contract-mini" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Quick Actions Row -->
      <div class="actions-wrap">
        <div class="actions-title">Aksi Cepat</div>
        <div class="actions-row" id="quick-actions">
          <!-- Rendered in JS -->
        </div>
      </div>



    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Gt(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,p=document.getElementById("jadwal-year-label");p&&(p.textContent=a);let h=document.getElementById("skel-jadwal"),g=document.getElementById("chart-jadwal");h&&(h.style.display="block",h.style.position="absolute"),g&&(g.style.display="none");let u=await ve(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Qt(u)}catch(r){console.warn("ScheduleChart render:",r),we("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,p=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",h=document.getElementById("skel-insp"),g=document.getElementById("chart-insp");h&&(h.style.display="block",h.style.position="absolute"),g&&(g.style.display="none");let u=await ve(p,{},8e3);try{Wt(u)}catch(r){console.warn("InspBar render:",r),we("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>cn(),5e3),await Gt(e)}async function Gt(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,a,p,h,g,u,r]=await Promise.all([ve("/api/dashboard/kpi",{},8e3),ve("/api/dashboard/issues-trend",{},8e3),ve("/api/dashboard/issues-summary",{},8e3),ve("/api/dashboard/stats",{},8e3),ve("/api/dashboard/calendar",[],8e3),ve("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ve(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ve("/api/dashboard/dropdown-aggregates",{},8e3)]),f=document.getElementById("filter-insp-month"),D=f?f.value:"",F=D?`/api/dashboard/inspection-bar?month=${D}`:"/api/dashboard/inspection-bar",S=await ve(F,{},8e3);window.dashboardAggregates=r||{};try{Yt(t)}catch(E){console.warn("KPI render:",E)}try{Vt(t)}catch(E){console.warn("MiniStats render:",E)}try{Qt(u)}catch(E){console.warn("ScheduleChart render:",E),we("skel-jadwal","chart-jadwal")}try{un(Array.isArray(a?.by_category)?a.by_category:[])}catch(E){console.warn("Donut render:",E),we("skel-donut","chart-donut")}try{gn(i)}catch(E){console.warn("Trend render:",E),we("skel-trend","chart-trend")}try{Wt(S)}catch(E){console.warn("InspBar render:",E),we("skel-insp","chart-insp")}try{let E=Array.isArray(p)?p:Array.isArray(p?.recent_issues)?p.recent_issues:[];bn(E)}catch(E){console.warn("IssuesTable render:",E)}try{let E=Array.isArray(p?.expiring_contracts)?p.expiring_contracts:[];mn(g)}catch(E){console.warn("ContractsTable render:",E)}try{hn(Array.isArray(h)?h:[])}catch(E){console.warn("Agenda render:",E)}try{fn()}catch(E){console.warn("Quick Actions render:",E)}}function Yt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=window.dashboardAggregates||{},p=r=>{let f=window.dashboardAggregates?.schedule_quarter_details?.[r];return f?`Insp: ${f.insp} \u2022 GC: ${f.gc} \u2022 DC: ${f.dc} \u2022 Fog: ${f.fog}`:""},h=r=>{let f=window.dashboardAggregates?.schedule_quarter_details?.[r];return f?`Rincian ${r}: Inspeksi: ${f.insp} | General Cleaning: ${f.gc} | Deep Cleaning: ${f.dc} | Fogging: ${f.fog}`:`Total Kegiatan ${r}`},g=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:a.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:p(i),subId:"mini-jadwal-detail",tooltip:h(i)}];t.innerHTML=g.map(r=>{let f=r.val!==void 0?fe(r.val,0):fe(e[r.key]?.current,0);return r.dropdown?`
        <div class="kpi-card ${r.color}" ${r.id?`id="${r.id}"`:""} ${r.tooltip?`title="${r.tooltip}"`:""} data-href="${r.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${r.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${f}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${f}</div>
                ${r.dropdown}
              </div>
              <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${r.label}</div>
              <div class="kpi-subtitle" ${r.subId?`id="${r.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${r.tooltip||""}">${r.sub}</div>
            </div>
          </div>
        </div>`:`
      <a href="${r.href}" class="kpi-card ${r.color}" ${r.id?`id="${r.id}"`:""} ${r.tooltip?`title="${r.tooltip}"`:""} style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${r.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="kpi-value" data-target="${f}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${f}</div>
              ${r.dropdown?r.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${r.label}</div>
            <div class="kpi-subtitle" ${r.subId?`id="${r.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${r.tooltip||""}">${r.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(r=>{r.addEventListener("click",f=>{f.target.closest("select")||r.dataset.href&&(window.location.hash=r.dataset.href)}),r.addEventListener("keydown",f=>{f.key==="Enter"&&!f.target.closest("select")&&r.dataset.href&&(window.location.hash=r.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(r=>{it(r,parseInt(r.dataset.target)||0)});let u=document.getElementById("dash-jadwal-period");if(u){let r=f=>{let F=(window.dashboardAggregates?.schedule_by_quarter||{})[f]||0,S=document.querySelector("#mini-jadwal .kpi-value");S&&(S.dataset.target=F,it(S,F,400));let E=document.getElementById("mini-jadwal-detail");E&&(E.textContent=p(f),E.title=h(f));let C=document.getElementById("mini-jadwal");if(C){let w=`#/timeline?dash_filter=period_${f.toLowerCase()}`;C.tagName==="A"?C.href=w:C.dataset.href=w,C.title=h(f)}};u.addEventListener("change",f=>r(f.target.value))}}function Vt(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),a=String(new Date().getMonth()+1).padStart(2,"0"),p=`${i}-${a}`,h=f=>`
    <select id="${f}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${i}-01" ${p===`${i}-01`?"selected":""}>Jan</option>
      <option value="${i}-02" ${p===`${i}-02`?"selected":""}>Feb</option>
      <option value="${i}-03" ${p===`${i}-03`?"selected":""}>Mar</option>
      <option value="${i}-04" ${p===`${i}-04`?"selected":""}>Apr</option>
      <option value="${i}-05" ${p===`${i}-05`?"selected":""}>Mei</option>
      <option value="${i}-06" ${p===`${i}-06`?"selected":""}>Jun</option>
      <option value="${i}-07" ${p===`${i}-07`?"selected":""}>Jul</option>
      <option value="${i}-08" ${p===`${i}-08`?"selected":""}>Agu</option>
      <option value="${i}-09" ${p===`${i}-09`?"selected":""}>Sep</option>
      <option value="${i}-10" ${p===`${i}-10`?"selected":""}>Okt</option>
      <option value="${i}-11" ${p===`${i}-11`?"selected":""}>Nov</option>
      <option value="${i}-12" ${p===`${i}-12`?"selected":""}>Des</option>
    </select>
  `,g=window.dashboardAggregates||{},u=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:h("dash-inspeksi-month"),val:g.inspeksi_by_month?.[p]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${p}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:h("dash-gcdc-month"),val:g.gcdc_by_month?.[p]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${p}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:h("dash-reliefer-month"),val:g.relievers_by_month?.[p]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${p}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:h("dash-fogging-month"),val:g.fogging_by_month?.[p]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${p}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=u.map(f=>f.dropdown?`
        <div class="mini-stat ${f.color}" style="cursor:pointer;" id="${f.id||""}" data-href="${f.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${f.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${fe(f.val)}">0</div>
              ${f.dropdown}
            </div>
            <div class="mini-stat-text">${f.label}</div>
            ${f.extra||""}
          </div>
        </div>`:`
      <a href="${f.href}" class="mini-stat ${f.color}" style="text-decoration:none" id="${f.id||""}">
        <div class="mini-stat-icon">${f.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${fe(f.val)}">0</div>
            ${f.dropdown?f.dropdown:""}
          </div>
          <div class="mini-stat-text">${f.label}</div>
          ${f.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(f=>{f.addEventListener("click",D=>{D.target.closest("select")||f.dataset.href&&(window.location.hash=f.dataset.href)}),f.addEventListener("keydown",D=>{D.key==="Enter"&&!D.target.closest("select")&&f.dataset.href&&(window.location.hash=f.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(f=>it(f,parseInt(f.dataset.target)||0,700));let r=(f,D,F,S)=>{let E=document.getElementById(f);if(E){let C=w=>{let b=(window.dashboardAggregates?.[F]||{})[w]||0,n=document.querySelector(`#${D} .mini-stat-value`);n&&(n.dataset.target=b,it(n,b,400));let o=document.getElementById(D);if(o){let s=`${S}&month=${w}`;o.tagName==="A"?o.href=s:o.dataset.href=s}};E.addEventListener("change",w=>C(w.target.value))}};r("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),r("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),r("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),r("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function un(e){we("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Le("donut");let a=(e||[]).filter(r=>fe(r.count)>0);if(!a.length){Ge(t,"Belum ada data permasalahan");return}let p=a.map(r=>`${Te(r.category,"Lainnya")}`),h=a.map(r=>fe(r.count)),g=h.reduce((r,f)=>r+f,0);i.innerHTML=a.map((r,f)=>{let D=vt[f%vt.length],F=g>0?Math.round(r.count/g*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${D}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${r.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${F}%)</span></div>
          <div class="donut-legend-label">${p[f]}</div>
        </div>
      </div>
    `}).join("");let u={id:"centerText",beforeDraw:function(r){let f=r.width,D=r.height,F=r.ctx;F.restore();let S=(D/80).toFixed(2);F.font="bold "+S+"em Inter",F.textBaseline="middle",F.fillStyle="#1E293B";let E=g.toString(),C=Math.round((f-F.measureText(E).width)/2),w=D/2;F.fillText(E,C,w-4),F.font="600 "+(S*.35).toFixed(2)+"em Inter",F.fillStyle="#64748B";let y="Total",b=Math.round((f-F.measureText(y).width)/2);F.fillText(y,b,w+10),F.save()}};Fe.donut=new Chart(t,{type:"doughnut",data:{labels:p,datasets:[{data:h,backgroundColor:vt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ke,titleFont:{...ke,weight:"700"},callbacks:{label:r=>` ${r.label}: ${r.parsed} kasus`}}},cutout:"75%"},plugins:[u]})}function gn(e){we("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Le("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(e.labels||[]).map(g=>{if(!g||typeof g!="string")return"";try{let[u,r]=g.split("-");return(i[Number(r)-1]||r)+" "+String(u).slice(-2)}catch{return g}}),p=(e.open||[]).map(g=>fe(g)),h=(e.closed||[]).map(g=>fe(g));if(!a.length){Ge(t,"Belum ada data trend");return}Fe.trend=new Chart(t,{type:"line",data:{labels:a,datasets:[{label:"Open",data:p,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:h,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:rt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:45,autoSkip:!0}},y:{grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:Ee},beginAtZero:!0}}})})}function Qt(e){we("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Le("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(r=>Array.isArray(r)&&r.some(f=>f>0))){Ge(t,"Belum ada data jadwal");return}let p=e["Inspeksi Hygiene"]||Array(12).fill(0),h=e["General Cleaning"]||Array(12).fill(0),g=e["Deep Cleaning"]||Array(12).fill(0),u=e.Fogging||Array(12).fill(0);Fe.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:p,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:h,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:g,backgroundColor:"#F59E0B"},{label:"Fogging",data:u,backgroundColor:"#EF4444"}]},options:rt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:Ee},min:0}}})})}function Wt(e){we("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Le("inspBar"),e=e||{};let i=e.labels||[],a=(e.fc||[]).map(h=>fe(h)),p=(e.spv||[]).map(h=>fe(h));if(!i.length){Ge(t,"Belum ada data inspeksi");return}Fe.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:p,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:rt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Ee,maxRotation:45,minRotation:45}},y:{grid:{color:Pe},ticks:{font:ke,color:Ee},min:0,max:100}}})})}function mn(e){we("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Le("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(e.labels||[]).map(g=>{let u=g.split("-")[1];return i[u]||g}),p=(e.data||[]).map(g=>fe(g));if(!a.length){Ge(t,"Belum ada data");return}let h=t.getContext("2d");Fe.contractMiniBar=new Chart(t,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:p,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:rt({onClick:(g,u)=>{if(u&&u.length>0){let r=u[0].index,f=(e.labels||[])[r];f&&(window.location.hash="#/contracts?month_expiry="+f)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Ee,maxRotation:0,autoSkip:!1}},y:{grid:{color:Pe,borderDash:[4,4],drawBorder:!1},ticks:{font:ke,color:Ee,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function bn(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${sn(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Te(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Te(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function hn(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,h=(e||[]).filter(g=>(g.event_date||"").startsWith(a)).slice(0,10);if(!h.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${h.map(g=>{let u="#3B82F6",r="#EFF6FF",f="Agenda",D=(g.title||"").toLowerCase();return D.includes("inspeksi")?(u="#10B981",r="#ECFDF5",f="Inspeksi"):D.includes("cleaning")||D.includes("gcdc")?(u="#3B82F6",r="#EFF6FF",f="Cleaning"):D.includes("reliefer")?(u="#F59E0B",r="#FFFBEB",f="Reliefer"):D.includes("fogging")&&(u="#8B5CF6",r="#F5F3FF",f="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(g.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${u};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Te(g.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Te(g.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${r};color:${u}">${f}</div>
        </div>
      `}).join("")}
    </div>
  `}function fn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function we(e,t){let i=document.getElementById(e),a=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let p=a.parentElement;if(p){let h=p.querySelector(".chart-empty");h&&h.remove()}}}function Ge(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let p=document.createElement("div");p.className="chart-empty",p.textContent=t,i.appendChild(p)}}V();async function Xt(e){document.getElementById("app").innerHTML=`
    <div class="login-page">
      <div class="login-card">

        <div class="login-header">
          <div class="login-logo-wrap">
            <div class="login-logo-icon">
              <svg width="32" height="32" viewBox="0 0 48 48" fill="none">
                <rect width="48" height="48" rx="14" fill="url(#lg)"/>
                <path d="M12 20h6v16h-6V20zm10-8h6v24h-6V12zm10 6h6v18h-6V18z" fill="#fff" fill-opacity=".9"/>
                <defs>
                  <linearGradient id="lg" x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
                    <stop stop-color="#2563EB"/>
                    <stop offset="1" stop-color="#0EA5E9"/>
                  </linearGradient>
                </defs>
              </svg>
            </div>
            <div>
              <h1 class="login-title">FCMS</h1>
              <p class="login-subtitle">Facility Care Management System</p>
            </div>
          </div>
          <div class="login-divider"></div>
          <p class="login-desc">Masuk untuk mengelola operasional Facility Care</p>
        </div>

        <form class="login-form" id="login-form" novalidate>
          <div class="form-group">
            <label class="form-label">Username / Email</label>
            <div class="input-with-icon">
              <svg class="input-prefix-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
              </svg>
              <input type="text" name="username" class="form-control has-prefix-icon"
                placeholder="Masukkan username atau email"
                required autofocus autocomplete="username">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Password</label>
            <div class="input-with-icon">
              <svg class="input-prefix-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>
              </svg>
              <input type="password" name="password" class="form-control has-prefix-icon"
                placeholder="\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022" required autocomplete="current-password" id="login-password">
              <button type="button" class="input-icon-btn" id="toggle-password" aria-label="Toggle password">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" id="icon-eye">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
                </svg>
              </button>
            </div>
          </div>

          <div id="login-error" class="alert alert-danger" style="display:none"></div>

          <button type="submit" class="btn btn-primary btn-full btn-lg" id="login-btn" style="margin-top:4px">
            <span class="btn-text">
              <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:6px">
                <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/>
              </svg>
              Masuk ke FCMS
            </span>
            <span class="btn-spinner" style="display:none">\u23F3 Memproses...</span>
          </button>
        </form>
        <div class="login-version">FCMS v2.0 \xB7 Facility Care Indonesia</div>
      </div>

      <!-- Decorative background blobs -->
      <div class="login-blob login-blob-1"></div>
      <div class="login-blob login-blob-2"></div>
    </div>
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),p=document.getElementById("toggle-password"),h=document.getElementById("login-password");p?.addEventListener("click",()=>{let g=h.type==="text";h.type=g?"password":"text",p.style.color=g?"":"var(--primary)"}),t?.addEventListener("submit",async g=>{g.preventDefault(),i.style.display="none";let u=t.username.value.trim(),r=t.password.value;if(!u||!r){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let f=await M("/api/auth/login",{method:"POST",body:JSON.stringify({username:u,password:r})});f.ok&&f.data.success?(ut(f.data.data.token),Oe(f.data.data.user),K("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=f.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}V();le();var kt=[{id:1,full_name:"ACHMAD NURYAINI",branch_id:7,branch_name:"007. Bekasi",division:"FACILITY CARE",phone:"081298765431",join_date:"2023-10-09",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:2,full_name:"ADAM JHORDY",branch_id:32,branch_name:"032. Cawang",division:"FACILITY CARE",phone:"085712345678",join_date:"2024-03-18",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:3,full_name:"ADE MUSTIKA REZDAFITRA",branch_id:16,branch_name:"016. Cibubur",division:"FACILITY CARE",phone:"081345678901",join_date:"2022-06-21",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:4,full_name:"ADE SURAHMAN",branch_id:0,branch_name:"000. Management",division:"FACILITY CARE",phone:"081223344556",join_date:"2021-01-15",status:"Aktif",notes:"Trainer & Supervisor Operasional",target_pindah_os:"",target_selesai_os:""},{id:5,full_name:"ADITYA",branch_id:26,branch_name:"026. Tangerang City",division:"FACILITY CARE",phone:"087811223344",join_date:"2024-10-19",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:6,full_name:"ADITYA PRATAMA",branch_id:54,branch_name:"054. Bogor Gunung Batu",division:"FACILITY CARE",phone:"089699887766",join_date:"2025-09-04",status:"Aktif",notes:"Lulusan Training Batch 5 nilai tertinggi",target_pindah_os:"",target_selesai_os:""},{id:7,full_name:"AGUNG SEPTIADI",branch_id:0,branch_name:"000. Management",division:"FC - RELIEFER",phone:"082155667788",join_date:"2026-02-12",status:"Aktif",notes:"Personil cover cabang mobile",target_pindah_os:"",target_selesai_os:""}],Zt=["FACILITY CARE","SECURITY","FC - RELIEFER"],xn=["Aktif","Tidak Aktif","Resign","Cut"];function yn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function vn(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function ea(e,t){let i=[],a=[],p="all",h="all",g="all",u="",r=!1,f=t?t.get("dash_filter"):null;f==="active"?p="Aktif":f==="reliefer"&&(h="FC - RELIEFER");let F=Z()?.role!=="viewer";try{a=await ee()}catch(n){console.error("Error fetching branches:",n)}async function S(){r=!0,b();try{let n=await M("/api/employees?limit=10000");if(n.ok){let o=n.data?.data||n.data||[];i=Array.isArray(o)&&o.length>0?o:kt}else i=kt}catch(n){console.warn("Gagal load employees, menggunakan data lokal:",n),i=kt}finally{r=!1,b()}}async function E(){try{let{downloadExcel:n}=await Promise.resolve().then(()=>(ie(),he)),o=i.map((s,d)=>({No:d+1,"Nama Lengkap":s.full_name||"",Cabang:s.branch_name||"",Divisi:s.division||"","No. HP":s.phone||"","Tgl Masuk":s.join_date||"",Status:s.status||"","Target Pindah OS":s.target_pindah_os||"","Target Selesai OS":s.target_selesai_os||"",Catatan:s.notes||""}));n(o,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),K("Berhasil mengekspor data karyawan ke Excel")}catch(n){N("Gagal ekspor Excel: "+n.message)}}function C(n=null){let o=!!n,s=n?.division||"FACILITY CARE",d=n?.status||"Aktif",m=document.createElement("div");m.className="emp-modal-body",m.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Achmad Nuryaini" value="${n?.full_name?String(n.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="08xx-xxxx-xxxx" value="${n?.phone?String(n.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${a.map(l=>`
                <option value="${l.value}" ${String(n?.branch_id)===String(l.value)?"selected":""}>
                  ${l.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Zt.map(l=>`
                <option value="${l}" ${s===l?"selected":""}>${l}</option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Masuk Kerja
            </label>
            <input type="date" id="emp-modal-joindate" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${n?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${xn.map(l=>`
                <option value="${l}" ${d===l?"selected":""}>${l}</option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Pindah OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-pindah" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${n?.target_pindah_os?String(n.target_pindah_os).slice(0,7):""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Selesai OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-selesai" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${n?.target_selesai_os?String(n.target_selesai_os).slice(0,7):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${n?.notes||""}</textarea>
        </div>
      </div>
    `,J({title:o?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:m,size:"md",confirmText:o?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(l,c)=>{let v=(m.querySelector("#emp-modal-name").value||"").trim(),k=(m.querySelector("#emp-modal-phone").value||"").trim(),T=m.querySelector("#emp-modal-branch").value,A=m.querySelector("#emp-modal-division").value,P=m.querySelector("#emp-modal-joindate").value||null,_=m.querySelector("#emp-modal-status").value,I=m.querySelector("#emp-modal-target-pindah").value||null,z=m.querySelector("#emp-modal-target-selesai").value||null,$=(m.querySelector("#emp-modal-notes").value||"").trim();if(!v||!T||!A||!_){N("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let x={full_name:v,phone:k,branch_id:Number(T),division:A,join_date:P,status:_,target_pindah_os:I,target_selesai_os:z,notes:$};try{let B;o?B=await M(`/api/employees/${n.id}`,{method:"PUT",body:JSON.stringify(x)}):B=await M("/api/employees",{method:"POST",body:JSON.stringify(x)}),B.ok?(K(o?"Data karyawan berhasil diperbarui!":"Karyawan baru berhasil ditambahkan!"),c(),await S()):N(B.data?.error||"Gagal menyimpan data karyawan")}catch(B){N("Terjadi kesalahan: "+B.message)}}})}async function w(n){W({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${n.full_name}</strong> dari cabang <em>${n.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let o=await M(`/api/employees/${n.id}`,{method:"DELETE"});o.ok?(K("Status karyawan berhasil diubah menjadi Tidak Aktif"),await S()):N(o.data?.error||"Gagal menonaktifkan data")}catch(o){N("Terjadi kesalahan: "+o.message)}}})}function y(){return i.filter(n=>{if(u){let o=u.toLowerCase(),s=(n.full_name||"").toLowerCase().includes(o),d=(n.branch_name||"").toLowerCase().includes(o),m=(n.phone||"").toLowerCase().includes(o),l=(n.division||"").toLowerCase().includes(o),c=(n.notes||"").toLowerCase().includes(o);if(!s&&!d&&!m&&!l&&!c)return!1}return!(p!=="all"&&(p==="Aktif"&&n.status!=="Aktif"||p==="Nonaktif"&&n.status==="Aktif")||h!=="all"&&n.division!==h||g!=="all"&&String(n.branch_id)!==String(g))})}function b(){let n=y(),o=i.length,s=i.filter($=>$.status==="Aktif").length,d=i.filter($=>$.division==="FACILITY CARE"&&$.status==="Aktif").length,m=i.filter($=>($.division==="SECURITY"||$.division==="FC - RELIEFER")&&$.status==="Aktif").length,l=i.length,c=i.filter($=>$.status==="Aktif").length,v=i.filter($=>$.status!=="Aktif").length;e.innerHTML=`
      <style>
        .emp-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .emp-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .emp-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(14, 165, 233, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .emp-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .emp-stat-box {
          background: rgba(255, 255, 255, 0.05);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 18px;
          display: flex;
          align-items: center;
          gap: 14px;
          backdrop-filter: blur(8px);
          transition: transform 0.15s ease, background 0.15s ease;
        }
        .emp-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .emp-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .emp-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .emp-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .emp-controls-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .emp-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .emp-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .emp-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .emp-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .emp-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .emp-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 170px;
          cursor: pointer;
        }
        .emp-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .emp-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .emp-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .emp-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .emp-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .emp-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .emp-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .emp-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .emp-pill.active .emp-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* Cards Grid */
        .emp-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
          gap: 16px;
        }
        .emp-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 14px;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .emp-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }
        .emp-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
        }
        
        .emp-div-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .emp-div-fc {
          background: #EFF6FF;
          color: #2563EB;
          border: 1px solid #DBEAFE;
        }
        .emp-div-sec {
          background: #F1F5F9;
          color: #334155;
          border: 1px solid #E2E8F0;
        }
        .emp-div-rel {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }

        .emp-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .emp-status-aktif {
          background: #DCFCE7;
          color: #16A34A;
          border: 1px solid #BBF7D0;
        }
        .emp-status-nonaktif {
          background: #FEE2E2;
          color: #DC2626;
          border: 1px solid #FCA5A5;
        }

        /* Profile Box */
        .emp-profile-box {
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .emp-avatar {
          width: 52px;
          height: 52px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.15rem;
          flex-shrink: 0;
          box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }
        .emp-avatar.sec {
          background: linear-gradient(135deg, #334155 0%, #0F172A 100%);
          box-shadow: 0 4px 10px rgba(15, 23, 42, 0.2);
        }
        .emp-avatar.rel {
          background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%);
          box-shadow: 0 4px 10px rgba(124, 58, 237, 0.2);
        }
        .emp-name {
          font-size: 1.05rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
        }
        .emp-subtitle {
          font-size: 0.76rem;
          color: #64748B;
          font-weight: 600;
          margin-top: 3px;
        }

        /* Info Container */
        .emp-info-card {
          background: #F8FAFC;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          font-size: 0.84rem;
        }
        .emp-info-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 8px;
        }
        .emp-info-label {
          color: #64748B;
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .emp-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
        }
        .emp-phone-link {
          color: #2563EB;
          text-decoration: none;
          font-weight: 600;
          display: inline-flex;
          align-items: center;
          gap: 4px;
        }
        .emp-phone-link:hover {
          text-decoration: underline;
        }

        /* Notes box */
        .emp-notes-box {
          font-size: 0.8rem;
          color: #475569;
          background: #FFFBEB;
          border: 1px solid #FEF3C7;
          border-radius: 8px;
          padding: 8px 10px;
          line-height: 1.4;
        }

        /* Card footer */
        .emp-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 12px;
          margin-top: auto;
          gap: 8px;
        }
        .emp-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
        }
        .emp-icon-btn {
          width: 34px;
          height: 34px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .emp-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .emp-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .emp-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .emp-hero-card {
            padding: 18px 16px;
          }
          .emp-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .emp-stat-box {
            padding: 10px 12px;
          }
          .emp-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .emp-stat-val {
            font-size: 1.3rem;
          }
          .emp-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .emp-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .emp-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .emp-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .emp-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .emp-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .emp-select {
            width: 100%;
          }
          .emp-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="emp-page">

        <!-- Hero Header Card -->
        <div class="emp-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u{1F465}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Master Karyawan FCMS
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pusat data personil Facility Care, Security, dan Reliefer. Mengatur status keaktifan, kontak, penempatan cabang, dan riwayat kerja.
              </p>
            </div>

            <div class="emp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${F?`
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              `:""}
              
              <button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${r?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(37, 99, 235, 0.2); color:#93C5FD;">\u{1F465}</div>
              <div>
                <div class="emp-stat-val">${o}</div>
                <div class="emp-stat-lbl">Total Karyawan</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val">${s}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(14, 165, 233, 0.2); color:#7DD3FC;">\u{1F3E2}</div>
              <div>
                <div class="emp-stat-val">${d}</div>
                <div class="emp-stat-lbl">Facility Care</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(124, 58, 237, 0.2); color:#C4B5FD;">\u{1F6E1}\uFE0F</div>
              <div>
                <div class="emp-stat-val">${m}</div>
                <div class="emp-stat-lbl">Security & Reliefer</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="emp-controls-card">
          <div class="emp-search-row">
            <div class="emp-search-wrap">
              <svg class="emp-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="emp-search-input" class="emp-search-input" 
                placeholder="Cari nama karyawan, no HP, cabang..." value="${u?String(u).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="emp-branch-filter" class="emp-select">
              <option value="all">Semua Cabang (${a.length})</option>
              ${a.map($=>`
                <option value="${$.value}" ${String(g)===String($.value)?"selected":""}>
                  ${$.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${Zt.map($=>`
                <option value="${$}" ${h===$?"selected":""}>
                  ${$}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${p==="all"&&h==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count">${l}</span>
            </div>
            <div class="emp-pill ${p==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count">${c}</span>
            </div>
            <div class="emp-pill ${h==="FC - RELIEFER"?"active":""}" data-status="all" data-div="FC - RELIEFER">
              <span class="emp-pill-dot" style="background:#8B5CF6;"></span>
              <span>FC - Reliefer</span>
              <span class="emp-pill-count">${i.filter($=>$.division==="FC - RELIEFER").length}</span>
            </div>
            <div class="emp-pill ${p==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count">${v}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="emp-cards-grid">
          ${n.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:n.map($=>{let x=$.division||"FACILITY CARE",B="emp-div-fc",L="";x==="SECURITY"?(B="emp-div-sec",L="sec"):x==="FC - RELIEFER"&&(B="emp-div-rel",L="rel");let H=$.status==="Aktif",O=H?"emp-status-aktif":"emp-status-nonaktif",j=H?"\u{1F7E2} Aktif":`\u26AA ${$.status||"Nonaktif"}`;return`
              <div class="emp-card" data-id="${$.id}">
                <!-- Top Badges Row -->
                <div class="emp-card-top">
                  <span class="emp-div-badge ${B}">
                    ${x==="SECURITY"?"\u{1F6E1}\uFE0F":x==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${x}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="emp-status-badge ${O}">
                      ${j}
                    </span>
                  </div>
                </div>

                <!-- Profile Box -->
                <div class="emp-profile-box">
                  <div class="emp-avatar ${L}">
                    ${vn($.full_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="emp-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${$.full_name||"-"}
                    </div>
                    <div class="emp-subtitle">
                      Personil ${x}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="emp-info-card">
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4CD} Cabang:</span>
                    <span class="emp-info-val">${$.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4DE} No. HP:</span>
                    <span class="emp-info-val">
                      ${$.phone?`
                        <a href="tel:${$.phone}" class="emp-phone-link">
                          <span>${$.phone}</span>
                        </a>
                      `:'<span style="color:#94A3B8;">-</span>'}
                    </span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F5D3}\uFE0F Tgl Masuk:</span>
                    <span class="emp-info-val">${yn($.join_date)}</span>
                  </div>
                  ${$.target_pindah_os||$.target_selesai_os?`
                    <div class="emp-info-row" style="border-top:1px dashed #E2E8F0; padding-top:6px; margin-top:2px;">
                      <span class="emp-info-label">\u{1F504} Target OS:</span>
                      <span class="emp-info-val" style="font-size:0.78rem; color:#6366F1;">
                        ${$.target_pindah_os?"Pindah: "+$.target_pindah_os:""}
                        ${$.target_selesai_os?" \u2192 Selesai: "+$.target_selesai_os:""}
                      </span>
                    </div>
                  `:""}
                </div>

                <!-- Notes box (if any) -->
                ${$.notes?`
                  <div class="emp-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${$.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="emp-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${$.id}
                  </div>

                  <div class="emp-action-btns">
                    ${$.phone?`
                      <a href="https://wa.me/${$.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${F?`
                      <button class="emp-icon-btn edit" data-action="edit" data-id="${$.id}" title="Edit Data Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="emp-icon-btn delete" data-action="delete" data-id="${$.id}" title="Nonaktifkan Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let k=e.querySelector("#emp-search-input");k&&k.addEventListener("input",$=>{u=$.target.value,b();let x=e.querySelector("#emp-search-input");x&&(x.focus(),x.setSelectionRange(x.value.length,x.value.length))});let T=e.querySelector("#emp-branch-filter");T&&T.addEventListener("change",$=>{g=$.target.value,b()});let A=e.querySelector("#emp-division-filter");A&&A.addEventListener("change",$=>{h=$.target.value,b()}),e.querySelectorAll(".emp-pill").forEach($=>{$.addEventListener("click",()=>{p=$.dataset.status;let x=$.dataset.div;x&&x!=="all"?h=x:$.dataset.status==="all"&&(h="all"),b()})});let P=e.querySelector("#emp-reset-filter");P&&P.addEventListener("click",()=>{u="",g="all",h="all",p="all",b()});let _=e.querySelector("#emp-btn-add");_&&_.addEventListener("click",()=>C());let I=e.querySelector("#emp-btn-excel");I&&I.addEventListener("click",()=>E());let z=e.querySelector("#emp-btn-refresh");z&&z.addEventListener("click",()=>S()),e.querySelectorAll(".emp-card").forEach($=>{let x=Number($.dataset.id),B=i.find(O=>O.id===x);if(!B)return;let L=$.querySelector('[data-action="edit"]');L&&L.addEventListener("click",()=>C(B));let H=$.querySelector('[data-action="delete"]');H&&H.addEventListener("click",()=>w(B))})}await S()}V();ie();var St=[],ta=[];async function kn(){St=await ee(),ta=await _e()}var wt=async e=>{let t=[],i=1;for(;;){let p=await(await Promise.resolve().then(()=>(V(),Rt))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${i}`);if(!p.ok)break;let h=p.data?.data||p.data||[],g=Array.isArray(h)?h:[];if(t=t.concat(g),g.length<100||p.data?.pagination&&i>=p.data.pagination.pages)break;i++}return t};async function ot(e,t){await kn(),se({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>xt(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,p)=>p.end_date&&String(p.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ft(a)},{key:"status",label:"Status",render:a=>pe(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:St},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[h,g]=await Promise.all([wt("/api/employees?status=Aktif"),wt("/api/contracts")]);if(h.length>0){let u=g.filter(F=>F.status==="Aktif"),r=new Set(u.map(F=>F.employee_id)),f=h.filter(F=>!r.has(F.id)),D=`<p style="margin-bottom:12px">Data yang terbaca: <b>${h.length}</b> Karyawan Aktif, dan <b>${u.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${f.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;f.forEach(F=>{let S=g.filter(C=>C.employee_id===F.id),E='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(S.length>0){let C=S[0];E=`Pernah ada kontrak (Status: <b style="color:#EF4444">${C.status}</b>, Selesai: ${window.formatDate(C.end_date)})`}D+=`<li style="margin-bottom:8px"><b>${F.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${F.branch_name||"-"} | ${E}</span></li>`}),D+="</ul>",Promise.resolve().then(()=>(le(),ht)).then(F=>F.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:D,cancelText:"Tutup"}))}}catch(h){console.error(h)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let p=document.querySelector(".page-actions");p&&p.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:ta,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:St,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await M(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let p=a.data.data.map(h=>({"Nama Lengkap":h.employee_name,Cabang:h.branch_name||"","Div / Bagian":h.division||"","Tanggal Mulai":h.start_date||"","Tanggal Selesai":h.end_date&&String(h.end_date).startsWith("2099")?"":h.end_date||"","Sisa Kontrak":h.end_date&&String(h.end_date).startsWith("2099")?"Tetap":h.days_remaining!==null&&h.days_remaining!==void 0?`${h.days_remaining} Hari`:"",Status:h.status||""}));Y(p,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[p,h]=await Promise.all([M("/api/branches?limit=10000"),wt("/api/employees")]),g=p.data?.data||[],u=h||[];console.log(`Total employee yang berhasil dimuat dari database : ${u.length}`),u.length>0&&(console.log("Contoh 5 employee pertama:"),u.slice(0,5).forEach((y,b)=>{console.log(`${b+1}. ID: ${y.id}, Name: ${y.full_name}, Status: ${y.status}`)}));let r=y=>{if(!y)return null;let b=String(y||"").replace(/\s+/g," ").toLowerCase().trim(),n=g.find(o=>String(o.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(o.code||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(o.name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return n?n.id:null},f=(y,b)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${b}`),console.log(`Nama dari Excel : "${y}"`),!y)return console.log("Alasan gagal mapping : Nama kosong"),null;let n=String(y||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${n}"`),console.log(`Jumlah employee di database : ${u.length}`);let o=u.find(s=>String(s.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===n);return o?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${o.id}`),o.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},D=y=>{if(!y)return"";if(y instanceof Date&&!isNaN(y.getTime()))return y.toISOString().slice(0,10);let b=String(y).trim();if(/^\d{4,5}(\.\d+)?$/.test(b)){let o=Math.floor(Number(b));if(o>2e4&&o<99999){let s=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(s.getTime())?"":s.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let n=b.split(/[\/\-\.]/);if(n.length===3){let[o,s,d]=n.map(m=>m.trim());if(o.length===4&&s.length<=2&&d.length<=2)return`${o}-${s.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&s.length<=2&&o.length<=2)return`${d}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`}return b},F=a.map((y,b)=>{let n=b+2,o=String(y["Nama Lengkap"]||"").trim(),s=y["Tanggal Mulai"],d=D(s);if(!d){let c=a.__worksheet,v=a.__headers||[],k=v.indexOf("Tanggal Mulai"),T="N/A",A="N/A",P="N/A";if(k!==-1&&c&&window.XLSX){let I=window.XLSX.utils.encode_cell({c:k,r:n-1});P=I;let z=c[I];z?(T=z.t||"undefined",A=z.w||"undefined"):T="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let _="Unknown";s==null||s===""?_="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":s instanceof Date&&isNaN(s.getTime())?_="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":_="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${n}`),console.log(`Employee Name : ${o}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${k})`),console.log(`Raw Cell Value : "${s}"`),console.log(`JavaScript Type : ${typeof s}`),console.log(`SheetJS Cell Type : ${T}`),console.log(`SheetJS Formatted Value : "${A}"`),console.log(`Value After Trim : "${String(s||"").trim()}"`),console.log(`Value After Date Parser : "${d}"`),console.log(`Is Empty : ${!s}`),console.log(`Is Invalid Date : ${s instanceof Date?isNaN(s.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${_}`),console.log(`Workbook Sheet : ${c?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${P}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(y,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(v)),console.log(`==========================
`)}let m=f(o,n),l=null;return m?d||(l="Tanggal Mulai kosong atau tidak berformat tanggal"):l="Karyawan tidak ditemukan di Database",{isValid:!!(m&&d),invalidReason:l,rowNum:n,data:{employee_id:m,branch_id:r(String(y.Cabang||"").trim()),division:String(y["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:d,end_date:D(y["Tanggal Selesai"])||"2099-12-31",status:String(y.Status||"").trim(),_rawName:o}}}),S=[],E=[];if(F.forEach(y=>{y.isValid?S.push(y.data):E.push({rowNum:y.rowNum,name:y.data._rawName,reason:y.invalidReason})}),console.log(`Split Validation - Valid: ${S.length}, Invalid: ${E.length}`),S.length===0){let y=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${E.length}

Daftar Kegagalan (Contoh):
`;E.slice(0,10).forEach(b=>{y+=`- Row ${b.rowNum} | Nama: ${b.name} | Alasan: ${b.reason}
`}),E.length>10&&(y+=`- ... dan ${E.length-10} lainnya.
`),alert(y);return}let C=await M("/api/contracts/import",{method:"POST",body:JSON.stringify(S)}),w=`IMPORT SUMMARY
======================
`;w+=`Total Baris Excel : ${a.length}
`,w+=`Baris Valid       : ${S.length}
`,w+=`Baris Invalid     : ${E.length}

`,C&&C.data&&C.data.metrics?(w+=`Berhasil INSERT   : ${C.data.metrics.inserted}
`,w+=`Berhasil UPDATE   : ${C.data.metrics.updated}
`):w+=`Berhasil diproses : ${S.length}
`,E.length>0&&(w+=`
DAFTAR DATA DILEWATI:
`,E.forEach(y=>{w+=`- Row ${y.rowNum} | ${y.name} | ${y.reason}
`})),alert(w),typeof ot=="function"&&ot()}}})}V();ie();var Et=[],Je=[];function wn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(a=>a.period===i))return i;return"Q3"}async function aa(e,t){Et=await ee();let i=await ge();Je=["Berlin Ariansyah","Ade Surahman"];let a=y=>y&&!Je.find(b=>String(typeof b=="object"?b.value:b).toLowerCase()===String(y).toLowerCase())?[...Je,y]:Je,p=await M(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),h=y=>{if(!y||y==="-"||String(y).trim()==="")return"";let b=String(y).split("-");return b.length===3&&b[0].length===4?`${b[2]}-${b[1]}-${b[0]}`:y},g=p.data?.data||[],u=wn(g),r=t?t.get("dash_filter"):null,f=new Date,D=`${f.getFullYear()}-${String(f.getMonth()+1).padStart(2,"0")}`,F={},S=t&&t.get("month")?t.get("month"):null;r==="inspeksi"?F={status:"Done",activity_type:"Inspeksi Hygiene",month:S}:r==="gcdc"?F={status:"Done",activity_type:"GCDC",month:S}:r&&r.startsWith("period_")&&(F={period:r.replace("period_","").toUpperCase()});let E=new Date().getFullYear(),w=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((y,b)=>{let n=String(b+1).padStart(2,"0");return{value:`${E}-${n}`,label:`${y} ${E}`}});se({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:F,onDataLoaded:y=>y.sort((b,n)=>{let o=b.opening_date?new Date(b.opening_date).getTime():0;return(n.opening_date?new Date(n.opening_date).getTime():0)-o}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:y=>yt(y)},{key:"period",label:"Periode",render:y=>ye(y)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:y=>h(y)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:y=>h(y)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:y=>h(y)},{key:"status",label:"Status",render:y=>pe(y)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Et},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:w},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Je}],formFields:y=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Et,value:y?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:y?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:y?.period},{name:"pic",label:"PIC",type:"combobox",options:a(y?.pic),value:y?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:y?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:y?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:y?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:y?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:y?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let y=await M(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(y.ok){let b=y.data.data.map(n=>({Cabang:n.branch_name||"",Kegiatan:n.activity_type||"",Periode:n.period||"",PIC:n.pic||"","Tgl Opening":n.opening_date||"","Tgl Target":n.target_date||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));Y(b,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async y=>{let n=(await M("/api/branches?all=1")).data?.data||[],o=l=>{if(!l)return null;let c=String(l||"").toLowerCase().trim(),v=c.match(/^(\d{3})/);if(v){let T=n.find(A=>String(A.code||"")===v[1]);if(T)return T.id}let k=n.find(T=>String(T.full_name||"").toLowerCase()===c||String(T.code||"").toLowerCase()===c||String(T.name||"").toLowerCase()===c);return k?k.id:null},s=l=>{if(l==null||l==="")return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let c=String(l).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let k=Number(c);if(k>2e4&&k<99999){let T=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}let v=c.split(/[\/\-\.]/);if(v.length===3){let[k,T,A]=v.map(P=>P.trim());if(k.length===4&&T.length<=2&&A.length<=2)return`${k}-${T.padStart(2,"0")}-${A.padStart(2,"0")}`;if(A.length===4&&T.length<=2&&k.length<=2)return`${A}-${T.padStart(2,"0")}-${k.padStart(2,"0")}`}return c},d=y.map(l=>({branch_id:o(String(l.Cabang||"").trim()),branch_name:String(l.Cabang||"").trim(),activity_type:String(l.Kegiatan||"").trim(),period:String(l.Periode||"").trim(),pic:String(l.PIC||l.Pic||"").trim(),opening_date:s(l["Tgl Opening"]||l["Tanggal Opening"]||l["Tgl Openir"]),target_date:s(l["Tgl Target"]||l["Tanggal Target"]),completion_date:s(l["Tgl Selesai"]||l["Tanggal Selesai"]),status:String(l.Status||"").trim(),notes:String(l.Catatan||l.Keterangan||"").trim()})).filter(l=>l.activity_type&&l.period),m=await M("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return be("schedule"),m.data}}})}V();le();var Ct=[{id:1,report_date:"2026-08-27",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"AM",complaint:"Keluhan AM fc atas nama Malik bau badan, dan sudah dibilangin berulang kali terkait kebersihan seragam dan deodoran.",employee_name:"MALIK",fc_specialist:"Mizwar",solution:"Diberikan teguran lisan, edukasi kebersihan diri, dan pendampingan seragam cadangan di klinik.",status:"Done",completion_date:"2026-08-28",day_count:1},{id:2,report_date:"2026-08-26",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Security atas nama Reza suhelmi meninggalkan area tugas jaga tanpa personil pengganti saat jam operasional klinik.",employee_name:"REZA SUHELMI",fc_specialist:"Ade Surahman",solution:"Pemanggilan investigasi dan penerbitan Surat Peringatan (SP) sesuai SOP security.",status:"Done",completion_date:"2026-08-27",day_count:1},{id:3,report_date:"2026-07-27",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"Terinfo kembali permasalahan FC Ciputat, terkait standar pembersihan area dental unit yang belum memenuhi checklist hygiene.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Re-training modul pembersihan dental unit dan monitoring harian oleh supervisor.",status:"In Progress",completion_date:null,day_count:null},{id:4,report_date:"2026-06-24",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"FC atas nama Hendra mengintimidasi rcp pada saat briefing pagi mengenai pembagian area kerja.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Mediasi bersama AM dan HR, diberikan pembinaan konseling perilaku kerja.",status:"Done",completion_date:"2026-06-25",day_count:1},{id:5,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Kinerja kurang dan ada komplain dari cabang mengenai ketidakhadiran berkala personil pengganti.",employee_name:"Indra",fc_specialist:"Ade",solution:"Evaluasi penempatan dan pembinaan kedisiplinan kerja.",status:"Done",completion_date:"2026-06-08",day_count:3},{id:6,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Ade",complaint:"Sering tidak masuk kerja dan memalsukan surat dokter saat izin sakit.",employee_name:"Kristian",fc_specialist:"Ade Surahman",solution:"Investigasi ke klinik dokter terkait dan eskalasi ke HRD untuk tindakan tegas.",status:"Done",completion_date:"2026-06-07",day_count:2},{id:7,report_date:"2026-05-25",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"SPV",complaint:"Tidak masuk kerja tanpa kabar (alpa), dan sudah di-one on one oleh SPV terkait absensi.",employee_name:"Malik",fc_specialist:"Mizwar",solution:"Pemberian SP 1 dan penandatanganan komitmen kehadiran kerja.",status:"Done",completion_date:"2026-05-26",day_count:1}],na=["SDM","Cleaning","Aset","K3","Lainnya"],Sn=["Open","In Progress","Done"],En=["SPV","AM","RCP","Perawat","FC","Security","Berlin","Ade","Pattrel","Dentrel"];function ia(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Cn(e){if(!e||e==="-")return"\u26A0\uFE0F";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function ra(e,t){let i=[],a=[],p=[],h="all",g="all",u="all",r="",f=!1;(t?t.get("dash_filter"):null)==="open"&&(h="Open");let S=Z()?.role!=="viewer";try{[a,p]=await Promise.all([ee(),ge()])}catch(s){console.error("Error fetching cache:",s)}async function E(){f=!0,o();try{let s=await M("/api/issues?limit=10000");if(s.ok){let d=s.data?.data||s.data||[];i=Array.isArray(d)&&d.length>0?d:Ct}else i=Ct}catch(s){console.warn("Gagal load issues, menggunakan data lokal:",s),i=Ct}finally{f=!1,o()}}async function C(){try{let{downloadExcel:s}=await Promise.resolve().then(()=>(ie(),he)),d=i.map((m,l)=>({No:l+1,"Tanggal Info":m.report_date||"",Cabang:m.branch_name||"",Kategori:m.category||"","Sumber Laporan":m.source||"",Keluhan:m.complaint||"","Nama FC / Security":m.employee_name||"","FC Spesialis":m.fc_specialist||"","Solusi / Tindak Lanjut":m.solution||"",Status:m.status||"","Tanggal Selesai":m.completion_date||"","Durasi Hari":m.day_count!=null?m.day_count:""}));s(d,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),K("Berhasil mengekspor data permasalahan ke Excel")}catch(s){N("Gagal ekspor Excel: "+s.message)}}async function w(s,d){try{let m=d==="Done"?s.completion_date||new Date().toISOString().split("T")[0]:d==="Open"?null:s.completion_date,l=s.day_count;if(s.report_date&&m){let k=new Date(s.report_date),T=new Date(m);l=Math.max(0,Math.floor((T-k)/(1e3*60*60*24)))}else d==="Open"&&(l=null);let c={...s,status:d,completion_date:m,day_count:l},v=await M(`/api/issues/${s.id}`,{method:"PUT",body:JSON.stringify(c)});v.ok?(K(`Status masalah diubah ke "${d}"`),await E()):N(v.data?.error||"Gagal memperbarui status")}catch(m){N("Terjadi kesalahan: "+m.message)}}function y(s=null){let d=!!s,m=new Date().toISOString().split("T")[0],l=s?.report_date||m,c=s?.category||"SDM",v=s?.source||"AM",k=s?.status||"Open",T=document.createElement("div");T.className="issue-modal-body",T.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${a.map($=>`
                <option value="${$.value}" ${String(s?.branch_id)===String($.value)?"selected":""}>
                  ${$.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Kategori <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-category" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${na.map($=>`
                <option value="${$}" ${c===$?"selected":""}>${$}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <input list="iss-source-options" id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: AM, SPV, Berlin..." value="${v}" />
            <datalist id="iss-source-options">
              ${En.map($=>`<option value="${$}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${s?.complaint||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <input list="iss-emp-options" id="iss-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Ketik nama karyawan terkait..." value="${s?.employee_name?String(s.employee_name).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-emp-options">
              ${p.map($=>`<option value="${$.label}">`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${s?.fc_specialist?String(s.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${p.map($=>`<option value="${$.label}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${s?.solution||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Sn.map($=>`
                <option value="${$}" ${k===$?"selected":""}>${$}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${s?.completion_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC;" 
              value="${s?.day_count!=null?s.day_count:""}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;let A=T.querySelector("#iss-modal-date"),P=T.querySelector("#iss-modal-completion"),_=T.querySelector("#iss-modal-daycount"),I=T.querySelector("#iss-modal-status");function z(){if(A.value&&P.value){let $=new Date(A.value),x=new Date(P.value),B=Math.floor((x-$)/(1e3*60*60*24));_.value=isNaN(B)?"":Math.max(0,B)}else _.value=""}A.addEventListener("change",z),P.addEventListener("change",z),I.addEventListener("change",()=>{I.value==="Done"&&!P.value?(P.value=new Date().toISOString().split("T")[0],z()):I.value==="Open"&&(P.value="",_.value="")}),J({title:d?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:T,size:"md",confirmText:d?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async($,x)=>{let B=T.querySelector("#iss-modal-date").value,L=T.querySelector("#iss-modal-branch").value,H=T.querySelector("#iss-modal-category").value,O=(T.querySelector("#iss-modal-source").value||"").trim(),j=(T.querySelector("#iss-modal-complaint").value||"").trim(),q=(T.querySelector("#iss-modal-employee").value||"").trim(),R=(T.querySelector("#iss-modal-spec").value||"").trim(),Q=(T.querySelector("#iss-modal-solution").value||"").trim(),ne=T.querySelector("#iss-modal-status").value,oe=T.querySelector("#iss-modal-completion").value||null;if(!B||!L||!j||!H){N("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let G=null;if(B&&oe){let X=new Date(B),U=new Date(oe);G=Math.max(0,Math.floor((U-X)/(1e3*60*60*24)))}let ae={report_date:B,branch_id:Number(L),category:H,source:O,complaint:j,employee_name:q,fc_specialist:R,solution:Q,status:ne,completion_date:oe,day_count:G};try{let X;d?X=await M(`/api/issues/${s.id}`,{method:"PUT",body:JSON.stringify(ae)}):X=await M("/api/issues",{method:"POST",body:JSON.stringify(ae)}),X.ok?(K(d?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),x(),await E()):N(X.data?.error||"Gagal menyimpan permasalahan")}catch(X){N("Terjadi kesalahan: "+X.message)}}})}async function b(s){W({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${s.branch_name||"cabang"}</strong>:<br><br><em>"${s.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let d=await M(`/api/issues/${s.id}`,{method:"DELETE"});d.ok?(K("Permasalahan berhasil dihapus"),await E()):N(d.data?.error||"Gagal menghapus data")}catch(d){N("Terjadi kesalahan: "+d.message)}}})}function n(){return i.filter(s=>{if(r){let d=r.toLowerCase(),m=(s.complaint||"").toLowerCase().includes(d),l=(s.employee_name||"").toLowerCase().includes(d),c=(s.branch_name||"").toLowerCase().includes(d),v=(s.source||"").toLowerCase().includes(d),k=(s.category||"").toLowerCase().includes(d),T=(s.solution||"").toLowerCase().includes(d);if(!m&&!l&&!c&&!v&&!k&&!T)return!1}return!(h!=="all"&&s.status!==h||g!=="all"&&s.category!==g||u!=="all"&&String(s.branch_id)!==String(u))})}function o(){let s=n(),d=i.length,m=i.filter(j=>j.status==="Open"||j.status==="In Progress").length,l=i.filter(j=>j.status==="Done").length,c=new Date,v=c.getFullYear(),k=c.getMonth(),T=i.filter(j=>{if(!j.report_date)return!1;let q=new Date(j.report_date);return!isNaN(q.getTime())&&q.getFullYear()===v&&q.getMonth()===k}).length,A=i.length,P=i.filter(j=>j.status==="Open").length,_=i.filter(j=>j.status==="In Progress").length,I=i.filter(j=>j.status==="Done").length;e.innerHTML=`
      <style>
        .issues-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .issues-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .issues-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(245, 158, 11, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .issues-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .issues-stat-box {
          background: rgba(255, 255, 255, 0.05);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 18px;
          display: flex;
          align-items: center;
          gap: 14px;
          backdrop-filter: blur(8px);
          transition: transform 0.15s ease, background 0.15s ease;
        }
        .issues-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .issues-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .issues-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .issues-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .issues-controls-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .issues-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .issues-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .issues-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .issues-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .issues-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .issues-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 160px;
          cursor: pointer;
        }
        .issues-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .issues-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .issues-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .issues-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .issues-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .issues-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .issues-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .issues-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .issues-pill.active .issues-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* Cards Grid */
        .issues-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(370px, 1fr));
          gap: 16px;
        }
        .issues-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 14px;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .issues-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }
        .issues-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
        }
        .issues-category-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .issues-category-sdm {
          background: #EEF2FF;
          color: #4F46E5;
          border: 1px solid #E0E7FF;
        }
        .issues-category-cleaning {
          background: #ECFEFF;
          color: #0891B2;
          border: 1px solid #CFFAFE;
        }
        .issues-category-aset {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }
        .issues-category-k3 {
          background: #FEF3C7;
          color: #D97706;
          border: 1px solid #FDE68A;
        }
        .issues-category-lainnya {
          background: #F1F5F9;
          color: #475569;
          border: 1px solid #E2E8F0;
        }

        .issues-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .issues-status-open {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .issues-status-inprogress {
          background: #DBEAFE;
          color: #1D4ED8;
          border: 1px solid #BFDBFE;
        }
        .issues-status-done {
          background: #DCFCE7;
          color: #15803D;
          border: 1px solid #BBF7D0;
        }

        /* Complaint Quote Box */
        .issues-complaint-box {
          background: #F8FAFC;
          border-left: 4px solid #F59E0B;
          border-radius: 0 10px 10px 0;
          padding: 12px 14px;
          font-size: 0.92rem;
          color: #0F172A;
          line-height: 1.5;
          font-weight: 600;
        }
        .issues-complaint-box.status-done {
          border-left-color: #10B981;
        }
        .issues-complaint-box.status-inprogress {
          border-left-color: #3B82F6;
        }

        /* Person / Location Box */
        .issues-person-box {
          background: #F8FAFC;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 12px;
          display: flex;
          align-items: center;
          gap: 12px;
        }
        .issues-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: #4F46E5;
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1rem;
          flex-shrink: 0;
        }
        .issues-person-name {
          font-size: 0.95rem;
          font-weight: 800;
          color: #0F172A;
        }
        .issues-person-sub {
          font-size: 0.74rem;
          color: #64748B;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.04em;
        }

        /* Details list */
        .issues-details-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 8px 12px;
          font-size: 0.82rem;
          color: #334155;
          padding: 2px 4px;
        }
        .issues-detail-item {
          display: flex;
          align-items: center;
          gap: 6px;
        }

        /* Solution box */
        .issues-solution-box {
          background: #F0FDF4;
          border: 1px solid #DCFCE7;
          border-radius: 10px;
          padding: 10px 12px;
          font-size: 0.82rem;
          color: #166534;
          line-height: 1.45;
        }
        .issues-solution-title {
          font-weight: 700;
          display: flex;
          align-items: center;
          gap: 6px;
          margin-bottom: 3px;
        }

        /* Card footer */
        .issues-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 12px;
          margin-top: auto;
          gap: 8px;
        }
        .issues-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .issues-icon-btn {
          width: 34px;
          height: 34px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .issues-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .issues-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }
        .issues-quick-btn {
          font-size: 0.76rem;
          font-weight: 700;
          padding: 6px 10px;
          border-radius: 8px;
          cursor: pointer;
          border: 1px solid transparent;
          display: inline-flex;
          align-items: center;
          gap: 4px;
          transition: all 0.15s ease;
        }
        .issues-quick-btn.start {
          background: #EFF6FF;
          color: #2563EB;
          border-color: #BFDBFE;
        }
        .issues-quick-btn.start:hover {
          background: #DBEAFE;
        }
        .issues-quick-btn.complete {
          background: #F0FDF4;
          color: #16A34A;
          border-color: #BBF7D0;
        }
        .issues-quick-btn.complete:hover {
          background: #DCFCE7;
        }
        .issues-quick-btn.reopen {
          background: #FFFBEB;
          color: #D97706;
          border-color: #FDE68A;
        }
        .issues-quick-btn.reopen:hover {
          background: #FEF3C7;
        }

        /* Empty state */
        .issues-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .issues-hero-card {
            padding: 18px 16px;
          }
          .issues-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .issues-stat-box {
            padding: 10px 12px;
          }
          .issues-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .issues-stat-val {
            font-size: 1.3rem;
          }
          .issues-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .issues-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .issues-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .issues-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .issues-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .issues-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .issues-select {
            width: 100%;
          }
          .issues-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="issues-page">

        <!-- Hero Header Card -->
        <div class="issues-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u26A0\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Log Permasalahan Operasional
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pusat monitoring keluhan operasional klinik, kendala personil FC & Security, serta rekaman investigasi dan progres penyelesaian masalah.
              </p>
            </div>

            <div class="issues-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${S?`
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              `:""}
              
              <button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${f?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val">${d}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val">${m}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val">${l}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">\u{1F4C5}</div>
              <div>
                <div class="issues-stat-val">${T}</div>
                <div class="issues-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="issues-controls-card">
          <div class="issues-search-row">
            <div class="issues-search-wrap">
              <svg class="issues-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="issues-search-input" class="issues-search-input" 
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${r?String(r).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${a.length})</option>
              ${a.map(j=>`
                <option value="${j.value}" ${String(u)===String(j.value)?"selected":""}>
                  ${j.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${na.map(j=>`
                <option value="${j}" ${g===j?"selected":""}>
                  Kategori: ${j}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${h==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count">${A}</span>
            </div>
            <div class="issues-pill ${h==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count">${P}</span>
            </div>
            <div class="issues-pill ${h==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count">${_}</span>
            </div>
            <div class="issues-pill ${h==="Done"?"active":""}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count">${I}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="issues-cards-grid">
          ${s.length===0?`
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:s.map(j=>{let q=j.category||"Lainnya",R=`issues-category-${q.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",Q=j.status||"Open",ne="issues-status-open",oe="\u26A0\uFE0F Open";Q==="In Progress"?(ne="issues-status-inprogress",oe="\u{1F504} In Progress"):Q==="Done"&&(ne="issues-status-done",oe="\u2705 Done");let G=Q==="Done"?"status-done":Q==="In Progress"?"status-inprogress":"";return`
              <div class="issues-card" data-id="${j.id}">
                <!-- Top Badges Row -->
                <div class="issues-card-top">
                  <span class="issues-category-badge ${R}">
                    \u{1F3F7}\uFE0F ${q}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="issues-status-badge ${ne}">
                      ${oe}
                    </span>
                    <span style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                      \u{1F5D3}\uFE0F ${ia(j.report_date)}
                    </span>
                  </div>
                </div>

                <!-- Complaint Box -->
                <div class="issues-complaint-box ${G}">
                  "${j.complaint||"-"}"
                </div>

                <!-- Person & Location Box -->
                <div class="issues-person-box">
                  <div class="issues-avatar-box">
                    ${Cn(j.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="issues-person-sub">PERSONIL FC / SECURITY</div>
                    <div class="issues-person-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${j.employee_name||"Tidak ada personil spesifik"}
                    </div>
                  </div>
                </div>

                <!-- Details Grid -->
                <div class="issues-details-grid">
                  <div class="issues-detail-item" style="grid-column: 1 / -1;">
                    <span style="color:#64748B;">\u{1F4CD} Cabang:</span>
                    <strong style="color:#0F172A;">${j.branch_name||"Cabang Tidak Diketahui"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F5E3}\uFE0F Sumber:</span>
                    <strong>${j.source||"-"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F6E1}\uFE0F Spesialis:</span>
                    <strong>${j.fc_specialist||"-"}</strong>
                  </div>
                </div>

                <!-- Solution & Completion Box (if available) -->
                ${j.solution||j.completion_date?`
                  <div class="issues-solution-box">
                    ${j.solution?`
                      <div class="issues-solution-title">
                        <span>\u{1F4A1} Tindak Lanjut / Solusi:</span>
                      </div>
                      <div style="margin-bottom:${j.completion_date?"6px":"0"};">
                        ${j.solution}
                      </div>
                    `:""}
                    ${j.completion_date?`
                      <div style="font-size:0.76rem; color:#15803D; font-weight:600; display:flex; align-items:center; gap:4px;">
                        <span>\u23F1\uFE0F Selesai: ${ia(j.completion_date)}</span>
                        ${j.day_count!=null?`<span>(Durasi: ${j.day_count} Hari)</span>`:""}
                      </div>
                    `:""}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="issues-card-footer">
                  <div>
                    ${S?Q==="Open"?`
                        <button class="issues-quick-btn start" data-action="quick-progress" data-id="${j.id}" title="Ubah status ke In Progress">
                          <span>\u{1F504} Mulai Tangani</span>
                        </button>
                      `:Q==="In Progress"?`
                        <button class="issues-quick-btn complete" data-action="quick-done" data-id="${j.id}" title="Tandai selesai">
                          <span>\u2705 Selesaikan</span>
                        </button>
                      `:`
                        <button class="issues-quick-btn reopen" data-action="quick-reopen" data-id="${j.id}" title="Buka kembali tiket masalah">
                          <span>\u21A9\uFE0F Buka Kembali</span>
                        </button>
                      `:""}
                  </div>

                  <div class="issues-action-btns">
                    ${S?`
                      <button class="issues-icon-btn edit" data-action="edit" data-id="${j.id}" title="Edit Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="issues-icon-btn delete" data-action="delete" data-id="${j.id}" title="Hapus Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let z=e.querySelector("#issues-search-input");z&&z.addEventListener("input",j=>{r=j.target.value,o();let q=e.querySelector("#issues-search-input");q&&(q.focus(),q.setSelectionRange(q.value.length,q.value.length))});let $=e.querySelector("#issues-branch-filter");$&&$.addEventListener("change",j=>{u=j.target.value,o()});let x=e.querySelector("#issues-category-filter");x&&x.addEventListener("change",j=>{g=j.target.value,o()}),e.querySelectorAll(".issues-pill").forEach(j=>{j.addEventListener("click",()=>{h=j.dataset.status,o()})});let B=e.querySelector("#issues-reset-filter");B&&B.addEventListener("click",()=>{r="",u="all",g="all",h="all",o()});let L=e.querySelector("#issues-btn-add");L&&L.addEventListener("click",()=>y());let H=e.querySelector("#issues-btn-excel");H&&H.addEventListener("click",()=>C());let O=e.querySelector("#issues-btn-refresh");O&&O.addEventListener("click",()=>E()),e.querySelectorAll(".issues-card").forEach(j=>{let q=Number(j.dataset.id),R=i.find(X=>X.id===q);if(!R)return;let Q=j.querySelector('[data-action="edit"]');Q&&Q.addEventListener("click",()=>y(R));let ne=j.querySelector('[data-action="delete"]');ne&&ne.addEventListener("click",()=>b(R));let oe=j.querySelector('[data-action="quick-progress"]');oe&&oe.addEventListener("click",()=>w(R,"In Progress"));let G=j.querySelector('[data-action="quick-done"]');G&&G.addEventListener("click",()=>w(R,"Done"));let ae=j.querySelector('[data-action="quick-reopen"]');ae&&ae.addEventListener("click",()=>w(R,"Open"))})}await E()}V();var Ie=[];function Fn(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function oa(e,t){let i=t?t.get("dash_filter"):null;Ie=await ee();let a=await ge(),p=["Ade","Berlin"],h=u=>u&&!a.find(r=>r.value===u)?[...a,{value:u,label:u}]:a,g=u=>u&&!p.find(r=>(typeof r=="object"?r.value:r)===u)?[...p,u]:p;se({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:u=>i?u.filter(r=>Fn(r,i)):u,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:u=>`<span title="${u||""}">${u?.length>50?u.slice(0,50)+"\u2026":u||"-"}</span>`},{key:"solution",label:"Solusi",render:u=>`<span title="${u||""}">${u?.length>40?u.slice(0,40)+"\u2026":u||"-"}</span>`},{key:"status",label:"Status",render:u=>pe(u)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:u=>window.formatDate(u)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:u=>u?`<a href="${u}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Ie},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async u=>{let r=new URLSearchParams(u||{}).toString(),f=await M(`/api/one-on-one?limit=10000&${r}`);if(f.ok){let D=f.data.data.map(S=>({Tanggal:S.meeting_date||"",Cabang:S.branch_name||"","Nama Karyawan":S.employee_name||"",PIC:S.pic||"",Masalah:S.problem||"",Solusi:S.solution||"",Status:S.status||"","Tgl Selesai":S.completion_date||"",Dokumen:S.document_link||""})),{downloadExcel:F}=await Promise.resolve().then(()=>(ie(),he));F(D,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let u=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(ie(),he));r(u,"Template_Import_OneOnOne")},onImport:async u=>{let r=S=>{if(!S)return null;let E=String(S||"").toLowerCase(),C=Ie.find(w=>String(w.label||"").toLowerCase()===E);return C?C.value:null},f=S=>{if(!S)return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let E=String(S).trim();if(/^\d{4,5}$/.test(E)){let w=Number(E);if(w>2e4&&w<99999){let y=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(E))return E.slice(0,10);let C=E.split(/[\/\-\.]/);if(C.length===3){let[w,y,b]=C.map(n=>n.trim());if(w.length===4&&y.length<=2&&b.length<=2)return`${w}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&y.length<=2&&w.length<=2)return`${b}-${y.padStart(2,"0")}-${w.padStart(2,"0")}`}return E},D=u.map(S=>({meeting_date:f(S.Tanggal),employee_name:String(S["Nama Karyawan"]||"").trim(),branch_id:r(String(S.Cabang||"").trim()),pic:String(S.PIC||"").trim(),problem:String(S.Masalah||"").trim(),solution:String(S.Solusi||"").trim(),status:String(S.Status||"").trim(),completion_date:f(S["Tgl Selesai"]),document_link:String(S.Dokumen||"").trim()})).filter(S=>S.meeting_date&&S.employee_name&&S.branch_id),F=await M("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:D,onDuplicate:"update"})});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}},formFields:u=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:u?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:u?.branch_id&&!Ie.find(r=>r.value==u.branch_id)?[...Ie,{value:u.branch_id,label:u.branch_name||u.branch_id}]:Ie,createApi:{path:"/api/branches",field:"full_name"},value:u?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:h(u?.employee_name),value:u?.employee_name},{name:"pic",label:"PIC",type:"select",options:g(u?.pic),createApi:{path:"/api/pic",field:"name"},value:u?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:u?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:u?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:u?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:u?.document_link}]})}V();le();var Ft=[{id:1,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:6,branch_name:"006. Kelapa Gading MOI",trainer:"Berlin Ariansyah",participants:"MUHAMMAD PRAKAS AGAZI",score:90,document_link:""},{id:2,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:11,branch_name:"011. Kuningan",trainer:"Berlin Ariansyah",participants:"KAMALLUDIN",score:88,document_link:""},{id:3,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:42,branch_name:"042. Muara Karang",trainer:"Berlin Ariansyah",participants:"SAHRUL ICHSANI",score:85,document_link:""},{id:4,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:51,branch_name:"051. Bintaro Sektor 9",trainer:"Berlin Ariansyah",participants:"MUHAMAD VICKRI FATHURROHMAN",score:92,document_link:""},{id:5,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:45,branch_name:"045. FX Mall",trainer:"Berlin Ariansyah",participants:"DEDE AHDORI",score:86,document_link:""},{id:6,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:54,branch_name:"054. Bogor Gunung Batu",trainer:"Berlin Ariansyah",participants:"ADITYA PRATAMA",score:95,document_link:""},{id:7,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:55,branch_name:"055. Gajah Mada",trainer:"Berlin Ariansyah",participants:"ELVAN JANGGA PRAMANA",score:89,document_link:""}],sa=["Batch 5","Batch 4","Batch 3","Batch 2","Batch 1"],la=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function $n(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Be(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function da(e){let t=[],i=[],a=[],p="all",h="all",g="all",u="",r=!1,D=Z()?.role!=="viewer";try{[i,a]=await Promise.all([ee(),ge()])}catch(n){console.error("Error fetching cache:",n)}async function F(){r=!0,y();try{let n=await M("/api/training?limit=10000");if(n.ok){let o=n.data?.data||n.data||[];t=Array.isArray(o)&&o.length>0?o:Ft}else t=Ft}catch(n){console.warn("Gagal load training, menggunakan data lokal:",n),t=Ft}finally{r=!1,y()}}async function S(){try{let{downloadExcel:n}=await Promise.resolve().then(()=>(ie(),he)),o=t.map((s,d)=>({No:d+1,Tanggal:s.training_date||"",Batch:s.batch||"",Materi:s.subject||"",Cabang:s.branch_name||"",Trainer:s.trainer||"",Peserta:Be(s.participants),Nilai:s.score!=null?s.score:"",Dokumen:s.document_link||"",Catatan:s.notes||""}));n(o,`Data_Training_${new Date().toISOString().slice(0,10)}`),K("Data training berhasil diekspor ke Excel!")}catch(n){N("Gagal ekspor Excel: "+n.message)}}async function E(n){if(!D)return;let o=t.find(s=>s.id===n);W({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${o?.subject||"ini"}" untuk ${Be(o?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let s=await M(`/api/training/${n}`,{method:"DELETE"});s.ok?(K("Data training berhasil dihapus"),t=t.filter(d=>d.id!==n),y()):N(s.data?.error||"Gagal menghapus data training")}catch(s){N("Gagal menghapus training: "+s.message)}}})}function C(n=null){if(!D)return;let o=!!n,s=n?.training_date?n.training_date.slice(0,10):new Date().toISOString().slice(0,10),d=n?.batch||"Batch 5",m=n?.trainer||"Berlin Ariansyah",l=Be(n?.participants)!=="-"?Be(n?.participants):"",c=document.createElement("div");c.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${s}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-batch-options" id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik batch..." value="${d}" required />
            <datalist id="tr-batch-options">
              ${sa.map(v=>`<option value="${v}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${n?.subject?String(n.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(v=>`
                <option value="${v.value}" ${String(n?.branch_id)===String(v.value)?"selected":""}>
                  ${v.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${m}" required />
            <datalist id="tr-trainer-options">
              ${la.map(v=>`<option value="${v}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${l}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${n?.score!=null?n.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${n?.document_link?String(n.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${n?.notes||""}</textarea>
        </div>
      </div>
    `,J({title:o?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:c,size:"md",confirmText:o?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(v,k)=>{let T=c.querySelector("#tr-modal-date").value,A=(c.querySelector("#tr-modal-batch").value||"").trim(),P=(c.querySelector("#tr-modal-subject").value||"").trim(),_=c.querySelector("#tr-modal-branch").value,I=(c.querySelector("#tr-modal-trainer").value||"").trim(),z=(c.querySelector("#tr-modal-participants").value||"").trim(),$=c.querySelector("#tr-modal-score").value,x=$!==""?Number($):null,B=(c.querySelector("#tr-modal-link").value||"").trim(),L=(c.querySelector("#tr-modal-notes").value||"").trim();if(!T){N("Tanggal training wajib diisi");return}if(!P){N("Materi / topik pelatihan wajib diisi");return}if(!_){N("Cabang klinik wajib dipilih");return}if(!I){N("Trainer wajib diisi");return}if(!z){N("Nama peserta training wajib diisi");return}let H=i.find(j=>String(j.value)===String(_)),O={training_date:T,batch:A,subject:P,branch_id:Number(_),branch_name:H?H.label:"Cabang",trainer:I,participants:z,score:x,document_link:B,notes:L};try{if(o){let j=await M(`/api/training/${n.id}`,{method:"PUT",body:JSON.stringify(O)});if(j.ok){K("Data training berhasil diperbarui");let q=t.findIndex(R=>R.id===n.id);q!==-1&&(t[q]={...t[q],...O}),k(),y()}else N(j.data?.error||"Gagal memperbarui data training")}else{let j=await M("/api/training",{method:"POST",body:JSON.stringify(O)});if(j.ok){K("Sesi training baru berhasil ditambahkan");let q=j.data?.id||Date.now();t.unshift({id:q,...O}),k(),y()}else N(j.data?.error||"Gagal menambahkan training")}}catch(j){N("Terjadi kesalahan: "+j.message)}}})}function w(){return t.filter(n=>{if(u){let o=u.toLowerCase(),s=(n.subject||"").toLowerCase().includes(o),d=(n.trainer||"").toLowerCase().includes(o),m=Be(n.participants).toLowerCase().includes(o),l=(n.branch_name||"").toLowerCase().includes(o),c=(n.batch||"").toLowerCase().includes(o);if(!s&&!d&&!m&&!l&&!c)return!1}return!(p!=="all"&&n.batch!==p||h!=="all"&&String(n.branch_id)!==String(h)||g!=="all"&&n.trainer!==g)})}function y(){let n=w(),o=t.length,s=new Set;t.forEach(A=>{Be(A.participants).split(",").forEach(_=>{let I=_.trim();I&&I!=="-"&&s.add(I)})});let d=s.size,l=new Set(t.map(A=>A.batch).filter(Boolean)).size,c=new Date,v=c.getFullYear(),k=c.getMonth(),T=t.filter(A=>{if(!A.training_date)return!1;let P=new Date(A.training_date);return!isNaN(P.getTime())&&P.getFullYear()===v&&P.getMonth()===k}).length;e.innerHTML=`
      <style>
        .training-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .training-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .training-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(147, 51, 234, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .training-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 22px;
        }
        .training-stat-box {
          background: rgba(255, 255, 255, 0.06);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 16px;
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .training-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .training-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          color: #F8FAFC;
        }
        .training-stat-lbl {
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #94A3B8;
          font-weight: 600;
          margin-top: 2px;
        }

        /* Controls Card */
        .training-controls-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .training-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .training-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .training-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .training-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .training-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .training-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 180px;
          cursor: pointer;
        }
        .training-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .training-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .training-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .training-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .training-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .training-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* Grid Cards View */
        .training-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .training-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          transition: transform 0.15s ease, box-shadow 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          position: relative;
        }
        .training-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.06);
          border-color: #CBD5E1;
        }
        .training-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: linear-gradient(135deg, #7C3AED 0%, #6D28D9 100%);
          color: white;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          box-shadow: 0 3px 6px rgba(124, 58, 237, 0.25);
        }
        .training-batch-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
          background: #F3E8FF;
          color: #7E22CE;
          border: 1px solid #E9D5FF;
        }
        .training-score-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
          background: #DCFCE7;
          color: #15803D;
          border: 1px solid #BBF7D0;
        }
        .training-info-box {
          background: #F8FAFC;
          border: 1px solid #E2E8F0;
          border-radius: 10px;
          padding: 10px 12px;
          margin: 12px 0;
          display: flex;
          flex-direction: column;
          gap: 6px;
        }
        .training-icon-btn {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .training-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .training-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .training-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .training-hero-card {
            padding: 18px 16px;
          }
          .training-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .training-stat-box {
            padding: 10px 12px;
          }
          .training-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .training-stat-val {
            font-size: 1.3rem;
          }
          .training-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .training-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .training-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .training-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .training-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .training-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .training-select {
            width: 100%;
          }
          .training-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="training-page">

        <!-- Hero Header Card -->
        <div class="training-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u{1F393}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Training Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Sistem evaluasi pengembangan kompetensi, sertifikasi keahlian, dan arsip riwayat pelatihan personil PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div class="training-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${D?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              <button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="tr-btn-refresh" class="btn btn-ghost training-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="training-stats-row">
            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(124,58,237,0.25); color:#A78BFA;">
                \u{1F393}
              </div>
              <div>
                <div class="training-stat-val">${o}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val">${d}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val">${l} Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val">${T}</div>
                <div class="training-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="training-controls-card">
          <div class="training-search-row">
            <div class="training-search-wrap">
              <span class="training-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${u}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(A=>`<option value="${A.value}" ${String(h)===String(A.value)?"selected":""}>${A.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${la.map(A=>`<option value="${A}" ${g===A?"selected":""}>${A}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${p==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${o}</span>
            </div>
            ${sa.map(A=>{let P=t.filter(_=>_.batch===A).length;return`
                <div class="training-pill ${p===A?"active":""}" data-batch="${A}">
                  <span>\u{1F393}</span>
                  <span>${A}</span>
                  ${P>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${P}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${n.length===0?`
            <div class="training-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Training Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="tr-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <div class="training-cards-grid">
              ${n.map(A=>{let P=Be(A.participants),I=P.split(",")[0].trim().split(" ").map(z=>z[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${A.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${A.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${A.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${A.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${$n(A.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${A.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${I}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${P}">
                            ${P}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${A.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${A.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${A.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${A.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${A.document_link?`
                          <a href="${A.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${D?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${A.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${A.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        </div>
                      `:""}
                    </div>

                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,b()}function b(){let n=e.querySelector("#tr-btn-add");n&&(n.onclick=()=>C());let o=e.querySelector("#tr-btn-excel");o&&(o.onclick=()=>S());let s=e.querySelector("#tr-btn-refresh");s&&(s.onclick=()=>F());let d=e.querySelector("#tr-btn-reset-filter");d&&(d.onclick=()=>{u="",h="all",p="all",g="all",y()});let m=e.querySelector("#tr-search-input");m&&(m.oninput=v=>{u=v.target.value,y();let k=e.querySelector("#tr-search-input");k&&(k.focus(),k.setSelectionRange(k.value.length,k.value.length))});let l=e.querySelector("#tr-filter-branch");l&&(l.onchange=v=>{h=v.target.value,y()});let c=e.querySelector("#tr-filter-trainer");c&&(c.onchange=v=>{g=v.target.value,y()}),e.querySelectorAll(".training-pill").forEach(v=>{v.onclick=()=>{p=v.getAttribute("data-batch")||"all",y()}}),e.querySelectorAll(".training-edit-btn").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id")),A=t.find(P=>P.id===T);A&&C(A)}}),e.querySelectorAll(".training-delete-btn").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id"));E(T)}})}await F()}V();le();ie();function pa({container:e,title:t,icon:i,apiPath:a,columns:p,formFields:h,filterFields:g,defaultFilters:u={},itemLabel:r="Data",canCreate:f=!0,canEdit:D=!0,canDelete:F=!0,onBeforeSubmit:S,onAfterLoad:E,onDataLoaded:C,extraActions:w=[],initialSearch:y="",exportOptions:b=null,bulkDelete:n=!1,paginationMode:o="server"}){let s=Z();s&&s.role==="viewer"&&(f=!1,D=!1,F=!1,n=!1,b=null),s&&s.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(f=!1,D=!1,F=!1,n=!1,b=null),s&&s.role==="input_lembur"&&(a==="/api/overtime"?(F=!1,n=!1,b=null):(f=!1,D=!1,F=!1,n=!1,b=null));let d=1,m={...u};y&&(m.search=y);let l=new Set;e.innerHTML=`
    ${n?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; justify-content:space-between; background:#2563EB; padding:12px 16px; border-radius:8px; margin-bottom:16px; box-shadow:0 4px 6px -1px rgba(37,99,235,0.2);">
      <button id="btn-bulk-cancel" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6L6 18M6 6l12 12"></path></svg>
      </button>
      <span id="bulk-count" style="font-weight:500; font-size:0.95rem; color:white;">0 item dipilih</span>
      <button id="btn-bulk-delete" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
      </button>
    </div>`:""}

    <div class="page-header">
      <h1 class="page-title">${i} ${t}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${f?`<button class="btn btn-primary" id="btn-create">+ Tambah ${r}</button>`:""}
        ${b?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${b.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${b.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${b.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${b.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
              </label>

            </div>
          </div>
          <style>
            .show-aksi-menu { display: flex !important; }
            .dropdown-item:hover { background-color: #F8FAFC !important; color: #2563EB !important; }
          </style>
        `:""}
      </div>
    </div>
    

    ${g&&g.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${g.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${m.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${g.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let B=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${B}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${m[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${g.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let B=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${B}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${m[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function c(){let x=document.getElementById("bulk-toolbar");if(!x)return;let B=document.getElementById("bulk-count"),L=document.getElementById("btn-bulk-delete"),H=document.getElementById("btn-bulk-cancel");B.textContent=`${l.size} item dipilih`,l.size>0?(x.style.display="flex",L.disabled=!1,H.disabled=!1):(x.style.display="none",L.disabled=!0,H.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let j=document.querySelectorAll(".row-checkbox");if(j.length>0){let q=[...j].every(Q=>Q.checked),R=[...j].some(Q=>Q.checked);O.checked=q,O.indeterminate=R&&!q}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{l.clear(),document.querySelectorAll(".row-checkbox").forEach(B=>B.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),c()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(l.size===0)return;let x=[...l],B=document.createElement("div");B.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",B.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${r}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(B),B.querySelector("#bulk-cancel-btn").addEventListener("click",()=>B.remove()),B.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let L=B.querySelector("#bulk-confirm-btn");L.disabled=!0,L.textContent="Menghapus...";let H=await M(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});B.remove(),H.ok?(K(`${x.length} ${r} berhasil dihapus.`),l.clear(),c(),be(a),_()):N(H.data?.error||"Gagal menghapus data.")})});let v=document.getElementById("filter-search"),k;v?.addEventListener("input",x=>{clearTimeout(k),k=setTimeout(()=>{m.search=x.target.value,d=1,l.clear(),c(),_()},400)}),g?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",B=>{m[x.name]=B.target.value;let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value=B.target.value),d=1,l.clear(),c(),_()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",B=>{m[x.name]=B.target.value;let L=document.getElementById(`filter-${x.name}`);L&&(L.value=B.target.value),d=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{m={},v&&(v.value=""),g?.forEach(x=>{let B=document.getElementById(`filter-${x.name}`);B&&(B.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),d=1,l.clear(),c(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{m={},v&&(v.value=""),g?.forEach(x=>{let B=document.getElementById(`filter-${x.name}`);B&&(B.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),d=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>z(null)),b&&document.addEventListener("click",function(x){let B=document.getElementById("aksi-menu-main"),L=document.getElementById("btn-aksi-main");B&&L&&!L.contains(x.target)&&!B.contains(x.target)&&B.classList.remove("show-aksi-menu")});let T=document.getElementById("btn-mobile-filter"),A=document.getElementById("filter-options-wrapper"),P=document.getElementById("btn-close-filter-sheet");if(T&&A&&(T.addEventListener("click",x=>{x.preventDefault(),A.classList.add("sheet-open")}),P&&P.addEventListener("click",x=>{x.preventDefault(),A.classList.remove("sheet-open")})),b){document.getElementById(`btn-export-${b.moduleName}`)?.addEventListener("click",async B=>{let L=B.target,H=L.innerHTML;L.innerHTML="\u23F3 Loading...",L.disabled=!0;try{await b.onExport()}catch{N("Gagal export data")}finally{L.innerHTML=H,L.disabled=!1}}),document.getElementById(`btn-template-${b.moduleName}`)?.addEventListener("click",()=>{b.onTemplate()});let x=document.getElementById(`input-import-${b.moduleName}`);x?.addEventListener("change",async B=>{let L=B.target.files[0];if(!L)return;let H=document.getElementById(`label-import-${b.moduleName}`),O=H?H.querySelector(".import-text"):null,j=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),H&&(H.style.pointerEvents="none"),x.disabled=!0;try{let q=await Ue(L);if(q.length===0)throw new Error("File kosong atau format salah");await b.onImport(q),K("Import berhasil!"),be(a),_()}catch(q){N(q.message||"Gagal import data")}finally{O&&(O.innerText=j),H&&(H.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){c();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let B=o==="client",L=B?1:d,H=B?Ae:20,O=new URLSearchParams({page:L,limit:H,...Object.fromEntries(Object.entries(m).filter(([,G])=>G))}),j=await M(`${a}?${O}`);if(!j.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${j.data?.error||"Error"}</p></div>`;return}let q=j.data?.data||j.data||[],R=j.data?.pagination,Q=q.length;if(B){q=C(q);let G=q.length,ae=20,X=Math.ceil(G/ae);d>X&&X>0&&(d=X);let U=(d-1)*ae,ue=d*ae;q=q.slice(U,ue),R={page:d,limit:ae,total:G,pages:X}}!1,E&&E(q);let ne=Xe({columns:p,data:q,onEdit:D?G=>z(G):null,actions:w.map(G=>({...G,handler:ae=>G.handler(ae,_)})),emptyText:`Tidak ada ${String(r||"").toLowerCase()}`,bulkSelect:n?{selectedIds:l,onToggle:c}:null});x.innerHTML="",x.appendChild(ne);let oe=document.getElementById("pagination-container");if(oe&&(oe.innerHTML="",R&&R.pages>1)){let G=Ze({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:ae=>{d=ae,_()}});G&&oe.appendChild(G)}}function I(x){let B=typeof h=="function"?h(x):h;return He(B)}function z(x){let B=!!x,L=document.createElement("form");if(L.noValidate=!0,L.innerHTML=I(x),B){let O=typeof h=="function"?h(x):h;tt(L,x)}let{close:H}=J({title:B?`Edit ${r}`:`Tambah ${r}`,content:L,size:"lg",confirmText:B?"Simpan Perubahan":`Tambah ${r}`,onConfirm:async(O,j)=>{if(!L.reportValidity())return;let q=O.querySelector(".modal-confirm");q.disabled=!0,q.textContent="Menyimpan...";let R=et(L),Q=typeof h=="function"?h(x):h,ne=async X=>{for(let U of X)if(U.type==="row")await ne(U.fields);else if(U.type==="combobox"&&R[U.name]){let ue=R[U.name],xe=(U.options||[]).find(de=>{let me=String(typeof de=="object"?de.value:de),dt=String(typeof de=="object"?de.label:de);return me===ue||dt===ue});if(xe)R[U.name]=typeof xe=="object"?xe.value:xe;else if(U.createApi){let de={};de[U.createApi.field]=ue,U.createApi.extra&&Object.assign(de,U.createApi.extra);let me=await M(U.createApi.path,{method:"POST",body:JSON.stringify(de)});if(me.ok&&me.data?.id)R[U.name]=me.data.id;else if(me.ok&&!me.data?.id)R[U.name]=ue;else throw new Error(`Gagal membuat master data: ${me.data?.error||"Unknown error"}`)}}};try{await ne(Q)}catch(X){N(X.message),q.disabled=!1,q.textContent=B?"Simpan Perubahan":`Tambah ${r}`;return}S&&(R=await S(R,x));let oe=B?"PUT":"POST",G=B?`${a}/${x.id}`:a,ae=await M(G,{method:oe,body:JSON.stringify(R)});ae.ok?(K(B?`${r} berhasil diperbarui.`:`${r} berhasil ditambahkan.`),j(),be(a),_()):(N(ae.data?.error||"Gagal menyimpan data."),q.disabled=!1,q.textContent=B?"Simpan Perubahan":`Tambah ${r}`)}})}function $(x){W(`Hapus ${r} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let B=await M(`${a}/${x.id}`,{method:"DELETE"});B.ok?(K(`${r} berhasil dihapus.`),be(a),_()):N(B.data?.error||"Gagal menghapus.")},`Hapus ${r}`)}return _(),_}V();ie();async function ca(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await ee(),a=await ge(),p=t?t.get("dash_filter"):null,h={};if(p==="reliever"){let S=new Date,E=`${S.getFullYear()}-${String(S.getMonth()+1).padStart(2,"0")}`;h={status:"Done",month:t&&t.get("month")?t.get("month"):E}}console.log("RAW",await _e()),console.log("OPTIONS",a);let g=S=>S&&!a.find(E=>E.value===S)?[...a,{value:S,label:S}]:a,u=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],r=S=>S&&!u.includes(S)?[...u,S]:u,f=new Date().getFullYear(),F=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((S,E)=>{let C=String(E+1).padStart(2,"0");return{value:`${f}-${C}`,label:`${S} ${f}`}});pa({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:h,onDataLoaded:S=>S.sort((E,C)=>{let w=E.backup_date?new Date(E.backup_date).getTime():0;return(C.backup_date?new Date(C.backup_date).getTime():0)-w}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:S=>ye(S)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:S=>window.formatDate(S)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:S=>window.formatDate(S)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:S=>S?`<span class="badge badge-info">${S}</span>`:"-"},{key:"status",label:"Status",render:S=>pe(S)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:u},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:F},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:S=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:S?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:S?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:g(S?.original_fc_name),value:S?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:r(S?.reliever_name),value:S?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:S?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:S?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:S?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:S?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:S?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let S=await M(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(S.ok){let E=S.data.data.map(C=>({Cabang:C.branch_name||"","Nama Facility care":C.original_fc_name||"",Periode:C.period||"",Relifer:C.reliever_name||"","Tanggal Back Up":C.backup_date||"","Tanggal Selesai":C.completion_date||"",Keterangan:C.reason||"",Shift:C.shift||"",Status:C.status||""}));E.length===0&&E.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),Y(E,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async S=>{let C=(await M("/api/branches?all=1")).data?.data||[],w=n=>{if(!n)return null;let o=String(n||"").toLowerCase(),s=C.find(d=>String(d.full_name||"").toLowerCase()===o||String(d.code||"").toLowerCase()===o||String(d.name||"").toLowerCase()===o);return s?s.id:null},y=S.map(n=>({branch_name:String(n.Cabang||"").trim(),backup_date:String(n["Tanggal Back Up"]||n["Tanggal Backup"]||"").trim(),original_fc_name:String(n["Nama Facility care"]||n["FC Digantikan"]||"").trim(),reliever_name:String(n.Relifer||n.Reliefer||"").trim(),period:String(n.Periode||"").trim(),reason:String(n.Keterangan||"").trim(),shift:String(n.Shift||"").trim(),completion_date:String(n["Tanggal Selesai"]||"").trim(),status:String(n.Status||"").trim()})).filter(n=>n.reliever_name&&n.backup_date),b=await M("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:y,onDuplicate:"update"})});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}V();ie();async function ua(e){let t=await ee(),i=Array.from({length:4},(a,p)=>String(new Date().getFullYear()-p));se({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ye(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>pe(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let p=new URLSearchParams(a||{}).toString(),h=await M(`/api/reports/inspection?limit=10000&${p}`);if(h.ok){let g=h.data.data.map(u=>({Cabang:u.branch_name||"",Periode:u.period||"",Tanggal:u.inspection_date||"","Point FC":u.fc_score!==null&&u.fc_score!==void 0?u.fc_score:"","Point SPV":u.spv_score!==null&&u.spv_score!==void 0?u.spv_score:"",Status:u.status||"","Link Dokumen":u.document_link||""}));Y(g,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let p=r=>{if(!r)return null;let f=String(r||"").toLowerCase(),D=t.find(F=>String(F.label||"").toLowerCase()===f);return D?D.value:null},h=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let f=String(r).trim();if(f===""||f==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);if(/^\d{4,5}$/.test(f)){let F=Number(f);if(F>2e4&&F<99999){let S=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let D=f.split(/[\/\-\.]/);if(D.length===3){let[F,S,E]=D.map(C=>C.trim());if(F.length===4&&S.length<=2&&E.length<=2)return`${F}-${S.padStart(2,"0")}-${E.padStart(2,"0")}`;if(E.length===4&&S.length<=2&&F.length<=2)return`${E}-${S.padStart(2,"0")}-${F.padStart(2,"0")}`}return f},g=a.map(r=>({branch_id:p(String(r.Cabang||"").trim()),period:String(r.Periode||"").trim(),inspection_date:h(r.Tanggal),fc_score:r["Point FC"]!==void 0&&r["Point FC"]!==""?Number(r["Point FC"]):null,spv_score:r["Point SPV"]!==void 0&&r["Point SPV"]!==""?Number(r["Point SPV"]):null,status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.period&&r.inspection_date),u=await M("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}V();ie();async function ga(e){let t=await ee(),i=Array.from({length:4},(a,p)=>String(new Date().getFullYear()-p));se({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ye(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>pe(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let p=new URLSearchParams(a||{}).toString(),h=await M(`/api/reports/cleaning?limit=10000&${p}`);if(h.ok){let g=h.data.data.map(u=>({Cabang:u.branch_name||"",Jenis:u.activity_type||"",Periode:u.period||"",Tanggal:u.activity_date||"",Status:u.status||"","Link Dokumen":u.document_link||""}));Y(g,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let p=r=>{if(!r)return null;let f=String(r||"").toLowerCase(),D=t.find(F=>String(F.label||"").toLowerCase()===f);return D?D.value:null},h=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let f=String(r).trim();if(f===""||f==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);if(/^\d{4,5}$/.test(f)){let F=Number(f);if(F>2e4&&F<99999){let S=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let D=f.split(/[\/\-\.]/);if(D.length===3){let[F,S,E]=D.map(C=>C.trim());if(F.length===4&&S.length<=2&&E.length<=2)return`${F}-${S.padStart(2,"0")}-${E.padStart(2,"0")}`;if(E.length===4&&S.length<=2&&F.length<=2)return`${E}-${S.padStart(2,"0")}-${F.padStart(2,"0")}`}return f},g=a.map(r=>({branch_id:p(String(r.Cabang||"").trim()),activity_type:String(r.Jenis||r.Kegiatan||"").trim(),period:String(r.Periode||"").trim(),activity_date:h(r.Tanggal),status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.activity_type&&r.period&&r.activity_date),u=await M("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}V();ie();async function ma(e,t){let i=await ee(),a=Array.from({length:4},(g,u)=>String(new Date().getFullYear()-u)),p=t?t.get("dash_filter"):null,h={};if(p==="fogging"){let g=new Date,u=String(g.getMonth()+1).padStart(2,"0"),r=String(g.getFullYear()),f=t?t.get("month"):null;f&&f.length===7&&(r=f.split("-")[0],u=f.split("-")[1]),h={status:"Done",month:u,year:r}}se({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:h,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:g=>`<span class="badge badge-warning">${g}</span>`},{key:"period",label:"Periode",render:g=>ye(g)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:g=>window.formatDate(g)},{key:"status",label:"Status",render:g=>pe(g)},{key:"document_link",label:"Dokumen",render:g=>g?`<a href="${g}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:g=>g||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:g=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:g?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:g?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:g?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:g?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:g?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:g?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async g=>{let u=new URLSearchParams(g||{}).toString(),r=await M(`/api/reports/fogging?limit=10000&${u}`);if(r.ok){let f=r.data.data.map(D=>({Cabang:D.branch_name||"",Jenis:D.activity_type||"Fogging",Periode:D.period||"",Tanggal:D.activity_date||"",Status:D.status||"","Link Dokumen":D.document_link||""}));Y(f,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async g=>{let u=F=>{if(!F)return null;let S=String(F||"").toLowerCase(),E=i.find(C=>String(C.label||"").toLowerCase()===S);return E?E.value:null},r=F=>{if(F==null||F==="")return"";if(F instanceof Date&&!isNaN(F.getTime()))return F.toISOString().slice(0,10);let S=String(F).trim();if(S===""||S==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);if(/^\d{4,5}$/.test(S)){let C=Number(S);if(C>2e4&&C<99999){let w=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let E=S.split(/[\/\-\.]/);if(E.length===3){let[C,w,y]=E.map(b=>b.trim());if(C.length===4&&w.length<=2&&y.length<=2)return`${C}-${w.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&w.length<=2&&C.length<=2)return`${y}-${w.padStart(2,"0")}-${C.padStart(2,"0")}`}return S},f=g.map(F=>({branch_id:u(String(F.Cabang||"").trim()),activity_type:String(F.Jenis||F.Kegiatan||"Fogging").trim(),period:String(F.Periode||"").trim(),activity_date:r(F.Tanggal),status:String(F.Status||"").trim(),document_link:String(F["Link Dokumen"]||"").trim(),notes:String(F.Catatan||F.Keterangan||"").trim()})).filter(F=>F.branch_id&&F.period&&F.activity_date),D=await M("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(f)});if(!D.ok)throw new Error(D.data?.error||"Import gagal");return D.data}}})}V();ie();async function ba(e){let t=await ee(),i=await ge(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],p=Array.from({length:4},(u,r)=>String(new Date().getFullYear()-r)),h=u=>u&&!i.find(r=>r.value===u)?[...i,{value:u,label:u}]:i,g=u=>u&&!a.find(r=>r.value===u)?[...a,{value:u,label:u}]:a;se({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:u=>`<span title="${u||""}">${u?.length>60?u.slice(0,60)+"\u2026":u||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:u=>window.formatDate(u)},{key:"status",label:"Status",render:u=>pe(u)},{key:"notes",label:"Keterangan",render:u=>u?.length>40?u.slice(0,40)+"\u2026":u||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:p}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:u?.branch_id},{name:"pic",label:"PIC",type:"select",options:g(u?.pic),value:u?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:u?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:u?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:u?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:u?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async u=>{let r=new URLSearchParams(u||{}).toString(),f=await M(`/api/reports/basecamp?limit=10000&${r}`);if(f.ok){let D=f.data.data.map(F=>({"Tgl Info":F.info_date||"",Cabang:F.branch_name||"",Permasalahan:F.problem||"",PIC:F.pic||"","Tgl Done":F.done_date||"",Status:F.status||"",Keterangan:F.notes||""}));Y(D,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async u=>{let r=S=>{if(!S)return null;let E=String(S||"").toLowerCase(),C=t.find(w=>String(w.label||"").toLowerCase()===E);return C?C.value:null},f=S=>{if(S==null||S==="")return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let E=String(S).trim();if(E===""||E==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(E))return E.slice(0,10);if(/^\d{4,5}$/.test(E)){let w=Number(E);if(w>2e4&&w<99999){let y=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let C=E.split(/[\/\-\.]/);if(C.length===3){let[w,y,b]=C.map(n=>n.trim());if(w.length===4&&y.length<=2&&b.length<=2)return`${w}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&y.length<=2&&w.length<=2)return`${b}-${y.padStart(2,"0")}-${w.padStart(2,"0")}`}return E},D=u.map(S=>({info_date:f(S["Tgl Info"]||S["Tanggal Info"]),branch_id:r(String(S.Cabang||"").trim()),problem:String(S.Permasalahan||"").trim(),pic:String(S.PIC||"").trim(),done_date:f(S["Tgl Done"]||S["Tanggal Done"]),status:String(S.Status||"").trim(),notes:String(S.Keterangan||S.Catatan||"").trim()})).filter(S=>S.info_date&&S.branch_id&&S.problem),F=await M("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(D)});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}}})}le();var _n=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],An=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],Bn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Bt="fcms_sop_fc_ik_v1",Dt="fcms_sop_fc_tools_v1",Tt="fcms_sop_fc_chem_v1";function ha(){try{let e=localStorage.getItem(Bt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(_n))}function $t(e){localStorage.setItem(Bt,JSON.stringify(e))}function fa(){try{let e=localStorage.getItem(Dt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(An))}function _t(e){localStorage.setItem(Dt,JSON.stringify(e))}function xa(){try{let e=localStorage.getItem(Tt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(Bn))}function At(e){localStorage.setItem(Tt,JSON.stringify(e))}async function ya(e){let t="ik",i="",a="all",p="tools",h="",g=ha(),u=fa(),r=xa();function f(){e.innerHTML=`
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
            <span>\u{1F4DA}</span> SOP Facility Care & Housekeeping
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
            Pedoman Standar Operasional Prosedur Kebersihan & Pemeliharaan Fasilitas PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <button id="btn-reset-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          <button id="btn-print-sop" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          <button id="btn-export-tools" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${D(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${F(t==="ik")}">${g.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${D(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${F(t==="tools")}">${u.length+r.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${D(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${D(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,S(),E()}function D(c){return c?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function F(c){return c?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function S(){e.querySelectorAll(".sop-tab-btn").forEach(T=>{T.addEventListener("click",()=>{t=T.dataset.tab,f()})});let c=e.querySelector("#btn-reset-default");c&&c.addEventListener("click",()=>{W("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Bt),localStorage.removeItem(Dt),localStorage.removeItem(Tt),g=ha(),u=fa(),r=xa(),K("Data SOP Facility Care berhasil di-reset ke standar asli"),f()},"Konfirmasi Reset Standar")});let v=e.querySelector("#btn-print-sop");v&&v.addEventListener("click",()=>{window.print()});let k=e.querySelector("#btn-export-tools");k&&k.addEventListener("click",()=>{m()})}function E(){let c=e.querySelector("#sop-content-area");c&&(t==="ik"?C(c):t==="tools"?b(c):t==="policy"?s(c):t==="links"&&d(c))}function C(c){c.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${a==="all"?"active":""}" data-cat="all">Semua (${g.length})</button>
              <button class="cat-pill ${a==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${g.filter(I=>I.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${a==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${g.filter(I=>I.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${a==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${g.filter(I=>I.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${a==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${g.filter(I=>I.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${a==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${g.filter(I=>I.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${g.length} dari ${g.length} Instruksi Kerja
            </div>
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let v=document.createElement("style");v.innerHTML=`
      .cat-pill {
        border: 1px solid var(--border);
        background: #F8FAFC;
        color: var(--text-2);
        font-size: 0.8rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .cat-pill:hover {
        background: #E2E8F0;
      }
      .cat-pill.active {
        background: #2563EB;
        color: #fff;
        border-color: #2563EB;
      }
      .ik-card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.08) !important;
        transform: translateY(-2px);
      }
      @media (max-width: 768px) {
        #cat-pills-wrap {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          scrollbar-width: none !important;
          padding-bottom: 4px !important;
          width: 100% !important;
        }
        #cat-pills-wrap::-webkit-scrollbar {
          display: none !important;
        }
        .sop-ik-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
      }
    `,c.appendChild(v);let k=c.querySelector("#ik-cards-grid-container"),T=c.querySelector("#ik-count-label");function A(){let I=g.filter(z=>{let $=a==="all"||z.category===a,x=i.toLowerCase().trim(),B=!x||z.title.toLowerCase().includes(x)||z.target.toLowerCase().includes(x)||z.tools.some(L=>L.toLowerCase().includes(x))||z.steps.some(L=>L.toLowerCase().includes(x));return $&&B});if(T&&(T.textContent=`Menampilkan ${I.length} dari ${g.length} Instruksi Kerja`),I.length===0){k.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}k.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${I.map(z=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${z.bg||"#EFF6FF"}; border-radius:10px;">${z.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${z.id}. ${z.title}
                      </h3>
                      <span class="badge" style="background:${z.bg||"#EFF6FF"}; color:${z.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${z.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${z.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${z.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${z.color||"#2563EB"};">
                  <strong>Area:</strong> ${z.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${z.tools.slice(0,3).map($=>`<li style="margin-bottom:2px;">${$}</li>`).join("")}
                    ${z.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${z.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${z.chemical_ratio&&z.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${z.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${z.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${z.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,k.querySelectorAll(".btn-view-ik").forEach(z=>{z.addEventListener("click",()=>{let $=parseInt(z.dataset.id,10),x=g.find(B=>B.id===$);x&&w(x)})}),k.querySelectorAll(".btn-edit-ik").forEach(z=>{z.addEventListener("click",()=>{let $=parseInt(z.dataset.id,10),x=g.find(B=>B.id===$);x&&y(x)})}),k.querySelectorAll(".btn-delete-ik").forEach(z=>{z.addEventListener("click",()=>{let $=parseInt(z.dataset.id,10),x=g.find(B=>B.id===$);x&&W(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${x.title}"?`,()=>{g=g.filter(B=>B.id!==$),$t(g),K("Instruksi Kerja berhasil dihapus"),f()},"Hapus Instruksi Kerja")})})}A();let P=c.querySelector("#ik-search");P&&P.addEventListener("input",I=>{i=I.target.value,A()}),c.querySelectorAll(".cat-pill").forEach(I=>{I.addEventListener("click",()=>{c.querySelectorAll(".cat-pill").forEach(z=>z.classList.remove("active")),I.classList.add("active"),a=I.dataset.cat,A()})});let _=c.querySelector("#btn-add-ik");_&&_.addEventListener("click",()=>{y()})}function w(c){J({title:`${c.icon||"\u{1F4CB}"} IK-${c.id}: ${c.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${c.bg||"#EFF6FF"}; color:${c.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${c.category}</span>
            <span>ID: IK-${String(c.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${c.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${c.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${c.tools.map(v=>`<li>${v}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${c.chemical_ratio&&c.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${c.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${c.steps.map(v=>`<li style="margin-bottom:6px;">${v}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${c.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${c.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(v,k)=>{k()}})}function y(c=null){let v=!!c,k=c?c.tools.join(`
`):"",T=c?c.steps.join(`
`):"",A=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${c?l(c.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${c?c.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${c?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${c?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${c?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${c?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${c?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${c?l(c.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${c?l(c.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${l(k)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${l(T)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${c?l(c.notes||""):""}</textarea>
        </div>
      </form>
    `;J({title:v?`\u270F\uFE0F Edit IK-${c.id}: ${c.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:A,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(P,_)=>{let I=P.querySelector("#form-ik-title").value.trim(),z=P.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",$=P.querySelector("#form-ik-cat").value,x=P.querySelector("#form-ik-ratio").value.trim(),B=P.querySelector("#form-ik-target").value.trim(),L=P.querySelector("#form-ik-tools").value.trim(),H=P.querySelector("#form-ik-steps").value.trim(),O=P.querySelector("#form-ik-notes").value.trim();if(!I||!B||!L||!H){N("Harap lengkapi seluruh field wajib (*)");return}let j=L.split(`
`).map(ne=>ne.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),q=H.split(`
`).map(ne=>ne.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),Q={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[$]||{color:"#2563EB",bg:"#EFF6FF"};if(v)c.title=I,c.icon=z,c.category=$,c.chemical_ratio=x,c.target=B,c.tools=j,c.steps=q,c.notes=O,c.color=Q.color,c.bg=Q.bg,$t(g),K(`Instruksi Kerja "${I}" berhasil diperbarui`);else{let ne=g.length>0?Math.max(...g.map(G=>G.id))+1:1,oe={id:ne,title:I,category:$,icon:z,color:Q.color,bg:Q.bg,target:B,chemical_ratio:x||"-",tools:j,steps:q,notes:O};g.push(oe),$t(g),K(`Instruksi Kerja baru IK-${ne} berhasil ditambahkan`)}_(),f()}})}function b(c){c.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${p==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${u.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${p==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${r.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${h}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${p==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let v=c.querySelector("#tool-table-container"),k=c.querySelector("#tool-search");function T(){let A=h.toLowerCase().trim();if(p==="tools"){let P=u.filter(_=>!A||_.name.toLowerCase().includes(A)||_.area.toLowerCase().includes(A));v.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${P.length} Item</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
                <thead>
                  <tr style="background:#F1F5F9; border-bottom:2px solid #E2E8F0; text-align:left;">
                    <th style="padding:10px 14px; width:60px;">No</th>
                    <th style="padding:10px 14px;">Nama Alat / Item</th>
                    <th style="padding:10px 14px; width:130px;">Kuantitas Standar</th>
                    <th style="padding:10px 14px;">Peruntukan Area & Kegunaan</th>
                    <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${P.map((_,I)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${I%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||I+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${l(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${l(_.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${l(_.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${_.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${_.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,v.querySelectorAll(".btn-edit-tool").forEach(_=>{_.addEventListener("click",()=>{let I=parseInt(_.dataset.no,10),z=u.find($=>$.no===I);z&&n(z)})}),v.querySelectorAll(".btn-delete-tool").forEach(_=>{_.addEventListener("click",()=>{let I=parseInt(_.dataset.no,10),z=u.find($=>$.no===I);z&&W(`Apakah Anda yakin ingin menghapus alat: "${z.name}"?`,()=>{u=u.filter($=>$.no!==I),u.forEach(($,x)=>{$.no=x+1}),_t(u),K("Alat kebersihan berhasil dihapus"),f()},"Hapus Alat Kebersihan")})})}else{let P=r.filter(_=>!A||_.name.toLowerCase().includes(A)||_.function.toLowerCase().includes(A));v.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${P.length} Chemical</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
                <thead>
                  <tr style="background:#F1F5F9; border-bottom:2px solid #E2E8F0; text-align:left;">
                    <th style="padding:10px 14px; width:60px;">No</th>
                    <th style="padding:10px 14px;">Nama Produk Chemical</th>
                    <th style="padding:10px 14px; width:110px;">Kemasan</th>
                    <th style="padding:10px 14px; width:220px;">Rasio & Standar Takaran</th>
                    <th style="padding:10px 14px;">Fungsi & Cara Penggunaan</th>
                    <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${P.map((_,I)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${I%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||I+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${l(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${l(_.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${l(_.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${l(_.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${_.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${_.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,v.querySelectorAll(".btn-edit-chem").forEach(_=>{_.addEventListener("click",()=>{let I=parseInt(_.dataset.no,10),z=r.find($=>$.no===I);z&&o(z)})}),v.querySelectorAll(".btn-delete-chem").forEach(_=>{_.addEventListener("click",()=>{let I=parseInt(_.dataset.no,10),z=r.find($=>$.no===I);z&&W(`Apakah Anda yakin ingin menghapus chemical: "${z.name}"?`,()=>{r=r.filter($=>$.no!==I),r.forEach(($,x)=>{$.no=x+1}),At(r),K("Bahan chemical berhasil dihapus"),f()},"Hapus Bahan Chemical")})})}}T(),c.querySelector("#view-tools-btn").addEventListener("click",()=>{p="tools",b(c)}),c.querySelector("#view-chem-btn").addEventListener("click",()=>{p="chemical",b(c)}),k&&k.addEventListener("input",A=>{h=A.target.value,T()}),c.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{p==="tools"?n():o()})}function n(c=null){let v=!!c,k=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${c?l(c.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${c?l(c.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${c?l(c.area):""}</textarea>
        </div>
      </form>
    `;J({title:v?`\u270F\uFE0F Edit Alat: ${c.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:k,confirmText:"Simpan",onConfirm:(T,A)=>{let P=T.querySelector("#form-tool-name").value.trim(),_=T.querySelector("#form-tool-qty").value.trim(),I=T.querySelector("#form-tool-area").value.trim();if(!P||!_||!I){N("Harap lengkapi semua kolom");return}if(v)c.name=P,c.qty=_,c.area=I,_t(u),K(`Alat "${P}" berhasil diperbarui`);else{let z=u.length>0?Math.max(...u.map($=>$.no))+1:1;u.push({no:z,name:P,qty:_,area:I}),_t(u),K(`Alat baru "${P}" berhasil ditambahkan`)}A(),f()}})}function o(c=null){let v=!!c,k=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${c?l(c.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${c?l(c.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${c?l(c.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${c?l(c.function):""}</textarea>
        </div>
      </form>
    `;J({title:v?`\u270F\uFE0F Edit Chemical: ${c.name}`:"\u2795 Tambah Chemical Baru",content:k,confirmText:"Simpan",onConfirm:(T,A)=>{let P=T.querySelector("#form-chem-name").value.trim(),_=T.querySelector("#form-chem-size").value.trim(),I=T.querySelector("#form-chem-ratio").value.trim(),z=T.querySelector("#form-chem-function").value.trim();if(!P||!_||!I||!z){N("Harap lengkapi semua kolom");return}if(v)c.name=P,c.size=_,c.ratio=I,c.function=z,At(r),K(`Chemical "${P}" berhasil diperbarui`);else{let $=r.length>0?Math.max(...r.map(x=>x.no))+1:1;r.push({no:$,name:P,size:_,ratio:I,function:z}),At(r),K(`Chemical baru "${P}" berhasil ditambahkan`)}A(),f()}})}function s(c){c.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(360px, 1fr)); gap:20px;">
        <!-- Card 1: Grooming -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#EFF6FF; padding:8px; border-radius:8px;">\u{1F454}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Standar Grooming Petugas (6.1)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">8 Persyaratan Wajib Penampilan Facility Care</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:var(--text-2); line-height:1.8;">
            <li>Menggunakan seragam resmi setiap hari (Senin s/d Minggu).</li>
            <li>Menggunakan celana bahan berwarna hitam (bukan jeans/chino kasual).</li>
            <li>Wajib mengenakan name badge atau ID Card di dada kiri.</li>
            <li>Menggunakan kaos kaki warna hitam.</li>
            <li>Menggunakan sepatu pantofel hitam yang bersih dan disemir.</li>
            <li>Menggunakan parfum dan deodorant untuk menjaga kesegaran badan.</li>
            <li>Merapikan rambut (pria pendek rapi, wanita diikat/hairnet rapi).</li>
            <li>Dilarang menggunakan aksesoris perhiasan selain jam tangan standar.</li>
          </ol>
        </div>

        <!-- Card 2: Jam Kerja -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF3C7; padding:8px; border-radius:8px;">\u23F0</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Pengaturan Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Sistem Kerja 6 Hari Kerja : 1 Hari Libur</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:10px; border-left:3px solid #2563EB;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u{1F3E5} Klinik Primer (1 Orang Facility Care):</strong>
              <span>Shift 1 (Normal): <strong>08.00 s/d 17.00</strong> (Istirahat 13.00 s/d 14.00)</span>
            </div>
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:12px; border-left:3px solid #059669;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u2728 Klinik Luxury (> 1 Orang Facility Care):</strong>
              <div style="margin-bottom:2px;">\u2022 Shift 1 (Pagi): <strong>07.00 s/d 16.00</strong> (Istirahat 12.00 s/d 13.00)</div>
              <div>\u2022 Shift 2 (Siang): <strong>13.00 s/d 22.00</strong> (Istirahat 17.00 s/d 18.00)</div>
            </div>
            <p style="margin:0; font-size:0.82rem; color:var(--text-3);">
              <strong>Ketentuan Backup:</strong> Saat Facility Care cabang libur, tim cabang wajib melakukan backup pekerjaan yang ditentukan oleh Supervisor (SPV) Cabang.
            </p>
          </div>
        </div>

        <!-- Card 3: Tanggung Jawab & Pembagian Area -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#ECFDF5; padding:8px; border-radius:8px;">\u{1F3E2}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Tanggung Jawab & Area (6.3)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Pembagian Pelaksana & Penanggung Jawab Area</span>
            </div>
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:14px;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; text-align:left;">
                <th style="padding:6px 10px;">Area Kerja</th>
                <th style="padding:6px 10px;">Pelaksana</th>
                <th style="padding:6px 10px;">PIC Pengawas</th>
              </tr>
            </thead>
            <tbody>
              <tr style="border-bottom:1px solid #F1F5F9;">
                <td style="padding:8px 10px; font-weight:600;">HQ & Gudang Pusat</td>
                <td style="padding:8px 10px;">Facility Care Pusat</td>
                <td style="padding:8px 10px; color:#2563EB; font-weight:600;">FC Supervisor</td>
              </tr>
              <tr>
                <td style="padding:8px 10px; font-weight:600;">Klinik / Cabang</td>
                <td style="padding:8px 10px;">Facility Care Cabang</td>
                <td style="padding:8px 10px; color:#059669; font-weight:600;">SPV Cabang</td>
              </tr>
            </tbody>
          </table>
          <div style="font-size:0.83rem; color:var(--text-2); line-height:1.6;">
            <div style="margin-bottom:6px;"><strong>Deep Cleaning:</strong> Pembersihan total eventual pasca renovasi, pembukaan cabang baru, atau relokasi cabang lama.</div>
            <div><strong>General Cleaning (GC):</strong> Pembersihan menyeluruh berkala untuk cabang eksisting sesuai jadwal yang telah ditetapkan.</div>
          </div>
        </div>

        <!-- Card 4: Monitoring & Evaluasi -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#F5F3FF; padding:8px; border-radius:8px;">\u{1F50D}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Sistem Monitoring & Inspeksi (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Mekanisme Evaluasi Kinerja Harian & Triwulanan</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="margin-bottom:10px;">
              <strong style="color:var(--text-1);">1. Monitor Harian & Mingguan:</strong>
              <div style="font-size:0.84rem; color:var(--text-2); margin-top:2px;">
                Ceklist harian diisi 2 kali sehari (Pagi & Siang) dan pekerjaan mingguan diisi seminggu sekali. Wajib ditandatangani oleh SPV Cabang / FC Supervisor.
              </div>
            </div>
            <div>
              <strong style="color:var(--text-1);">2. Inspeksi Cabang Triwulanan (3 Bulan):</strong>
              <div style="font-size:0.84rem; color:var(--text-2); margin-top:2px;">
                Dilakukan oleh FC Specialist langsung ke lokasi klinik. Laporan inspeksi diunggah ke Basecamp dan wajib diberikan tindak lanjut oleh PIC cabang maksimal <strong>3 hari</strong> setelah laporan terbit. Hasil penilaian mempengaruhi performance cabang.
              </div>
            </div>
          </div>
        </div>
      </div>
    `}function d(c){c.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Ceklist Harian Google Sheets -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4CA}</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Ceklist Harian & Mingguan</h3>
                <span class="badge" style="background:#E0F2FE; color:#0369A1; font-size:0.75rem; font-weight:700;">Google Sheets</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Format lembar ceklist kebersihan harian (Shift Pagi & Siang) dan ceklist pembersihan berkala mingguan yang wajib diisi oleh Facility Care dan disetujui Supervisor Cabang.
            </p>
          </div>
          <a href="https://docs.google.com/spreadsheets/d/1alq6ppT7syB8Q0_RjjeqwYhjeHgrWvVe-KJ3dHXWrss/edit?usp=sharing" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px;">
            \u{1F4C4} Buka Google Sheets Ceklist \u2794
          </a>
        </div>

        <!-- Card: Form Inspeksi Hygiene Google Forms -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4D1}</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Inspeksi Hygiene Cabang</h3>
                <span class="badge" style="background:#F3E8FF; color:#7E22CE; font-size:0.75rem; font-weight:700;">Google Forms</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi evaluasi kebersihan dan kepatuhan hygiene cabang yang diisi oleh Facility Care Specialist setiap periode 3 bulanan.
            </p>
          </div>
          <a href="https://docs.google.com/forms/d/e/1FAIpQLScn8iPabIt51E_GORkbD86IURSS-EsUYINEQ_Zwdq_uC3XOCQ/viewform" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#7E22CE; border-color:#7E22CE;">
            \u{1F4CB} Buka Google Form Inspeksi \u2794
          </a>
        </div>

        <!-- Card: Basecamp Monitoring Project -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F3D5}\uFE0F</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Laporan Inspeksi Basecamp</h3>
                <span class="badge" style="background:#FEF3C7; color:#92400E; font-size:0.75rem; font-weight:700;">Basecamp 3 Project</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              To-do list dan kanal pelaporan temuan inspeksi cabang. PIC Cabang wajib memberikan respon tindak lanjut maksimal 3 hari kerja setelah laporan terbit.
            </p>
          </div>
          <a href="https://3.basecamp.com/5259614/buckets/33617501/todosets/6367925585" target="_blank" rel="noopener" class="btn btn-outline" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#fff;">
            \u{1F680} Buka Laporan di Basecamp \u2794
          </a>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#fff; border:1px solid var(--border); border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
          <span>\u{1F4DC}</span> Validasi & Pengesahan Dokumen SOP
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3px solid #2563EB; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3px solid #059669; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.8rem; color:var(--text-2);">General Support Manager</div>
          </div>
          <div style="border-left:3px solid #7C3AED; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `}async function m(){try{let{downloadExcel:c}=await Promise.resolve().then(()=>(ie(),he)),v=g.map(T=>({"No IK":`IK-${T.id}`,"Judul Pekerjaan":T.title,Kategori:T.category,"Sasaran Target Area":T.target,"Rasio Chemical":T.chemical_ratio||"-","Alat & Bahan":T.tools.join(", "),"Langkah Kerja":T.steps.map((A,P)=>`${P+1}. ${A}`).join(" | "),"Catatan Safety":T.notes||"-"})),k=new Date().toISOString().slice(0,10);c(v,`SOP_Facility_Care_Instruksi_Kerja_${k}`),K("Data Excel Instruksi Kerja berhasil diunduh")}catch(c){console.error(c),N("Gagal mendownload data Excel: "+c.message)}}function l(c){return c?String(c).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}f()}le();var Dn=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],Tn=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],va={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},It="fcms_sop_sec_ik_v1",Mt="fcms_sop_sec_assets_v1";function ka(){try{let e=localStorage.getItem(It);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Dn))}function Pt(e){localStorage.setItem(It,JSON.stringify(e))}function wa(){try{let e=localStorage.getItem(Mt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(Tn))}function Lt(e){localStorage.setItem(Mt,JSON.stringify(e))}async function Sa(e){let t="ik",i="",a="all",p="",h="cabang",g=ka(),u=wa();function r(){e.innerHTML=`
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
            <span>\u{1F6E1}\uFE0F</span> SOP Pengelolaan Keamanan & Aset Perusahaan
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
            Standar Operasional Prosedur Pengawasan Keamanan Area, Fasilitas, dan Perlindungan Aset PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          <button id="btn-print-sec" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          <button id="btn-export-sec" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${f(t==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${D(t==="ik")}">${g.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="zonasi"?"active":""}" data-tab="zonasi" style="${f(t==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${t==="assets"?"active":""}" data-tab="assets" style="${f(t==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${D(t==="assets")}">${u.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${f(t==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${f(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,F(),S()}function f(l){return l?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function D(l){return l?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function F(){e.querySelectorAll(".sec-tab-btn").forEach(k=>{k.addEventListener("click",()=>{t=k.dataset.tab,r()})});let l=e.querySelector("#btn-reset-sec-default");l&&l.addEventListener("click",()=>{W("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(It),localStorage.removeItem(Mt),g=ka(),u=wa(),K("Data SOP Security berhasil di-reset ke standar dokumen"),r()},"Reset ke Standar Dokumen")});let c=e.querySelector("#btn-print-sec");c&&c.addEventListener("click",()=>{window.print()});let v=e.querySelector("#btn-export-sec");v&&v.addEventListener("click",()=>{d()})}function S(){let l=e.querySelector("#sec-content-area");l&&(t==="ik"?E(l):t==="zonasi"?y(l):t==="assets"?b(l):t==="policy"?o(l):t==="links"&&s(l))}function E(l){l.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${a==="all"?"active":""}" data-cat="all">Semua (${g.length})</button>
              <button class="sec-cat-pill ${a==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${g.filter(_=>_.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${a==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${g.filter(_=>_.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${a==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${g.filter(_=>_.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${g.length} Instruksi Kerja
            </div>
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let c=document.createElement("style");c.innerHTML=`
      .sec-cat-pill {
        border: 1px solid var(--border);
        background: #F8FAFC;
        color: var(--text-2);
        font-size: 0.8rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .sec-cat-pill:hover { background: #E2E8F0; }
      .sec-cat-pill.active {
        background: #1E40AF;
        color: #fff;
        border-color: #1E40AF;
      }
      .sec-ik-card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.08) !important;
        transform: translateY(-2px);
      }
      @media (max-width: 768px) {
        #sec-cat-pills-wrap {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          scrollbar-width: none !important;
          padding-bottom: 4px !important;
          width: 100% !important;
        }
        #sec-cat-pills-wrap::-webkit-scrollbar {
          display: none !important;
        }
        .sec-ik-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
      }
    `,l.appendChild(c);let v=l.querySelector("#sec-ik-cards-grid-container"),k=l.querySelector("#sec-ik-count-label");function T(){let _=g.filter(I=>{let z=a==="all"||I.category===a,$=i.toLowerCase().trim(),x=!$||I.title.toLowerCase().includes($)||I.target.toLowerCase().includes($)||I.tools.some(B=>B.toLowerCase().includes($))||I.steps.some(B=>B.toLowerCase().includes($));return z&&x});if(k&&(k.textContent=`Menampilkan ${_.length} dari ${g.length} Instruksi Kerja`),_.length===0){v.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}v.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${_.map(I=>`
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${I.bg||"#EFF6FF"}; border-radius:10px;">${I.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(I.id).padStart(2,"0")}. ${I.title}
                      </h3>
                      <span class="badge" style="background:${I.bg||"#EFF6FF"}; color:${I.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${I.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${I.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${I.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${I.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${I.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${I.tools.slice(0,3).map(z=>`<li style="margin-bottom:2px;">${m(z)}</li>`).join("")}
                    ${I.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${I.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${I.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${I.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,v.querySelectorAll(".btn-view-sec-ik").forEach(I=>{I.addEventListener("click",()=>{let z=parseInt(I.dataset.id,10),$=g.find(x=>x.id===z);$&&C($)})}),v.querySelectorAll(".btn-edit-sec-ik").forEach(I=>{I.addEventListener("click",()=>{let z=parseInt(I.dataset.id,10),$=g.find(x=>x.id===z);$&&w($)})}),v.querySelectorAll(".btn-delete-sec-ik").forEach(I=>{I.addEventListener("click",()=>{let z=parseInt(I.dataset.id,10),$=g.find(x=>x.id===z);$&&W(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${$.title}"?`,()=>{g=g.filter(x=>x.id!==z),Pt(g),K("Instruksi Kerja Security berhasil dihapus"),r()},"Hapus Instruksi Kerja")})})}T();let A=l.querySelector("#sec-ik-search");A&&A.addEventListener("input",_=>{i=_.target.value,T()}),l.querySelectorAll(".sec-cat-pill").forEach(_=>{_.addEventListener("click",()=>{l.querySelectorAll(".sec-cat-pill").forEach(I=>I.classList.remove("active")),_.classList.add("active"),a=_.dataset.cat,T()})});let P=l.querySelector("#btn-add-sec-ik");P&&P.addEventListener("click",()=>{w()})}function C(l){J({title:`${l.icon||"\u{1F6E1}\uFE0F"} IK-${String(l.id).padStart(2,"0")}: ${l.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${l.bg||"#EFF6FF"}; color:${l.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${l.category}</span>
            <span>Kode: IK-SEC-${String(l.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${l.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${l.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${l.tools.map(c=>{if(c.includes("http")){let v=c.split(": ");return`<li>${v[0]}: <a href="${v.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${m(c)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${l.steps.map(c=>`<li style="margin-bottom:6px;">${m(c)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${l.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${m(l.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(c,v)=>{v()}})}function w(l=null){let c=!!l,v=l?l.tools.join(`
`):"",k=l?l.steps.join(`
`):"",T=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${l?m(l.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${l?l.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${l?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${l?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${l?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${l?m(l.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${m(v)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${m(k)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${l?m(l.notes||""):""}</textarea>
        </div>
      </form>
    `;J({title:c?`\u270F\uFE0F Edit IK-${l.id}: ${l.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:T,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(A,P)=>{let _=A.querySelector("#form-sec-ik-title").value.trim(),I=A.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",z=A.querySelector("#form-sec-ik-cat").value,$=A.querySelector("#form-sec-ik-target").value.trim(),x=A.querySelector("#form-sec-ik-tools").value.trim(),B=A.querySelector("#form-sec-ik-steps").value.trim(),L=A.querySelector("#form-sec-ik-notes").value.trim();if(!_||!$||!x||!B){N("Harap lengkapi semua kolom bertanda bintang (*)");return}let H=x.split(`
`).map(R=>R.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),O=B.split(`
`).map(R=>R.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),q={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[z]||{color:"#1E40AF",bg:"#EFF6FF"};if(c)l.title=_,l.icon=I,l.category=z,l.target=$,l.tools=H,l.steps=O,l.notes=L,l.color=q.color,l.bg=q.bg,Pt(g),K(`Instruksi Kerja "${_}" berhasil diperbarui`);else{let R=g.length>0?Math.max(...g.map(ne=>ne.id))+1:1,Q={id:R,title:_,category:z,icon:I,color:q.color,bg:q.bg,target:$,tools:H,steps:O,notes:L};g.push(Q),Pt(g),K(`Instruksi Kerja baru IK-${R} berhasil ditambahkan`)}P(),r()}})}function y(l){l.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${h==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${h==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E2} Zonasi Head Quarter (5 Lantai)
          </button>
        </div>
        <div style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Ketentuan Pembagian Akses Ruangan (Lampiran 8.1 SOP)
        </div>
      </div>

      <!-- Definisi 3 Zona Header Cards -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px; margin-bottom:22px;">
        <div class="card" style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#DC2626; display:inline-block;"></span>
            <strong style="color:#991B1B; font-size:0.95rem;">ZONA MERAH (Restricted Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#7F1D1D; line-height:1.5;">
            Area sangat terbatas. Hanya personal yang bertugas spesifik di area tersebut dan BOD. Tamu WAJIB mendapat izin tertulis Manajemen dan didampingi penuh.
          </p>
        </div>

        <div class="card" style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#D97706; display:inline-block;"></span>
            <strong style="color:#92400E; font-size:0.95rem;">ZONA KUNING (Limited Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#78350F; line-height:1.5;">
            Area dengan pengamanan terbatas. Karyawan umum atau tamu yang memasuki area ini perlu mendapatkan arahan dan didampingi staf area terkait.
          </p>
        </div>

        <div class="card" style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#059669; display:inline-block;"></span>
            <strong style="color:#065F46; font-size:0.95rem;">ZONA HIJAU (Public Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#064E3B; line-height:1.5;">
            Area umum dengan pengawasan security. Diperbolehkan untuk seluruh tamu dan pasien setelah diarahkan oleh petugas security.
          </p>
        </div>
      </div>

      <!-- Zonasi Content Area -->
      <div id="zonasi-content-container"></div>
    `;let c=l.querySelector("#zonasi-content-container");h==="cabang"?c.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${va.cabang.map(v=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${v.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${v.color}; display:inline-block;"></span>
                  ${v.zone}
                </h3>
                <span class="badge" style="background:${v.bg}; color:${v.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${v.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${v.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${v.rooms.map(k=>`<li>${k}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:c.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${va.hq.map(v=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${v.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${v.hijau.map(k=>`<li>${k}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${v.kuning.map(k=>`<li>${k}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${v.merah.map(k=>`<li>${k}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,l.querySelector("#view-cabang-btn").addEventListener("click",()=>{h="cabang",y(l)}),l.querySelector("#view-hq-btn").addEventListener("click",()=>{h="hq",y(l)})}function b(l){l.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div>
          <h3 style="margin:0 0 4px 0; font-size:1.1rem; font-weight:700; color:var(--text-1);">
            Daftar Klasifikasi Aset yang Dipantau Security (Lampiran 8.3)
          </h3>
          <p style="margin:0; font-size:0.85rem; color:var(--text-3);">
            Pengawasan dan perlindungan sarana prasarana fisik, elektronik, dan peralatan klinik
          </p>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${p}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let c=l.querySelector("#sec-asset-table-container"),v=l.querySelector("#sec-asset-search");function k(){let A=p.toLowerCase().trim(),P=u.filter(_=>!A||_.item.toLowerCase().includes(A)||_.type.toLowerCase().includes(A)||_.description.toLowerCase().includes(A)||_.pic&&_.pic.toLowerCase().includes(A));c.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${P.length} Aset</span>
          </div>
          <div style="overflow-x:auto;">
            <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
              <thead>
                <tr style="background:#F1F5F9; border-bottom:2px solid #E2E8F0; text-align:left;">
                  <th style="padding:10px 14px; width:60px;">No</th>
                  <th style="padding:10px 14px; width:140px;">Tipe Aset</th>
                  <th style="padding:10px 14px; width:220px;">Nama Aset / Kelompok</th>
                  <th style="padding:10px 14px;">Deskripsi / Contoh Item</th>
                  <th style="padding:10px 14px; width:180px;">PIC Pengawas</th>
                  <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${P.map((_,I)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${I%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||I+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${T(_.type)}">${m(_.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${m(_.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${m(_.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${m(_.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${_.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${_.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,c.querySelectorAll(".btn-edit-asset").forEach(_=>{_.addEventListener("click",()=>{let I=parseInt(_.dataset.no,10),z=u.find($=>$.no===I);z&&n(z)})}),c.querySelectorAll(".btn-delete-asset").forEach(_=>{_.addEventListener("click",()=>{let I=parseInt(_.dataset.no,10),z=u.find($=>$.no===I);z&&W(`Apakah Anda yakin ingin menghapus aset: "${z.item}"?`,()=>{u=u.filter($=>$.no!==I),u.forEach(($,x)=>{$.no=x+1}),Lt(u),K("Aset berhasil dihapus"),r()},"Hapus Aset")})})}function T(A){return A==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":A==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}k(),v&&v.addEventListener("input",A=>{p=A.target.value,k()}),l.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{n()})}function n(l=null){let c=!!l,v=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${l?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${l?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${l?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${l?m(l.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${l?m(l.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${l?m(l.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;J({title:c?`\u270F\uFE0F Edit Aset: ${l.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:v,confirmText:"Simpan",onConfirm:(k,T)=>{let A=k.querySelector("#form-asset-type").value,P=k.querySelector("#form-asset-item").value.trim(),_=k.querySelector("#form-asset-desc").value.trim(),I=k.querySelector("#form-asset-pic").value.trim();if(!P||!_||!I){N("Harap lengkapi semua kolom");return}if(c)l.type=A,l.item=P,l.description=_,l.pic=I,Lt(u),K(`Aset "${P}" berhasil diperbarui`);else{let z=u.length>0?Math.max(...u.map($=>$.no))+1:1;u.push({no:z,type:A,item:P,description:_,pic:I}),Lt(u),K(`Aset baru "${P}" berhasil ditambahkan`)}T(),r()}})}function o(l){l.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(340px, 1fr)); gap:20px;">
        <!-- Card 1: Shift Kerja -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#EFF6FF; padding:8px; border-radius:8px;">\u23F0</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Sistem Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Ketentuan 12 Jam Kerja dengan 2 Hari Libur per Minggu</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:10px; border-left:3px solid #2563EB;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u2600\uFE0F Shift 1 (Pagi - Siang):</strong>
              <span>Pukul <strong>07.00 s/d 19.00 WIB</strong></span>
            </div>
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:12px; border-left:3px solid #1E293B;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u{1F319} Shift 2 (Malam):</strong>
              <span>Pukul <strong>19.00 s/d 07.00 WIB</strong></span>
            </div>
            <p style="margin:0; font-size:0.82rem; color:var(--text-3);">
              <strong>Ketentuan Off:</strong> Berlaku 2 hari off dalam setiap pekannya sesuai jadwal roster yang ditetapkan oleh Facility Care Supervisor / General Support.
            </p>
          </div>
        </div>

        <!-- Card 2: Jadwal Safety Patrol -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF3C7; padding:8px; border-radius:8px;">\u{1F6B6}\u200D\u2642\uFE0F</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Jadwal Safety Patrol Harian (6.2 B)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Keliling Pengecekan Area & Form Checklist Digital</span>
            </div>
          </div>
          <div style="display:flex; flex-direction:column; gap:8px; font-size:0.86rem;">
            <div style="background:#F8FAFC; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 1 (Pagi):</strong>
              <span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:700;">07.30 - 08.30 WIB</span>
            </div>
            <div style="background:#F8FAFC; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 2 (Malam):</strong>
              <span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:700;">19.30 - 20.30 WIB</span>
            </div>
            <div style="background:#F8FAFC; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 3 (Dini Hari):</strong>
              <span class="badge" style="background:#F3E8FF; color:#7E22CE; font-weight:700;">04.00 - 05.00 WIB</span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); margin-top:4px;">
              * <em>Patroli 3 khusus berlaku untuk area Head Quarter atau lokasi yang memiliki lebih dari 1 personil security.</em>
            </div>
          </div>
        </div>

        <!-- Card 3: Penanganan Pelanggaran Disiplin -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#ECFDF5; padding:8px; border-radius:8px;">\u26A0\uFE0F</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Penanganan Pelanggaran Disiplin (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Pelanggaran SK Tata Tertib Perusahaan</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.85rem; color:var(--text-2); line-height:1.7;">
            <li>Petugas security membuat <strong>Berita Acara Umum (BA)</strong> kejadian secara lengkap.</li>
            <li>Melaporkan kronologis tertulis kepada atasan (General Support Manager / FC Supervisor).</li>
            <li>Melaporkan dan meneruskan berkas kepada atasan langsung dari karyawan yang bersangkutan untuk penanganan HR/HC.</li>
          </ol>
        </div>

        <!-- Card 4: Penanganan Pelanggaran Kriminal -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF2F2; padding:8px; border-radius:8px;">\u{1F6A8}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:#DC2626;">Penanganan Pelanggaran Kriminal (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Tindak Pidana, Pencurian, & Pelanggaran Hukum</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.84rem; color:var(--text-2); line-height:1.7;">
            <li>Buat <strong>Berita Acara Umum</strong> secara resmi.</li>
            <li>Periksa dan amankan rekaman bukti visual melalui <strong>CCTV</strong>.</li>
            <li>Laporkan ke atasan dari karyawan terkait dengan rincian kronologis (bila pelaku karyawan).</li>
            <li>Minta dan catat detail identitas <strong>KTP / ID Card</strong> dari pihak yang bersangkutan.</li>
            <li>Amankan dan netralisir <strong>Tempat Kejadian Perkara (TKP)</strong> dari gangguan orang lain.</li>
            <li>Laporkan ke <strong>Kepolisian terdekat</strong> untuk proses hukum tindak lanjut.</li>
          </ol>
        </div>
      </div>
    `}function s(l){l.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Google Form Safety Patrol -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4D1}</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Safety Patrol Audy</h3>
                <span class="badge" style="background:#E0F2FE; color:#0369A1; font-size:0.75rem; font-weight:700;">Google Forms Resmi</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi keamanan yang wajib diisi oleh tim security saat melakukan Safety Patrol rutin di klinik dan kantor pusat Audy Dental.
            </p>
          </div>
          <a href="https://forms.gle/1EAvkqDJVdctXEGn7" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#1E40AF; border-color:#1E40AF;">
            \u{1F4CB} Buka Form Safety Patrol \u2794
          </a>
        </div>

        <!-- Card: Dasar Hukum & Ruang Lingkup Perusahaan -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u2696\uFE0F</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Dasar Hukum & Ruang Lingkup</h3>
                <span class="badge" style="background:#F1F5F9; color:var(--text-2); font-size:0.75rem; font-weight:700;">Legalitas</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:12px;">
              SOP ini diterapkan sebagai tata cara Pengelolaan Keamanan Aset Perusahaan di seluruh lokasi Audy Dental Klinik di bawah naungan:
            </p>
            <ul style="margin:0 0 0 16px; padding:0; font-size:0.84rem; color:var(--text-2); line-height:1.6;">
              <li><strong>PT. Audy Mandiri Indonesia</strong></li>
              <li><strong>PT. Kastara Ananta</strong></li>
              <li><strong>PT. Nara Adidaya</strong></li>
            </ul>
          </div>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#fff; border:1px solid var(--border); border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
          <span>\u{1F4DC}</span> Validasi & Pengesahan Dokumen SOP Security
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3px solid #2563EB; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3px solid #059669; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.8rem; color:var(--text-2);">General Support Manager</div>
          </div>
          <div style="border-left:3px solid #7C3AED; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `}async function d(){try{let{downloadExcel:l}=await Promise.resolve().then(()=>(ie(),he)),c=g.map(k=>({"Kode IK":`IK-SEC-${String(k.id).padStart(2,"0")}`,"Judul Prosedur":k.title,Kategori:k.category,"Sasaran Target":k.target,"Alat & Perlengkapan":k.tools.join(", "),"Langkah Pelaksanaan":k.steps.map((T,A)=>`${A+1}. ${T}`).join(" | "),"Catatan Khusus":k.notes||"-"})),v=new Date().toISOString().slice(0,10);l(c,`SOP_Security_Instruksi_Kerja_${v}`),K("Data Excel SOP Security berhasil diunduh")}catch(l){console.error(l),N("Gagal mendownload data Excel: "+l.message)}}function m(l){return l?String(l).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}r()}V();le();var Ea=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Ca(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Fa(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function $a(e){let t=[],i="all",a="",p="fcms-master-doc-styles";if(!document.getElementById(p)){let C=document.createElement("style");C.id=p,C.textContent=`
      .doc-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .doc-hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: #ffffff;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .doc-hero-card::after {
        content: '';
        position: absolute;
        right: -60px;
        top: -60px;
        width: 240px;
        height: 240px;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.18) 0%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
      }
      .doc-hero-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
      }
      .doc-hero-title {
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
        letter-spacing: -0.02em;
      }
      .doc-hero-subtitle {
        color: #94A3B8;
        font-size: 0.92rem;
        margin: 0;
        max-width: 720px;
        line-height: 1.5;
      }
      .doc-stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin-top: 24px;
      }
      .doc-stat-box {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .doc-stat-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        background: rgba(59, 130, 246, 0.18);
        color: #60A5FA;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        flex-shrink: 0;
      }
      .doc-stat-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.2;
      }
      .doc-stat-label {
        font-size: 0.78rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
      }
      .doc-controls-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .doc-controls-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 14px;
        flex-wrap: wrap;
      }
      .doc-search-wrapper {
        position: relative;
        flex: 1;
        min-width: 260px;
      }
      .doc-search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94A3B8;
        font-size: 0.95rem;
        pointer-events: none;
      }
      .doc-search-input {
        width: 100%;
        padding: 10px 14px 10px 40px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.9rem;
        background: #F8FAFC;
        transition: all 0.2s ease;
      }
      .doc-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #FFFFFF;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .doc-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .doc-pill {
        border: 1px solid #E2E8F0;
        background: #F8FAFC;
        color: #475569;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 30px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .doc-pill:hover {
        background: #EEF2F6;
        border-color: #CBD5E1;
      }
      .doc-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .doc-pill-count {
        background: rgba(0, 0, 0, 0.08);
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 0.75rem;
      }
      .doc-pill.active .doc-pill-count {
        background: rgba(255, 255, 255, 0.25);
        color: #FFFFFF;
      }
      /* Grid View */
      .doc-cards-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
        gap: 20px;
      }
      .doc-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
        overflow: hidden;
      }
      .doc-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.1);
        border-color: #CBD5E1;
      }
      .doc-card-top {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 12px;
        margin-bottom: 14px;
      }
      .doc-category-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        text-transform: uppercase;
      }
      .doc-platform-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        font-size: 0.74rem;
        font-weight: 600;
        color: #64748B;
        background: #F1F5F9;
        padding: 3px 8px;
        border-radius: 6px;
      }
      .doc-card-title {
        font-size: 1.08rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0 0 8px 0;
        line-height: 1.4;
      }
      .doc-card-desc {
        color: #64748B;
        font-size: 0.85rem;
        line-height: 1.5;
        margin: 0 0 18px 0;
        flex-grow: 1;
        display: -webkit-box;
        -webkit-line-clamp: 3;
        -webkit-box-orient: vertical;
        overflow: hidden;
      }
      .doc-card-footer {
        padding-top: 14px;
        border-top: 1px solid #F1F5F9;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
      }
      .doc-btn-open {
        background: #2563EB;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        padding: 8px 14px;
        font-size: 0.84rem;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        text-decoration: none;
        cursor: pointer;
        transition: background 0.15s ease;
      }
      .doc-btn-open:hover {
        background: #1D4ED8;
        color: #FFFFFF;
      }
      .doc-btn-open.disabled {
        background: #E2E8F0;
        color: #94A3B8;
        pointer-events: none;
      }
      .doc-card-actions {
        display: flex;
        align-items: center;
        gap: 4px;
      }
      .doc-icon-btn {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        color: #64748B;
        width: 34px;
        height: 34px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.15s ease;
      }
      .doc-icon-btn:hover {
        background: #EDF2F7;
        color: #1E293B;
        border-color: #CBD5E1;
      }
      .doc-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      /* Table View */
      .doc-table-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        overflow: hidden;
      }
      .doc-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
      }
      .doc-table th {
        background: #F8FAFC;
        padding: 14px 18px;
        font-size: 0.78rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        border-bottom: 1.5px solid #E2E8F0;
      }
      .doc-table td {
        padding: 16px 18px;
        font-size: 0.88rem;
        color: #1E293B;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: middle;
      }
      .doc-table tr:hover td {
        background: #F8FAFC;
      }
      .doc-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #FFFFFF;
        border-radius: 16px;
        border: 1px dashed #CBD5E1;
      }

      /* \u2500\u2500 TAMPILAN RESPONSIVE HP / MOBILE (\u2264 768px) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */
      @media (max-width: 768px) {
        .doc-page-container {
          padding: 12px 14px !important;
        }
        .doc-hero-card {
          padding: 18px 16px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
        }
        .doc-hero-title {
          font-size: 1.25rem !important;
          gap: 8px !important;
        }
        .doc-hero-subtitle {
          font-size: 0.82rem !important;
          line-height: 1.4 !important;
        }
        .doc-hero-header {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 12px !important;
        }
        .doc-hero-header > div:last-child {
          display: flex !important;
          width: 100% !important;
          gap: 8px !important;
        }
        #doc-btn-add {
          flex: 2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 12px !important;
        }
        #doc-btn-excel {
          flex: 1.2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 10px !important;
        }
        #doc-btn-refresh {
          flex: 0 0 38px !important;
          justify-content: center !important;
          padding: 9px 0 !important;
        }
        /* 2x2 compact stats grid di layar HP */
        .doc-stats-grid {
          grid-template-columns: repeat(2, 1fr) !important;
          gap: 8px !important;
          margin-top: 14px !important;
        }
        .doc-stat-box {
          padding: 10px 12px !important;
          gap: 10px !important;
          border-radius: 10px !important;
        }
        .doc-stat-icon {
          width: 36px !important;
          height: 36px !important;
          font-size: 1.1rem !important;
          border-radius: 8px !important;
        }
        .doc-stat-value {
          font-size: 1.15rem !important;
        }
        .doc-stat-label {
          font-size: 0.68rem !important;
        }
        .doc-controls-panel {
          padding: 14px 12px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
          gap: 12px !important;
        }
        .doc-controls-top {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 10px !important;
        }
        .doc-search-wrapper {
          min-width: 100% !important;
          width: 100% !important;
        }
        /* PITA FILTER SWIPEABLE HORIZONTAL DI HP (TIDAK LAGI NUMPUK VERTIKAL) */
        .doc-filter-pills {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          padding-bottom: 6px !important;
          padding-top: 2px !important;
          scrollbar-width: none !important;
          -ms-overflow-style: none !important;
        }
        .doc-filter-pills::-webkit-scrollbar {
          display: none !important;
        }
        .doc-pill {
          flex-shrink: 0 !important;
          white-space: nowrap !important;
          padding: 6px 12px !important;
          font-size: 0.8rem !important;
        }
        .doc-cards-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
        .doc-card {
          padding: 16px !important;
          border-radius: 14px !important;
        }
        .doc-card-title {
          font-size: 1rem !important;
        }
        .doc-card-desc {
          font-size: 0.82rem !important;
        }
        .doc-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .doc-table {
          min-width: 650px !important;
        }
        .doc-table th, .doc-table td {
          padding: 10px 12px !important;
          font-size: 0.82rem !important;
        }
      }
    `,document.head.appendChild(C)}async function h(){try{let C=await M("/api/checklist?limit=1000");if(C.ok&&C.data?.data&&Array.isArray(C.data.data)&&C.data.data.length>0)t=C.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t));else{let w=localStorage.getItem("fcms_master_doc_cache");w?t=JSON.parse(w):(t=[...Ea],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)))}}catch(C){console.warn("Fallback loading documents from cache/seed:",C);let w=localStorage.getItem("fcms_master_doc_cache");t=w?JSON.parse(w):[...Ea]}S()}function g(){return t.filter(C=>{let w=i==="all"||C.category===i,y=a.toLowerCase().trim(),b=!y||(C.name||"").toLowerCase().includes(y)||(C.category||"").toLowerCase().includes(y)||(C.description||"").toLowerCase().includes(y);return w&&b})}function u(){let C={all:t.length};return t.forEach(w=>{let y=w.category||"Lainnya";C[y]=(C[y]||0)+1}),C}function r(C=null){let w=!!C,y=Array.from(new Set(t.map(n=>n.category).filter(Boolean))),b=document.createElement("div");b.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${C?.name?String(C.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${C?.category?String(C.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${y.map(n=>`<option value="${n}">${n}</option>`).join("")}
            <option value="Berita Acara">Berita Acara</option>
            <option value="Apar & K3">Apar & K3</option>
            <option value="Kegiatan Deep Cleaning">Kegiatan Deep Cleaning</option>
            <option value="Kegiatan General Cleaning">Kegiatan General Cleaning</option>
            <option value="Drain Cleaning">Drain Cleaning</option>
            <option value="Form Operasional">Form Operasional</option>
          </datalist>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Tautan / Link Dokumen (Google Drive, Docs, Sheets, atau Forms)
          </label>
          <input type="url" id="doc-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://docs.google.com/..." value="${C?.document_link?String(C.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${C?.description||""}</textarea>
        </div>
      </div>
    `,J({title:w?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:b,size:"md",confirmText:w?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(n,o)=>{let s=b.querySelector("#doc-modal-name"),d=b.querySelector("#doc-modal-category"),m=b.querySelector("#doc-modal-link"),l=b.querySelector("#doc-modal-desc"),c=(s.value||"").trim(),v=(d.value||"").trim(),k=(m.value||"").trim(),T=(l.value||"").trim();if(!c){N("Nama Dokumen wajib diisi"),s.focus();return}if(!v){N("Kategori Dokumen wajib diisi"),d.focus();return}let A={name:c,category:v,document_link:k,description:T};try{if(w)if((await M(`/api/checklist/${C.id}`,{method:"PUT",body:JSON.stringify(A)})).ok){let _=t.findIndex(I=>I.id===C.id);_!==-1&&(t[_]={...t[_],...A}),K("Dokumen berhasil diperbarui")}else{let _=t.findIndex(I=>I.id===C.id);_!==-1&&(t[_]={...t[_],...A}),K("Dokumen disimpan secara lokal")}else{let P=await M("/api/checklist",{method:"POST",body:JSON.stringify(A)});if(P.ok&&P.data?.id)t.unshift({id:P.data.id,...A}),K("Dokumen baru berhasil ditambahkan");else{let _=Date.now();t.unshift({id:_,...A}),K("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),o(),S()}catch(P){console.error("Error saving document:",P),N("Gagal menyimpan dokumen: "+P.message)}}})}function f(C){W(`Apakah Anda yakin ingin menghapus dokumen <strong>"${C.name}"</strong>?`,async()=>{try{let w=await M(`/api/checklist/${C.id}`,{method:"DELETE"});w.ok||console.warn("API delete warning:",w),t=t.filter(y=>y.id!==C.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),K(`Dokumen "${C.name}" berhasil dihapus`),S()}catch{t=t.filter(y=>y.id!==C.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),K(`Dokumen "${C.name}" berhasil dihapus`),S()}},"Hapus Dokumen")}function D(C,w){if(!C){bt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(C).then(()=>{K(`Tautan "${w}" disalin ke clipboard!`)}).catch(()=>{bt(`Tautan: ${C}`)})}async function F(){try{let{downloadExcel:C}=await Promise.resolve().then(()=>(ie(),he)),w=t.map((y,b)=>({No:b+1,"Nama Dokumen":y.name||"",Kategori:y.category||"","Platform / Format":Fa(y.document_link).type,"Link Dokumen":y.document_link||"","Deskripsi / Keterangan":y.description||""}));C(w,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),K("Berhasil mengunduh data dokumen ke Excel.")}catch(C){N("Gagal mengekspor data: "+C.message)}}function S(){let C=g(),w=u(),y=Object.keys(w).filter(n=>n!=="all"),b=t.filter(n=>!!n.document_link).length;e.innerHTML=`
      <div class="doc-page-container">
        
        <!-- Hero Header -->
        <div class="doc-hero-card">
          <div class="doc-hero-header">
            <div>
              <h1 class="doc-hero-title">
                <span>\u{1F4CB}</span> Master Document & Template Standar
              </h1>
              <p class="doc-hero-subtitle">
                Pusat repositori formulir digital, checklist inspeksi kebersihan, berita acara, dan dokumen kerja standar PT Audy Mandiri Indonesia.
              </p>
            </div>
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; border:none; border-radius:10px; padding:10px 18px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>\u2795</span> Tambah Dokumen Baru
              </button>
              <button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 16px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>\u{1F4E5}</span> Ekspor Excel
              </button>
              <button id="doc-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 14px; font-weight:600;">
                <span>\u{1F504}</span>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="doc-stats-grid">
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F4C4}</div>
              <div>
                <div class="doc-stat-value">${t.length}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value">${y.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${b} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${t.length}</span></div>
                <div class="doc-stat-label">Tautan Online</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(168,85,247,0.18); color:#C084FC;">\u26A1</div>
              <div>
                <div class="doc-stat-value">Aktif</div>
                <div class="doc-stat-label">Status Repositori</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls: Search, View Switch, Category Pills -->
        <div class="doc-controls-panel">
          <div class="doc-controls-top">
            <div class="doc-search-wrapper" style="width:100%;">
              <span class="doc-search-icon">\u{1F50D}</span>
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${a?a.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${i==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${w.all||0}</span>
            </button>
            ${y.map(n=>{let o=Ca(n);return`
                <button class="doc-pill ${i===n?"active":""}" data-cat="${n}">
                  <span>${o.icon}</span>
                  <span>${n}</span>
                  <span class="doc-pill-count">${w[n]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${C.length===0?`
            <div class="doc-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Dokumen Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Coba sesuaikan kata kunci pencarian atau pilih kategori lain.</p>
              <button id="doc-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Tampilkan Semua Dokumen
              </button>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="doc-cards-grid">
              ${C.map(n=>{let o=Ca(n.category),s=Fa(n.document_link);return`
                  <div class="doc-card" data-id="${n.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${o.bg}; color:${o.color}; border:1px solid ${o.border};">
                          <span>${o.icon}</span>
                          <span>${o.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${s.icon}</span>
                          <span>${s.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${n.name}</h3>
                      <p class="doc-card-desc">${n.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${n.document_link?`
                        <a href="${n.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${n.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${n.document_link}" data-title="${n.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${n.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${n.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,E()}function E(){let C=e.querySelector("#doc-btn-add");C&&(C.onclick=()=>r());let w=e.querySelector("#doc-btn-excel");w&&(w.onclick=()=>F());let y=e.querySelector("#doc-btn-refresh");y&&(y.onclick=()=>h());let b=e.querySelector("#doc-btn-reset-filter");b&&(b.onclick=()=>{a="",i="all",S()});let n=e.querySelector("#doc-search-input");n&&(n.oninput=o=>{a=o.target.value;let s=g();S();let d=e.querySelector("#doc-search-input");d&&(d.focus(),d.setSelectionRange(d.value.length,d.value.length))}),e.querySelectorAll(".doc-pill").forEach(o=>{o.onclick=()=>{i=o.getAttribute("data-cat")||"all",S()}}),e.querySelectorAll(".copy-link-btn").forEach(o=>{o.onclick=s=>{s.stopPropagation();let d=o.getAttribute("data-link"),m=o.getAttribute("data-title");D(d,m)}}),e.querySelectorAll(".edit-doc-btn").forEach(o=>{o.onclick=s=>{s.stopPropagation();let d=Number(o.getAttribute("data-id")),m=t.find(l=>l.id===d);m&&r(m)}}),e.querySelectorAll(".delete-doc-btn").forEach(o=>{o.onclick=s=>{s.stopPropagation();let d=Number(o.getAttribute("data-id")),m=t.find(l=>l.id===d);m&&f(m)}})}await h()}V();le();ie();var Me={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function zt(e){for(let[t,i]of Object.entries(Me))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function _a(e){e.innerHTML=`
    <!-- Header Section -->
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>\u2728</span> Master Standar Kebersihan Fisik
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Acuan Standar Kebersihan Fisik Seluruh Area & Ruangan Klinik (Audy Dental \xB7 Facility Care)
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-table-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:6px 12px; cursor:pointer;">
            \u{1F4CB} Tabel Standar
          </button>
          <button id="view-zone-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            \u{1F5C2}\uFE0F Per Zona Area
          </button>
        </div>
        <button id="btn-print-hygiene" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;" title="Cetak Format SOP Standar Kebersihan">
          \u{1F5A8}\uFE0F Cetak SOP
        </button>
        <button id="btn-export-hygiene" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        <button id="btn-add-hygiene" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Parameter
        </button>
      </div>
    </div>

    <!-- KPI Summary Cards -->
    <div id="hygiene-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#EFF6FF; color:#2563EB; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F4CB}
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Total Parameter</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-params">104</div>
          <div style="font-size:0.75rem; color:#16A34A; font-weight:500;">\u2713 Terverifikasi Standar SOP</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#F5F3FF; color:#7C3AED; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F3E2}
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Kategori Zona</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-zones">8</div>
          <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">Klinis, Publik, Sanitair, dll.</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#ECFDF5; color:#059669; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F6AA}
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Ruangan / Lokasi</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-rooms">35</div>
          <div style="font-size:0.75rem; color:#059669; font-weight:500;">Termasuk Ruang Tindakan</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#FFFBEB; color:#D97706; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F6E1}\uFE0F
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Kondisi Fisik</div>
          <div style="font-size:1.6rem; font-weight:800; color:#16A34A;">100%</div>
          <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">Bebas Debu, Kerak & Bau</div>
        </div>
      </div>
    </div>

    <!-- Zona Switcher Tabs -->
    <div style="overflow-x:auto; padding-bottom:8px; margin-bottom:18px;">
      <div id="hygiene-zone-tabs" style="display:flex; gap:8px; min-width:max-content; border-bottom:2px solid #E2E8F0; padding-bottom:8px;">
        <button class="zone-tab-btn active" data-zone="all" style="border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
          \u{1F31F} Semua Zona <span class="badge" style="background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-all">104</span>
        </button>
        ${Object.entries(Me).map(([o,s])=>`
          <button class="zone-tab-btn" data-zone="${o}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${s.icon}</span> ${o}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${o.replace(/[\s&]+/g,"-")}">-</span>
          </button>
        `).join("")}
      </div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:220px; max-width:360px;">
            <input type="text" id="hygiene-search" class="form-control" placeholder="Cari ruangan, item objek, atau standar kebersihan..." style="padding-left:36px; height:40px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:12px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.95rem;">\u{1F50D}</span>
          </div>
          <select id="hygiene-room-filter" class="form-control" style="width:auto; height:40px; border-radius:8px; font-size:0.875rem; min-width:180px;">
            <option value="">Semua Ruangan</option>
          </select>
          <button id="btn-reset-filters" class="btn btn-sm btn-ghost" style="color:var(--text-3); font-weight:600; cursor:pointer; padding:6px 10px;">
            \u2715 Reset
          </button>
        </div>
        <div id="hygiene-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="hygiene-view-container">
      <div class="skeleton" style="height:320px; border-radius:12px;"></div>
    </div>
  `;let t="all",i="table",a=[],p=Z(),h=p&&["superadmin","admin","manager","spv"].includes(p.role);async function g(){try{let o=await M("/api/hygiene-standards?all=1");if(!o.ok)throw new Error("Gagal memuat standar hygiene");a=Array.isArray(o.data)?o.data:o.data?.data||[],r(),u(),D()}catch(o){console.error(o),N(o.message||"Terjadi kesalahan memuat data");let s=e.querySelector("#hygiene-view-container");s&&(s.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${o.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,s.querySelector("#btn-retry-hygiene")?.addEventListener("click",g))}}function u(){let o=e.querySelector("#badge-count-all");o&&(o.textContent=a.length),Object.entries(Me).forEach(([m,l])=>{let c=m.replace(/[\s&]+/g,"-"),v=e.querySelector(`#badge-count-${c}`);if(v){let k=a.filter(T=>l.rooms.includes(T.room_name)).length;v.textContent=k}});let s=new Set(a.map(m=>m.room_name)),d=e.querySelector("#card-total-rooms");d&&(d.textContent=s.size)}function r(){let o=e.querySelector("#hygiene-room-filter");if(!o)return;let s=o.value,d=[];if(t==="all")d=Array.from(new Set(a.map(m=>m.room_name))).filter(Boolean).sort();else{let m=Me[t];m&&(d=m.rooms.filter(l=>a.some(c=>c.room_name===l)))}o.innerHTML=`<option value="">Semua Ruangan (${d.length})</option>`+d.map(m=>`<option value="${m}" ${m===s?"selected":""}>${m}</option>`).join("")}function f(){let o=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),s=e.querySelector("#hygiene-room-filter")?.value||"";return a.filter(d=>{if(t!=="all"){let m=Me[t];if(!m||!m.rooms.includes(d.room_name))return!1}if(s&&d.room_name!==s)return!1;if(o){let m=(d.room_name||"").toLowerCase(),l=(d.item_name||"").toLowerCase(),c=(d.cleanliness_standard||"").toLowerCase();if(!m.includes(o)&&!l.includes(o)&&!c.includes(o))return!1}return!0})}function D(){let o=f(),s=e.querySelector("#hygiene-stats-text");s&&(s.innerHTML=`Menampilkan <strong>${o.length}</strong> dari <strong>${a.length}</strong> parameter`);let d=e.querySelector("#hygiene-view-container");if(d){if(o.length===0){d.innerHTML=`
        <div class="empty-state" style="text-align:center; padding:50px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
          <h3 style="margin:0 0 8px 0; color:var(--text-1); font-weight:700;">Tidak Ada Parameter yang Cocok</h3>
          <p style="color:var(--text-3); max-width:400px; margin:0 auto 16px auto; font-size:0.9rem;">
            Tidak ditemukan parameter kebersihan dengan kata kunci atau filter yang Anda pilih.
          </p>
          <button id="btn-clear-search-empty" class="btn btn-outline" style="border-radius:8px; padding:8px 16px; font-weight:600;">
            Reset Filter
          </button>
        </div>
      `,d.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let m=e.querySelector("#hygiene-search"),l=e.querySelector("#hygiene-room-filter");m&&(m.value=""),l&&(l.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(c=>c.classList.toggle("active",c.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(c=>{c.dataset.zone==="all"?(c.style.background="#2563EB",c.style.color="#fff"):(c.style.background="#F1F5F9",c.style.color="var(--text-2)")}),r(),D()});return}i==="table"?F(d,o):S(d,o)}}function F(o,s){o.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${h?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${s.map((d,m)=>{let l=zt(d.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${m+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${l.bg}; color:${l.color}; border:1px solid ${l.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${l.icon}</span> ${d.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${l.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${d.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${d.cleanliness_standard}
                      </div>
                    </td>
                    ${h?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${d.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${d.id}" data-name="${d.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
                            \u{1F5D1}\uFE0F
                          </button>
                        </div>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,E(o)}function S(o,s){let d={};s.forEach(m=>{let l=zt(m.room_name);d[l.zoneName]||(d[l.zoneName]={info:l,rooms:{}}),d[l.zoneName].rooms[m.room_name]||(d[l.zoneName].rooms[m.room_name]=[]),d[l.zoneName].rooms[m.room_name].push(m)}),o.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(d).map(([m,l])=>{let c=Object.values(l.rooms).reduce((v,k)=>v+k.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${l.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${l.info.bg}; border-bottom:1px solid ${l.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${l.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${l.info.color};">
                      ${m}
                    </h3>
                    <div style="font-size:0.78rem; color:${l.info.color}; opacity:0.85;">
                      ${Object.keys(l.rooms).length} Ruangan \xB7 ${c} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${l.info.color}; border:1px solid ${l.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${c} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(l.rooms).map(([v,k])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${v}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${k.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${h?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${k.map((T,A)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${A+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${T.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${T.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${h?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${T.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${T.id}" data-name="${T.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                                  </div>
                                </td>
                              `:""}
                            </tr>
                          `).join("")}
                        </tbody>
                      </table>
                    </div>
                  </div>
                `).join("")}
              </div>
            </div>
          `}).join("")}
      </div>
    `,E(o)}function E(o){o.querySelectorAll(".btn-edit-item").forEach(s=>{s.addEventListener("click",()=>{let d=s.dataset.id,m=a.find(l=>String(l.id)===String(d));m&&C(m)})}),o.querySelectorAll(".btn-delete-item").forEach(s=>{s.addEventListener("click",()=>{let d=s.dataset.id,m=s.dataset.name;W(`Yakin ingin menghapus parameter standar <strong>"${m}"</strong>?`,async()=>{try{let l=await M(`/api/hygiene-standards/${d}`,{method:"DELETE"});if(!l.ok)throw new Error(l.data?.error||"Gagal menghapus parameter");K("Parameter kebersihan berhasil dihapus"),g()}catch(l){N(l.message)}},"Hapus Parameter Standar")})})}function C(o=null){let s=!!o,d=Array.from(new Set(a.map(l=>l.room_name))).filter(Boolean).sort(),m=document.createElement("div");m.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${o?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${d.map(l=>`<option value="${l}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${o?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${o?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,J({title:s?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:m,confirmText:s?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(l,c)=>{let v=m.querySelector("#modal-room-name").value.trim(),k=m.querySelector("#modal-item-name").value.trim(),T=m.querySelector("#modal-standard").value.trim();if(!v||!k||!T){N("Semua kolom wajib diisi");return}try{let A={room_name:v,item_name:k,cleanliness_standard:T},P=s?`/api/hygiene-standards/${o.id}`:"/api/hygiene-standards",I=await M(P,{method:s?"PUT":"POST",body:JSON.stringify(A)});if(!I.ok)throw new Error(I.data?.error||"Gagal menyimpan parameter");K(s?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),c(),g()}catch(A){N(A.message)}}})}function w(){let o=window.open("","_blank");if(!o){N("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let d=`
      <!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="UTF-8">
        <title>SOP Master Standar Kebersihan Fisik - Audy Dental</title>
        <style>
          @page { size: A4 portrait; margin: 12mm 15mm 15mm 15mm; }
          body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; color: #1e293b; line-height: 1.4; margin: 0; padding: 0; }
          .header { border-bottom: 2px solid #1e3a8a; padding-bottom: 8px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: flex-end; }
          .logo { font-size: 16pt; font-weight: 800; color: #1e3a8a; letter-spacing: -0.5px; }
          .sublogo { font-size: 8.5pt; color: #64748b; font-weight: 600; text-transform: uppercase; }
          .doc-info { text-align: right; font-size: 8pt; color: #475569; }
          .title { text-align: center; margin: 12px 0 16px 0; }
          .title h1 { font-size: 13pt; margin: 0 0 4px 0; color: #0f172a; text-transform: uppercase; }
          .title p { margin: 0; font-size: 8.5pt; color: #64748b; }
          .zone-title { background: #f1f5f9; border-left: 4px solid #2563eb; padding: 6px 10px; font-size: 9.5pt; font-weight: 700; color: #1e3a8a; margin: 14px 0 8px 0; page-break-after: avoid; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 8.5pt; page-break-inside: auto; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          th { background: #f8fafc; border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; font-weight: 700; color: #334155; font-size: 8pt; text-transform: uppercase; }
          td { border: 1px solid #cbd5e1; padding: 6px 8px; vertical-align: top; }
          .td-no { text-align: center; width: 28px; }
          .td-room { width: 120px; font-weight: 600; }
          .td-obj { width: 150px; font-weight: 600; color: #0f172a; }
          .footer-sign { margin-top: 24px; display: flex; justify-content: space-between; page-break-inside: avoid; }
          .sign-box { text-align: center; width: 180px; font-size: 8.5pt; }
          .sign-line { margin-top: 45px; border-top: 1px solid #0f172a; font-weight: 600; padding-top: 4px; }
          @media print {
            .no-print { display: none !important; }
          }
        </style>
      </head>
      <body>
        <div class="no-print" style="background:#fef3c7; border:1px solid #f59e0b; padding:10px 14px; border-radius:6px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center;">
          <span style="font-size:9pt; color:#92400e;">\u{1F4A1} <strong>Format Cetak SOP Standar Kebersihan Fisik FCMS</strong>. Siap dicetak untuk display janitor atau arsip operasional.</span>
          <button onclick="window.print()" style="background:#2563eb; color:#fff; border:none; padding:6px 14px; border-radius:4px; font-weight:bold; cursor:pointer;">Cetak Sekarang</button>
        </div>

        <div class="header">
          <div>
            <div class="logo">AUDY DENTAL</div>
            <div class="sublogo">Facility Care Management System (FCMS)</div>
          </div>
          <div class="doc-info">
            <div>Dokumen: <strong>FCMS-SOP-HYG-01</strong></div>
            <div>Tanggal Cetak: ${new Date().toLocaleDateString("id-ID",{day:"numeric",month:"long",year:"numeric"})}</div>
            <div>Total: <strong>${a.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(Me).map(([m,l])=>{let c=a.filter(v=>l.rooms.includes(v.room_name));return c.length===0?"":`
            <div class="zone-title">${l.icon} ${m.toUpperCase()} (${c.length} PARAMETER)</div>
            <table>
              <thead>
                <tr>
                  <th class="td-no">No</th>
                  <th class="td-room">Ruangan / Lokasi</th>
                  <th class="td-obj">Item / Objek Inspeksi</th>
                  <th>Standar Kebersihan Fisik</th>
                </tr>
              </thead>
              <tbody>
                ${c.map((v,k)=>`
                  <tr>
                    <td class="td-no">${k+1}</td>
                    <td class="td-room">${v.room_name}</td>
                    <td class="td-obj">${v.item_name}</td>
                    <td>${v.cleanliness_standard}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          `}).join("")}

        <div class="footer-sign">
          <div class="sign-box">
            <div>Dibuat Oleh,</div>
            <div class="sign-line">Facility Care Specialist</div>
          </div>
          <div class="sign-box">
            <div>Ditinjau Oleh,</div>
            <div class="sign-line">Supervisor Operasional</div>
          </div>
          <div class="sign-box">
            <div>Disetujui Oleh,</div>
            <div class="sign-line">Head of Operations</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() { window.print(); }, 400);
          };
        <\/script>
      </body>
      </html>
    `;o.document.open(),o.document.write(d),o.document.close()}function y(){let o=a.map((d,m)=>{let l=zt(d.room_name);return{No:m+1,"Kategori Zona":l.zoneName,"Ruangan / Lokasi":d.room_name,"Item / Objek Inspeksi":d.item_name,"Standar Kebersihan Fisik":d.cleanliness_standard}}),s=new Date().toISOString().slice(0,10);Y(o,`Master_Standar_Kebersihan_FCMS_${s}`),K("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>D()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>D()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let o=e.querySelector("#hygiene-search"),s=e.querySelector("#hygiene-room-filter");o&&(o.value=""),s&&(s.value=""),D()}),e.querySelectorAll(".zone-tab-btn").forEach(o=>{o.addEventListener("click",()=>{t=o.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(s=>{let d=s.dataset.zone===t;if(s.classList.toggle("active",d),d){s.style.background="#2563EB",s.style.color="#fff";let m=s.querySelector(".badge");m&&(m.style.background="rgba(255,255,255,0.25)",m.style.color="#fff")}else{s.style.background="#F1F5F9",s.style.color="var(--text-2)";let m=s.querySelector(".badge");m&&(m.style.background="#CBD5E1",m.style.color="#334155")}}),r(),D()})});let b=e.querySelector("#view-table-btn"),n=e.querySelector("#view-zone-btn");b?.addEventListener("click",()=>{i="table",b.style.background="#fff",b.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",n.style.background="transparent",n.style.boxShadow="none",D()}),n?.addEventListener("click",()=>{i="zone",n.style.background="#fff",n.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",b.style.background="transparent",b.style.boxShadow="none",D()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",w),e.querySelector("#btn-export-hygiene")?.addEventListener("click",y),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>C(null)),g()}V();le();ie();async function Aa(e,t="supply"){return Pn(e)}async function Pn(e){let i=((await M("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));se({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,p)=>p.branch_name_ref||p.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let p=JSON.parse(a);return Array.isArray(p)?p.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let p=JSON.parse(a);return Array.isArray(p)?p.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>pe(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let p=a?.tools_items;try{p=Array.isArray(JSON.parse(p))?JSON.parse(p).join(", "):p}catch{}let h=a?.chemical_items;try{h=Array.isArray(JSON.parse(h))?JSON.parse(h).join(", "):h}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(g=>g.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:p},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:h},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let p=new URLSearchParams(a||{}).toString(),h=await M(`/api/reports/supply?limit=10000&${p}`);if(h.ok){let g=h.data.data.map(u=>{let r=u.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let f=u.chemical_items;try{f=Array.isArray(JSON.parse(f))?JSON.parse(f).join(", "):f}catch{}return{Waktu:u.submitted_at||"",Pengirim:u.submitter_name||"",Cabang:u.branch_name_ref||u.branch_name||"","Alat/Barang":r||"",Chemical:f||"",Catatan:u.additional_notes||"",Status:u.status||"","Diproses Oleh":u.processed_by||""}});Y(g,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let h=(await M("/api/branches?all=1")).data?.data||[],g=D=>{if(!D)return null;let F=String(D||"").toLowerCase(),S=h.find(E=>String(E.full_name||"").toLowerCase()===F||String(E.code||"").toLowerCase()===F||String(E.name||"").toLowerCase()===F);return S?S.id:null},u=D=>{if(D==null||D==="")return"";if(D instanceof Date&&!isNaN(D.getTime()))return D.toISOString().slice(0,10);let F=String(D).trim();if(F===""||F==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(F))return F.slice(0,10);if(/^\d{4,5}$/.test(F)){let E=Number(F);if(E>2e4&&E<99999){let C=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let S=F.split(/[\/\-\.]/);if(S.length===3){let[E,C,w]=S.map(y=>y.trim());if(E.length===4&&C.length<=2&&w.length<=2)return`${E}-${C.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&C.length<=2&&E.length<=2)return`${w}-${C.padStart(2,"0")}-${E.padStart(2,"0")}`}return F},r=a.map(D=>({submitted_at:u(D.Waktu||D.Tanggal),submitter_name:String(D.Pengirim||"").trim(),branch_id:g(String(D.Cabang||"").trim()),tools_items:String(D["Alat/Barang"]||D.Alat||"").trim(),chemical_items:String(D.Chemical||"").trim(),additional_notes:String(D.Catatan||D.Keterangan||"").trim(),status:String(D.Status||"").trim(),processed_by:String(D["Diproses Oleh"]||D.PIC||"").trim()})).filter(D=>D.submitted_at&&D.submitter_name&&D.branch_id),f=await M("/api/reports/supply/import",{method:"POST",body:JSON.stringify(r)});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,p)=>{let h=J({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${a.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${a.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${a.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${a.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(g,u)=>{let r=g.querySelector("#supply-status").value,f=g.querySelector("#supply-processed-by").value;(await M(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:r,processed_by:f})})).ok?(K("Status diperbarui."),u(),p()):N("Gagal update status.")}})}}]})}V();ie();async function Ba(e){let t=Z();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}se({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await M(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(p=>({"Nama Lengkap":p.full_name||"",Username:p.username||"",Email:p.email||"",Role:p.role||"",Status:p.is_active?"Aktif":"Nonaktif"}));Y(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(h=>({full_name:String(h["Nama Lengkap"]||"").trim(),username:String(h.Username||"").trim(),email:String(h.Email||"").trim(),role:String(h.Role||"").trim()||"viewer",password:String(h.Password||"").trim()})).filter(h=>h.username&&h.password&&h.email&&h.full_name),p=await M("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}V();ie();async function Da(e){se({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await M(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)Y(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{Y([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(p=>({code:String(p["Kode Cabang"]||"").trim(),name:String(p["Nama Pendek"]||"").trim(),full_name:String(p["Nama Lengkap"]||"").trim(),city:String(p.Kota||"").trim()})).filter(p=>p.code&&p.name),a=await M("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}V();async function Ta(e){let t=new Date,i=[];e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F4C5} Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">\u2039 Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next \u203A</button>
        <div class="calendar-filters">
          <label class="filter-check"><input type="checkbox" value="schedule"        checked class="cal-filter"> Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="issue"           checked class="cal-filter"> Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="reliever"        checked class="cal-filter"> Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="training"        checked class="cal-filter"> Training</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> Kontrak Habis</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid" style="min-height:400px"></div>
      </div>
    </div>
    <!-- Event detail popup -->
    <div id="cal-event-list" class="cal-event-sidebar" style="display:none">
      <div class="cal-event-header">
        <span id="cal-event-date"></span>
        <button class="btn btn-ghost btn-sm" id="cal-event-close">&times;</button>
      </div>
      <div id="cal-event-items"></div>
    </div>
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),p()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),p()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(g=>g.addEventListener("change",p));async function a(){try{let g=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await M(`/api/dashboard/calendar?month=${g}`)).data?.data||[]}catch(g){console.warn("[Calendar] Failed to load events, rendering empty grid:",g),i=[]}}async function p(){let g=document.getElementById("calendar-grid");if(g){g.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let u=t.getFullYear(),r=t.getMonth(),f=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),D=document.getElementById("cal-month-label");D&&(D.textContent=f);let F=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(d=>d.value)),S=i.filter(d=>F.has(d.type)),E={};S.forEach(d=>{let m=Ln(d.event_date);m&&(E[m]||(E[m]=[]),E[m].push(d))});let C=new Date(u,r,1).getDay(),w=new Date(u,r+1,0).getDate(),y=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],b=new Date().toISOString().slice(0,10),n='<div class="calendar-grid">';y.forEach(d=>{n+=`<div class="cal-day-header">${d}</div>`});for(let d=0;d<C;d++)n+='<div class="cal-cell cal-cell-empty"></div>';for(let d=1;d<=w;d++){let m=`${u}-${String(r+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`,l=E[m]||[],c=m===b;n+=`
          <div class="cal-cell ${c?"cal-today":""} ${l.length?"cal-has-events":""}"
               data-date="${m}" tabindex="0" role="button" aria-label="${m}">
            <div class="cal-day-num ${c?"today-num":""}">${d}</div>
            <div class="cal-events-preview">
              ${l.slice(0,3).map(v=>`
                <div class="cal-event-dot cal-color-${v.color||"gray"}" title="${Se(v.title||v.type)}">
                  <span class="cal-event-dot-label">${In(v.title||v.branch_name||v.type,18)}</span>
                </div>
              `).join("")}
              ${l.length>3?`<div class="cal-more">+${l.length-3} lagi</div>`:""}
            </div>
          </div>`}let s=(C+w)%7;if(s!==0)for(let d=0;d<7-s;d++)n+='<div class="cal-cell cal-cell-empty"></div>';n+="</div>",g.innerHTML=n,g.querySelectorAll(".cal-cell[data-date]").forEach(d=>{d.addEventListener("click",()=>{let m=d.dataset.date,l=E[m]||[];if(!l.length)return;let c=document.getElementById("cal-event-list"),v=new Date(m+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=v,document.getElementById("cal-event-items").innerHTML=l.map(k=>`
            <div class="cal-event-item cal-color-border-${k.color||"gray"}">
              <div class="cal-event-type">
                ${Mn(k.type)}
                ${k.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${Se(k.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${Se(k.title||"-")}</div>
              ${k.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${Se(k.branch_name)}</div>`:""}
              ${k.pic||k.period||k.opening_date||k.target_date||k.completion_date||k.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${k.pic?`<div>\u{1F464} <b>PIC:</b> ${Se(k.pic)}</div>`:""}
                  ${k.period?`<div>\u{1F4C5} <b>Periode:</b> ${Se(k.period)}</div>`:""}
                  ${k.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${Se(jt(k.opening_date))}</div>`:""}
                  ${k.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${Se(jt(k.target_date))}</div>`:""}
                  ${k.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${Se(jt(k.completion_date))}</div>`:""}
                  ${k.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${Se(k.notes)}</div>`:""}
                </div>
              `:""}
              ${k.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${k.status==="Done"?"badge-success":"badge-warning"}">${Se(k.status)}</span></div>`:""}
              ${k.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${k.days_remaining} hari</div>`:""}
            </div>
          `).join(""),c.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),g&&(g.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let h=()=>{a().then(()=>p())};at.on("data:changed",h),p()}function Ln(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[a,p,h]=i.map(g=>g.trim());if(h.length===4)return`${h}-${p.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${p.padStart(2,"0")}-${h.padStart(2,"0")}`}return t.slice(0,10)}function jt(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[a,p,h]=i.map(g=>g.trim());if(a.length===4)return`${h.padStart(2,"0")}-${p.padStart(2,"0")}-${a}`;if(h.length===4)return`${a.padStart(2,"0")}-${p.padStart(2,"0")}-${h}`}return e}function In(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function Se(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Mn(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}V();async function Pa(e){let t=Z(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),p={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${p},${p}99)">
            ${i}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${p}18;color:${p};margin-top:6px">
              ${(t?.role||"viewer").toUpperCase()}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">\u{1F4E7} Email</span>
            <span class="info-value">${t?.email||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F464} Username</span>
            <span class="info-value">${t?.username||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F3AF} Role</span>
            <span class="info-value" style="color:${p};font-weight:700">${t?.role||"\u2014"}</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: Change Password -->
      <div class="chart-card">
        ${t?.role!=="viewer"?`
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">\u{1F511} Ganti Password</div>
            <div class="chart-card-subtitle">Gunakan password yang kuat, minimal 6 karakter</div>
          </div>
        </div>

        <form id="change-pwd-form" novalidate style="margin-top:8px">
          <div class="form-group">
            <label class="form-label">Password Lama <span class="required">*</span></label>
            <input type="password" name="current_password" class="form-control"
              required placeholder="Masukkan password saat ini" autocomplete="current-password">
          </div>
          <div class="form-group">
            <label class="form-label">Password Baru <span class="required">*</span></label>
            <input type="password" name="new_password" class="form-control"
              required placeholder="Minimal 6 karakter" autocomplete="new-password">
          </div>
          <div class="form-group">
            <label class="form-label">Konfirmasi Password Baru <span class="required">*</span></label>
            <input type="password" name="confirm_password" class="form-control"
              required placeholder="Ulangi password baru" autocomplete="new-password">
          </div>

          <div id="pwd-error" class="alert alert-danger" style="display:none"></div>
          <div id="pwd-success" class="alert alert-success" style="display:none"></div>

          <button type="submit" class="btn btn-primary" id="btn-save-pwd">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px">
              <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
              <polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
            </svg>
            Simpan Password
          </button>
        </form>
        `:`
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">\u{1F511} Ganti Password</div>
            <div class="chart-card-subtitle" style="color: #ef4444; font-weight:500;">Perubahan password dinonaktifkan untuk Viewer. Silakan hubungi Administrator.</div>
          </div>
        </div>
        `}

        <hr class="profile-divider" style="margin-top:28px">

        <div class="chart-card-title" style="margin-bottom:12px">\u{1F510} Keamanan Akun</div>
        <div class="info-list">
          <div class="info-row">
            <span class="info-key">Token Login</span>
            <span class="info-value">
              <span class="badge badge-success">Aktif</span>
            </span>
          </div>
          <div class="info-row">
            <span class="info-key">Session</span>
            <span class="info-value" id="session-info">Memuat...</span>
          </div>
        </div>
        <button class="btn btn-danger btn-sm" id="btn-logout" style="margin-top:16px">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:4px">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
            <polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/>
          </svg>
          Keluar dari Semua Sesi
        </button>
      </div>

    </div>
  `;let h=localStorage.getItem("fm_token"),g=document.getElementById("session-info");if(h&&g)try{let u=JSON.parse(atob(h.split(".")[1])),r=new Date(u.exp*1e3);g.textContent=`Berakhir: ${r.toLocaleString("id-ID")}`}catch{g.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async u=>{u.preventDefault();let r=document.getElementById("pwd-error"),f=document.getElementById("pwd-success"),D=document.getElementById("btn-save-pwd");r.style.display="none",f.style.display="none";let F=u.target,S=F.current_password.value,E=F.new_password.value,C=F.confirm_password.value;if(E!==C){r.textContent="\u274C Konfirmasi password tidak cocok.",r.style.display="block";return}if(E.length<6){r.textContent="\u274C Password baru minimal 6 karakter.",r.style.display="block";return}D.disabled=!0,D.textContent="\u23F3 Menyimpan...";let w=await M("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:S,new_password:E})});D.disabled=!1,D.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',w.ok?(f.textContent="\u2705 Password berhasil diubah.",f.style.display="block",F.reset(),K("Password berhasil diubah.")):(r.textContent=w.data?.error||"Gagal mengubah password.",r.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}V();var st={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ce(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let p=Number(t);if(p>2e4&&p<99999){let h=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(h.getTime())?null:h.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[p,h,g]=i.map(D=>D.trim()),u=Number(p),r=Number(h),f=Number(g);if(p.length===4&&u>1900)return`${p}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&f>1900)return u>12?`${g}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`:r>12?`${g}-${p.padStart(2,"0")}-${h.padStart(2,"0")}`:`${g}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`;if(g.length===2&&!isNaN(f)){let D=f>=50?`19${g}`:`20${g}`;return u>12?`${D}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`:`${D}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`}}let a=new Date(t);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function La(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var zn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:ce(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:ce(e["Target Pindah OS"]),target_selesai_os:ce(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:ce(e["Tanggal Mulai"]),end_date:ce(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:ce(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:ce(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:ce(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:ce(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:ce(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:ce(e["Tanggal Target"]||e["Tgl Target"]),completion_date:ce(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:ce(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:ce(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:ce(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:ce(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:ce(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:ce(e["Tanggal Back Up"]),completion_date:ce(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:ce(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:ce(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function jn(e,t){let i=st[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=zn[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let p=[],h=[],g=[];return t.filter(r=>!La(r)).forEach((r,f)=>{let D=t.indexOf(r)+2,F=[];a.required.forEach(({key:E,label:C})=>{let w=r[E];if(w==null||String(w).trim()===""){let y=Object.keys(r).filter(b=>b.trim()).join(", ");F.push({column:C,originalValue:w||"",reason:`Kolom "${C}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${y.slice(0,120)}`})}});let S=a.map(r);F.length>0?h.push({row:D,data:S,raw:r,errors:F}):(p.push(r),g.push(S))}),{valid:p,errors:h,mapped:g}}function Ia(e){let t=[];return e.SheetNames.forEach(i=>{let a=st[i];if(!a)return;let p=e.Sheets[i],h=window.XLSX.utils.sheet_to_json(p,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),g=jn(i,h),u=h.filter(r=>!La(r));t.push({sheetName:i,module:a.module,label:a.label,total:u.length,valid:g.mapped.length,errorCount:g.errors.length,errors:g.errors,mapped:g.mapped,skipped:!1})}),t}function Ma(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,p])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(p),a)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function za(e){let t=window.XLSX,i=t.utils.book_new(),a=!1;return e.forEach(p=>{if(!p.errors||p.errors.length===0)return;a=!0;let h=p.errors.map(u=>({"No. Baris":u.row,"Kolom Gagal":(u.errors||[]).map(r=>r.column||r).join("; "),"Alasan Error":(u.errors||[]).map(r=>r.reason||r).join("; "),...Object.fromEntries(Object.entries(u.data||{}).map(([r,f])=>[r,f??""]))})),g=t.utils.json_to_sheet(h);t.utils.book_append_sheet(i,g,p.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Nn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function ja(e){e.innerHTML=`
    <div class="page-header">
      <div class="header-content">
        <h1 class="page-title"><span class="title-icon">\u{1F4E5}</span> Import Data Awal</h1>
        <p class="page-subtitle">Unggah file Excel untuk mengisi data aplikasi, atau sinkronkan langsung dari Google Sheets.</p>
      </div>
      <div class="page-actions" style="display:flex;gap:8px">
        <button id="btn-sync-google" class="btn btn-secondary">
          <span>\u{1F504} Tarik Data dari Google Sheets</span>
        </button>
        <button class="btn btn-warning" id="btn-backup-db">\u{1F4E6} Backup Database</button>
        <button class="btn btn-secondary" id="btn-download-template">\u2B07\uFE0F Download Template</button>
      </div>
    </div>

    <!-- STEP 1: Upload -->
    <div id="step-upload" class="import-step">
      <div class="card">
        <div class="card-body">
          <div class="import-info-box">
            <h3>\u{1F4CB} Petunjuk Import Data Awal</h3>
            <p>Upload file Excel (.xlsx) yang sudah diisi sesuai template. Sistem akan membaca seluruh sheet secara otomatis dan memvalidasi sebelum data disimpan.</p>
            <div class="import-sheet-list">
              ${Object.entries(st).map(([w,{label:y}])=>`<span class="import-sheet-tag">\u{1F4C4} ${w} \u2192 ${y}</span>`).join("")}
            </div>
          </div>

          <div class="import-upload-zone" id="upload-zone">
            <div class="upload-icon">\u{1F4C2}</div>
            <div class="upload-text">
              <strong>Drag & Drop file Excel di sini</strong>
              <span>atau klik untuk memilih file</span>
            </div>
            <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none">
            <button class="btn btn-primary" id="btn-browse">Pilih File Excel</button>
            <div class="upload-hint">Format: .xlsx | Ukuran maks: 20MB</div>
          </div>
          
          <div id="file-info" style="display:none" class="file-info-bar">
            <span id="file-name-display"></span>
            <button class="btn btn-ghost btn-sm" id="btn-clear-file">\u2715 Ganti</button>
          </div>
        </div>
      </div>
    </div>

    <!-- STEP 2: Validating (progress) -->
    <div id="step-validating" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body text-center">
          <div class="import-progress-wrap">
            <div class="spinner" style="margin:0 auto 16px"></div>
            <div id="validation-status" class="import-status-text">Membaca file Excel...</div>
            <div class="import-progress-bar"><div class="import-progress-fill" id="validation-bar" style="width:0%"></div></div>
          </div>
        </div>
      </div>
    </div>

    <!-- STEP 3: Preview -->
    <div id="step-preview" class="import-step" style="display:none">
      <!-- Duplicate Strategy -->
      <div class="card mb-12">
        <div class="card-body">
          <h3 style="margin-bottom:12px">\u2699\uFE0F Pengaturan Duplikat</h3>
          <div class="dup-options">
            <label class="dup-option">
              <input type="radio" name="dup-strategy" value="skip" checked>
              <div class="dup-option-text">
                <strong>Lewati Data Duplikat</strong>
                <span>Data yang sudah ada di database tidak akan diubah</span>
              </div>
            </label>
            <label class="dup-option">
              <input type="radio" name="dup-strategy" value="update">
              <div class="dup-option-text">
                <strong>Perbarui Data yang Sudah Ada</strong>
                <span>Data lama akan ditimpa dengan data dari Excel</span>
              </div>
            </label>
          </div>
        </div>
      </div>

      <!-- Preview Table -->
      <div class="card mb-12">
        <div class="card-body p-0">
          <div class="preview-header">
            <h3>\u{1F4CA} Preview Validasi per Sheet</h3>
            <div id="preview-summary-badges"></div>
          </div>
          <div id="preview-table-container"></div>
        </div>
      </div>

      <!-- Error Detail -->
      <div id="error-detail-section" style="display:none" class="card mb-12">
        <div class="card-body p-0">
          <div class="preview-header">
            <h3>\u274C Detail Error</h3>
            <button class="btn btn-secondary btn-sm" id="btn-download-log">\u2B07\uFE0F Download Log Error</button>
          </div>
          <div id="error-detail-container"></div>
        </div>
      </div>

      <!-- Actions -->
      <div class="import-action-bar">
        <button class="btn btn-ghost" id="btn-back-to-upload">\u2190 Upload Ulang</button>
        <button class="btn btn-primary" id="btn-start-import" disabled>
          \u{1F680} Mulai Import
        </button>
      </div>
    </div>

    <!-- STEP 4: Importing -->
    <div id="step-importing" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body">
          <h3 style="margin-bottom:20px;text-align:center">\u23F3 Sedang Mengimport Data...</h3>
          <div id="import-steps-list" class="import-steps-list"></div>
          <div class="import-progress-bar" style="margin-top:20px">
            <div class="import-progress-fill" id="import-bar" style="width:0%"></div>
          </div>
          <div id="import-current-status" class="import-status-text" style="margin-top:8px;text-align:center"></div>
        </div>
      </div>
    </div>

    <!-- STEP 5: Summary -->
    <div id="step-summary" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body">
          <div class="import-summary-header" id="summary-status-icon"></div>
          <div class="import-summary-stats" id="summary-stats"></div>
          <div id="summary-module-results"></div>
          <div class="import-action-bar" style="margin-top:24px">
            <button class="btn btn-secondary" id="btn-import-again">\u{1F504} Import Lagi</button>
            <button class="btn btn-primary" id="btn-go-to-dashboard">\u{1F4CA} Ke Dashboard</button>
          </div>
        </div>
      </div>
    </div>
  `;let t=null,i=null,a=0,p={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function h(w){Object.entries(p).forEach(([y,b])=>{b.style.display=y===w?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let w=document.getElementById("btn-backup-db");w.disabled=!0,w.textContent="\u23F3 Memproses Backup...";try{let y=await M("/api/import/backup");if(y.ok){if(!window.XLSX){N("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let b=window.XLSX,n=b.utils.book_new();Object.entries(y.data.database).forEach(([o,s])=>{let d=s.length>0?s:[{}],m=b.utils.json_to_sheet(d);b.utils.book_append_sheet(n,m,o.substring(0,31))}),b.writeFile(n,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),K("Backup berhasil diunduh!")}else N("Gagal memproses backup: "+(y.data?.error||"Unknown error"))}catch(y){N("Gagal memproses backup: "+y.message)}finally{w.disabled=!1,w.textContent="\u{1F4E6} Backup Database"}});let g=document.getElementById("btn-sync-google");g&&g.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let w=g.innerHTML;g.innerHTML='<span class="spinner"></span> Menyinkronkan...',g.disabled=!0;try{let y=await M("/api/sync/google-sheets",{method:"POST"});y.ok?alert("Sinkronisasi Berhasil: "+(y.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(y.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{g.innerHTML=w,g.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Ma(),K("Template Excel berhasil didownload!")});let u=document.getElementById("file-input"),r=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",w=>{w.stopPropagation(),u.click()}),u.addEventListener("change",w=>{w.target.files[0]&&f(w.target.files[0])}),r.addEventListener("dragover",w=>{w.preventDefault(),r.classList.add("drag-over")}),r.addEventListener("dragleave",()=>r.classList.remove("drag-over")),r.addEventListener("drop",w=>{w.preventDefault(),r.classList.remove("drag-over");let y=w.dataTransfer.files[0];y&&y.name.match(/\.xlsx?$/i)?f(y):N("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,u.value="",document.getElementById("file-info").style.display="none",r.style.display="",h("upload")});async function f(w){t=w,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${w.name} (${(w.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",r.style.display="none",await D(w)}async function D(w){h("validating");let y=document.getElementById("validation-status"),b=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");y.textContent="Membaca file Excel...",b.style.width="20%",await Ye(200);let n=await w.arrayBuffer(),o=window.XLSX.read(n,{type:"array",cellDates:!0});y.textContent=`Memvalidasi ${o.SheetNames.length} sheet...`,b.style.width="50%",await Ye(100),i=Ia(o),b.style.width="100%",y.textContent="Validasi selesai!",await Ye(300),F()}catch(n){h("upload"),N("Gagal memproses file: "+n.message),document.getElementById("file-info").style.display="flex",r.style.display="none"}}function F(){h("preview");let w=i.filter(c=>!c.skipped).length,y=i.reduce((c,v)=>c+v.total,0),b=i.reduce((c,v)=>c+v.valid,0),n=i.reduce((c,v)=>c+v.errorCount,0),o=y>0?Math.round(b/y*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${w} sheet</span>
      <span class="badge badge-secondary">${y} baris</span>
      <span class="badge badge-success">${b} valid (${o}%)</span>
      ${n>0?`<span class="badge badge-danger">${n} error</span>`:""}
    `;let s=document.getElementById("preview-table-container");s.innerHTML=`
      <table class="data-table">
        <thead>
          <tr>
            <th>Sheet (Excel)</th>
            <th>Modul</th>
            <th style="text-align:center">Total</th>
            <th style="text-align:center">Valid</th>
            <th style="text-align:center">Error</th>
            <th style="text-align:center">Status</th>
            <th style="text-align:center">Detail</th>
          </tr>
        </thead>
        <tbody>
          ${i.map((c,v)=>`
            <tr class="${c.errorCount>0?"row-error":c.skipped?"row-skipped":"row-ok"}">
              <td><strong>${c.sheetName}</strong></td>
              <td>${c.label}</td>
              <td style="text-align:center">${c.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${c.valid}</span></td>
              <td style="text-align:center">${c.errorCount>0?`<span class="badge badge-danger">${c.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${c.skipped?'<span class="badge badge-neutral">Dilewati</span>':c.errorCount>0&&c.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':c.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':c.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${c.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${v}">\u{1F50D} ${c.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,s.querySelectorAll(".btn-detail-error").forEach(c=>{c.addEventListener("click",()=>{let v=i[Number(c.dataset.idx)];S(v)})});let d=document.getElementById("error-detail-section"),m=document.getElementById("error-detail-container");m.innerHTML="",d.style.display="none";let l=document.getElementById("btn-start-import");b===0?(l.disabled=!0,l.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(l.disabled=!1,n>0?(l.innerHTML=`\u{1F680} Import ${b} Data Valid (${n} dilewati)`,l.title="Baris error akan dilewati, baris valid tetap diimport"):l.innerHTML=`\u{1F680} Mulai Import ${b} Data`)}function S(w){let y=document.getElementById("error-detail-section"),b=document.getElementById("error-detail-container");y.style.display="";let n=w.errors.slice(0,100).map(o=>(Array.isArray(o.errors)?o.errors:[]).map(d=>{let m=typeof d=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${o.row}</span></td>
            <td><strong>${m?d.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${m&&d.originalValue!==void 0?d.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${m?d.reason:d}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${m&&d.aliases?`Gunakan salah satu nama kolom:<br><em>${d.aliases}</em>`:m&&d.hint?d.hint:""}
            </td>
          </tr>
        `}).join("")).join("");b.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${w.sheetName} \u2014 ${w.errorCount} baris error dari ${w.total} total
          ${w.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
        </div>
        <div style="overflow-x:auto">
          <table class="data-table error-table" style="min-width:700px">
            <thead>
              <tr>
                <th style="width:80px">Baris</th>
                <th style="width:140px">Kolom Gagal</th>
                <th style="width:140px">Nilai di Excel</th>
                <th>Alasan Error</th>
                <th style="width:220px">\u{1F4A1} Cara Memperbaiki</th>
              </tr>
            </thead>
            <tbody>${n||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${w.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,y.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{h("upload"),document.getElementById("file-info").style.display="none",r.style.display="",t=null,u.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;za(i)?K("Log error berhasil didownload."):K("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let w=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";E(w)});async function E(w){h("importing"),a=Date.now();let y=[];Nn.forEach(d=>{let m=i?.find(l=>l.module===d&&l.mapped?.length>0);m&&y.push(m)});let b=document.getElementById("import-steps-list");b.innerHTML=y.map(d=>`
      <div class="import-step-item" id="step-item-${d.module}">
        <span class="step-item-icon" id="step-icon-${d.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${d.label} <span class="step-item-count">(${d.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${d.module}"></span>
      </div>
    `).join("");let n=document.getElementById("import-bar"),o=document.getElementById("import-current-status"),s={totalSheets:y.length,totalRows:y.reduce((d,m)=>d+m.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let d=0;d<y.length;d++){let m=y[d],l=document.getElementById(`step-icon-${m.module}`),c=document.getElementById(`step-status-${m.module}`);l.textContent="\u{1F504}",c.textContent="Mengimport...",o.textContent=`Mengimport ${m.label}...`,n.style.width=`${Math.round(d/y.length*100)}%`;try{let v=await M(`/api/import/${m.module}`,{method:"POST",body:JSON.stringify({rows:m.mapped,onDuplicate:w})});if(v.ok){let k=v.data;s.inserted+=k.inserted||0,s.skipped+=k.skipped||0,s.moduleResults.push({label:m.label,inserted:k.inserted||0,skipped:k.skipped||0,status:"ok"}),l.textContent="\u2705",c.innerHTML=`<span class="badge badge-success">${k.inserted||0} berhasil</span>${k.skipped>0?` <span class="badge badge-neutral">${k.skipped} skip</span>`:""}`}else s.failed++,s.moduleResults.push({label:m.label,inserted:0,skipped:0,status:"error",error:v.data?.error}),l.textContent="\u274C",c.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(v){s.failed++,s.moduleResults.push({label:m.label,inserted:0,skipped:0,status:"error",error:v.message}),l.textContent="\u274C",c.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ye(150)}n.style.width="100%",o.textContent="Selesai!",await Ye(400),C(s)}function C(w){h("summary");let y=((Date.now()-a)/1e3).toFixed(1),b=w.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${b?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${b?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${w.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${w.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${w.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${w.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${w.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${w.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${y}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${w.moduleResults.map(n=>`
            <tr>
              <td>${n.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${n.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${n.skipped}</span></td>
              <td style="text-align:center">
                ${n.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${n.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,u.value="",document.getElementById("file-info").style.display="none",r.style.display="",h("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ye(e){return new Promise(t=>setTimeout(t,e))}V();le();var Na=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function ze(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let i=new Date(e);i.setHours(0,0,0,0);let a=i.getTime()-t.getTime(),p=Math.ceil(a/(1e3*60*60*24));return p<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(p)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:p,isExpired:!0,isExpiringSoon:!1}:p<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${p} hari)`,shortLabel:`Sisa ${p} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:p,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${p} hari)`,shortLabel:`Aktif (${p} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:p,isExpired:!1,isExpiringSoon:!1}}function Ka(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function lt(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function qa(e){let t=[],i=[],a=[],p="all",h="all",g="all",u="",r="fcms-sp-page-styles";if(!document.getElementById(r)){let n=document.createElement("style");n.id=r,n.textContent=`
      .sp-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .sp-hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .sp-hero-card::after {
        content: '';
        position: absolute;
        right: -50px;
        top: -50px;
        width: 220px;
        height: 220px;
        background: radial-gradient(circle, rgba(239, 68, 68, 0.15) 0%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
      }
      .sp-hero-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
      }
      .sp-hero-title {
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
        letter-spacing: -0.02em;
      }
      .sp-hero-subtitle {
        color: #94A3B8;
        font-size: 0.92rem;
        margin: 0;
        max-width: 720px;
        line-height: 1.5;
      }
      .sp-stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin-top: 24px;
      }
      .sp-stat-box {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .sp-stat-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        flex-shrink: 0;
      }
      .sp-stat-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.2;
      }
      .sp-stat-label {
        font-size: 0.78rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
      }
      .sp-controls-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .sp-controls-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
      }
      .sp-search-wrapper {
        position: relative;
        flex: 1;
        min-width: 250px;
      }
      .sp-search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94A3B8;
        font-size: 0.95rem;
        pointer-events: none;
      }
      .sp-search-input {
        width: 100%;
        padding: 10px 14px 10px 40px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.9rem;
        background: #F8FAFC;
        transition: all 0.2s ease;
      }
      .sp-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #FFFFFF;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .sp-filter-select {
        padding: 9px 14px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.88rem;
        background: #F8FAFC;
        color: #334155;
        font-weight: 500;
        cursor: pointer;
      }
      .sp-filter-select:focus {
        outline: none;
        border-color: #2563EB;
        background: #FFFFFF;
      }
      .sp-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .sp-pill {
        border: 1px solid #E2E8F0;
        background: #F8FAFC;
        color: #475569;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 30px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .sp-pill:hover {
        background: #EEF2F6;
        border-color: #CBD5E1;
      }
      .sp-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .sp-pill-count {
        background: rgba(0, 0, 0, 0.08);
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 0.75rem;
      }
      .sp-pill.active .sp-pill-count {
        background: rgba(255, 255, 255, 0.25);
        color: #FFFFFF;
      }
      /* Modern Table View */
      .sp-table-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        overflow: hidden;
      }
      .sp-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
      }
      .sp-table th {
        background: #F8FAFC;
        padding: 14px 18px;
        font-size: 0.78rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        border-bottom: 1.5px solid #E2E8F0;
      }
      .sp-table td {
        padding: 16px 18px;
        font-size: 0.88rem;
        color: #1E293B;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: middle;
      }
      .sp-table tr:hover td {
        background: #F8FAFC;
      }
      /* Grid View */
      .sp-cards-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
        gap: 20px;
      }
      .sp-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
      }
      .sp-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.1);
        border-color: #CBD5E1;
      }
      .sp-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.02em;
      }
      .sp-type-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 4px 9px;
        border-radius: 6px;
        font-size: 0.74rem;
        font-weight: 700;
      }
      .sp-avatar-box {
        width: 42px;
        height: 42px;
        border-radius: 10px;
        background: #EFF6FF;
        color: #2563EB;
        font-weight: 700;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.95rem;
        flex-shrink: 0;
      }
      .sp-icon-btn {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        color: #64748B;
        width: 32px;
        height: 32px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.15s ease;
      }
      .sp-icon-btn:hover {
        background: #EDF2F7;
        color: #1E293B;
        border-color: #CBD5E1;
      }
      .sp-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      .sp-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #FFFFFF;
        border-radius: 16px;
        border: 1px dashed #CBD5E1;
      }

      /* \u2500\u2500 TAMPILAN RESPONSIVE HP / MOBILE (\u2264 768px) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */
      @media (max-width: 768px) {
        .sp-page-container {
          padding: 12px 14px !important;
        }
        .sp-hero-card {
          padding: 18px 16px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
        }
        .sp-hero-title {
          font-size: 1.25rem !important;
          gap: 8px !important;
        }
        .sp-hero-subtitle {
          font-size: 0.82rem !important;
          line-height: 1.4 !important;
        }
        .sp-hero-header {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 12px !important;
        }
        .sp-hero-actions {
          width: 100% !important;
          display: flex !important;
          align-items: center !important;
          gap: 8px !important;
          flex-wrap: nowrap !important;
        }
        .sp-btn-add {
          flex: 1 1 auto !important;
          justify-content: center !important;
          padding: 8px 10px !important;
          font-size: 0.82rem !important;
          white-space: nowrap !important;
        }
        .sp-btn-excel {
          flex: 1 1 auto !important;
          justify-content: center !important;
          padding: 8px 10px !important;
          font-size: 0.82rem !important;
          white-space: nowrap !important;
        }
        .sp-btn-refresh {
          width: 38px !important;
          height: 38px !important;
          padding: 0 !important;
          flex-shrink: 0 !important;
          justify-content: center !important;
        }
        /* 2x2 compact stats grid di layar HP */
        .sp-stats-grid {
          grid-template-columns: repeat(2, 1fr) !important;
          gap: 8px !important;
          margin-top: 14px !important;
        }
        .sp-stat-box {
          padding: 10px 12px !important;
          gap: 10px !important;
          border-radius: 10px !important;
        }
        .sp-stat-icon {
          width: 36px !important;
          height: 36px !important;
          font-size: 1.1rem !important;
          border-radius: 8px !important;
        }
        .sp-stat-value {
          font-size: 1.15rem !important;
        }
        .sp-stat-label {
          font-size: 0.68rem !important;
        }
        .sp-controls-panel {
          padding: 14px 12px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
          gap: 12px !important;
        }
        .sp-controls-top {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 10px !important;
        }
        .sp-search-wrapper {
          min-width: 100% !important;
          width: 100% !important;
        }
        /* PITA FILTER SWIPEABLE HORIZONTAL DI HP (TIDAK LAGI NUMPUK VERTIKAL) */
        .sp-filter-pills {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          padding-bottom: 6px !important;
          padding-top: 2px !important;
          scrollbar-width: none !important;
          -ms-overflow-style: none !important;
        }
        .sp-filter-pills::-webkit-scrollbar {
          display: none !important;
        }
        .sp-pill {
          flex-shrink: 0 !important;
          white-space: nowrap !important;
          padding: 6px 12px !important;
          font-size: 0.8rem !important;
        }
        .sp-cards-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
        .sp-card {
          padding: 16px !important;
          border-radius: 14px !important;
        }
        .sp-card-header {
          flex-direction: row !important;
          align-items: flex-start !important;
        }
        .sp-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .sp-table {
          min-width: 650px !important;
        }
        .sp-table th, .sp-table td {
          padding: 10px 12px !important;
          font-size: 0.82rem !important;
        }
        .sp-form-grid, .sp-detail-grid {
          grid-template-columns: 1fr !important;
          gap: 10px !important;
        }
      }
    `,document.head.appendChild(n)}async function f(){try{let[n,o]=await Promise.all([ee(),ge()]);i=n||[],a=o||[];let s=await M("/api/sp?limit=1000");if(s.ok&&s.data?.data&&Array.isArray(s.data.data)&&s.data.data.length>0)t=s.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t));else{let d=localStorage.getItem("fcms_sp_data_cache");d?t=JSON.parse(d):(t=[...Na],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)))}}catch(n){console.warn("Fallback loading SP data from cache/seed:",n);let o=localStorage.getItem("fcms_sp_data_cache");t=o?JSON.parse(o):[...Na]}y()}function D(){return t.filter(n=>{let o=u.toLowerCase().trim(),s=!o||(n.employee_name||"").toLowerCase().includes(o)||(n.branch_name||"").toLowerCase().includes(o)||(n.division||"").toLowerCase().includes(o)||(n.sp_type||"").toLowerCase().includes(o),d=h==="all"||String(n.branch_id)===String(h),m=g==="all"||n.division===g,l=ze(n.akhir_sp),c=!0;return p==="active"?c=!l.isExpired:p==="expiring"?c=l.isExpiringSoon:p==="expired"?c=l.isExpired:p==="sp1"?c=(n.sp_type||"").toUpperCase()==="SP 1":p==="sp2"?c=(n.sp_type||"").toUpperCase()==="SP 2":p==="sp3"?c=(n.sp_type||"").toUpperCase()==="SP 3":p==="teguran"&&(c=(n.sp_type||"").toUpperCase().includes("TEGURAN")),s&&d&&m&&c})}function F(){let n=0,o=0,s=0,d=0,m=0,l=0,c=0,v=new Set;return t.forEach(k=>{k.employee_name&&v.add(k.employee_name.toLowerCase().trim());let T=ze(k.akhir_sp);T.isExpired?s++:n++,T.isExpiringSoon&&o++;let A=(k.sp_type||"").toUpperCase();A==="SP 1"?d++:A==="SP 2"?m++:A==="SP 3"?l++:A.includes("TEGURAN")&&c++}),{total:t.length,active:n,expiring:o,expired:s,sp1:d,sp2:m,sp3:l,teguran:c,uniqueEmployees:v.size}}function S(n=null){let o=!!n,s=n?.tanggal?n.tanggal.slice(0,10):new Date().toISOString().slice(0,10),d=n?.akhir_sp?n.akhir_sp.slice(0,10):"";if(!d){let v=new Date(s);v.setMonth(v.getMonth()+6),d=v.toISOString().slice(0,10)}let m=document.createElement("div");m.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${n?.employee_name?String(n.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${a.map(v=>`<option value="${v.label}">${v.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${n?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${n?.division==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">Pilih Cabang...</option>
              ${i.map(v=>`<option value="${v.value}" ${String(n?.branch_id)===String(v.value)?"selected":""}>${v.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${n?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${n?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${n?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${n?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${s}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${n?.document_link?String(n.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${n?.notes||n?.description||""}</textarea>
        </div>
      </div>
    `;let l=m.querySelector("#sp-modal-tanggal"),c=m.querySelector("#sp-modal-akhir");l.addEventListener("change",()=>{if(l.value){let v=new Date(l.value);v.setMonth(v.getMonth()+6),c.value=v.toISOString().slice(0,10)}}),J({title:o?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:m,size:"md",confirmText:o?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(v,k)=>{let T=(m.querySelector("#sp-modal-emp").value||"").trim(),A=m.querySelector("#sp-modal-div").value,P=m.querySelector("#sp-modal-branch").value,_=m.querySelector("#sp-modal-type").value,I=m.querySelector("#sp-modal-tanggal").value,z=m.querySelector("#sp-modal-akhir").value,$=(m.querySelector("#sp-modal-link").value||"").trim(),x=(m.querySelector("#sp-modal-notes").value||"").trim();if(!T){N("Nama Karyawan wajib diisi");return}if(!P){N("Cabang penempatan wajib dipilih");return}if(!I||!z){N("Tanggal SP dan Akhir SP wajib diisi");return}let B=i.find(O=>String(O.value)===String(P)),L=B?B.label:"Cabang",H={employee_name:T,division:A,branch_id:Number(P),branch_name:L,sp_type:_,tanggal:I,akhir_sp:z,document_link:$,notes:x,status:ze(z).status};try{if(o){let O=await M(`/api/sp/${n.id}`,{method:"PUT",body:JSON.stringify(H)}),j=t.findIndex(q=>q.id===n.id);j!==-1&&(t[j]={...t[j],...H}),K("Data SP berhasil diperbarui")}else{let O=await M("/api/sp",{method:"POST",body:JSON.stringify(H)}),j=O.ok&&O.data?.id?O.data.id:Date.now();t.unshift({id:j,...H}),K("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),k(),y()}catch(O){console.error("Error saving SP:",O),N("Gagal menyimpan SP: "+O.message)}}})}function E(n){W(`Hapus data Surat Peringatan untuk karyawan <strong>${n.employee_name}</strong> (${n.sp_type})?`,async()=>{try{await M(`/api/sp/${n.id}`,{method:"DELETE"})}catch(o){console.warn("API delete error fallback to local:",o)}t=t.filter(o=>o.id!==n.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),K(`Data SP ${n.employee_name} berhasil dihapus`),y()},"Hapus Data SP")}function C(n){let o=ze(n.akhir_sp),s=Ka(n.sp_type),d=document.createElement("div");d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${n.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${n.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${n.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${n.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${s.bg}; color:${s.color}; border:1px solid ${s.border}; font-size:0.85rem;">
              <span>${s.icon}</span>
              <span>${s.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${o.bg}; color:${o.color}; border:1px solid ${o.border}; font-size:0.85rem;">
              <span>${o.icon}</span>
              <span>${o.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${lt(n.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${lt(n.akhir_sp)}</div>
          </div>
        </div>

        ${n.notes||n.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${n.notes||n.description}</div>
          </div>
        `:""}

        ${n.document_link?`
          <div>
            <a href="${n.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,J({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:d,size:"md",cancelText:"Tutup"})}async function w(){try{let{downloadExcel:n}=await Promise.resolve().then(()=>(ie(),he)),o=t.map((s,d)=>{let m=ze(s.akhir_sp);return{No:d+1,"Nama Karyawan":s.employee_name||"",Divisi:s.division||"",Cabang:s.branch_name||"","Jenis SP":s.sp_type||"","Tanggal SP":s.tanggal||"","Akhir SP":s.akhir_sp||"","Status Masa Berlaku":m.status,"Sisa Hari":m.daysLeft>0?m.daysLeft:0,"Link BAP / Foto":s.document_link||"","Catatan / Kronologi":s.notes||s.description||""}});n(o,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),K("Berhasil mengunduh data SP ke format Excel.")}catch(n){N("Gagal ekspor: "+n.message)}}function y(){let n=D(),o=F();e.innerHTML=`
      <div class="sp-page-container">
        
        <!-- Hero Header -->
        <div class="sp-hero-card">
          <div class="sp-hero-header">
            <div>
              <h1 class="sp-hero-title">
                <span>\u{1F4DC}</span> Monitoring Data SP (Surat Peringatan)
              </h1>
              <p class="sp-hero-subtitle">
                Sistem pengawasan kedisiplinan kerja, masa aktif sanksi pembinaan, dan arsip dokumen BAP karyawan PT Audy Mandiri Indonesia.
              </p>
            </div>
            <div class="sp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              <button id="sp-btn-add" class="btn btn-primary sp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                <span>Terbitkan SP</span>
              </button>
              
              <button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="sp-btn-refresh" class="btn btn-ghost sp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="sp-stats-grid">
            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(239,68,68,0.2); color:#F87171;">\u{1F534}</div>
              <div>
                <div class="sp-stat-value">${o.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${o.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${o.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${o.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
                <div class="sp-stat-label">Total Terkena Sanksi</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls Panel -->
        <div class="sp-controls-panel">
          <div class="sp-controls-top">
            <div class="sp-search-wrapper">
              <span class="sp-search-icon">\u{1F50D}</span>
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${u?u.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(s=>`<option value="${s.value}" ${String(h)===String(s.value)?"selected":""}>${s.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${g==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${g==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${g==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${p==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${o.total}</span>
            </button>
            <button class="sp-pill ${p==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${o.active}</span>
            </button>
            <button class="sp-pill ${p==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${o.expiring}</span>
            </button>
            <button class="sp-pill ${p==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${o.expired}</span>
            </button>
            <button class="sp-pill ${p==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${o.sp1}</span>
            </button>
            <button class="sp-pill ${p==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${o.sp2}</span>
            </button>
            <button class="sp-pill ${p==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${o.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${n.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="sp-cards-grid">
              ${n.map(s=>{let d=ze(s.akhir_sp),m=Ka(s.sp_type);return`
                  <div class="sp-card" data-id="${s.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${m.bg}; color:${m.color}; border:1px solid ${m.border};">
                          <span>${m.icon}</span>
                          <span>${m.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border};">
                          <span>${d.icon}</span>
                          <span>${d.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${s.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${s.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${s.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${s.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${lt(s.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${lt(s.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${s.notes||s.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${s.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${s.document_link?`
                          <a href="${s.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${s.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${s.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,b()}function b(){let n=e.querySelector("#sp-btn-add");n&&(n.onclick=()=>S());let o=e.querySelector("#sp-btn-excel");o&&(o.onclick=()=>w());let s=e.querySelector("#sp-btn-refresh");s&&(s.onclick=()=>f());let d=e.querySelector("#sp-btn-reset-filter");d&&(d.onclick=()=>{u="",h="all",g="all",p="all",y()});let m=e.querySelector("#sp-search-input");m&&(m.oninput=v=>{u=v.target.value,y();let k=e.querySelector("#sp-search-input");k&&(k.focus(),k.setSelectionRange(k.value.length,k.value.length))});let l=e.querySelector("#sp-filter-branch");l&&(l.onchange=v=>{h=v.target.value,y()});let c=e.querySelector("#sp-filter-division");c&&(c.onchange=v=>{g=v.target.value,y()}),e.querySelectorAll(".sp-pill").forEach(v=>{v.onclick=()=>{p=v.getAttribute("data-filter")||"all",y()}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id")),A=t.find(P=>P.id===T);A&&C(A)}}),e.querySelectorAll(".sp-edit-btn").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id")),A=t.find(P=>P.id===T);A&&S(A)}}),e.querySelectorAll(".sp-delete-btn").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id")),A=t.find(P=>P.id===T);A&&E(A)}})}await f()}V();le();var Nt=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function Kn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Oa(e){let t=[],i=[],a=[],p="all",h="all",g="all",u="",r=!1,D=Z()?.role!=="viewer";try{[i,a]=await Promise.all([ee(),ge()])}catch(n){console.error("Error fetching cache:",n)}async function F(){r=!0,y();try{let n=await M("/api/mutasi?limit=10000");if(n.ok){let o=n.data?.data||n.data||[];t=Array.isArray(o)&&o.length>0?o:Nt}else t=Nt}catch(n){console.warn("Gagal load mutasi, menggunakan data lokal:",n),t=Nt}finally{r=!1,y()}}async function S(){try{let{downloadExcel:n}=await Promise.resolve().then(()=>(ie(),he)),o=t.map((s,d)=>({No:d+1,Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"","Cabang Asal":s.from_branch_name||"","Cabang Tujuan":s.to_branch_name||"",Status:s.status||"","Link Dokumen":s.document_link||""}));n(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),K("Data mutasi berhasil diekspor ke Excel!")}catch(n){N("Gagal ekspor Excel: "+n.message)}}async function E(n){if(!D)return;let o=t.find(s=>s.id===n);W({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${o?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let s=await M(`/api/mutasi/${n}`,{method:"DELETE"});s.ok?(K("Data mutasi berhasil dihapus"),t=t.filter(d=>d.id!==n),y()):N(s.data?.error||"Gagal menghapus mutasi")}catch(s){N("Gagal menghapus data mutasi: "+s.message)}}})}function C(n=null){if(!D)return;let o=!!n,s=n?.tanggal?n.tanggal.slice(0,10):new Date().toISOString().slice(0,10),d=document.createElement("div");d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${s}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${n?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${n?.status==="Selesai"||!n?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <input list="emp-list-options" id="mutasi-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik atau pilih nama karyawan..." value="${n?.employee_name?String(n.employee_name).replace(/"/g,"&quot;"):""}" required />
          <datalist id="emp-list-options">
            ${a.map(m=>`<option value="${typeof m=="string"?m:m.label||m.name||""}">`).join("")}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${i.map(m=>`
                <option value="${m.value}" ${String(n?.from_branch_id)===String(m.value)?"selected":""}>
                  ${m.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Tujuan <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-to" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Tujuan --</option>
              ${i.map(m=>`
                <option value="${m.value}" ${String(n?.to_branch_id)===String(m.value)?"selected":""}>
                  ${m.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas SK Mutasi / Dokumen (Opsional)
          </label>
          <input type="url" id="mutasi-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/file/d/..." value="${n?.document_link?String(n.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `,J({title:o?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:d,size:"md",confirmText:o?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(m,l)=>{let c=d.querySelector("#mutasi-modal-tanggal").value,v=(d.querySelector("#mutasi-modal-emp").value||"").trim(),k=d.querySelector("#mutasi-modal-from").value,T=d.querySelector("#mutasi-modal-to").value,A=d.querySelector("#mutasi-modal-status").value,P=(d.querySelector("#mutasi-modal-link").value||"").trim();if(!c){N("Tanggal mutasi wajib diisi");return}if(!v){N("Nama karyawan wajib diisi");return}if(!k){N("Cabang asal wajib dipilih");return}if(!T){N("Cabang tujuan wajib dipilih");return}let _=i.find($=>String($.value)===String(k)),I=i.find($=>String($.value)===String(T)),z={tanggal:c,employee_name:v,from_branch_id:Number(k),from_branch_name:_?_.label:"Cabang Asal",to_branch_id:Number(T),to_branch_name:I?I.label:"Cabang Tujuan",status:A,document_link:P};try{if(o){let $=await M(`/api/mutasi/${n.id}`,{method:"PUT",body:JSON.stringify(z)});if($.ok){K("Data mutasi berhasil diperbarui");let x=t.findIndex(B=>B.id===n.id);x!==-1&&(t[x]={...t[x],...z}),l(),y()}else N($.data?.error||"Gagal memperbarui data mutasi")}else{let $=await M("/api/mutasi",{method:"POST",body:JSON.stringify(z)});if($.ok){K("Data mutasi baru berhasil ditambahkan");let x=$.data?.id||Date.now();t.unshift({id:x,...z}),l(),y()}else N($.data?.error||"Gagal menambahkan mutasi")}}catch($){N("Terjadi kesalahan: "+$.message)}}})}function w(){return t.filter(n=>{if(u){let o=u.toLowerCase(),s=(n.employee_name||"").toLowerCase().includes(o),d=(n.from_branch_name||"").toLowerCase().includes(o),m=(n.to_branch_name||"").toLowerCase().includes(o);if(!s&&!d&&!m)return!1}return!(h!=="all"&&String(n.from_branch_id)!==String(h)||g!=="all"&&String(n.to_branch_id)!==String(g)||p!=="all"&&n.status!==p)})}function y(){let n=w(),o=t.length,s=t.filter(k=>k.status==="Proses").length,d=t.filter(k=>k.status==="Selesai").length,m=new Date,l=m.getFullYear(),c=m.getMonth(),v=t.filter(k=>{if(!k.tanggal)return!1;let T=new Date(k.tanggal);return!isNaN(T.getTime())&&T.getFullYear()===l&&T.getMonth()===c}).length;e.innerHTML=`
      <style>
        .mutasi-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .mutasi-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .mutasi-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.15) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .mutasi-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 22px;
        }
        .mutasi-stat-box {
          background: rgba(255, 255, 255, 0.06);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 16px;
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .mutasi-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .mutasi-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          color: #F8FAFC;
        }
        .mutasi-stat-lbl {
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #94A3B8;
          font-weight: 600;
          margin-top: 2px;
        }

        /* Controls Card */
        .mutasi-controls-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .mutasi-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .mutasi-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .mutasi-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .mutasi-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .mutasi-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .mutasi-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 180px;
          cursor: pointer;
        }
        .mutasi-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .mutasi-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .mutasi-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .mutasi-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .mutasi-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .mutasi-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* Grid Cards View */
        .mutasi-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .mutasi-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          transition: transform 0.15s ease, box-shadow 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          position: relative;
        }
        .mutasi-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.06);
          border-color: #CBD5E1;
        }
        .mutasi-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          box-shadow: 0 3px 6px rgba(37, 99, 235, 0.2);
        }
        .mutasi-route-box {
          background: #F8FAFC;
          border: 1px dashed #CBD5E1;
          border-radius: 10px;
          padding: 12px 14px;
          margin: 12px 0;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
        }
        .mutasi-route-node {
          flex: 1;
        }
        .mutasi-route-label {
          font-size: 0.7rem;
          text-transform: uppercase;
          font-weight: 700;
          letter-spacing: 0.04em;
          color: #94A3B8;
          margin-bottom: 2px;
        }
        .mutasi-route-branch {
          font-size: 0.85rem;
          font-weight: 700;
          color: #1E293B;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .mutasi-route-arrow {
          display: flex;
          align-items: center;
          justify-content: center;
          width: 32px;
          height: 32px;
          border-radius: 50%;
          background: #EFF6FF;
          color: #2563EB;
          font-size: 1rem;
          flex-shrink: 0;
        }
        .mutasi-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
        }
        .mutasi-status-badge.selesai {
          background: #DCFCE7;
          color: #15803D;
          border: 1px solid #BBF7D0;
        }
        .mutasi-status-badge.proses {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .mutasi-icon-btn {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .mutasi-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .mutasi-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .mutasi-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .mutasi-hero-card {
            padding: 18px 16px;
          }
          .mutasi-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .mutasi-stat-box {
            padding: 10px 12px;
          }
          .mutasi-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .mutasi-stat-val {
            font-size: 1.3rem;
          }
          .mutasi-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .mutasi-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .mutasi-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .mutasi-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .mutasi-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .mutasi-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .mutasi-select {
            width: 100%;
          }
          .mutasi-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="mutasi-page">

        <!-- Hero Header Card -->
        <div class="mutasi-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u{1F501}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Mutasi Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Sistem pengelolaan rotasi penempatan tugas, arsip SK mutasi cabang, dan riwayat perpindahan personil PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div class="mutasi-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${D?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              <button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="mutasi-btn-refresh" class="btn btn-ghost mutasi-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="mutasi-stats-row">
            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4CB}
              </div>
              <div>
                <div class="mutasi-stat-val">${o}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${s}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val">${d}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val">${v}</div>
                <div class="mutasi-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="mutasi-controls-card">
          <div class="mutasi-search-row">
            <div class="mutasi-search-wrap">
              <span class="mutasi-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${u}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${i.length})</option>
              ${i.map(k=>`<option value="${k.value}" ${String(h)===String(k.value)?"selected":""}>${k.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${i.length})</option>
              ${i.map(k=>`<option value="${k.value}" ${String(g)===String(k.value)?"selected":""}>${k.label}</option>`).join("")}
            </select>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${p==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${o}</span>
            </div>
            <div class="mutasi-pill ${p==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${s}</span>
            </div>
            <div class="mutasi-pill ${p==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${d}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${n.length===0?`
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <div class="mutasi-cards-grid">
              ${n.map(k=>{let T=k.status==="Selesai",A=(k.employee_name||"Karyawan").split(" ").map(P=>P[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${k.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${T?"selesai":"proses"}">
                          <span>${T?"\u2705":"\u23F3"}</span>
                          <span>${k.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${Kn(k.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${A}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${k.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${k.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${k.from_branch_name||"-"}">
                            ${k.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node" style="text-align:right;">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${k.to_branch_name||"-"}">
                            ${k.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${k.document_link?`
                          <a href="${k.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${D?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${k.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${k.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        </div>
                      `:""}
                    </div>

                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,b()}function b(){let n=e.querySelector("#mutasi-btn-add");n&&(n.onclick=()=>C());let o=e.querySelector("#mutasi-btn-excel");o&&(o.onclick=()=>S());let s=e.querySelector("#mutasi-btn-refresh");s&&(s.onclick=()=>F());let d=e.querySelector("#mutasi-btn-reset-filter");d&&(d.onclick=()=>{u="",h="all",g="all",p="all",y()});let m=e.querySelector("#mutasi-search-input");m&&(m.oninput=v=>{u=v.target.value,y();let k=e.querySelector("#mutasi-search-input");k&&(k.focus(),k.setSelectionRange(k.value.length,k.value.length))});let l=e.querySelector("#mutasi-filter-from");l&&(l.onchange=v=>{h=v.target.value,y()});let c=e.querySelector("#mutasi-filter-to");c&&(c.onchange=v=>{g=v.target.value,y()}),e.querySelectorAll(".mutasi-pill").forEach(v=>{v.onclick=()=>{p=v.getAttribute("data-status")||"all",y()}}),e.querySelectorAll(".mutasi-edit-btn").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id")),A=t.find(P=>P.id===T);A&&C(A)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(v=>{v.onclick=k=>{k.stopPropagation();let T=Number(v.getAttribute("data-id"));E(T)}})}await F()}V();le();var Kt=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],Ra=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function Ha(e,t,i){if(!e||!t)return 0;try{let[a,p]=e.split(":").map(Number),[h,g]=t.split(":").map(Number),u=h*60+g-(a*60+p);u<0&&(u+=24*60);let r=u/60-(parseFloat(i)||0);return r<0&&(r=0),Number.isInteger(r)?r:parseFloat(r.toFixed(2))}catch{return 0}}function Ua(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function Ga(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Ja(e){let t=[],i=[],a=[],p="all",h="all",g="",u=!1,r=Z(),f=r?.role==="viewer",D=r?.role==="input_lembur",F=!f,S=!f,E=!f&&!D;try{[i,a]=await Promise.all([ee(),_e()])}catch(d){console.error("Error fetching cache:",d)}async function C(){u=!0,o();try{let d=await M("/api/overtime?limit=10000");if(d.ok){let m=d.data?.data||d.data||[];t=Array.isArray(m)&&m.length>0?m:Kt}else t=Kt}catch(d){console.warn("Gagal load lembur, menggunakan data lokal:",d),t=Kt}finally{u=!1,o()}}async function w(){try{let{downloadExcel:d}=await Promise.resolve().then(()=>(ie(),he)),m=t.map((l,c)=>({No:c+1,Tanggal:l.date||"","Cabang Lembur":l.branch_name||"","Nama Karyawan":l.employee_name||"","Jam Mulai":l.start_time||"","Jam Selesai":l.end_time||"","Istirahat (Jam)":l.break_hours!=null?l.break_hours:0,"Total Jam":l.total_hours!=null?l.total_hours:0,"Alasan / Keterangan":l.reason||""}));d(m,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),K("Data lembur berhasil diekspor ke Excel!")}catch(d){N("Gagal ekspor Excel: "+d.message)}}async function y(d){if(!E)return;let m=t.find(l=>l.id===d);W({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${m?.employee_name||"karyawan ini"}" pada tanggal ${Ga(m?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let l=await M(`/api/overtime/${d}`,{method:"DELETE"});l.ok?(K("Catatan lembur berhasil dihapus"),t=t.filter(c=>c.id!==d),o()):N(l.data?.error||"Gagal menghapus lembur")}catch(l){N("Gagal menghapus catatan lembur: "+l.message)}}})}function b(d=null){if(!S)return;let m=!!d,l=d?.date?d.date.slice(0,10):new Date().toISOString().slice(0,10),c=d?.start_time||"22:00",v=d?.end_time||"08:00",k=d?.break_hours!=null?d.break_hours:1,T=d?.total_hours!=null?d.total_hours:Ha(c,v,k),A=document.createElement("div");A.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(x=>`
                <option value="${x.value}" ${String(d?.branch_id)===String(x.value)?"selected":""}>
                  ${x.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            <option value="">-- Pilih Karyawan --</option>
            ${a.map(x=>{let B=typeof x=="string"?x:x.label||x.full_name||x.name||"",L=typeof x=="object"&&x!==null?x.value!=null?x.value:x.id!=null?x.id:B:x,H=String(d?.employee_id)===String(L)||d?.employee_name===B;return`<option value="${L}" ${H?"selected":""}>${B}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${c}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${v}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${k}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${T}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${Ra.map(x=>`
              <option value="${x}" ${d?.reason===x?"selected":""}>${x}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let P=A.querySelector("#ot-modal-start"),_=A.querySelector("#ot-modal-end"),I=A.querySelector("#ot-modal-break"),z=A.querySelector("#ot-modal-total");function $(){let x=P.value,B=_.value,L=I.value;z.value=Ha(x,B,L)}P.addEventListener("input",$),_.addEventListener("input",$),I.addEventListener("input",$),J({title:m?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:A,size:"md",confirmText:m?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(x,B)=>{let L=A.querySelector("#ot-modal-date").value,H=A.querySelector("#ot-modal-branch").value,O=A.querySelector("#ot-modal-emp").value,j=A.querySelector("#ot-modal-start").value,q=A.querySelector("#ot-modal-end").value,R=parseFloat(A.querySelector("#ot-modal-break").value)||0,Q=parseFloat(A.querySelector("#ot-modal-total").value)||0,ne=A.querySelector("#ot-modal-reason").value;if(!L){N("Tanggal lembur wajib diisi");return}if(!H){N("Cabang penugasan wajib dipilih");return}if(!O){N("Nama karyawan wajib dipilih");return}if(!j||!q){N("Jam mulai dan jam selesai wajib diisi");return}let oe=i.find(U=>String(U.value)===String(H)),G=a.find(U=>String(U.value)===String(O)||String(U.id)===String(O)),ae=G?G.label||G.full_name||G.name||"Karyawan":typeof O=="string"&&isNaN(Number(O))?O:"Karyawan",X={date:L,branch_id:Number(H),branch_name:oe?oe.label:"Cabang Lembur",employee_id:Number(O),employee_name:G?G.label:"Karyawan",start_time:j,end_time:q,break_hours:R,total_hours:Q,reason:ne};try{if(m){let U=await M(`/api/overtime/${d.id}`,{method:"PUT",body:JSON.stringify(X)});if(U.ok){K("Data lembur berhasil diperbarui");let ue=t.findIndex(xe=>xe.id===d.id);ue!==-1&&(t[ue]={...t[ue],...X}),B(),o()}else N(U.data?.error||"Gagal memperbarui data lembur")}else{let U=await M("/api/overtime",{method:"POST",body:JSON.stringify(X)});if(U.ok){K("Data lembur berhasil ditambahkan");let ue=U.data?.id||Date.now();t.unshift({id:ue,...X}),B(),o()}else N(U.data?.error||"Gagal menambahkan lembur")}}catch(U){N("Terjadi kesalahan: "+U.message)}}})}function n(){return t.filter(d=>{if(g){let m=g.toLowerCase(),l=(d.employee_name||"").toLowerCase().includes(m),c=(d.branch_name||"").toLowerCase().includes(m),v=(d.reason||"").toLowerCase().includes(m);if(!l&&!c&&!v)return!1}return!(h!=="all"&&String(d.branch_id)!==String(h)||p!=="all"&&d.reason!==p)})}function o(){let d=n(),m=t.length,l=t.reduce((P,_)=>P+(parseFloat(_.total_hours)||0),0),c=m>0?(l/m).toFixed(1):0,v=new Date,k=v.getFullYear(),T=v.getMonth(),A=t.filter(P=>{if(!P.date)return!1;let _=new Date(P.date);return!isNaN(_.getTime())&&_.getFullYear()===k&&_.getMonth()===T}).length;e.innerHTML=`
      <style>
        .lembur-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .lembur-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .lembur-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(99, 102, 241, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .lembur-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 22px;
        }
        .lembur-stat-box {
          background: rgba(255, 255, 255, 0.06);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 16px;
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .lembur-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .lembur-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          color: #F8FAFC;
        }
        .lembur-stat-lbl {
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #94A3B8;
          font-weight: 600;
          margin-top: 2px;
        }

        /* Controls Card */
        .lembur-controls-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .lembur-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .lembur-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .lembur-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .lembur-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .lembur-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .lembur-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 200px;
          cursor: pointer;
        }
        .lembur-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .lembur-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .lembur-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .lembur-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .lembur-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .lembur-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* Grid Cards View */
        .lembur-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .lembur-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          transition: transform 0.15s ease, box-shadow 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          position: relative;
        }
        .lembur-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.06);
          border-color: #CBD5E1;
        }
        .lembur-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: linear-gradient(135deg, #4F46E5 0%, #3730A3 100%);
          color: white;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          box-shadow: 0 3px 6px rgba(79, 70, 229, 0.25);
        }
        .lembur-time-box {
          background: #F8FAFC;
          border: 1px solid #E2E8F0;
          border-radius: 10px;
          padding: 10px 12px;
          margin: 12px 0;
          display: flex;
          align-items: center;
          justify-content: space-between;
        }
        .lembur-hours-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 5px 12px;
          border-radius: 8px;
          font-size: 0.88rem;
          font-weight: 800;
          background: #EEF2FF;
          color: #4F46E5;
          border: 1px solid #C7D2FE;
        }
        .lembur-reason-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .lembur-icon-btn {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .lembur-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .lembur-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .lembur-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .lembur-hero-card {
            padding: 18px 16px;
          }
          .lembur-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .lembur-stat-box {
            padding: 10px 12px;
          }
          .lembur-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .lembur-stat-val {
            font-size: 1.3rem;
          }
          .lembur-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .lembur-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .lembur-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .lembur-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .lembur-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .lembur-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .lembur-select {
            width: 100%;
          }
          .lembur-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="lembur-page">

        <!-- Hero Header Card -->
        <div class="lembur-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u23F1\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Lembur Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pencatatan jam kerja lembur, approval penugasan overtime cabang, dan rekapitulasi durasi jam kerja PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div class="lembur-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${F?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              <button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="lembur-btn-refresh" class="btn btn-ghost lembur-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="lembur-stats-row">
            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(79,70,229,0.25); color:#818CF8;">
                \u23F1\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${l} Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val">${m}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val">${c} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${A}</div>
                <div class="lembur-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="lembur-controls-card">
          <div class="lembur-search-row">
            <div class="lembur-search-wrap">
              <span class="lembur-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${g}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(P=>`<option value="${P.value}" ${String(h)===String(P.value)?"selected":""}>${P.label}</option>`).join("")}
            </select>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${p==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${m}</span>
            </div>
            ${Ra.map(P=>{let _=t.filter(z=>z.reason===P).length,I=Ua(P);return`
                <div class="lembur-pill ${p===P?"active":""}" data-reason="${P}">
                  <span>${I.icon}</span>
                  <span>${P}</span>
                  ${_>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${_}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${d.length===0?`
            <div class="lembur-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Lembur Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="lembur-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <div class="lembur-cards-grid">
              ${d.map(P=>{let _=(P.employee_name||"Karyawan").split(" ").map(z=>z[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),I=Ua(P.reason);return`
                  <div class="lembur-card" data-id="${P.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${P.total_hours!=null?P.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${P.break_hours!=null?P.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${_}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${P.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${P.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${Ga(P.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${P.start_time||"00:00"} - ${P.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${I.bg}; color:${I.color}; border:1px solid ${I.border};">
                          <span>${I.icon}</span>
                          <span>${P.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${S?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${P.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        ${E?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${P.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        `:""}
                      </div>
                    `:""}

                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,s()}function s(){let d=e.querySelector("#lembur-btn-add");d&&(d.onclick=()=>b());let m=e.querySelector("#lembur-btn-excel");m&&(m.onclick=()=>w());let l=e.querySelector("#lembur-btn-refresh");l&&(l.onclick=()=>C());let c=e.querySelector("#lembur-btn-reset-filter");c&&(c.onclick=()=>{g="",h="all",p="all",o()});let v=e.querySelector("#lembur-search-input");v&&(v.oninput=T=>{g=T.target.value,o();let A=e.querySelector("#lembur-search-input");A&&(A.focus(),A.setSelectionRange(A.value.length,A.value.length))});let k=e.querySelector("#lembur-filter-branch");k&&(k.onchange=T=>{h=T.target.value,o()}),e.querySelectorAll(".lembur-pill").forEach(T=>{T.onclick=()=>{p=T.getAttribute("data-reason")||"all",o()}}),e.querySelectorAll(".lembur-edit-btn").forEach(T=>{T.onclick=A=>{A.stopPropagation();let P=Number(T.getAttribute("data-id")),_=t.find(I=>I.id===P);_&&b(_)}}),e.querySelectorAll(".lembur-delete-btn").forEach(T=>{T.onclick=A=>{A.stopPropagation();let P=Number(T.getAttribute("data-id"));y(P)}})}await C()}V();async function Ya(e){let t=Z();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}se({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let h=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(f=>String(f.id)===String(i));if(!h)return alert("Data tidak ditemukan");let g=f=>{if(!f)return"Tidak ada data";try{return JSON.stringify(JSON.parse(f),null,2)}catch{return f}},u=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${g(h.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${g(h.new_data)}</pre>
           </div>
         </div>
       `,{createModal:r}=await Promise.resolve().then(()=>(le(),ht));r({title:`Detail Audit Log #${i}`,content:u,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}V();le();ie();var qt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ve={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function Va(e){e.innerHTML=`
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>\u{1F504}</span> Time Line Migrasi OS
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Jadwal & Status Perpindahan Vendor Outsourcing (AEON & Triple S) ke Seluruh Cabang Audy Dental
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-board-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:6px 12px; cursor:pointer;">
            \u{1F5C2}\uFE0F Matrix Board
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            \u{1F4CB} Tabel Data
          </button>
        </div>
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Cabang
        </button>
      </div>
    </div>

    <!-- Summary KPI Cards -->
    <div id="migration-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:16px 20px; margin-bottom:24px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
      <div style="display:flex; flex-wrap:wrap; gap:14px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="migration-search" class="form-control" placeholder="Cari cabang, vendor, catatan..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
          </div>
          <select id="filter-vendor" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:140px;">
            <option value="">Semua Vendor</option>
            <option value="AEON">AEON</option>
            <option value="Triple S">Triple S</option>
          </select>
          <select id="filter-month" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:150px;">
            <option value="">Semua Bulan</option>
            <option value="2026-09">September 2026</option>
            <option value="2026-10">Oktober 2026</option>
            <option value="2026-11">November 2026</option>
            <option value="2026-12">Desember 2026</option>
          </select>
          <select id="filter-status" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:140px;">
            <option value="">Semua Status</option>
            <option value="scheduled">\u{1F7E1} Terjadwal</option>
            <option value="in_progress">\u{1F535} Sedang Proses</option>
            <option value="completed">\u{1F7E2} Selesai</option>
          </select>
        </div>
        <div id="filter-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="migration-view-container"></div>
  `;let t="board",i=null,a=[],p=Z(),h=p&&["superadmin","admin","manager","spv"].includes(p.role);M("/api/branches?all=1").then(b=>{b.ok&&b.data?.data&&(a=b.data.data)}).catch(()=>{});async function g(){try{let b=await M("/api/os-migrations/summary");if(!b.ok)throw new Error("Gagal memuat data migrasi");i=b.data?.data||b.data,u(i),f()}catch(b){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${b.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function u(b){let n=document.getElementById("migration-summary-cards");if(!n||!b)return;let o=b.total_branches||0,s=b.vendors?.AEON?.total||0,d=b.vendors?.["Triple S"]?.total||0,m=b.by_status?.completed||0,l=b.by_status?.in_progress||0,c=b.by_status?.scheduled||0,v=o>0?Math.round(m/o*100):0;n.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${o} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${s} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${d} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${v}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${v}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${m}</strong></span>
          <span>\u{1F535} Proses: <strong>${l}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${c}</strong></span>
        </div>
      </div>
    `}function r(){if(!i||!i.matrix)return[];let b=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),n=document.getElementById("filter-vendor")?.value||"",o=document.getElementById("filter-month")?.value||"",s=document.getElementById("filter-status")?.value||"";return i.matrix.filter(d=>!(n&&d.vendor!==n||o&&d.target_month!==o||s&&d.status!==s||b&&!`${d.branch_code} ${d.branch_name} ${d.branch_full_name} ${d.vendor} ${d.notes||""} ${d.pic_vendor||""}`.toLowerCase().includes(b)))}function f(){let b=r(),n=document.getElementById("filter-stats-text");n&&(n.innerHTML=`Menampilkan <strong>${b.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let o=document.getElementById("migration-view-container");t==="board"?D(o,b):F(o,b)}function D(b,n){if(!n.length){b.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let o=["AEON","Triple S"],s=["2026-09","2026-10","2026-11","2026-12"],d="";o.forEach(m=>{let l=n.filter(P=>P.vendor===m);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==m||l.length===0)return;let c=l.length,v=m==="AEON"?"#1E40AF":"#0F766E",k=m==="AEON"?"#EFF6FF":"#F0FDFA",T=m==="AEON"?"#BFDBFE":"#99F6E4",A=s.filter(P=>l.some(_=>_.target_month===P));d+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${k}; border-bottom:1px solid ${T}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${v};">${m==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${v}; border:1px solid ${T}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${c} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${A.map(P=>{let _=l.filter($=>$.target_month===P);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(qt[P]||P).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${_.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${_.map($=>{let x=Ve[$.status]||Ve.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${$.branch_code}.</span>${$.branch_name}
                            </div>
                            ${h?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${$.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${$.id}" data-name="${$.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${$.id}" data-status="${$.status}" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${x.icon} ${x.label} \u25BE
                              </button>
                            </div>

                            ${$.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${$.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${$.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${$.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${$.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),b.innerHTML=d,S(b)}function F(b,n){if(!n.length){b.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}b.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">KODE & CABANG</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">VENDOR</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">TARGET BULAN</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">STATUS</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">REALISASI</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SDM FC</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">CATATAN</th>
                ${h?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${n.map((o,s)=>{let d=Ve[o.status]||Ve.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${s+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${o.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${o.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${o.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${o.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${qt[o.target_month]||o.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${o.id}" data-status="${o.status}" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${d.icon} ${d.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${o.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${o.employee_count>0?`\u{1F464} ${o.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${o.notes||"-"}
                    </td>
                    ${h?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${o.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${o.id}" data-name="${o.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,S(b)}function S(b){b.querySelectorAll(".status-toggle-btn").forEach(n=>{n.addEventListener("click",o=>{o.stopPropagation();let s=n.dataset.id,d=n.dataset.status;E(n,s,d)})}),b.querySelectorAll(".edit-card-btn").forEach(n=>{n.addEventListener("click",()=>{let o=n.dataset.id,s=i?.matrix?.find(d=>String(d.id)===String(o));s&&C(s)})}),b.querySelectorAll(".delete-card-btn").forEach(n=>{n.addEventListener("click",()=>{let o=n.dataset.id,s=n.dataset.name||"Cabang";W(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${s}</strong>?`,async()=>{try{if(!(await M(`/api/os-migrations/${o}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");K(`Jadwal migrasi ${s} berhasil dihapus`),g()}catch(d){N(d.message)}})})})}function E(b,n,o){document.querySelectorAll(".status-popup-menu").forEach(c=>c.remove());let s=b.getBoundingClientRect(),d=document.createElement("div");d.className="status-popup-menu",d.style.cssText=`
      position: fixed;
      top: ${s.bottom+window.scrollY+4}px;
      left: ${s.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let m=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];d.innerHTML=m.map(c=>`
      <button class="status-option-item" data-key="${c.key}" style="
        width: 100%;
        text-align: left;
        padding: 8px 14px;
        background: none;
        border: none;
        font-size: 0.85rem;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 8px;
        color: ${c.key===o?"#2563EB":"var(--text-1)"};
        font-weight: ${c.key===o?"700":"500"};
      ">
        ${c.label} ${c.key===o?"\u2713":""}
      </button>
    `).join("");let l=c=>{!d.contains(c.target)&&c.target!==b&&(d.remove(),document.removeEventListener("click",l))};setTimeout(()=>document.addEventListener("click",l),50),d.querySelectorAll(".status-option-item").forEach(c=>{c.addEventListener("click",async()=>{let v=c.dataset.key;if(d.remove(),document.removeEventListener("click",l),v!==o)try{if(!(await M(`/api/os-migrations/${n}/status`,{method:"PATCH",body:JSON.stringify({status:v})})).ok)throw new Error("Gagal memperbarui status");K("Status berhasil diubah"),g()}catch(k){N(k.message)}})}),document.body.appendChild(d)}function C(b=null){let n=!!b,o=n?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",d=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${a.map(m=>`
      <option value="${m.id}" ${b&&b.branch_id===m.id?"selected":""}>
        ${m.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${b?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${b?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${b&&!["AEON","Triple S"].includes(b.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${b?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${b?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${b?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${b?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${b?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${b?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${b?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${b?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${b?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${b?.notes||""}</textarea>
        </div>
      </form>
    `;J({title:o,content:d,confirmText:n?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(m,l)=>{let c=m.querySelector("#migration-form");if(!c.checkValidity()){c.reportValidity();return}let v=new FormData(c),k={branch_id:parseInt(v.get("branch_id"),10),vendor:v.get("vendor"),target_month:v.get("target_month"),status:v.get("status"),actual_date:v.get("actual_date")||null,pic_vendor:v.get("pic_vendor")||null,notes:v.get("notes")||null};try{let T=n?`/api/os-migrations/${b.id}`:"/api/os-migrations",P=await M(T,{method:n?"PUT":"POST",body:JSON.stringify(k)});if(!P.ok){let _=await P.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan data")}K(n?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),l(),g()}catch(T){N(T.message)}}})}let w=document.getElementById("view-board-btn"),y=document.getElementById("view-table-btn");w.addEventListener("click",()=>{t="board",w.style.background="#fff",w.classList.remove("btn-ghost"),y.style.background="transparent",y.classList.add("btn-ghost"),f()}),y.addEventListener("click",()=>{t="table",y.style.background="#fff",y.classList.remove("btn-ghost"),w.style.background="transparent",w.classList.add("btn-ghost"),f()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(b=>{let n=document.getElementById(b);n&&(n.addEventListener("input",f),n.addEventListener("change",f))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{C()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let b=r();if(!b.length){N("Tidak ada data untuk diexport");return}let n=b.map((o,s)=>({No:s+1,"Kode Cabang":o.branch_code,"Nama Cabang":o.branch_name,"Cabang Lengkap":o.branch_full_name,"Vendor OS":o.vendor,"Target Bulan":qt[o.target_month]||o.target_month,Status:Ve[o.status]?.label||o.status,"Tanggal Realisasi":o.actual_date||"","Jml Personil FC":o.employee_count||0,"PIC Vendor":o.pic_vendor||"",Catatan:o.notes||""}));Y(n,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),K("File Excel berhasil diunduh")}),g()}V();le();ie();var Qa={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Wa(e){e.innerHTML=`
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>\u23F1\uFE0F</span> Master DAC (Daily Activity Cleaning)
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Standar Acuan Alur Rutinitas Pembersihan Harian Facility Care (Shift Pagi & Shift Siang)
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-timeline-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:6px 12px; cursor:pointer;">
            \u23F1\uFE0F Alur Timeline
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            \u{1F4CB} Tabel Standar
          </button>
        </div>
        <button id="btn-print-dac" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;" title="Cetak Format SOP untuk Ruang Janitor">
          \u{1F5A8}\uFE0F Cetak SOP
        </button>
        <button id="btn-export-dac" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        <button id="btn-add-dac" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Aktivitas
        </button>
      </div>
    </div>

    <!-- Shift Switcher Tabs -->
    <div style="display:flex; gap:8px; margin-bottom:20px; border-bottom:2px solid #E2E8F0; padding-bottom:8px;">
      <button class="shift-tab-btn active" data-shift="pagi" style="border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        \u{1F305} Shift Pagi (08:00 - 17:00) <span class="badge" style="background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-pagi">31</span>
      </button>
      <button class="shift-tab-btn" data-shift="siang" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        \u{1F307} Shift Siang (13:00 - 22:00) <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-siang">33</span>
      </button>
      <button class="shift-tab-btn" data-shift="all" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        \u{1F4D1} Semua Shift <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-all">64</span>
      </button>
    </div>

    <!-- Summary KPI Cards -->
    <div id="dac-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:24px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="dac-search" class="form-control" placeholder="Cari lokasi, aktivitas, jam..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
          </div>
          <select id="dac-category-filter" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:160px;">
            <option value="">Semua Kategori</option>
          </select>
          <select id="dac-period-filter" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:150px;">
            <option value="">Semua Waktu</option>
            <option value="morning">Sesi Awal (08:00 - 12:00 / 13:00 - 17:00)</option>
            <option value="afternoon">Sesi Lanjutan (12:00 - 17:00 / 17:00 - 22:00)</option>
          </select>
        </div>
        <div id="dac-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="dac-view-container"></div>
  `;let t="pagi",i="timeline",a=[],p=Z(),h=p&&["superadmin","admin","manager","spv"].includes(p.role);async function g(){try{let b=await M("/api/dac?all=1");if(!b.ok)throw new Error("Gagal memuat standar DAC");a=b.data?.data||b.data||[];let n=a.filter(m=>m.shift==="pagi").length,o=a.filter(m=>m.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=n,document.getElementById("badge-count-siang").textContent=o,document.getElementById("badge-count-all").textContent=a.length;let s=document.getElementById("dac-category-filter"),d=Array.from(new Set(a.map(m=>m.category))).sort();s.innerHTML='<option value="">Semua Kategori</option>'+d.map(m=>`<option value="${m}">${m}</option>`).join(""),r(),f()}catch(b){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${b.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function u(){let b=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),n=document.getElementById("dac-category-filter")?.value||"",o=document.getElementById("dac-period-filter")?.value||"";return a.filter(s=>{if(t!=="all"&&s.shift!==t||n&&s.category!==n)return!1;if(o==="morning"){let d=parseInt(s.start_time.split(":")[0],10);if(s.shift==="pagi"&&d>=12||s.shift==="siang"&&d>=17)return!1}else if(o==="afternoon"){let d=parseInt(s.start_time.split(":")[0],10);if(s.shift==="pagi"&&d<12||s.shift==="siang"&&d<17)return!1}return!(b&&!`${s.start_time} ${s.end_time} ${s.category} ${s.location} ${s.activity}`.toLowerCase().includes(b))})}function r(){let b=document.getElementById("dac-summary-cards");if(!b)return;let n=a.filter(l=>l.shift==="pagi"),o=a.filter(l=>l.shift==="siang"),s=t==="pagi"?n:t==="siang"?o:a,d=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",m=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";b.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${m})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${s.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${d}</div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">9 Jam Operasional (Break 1 Jam)</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Cakupan Operasional</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">12 <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Kategori</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Tindakan, Toilet, Area Umum, Medis, dll.</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #F59E0B; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#B45309; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Poin Wajib Harian</div>
        <div style="font-size:1.1rem; font-weight:700; color:#B45309; line-height:1.4;">Check Sheet & Sampah Medis</div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Kontrol mobilitas bertahap & sterilisasi</div>
      </div>
    `}function f(){let b=u(),n=document.getElementById("dac-stats-text");n&&(n.innerHTML=`Menampilkan <strong>${b.length}</strong> aktivitas`);let o=document.getElementById("dac-view-container");i==="timeline"?D(o,b):F(o,b)}function D(b,n){if(!n.length){b.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}b.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${n.map(o=>{let s=Qa[o.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},d=o.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",m=o.shift==="pagi"?"#FEF3C7":"#E0E7FF",l=o.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${o.start_time} - ${o.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${o.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${m}; color:${l}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${d}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${s.bg}; color:${s.color}; border:1px solid ${s.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${o.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${o.location}
                    </span>
                  </div>
                  ${h?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${o.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${o.id}" data-name="${o.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${s.color}; margin-top:8px;">
                  ${o.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,S(b)}function F(b,n){if(!n.length){b.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}b.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                ${t==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${h?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${n.map(o=>{let s=Qa[o.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${o.item_no}</td>
                    ${t==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${o.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${o.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${o.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${o.start_time} - ${o.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${s.bg}; color:${s.color}; border:1px solid ${s.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${o.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${o.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${o.activity}
                    </td>
                    ${h?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${o.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${o.id}" data-name="${o.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,S(b)}function S(b){b.querySelectorAll(".edit-dac-btn").forEach(n=>{n.addEventListener("click",()=>{let o=n.dataset.id,s=a.find(d=>String(d.id)===String(o));s&&E(s)})}),b.querySelectorAll(".delete-dac-btn").forEach(n=>{n.addEventListener("click",()=>{let o=n.dataset.id,s=n.dataset.name||"item";W(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${s}</strong>?`,async()=>{try{if(!(await M(`/api/dac/${o}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");K("Aktivitas DAC berhasil dihapus"),g()}catch(d){N(d.message)}})})})}function E(b=null){let n=!!b,o=n?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",s=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],d=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${b?.shift==="pagi"||!b&&t==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${b?.shift==="siang"||!b&&t==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${b?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${b?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${b?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${b?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${s.map(m=>`<option value="${m}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${b?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${b?.activity||""}</textarea>
        </div>
      </form>
    `;J({title:o,content:d,confirmText:n?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(m,l)=>{let c=m.querySelector("#dac-form");if(!c.checkValidity()){c.reportValidity();return}let v=new FormData(c),k={shift:v.get("shift"),item_no:v.get("item_no")?parseInt(v.get("item_no"),10):void 0,start_time:v.get("start_time"),end_time:v.get("end_time"),category:v.get("category"),location:v.get("location"),activity:v.get("activity")};try{let T=n?`/api/dac/${b.id}`:"/api/dac",P=await M(T,{method:n?"PUT":"POST",body:JSON.stringify(k)});if(!P.ok){let _=await P.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan aktivitas")}K(n?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),l(),g()}catch(T){N(T.message)}}})}function C(){let b=u();if(!b.length){N("Tidak ada data untuk dicetak");return}let n=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",o=window.open("","_blank");if(!o){N("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}o.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${n}</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; margin: 20px; color: #1e293b; font-size: 11pt; }
          .header { text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 12px; margin-bottom: 20px; }
          .title { font-size: 16pt; font-weight: bold; margin: 0; color: #0f172a; }
          .subtitle { font-size: 11pt; color: #475569; margin: 4px 0 0 0; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #cbd5e1; padding: 7px 10px; text-align: left; vertical-align: top; }
          th { background-color: #f1f5f9; font-weight: bold; font-size: 10pt; text-transform: uppercase; }
          .jam { font-family: monospace; font-weight: bold; white-space: nowrap; }
          .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 8pt; font-weight: bold; }
          @media print {
            body { margin: 10mm; }
            button { display: none; }
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="title">FACILITY CARE MANAGEMENT SYSTEM (FCMS)</div>
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${n}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:35px; text-align:center;">NO</th>
              ${t==="all"?'<th style="width:55px;">SHIFT</th>':""}
              <th style="width:105px;">JAM</th>
              <th style="width:130px;">KATEGORI</th>
              <th style="width:180px;">LOKASI / AREA</th>
              <th>RINCIAN AKTIVITAS</th>
            </tr>
          </thead>
          <tbody>
            ${b.map(s=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${s.item_no}</td>
                ${t==="all"?`<td>${s.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${s.start_time} - ${s.end_time}</td>
                <td><strong>${s.category}</strong></td>
                <td><strong>${s.location}</strong></td>
                <td>${s.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),o.document.close(),o.focus(),setTimeout(()=>{o.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(b=>{b.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(n=>{n.style.background="#F1F5F9",n.style.color="var(--text-2)"}),b.style.background="#2563EB",b.style.color="#fff",t=b.dataset.shift,r(),f()})});let w=document.getElementById("view-timeline-btn"),y=document.getElementById("view-table-btn");w.addEventListener("click",()=>{i="timeline",w.style.background="#fff",w.classList.remove("btn-ghost"),y.style.background="transparent",y.classList.add("btn-ghost"),f()}),y.addEventListener("click",()=>{i="table",y.style.background="#fff",y.classList.remove("btn-ghost"),w.style.background="transparent",w.classList.add("btn-ghost"),f()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(b=>{let n=document.getElementById(b);n&&(n.addEventListener("input",f),n.addEventListener("change",f))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{E()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{C()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let b=u();if(!b.length){N("Tidak ada data untuk diexport");return}let n=b.map(o=>({No:o.item_no,Shift:o.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":o.start_time,"Jam Selesai":o.end_time,Kategori:o.category,"Lokasi / Area":o.location,"Rincian Aktivitas":o.activity}));Y(n,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),K("File Excel berhasil diunduh")}),g()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],a=parseInt(t[2],10),p=parseInt(t[3],10);if(a>12&&p<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],p=parseInt(i[2],10),h=a[parseInt(i[1],10)-1];return`${p} ${h} ${i[0]}`}return t};function re(e){return async t=>{if(!Ke()){$e("/login");return}return e(t)}}var Qe=null;function qn(){Qe&&clearInterval(Qe);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),p=document.getElementById("header-clock-time"),h=document.getElementById("header-clock-date");p&&(p.textContent=i),h&&(h.textContent=a)};e(),Qe=setInterval(e,1e3)}async function On(){try{let e=await M("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(a,p)=>{let h=document.getElementById(a);h&&(h.textContent=p>0?p:"",h.style.display=p>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var je=[];async function Rn(){try{let e=await M("/api/dashboard/notifications");if(!e.ok)return;je=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=je.length>0?"block":"none",t.textContent=je.length)}catch{}}function Hn(){if(!je.length){J({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${je.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;J({title:`Notifikasi (${je.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function Xa(){let e=Z(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
    <div class="app-layout">
      <!-- Sidebar dark -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <span class="logo-icon-wrap">\u{1F3E5}</span>
            <span class="logo-text">FC<strong>MS</strong></span>
          </div>
          <button class="sidebar-close" id="sidebar-close" aria-label="Close">\u2715</button>
        </div>

        <nav class="sidebar-nav" id="sidebar-nav">

          <!-- Utama -->
          <div class="nav-section">
            <span class="nav-section-label">UTAMA</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
              </span>
              <span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
              </span>
              <span class="nav-label">Kalender</span>
            </a>
          </div>

          <!-- SDM -->
          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="9" cy="7" r="4"/><path d="M3 21v-2a4 4 0 014-4h4a4 4 0 014 4v2"/><circle cx="19" cy="7" r="2"/><path d="M23 21v-1a3 3 0 00-3-3"/></svg>
              </span>
              <span class="nav-label">Master Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
              </span>
              <span class="nav-label">Data Kontrak</span>
              <span class="nav-badge" id="badge-contracts"></span>
            </a>
            <a href="#/sp" class="nav-item" data-route="/sp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
              </span>
              <span class="nav-label">Data Sp</span>
            </a>
            <a href="#/mutasi" class="nav-item" data-route="/mutasi">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Data Mutasi</span>
            </a>
            <a href="#/overtime" class="nav-item" data-route="/overtime">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Data Lembur</span>
            </a>
          </div>

          <!-- Operasional -->
          <div class="nav-section">
            <span class="nav-section-label">OPERASIONAL</span>
            <a href="#/timeline" class="nav-item" data-route="/timeline">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Time Line</span>
              <span class="nav-badge" id="badge-schedule"></span>
            </a>
            <a href="#/os-migrations" class="nav-item" data-route="/os-migrations">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Migrasi OS</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Permasalahan</span>
              <span class="nav-badge badge-danger" id="badge-issues"></span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
              </span>
              <span class="nav-label">Training</span>
            </a>
          </div>

          <!-- Laporan -->
          <div class="nav-section">
            <span class="nav-section-label">LAPORAN</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </span>
              <span class="nav-label">Report Inspeksi Hygiene 2026</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
              </span>
              <span class="nav-label">Report GCDC 2026</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
              </span>
              <span class="nav-label">Report Fogging 2026</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
              </span>
              <span class="nav-label">Rekap Laporan Basecamp</span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop-fc" class="nav-item" data-route="/sop-fc">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
              </span>
              <span class="nav-label">SOP Facility Care</span>
            </a>
            <a href="#/sop-security" class="nav-item" data-route="/sop-security">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
              </span>
              <span class="nav-label">SOP Security</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Document</span>
            </a>
            <a href="#/hygiene-standards" class="nav-item" data-route="/hygiene-standards">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
              </span>
              <span class="nav-label">Master Hygiene</span>
            </a>
            <a href="#/dac" class="nav-item" data-route="/dac">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Master DAC</span>
            </a>
          </div>

          <!-- Admin -->
          ${e&&(e.role==="superadmin"||e.role==="admin")?`
          <div class="nav-section">
            <span class="nav-section-label">ADMIN</span>
            <a href="#/users" class="nav-item" data-route="/users">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
              </span>
              <span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
            <a href="#/audit-logs" class="nav-item" data-route="/audit-logs">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
              </span>
              <span class="nav-label">Audit Log</span>
            </a>
            <a href="#/settings/import" class="nav-item" data-route="/settings/import">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              </span>
              <span class="nav-label">Import Data Awal</span>
            </a>
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar">${(e?.full_name||e?.username||"U")[0].toUpperCase()}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${e?.full_name||e?.username||"Guest"}</div>
              <div class="sidebar-user-role">${e?.role||"Viewer"}</div>
            </div>
          </div>
          <button class="sidebar-logout" id="logout-btn">
            <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
            Keluar
          </button>
        </div>
      </aside>

      <!-- Mobile overlay -->
      <div class="sidebar-overlay" id="sidebar-overlay"></div>

      <!-- Main wrapper -->
      <div class="main-wrapper">
        <!-- Topbar -->
        <header class="topbar" id="topbar">
          <div class="topbar-left">
            <button class="topbar-menu-btn" id="topbar-menu-btn" aria-label="Menu">
              <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
            </button>
            <div class="topbar-welcome">
              <div class="topbar-greeting">
                <span class="topbar-greeting-time">${(()=>{let f=new Date().getHours();return f>=4&&f<11?"Selamat Pagi":f>=11&&f<15?"Selamat Siang":f>=15&&f<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
              </div>
              <div class="topbar-subtitle">
                Ringkasan Operasional FCMS Hari Ini
              </div>
            </div>
          </div>

          <div class="topbar-center" id="topbar-clock">
            <div class="header-clock">
              <div class="header-clock-time" id="header-clock-time">00:00:00</div>
              <div class="header-clock-date" id="header-clock-date">Memuat...</div>
            </div>
          </div>

          <div class="topbar-right">
            <button class="topbar-icon-btn" id="btn-fullscreen" title="Fullscreen" aria-label="Fullscreen">
              <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
            </button>
            <button class="topbar-icon-btn notif-btn" id="btn-notif" title="Notifikasi" aria-label="Notifikasi">
              <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg>
              <span class="notif-dot" id="notif-dot" style="display:none"></span>
            </button>
            <a href="#/profile" class="topbar-user-btn" title="Profil">
              <img src="https://ui-avatars.com/api/?name=${encodeURIComponent(e?.full_name||e?.username||"Guest")}&background=2563EB&color=fff&bold=true" class="topbar-avatar" alt="Avatar" />
              <div class="topbar-user-text">
                <span class="topbar-user-name">${e?.full_name||e?.username||"Guest"}</span>
                <span class="topbar-user-role-mini">${e?.role||"Viewer"}</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--gray-400)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),p=document.getElementById("topbar-menu-btn"),h=document.getElementById("sidebar-close"),g=()=>{i.classList.add("open"),a.classList.add("show")},u=()=>{i.classList.remove("open"),a.classList.remove("show")};p?.addEventListener("click",g),h?.addEventListener("click",u),a?.addEventListener("click",u),document.querySelectorAll(".nav-item").forEach(f=>f.addEventListener("click",u)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(f=>{f.dataset.route!=="/overtime"&&(f.style.display="none")}),document.querySelectorAll(".nav-section").forEach(f=>{Array.from(f.querySelectorAll(".nav-item")).some(F=>F.style.display!=="none")||(f.style.display="none")}));function r(){let f=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(S=>{let E=S.dataset.route;S.classList.toggle("active",f===E||E!=="/dashboard"&&f.startsWith(E))});let D=document.getElementById("topbar-title"),F=document.querySelector(".nav-item.active .nav-label");D&&F&&(D.textContent=F.textContent)}window.addEventListener("hashchange",r),r(),qn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await M("/api/auth/logout",{method:"POST"}),qe(),Qe&&clearInterval(Qe),$e("/login")}),On(),Rn(),document.getElementById("btn-notif")?.addEventListener("click",f=>{f.preventDefault(),Hn()})}async function Un(){te("/login",({main:t})=>Xt(t)),te("/dashboard",re(({main:t})=>Jt(t))),te("/calendar",re(({main:t})=>Ta(t))),te("/employees",re(({main:t,params:i})=>ea(t,i))),te("/contracts",re(({main:t,params:i})=>ot(t,i))),te("/sp",re(({main:t})=>qa(t))),te("/mutasi",re(({main:t})=>Oa(t))),te("/overtime",re(({main:t})=>Ja(t))),te("/timeline",re(({main:t,params:i})=>aa(t,i))),te("/os-migrations",re(({main:t})=>Va(t))),te("/issues",re(({main:t,params:i})=>ra(t,i))),te("/one-on-one",re(({main:t,params:i})=>oa(t,i))),te("/training",re(({main:t})=>da(t))),te("/relievers",re(({main:t,params:i})=>ca(t,i))),te("/reports/inspection",re(({main:t})=>ua(t))),te("/reports/cleaning",re(({main:t})=>ga(t))),te("/reports/fogging",re(({main:t})=>ma(t))),te("/reports/basecamp",re(({main:t})=>ba(t))),te("/reports/supply",re(({main:t})=>Aa(t,"supply"))),te("/sop",()=>{window.location.hash="#/sop-fc"}),te("/sop-fc",re(({main:t})=>ya(t))),te("/sop-security",re(({main:t})=>Sa(t))),te("/checklist",re(({main:t})=>$a(t))),te("/hygiene-standards",re(({main:t})=>_a(t))),te("/dac",re(({main:t})=>Wa(t))),te("/users",re(({main:t})=>Ba(t))),te("/branches",re(({main:t})=>Da(t))),te("/profile",re(({main:t})=>Pa(t))),te("/settings/import",re(({main:t})=>ja(t))),te("/audit-logs",re(({main:t})=>Ya(t)));let e=Ke();if(!e&&window.location.hash!=="#/login"&&$e("/login"),e){let t=await M("/api/auth/me");if(t.ok){Oe(t.data.data),Xa();let i=Z(),a=window.location.hash;i&&i.role==="input_lembur"&&(a==="#/dashboard"||a===""||a==="#/")&&$e("/overtime")}else qe(),$e("/login")}window.addEventListener("fm:login",()=>{Xa();let t=Z();t&&t.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),Ht()}Un();
