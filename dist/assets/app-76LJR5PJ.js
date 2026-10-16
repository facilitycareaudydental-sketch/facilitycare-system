var Ya=Object.defineProperty;var gt=(e,t)=>()=>(e&&(t=e(e=0)),t);var bt=(e,t)=>{for(var i in t)Ya(e,i,{get:t[i],enumerable:!0})};var Ht={};bt(Ht,{API:()=>qt,CLIENT_SIDE_MAX_ROWS:()=>Te,IS_DEVELOPMENT:()=>Ke,apiFetch:()=>I,clearToken:()=>Oe,getToken:()=>je,getUser:()=>ne,setToken:()=>ht,setUser:()=>Re});function je(){return localStorage.getItem("fm_token")}function ht(e){localStorage.setItem("fm_token",e)}function Oe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ne(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Re(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function I(e,t={}){let i=je(),n={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let r=`cb=${Date.now()}`,b=e.includes("?")?"&":"?",p=`${qt}${e}${b}${r}`,l=await fetch(p,{...t,headers:n}),a;try{let u=await l.text();try{a=JSON.parse(u)}catch{a={error:`Server Error (${l.status}): ${u.substring(0,80)}...`}}}catch{a={error:"Gagal membaca respon dari server"}}return l.status===401&&(Oe(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:a}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var Ke,Wa,qt,Te,V=gt(()=>{Ke=!1,Wa="https://fm-operations-api.facilitycare-audydental.workers.dev",qt=Wa,Te=1e4});var vt={};bt(vt,{confirmDialog:()=>ee,createModal:()=>Y});function Y({title:e,content:t,onConfirm:i,onCancel:n,confirmText:r="Simpan",cancelText:b="Batal",size:p="md",confirmClass:l="btn-primary"}){let a={sm:"400px",md:"560px",lg:"720px",xl:"900px"},u=document.createElement("div");u.className="modal-overlay",u.innerHTML=`
    <div class="modal" style="max-width:${a[p]||a.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${b}</button>
        ${i?`<button class="btn ${l} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&u.querySelector(".modal-body").appendChild(t);let F=()=>{u.classList.remove("show"),setTimeout(()=>u.remove(),250)};return u.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),F()}),u.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),F()}),i&&u.querySelector(".modal-confirm").addEventListener("click",()=>i(u,F)),u.addEventListener("click",$=>{$.target===u&&(n&&n(),F())}),document.body.appendChild(u),requestAnimationFrame(()=>u.classList.add("show")),{overlay:u,close:F}}function ee(e,t,i="Konfirmasi"){return Y({title:i,content:`<p>${e}</p>`,onConfirm:(n,r)=>{t(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var me=gt(()=>{});var ye={};bt(ye,{downloadExcel:()=>J,parseExcel:()=>Je,renderExcelButtons:()=>Xa});function Je(e){return new Promise((t,i)=>{let n=new FileReader;n.onload=r=>{try{let b=new Uint8Array(r.target.result),p=XLSX.read(b,{type:"array"}),l=p.SheetNames[0],a=p.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${p.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let u=XLSX.utils.decode_range(a["!ref"]||"A1:A1"),F=u.e.r-u.s.r+1,$=u.e.c-u.s.c+1;console.log(`Total Rows (including empty): ${F}`),console.log(`Total Columns: ${$}`);let S=[];for(let k=u.s.c;k<=u.e.c;++k){let w=a[XLSX.utils.encode_cell({c:k,r:u.s.r})];w&&w.v&&S.push(w.v)}console.log(`Headers Found: ${S.join(", ")}`),console.log("---------------------------");let v=XLSX.utils.sheet_to_json(a,{defval:""});Object.defineProperty(v,"__worksheet",{value:a,enumerable:!1}),Object.defineProperty(v,"__headers",{value:S,enumerable:!1}),t(v)}catch(b){i(b)}},n.onerror=r=>i(r),n.readAsArrayBuffer(e)})}function J(e,t){try{let i=XLSX.utils.json_to_sheet(e),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,i,"Data"),XLSX.writeFile(n,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Xa(e){return`
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
  `}var X=gt(()=>{});V();var ft={},Qe=null;function Z(e,t){ft[e]=t}function $e(e){window.location.hash=e}function Jt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...n]=t.split("?"),r=ft[i];if(!r){for(let[p,l]of Object.entries(ft))if(p.endsWith("/*")&&i.startsWith(p.slice(0,-2))){r=l;break}}Qe&&(Qe(),Qe=null);let b=document.getElementById("main-content");if(b&&(b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let p=new URLSearchParams(n.join("?")),l=i.split("/").filter(Boolean),a=await r({path:i,params:p,segments:l,main:b});a&&(Qe=a)}else{let p=b||document.getElementById("app");p&&(p.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var qe;function Qa(){return qe||(qe=document.createElement("div"),qe.id="toast-container",document.body.appendChild(qe)),qe}function yt(e,t="info",i=3500){let n=Qa(),r=document.createElement("div");r.className=`toast toast-${t}`;let b={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${b[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,n.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var K=e=>yt(e,"success"),N=e=>yt(e,"error");var xt=e=>yt(e,"info");me();V();V();me();function Xe({columns:e,data:t,onEdit:i,onDelete:n,onView:r,actions:b=[],emptyText:p="Tidak ada data",bulkSelect:l=null}){let a=document.createElement("div");if(a.className="table-wrapper",!t||t.length===0)return a.innerHTML=`<div class="empty-state"><p>${p}</p></div>`,a;let u=document.createElement("table");u.className="data-table";let F=document.createElement("thead"),$=document.createElement("tr");if(l){let v=document.createElement("th");v.style.width="40px",v.style.textAlign="center";let k=document.createElement("input");k.type="checkbox",k.id="select-all-checkbox",k.title="Pilih semua",k.addEventListener("change",()=>{t.forEach(w=>{k.checked?l.selectedIds.add(w.id):l.selectedIds.delete(w.id)}),a.querySelectorAll(".row-checkbox").forEach(w=>w.checked=k.checked),l.onToggle()}),v.appendChild(k),$.appendChild(v)}if(e.forEach(v=>{let k=document.createElement("th");k.textContent=v.label,v.width&&(k.style.width=v.width),$.appendChild(k)}),i||n||r||b.length>0){let v=document.createElement("th");v.textContent="Aksi",v.style.width="120px",$.appendChild(v)}F.appendChild($),u.appendChild(F);let S=document.createElement("tbody");return t.forEach(v=>{let k=document.createElement("tr");if(l){let w=document.createElement("td");w.style.textAlign="center",w.style.width="40px";let f=document.createElement("input");f.type="checkbox",f.className="row-checkbox",f.checked=l.selectedIds.has(v.id),f.addEventListener("change",()=>{if(f.checked)l.selectedIds.add(v.id);else{l.selectedIds.delete(v.id);let m=document.getElementById("select-all-checkbox");m&&(m.checked=!1)}l.onToggle()}),w.appendChild(f),k.appendChild(w)}if(e.forEach(w=>{let f=document.createElement("td");if(w.render){let m=w.render(v[w.key],v);m instanceof HTMLElement?f.appendChild(m):f.innerHTML=m||""}else f.textContent=v[w.key]!==null&&v[w.key]!==void 0&&v[w.key]!==""?v[w.key]:"";w.nowrap&&(f.style.whiteSpace="nowrap"),k.appendChild(f)}),i||n||r||b.length>0){let w=document.createElement("td");w.className="actions-cell";let f=document.createElement("div");if(f.className="btn-group",r){let m=document.createElement("button");m.className="btn btn-xs btn-ghost",m.innerHTML="\u{1F441}",m.title="Lihat",m.addEventListener("click",()=>r(v)),f.appendChild(m)}if(i){let m=document.createElement("button");m.className="btn btn-xs btn-secondary",m.innerHTML="\u270F\uFE0F",m.title="Edit",m.addEventListener("click",()=>i(v)),f.appendChild(m)}b.forEach(m=>{let o=document.createElement("button");o.className=`btn btn-xs ${m.class||"btn-ghost"}`,o.innerHTML=m.icon||m.label,o.title=m.label,o.addEventListener("click",()=>m.handler(v)),f.appendChild(o)}),w.appendChild(f),k.appendChild(w)}S.appendChild(k)}),u.appendChild(S),a.appendChild(u),a}function Ze({page:e,pages:t,total:i,limit:n,onPage:r}){if(t<=1)return null;let b=document.createElement("div");b.className="pagination";let p=document.createElement("span");p.className="pagination-info",p.textContent=`Total: ${i} data`,b.appendChild(p);let l=document.createElement("div");l.className="pagination-btns";let a=($,S,v=!1,k=!1)=>{let w=document.createElement("button");w.className=`btn btn-sm ${k?"btn-primary":"btn-ghost"} pagination-btn`,w.textContent=$,w.disabled=v,w.addEventListener("click",()=>r(S)),l.appendChild(w)};a("\xAB",1,e===1),a("\u2039",e-1,e===1);let u=Math.max(1,e-2),F=Math.min(t,e+2);for(let $=u;$<=F;$++)a($,$,!1,$===e);return a("\u203A",e+1,e===t),a("\xBB",t,e===t),b.appendChild(l),b}me();function He(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${He(t.fields)}</div>`;let i=t.required?"required":"",n=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(t.type){case"textarea":r=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let p=(t.options||[]).map(v=>{let k=typeof v=="object"?v.value:v,w=typeof v=="object"?v.label:v,f=t.value==k?"selected":"";return`<option value="${k}" ${f}>${w}</option>`}).join("");r=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${p}</select>`;break;case"combobox":let l=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,a=(t.options||[]).map(v=>{let k=typeof v=="object"?v.value:v,w=typeof v=="object"?v.label||v.value||"":v||"";return(w==="undefined"||w==="[object Object]"||w==="null")&&(w=""),w?`<option value="${w}"></option>`:""}).join(""),u=t.value||"";if(t.value){let v=(t.options||[]).find(k=>(typeof k=="object"?k.value:k)==t.value);if(v){let k=typeof v=="object"?v.label||v.value||"":v||"";k&&k!=="undefined"&&k!=="[object Object]"&&k!=="null"&&(u=k)}}r=`
          <input type="text" name="${t.name}" list="${l}" class="form-control" value="${u}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${l}">${a}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let F=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";r=`<input type="date" name="${t.name}" class="form-control" value="${F}" ${i}>`;break;case"month":let $=t.value?String(t.value).slice(0,7):"";r=`<input type="month" name="${t.name}" class="form-control" value="${$}" ${i}>`;break;case"number":let S=t.readonly?"readonly":"";r=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${S}>`;break;case"email":r=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let b=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${n}${r}${b}</div>`}).join("")}function et(e){let t={},i=new FormData(e);for(let[n,r]of i.entries())t[n]=r===""?null:r;return e.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(t[n.name]=null)}),t}function tt(e,t){t&&Object.entries(t).forEach(([i,n])=>{let r=e.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!n:r.type==="date"&&n&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(n):r.type==="month"&&n?r.value=String(n).slice(0,7):r.value=n??""))})}X();var Ce={},at={on(e,t){Ce[e]||(Ce[e]=new Set),Ce[e].add(t)},off(e,t){Ce[e]&&Ce[e].delete(t)},emit(e,t){Ce[e]&&Ce[e].forEach(i=>{try{i(t)}catch(n){console.warn("[calendarBus] Handler error:",n)}})},clear(){Object.keys(Ce).forEach(e=>delete Ce[e])}},Za=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function be(e){if(!e){at.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");at.emit("data:changed",{module:t,relevant:Za.has(t)})}function W({container:e,title:t,icon:i,apiPath:n,columns:r,formFields:b,filterFields:p,defaultFilters:l={},itemLabel:a="Data",canCreate:u=!0,canEdit:F=!0,canDelete:$=!0,onBeforeSubmit:S,onAfterLoad:v,onDataLoaded:k,extraActions:w=[],initialSearch:f="",exportOptions:m=null,bulkDelete:o=!1,paginationMode:d="server"}){let h=ne();h&&h.role==="viewer"&&(u=!1,F=!1,$=!1,o=!1,m=null),h&&h.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(u=!1,F=!1,$=!1,o=!1,m=null),h&&h.role==="input_lembur"&&(n==="/api/overtime"?($=!1,o=!1,m=null):(u=!1,F=!1,$=!1,o=!1,m=null));let c=1,y={...l};f&&(y.search=f);let s=new Set;e.innerHTML=`
    ${o?`
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
        ${u?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
        ${m?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${m.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${m.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${m.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${m.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${p&&p.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${p.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${y.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${p.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${y[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${p.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${y[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function g(){let x=document.getElementById("bulk-toolbar");if(!x)return;let T=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),q=document.getElementById("btn-bulk-cancel");T.textContent=`${s.size} item dipilih`,s.size>0?(x.style.display="flex",A.disabled=!1,q.disabled=!1):(x.style.display="none",A.disabled=!0,q.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let G=document.querySelectorAll(".row-checkbox");if(G.length>0){let j=[...G].every(ie=>ie.checked),R=[...G].some(ie=>ie.checked);O.checked=j,O.indeterminate=R&&!j}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{s.clear(),document.querySelectorAll(".row-checkbox").forEach(T=>T.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),g()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(s.size===0)return;let x=[...s],T=document.createElement("div");T.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",T.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(T),T.querySelector("#bulk-cancel-btn").addEventListener("click",()=>T.remove()),T.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=T.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let q=await I(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});T.remove(),q.ok?(K(`${x.length} ${a} berhasil dihapus.`),s.clear(),g(),be(n),_()):N(q.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),C;E?.addEventListener("input",x=>{clearTimeout(C),C=setTimeout(()=>{y.search=x.target.value,c=1,s.clear(),g(),_()},400)}),p?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",T=>{y[x.name]=T.target.value;let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value=T.target.value),c=1,s.clear(),g(),_()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",T=>{y[x.name]=T.target.value;let A=document.getElementById(`filter-${x.name}`);A&&(A.value=T.target.value),c=1,s.clear(),g(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{y={},E&&(E.value=""),p?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),c=1,s.clear(),g(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{y={},E&&(E.value=""),p?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),c=1,s.clear(),g(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>M(null)),m&&document.addEventListener("click",function(x){let T=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");T&&A&&!A.contains(x.target)&&!T.contains(x.target)&&T.classList.remove("show-aksi-menu")});let B=document.getElementById("btn-mobile-filter"),P=document.getElementById("filter-options-wrapper"),D=document.getElementById("btn-close-filter-sheet");if(B&&P&&(B.addEventListener("click",x=>{x.preventDefault(),P.classList.add("sheet-open")}),D&&D.addEventListener("click",x=>{x.preventDefault(),P.classList.remove("sheet-open")})),m){document.getElementById(`btn-export-${m.moduleName}`)?.addEventListener("click",async T=>{let A=T.target,q=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await m.onExport()}catch{N("Gagal export data")}finally{A.innerHTML=q,A.disabled=!1}}),document.getElementById(`btn-template-${m.moduleName}`)?.addEventListener("click",()=>{m.onTemplate()});let x=document.getElementById(`input-import-${m.moduleName}`);x?.addEventListener("change",async T=>{let A=T.target.files[0];if(!A)return;let q=document.getElementById(`label-import-${m.moduleName}`),O=q?q.querySelector(".import-text"):null,G=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),q&&(q.style.pointerEvents="none"),x.disabled=!0;try{let j=await Je(A);if(j.length===0)throw new Error("File kosong atau format salah");await m.onImport(j),K("Import berhasil!"),be(n),_()}catch(j){N(j.message||"Gagal import data")}finally{O&&(O.innerText=G),q&&(q.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){g();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let T=d==="client",A=T?1:c,q=T?Te:20,O=new URLSearchParams({page:A,limit:q,...Object.fromEntries(Object.entries(y).filter(([,U])=>U))}),G=await I(`${n}?${O}`);if(!G.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${G.data?.error||"Error"}</p></div>`;return}let j=G.data?.data||G.data||[],R=G.data?.pagination,ie=j.length;if(T){j=k(j);let U=j.length,re=20,oe=Math.ceil(U/re);c>oe&&oe>0&&(c=oe);let H=(c-1)*re,pe=c*re;j=j.slice(H,pe),R={page:c,limit:re,total:U,pages:oe}}!1,v&&v(j);let de=Xe({columns:r,data:j,onEdit:F?U=>M(U):null,actions:w.map(U=>({...U,handler:re=>U.handler(re,_)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:o?{selectedIds:s,onToggle:g}:null});x.innerHTML="",x.appendChild(de);let ge=document.getElementById("pagination-container");if(ge&&(ge.innerHTML="",R&&R.pages>1)){let U=Ze({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:re=>{c=re,_()}});U&&ge.appendChild(U)}}function L(x){let T=typeof b=="function"?b(x):b;return He(T)}function M(x){let T=!!x,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=L(x),T){let O=typeof b=="function"?b(x):b;tt(A,x)}let{close:q}=Y({title:T?`Edit ${a}`:`Tambah ${a}`,content:A,size:"lg",confirmText:T?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(O,G)=>{if(!A.reportValidity())return;let j=O.querySelector(".modal-confirm");j.disabled=!0,j.textContent="Menyimpan...";let R=et(A),ie=typeof b=="function"?b(x):b,de=async oe=>{for(let H of oe)if(H.type==="row")await de(H.fields);else if(H.type==="combobox"&&R[H.name]){let pe=R[H.name],fe=(H.options||[]).find(se=>{let ue=String(typeof se=="object"?se.value:se),mt=String(typeof se=="object"?se.label:se);return ue===pe||mt===pe});if(fe)R[H.name]=typeof fe=="object"?fe.value:fe;else if(H.createApi){let se={};se[H.createApi.field]=pe,H.createApi.extra&&Object.assign(se,H.createApi.extra);let ue=await I(H.createApi.path,{method:"POST",body:JSON.stringify(se)});if(ue.ok&&ue.data?.id)R[H.name]=ue.data.id;else if(ue.ok&&!ue.data?.id)R[H.name]=pe;else throw new Error(`Gagal membuat master data: ${ue.data?.error||"Unknown error"}`)}}};try{await de(ie)}catch(oe){N(oe.message),j.disabled=!1,j.textContent=T?"Simpan Perubahan":`Tambah ${a}`;return}S&&(R=await S(R,x));let ge=T?"PUT":"POST",U=T?`${n}/${x.id}`:n,re=await I(U,{method:ge,body:JSON.stringify(R)});re.ok?(K(T?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),G(),be(n),_()):(N(re.data?.error||"Gagal menyimpan data."),j.disabled=!1,j.textContent=T?"Simpan Perubahan":`Tambah ${a}`)}})}function z(x){ee(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let T=await I(`${n}/${x.id}`,{method:"DELETE"});T.ok?(K(`${a} berhasil dihapus.`),be(n),_()):N(T.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return _(),_}V();V();var Ae=null,nt=null;async function _e(e=!1){if(Ae&&!e)return console.log("Employees Raw (Cache Hit)",Ae.slice(0,5)),Ae;let t=await I(`/api/employees?limit=${Te}&status=Aktif`);return Ae=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Ae.slice(0,5)),Ae}async function ce(e=!1){let i=(await _e(e)).map(n=>({value:n.label,label:n.label}));return console.log("Employee Options",i.slice(0,5)),i}async function Q(e=!1){return nt&&!e||(nt=((await I("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),nt}function ae(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function kt(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function De(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function wt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function xe(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}X();V();X();V();X();V();X();function it(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let r=new Date(t.value),b=new Date(i.value),p=Math.floor((b-r)/864e5);n.value=isNaN(p)?"":p}else n.value=""}}document.body.removeEventListener("input",it);document.body.addEventListener("input",it);document.body.removeEventListener("change",it);document.body.addEventListener("change",it);V();var Fe={};function Le(e){if(Fe[e]){try{Fe[e].destroy()}catch{}delete Fe[e]}}function en(){Object.keys(Fe).forEach(Le)}var he=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},Be=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function rt(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let n=Math.max(0,Math.round(he(t))),r=e.textContent?e.textContent.replace(/[^\d]/g,""):"",b=r!==""&&parseInt(r)||0;if(n===b){e.textContent=n.toLocaleString("id-ID"),e._animFrame=null;return}let p=Date.now(),l=()=>{let a=Math.min((Date.now()-p)/i,1),u=1-Math.pow(1-a,3);e.textContent=Math.round(b+u*(n-b)).toLocaleString("id-ID"),a<1?e._animFrame=requestAnimationFrame(l):(e.textContent=n.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(l)}var tn={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},an=e=>{let t=Be(e,"\u2014");return`<span class="status-pill ${tn[t]||"pill-neutral"}">${t}</span>`};var ke={family:"Inter",size:11},Ee="#94A3B8",Pe="#F1F5F9",St=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],nn=()=>window.innerWidth<768;function ot(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:nn()?"bottom":"top",labels:{font:ke,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ke,titleFont:{...ke,weight:"700"}}},scales:{x:{grid:{color:Pe},ticks:{font:ke,color:Ee,maxRotation:0}},y:{grid:{color:Pe},ticks:{font:ke,color:Ee},beginAtZero:!0}},...e}}var rn=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),on=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Gt(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ve(e,t,i=8e3){try{let n=new AbortController,r=setTimeout(()=>n.abort(),i),b=await I(e,{signal:n.signal}).catch(()=>null);if(clearTimeout(r),!b||!b.ok)return t;let p=b.data;return p?p.data!==void 0?p.data??t:p:t}catch{return t}}function sn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(n=>{let r=document.getElementById(n);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(n=>{let r=document.getElementById(n);if(r&&r.style.display==="none"){r.style.display="block";let b=r.parentElement;if(b&&!b.querySelector(".chart-empty")){let p=document.createElement("div");p.className="chart-empty",p.textContent="Belum ada data",r.style.display="none",b.appendChild(p)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Yt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Wt({}),["table-contracts","table-issues"].forEach(n=>{let r=document.getElementById(n);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Vt(e){en(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${rn()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${on()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Gt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Gt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ut(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let n=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=n);let b=document.getElementById("skel-jadwal"),p=document.getElementById("chart-jadwal");b&&(b.style.display="block",b.style.position="absolute"),p&&(p.style.display="none");let l=await ve(`/api/dashboard/schedule-chart?year=${n}`,{},8e3);try{Qt(l)}catch(a){console.warn("ScheduleChart render:",a),we("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let n=i.target.value,r=n?`/api/dashboard/inspection-bar?month=${n}`:"/api/dashboard/inspection-bar",b=document.getElementById("skel-insp"),p=document.getElementById("chart-insp");b&&(b.style.display="block",b.style.position="absolute"),p&&(p.style.display="none");let l=await ve(r,{},8e3);try{Xt(l)}catch(a){console.warn("InspBar render:",a),we("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>sn(),5e3),await Ut(e)}async function Ut(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,n,r,b,p,l,a]=await Promise.all([ve("/api/dashboard/kpi",{},8e3),ve("/api/dashboard/issues-trend",{},8e3),ve("/api/dashboard/issues-summary",{},8e3),ve("/api/dashboard/stats",{},8e3),ve("/api/dashboard/calendar",[],8e3),ve("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ve(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ve("/api/dashboard/dropdown-aggregates",{},8e3)]),u=document.getElementById("filter-insp-month"),F=u?u.value:"",$=F?`/api/dashboard/inspection-bar?month=${F}`:"/api/dashboard/inspection-bar",S=await ve($,{},8e3);window.dashboardAggregates=a||{};try{Yt(t)}catch(v){console.warn("KPI render:",v)}try{Wt(t)}catch(v){console.warn("MiniStats render:",v)}try{Qt(l)}catch(v){console.warn("ScheduleChart render:",v),we("skel-jadwal","chart-jadwal")}try{ln(Array.isArray(n?.by_category)?n.by_category:[])}catch(v){console.warn("Donut render:",v),we("skel-donut","chart-donut")}try{dn(i)}catch(v){console.warn("Trend render:",v),we("skel-trend","chart-trend")}try{Xt(S)}catch(v){console.warn("InspBar render:",v),we("skel-insp","chart-insp")}try{let v=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];cn(v)}catch(v){console.warn("IssuesTable render:",v)}try{let v=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];pn(p)}catch(v){console.warn("ContractsTable render:",v)}try{un(Array.isArray(b)?b:[])}catch(v){console.warn("Agenda render:",v)}try{mn()}catch(v){console.warn("Quick Actions render:",v)}}function Yt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,n=window.dashboardAggregates||{},r=a=>{let u=window.dashboardAggregates?.schedule_quarter_details?.[a];return u?`Insp: ${u.insp} \u2022 GC: ${u.gc} \u2022 DC: ${u.dc} \u2022 Fog: ${u.fog}`:""},b=a=>{let u=window.dashboardAggregates?.schedule_quarter_details?.[a];return u?`Rincian ${a}: Inspeksi: ${u.insp} | General Cleaning: ${u.gc} | Deep Cleaning: ${u.dc} | Fogging: ${u.fog}`:`Total Kegiatan ${a}`},p=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:n.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:r(i),subId:"mini-jadwal-detail",tooltip:b(i)}];t.innerHTML=p.map(a=>{let u=a.val!==void 0?he(a.val,0):he(e[a.key]?.current,0);return a.dropdown?`
        <div class="kpi-card ${a.color}" ${a.id?`id="${a.id}"`:""} ${a.tooltip?`title="${a.tooltip}"`:""} data-href="${a.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${u}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${u}</div>
                ${a.dropdown}
              </div>
              <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
              <div class="kpi-subtitle" ${a.subId?`id="${a.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${a.tooltip||""}">${a.sub}</div>
            </div>
          </div>
        </div>`:`
      <a href="${a.href}" class="kpi-card ${a.color}" ${a.id?`id="${a.id}"`:""} ${a.tooltip?`title="${a.tooltip}"`:""} style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="kpi-value" data-target="${u}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${u}</div>
              ${a.dropdown?a.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" ${a.subId?`id="${a.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${a.tooltip||""}">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(a=>{a.addEventListener("click",u=>{u.target.closest("select")||a.dataset.href&&(window.location.hash=a.dataset.href)}),a.addEventListener("keydown",u=>{u.key==="Enter"&&!u.target.closest("select")&&a.dataset.href&&(window.location.hash=a.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(a=>{rt(a,parseInt(a.dataset.target)||0)});let l=document.getElementById("dash-jadwal-period");if(l){let a=u=>{let $=(window.dashboardAggregates?.schedule_by_quarter||{})[u]||0,S=document.querySelector("#mini-jadwal .kpi-value");S&&(S.dataset.target=$,rt(S,$,400));let v=document.getElementById("mini-jadwal-detail");v&&(v.textContent=r(u),v.title=b(u));let k=document.getElementById("mini-jadwal");if(k){let w=`#/timeline?dash_filter=period_${u.toLowerCase()}`;k.tagName==="A"?k.href=w:k.dataset.href=w,k.title=b(u)}};l.addEventListener("change",u=>a(u.target.value))}}function Wt(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),n=String(new Date().getMonth()+1).padStart(2,"0"),r=`${i}-${n}`,b=u=>`
    <select id="${u}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${i}-01" ${r===`${i}-01`?"selected":""}>Jan</option>
      <option value="${i}-02" ${r===`${i}-02`?"selected":""}>Feb</option>
      <option value="${i}-03" ${r===`${i}-03`?"selected":""}>Mar</option>
      <option value="${i}-04" ${r===`${i}-04`?"selected":""}>Apr</option>
      <option value="${i}-05" ${r===`${i}-05`?"selected":""}>Mei</option>
      <option value="${i}-06" ${r===`${i}-06`?"selected":""}>Jun</option>
      <option value="${i}-07" ${r===`${i}-07`?"selected":""}>Jul</option>
      <option value="${i}-08" ${r===`${i}-08`?"selected":""}>Agu</option>
      <option value="${i}-09" ${r===`${i}-09`?"selected":""}>Sep</option>
      <option value="${i}-10" ${r===`${i}-10`?"selected":""}>Okt</option>
      <option value="${i}-11" ${r===`${i}-11`?"selected":""}>Nov</option>
      <option value="${i}-12" ${r===`${i}-12`?"selected":""}>Des</option>
    </select>
  `,p=window.dashboardAggregates||{},l=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:b("dash-inspeksi-month"),val:p.inspeksi_by_month?.[r]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${r}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:b("dash-gcdc-month"),val:p.gcdc_by_month?.[r]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${r}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:b("dash-reliefer-month"),val:p.relievers_by_month?.[r]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${r}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:b("dash-fogging-month"),val:p.fogging_by_month?.[r]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${r}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=l.map(u=>u.dropdown?`
        <div class="mini-stat ${u.color}" style="cursor:pointer;" id="${u.id||""}" data-href="${u.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${u.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${he(u.val)}">0</div>
              ${u.dropdown}
            </div>
            <div class="mini-stat-text">${u.label}</div>
            ${u.extra||""}
          </div>
        </div>`:`
      <a href="${u.href}" class="mini-stat ${u.color}" style="text-decoration:none" id="${u.id||""}">
        <div class="mini-stat-icon">${u.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${he(u.val)}">0</div>
            ${u.dropdown?u.dropdown:""}
          </div>
          <div class="mini-stat-text">${u.label}</div>
          ${u.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(u=>{u.addEventListener("click",F=>{F.target.closest("select")||u.dataset.href&&(window.location.hash=u.dataset.href)}),u.addEventListener("keydown",F=>{F.key==="Enter"&&!F.target.closest("select")&&u.dataset.href&&(window.location.hash=u.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(u=>rt(u,parseInt(u.dataset.target)||0,700));let a=(u,F,$,S)=>{let v=document.getElementById(u);if(v){let k=w=>{let m=(window.dashboardAggregates?.[$]||{})[w]||0,o=document.querySelector(`#${F} .mini-stat-value`);o&&(o.dataset.target=m,rt(o,m,400));let d=document.getElementById(F);if(d){let h=`${S}&month=${w}`;d.tagName==="A"?d.href=h:d.dataset.href=h}};v.addEventListener("change",w=>k(w.target.value))}};a("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),a("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),a("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),a("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function ln(e){we("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Le("donut");let n=(e||[]).filter(a=>he(a.count)>0);if(!n.length){Ge(t,"Belum ada data permasalahan");return}let r=n.map(a=>`${Be(a.category,"Lainnya")}`),b=n.map(a=>he(a.count)),p=b.reduce((a,u)=>a+u,0);i.innerHTML=n.map((a,u)=>{let F=St[u%St.length],$=p>0?Math.round(a.count/p*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${F}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${a.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${$}%)</span></div>
          <div class="donut-legend-label">${r[u]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(a){let u=a.width,F=a.height,$=a.ctx;$.restore();let S=(F/80).toFixed(2);$.font="bold "+S+"em Inter",$.textBaseline="middle",$.fillStyle="#1E293B";let v=p.toString(),k=Math.round((u-$.measureText(v).width)/2),w=F/2;$.fillText(v,k,w-4),$.font="600 "+(S*.35).toFixed(2)+"em Inter",$.fillStyle="#64748B";let f="Total",m=Math.round((u-$.measureText(f).width)/2);$.fillText(f,m,w+10),$.save()}};Fe.donut=new Chart(t,{type:"doughnut",data:{labels:r,datasets:[{data:b,backgroundColor:St,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ke,titleFont:{...ke,weight:"700"},callbacks:{label:a=>` ${a.label}: ${a.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function dn(e){we("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Le("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],n=(e.labels||[]).map(p=>{if(!p||typeof p!="string")return"";try{let[l,a]=p.split("-");return(i[Number(a)-1]||a)+" "+String(l).slice(-2)}catch{return p}}),r=(e.open||[]).map(p=>he(p)),b=(e.closed||[]).map(p=>he(p));if(!n.length){Ge(t,"Belum ada data trend");return}Fe.trend=new Chart(t,{type:"line",data:{labels:n,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:b,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:ot({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:45,autoSkip:!0}},y:{grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:Ee},beginAtZero:!0}}})})}function Qt(e){we("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Le("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(a=>Array.isArray(a)&&a.some(u=>u>0))){Ge(t,"Belum ada data jadwal");return}let r=e["Inspeksi Hygiene"]||Array(12).fill(0),b=e["General Cleaning"]||Array(12).fill(0),p=e["Deep Cleaning"]||Array(12).fill(0),l=e.Fogging||Array(12).fill(0);Fe.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:b,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:p,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:ot({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:Ee},min:0}}})})}function Xt(e){we("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Le("inspBar"),e=e||{};let i=e.labels||[],n=(e.fc||[]).map(b=>he(b)),r=(e.spv||[]).map(b=>he(b));if(!i.length){Ge(t,"Belum ada data inspeksi");return}Fe.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:ot({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Ee,maxRotation:45,minRotation:45}},y:{grid:{color:Pe},ticks:{font:ke,color:Ee},min:0,max:100}}})})}function pn(e){we("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Le("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},n=(e.labels||[]).map(p=>{let l=p.split("-")[1];return i[l]||p}),r=(e.data||[]).map(p=>he(p));if(!n.length){Ge(t,"Belum ada data");return}let b=t.getContext("2d");Fe.contractMiniBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:ot({onClick:(p,l)=>{if(l&&l.length>0){let a=l[0].index,u=(e.labels||[])[a];u&&(window.location.hash="#/contracts?month_expiry="+u)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Ee,maxRotation:0,autoSkip:!1}},y:{grid:{color:Pe,borderDash:[4,4],drawBorder:!1},ticks:{font:ke,color:Ee,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function cn(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(n=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${an(n.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Be(n.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Be(n.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function un(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,n=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,b=(e||[]).filter(p=>(p.event_date||"").startsWith(n)).slice(0,10);if(!b.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${b.map(p=>{let l="#3B82F6",a="#EFF6FF",u="Agenda",F=(p.title||"").toLowerCase();return F.includes("inspeksi")?(l="#10B981",a="#ECFDF5",u="Inspeksi"):F.includes("cleaning")||F.includes("gcdc")?(l="#3B82F6",a="#EFF6FF",u="Cleaning"):F.includes("reliefer")?(l="#F59E0B",a="#FFFBEB",u="Reliefer"):F.includes("fogging")&&(l="#8B5CF6",a="#F5F3FF",u="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(p.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Be(p.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Be(p.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${a};color:${l}">${u}</div>
        </div>
      `}).join("")}
    </div>
  `}function mn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function we(e,t){let i=document.getElementById(e),n=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),n){n.style.display="block";let r=n.parentElement;if(r){let b=r.querySelector(".chart-empty");b&&b.remove()}}}function Ge(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=t,i.appendChild(r)}}V();async function Zt(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),n=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),b=document.getElementById("login-password");r?.addEventListener("click",()=>{let p=b.type==="text";b.type=p?"password":"text",r.style.color=p?"":"var(--primary)"}),t?.addEventListener("submit",async p=>{p.preventDefault(),i.style.display="none";let l=t.username.value.trim(),a=t.password.value;if(!l||!a){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let u=await I("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:a})});u.ok&&u.data.success?(ht(u.data.data.token),Re(u.data.data.user),K("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=u.data.error||"Username atau password salah.",i.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}V();X();async function gn(){return await Q()}function st(e){if(!e||e==="-"||String(e).trim()==="")return"-";let t=String(e).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],n=t.match(/^(\d{4})-(\d{1,2})/);if(n){let b=n[1],p=parseInt(n[2],10)-1;if(p>=0&&p<12)return`${i[p]} ${b}`}let r=t.split(/[\/\-\.]/);if(r.length===2&&r[1].length===4){let b=parseInt(r[0],10)-1;if(b>=0&&b<12)return`${i[b]} ${r[1]}`}if(r.length===3&&r[2].length===4){let b=parseInt(r[1],10)-1;if(b>=0&&b<12)return`${i[b]} ${r[2]}`}for(let b=0;b<12;b++)if(t.toLowerCase().includes(i[b].toLowerCase()))return t;return t}function bn(e,t){let i=String(e.status||"").toLowerCase();return t==="active"?i==="aktif":t==="reliefer"?e.division==="FC - RELIEFER"&&i==="aktif":!1}async function ea(e,t){let i=await gn(),n=t?t.get("dash_filter"):null;W({container:e,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>n?r.filter(b=>bn(b,n)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>De(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>ae(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await I(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let b=r.data.data.map(p=>({"Nama Lengkap":p.full_name,Cabang:p.branch_name||"",Divisi:p.division||"","No. HP":p.phone||"","Tgl Masuk":p.join_date||"",Status:p.status||"","Target Pindah OS":st(p.target_pindah_os)==="-"?"":st(p.target_pindah_os),"Target Selesai OS":st(p.target_selesai_os)==="-"?"":st(p.target_selesai_os)}));J(b,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async r=>{let b=a=>{if(!a)return null;let u=String(a||"").toLowerCase(),F=i.find($=>String($.label||"").toLowerCase()===u);return F?F.value:null},p=r.map(a=>({full_name:String(a["Nama Lengkap"]||"").trim(),branch_id:b(String(a.Cabang||"").trim()),division:String(a.Divisi||a["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(a["No. HP"]||a["No. Hp"]||"").trim(),join_date:String(a["Tgl Masuk"]||a["Tanggal Masuk"]||"").trim(),status:String(a.Status||"").trim(),target_pindah_os:String(a["Target Pindah OS"]||"").trim(),target_selesai_os:String(a["Target Selesai OS"]||"").trim(),notes:String(a.Catatan||"").trim()})).filter(a=>a.full_name),l=await I("/api/import/employees",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}V();X();var Ct=[],ta=[];async function hn(){Ct=await Q(),ta=await _e()}var Et=async e=>{let t=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(V(),Ht))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let b=r.data?.data||r.data||[],p=Array.isArray(b)?b:[];if(t=t.concat(p),p.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return t};async function lt(e,t){await hn(),W({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:n=>n,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:n=>De(n)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:n=>window.formatDate(n)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:n=>!n||String(n).startsWith("2099")?"Tetap / PKWTT":window.formatDate(n)},{key:"days_remaining",label:"Sisa Kontrak",render:(n,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':kt(n)},{key:"status",label:"Status",render:n=>ae(n)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ct},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:n=>(n.end_date||(n.end_date="2099-12-31"),n),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let n=document.createElement("button");n.id="btn-find-missing",n.className="btn btn-ghost",n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.style.marginLeft="8px",n.style.color="#EF4444",n.style.border="1px solid currentColor",n.onclick=async()=>{n.innerHTML="\u231B Mencari...",n.disabled=!0;try{let[b,p]=await Promise.all([Et("/api/employees?status=Aktif"),Et("/api/contracts")]);if(b.length>0){let l=p.filter($=>$.status==="Aktif"),a=new Set(l.map($=>$.employee_id)),u=b.filter($=>!a.has($.id)),F=`<p style="margin-bottom:12px">Data yang terbaca: <b>${b.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${u.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;u.forEach($=>{let S=p.filter(k=>k.employee_id===$.id),v='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(S.length>0){let k=S[0];v=`Pernah ada kontrak (Status: <b style="color:#EF4444">${k.status}</b>, Selesai: ${window.formatDate(k.end_date)})`}F+=`<li style="margin-bottom:8px"><b>${$.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${$.branch_name||"-"} | ${v}</span></li>`}),F+="</ul>",Promise.resolve().then(()=>(me(),vt)).then($=>$.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:F,cancelText:"Tutup"}))}}catch(b){console.error(b)}n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(n)}},formFields:n=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:ta,value:n?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:Ct,value:n?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:n?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:n?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:n?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:n?.end_date&&!String(n.end_date).startsWith("2099")?n.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:n?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:n?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let n=await I(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let r=n.data.data.map(b=>({"Nama Lengkap":b.employee_name,Cabang:b.branch_name||"","Div / Bagian":b.division||"","Tanggal Mulai":b.start_date||"","Tanggal Selesai":b.end_date&&String(b.end_date).startsWith("2099")?"":b.end_date||"","Sisa Kontrak":b.end_date&&String(b.end_date).startsWith("2099")?"Tetap":b.days_remaining!==null&&b.days_remaining!==void 0?`${b.days_remaining} Hari`:"",Status:b.status||""}));J(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async n=>{let[r,b]=await Promise.all([I("/api/branches?limit=10000"),Et("/api/employees")]),p=r.data?.data||[],l=b||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((f,m)=>{console.log(`${m+1}. ID: ${f.id}, Name: ${f.full_name}, Status: ${f.status}`)}));let a=f=>{if(!f)return null;let m=String(f||"").replace(/\s+/g," ").toLowerCase().trim(),o=p.find(d=>String(d.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===m||String(d.code||"").replace(/\s+/g," ").toLowerCase().trim()===m||String(d.name||"").replace(/\s+/g," ").toLowerCase().trim()===m);return o?o.id:null},u=(f,m)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${m}`),console.log(`Nama dari Excel : "${f}"`),!f)return console.log("Alasan gagal mapping : Nama kosong"),null;let o=String(f||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${o}"`),console.log(`Jumlah employee di database : ${l.length}`);let d=l.find(h=>String(h.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===o);return d?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${d.id}`),d.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},F=f=>{if(!f)return"";if(f instanceof Date&&!isNaN(f.getTime()))return f.toISOString().slice(0,10);let m=String(f).trim();if(/^\d{4,5}(\.\d+)?$/.test(m)){let d=Math.floor(Number(m));if(d>2e4&&d<99999){let h=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let o=m.split(/[\/\-\.]/);if(o.length===3){let[d,h,c]=o.map(y=>y.trim());if(d.length===4&&h.length<=2&&c.length<=2)return`${d}-${h.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&h.length<=2&&d.length<=2)return`${c}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`}return m},$=n.map((f,m)=>{let o=m+2,d=String(f["Nama Lengkap"]||"").trim(),h=f["Tanggal Mulai"],c=F(h);if(!c){let g=n.__worksheet,E=n.__headers||[],C=E.indexOf("Tanggal Mulai"),B="N/A",P="N/A",D="N/A";if(C!==-1&&g&&window.XLSX){let L=window.XLSX.utils.encode_cell({c:C,r:o-1});D=L;let M=g[L];M?(B=M.t||"undefined",P=M.w||"undefined"):B="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let _="Unknown";h==null||h===""?_="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":h instanceof Date&&isNaN(h.getTime())?_="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":_="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${o}`),console.log(`Employee Name : ${d}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${C})`),console.log(`Raw Cell Value : "${h}"`),console.log(`JavaScript Type : ${typeof h}`),console.log(`SheetJS Cell Type : ${B}`),console.log(`SheetJS Formatted Value : "${P}"`),console.log(`Value After Trim : "${String(h||"").trim()}"`),console.log(`Value After Date Parser : "${c}"`),console.log(`Is Empty : ${!h}`),console.log(`Is Invalid Date : ${h instanceof Date?isNaN(h.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${_}`),console.log(`Workbook Sheet : ${g?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${D}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(f,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let y=u(d,o),s=null;return y?c||(s="Tanggal Mulai kosong atau tidak berformat tanggal"):s="Karyawan tidak ditemukan di Database",{isValid:!!(y&&c),invalidReason:s,rowNum:o,data:{employee_id:y,branch_id:a(String(f.Cabang||"").trim()),division:String(f["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:c,end_date:F(f["Tanggal Selesai"])||"2099-12-31",status:String(f.Status||"").trim(),_rawName:d}}}),S=[],v=[];if($.forEach(f=>{f.isValid?S.push(f.data):v.push({rowNum:f.rowNum,name:f.data._rawName,reason:f.invalidReason})}),console.log(`Split Validation - Valid: ${S.length}, Invalid: ${v.length}`),S.length===0){let f=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${n.length}
Valid: 0
Invalid: ${v.length}

Daftar Kegagalan (Contoh):
`;v.slice(0,10).forEach(m=>{f+=`- Row ${m.rowNum} | Nama: ${m.name} | Alasan: ${m.reason}
`}),v.length>10&&(f+=`- ... dan ${v.length-10} lainnya.
`),alert(f);return}let k=await I("/api/contracts/import",{method:"POST",body:JSON.stringify(S)}),w=`IMPORT SUMMARY
======================
`;w+=`Total Baris Excel : ${n.length}
`,w+=`Baris Valid       : ${S.length}
`,w+=`Baris Invalid     : ${v.length}

`,k&&k.data&&k.data.metrics?(w+=`Berhasil INSERT   : ${k.data.metrics.inserted}
`,w+=`Berhasil UPDATE   : ${k.data.metrics.updated}
`):w+=`Berhasil diproses : ${S.length}
`,v.length>0&&(w+=`
DAFTAR DATA DILEWATI:
`,v.forEach(f=>{w+=`- Row ${f.rowNum} | ${f.name} | ${f.reason}
`})),alert(w),typeof lt=="function"&&lt()}}})}V();X();var Ft=[],Ue=[];function fn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(n=>n.period===i))return i;return"Q3"}async function aa(e,t){Ft=await Q();let i=await ce();Ue=["Berlin Ariansyah","Ade Surahman"];let n=f=>f&&!Ue.find(m=>String(typeof m=="object"?m.value:m).toLowerCase()===String(f).toLowerCase())?[...Ue,f]:Ue,r=await I(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),b=f=>{if(!f||f==="-"||String(f).trim()==="")return"";let m=String(f).split("-");return m.length===3&&m[0].length===4?`${m[2]}-${m[1]}-${m[0]}`:f},p=r.data?.data||[],l=fn(p),a=t?t.get("dash_filter"):null,u=new Date,F=`${u.getFullYear()}-${String(u.getMonth()+1).padStart(2,"0")}`,$={},S=t&&t.get("month")?t.get("month"):null;a==="inspeksi"?$={status:"Done",activity_type:"Inspeksi Hygiene",month:S}:a==="gcdc"?$={status:"Done",activity_type:"GCDC",month:S}:a&&a.startsWith("period_")&&($={period:a.replace("period_","").toUpperCase()});let v=new Date().getFullYear(),w=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((f,m)=>{let o=String(m+1).padStart(2,"0");return{value:`${v}-${o}`,label:`${f} ${v}`}});W({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:$,onDataLoaded:f=>f.sort((m,o)=>{let d=m.opening_date?new Date(m.opening_date).getTime():0;return(o.opening_date?new Date(o.opening_date).getTime():0)-d}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:f=>wt(f)},{key:"period",label:"Periode",render:f=>xe(f)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:f=>b(f)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:f=>b(f)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:f=>b(f)},{key:"status",label:"Status",render:f=>ae(f)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Ft},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:w},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ue}],formFields:f=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ft,value:f?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:f?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:f?.period},{name:"pic",label:"PIC",type:"combobox",options:n(f?.pic),value:f?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:f?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:f?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:f?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:f?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:f?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let f=await I(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(f.ok){let m=f.data.data.map(o=>({Cabang:o.branch_name||"",Kegiatan:o.activity_type||"",Periode:o.period||"",PIC:o.pic||"","Tgl Opening":o.opening_date||"","Tgl Target":o.target_date||"","Tgl Selesai":o.completion_date||"",Status:o.status||""}));J(m,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async f=>{let o=(await I("/api/branches?all=1")).data?.data||[],d=s=>{if(!s)return null;let g=String(s||"").toLowerCase().trim(),E=g.match(/^(\d{3})/);if(E){let B=o.find(P=>String(P.code||"")===E[1]);if(B)return B.id}let C=o.find(B=>String(B.full_name||"").toLowerCase()===g||String(B.code||"").toLowerCase()===g||String(B.name||"").toLowerCase()===g);return C?C.id:null},h=s=>{if(s==null||s==="")return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let g=String(s).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let C=Number(g);if(C>2e4&&C<99999){let B=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(B.getTime())?"":B.toISOString().slice(0,10)}}let E=g.split(/[\/\-\.]/);if(E.length===3){let[C,B,P]=E.map(D=>D.trim());if(C.length===4&&B.length<=2&&P.length<=2)return`${C}-${B.padStart(2,"0")}-${P.padStart(2,"0")}`;if(P.length===4&&B.length<=2&&C.length<=2)return`${P}-${B.padStart(2,"0")}-${C.padStart(2,"0")}`}return g},c=f.map(s=>({branch_id:d(String(s.Cabang||"").trim()),branch_name:String(s.Cabang||"").trim(),activity_type:String(s.Kegiatan||"").trim(),period:String(s.Periode||"").trim(),pic:String(s.PIC||s.Pic||"").trim(),opening_date:h(s["Tgl Opening"]||s["Tanggal Opening"]||s["Tgl Openir"]),target_date:h(s["Tgl Target"]||s["Tanggal Target"]),completion_date:h(s["Tgl Selesai"]||s["Tanggal Selesai"]),status:String(s.Status||"").trim(),notes:String(s.Catatan||s.Keterangan||"").trim()})).filter(s=>s.activity_type&&s.period),y=await I("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:c,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return be("schedule"),y.data}}})}V();X();var $t=[],dt=[];function pt(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let r=new Date(t.value),b=new Date(i.value),p=Math.floor((b-r)/864e5);n.value=isNaN(p)?"":p}else n.value=""}}document.body.removeEventListener("input",pt);document.body.addEventListener("input",pt);document.body.removeEventListener("change",pt);document.body.addEventListener("change",pt);function yn(e,t){let i=String(e.status||"").toLowerCase();return t==="open"?i==="open":!1}async function na(e,t){let i=t?t.get("dash_filter"):null;$t=await Q(),dt=await ce();let n=a=>a&&!dt.find(u=>u.value===a)?[...dt,{value:a,label:a}]:dt,r=new Date().getFullYear(),b=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((a,u)=>{let F=String(u+1).padStart(2,"0");return{value:`${r}-${F}`,label:`${a} ${r}`}});W({container:e,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:a=>i?a.filter(u=>yn(u,i)):a,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:a=>window.formatDate(a)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:a=>`<span class="badge badge-secondary">${a}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:a=>`<span title="${a}">${a?.length>50?a.slice(0,50)+"...":a}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:a=>`<span title="${a||""}">${a?.length>40?a.slice(0,40)+"...":a||"-"}</span>`},{key:"status",label:"Status",render:a=>ae(a)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:a=>window.formatDate(a)},{key:"day_count",label:"Day",render:a=>a??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:$t},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:b}],formFields:a=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:a?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:$t,value:a?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:a?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:a?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:a?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:n(a?.employee_name),value:a?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:n(a?.fc_specialist),value:a?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:a?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:a?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:a?.completion_date},{name:"day_count",label:"Day",type:"number",value:a?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let a=await I(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let u=a.data.data.map(F=>({"Tanggal Info":F.report_date||"",Cabang:F.branch_name||"",Kategori:F.category||"","Sumber Laporan":F.source||"",Keluhan:F.complaint||"","Nama FC":F.employee_name||"","FC Spesialis":F.fc_specialist||"",Solusi:F.solution||"",Status:F.status||"","Tanggal Selesai":F.completion_date||"",Day:F.day_count!==null?F.day_count:""}));J(u,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async a=>{let F=(await I("/api/branches?all=1")).data?.data||[],$=k=>{if(!k)return null;let w=String(k||"").toLowerCase(),f=F.find(m=>String(m.full_name||"").toLowerCase()===w||String(m.code||"").toLowerCase()===w||String(m.name||"").toLowerCase()===w);return f?f.id:null},S=a.map(k=>({branch_id:$(String(k.Cabang||"").trim()),report_date:String(k["Tanggal Info"]||k.Tanggal||"").trim(),category:String(k.Kategori||"").trim(),source:String(k["Sumber Laporan"]||k.Sumber||"").trim(),complaint:String(k.Keluhan||"").trim(),employee_name:String(k["Nama FC"]||"").trim(),fc_specialist:String(k["FC Spesialis"]||"").trim(),solution:String(k.Solusi||"").trim(),completion_date:String(k["Tanggal Selesai"]||k["Tgl Selesai"]||"").trim(),status:String(k.Status||"").trim(),day_count:k.Day||k.Hari||null})).filter(k=>k.report_date&&k.complaint&&k.category),v=await I("/api/import/issues",{method:"POST",body:JSON.stringify({rows:S,onDuplicate:"update"})});if(!v.ok)throw new Error(v.data?.error||"Import gagal");return v.data}}})}V();var Ie=[];function xn(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function ia(e,t){let i=t?t.get("dash_filter"):null;Ie=await Q();let n=await ce(),r=["Ade","Berlin"],b=l=>l&&!n.find(a=>a.value===l)?[...n,{value:l,label:l}]:n,p=l=>l&&!r.find(a=>(typeof a=="object"?a.value:a)===l)?[...r,l]:r;W({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>i?l.filter(a=>xn(a,i)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>ae(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Ie},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let a=new URLSearchParams(l||{}).toString(),u=await I(`/api/one-on-one?limit=10000&${a}`);if(u.ok){let F=u.data.data.map(S=>({Tanggal:S.meeting_date||"",Cabang:S.branch_name||"","Nama Karyawan":S.employee_name||"",PIC:S.pic||"",Masalah:S.problem||"",Solusi:S.solution||"",Status:S.status||"","Tgl Selesai":S.completion_date||"",Dokumen:S.document_link||""})),{downloadExcel:$}=await Promise.resolve().then(()=>(X(),ye));$(F,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:a}=await Promise.resolve().then(()=>(X(),ye));a(l,"Template_Import_OneOnOne")},onImport:async l=>{let a=S=>{if(!S)return null;let v=String(S||"").toLowerCase(),k=Ie.find(w=>String(w.label||"").toLowerCase()===v);return k?k.value:null},u=S=>{if(!S)return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let v=String(S).trim();if(/^\d{4,5}$/.test(v)){let w=Number(v);if(w>2e4&&w<99999){let f=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(v))return v.slice(0,10);let k=v.split(/[\/\-\.]/);if(k.length===3){let[w,f,m]=k.map(o=>o.trim());if(w.length===4&&f.length<=2&&m.length<=2)return`${w}-${f.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&f.length<=2&&w.length<=2)return`${m}-${f.padStart(2,"0")}-${w.padStart(2,"0")}`}return v},F=l.map(S=>({meeting_date:u(S.Tanggal),employee_name:String(S["Nama Karyawan"]||"").trim(),branch_id:a(String(S.Cabang||"").trim()),pic:String(S.PIC||"").trim(),problem:String(S.Masalah||"").trim(),solution:String(S.Solusi||"").trim(),status:String(S.Status||"").trim(),completion_date:u(S["Tgl Selesai"]),document_link:String(S.Dokumen||"").trim()})).filter(S=>S.meeting_date&&S.employee_name&&S.branch_id),$=await I("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:F,onDuplicate:"update"})});if(!$.ok)throw new Error($.data?.error||"Import gagal");return $.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Ie.find(a=>a.value==l.branch_id)?[...Ie,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Ie,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:b(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:p(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}V();async function ra(e){let t=await Q(),i=await ce(),n=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=a=>a&&!i.find(u=>u.value===a)?[...i,{value:a,label:a}]:i,b=a=>a&&!n.find(u=>(typeof u=="object"?u.value:u)===a)?[...n,a]:n,p=Array.from({length:5},(a,u)=>String(new Date().getFullYear()-u));W({container:e,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:a=>{try{let u=JSON.parse(a);return Array.isArray(u)?u.join(", "):a||"-"}catch{return a||"-"}}},{key:"score",label:"Nilai",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:p}],exportOptions:{moduleName:"training",onExport:async a=>{let u=new URLSearchParams(a||{}).toString(),F=await I(`/api/training?limit=10000&${u}`);if(F.ok){let $=F.data.data.map(v=>{let k=v.participants||"";try{let w=JSON.parse(k);k=Array.isArray(w)?w.join(", "):k}catch{}return{Tanggal:v.training_date||"",Batch:v.batch||"",Materi:v.subject||"",Cabang:v.branch_name||"",Trainer:v.trainer||"",Peserta:k,Nilai:v.score!==null&&v.score!==void 0?v.score:"",Dokumen:v.document_link||""}}),{downloadExcel:S}=await Promise.resolve().then(()=>(X(),ye));S($,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let a=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:u}=await Promise.resolve().then(()=>(X(),ye));u(a,"Template_Import_Training")},onImport:async a=>{let u=v=>{if(!v)return null;let k=String(v||"").toLowerCase(),w=t.find(f=>String(f.label||"").toLowerCase()===k);return w?w.value:null},F=v=>{if(!v)return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let k=String(v).trim();if(/^\d{4,5}$/.test(k)){let f=Number(k);if(f>2e4&&f<99999){let m=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(k))return k.slice(0,10);let w=k.split(/[\/\-\.]/);if(w.length===3){let[f,m,o]=w.map(d=>d.trim());if(f.length===4&&m.length<=2&&o.length<=2)return`${f}-${m.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&m.length<=2&&f.length<=2)return`${o}-${m.padStart(2,"0")}-${f.padStart(2,"0")}`}return k},$=a.map(v=>({training_date:F(v.Tanggal),batch:String(v.Batch||"").trim(),subject:String(v.Materi||"").trim(),branch_id:u(String(v.Cabang||"").trim()),trainer:String(v.Trainer||"").trim(),participants:String(v.Peserta||"").trim(),score:v.Nilai?Number(v.Nilai):null,document_link:String(v.Dokumen||"").trim()})).filter(v=>v.training_date&&v.subject&&v.branch_id),S=await I("/api/training/import",{method:"POST",body:JSON.stringify($)});if(!S.ok)throw new Error(S.data?.error||"Import gagal")}},formFields:a=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:a?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:a?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:a?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,value:a?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:b(a?.trainer),value:a?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let u=JSON.parse(a?.participants);return Array.isArray(u)?u.join(", "):a?.participants||""}catch{return a?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:a?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],onBeforeSubmit:async a=>a})}V();me();X();function oa({container:e,title:t,icon:i,apiPath:n,columns:r,formFields:b,filterFields:p,defaultFilters:l={},itemLabel:a="Data",canCreate:u=!0,canEdit:F=!0,canDelete:$=!0,onBeforeSubmit:S,onAfterLoad:v,onDataLoaded:k,extraActions:w=[],initialSearch:f="",exportOptions:m=null,bulkDelete:o=!1,paginationMode:d="server"}){let h=ne();h&&h.role==="viewer"&&(u=!1,F=!1,$=!1,o=!1,m=null),h&&h.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(u=!1,F=!1,$=!1,o=!1,m=null),h&&h.role==="input_lembur"&&(n==="/api/overtime"?($=!1,o=!1,m=null):(u=!1,F=!1,$=!1,o=!1,m=null));let c=1,y={...l};f&&(y.search=f);let s=new Set;e.innerHTML=`
    ${o?`
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
        ${u?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
        ${m?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${m.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${m.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${m.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${m.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${p&&p.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${p.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${y.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${p.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${y[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${p.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${y[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function g(){let x=document.getElementById("bulk-toolbar");if(!x)return;let T=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),q=document.getElementById("btn-bulk-cancel");T.textContent=`${s.size} item dipilih`,s.size>0?(x.style.display="flex",A.disabled=!1,q.disabled=!1):(x.style.display="none",A.disabled=!0,q.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let G=document.querySelectorAll(".row-checkbox");if(G.length>0){let j=[...G].every(ie=>ie.checked),R=[...G].some(ie=>ie.checked);O.checked=j,O.indeterminate=R&&!j}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{s.clear(),document.querySelectorAll(".row-checkbox").forEach(T=>T.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),g()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(s.size===0)return;let x=[...s],T=document.createElement("div");T.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",T.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(T),T.querySelector("#bulk-cancel-btn").addEventListener("click",()=>T.remove()),T.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=T.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let q=await I(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});T.remove(),q.ok?(K(`${x.length} ${a} berhasil dihapus.`),s.clear(),g(),be(n),_()):N(q.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),C;E?.addEventListener("input",x=>{clearTimeout(C),C=setTimeout(()=>{y.search=x.target.value,c=1,s.clear(),g(),_()},400)}),p?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",T=>{y[x.name]=T.target.value;let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value=T.target.value),c=1,s.clear(),g(),_()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",T=>{y[x.name]=T.target.value;let A=document.getElementById(`filter-${x.name}`);A&&(A.value=T.target.value),c=1,s.clear(),g(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{y={},E&&(E.value=""),p?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),c=1,s.clear(),g(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{y={},E&&(E.value=""),p?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),c=1,s.clear(),g(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>M(null)),m&&document.addEventListener("click",function(x){let T=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");T&&A&&!A.contains(x.target)&&!T.contains(x.target)&&T.classList.remove("show-aksi-menu")});let B=document.getElementById("btn-mobile-filter"),P=document.getElementById("filter-options-wrapper"),D=document.getElementById("btn-close-filter-sheet");if(B&&P&&(B.addEventListener("click",x=>{x.preventDefault(),P.classList.add("sheet-open")}),D&&D.addEventListener("click",x=>{x.preventDefault(),P.classList.remove("sheet-open")})),m){document.getElementById(`btn-export-${m.moduleName}`)?.addEventListener("click",async T=>{let A=T.target,q=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await m.onExport()}catch{N("Gagal export data")}finally{A.innerHTML=q,A.disabled=!1}}),document.getElementById(`btn-template-${m.moduleName}`)?.addEventListener("click",()=>{m.onTemplate()});let x=document.getElementById(`input-import-${m.moduleName}`);x?.addEventListener("change",async T=>{let A=T.target.files[0];if(!A)return;let q=document.getElementById(`label-import-${m.moduleName}`),O=q?q.querySelector(".import-text"):null,G=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),q&&(q.style.pointerEvents="none"),x.disabled=!0;try{let j=await Je(A);if(j.length===0)throw new Error("File kosong atau format salah");await m.onImport(j),K("Import berhasil!"),be(n),_()}catch(j){N(j.message||"Gagal import data")}finally{O&&(O.innerText=G),q&&(q.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){g();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let T=d==="client",A=T?1:c,q=T?Te:20,O=new URLSearchParams({page:A,limit:q,...Object.fromEntries(Object.entries(y).filter(([,U])=>U))}),G=await I(`${n}?${O}`);if(!G.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${G.data?.error||"Error"}</p></div>`;return}let j=G.data?.data||G.data||[],R=G.data?.pagination,ie=j.length;if(T){j=k(j);let U=j.length,re=20,oe=Math.ceil(U/re);c>oe&&oe>0&&(c=oe);let H=(c-1)*re,pe=c*re;j=j.slice(H,pe),R={page:c,limit:re,total:U,pages:oe}}!1,v&&v(j);let de=Xe({columns:r,data:j,onEdit:F?U=>M(U):null,actions:w.map(U=>({...U,handler:re=>U.handler(re,_)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:o?{selectedIds:s,onToggle:g}:null});x.innerHTML="",x.appendChild(de);let ge=document.getElementById("pagination-container");if(ge&&(ge.innerHTML="",R&&R.pages>1)){let U=Ze({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:re=>{c=re,_()}});U&&ge.appendChild(U)}}function L(x){let T=typeof b=="function"?b(x):b;return He(T)}function M(x){let T=!!x,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=L(x),T){let O=typeof b=="function"?b(x):b;tt(A,x)}let{close:q}=Y({title:T?`Edit ${a}`:`Tambah ${a}`,content:A,size:"lg",confirmText:T?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(O,G)=>{if(!A.reportValidity())return;let j=O.querySelector(".modal-confirm");j.disabled=!0,j.textContent="Menyimpan...";let R=et(A),ie=typeof b=="function"?b(x):b,de=async oe=>{for(let H of oe)if(H.type==="row")await de(H.fields);else if(H.type==="combobox"&&R[H.name]){let pe=R[H.name],fe=(H.options||[]).find(se=>{let ue=String(typeof se=="object"?se.value:se),mt=String(typeof se=="object"?se.label:se);return ue===pe||mt===pe});if(fe)R[H.name]=typeof fe=="object"?fe.value:fe;else if(H.createApi){let se={};se[H.createApi.field]=pe,H.createApi.extra&&Object.assign(se,H.createApi.extra);let ue=await I(H.createApi.path,{method:"POST",body:JSON.stringify(se)});if(ue.ok&&ue.data?.id)R[H.name]=ue.data.id;else if(ue.ok&&!ue.data?.id)R[H.name]=pe;else throw new Error(`Gagal membuat master data: ${ue.data?.error||"Unknown error"}`)}}};try{await de(ie)}catch(oe){N(oe.message),j.disabled=!1,j.textContent=T?"Simpan Perubahan":`Tambah ${a}`;return}S&&(R=await S(R,x));let ge=T?"PUT":"POST",U=T?`${n}/${x.id}`:n,re=await I(U,{method:ge,body:JSON.stringify(R)});re.ok?(K(T?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),G(),be(n),_()):(N(re.data?.error||"Gagal menyimpan data."),j.disabled=!1,j.textContent=T?"Simpan Perubahan":`Tambah ${a}`)}})}function z(x){ee(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let T=await I(`${n}/${x.id}`,{method:"DELETE"});T.ok?(K(`${a} berhasil dihapus.`),be(n),_()):N(T.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return _(),_}V();X();async function sa(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await Q(),n=await ce(),r=t?t.get("dash_filter"):null,b={};if(r==="reliever"){let S=new Date,v=`${S.getFullYear()}-${String(S.getMonth()+1).padStart(2,"0")}`;b={status:"Done",month:t&&t.get("month")?t.get("month"):v}}console.log("RAW",await _e()),console.log("OPTIONS",n);let p=S=>S&&!n.find(v=>v.value===S)?[...n,{value:S,label:S}]:n,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],a=S=>S&&!l.includes(S)?[...l,S]:l,u=new Date().getFullYear(),$=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((S,v)=>{let k=String(v+1).padStart(2,"0");return{value:`${u}-${k}`,label:`${S} ${u}`}});oa({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:b,onDataLoaded:S=>S.sort((v,k)=>{let w=v.backup_date?new Date(v.backup_date).getTime():0;return(k.backup_date?new Date(k.backup_date).getTime():0)-w}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:S=>xe(S)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:S=>window.formatDate(S)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:S=>window.formatDate(S)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:S=>S?`<span class="badge badge-info">${S}</span>`:"-"},{key:"status",label:"Status",render:S=>ae(S)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:$},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:S=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:S?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:S?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:p(S?.original_fc_name),value:S?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:a(S?.reliever_name),value:S?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:S?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:S?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:S?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:S?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:S?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let S=await I(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(S.ok){let v=S.data.data.map(k=>({Cabang:k.branch_name||"","Nama Facility care":k.original_fc_name||"",Periode:k.period||"",Relifer:k.reliever_name||"","Tanggal Back Up":k.backup_date||"","Tanggal Selesai":k.completion_date||"",Keterangan:k.reason||"",Shift:k.shift||"",Status:k.status||""}));v.length===0&&v.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),J(v,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async S=>{let k=(await I("/api/branches?all=1")).data?.data||[],w=o=>{if(!o)return null;let d=String(o||"").toLowerCase(),h=k.find(c=>String(c.full_name||"").toLowerCase()===d||String(c.code||"").toLowerCase()===d||String(c.name||"").toLowerCase()===d);return h?h.id:null},f=S.map(o=>({branch_name:String(o.Cabang||"").trim(),backup_date:String(o["Tanggal Back Up"]||o["Tanggal Backup"]||"").trim(),original_fc_name:String(o["Nama Facility care"]||o["FC Digantikan"]||"").trim(),reliever_name:String(o.Relifer||o.Reliefer||"").trim(),period:String(o.Periode||"").trim(),reason:String(o.Keterangan||"").trim(),shift:String(o.Shift||"").trim(),completion_date:String(o["Tanggal Selesai"]||"").trim(),status:String(o.Status||"").trim()})).filter(o=>o.reliever_name&&o.backup_date),m=await I("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:f,onDuplicate:"update"})});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}V();X();async function la(e){let t=await Q(),i=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r));W({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:n=>xe(n)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"fc_score",label:"Point FC",render:n=>n!=null?`<strong class="${n>=80?"text-success":n>=60?"text-warning":"text-danger"}">${n}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"status",label:"Status",render:n=>ae(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:n?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:n?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:n?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),b=await I(`/api/reports/inspection?limit=10000&${r}`);if(b.ok){let p=b.data.data.map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));J(p,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async n=>{let r=a=>{if(!a)return null;let u=String(a||"").toLowerCase(),F=t.find($=>String($.label||"").toLowerCase()===u);return F?F.value:null},b=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let u=String(a).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let $=Number(u);if($>2e4&&$<99999){let S=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let F=u.split(/[\/\-\.]/);if(F.length===3){let[$,S,v]=F.map(k=>k.trim());if($.length===4&&S.length<=2&&v.length<=2)return`${$}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&S.length<=2&&$.length<=2)return`${v}-${S.padStart(2,"0")}-${$.padStart(2,"0")}`}return u},p=n.map(a=>({branch_id:r(String(a.Cabang||"").trim()),period:String(a.Periode||"").trim(),inspection_date:b(a.Tanggal),fc_score:a["Point FC"]!==void 0&&a["Point FC"]!==""?Number(a["Point FC"]):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==""?Number(a["Point SPV"]):null,status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.period&&a.inspection_date),l=await I("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}V();X();async function da(e){let t=await Q(),i=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r));W({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge ${n==="Deep Cleaning"?"badge-purple":"badge-success"}">${n}</span>`},{key:"period",label:"Periode",render:n=>xe(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>ae(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:n?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:n?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:n?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),b=await I(`/api/reports/cleaning?limit=10000&${r}`);if(b.ok){let p=b.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));J(p,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async n=>{let r=a=>{if(!a)return null;let u=String(a||"").toLowerCase(),F=t.find($=>String($.label||"").toLowerCase()===u);return F?F.value:null},b=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let u=String(a).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let $=Number(u);if($>2e4&&$<99999){let S=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let F=u.split(/[\/\-\.]/);if(F.length===3){let[$,S,v]=F.map(k=>k.trim());if($.length===4&&S.length<=2&&v.length<=2)return`${$}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&S.length<=2&&$.length<=2)return`${v}-${S.padStart(2,"0")}-${$.padStart(2,"0")}`}return u},p=n.map(a=>({branch_id:r(String(a.Cabang||"").trim()),activity_type:String(a.Jenis||a.Kegiatan||"").trim(),period:String(a.Periode||"").trim(),activity_date:b(a.Tanggal),status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.activity_type&&a.period&&a.activity_date),l=await I("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}V();X();async function pa(e,t){let i=await Q(),n=Array.from({length:4},(p,l)=>String(new Date().getFullYear()-l)),r=t?t.get("dash_filter"):null,b={};if(r==="fogging"){let p=new Date,l=String(p.getMonth()+1).padStart(2,"0"),a=String(p.getFullYear()),u=t?t.get("month"):null;u&&u.length===7&&(a=u.split("-")[0],l=u.split("-")[1]),b={status:"Done",month:l,year:a}}W({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:b,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:p=>`<span class="badge badge-warning">${p}</span>`},{key:"period",label:"Periode",render:p=>xe(p)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:p=>window.formatDate(p)},{key:"status",label:"Status",render:p=>ae(p)},{key:"document_link",label:"Dokumen",render:p=>p?`<a href="${p}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:p=>p||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:p=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:p?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:p?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:p?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:p?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:p?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:p?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async p=>{let l=new URLSearchParams(p||{}).toString(),a=await I(`/api/reports/fogging?limit=10000&${l}`);if(a.ok){let u=a.data.data.map(F=>({Cabang:F.branch_name||"",Jenis:F.activity_type||"Fogging",Periode:F.period||"",Tanggal:F.activity_date||"",Status:F.status||"","Link Dokumen":F.document_link||""}));J(u,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async p=>{let l=$=>{if(!$)return null;let S=String($||"").toLowerCase(),v=i.find(k=>String(k.label||"").toLowerCase()===S);return v?v.value:null},a=$=>{if($==null||$==="")return"";if($ instanceof Date&&!isNaN($.getTime()))return $.toISOString().slice(0,10);let S=String($).trim();if(S===""||S==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);if(/^\d{4,5}$/.test(S)){let k=Number(S);if(k>2e4&&k<99999){let w=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let v=S.split(/[\/\-\.]/);if(v.length===3){let[k,w,f]=v.map(m=>m.trim());if(k.length===4&&w.length<=2&&f.length<=2)return`${k}-${w.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&w.length<=2&&k.length<=2)return`${f}-${w.padStart(2,"0")}-${k.padStart(2,"0")}`}return S},u=p.map($=>({branch_id:l(String($.Cabang||"").trim()),activity_type:String($.Jenis||$.Kegiatan||"Fogging").trim(),period:String($.Periode||"").trim(),activity_date:a($.Tanggal),status:String($.Status||"").trim(),document_link:String($["Link Dokumen"]||"").trim(),notes:String($.Catatan||$.Keterangan||"").trim()})).filter($=>$.branch_id&&$.period&&$.activity_date),F=await I("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(u)});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}}})}V();X();async function ca(e){let t=await Q(),i=await ce(),n=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(l,a)=>String(new Date().getFullYear()-a)),b=l=>l&&!i.find(a=>a.value===l)?[...i,{value:l,label:l}]:i,p=l=>l&&!n.find(a=>a.value===l)?[...n,{value:l,label:l}]:n;W({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>ae(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:p(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let a=new URLSearchParams(l||{}).toString(),u=await I(`/api/reports/basecamp?limit=10000&${a}`);if(u.ok){let F=u.data.data.map($=>({"Tgl Info":$.info_date||"",Cabang:$.branch_name||"",Permasalahan:$.problem||"",PIC:$.pic||"","Tgl Done":$.done_date||"",Status:$.status||"",Keterangan:$.notes||""}));J(F,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let a=S=>{if(!S)return null;let v=String(S||"").toLowerCase(),k=t.find(w=>String(w.label||"").toLowerCase()===v);return k?k.value:null},u=S=>{if(S==null||S==="")return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let v=String(S).trim();if(v===""||v==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(v))return v.slice(0,10);if(/^\d{4,5}$/.test(v)){let w=Number(v);if(w>2e4&&w<99999){let f=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let k=v.split(/[\/\-\.]/);if(k.length===3){let[w,f,m]=k.map(o=>o.trim());if(w.length===4&&f.length<=2&&m.length<=2)return`${w}-${f.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&f.length<=2&&w.length<=2)return`${m}-${f.padStart(2,"0")}-${w.padStart(2,"0")}`}return v},F=l.map(S=>({info_date:u(S["Tgl Info"]||S["Tanggal Info"]),branch_id:a(String(S.Cabang||"").trim()),problem:String(S.Permasalahan||"").trim(),pic:String(S.PIC||"").trim(),done_date:u(S["Tgl Done"]||S["Tanggal Done"]),status:String(S.Status||"").trim(),notes:String(S.Keterangan||S.Catatan||"").trim()})).filter(S=>S.info_date&&S.branch_id&&S.problem),$=await I("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(F)});if(!$.ok)throw new Error($.data?.error||"Import gagal");return $.data}}})}me();var vn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],kn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],wn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Dt="fcms_sop_fc_ik_v1",Bt="fcms_sop_fc_tools_v1",Pt="fcms_sop_fc_chem_v1";function ua(){try{let e=localStorage.getItem(Dt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(vn))}function _t(e){localStorage.setItem(Dt,JSON.stringify(e))}function ma(){try{let e=localStorage.getItem(Bt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(kn))}function Tt(e){localStorage.setItem(Bt,JSON.stringify(e))}function ga(){try{let e=localStorage.getItem(Pt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(wn))}function At(e){localStorage.setItem(Pt,JSON.stringify(e))}async function ba(e){let t="ik",i="",n="all",r="tools",b="",p=ua(),l=ma(),a=ga();function u(){e.innerHTML=`
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
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${F(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${$(t==="ik")}">${p.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${F(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${$(t==="tools")}">${l.length+a.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${F(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${F(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,S(),v()}function F(g){return g?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function $(g){return g?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function S(){e.querySelectorAll(".sop-tab-btn").forEach(B=>{B.addEventListener("click",()=>{t=B.dataset.tab,u()})});let g=e.querySelector("#btn-reset-default");g&&g.addEventListener("click",()=>{ee("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Dt),localStorage.removeItem(Bt),localStorage.removeItem(Pt),p=ua(),l=ma(),a=ga(),K("Data SOP Facility Care berhasil di-reset ke standar asli"),u()},"Konfirmasi Reset Standar")});let E=e.querySelector("#btn-print-sop");E&&E.addEventListener("click",()=>{window.print()});let C=e.querySelector("#btn-export-tools");C&&C.addEventListener("click",()=>{y()})}function v(){let g=e.querySelector("#sop-content-area");g&&(t==="ik"?k(g):t==="tools"?m(g):t==="policy"?h(g):t==="links"&&c(g))}function k(g){g.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${p.length})</button>
              <button class="cat-pill ${n==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${p.filter(L=>L.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${n==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${p.filter(L=>L.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${n==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${p.filter(L=>L.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${n==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${p.filter(L=>L.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${n==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${p.filter(L=>L.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${p.length} dari ${p.length} Instruksi Kerja
            </div>
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let E=document.createElement("style");E.innerHTML=`
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
    `,g.appendChild(E);let C=g.querySelector("#ik-cards-grid-container"),B=g.querySelector("#ik-count-label");function P(){let L=p.filter(M=>{let z=n==="all"||M.category===n,x=i.toLowerCase().trim(),T=!x||M.title.toLowerCase().includes(x)||M.target.toLowerCase().includes(x)||M.tools.some(A=>A.toLowerCase().includes(x))||M.steps.some(A=>A.toLowerCase().includes(x));return z&&T});if(B&&(B.textContent=`Menampilkan ${L.length} dari ${p.length} Instruksi Kerja`),L.length===0){C.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}C.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${L.map(M=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${M.bg||"#EFF6FF"}; border-radius:10px;">${M.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${M.id}. ${M.title}
                      </h3>
                      <span class="badge" style="background:${M.bg||"#EFF6FF"}; color:${M.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${M.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${M.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${M.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${M.color||"#2563EB"};">
                  <strong>Area:</strong> ${M.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${M.tools.slice(0,3).map(z=>`<li style="margin-bottom:2px;">${z}</li>`).join("")}
                    ${M.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${M.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${M.chemical_ratio&&M.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${M.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${M.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${M.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,C.querySelectorAll(".btn-view-ik").forEach(M=>{M.addEventListener("click",()=>{let z=parseInt(M.dataset.id,10),x=p.find(T=>T.id===z);x&&w(x)})}),C.querySelectorAll(".btn-edit-ik").forEach(M=>{M.addEventListener("click",()=>{let z=parseInt(M.dataset.id,10),x=p.find(T=>T.id===z);x&&f(x)})}),C.querySelectorAll(".btn-delete-ik").forEach(M=>{M.addEventListener("click",()=>{let z=parseInt(M.dataset.id,10),x=p.find(T=>T.id===z);x&&ee(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${x.title}"?`,()=>{p=p.filter(T=>T.id!==z),_t(p),K("Instruksi Kerja berhasil dihapus"),u()},"Hapus Instruksi Kerja")})})}P();let D=g.querySelector("#ik-search");D&&D.addEventListener("input",L=>{i=L.target.value,P()}),g.querySelectorAll(".cat-pill").forEach(L=>{L.addEventListener("click",()=>{g.querySelectorAll(".cat-pill").forEach(M=>M.classList.remove("active")),L.classList.add("active"),n=L.dataset.cat,P()})});let _=g.querySelector("#btn-add-ik");_&&_.addEventListener("click",()=>{f()})}function w(g){Y({title:`${g.icon||"\u{1F4CB}"} IK-${g.id}: ${g.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${g.bg||"#EFF6FF"}; color:${g.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${g.category}</span>
            <span>ID: IK-${String(g.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${g.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${g.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${g.tools.map(E=>`<li>${E}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${g.chemical_ratio&&g.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${g.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${g.steps.map(E=>`<li style="margin-bottom:6px;">${E}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${g.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${g.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(E,C)=>{C()}})}function f(g=null){let E=!!g,C=g?g.tools.join(`
`):"",B=g?g.steps.join(`
`):"",P=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${g?s(g.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${g?g.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${g?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${g?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${g?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${g?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${g?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${g?s(g.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${g?s(g.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${s(C)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${s(B)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${g?s(g.notes||""):""}</textarea>
        </div>
      </form>
    `;Y({title:E?`\u270F\uFE0F Edit IK-${g.id}: ${g.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:P,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(D,_)=>{let L=D.querySelector("#form-ik-title").value.trim(),M=D.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",z=D.querySelector("#form-ik-cat").value,x=D.querySelector("#form-ik-ratio").value.trim(),T=D.querySelector("#form-ik-target").value.trim(),A=D.querySelector("#form-ik-tools").value.trim(),q=D.querySelector("#form-ik-steps").value.trim(),O=D.querySelector("#form-ik-notes").value.trim();if(!L||!T||!A||!q){N("Harap lengkapi seluruh field wajib (*)");return}let G=A.split(`
`).map(de=>de.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),j=q.split(`
`).map(de=>de.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),ie={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[z]||{color:"#2563EB",bg:"#EFF6FF"};if(E)g.title=L,g.icon=M,g.category=z,g.chemical_ratio=x,g.target=T,g.tools=G,g.steps=j,g.notes=O,g.color=ie.color,g.bg=ie.bg,_t(p),K(`Instruksi Kerja "${L}" berhasil diperbarui`);else{let de=p.length>0?Math.max(...p.map(U=>U.id))+1:1,ge={id:de,title:L,category:z,icon:M,color:ie.color,bg:ie.bg,target:T,chemical_ratio:x||"-",tools:G,steps:j,notes:O};p.push(ge),_t(p),K(`Instruksi Kerja baru IK-${de} berhasil ditambahkan`)}_(),u()}})}function m(g){g.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${r==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${l.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${r==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${a.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${b}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${r==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let E=g.querySelector("#tool-table-container"),C=g.querySelector("#tool-search");function B(){let P=b.toLowerCase().trim();if(r==="tools"){let D=l.filter(_=>!P||_.name.toLowerCase().includes(P)||_.area.toLowerCase().includes(P));E.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${D.length} Item</span>
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
                  ${D.map((_,L)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${L%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||L+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${s(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${s(_.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${s(_.area)}</td>
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
        `,E.querySelectorAll(".btn-edit-tool").forEach(_=>{_.addEventListener("click",()=>{let L=parseInt(_.dataset.no,10),M=l.find(z=>z.no===L);M&&o(M)})}),E.querySelectorAll(".btn-delete-tool").forEach(_=>{_.addEventListener("click",()=>{let L=parseInt(_.dataset.no,10),M=l.find(z=>z.no===L);M&&ee(`Apakah Anda yakin ingin menghapus alat: "${M.name}"?`,()=>{l=l.filter(z=>z.no!==L),l.forEach((z,x)=>{z.no=x+1}),Tt(l),K("Alat kebersihan berhasil dihapus"),u()},"Hapus Alat Kebersihan")})})}else{let D=a.filter(_=>!P||_.name.toLowerCase().includes(P)||_.function.toLowerCase().includes(P));E.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${D.length} Chemical</span>
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
                  ${D.map((_,L)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${L%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||L+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${s(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${s(_.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${s(_.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${s(_.function)}</td>
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
        `,E.querySelectorAll(".btn-edit-chem").forEach(_=>{_.addEventListener("click",()=>{let L=parseInt(_.dataset.no,10),M=a.find(z=>z.no===L);M&&d(M)})}),E.querySelectorAll(".btn-delete-chem").forEach(_=>{_.addEventListener("click",()=>{let L=parseInt(_.dataset.no,10),M=a.find(z=>z.no===L);M&&ee(`Apakah Anda yakin ingin menghapus chemical: "${M.name}"?`,()=>{a=a.filter(z=>z.no!==L),a.forEach((z,x)=>{z.no=x+1}),At(a),K("Bahan chemical berhasil dihapus"),u()},"Hapus Bahan Chemical")})})}}B(),g.querySelector("#view-tools-btn").addEventListener("click",()=>{r="tools",m(g)}),g.querySelector("#view-chem-btn").addEventListener("click",()=>{r="chemical",m(g)}),C&&C.addEventListener("input",P=>{b=P.target.value,B()}),g.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{r==="tools"?o():d()})}function o(g=null){let E=!!g,C=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${g?s(g.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${g?s(g.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${g?s(g.area):""}</textarea>
        </div>
      </form>
    `;Y({title:E?`\u270F\uFE0F Edit Alat: ${g.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:C,confirmText:"Simpan",onConfirm:(B,P)=>{let D=B.querySelector("#form-tool-name").value.trim(),_=B.querySelector("#form-tool-qty").value.trim(),L=B.querySelector("#form-tool-area").value.trim();if(!D||!_||!L){N("Harap lengkapi semua kolom");return}if(E)g.name=D,g.qty=_,g.area=L,Tt(l),K(`Alat "${D}" berhasil diperbarui`);else{let M=l.length>0?Math.max(...l.map(z=>z.no))+1:1;l.push({no:M,name:D,qty:_,area:L}),Tt(l),K(`Alat baru "${D}" berhasil ditambahkan`)}P(),u()}})}function d(g=null){let E=!!g,C=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${g?s(g.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${g?s(g.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${g?s(g.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${g?s(g.function):""}</textarea>
        </div>
      </form>
    `;Y({title:E?`\u270F\uFE0F Edit Chemical: ${g.name}`:"\u2795 Tambah Chemical Baru",content:C,confirmText:"Simpan",onConfirm:(B,P)=>{let D=B.querySelector("#form-chem-name").value.trim(),_=B.querySelector("#form-chem-size").value.trim(),L=B.querySelector("#form-chem-ratio").value.trim(),M=B.querySelector("#form-chem-function").value.trim();if(!D||!_||!L||!M){N("Harap lengkapi semua kolom");return}if(E)g.name=D,g.size=_,g.ratio=L,g.function=M,At(a),K(`Chemical "${D}" berhasil diperbarui`);else{let z=a.length>0?Math.max(...a.map(x=>x.no))+1:1;a.push({no:z,name:D,size:_,ratio:L,function:M}),At(a),K(`Chemical baru "${D}" berhasil ditambahkan`)}P(),u()}})}function h(g){g.innerHTML=`
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
    `}function c(g){g.innerHTML=`
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
    `}async function y(){try{let{downloadExcel:g}=await Promise.resolve().then(()=>(X(),ye)),E=p.map(B=>({"No IK":`IK-${B.id}`,"Judul Pekerjaan":B.title,Kategori:B.category,"Sasaran Target Area":B.target,"Rasio Chemical":B.chemical_ratio||"-","Alat & Bahan":B.tools.join(", "),"Langkah Kerja":B.steps.map((P,D)=>`${D+1}. ${P}`).join(" | "),"Catatan Safety":B.notes||"-"})),C=new Date().toISOString().slice(0,10);g(E,`SOP_Facility_Care_Instruksi_Kerja_${C}`),K("Data Excel Instruksi Kerja berhasil diunduh")}catch(g){console.error(g),N("Gagal mendownload data Excel: "+g.message)}}function s(g){return g?String(g).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}u()}me();var Sn=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],En=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],ha={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},Mt="fcms_sop_sec_ik_v1",zt="fcms_sop_sec_assets_v1";function fa(){try{let e=localStorage.getItem(Mt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Sn))}function Lt(e){localStorage.setItem(Mt,JSON.stringify(e))}function ya(){try{let e=localStorage.getItem(zt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(En))}function It(e){localStorage.setItem(zt,JSON.stringify(e))}async function xa(e){let t="ik",i="",n="all",r="",b="cabang",p=fa(),l=ya();function a(){e.innerHTML=`
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
        <button class="sec-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${u(t==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${F(t==="ik")}">${p.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="zonasi"?"active":""}" data-tab="zonasi" style="${u(t==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${t==="assets"?"active":""}" data-tab="assets" style="${u(t==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${F(t==="assets")}">${l.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${u(t==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${u(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,$(),S()}function u(s){return s?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function F(s){return s?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function $(){e.querySelectorAll(".sec-tab-btn").forEach(C=>{C.addEventListener("click",()=>{t=C.dataset.tab,a()})});let s=e.querySelector("#btn-reset-sec-default");s&&s.addEventListener("click",()=>{ee("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(Mt),localStorage.removeItem(zt),p=fa(),l=ya(),K("Data SOP Security berhasil di-reset ke standar dokumen"),a()},"Reset ke Standar Dokumen")});let g=e.querySelector("#btn-print-sec");g&&g.addEventListener("click",()=>{window.print()});let E=e.querySelector("#btn-export-sec");E&&E.addEventListener("click",()=>{c()})}function S(){let s=e.querySelector("#sec-content-area");s&&(t==="ik"?v(s):t==="zonasi"?f(s):t==="assets"?m(s):t==="policy"?d(s):t==="links"&&h(s))}function v(s){s.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${p.length})</button>
              <button class="sec-cat-pill ${n==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${p.filter(_=>_.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${n==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${p.filter(_=>_.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${n==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${p.filter(_=>_.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${p.length} Instruksi Kerja
            </div>
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let g=document.createElement("style");g.innerHTML=`
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
    `,s.appendChild(g);let E=s.querySelector("#sec-ik-cards-grid-container"),C=s.querySelector("#sec-ik-count-label");function B(){let _=p.filter(L=>{let M=n==="all"||L.category===n,z=i.toLowerCase().trim(),x=!z||L.title.toLowerCase().includes(z)||L.target.toLowerCase().includes(z)||L.tools.some(T=>T.toLowerCase().includes(z))||L.steps.some(T=>T.toLowerCase().includes(z));return M&&x});if(C&&(C.textContent=`Menampilkan ${_.length} dari ${p.length} Instruksi Kerja`),_.length===0){E.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}E.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${_.map(L=>`
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${L.bg||"#EFF6FF"}; border-radius:10px;">${L.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(L.id).padStart(2,"0")}. ${L.title}
                      </h3>
                      <span class="badge" style="background:${L.bg||"#EFF6FF"}; color:${L.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${L.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${L.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${L.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${L.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${L.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${L.tools.slice(0,3).map(M=>`<li style="margin-bottom:2px;">${y(M)}</li>`).join("")}
                    ${L.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${L.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${L.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${L.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,E.querySelectorAll(".btn-view-sec-ik").forEach(L=>{L.addEventListener("click",()=>{let M=parseInt(L.dataset.id,10),z=p.find(x=>x.id===M);z&&k(z)})}),E.querySelectorAll(".btn-edit-sec-ik").forEach(L=>{L.addEventListener("click",()=>{let M=parseInt(L.dataset.id,10),z=p.find(x=>x.id===M);z&&w(z)})}),E.querySelectorAll(".btn-delete-sec-ik").forEach(L=>{L.addEventListener("click",()=>{let M=parseInt(L.dataset.id,10),z=p.find(x=>x.id===M);z&&ee(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${z.title}"?`,()=>{p=p.filter(x=>x.id!==M),Lt(p),K("Instruksi Kerja Security berhasil dihapus"),a()},"Hapus Instruksi Kerja")})})}B();let P=s.querySelector("#sec-ik-search");P&&P.addEventListener("input",_=>{i=_.target.value,B()}),s.querySelectorAll(".sec-cat-pill").forEach(_=>{_.addEventListener("click",()=>{s.querySelectorAll(".sec-cat-pill").forEach(L=>L.classList.remove("active")),_.classList.add("active"),n=_.dataset.cat,B()})});let D=s.querySelector("#btn-add-sec-ik");D&&D.addEventListener("click",()=>{w()})}function k(s){Y({title:`${s.icon||"\u{1F6E1}\uFE0F"} IK-${String(s.id).padStart(2,"0")}: ${s.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${s.bg||"#EFF6FF"}; color:${s.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${s.category}</span>
            <span>Kode: IK-SEC-${String(s.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${s.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${s.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${s.tools.map(g=>{if(g.includes("http")){let E=g.split(": ");return`<li>${E[0]}: <a href="${E.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${y(g)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${s.steps.map(g=>`<li style="margin-bottom:6px;">${y(g)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${s.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${y(s.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(g,E)=>{E()}})}function w(s=null){let g=!!s,E=s?s.tools.join(`
`):"",C=s?s.steps.join(`
`):"",B=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${s?y(s.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${s?s.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${s?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${s?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${s?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${s?y(s.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${y(E)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${y(C)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${s?y(s.notes||""):""}</textarea>
        </div>
      </form>
    `;Y({title:g?`\u270F\uFE0F Edit IK-${s.id}: ${s.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:B,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(P,D)=>{let _=P.querySelector("#form-sec-ik-title").value.trim(),L=P.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",M=P.querySelector("#form-sec-ik-cat").value,z=P.querySelector("#form-sec-ik-target").value.trim(),x=P.querySelector("#form-sec-ik-tools").value.trim(),T=P.querySelector("#form-sec-ik-steps").value.trim(),A=P.querySelector("#form-sec-ik-notes").value.trim();if(!_||!z||!x||!T){N("Harap lengkapi semua kolom bertanda bintang (*)");return}let q=x.split(`
`).map(R=>R.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),O=T.split(`
`).map(R=>R.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),j={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[M]||{color:"#1E40AF",bg:"#EFF6FF"};if(g)s.title=_,s.icon=L,s.category=M,s.target=z,s.tools=q,s.steps=O,s.notes=A,s.color=j.color,s.bg=j.bg,Lt(p),K(`Instruksi Kerja "${_}" berhasil diperbarui`);else{let R=p.length>0?Math.max(...p.map(de=>de.id))+1:1,ie={id:R,title:_,category:M,icon:L,color:j.color,bg:j.bg,target:z,tools:q,steps:O,notes:A};p.push(ie),Lt(p),K(`Instruksi Kerja baru IK-${R} berhasil ditambahkan`)}D(),a()}})}function f(s){s.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${b==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${b==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let g=s.querySelector("#zonasi-content-container");b==="cabang"?g.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${ha.cabang.map(E=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${E.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${E.color}; display:inline-block;"></span>
                  ${E.zone}
                </h3>
                <span class="badge" style="background:${E.bg}; color:${E.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${E.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${E.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${E.rooms.map(C=>`<li>${C}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:g.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${ha.hq.map(E=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${E.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${E.hijau.map(C=>`<li>${C}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${E.kuning.map(C=>`<li>${C}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${E.merah.map(C=>`<li>${C}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,s.querySelector("#view-cabang-btn").addEventListener("click",()=>{b="cabang",f(s)}),s.querySelector("#view-hq-btn").addEventListener("click",()=>{b="hq",f(s)})}function m(s){s.innerHTML=`
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
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${r}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let g=s.querySelector("#sec-asset-table-container"),E=s.querySelector("#sec-asset-search");function C(){let P=r.toLowerCase().trim(),D=l.filter(_=>!P||_.item.toLowerCase().includes(P)||_.type.toLowerCase().includes(P)||_.description.toLowerCase().includes(P)||_.pic&&_.pic.toLowerCase().includes(P));g.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${D.length} Aset</span>
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
                ${D.map((_,L)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${L%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||L+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${B(_.type)}">${y(_.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${y(_.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${y(_.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${y(_.pic||"-")}</td>
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
      `,g.querySelectorAll(".btn-edit-asset").forEach(_=>{_.addEventListener("click",()=>{let L=parseInt(_.dataset.no,10),M=l.find(z=>z.no===L);M&&o(M)})}),g.querySelectorAll(".btn-delete-asset").forEach(_=>{_.addEventListener("click",()=>{let L=parseInt(_.dataset.no,10),M=l.find(z=>z.no===L);M&&ee(`Apakah Anda yakin ingin menghapus aset: "${M.item}"?`,()=>{l=l.filter(z=>z.no!==L),l.forEach((z,x)=>{z.no=x+1}),It(l),K("Aset berhasil dihapus"),a()},"Hapus Aset")})})}function B(P){return P==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":P==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}C(),E&&E.addEventListener("input",P=>{r=P.target.value,C()}),s.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{o()})}function o(s=null){let g=!!s,E=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${s?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${s?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${s?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${s?y(s.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${s?y(s.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${s?y(s.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;Y({title:g?`\u270F\uFE0F Edit Aset: ${s.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:E,confirmText:"Simpan",onConfirm:(C,B)=>{let P=C.querySelector("#form-asset-type").value,D=C.querySelector("#form-asset-item").value.trim(),_=C.querySelector("#form-asset-desc").value.trim(),L=C.querySelector("#form-asset-pic").value.trim();if(!D||!_||!L){N("Harap lengkapi semua kolom");return}if(g)s.type=P,s.item=D,s.description=_,s.pic=L,It(l),K(`Aset "${D}" berhasil diperbarui`);else{let M=l.length>0?Math.max(...l.map(z=>z.no))+1:1;l.push({no:M,type:P,item:D,description:_,pic:L}),It(l),K(`Aset baru "${D}" berhasil ditambahkan`)}B(),a()}})}function d(s){s.innerHTML=`
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
    `}function h(s){s.innerHTML=`
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
    `}async function c(){try{let{downloadExcel:s}=await Promise.resolve().then(()=>(X(),ye)),g=p.map(C=>({"Kode IK":`IK-SEC-${String(C.id).padStart(2,"0")}`,"Judul Prosedur":C.title,Kategori:C.category,"Sasaran Target":C.target,"Alat & Perlengkapan":C.tools.join(", "),"Langkah Pelaksanaan":C.steps.map((B,P)=>`${P+1}. ${B}`).join(" | "),"Catatan Khusus":C.notes||"-"})),E=new Date().toISOString().slice(0,10);s(g,`SOP_Security_Instruksi_Kerja_${E}`),K("Data Excel SOP Security berhasil diunduh")}catch(s){console.error(s),N("Gagal mendownload data Excel: "+s.message)}}function y(s){return s?String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}a()}V();me();var va=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function ka(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function wa(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function Sa(e){let t=[],i="all",n="",r="fcms-master-doc-styles";if(!document.getElementById(r)){let k=document.createElement("style");k.id=r,k.textContent=`
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
    `,document.head.appendChild(k)}async function b(){try{let k=await I("/api/checklist?limit=1000");if(k.ok&&k.data?.data&&Array.isArray(k.data.data)&&k.data.data.length>0)t=k.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t));else{let w=localStorage.getItem("fcms_master_doc_cache");w?t=JSON.parse(w):(t=[...va],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)))}}catch(k){console.warn("Fallback loading documents from cache/seed:",k);let w=localStorage.getItem("fcms_master_doc_cache");t=w?JSON.parse(w):[...va]}S()}function p(){return t.filter(k=>{let w=i==="all"||k.category===i,f=n.toLowerCase().trim(),m=!f||(k.name||"").toLowerCase().includes(f)||(k.category||"").toLowerCase().includes(f)||(k.description||"").toLowerCase().includes(f);return w&&m})}function l(){let k={all:t.length};return t.forEach(w=>{let f=w.category||"Lainnya";k[f]=(k[f]||0)+1}),k}function a(k=null){let w=!!k,f=Array.from(new Set(t.map(o=>o.category).filter(Boolean))),m=document.createElement("div");m.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${k?.name?String(k.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${k?.category?String(k.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${f.map(o=>`<option value="${o}">${o}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${k?.document_link?String(k.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${k?.description||""}</textarea>
        </div>
      </div>
    `,Y({title:w?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:m,size:"md",confirmText:w?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(o,d)=>{let h=m.querySelector("#doc-modal-name"),c=m.querySelector("#doc-modal-category"),y=m.querySelector("#doc-modal-link"),s=m.querySelector("#doc-modal-desc"),g=(h.value||"").trim(),E=(c.value||"").trim(),C=(y.value||"").trim(),B=(s.value||"").trim();if(!g){N("Nama Dokumen wajib diisi"),h.focus();return}if(!E){N("Kategori Dokumen wajib diisi"),c.focus();return}let P={name:g,category:E,document_link:C,description:B};try{if(w)if((await I(`/api/checklist/${k.id}`,{method:"PUT",body:JSON.stringify(P)})).ok){let _=t.findIndex(L=>L.id===k.id);_!==-1&&(t[_]={...t[_],...P}),K("Dokumen berhasil diperbarui")}else{let _=t.findIndex(L=>L.id===k.id);_!==-1&&(t[_]={...t[_],...P}),K("Dokumen disimpan secara lokal")}else{let D=await I("/api/checklist",{method:"POST",body:JSON.stringify(P)});if(D.ok&&D.data?.id)t.unshift({id:D.data.id,...P}),K("Dokumen baru berhasil ditambahkan");else{let _=Date.now();t.unshift({id:_,...P}),K("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),d(),S()}catch(D){console.error("Error saving document:",D),N("Gagal menyimpan dokumen: "+D.message)}}})}function u(k){ee(`Apakah Anda yakin ingin menghapus dokumen <strong>"${k.name}"</strong>?`,async()=>{try{let w=await I(`/api/checklist/${k.id}`,{method:"DELETE"});w.ok||console.warn("API delete warning:",w),t=t.filter(f=>f.id!==k.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),K(`Dokumen "${k.name}" berhasil dihapus`),S()}catch{t=t.filter(f=>f.id!==k.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),K(`Dokumen "${k.name}" berhasil dihapus`),S()}},"Hapus Dokumen")}function F(k,w){if(!k){xt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(k).then(()=>{K(`Tautan "${w}" disalin ke clipboard!`)}).catch(()=>{xt(`Tautan: ${k}`)})}async function $(){try{let{downloadExcel:k}=await Promise.resolve().then(()=>(X(),ye)),w=t.map((f,m)=>({No:m+1,"Nama Dokumen":f.name||"",Kategori:f.category||"","Platform / Format":wa(f.document_link).type,"Link Dokumen":f.document_link||"","Deskripsi / Keterangan":f.description||""}));k(w,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),K("Berhasil mengunduh data dokumen ke Excel.")}catch(k){N("Gagal mengekspor data: "+k.message)}}function S(){let k=p(),w=l(),f=Object.keys(w).filter(o=>o!=="all"),m=t.filter(o=>!!o.document_link).length;e.innerHTML=`
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
                <div class="doc-stat-value">${f.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${m} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${t.length}</span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${n?n.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${i==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${w.all||0}</span>
            </button>
            ${f.map(o=>{let d=ka(o);return`
                <button class="doc-pill ${i===o?"active":""}" data-cat="${o}">
                  <span>${d.icon}</span>
                  <span>${o}</span>
                  <span class="doc-pill-count">${w[o]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${k.length===0?`
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
              ${k.map(o=>{let d=ka(o.category),h=wa(o.document_link);return`
                  <div class="doc-card" data-id="${o.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border};">
                          <span>${d.icon}</span>
                          <span>${d.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${h.icon}</span>
                          <span>${h.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${o.name}</h3>
                      <p class="doc-card-desc">${o.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${o.document_link?`
                        <a href="${o.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${o.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${o.document_link}" data-title="${o.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${o.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${o.id}">
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
    `,v()}function v(){let k=e.querySelector("#doc-btn-add");k&&(k.onclick=()=>a());let w=e.querySelector("#doc-btn-excel");w&&(w.onclick=()=>$());let f=e.querySelector("#doc-btn-refresh");f&&(f.onclick=()=>b());let m=e.querySelector("#doc-btn-reset-filter");m&&(m.onclick=()=>{n="",i="all",S()});let o=e.querySelector("#doc-search-input");o&&(o.oninput=d=>{n=d.target.value;let h=p();S();let c=e.querySelector("#doc-search-input");c&&(c.focus(),c.setSelectionRange(c.value.length,c.value.length))}),e.querySelectorAll(".doc-pill").forEach(d=>{d.onclick=()=>{i=d.getAttribute("data-cat")||"all",S()}}),e.querySelectorAll(".copy-link-btn").forEach(d=>{d.onclick=h=>{h.stopPropagation();let c=d.getAttribute("data-link"),y=d.getAttribute("data-title");F(c,y)}}),e.querySelectorAll(".edit-doc-btn").forEach(d=>{d.onclick=h=>{h.stopPropagation();let c=Number(d.getAttribute("data-id")),y=t.find(s=>s.id===c);y&&a(y)}}),e.querySelectorAll(".delete-doc-btn").forEach(d=>{d.onclick=h=>{h.stopPropagation();let c=Number(d.getAttribute("data-id")),y=t.find(s=>s.id===c);y&&u(y)}})}await b()}V();me();X();var Me={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Nt(e){for(let[t,i]of Object.entries(Me))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function Ea(e){e.innerHTML=`
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
        ${Object.entries(Me).map(([d,h])=>`
          <button class="zone-tab-btn" data-zone="${d}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${h.icon}</span> ${d}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${d.replace(/[\s&]+/g,"-")}">-</span>
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
  `;let t="all",i="table",n=[],r=ne(),b=r&&["superadmin","admin","manager","spv"].includes(r.role);async function p(){try{let d=await I("/api/hygiene-standards?all=1");if(!d.ok)throw new Error("Gagal memuat standar hygiene");n=Array.isArray(d.data)?d.data:d.data?.data||[],a(),l(),F()}catch(d){console.error(d),N(d.message||"Terjadi kesalahan memuat data");let h=e.querySelector("#hygiene-view-container");h&&(h.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${d.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,h.querySelector("#btn-retry-hygiene")?.addEventListener("click",p))}}function l(){let d=e.querySelector("#badge-count-all");d&&(d.textContent=n.length),Object.entries(Me).forEach(([y,s])=>{let g=y.replace(/[\s&]+/g,"-"),E=e.querySelector(`#badge-count-${g}`);if(E){let C=n.filter(B=>s.rooms.includes(B.room_name)).length;E.textContent=C}});let h=new Set(n.map(y=>y.room_name)),c=e.querySelector("#card-total-rooms");c&&(c.textContent=h.size)}function a(){let d=e.querySelector("#hygiene-room-filter");if(!d)return;let h=d.value,c=[];if(t==="all")c=Array.from(new Set(n.map(y=>y.room_name))).filter(Boolean).sort();else{let y=Me[t];y&&(c=y.rooms.filter(s=>n.some(g=>g.room_name===s)))}d.innerHTML=`<option value="">Semua Ruangan (${c.length})</option>`+c.map(y=>`<option value="${y}" ${y===h?"selected":""}>${y}</option>`).join("")}function u(){let d=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),h=e.querySelector("#hygiene-room-filter")?.value||"";return n.filter(c=>{if(t!=="all"){let y=Me[t];if(!y||!y.rooms.includes(c.room_name))return!1}if(h&&c.room_name!==h)return!1;if(d){let y=(c.room_name||"").toLowerCase(),s=(c.item_name||"").toLowerCase(),g=(c.cleanliness_standard||"").toLowerCase();if(!y.includes(d)&&!s.includes(d)&&!g.includes(d))return!1}return!0})}function F(){let d=u(),h=e.querySelector("#hygiene-stats-text");h&&(h.innerHTML=`Menampilkan <strong>${d.length}</strong> dari <strong>${n.length}</strong> parameter`);let c=e.querySelector("#hygiene-view-container");if(c){if(d.length===0){c.innerHTML=`
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
      `,c.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let y=e.querySelector("#hygiene-search"),s=e.querySelector("#hygiene-room-filter");y&&(y.value=""),s&&(s.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(g=>g.classList.toggle("active",g.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(g=>{g.dataset.zone==="all"?(g.style.background="#2563EB",g.style.color="#fff"):(g.style.background="#F1F5F9",g.style.color="var(--text-2)")}),a(),F()});return}i==="table"?$(c,d):S(c,d)}}function $(d,h){d.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${b?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${h.map((c,y)=>{let s=Nt(c.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${y+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${s.bg}; color:${s.color}; border:1px solid ${s.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${s.icon}</span> ${c.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${s.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${c.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${c.cleanliness_standard}
                      </div>
                    </td>
                    ${b?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${c.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${c.id}" data-name="${c.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,v(d)}function S(d,h){let c={};h.forEach(y=>{let s=Nt(y.room_name);c[s.zoneName]||(c[s.zoneName]={info:s,rooms:{}}),c[s.zoneName].rooms[y.room_name]||(c[s.zoneName].rooms[y.room_name]=[]),c[s.zoneName].rooms[y.room_name].push(y)}),d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(c).map(([y,s])=>{let g=Object.values(s.rooms).reduce((E,C)=>E+C.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${s.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${s.info.bg}; border-bottom:1px solid ${s.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${s.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${s.info.color};">
                      ${y}
                    </h3>
                    <div style="font-size:0.78rem; color:${s.info.color}; opacity:0.85;">
                      ${Object.keys(s.rooms).length} Ruangan \xB7 ${g} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${s.info.color}; border:1px solid ${s.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${g} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(s.rooms).map(([E,C])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${E}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${C.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${b?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${C.map((B,P)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${P+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${B.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${B.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${b?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${B.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${B.id}" data-name="${B.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,v(d)}function v(d){d.querySelectorAll(".btn-edit-item").forEach(h=>{h.addEventListener("click",()=>{let c=h.dataset.id,y=n.find(s=>String(s.id)===String(c));y&&k(y)})}),d.querySelectorAll(".btn-delete-item").forEach(h=>{h.addEventListener("click",()=>{let c=h.dataset.id,y=h.dataset.name;ee(`Yakin ingin menghapus parameter standar <strong>"${y}"</strong>?`,async()=>{try{let s=await I(`/api/hygiene-standards/${c}`,{method:"DELETE"});if(!s.ok)throw new Error(s.data?.error||"Gagal menghapus parameter");K("Parameter kebersihan berhasil dihapus"),p()}catch(s){N(s.message)}},"Hapus Parameter Standar")})})}function k(d=null){let h=!!d,c=Array.from(new Set(n.map(s=>s.room_name))).filter(Boolean).sort(),y=document.createElement("div");y.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${d?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${c.map(s=>`<option value="${s}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${d?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${d?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,Y({title:h?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:y,confirmText:h?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(s,g)=>{let E=y.querySelector("#modal-room-name").value.trim(),C=y.querySelector("#modal-item-name").value.trim(),B=y.querySelector("#modal-standard").value.trim();if(!E||!C||!B){N("Semua kolom wajib diisi");return}try{let P={room_name:E,item_name:C,cleanliness_standard:B},D=h?`/api/hygiene-standards/${d.id}`:"/api/hygiene-standards",L=await I(D,{method:h?"PUT":"POST",body:JSON.stringify(P)});if(!L.ok)throw new Error(L.data?.error||"Gagal menyimpan parameter");K(h?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),g(),p()}catch(P){N(P.message)}}})}function w(){let d=window.open("","_blank");if(!d){N("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let c=`
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
            <div>Total: <strong>${n.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(Me).map(([y,s])=>{let g=n.filter(E=>s.rooms.includes(E.room_name));return g.length===0?"":`
            <div class="zone-title">${s.icon} ${y.toUpperCase()} (${g.length} PARAMETER)</div>
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
                ${g.map((E,C)=>`
                  <tr>
                    <td class="td-no">${C+1}</td>
                    <td class="td-room">${E.room_name}</td>
                    <td class="td-obj">${E.item_name}</td>
                    <td>${E.cleanliness_standard}</td>
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
    `;d.document.open(),d.document.write(c),d.document.close()}function f(){let d=n.map((c,y)=>{let s=Nt(c.room_name);return{No:y+1,"Kategori Zona":s.zoneName,"Ruangan / Lokasi":c.room_name,"Item / Objek Inspeksi":c.item_name,"Standar Kebersihan Fisik":c.cleanliness_standard}}),h=new Date().toISOString().slice(0,10);J(d,`Master_Standar_Kebersihan_FCMS_${h}`),K("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>F()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>F()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let d=e.querySelector("#hygiene-search"),h=e.querySelector("#hygiene-room-filter");d&&(d.value=""),h&&(h.value=""),F()}),e.querySelectorAll(".zone-tab-btn").forEach(d=>{d.addEventListener("click",()=>{t=d.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(h=>{let c=h.dataset.zone===t;if(h.classList.toggle("active",c),c){h.style.background="#2563EB",h.style.color="#fff";let y=h.querySelector(".badge");y&&(y.style.background="rgba(255,255,255,0.25)",y.style.color="#fff")}else{h.style.background="#F1F5F9",h.style.color="var(--text-2)";let y=h.querySelector(".badge");y&&(y.style.background="#CBD5E1",y.style.color="#334155")}}),a(),F()})});let m=e.querySelector("#view-table-btn"),o=e.querySelector("#view-zone-btn");m?.addEventListener("click",()=>{i="table",m.style.background="#fff",m.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",o.style.background="transparent",o.style.boxShadow="none",F()}),o?.addEventListener("click",()=>{i="zone",o.style.background="#fff",o.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",m.style.background="transparent",m.style.boxShadow="none",F()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",w),e.querySelector("#btn-export-hygiene")?.addEventListener("click",f),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>k(null)),p()}V();me();X();async function Ca(e,t="supply"){return Cn(e)}async function Cn(e){let i=((await I("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));W({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>ae(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let r=n?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let b=n?.chemical_items;try{b=Array.isArray(JSON.parse(b))?JSON.parse(b).join(", "):b}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!i.find(p=>p.value==n.branch_id)?[...i,{value:n.branch_id,label:n.branch_name||n.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:b},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),b=await I(`/api/reports/supply?limit=10000&${r}`);if(b.ok){let p=b.data.data.map(l=>{let a=l.tools_items;try{a=Array.isArray(JSON.parse(a))?JSON.parse(a).join(", "):a}catch{}let u=l.chemical_items;try{u=Array.isArray(JSON.parse(u))?JSON.parse(u).join(", "):u}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":a||"",Chemical:u||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});J(p,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let b=(await I("/api/branches?all=1")).data?.data||[],p=F=>{if(!F)return null;let $=String(F||"").toLowerCase(),S=b.find(v=>String(v.full_name||"").toLowerCase()===$||String(v.code||"").toLowerCase()===$||String(v.name||"").toLowerCase()===$);return S?S.id:null},l=F=>{if(F==null||F==="")return"";if(F instanceof Date&&!isNaN(F.getTime()))return F.toISOString().slice(0,10);let $=String(F).trim();if($===""||$==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test($))return $.slice(0,10);if(/^\d{4,5}$/.test($)){let v=Number($);if(v>2e4&&v<99999){let k=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let S=$.split(/[\/\-\.]/);if(S.length===3){let[v,k,w]=S.map(f=>f.trim());if(v.length===4&&k.length<=2&&w.length<=2)return`${v}-${k.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&k.length<=2&&v.length<=2)return`${w}-${k.padStart(2,"0")}-${v.padStart(2,"0")}`}return $},a=n.map(F=>({submitted_at:l(F.Waktu||F.Tanggal),submitter_name:String(F.Pengirim||"").trim(),branch_id:p(String(F.Cabang||"").trim()),tools_items:String(F["Alat/Barang"]||F.Alat||"").trim(),chemical_items:String(F.Chemical||"").trim(),additional_notes:String(F.Catatan||F.Keterangan||"").trim(),status:String(F.Status||"").trim(),processed_by:String(F["Diproses Oleh"]||F.PIC||"").trim()})).filter(F=>F.submitted_at&&F.submitter_name&&F.branch_id),u=await I("/api/reports/supply/import",{method:"POST",body:JSON.stringify(a)});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,r)=>{let b=Y({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${n.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${n.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${n.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${n.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(p,l)=>{let a=p.querySelector("#supply-status").value,u=p.querySelector("#supply-processed-by").value;(await I(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:a,processed_by:u})})).ok?(K("Status diperbarui."),l(),r()):N("Gagal update status.")}})}}]})}V();X();async function Fa(e){let t=ne();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}W({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let n=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await I(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let n=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));J(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let n=i.map(b=>({full_name:String(b["Nama Lengkap"]||"").trim(),username:String(b.Username||"").trim(),email:String(b.Email||"").trim(),role:String(b.Role||"").trim()||"viewer",password:String(b.Password||"").trim()})).filter(b=>b.username&&b.password&&b.email&&b.full_name),r=await I("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}V();X();async function $a(e){W({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await I(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)J(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{J([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),n=await I("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}V();async function _a(e){let t=new Date,i=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(p=>p.addEventListener("change",r));async function n(){try{let p=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await I(`/api/dashboard/calendar?month=${p}`)).data?.data||[]}catch(p){console.warn("[Calendar] Failed to load events, rendering empty grid:",p),i=[]}}async function r(){let p=document.getElementById("calendar-grid");if(p){p.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let l=t.getFullYear(),a=t.getMonth(),u=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),F=document.getElementById("cal-month-label");F&&(F.textContent=u);let $=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(c=>c.value)),S=i.filter(c=>$.has(c.type)),v={};S.forEach(c=>{let y=Fn(c.event_date);y&&(v[y]||(v[y]=[]),v[y].push(c))});let k=new Date(l,a,1).getDay(),w=new Date(l,a+1,0).getDate(),f=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],m=new Date().toISOString().slice(0,10),o='<div class="calendar-grid">';f.forEach(c=>{o+=`<div class="cal-day-header">${c}</div>`});for(let c=0;c<k;c++)o+='<div class="cal-cell cal-cell-empty"></div>';for(let c=1;c<=w;c++){let y=`${l}-${String(a+1).padStart(2,"0")}-${String(c).padStart(2,"0")}`,s=v[y]||[],g=y===m;o+=`
          <div class="cal-cell ${g?"cal-today":""} ${s.length?"cal-has-events":""}"
               data-date="${y}" tabindex="0" role="button" aria-label="${y}">
            <div class="cal-day-num ${g?"today-num":""}">${c}</div>
            <div class="cal-events-preview">
              ${s.slice(0,3).map(E=>`
                <div class="cal-event-dot cal-color-${E.color||"gray"}" title="${Se(E.title||E.type)}">
                  <span class="cal-event-dot-label">${$n(E.title||E.branch_name||E.type,18)}</span>
                </div>
              `).join("")}
              ${s.length>3?`<div class="cal-more">+${s.length-3} lagi</div>`:""}
            </div>
          </div>`}let h=(k+w)%7;if(h!==0)for(let c=0;c<7-h;c++)o+='<div class="cal-cell cal-cell-empty"></div>';o+="</div>",p.innerHTML=o,p.querySelectorAll(".cal-cell[data-date]").forEach(c=>{c.addEventListener("click",()=>{let y=c.dataset.date,s=v[y]||[];if(!s.length)return;let g=document.getElementById("cal-event-list"),E=new Date(y+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=E,document.getElementById("cal-event-items").innerHTML=s.map(C=>`
            <div class="cal-event-item cal-color-border-${C.color||"gray"}">
              <div class="cal-event-type">
                ${_n(C.type)}
                ${C.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${Se(C.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${Se(C.title||"-")}</div>
              ${C.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${Se(C.branch_name)}</div>`:""}
              ${C.pic||C.period||C.opening_date||C.target_date||C.completion_date||C.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${C.pic?`<div>\u{1F464} <b>PIC:</b> ${Se(C.pic)}</div>`:""}
                  ${C.period?`<div>\u{1F4C5} <b>Periode:</b> ${Se(C.period)}</div>`:""}
                  ${C.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${Se(Kt(C.opening_date))}</div>`:""}
                  ${C.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${Se(Kt(C.target_date))}</div>`:""}
                  ${C.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${Se(Kt(C.completion_date))}</div>`:""}
                  ${C.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${Se(C.notes)}</div>`:""}
                </div>
              `:""}
              ${C.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${C.status==="Done"?"badge-success":"badge-warning"}">${Se(C.status)}</span></div>`:""}
              ${C.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${C.days_remaining} hari</div>`:""}
            </div>
          `).join(""),g.style.display="block"})})}catch(l){console.error("[Calendar] Render error:",l),p&&(p.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let b=()=>{n().then(()=>r())};at.on("data:changed",b),r()}function Fn(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[n,r,b]=i.map(p=>p.trim());if(b.length===4)return`${b}-${r.padStart(2,"0")}-${n.padStart(2,"0")}`;if(n.length===4)return`${n}-${r.padStart(2,"0")}-${b.padStart(2,"0")}`}return t.slice(0,10)}function Kt(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[n,r,b]=i.map(p=>p.trim());if(n.length===4)return`${b.padStart(2,"0")}-${r.padStart(2,"0")}-${n}`;if(b.length===4)return`${n.padStart(2,"0")}-${r.padStart(2,"0")}-${b}`}return e}function $n(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function Se(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function _n(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}V();async function Ta(e){let t=ne(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${r},${r}99)">
            ${i}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${r}18;color:${r};margin-top:6px">
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
            <span class="info-value" style="color:${r};font-weight:700">${t?.role||"\u2014"}</span>
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
  `;let b=localStorage.getItem("fm_token"),p=document.getElementById("session-info");if(b&&p)try{let l=JSON.parse(atob(b.split(".")[1])),a=new Date(l.exp*1e3);p.textContent=`Berakhir: ${a.toLocaleString("id-ID")}`}catch{p.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let a=document.getElementById("pwd-error"),u=document.getElementById("pwd-success"),F=document.getElementById("btn-save-pwd");a.style.display="none",u.style.display="none";let $=l.target,S=$.current_password.value,v=$.new_password.value,k=$.confirm_password.value;if(v!==k){a.textContent="\u274C Konfirmasi password tidak cocok.",a.style.display="block";return}if(v.length<6){a.textContent="\u274C Password baru minimal 6 karakter.",a.style.display="block";return}F.disabled=!0,F.textContent="\u23F3 Menyimpan...";let w=await I("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:S,new_password:v})});F.disabled=!1,F.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',w.ok?(u.textContent="\u2705 Password berhasil diubah.",u.style.display="block",$.reset(),K("Password berhasil diubah.")):(a.textContent=w.data?.error||"Gagal mengubah password.",a.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}V();var ct={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function le(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let r=Number(t);if(r>2e4&&r<99999){let b=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(b.getTime())?null:b.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[r,b,p]=i.map(F=>F.trim()),l=Number(r),a=Number(b),u=Number(p);if(r.length===4&&l>1900)return`${r}-${b.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&u>1900)return l>12?`${p}-${b.padStart(2,"0")}-${r.padStart(2,"0")}`:a>12?`${p}-${r.padStart(2,"0")}-${b.padStart(2,"0")}`:`${p}-${b.padStart(2,"0")}-${r.padStart(2,"0")}`;if(p.length===2&&!isNaN(u)){let F=u>=50?`19${p}`:`20${p}`;return l>12?`${F}-${b.padStart(2,"0")}-${r.padStart(2,"0")}`:`${F}-${b.padStart(2,"0")}-${r.padStart(2,"0")}`}}let n=new Date(t);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Aa(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Tn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:le(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:le(e["Target Pindah OS"]),target_selesai_os:le(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:le(e["Tanggal Mulai"]),end_date:le(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:le(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:le(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:le(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:le(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:le(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:le(e["Tanggal Target"]||e["Tgl Target"]),completion_date:le(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:le(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:le(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:le(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:le(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:le(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:le(e["Tanggal Back Up"]),completion_date:le(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:le(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:le(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function An(e,t){let i=ct[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let n=Tn[i.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],b=[],p=[];return t.filter(a=>!Aa(a)).forEach((a,u)=>{let F=t.indexOf(a)+2,$=[];n.required.forEach(({key:v,label:k})=>{let w=a[v];if(w==null||String(w).trim()===""){let f=Object.keys(a).filter(m=>m.trim()).join(", ");$.push({column:k,originalValue:w||"",reason:`Kolom "${k}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${f.slice(0,120)}`})}});let S=n.map(a);$.length>0?b.push({row:F,data:S,raw:a,errors:$}):(r.push(a),p.push(S))}),{valid:r,errors:b,mapped:p}}function Da(e){let t=[];return e.SheetNames.forEach(i=>{let n=ct[i];if(!n)return;let r=e.Sheets[i],b=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),p=An(i,b),l=b.filter(a=>!Aa(a));t.push({sheetName:i,module:n.module,label:n.label,total:l.length,valid:p.mapped.length,errorCount:p.errors.length,errors:p.errors,mapped:p.mapped,skipped:!1})}),t}function Ba(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([n,r])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(r),n)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function Pa(e){let t=window.XLSX,i=t.utils.book_new(),n=!1;return e.forEach(r=>{if(!r.errors||r.errors.length===0)return;n=!0;let b=r.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(a=>a.column||a).join("; "),"Alasan Error":(l.errors||[]).map(a=>a.reason||a).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([a,u])=>[a,u??""]))})),p=t.utils.json_to_sheet(b);t.utils.book_append_sheet(i,p,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Dn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function La(e){e.innerHTML=`
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
              ${Object.entries(ct).map(([w,{label:f}])=>`<span class="import-sheet-tag">\u{1F4C4} ${w} \u2192 ${f}</span>`).join("")}
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
  `;let t=null,i=null,n=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function b(w){Object.entries(r).forEach(([f,m])=>{m.style.display=f===w?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let w=document.getElementById("btn-backup-db");w.disabled=!0,w.textContent="\u23F3 Memproses Backup...";try{let f=await I("/api/import/backup");if(f.ok){if(!window.XLSX){N("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let m=window.XLSX,o=m.utils.book_new();Object.entries(f.data.database).forEach(([d,h])=>{let c=h.length>0?h:[{}],y=m.utils.json_to_sheet(c);m.utils.book_append_sheet(o,y,d.substring(0,31))}),m.writeFile(o,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),K("Backup berhasil diunduh!")}else N("Gagal memproses backup: "+(f.data?.error||"Unknown error"))}catch(f){N("Gagal memproses backup: "+f.message)}finally{w.disabled=!1,w.textContent="\u{1F4E6} Backup Database"}});let p=document.getElementById("btn-sync-google");p&&p.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let w=p.innerHTML;p.innerHTML='<span class="spinner"></span> Menyinkronkan...',p.disabled=!0;try{let f=await I("/api/sync/google-sheets",{method:"POST"});f.ok?alert("Sinkronisasi Berhasil: "+(f.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(f.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{p.innerHTML=w,p.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Ba(),K("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),a=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",w=>{w.stopPropagation(),l.click()}),l.addEventListener("change",w=>{w.target.files[0]&&u(w.target.files[0])}),a.addEventListener("dragover",w=>{w.preventDefault(),a.classList.add("drag-over")}),a.addEventListener("dragleave",()=>a.classList.remove("drag-over")),a.addEventListener("drop",w=>{w.preventDefault(),a.classList.remove("drag-over");let f=w.dataTransfer.files[0];f&&f.name.match(/\.xlsx?$/i)?u(f):N("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,l.value="",document.getElementById("file-info").style.display="none",a.style.display="",b("upload")});async function u(w){t=w,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${w.name} (${(w.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",a.style.display="none",await F(w)}async function F(w){b("validating");let f=document.getElementById("validation-status"),m=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");f.textContent="Membaca file Excel...",m.style.width="20%",await Ve(200);let o=await w.arrayBuffer(),d=window.XLSX.read(o,{type:"array",cellDates:!0});f.textContent=`Memvalidasi ${d.SheetNames.length} sheet...`,m.style.width="50%",await Ve(100),i=Da(d),m.style.width="100%",f.textContent="Validasi selesai!",await Ve(300),$()}catch(o){b("upload"),N("Gagal memproses file: "+o.message),document.getElementById("file-info").style.display="flex",a.style.display="none"}}function $(){b("preview");let w=i.filter(g=>!g.skipped).length,f=i.reduce((g,E)=>g+E.total,0),m=i.reduce((g,E)=>g+E.valid,0),o=i.reduce((g,E)=>g+E.errorCount,0),d=f>0?Math.round(m/f*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${w} sheet</span>
      <span class="badge badge-secondary">${f} baris</span>
      <span class="badge badge-success">${m} valid (${d}%)</span>
      ${o>0?`<span class="badge badge-danger">${o} error</span>`:""}
    `;let h=document.getElementById("preview-table-container");h.innerHTML=`
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
          ${i.map((g,E)=>`
            <tr class="${g.errorCount>0?"row-error":g.skipped?"row-skipped":"row-ok"}">
              <td><strong>${g.sheetName}</strong></td>
              <td>${g.label}</td>
              <td style="text-align:center">${g.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${g.valid}</span></td>
              <td style="text-align:center">${g.errorCount>0?`<span class="badge badge-danger">${g.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${g.skipped?'<span class="badge badge-neutral">Dilewati</span>':g.errorCount>0&&g.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':g.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':g.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${g.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${E}">\u{1F50D} ${g.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,h.querySelectorAll(".btn-detail-error").forEach(g=>{g.addEventListener("click",()=>{let E=i[Number(g.dataset.idx)];S(E)})});let c=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");y.innerHTML="",c.style.display="none";let s=document.getElementById("btn-start-import");m===0?(s.disabled=!0,s.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(s.disabled=!1,o>0?(s.innerHTML=`\u{1F680} Import ${m} Data Valid (${o} dilewati)`,s.title="Baris error akan dilewati, baris valid tetap diimport"):s.innerHTML=`\u{1F680} Mulai Import ${m} Data`)}function S(w){let f=document.getElementById("error-detail-section"),m=document.getElementById("error-detail-container");f.style.display="";let o=w.errors.slice(0,100).map(d=>(Array.isArray(d.errors)?d.errors:[]).map(c=>{let y=typeof c=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${d.row}</span></td>
            <td><strong>${y?c.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${y&&c.originalValue!==void 0?c.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${y?c.reason:c}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${y&&c.aliases?`Gunakan salah satu nama kolom:<br><em>${c.aliases}</em>`:y&&c.hint?c.hint:""}
            </td>
          </tr>
        `}).join("")).join("");m.innerHTML=`
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
            <tbody>${o||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${w.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,f.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{b("upload"),document.getElementById("file-info").style.display="none",a.style.display="",t=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;Pa(i)?K("Log error berhasil didownload."):K("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let w=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";v(w)});async function v(w){b("importing"),n=Date.now();let f=[];Dn.forEach(c=>{let y=i?.find(s=>s.module===c&&s.mapped?.length>0);y&&f.push(y)});let m=document.getElementById("import-steps-list");m.innerHTML=f.map(c=>`
      <div class="import-step-item" id="step-item-${c.module}">
        <span class="step-item-icon" id="step-icon-${c.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${c.label} <span class="step-item-count">(${c.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${c.module}"></span>
      </div>
    `).join("");let o=document.getElementById("import-bar"),d=document.getElementById("import-current-status"),h={totalSheets:f.length,totalRows:f.reduce((c,y)=>c+y.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let c=0;c<f.length;c++){let y=f[c],s=document.getElementById(`step-icon-${y.module}`),g=document.getElementById(`step-status-${y.module}`);s.textContent="\u{1F504}",g.textContent="Mengimport...",d.textContent=`Mengimport ${y.label}...`,o.style.width=`${Math.round(c/f.length*100)}%`;try{let E=await I(`/api/import/${y.module}`,{method:"POST",body:JSON.stringify({rows:y.mapped,onDuplicate:w})});if(E.ok){let C=E.data;h.inserted+=C.inserted||0,h.skipped+=C.skipped||0,h.moduleResults.push({label:y.label,inserted:C.inserted||0,skipped:C.skipped||0,status:"ok"}),s.textContent="\u2705",g.innerHTML=`<span class="badge badge-success">${C.inserted||0} berhasil</span>${C.skipped>0?` <span class="badge badge-neutral">${C.skipped} skip</span>`:""}`}else h.failed++,h.moduleResults.push({label:y.label,inserted:0,skipped:0,status:"error",error:E.data?.error}),s.textContent="\u274C",g.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(E){h.failed++,h.moduleResults.push({label:y.label,inserted:0,skipped:0,status:"error",error:E.message}),s.textContent="\u274C",g.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ve(150)}o.style.width="100%",d.textContent="Selesai!",await Ve(400),k(h)}function k(w){b("summary");let f=((Date.now()-n)/1e3).toFixed(1),m=w.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${m?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${m?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${f}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${w.moduleResults.map(o=>`
            <tr>
              <td>${o.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${o.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${o.skipped}</span></td>
              <td style="text-align:center">
                ${o.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${o.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,l.value="",document.getElementById("file-info").style.display="none",a.style.display="",b("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ve(e){return new Promise(t=>setTimeout(t,e))}V();me();var Ia=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function ze(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let i=new Date(e);i.setHours(0,0,0,0);let n=i.getTime()-t.getTime(),r=Math.ceil(n/(1e3*60*60*24));return r<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(r)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:r,isExpired:!0,isExpiringSoon:!1}:r<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${r} hari)`,shortLabel:`Sisa ${r} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:r,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${r} hari)`,shortLabel:`Aktif (${r} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:r,isExpired:!1,isExpiringSoon:!1}}function Ma(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function ut(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function za(e){let t=[],i=[],n=[],r="all",b="all",p="all",l="",a="fcms-sp-page-styles";if(!document.getElementById(a)){let o=document.createElement("style");o.id=a,o.textContent=`
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
        .sp-hero-header > div:last-child {
          display: flex !important;
          width: 100% !important;
          gap: 8px !important;
          flex-wrap: nowrap !important;
        }
        #sp-btn-add {
          flex: 2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 12px !important;
        }
        #sp-btn-excel {
          flex: 1.2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 10px !important;
        }
        #sp-btn-refresh {
          flex: 0 0 38px !important;
          justify-content: center !important;
          padding: 9px 0 !important;
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
    `,document.head.appendChild(o)}async function u(){try{let[o,d]=await Promise.all([Q(),ce()]);i=o||[],n=d||[];let h=await I("/api/sp?limit=1000");if(h.ok&&h.data?.data&&Array.isArray(h.data.data)&&h.data.data.length>0)t=h.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t));else{let c=localStorage.getItem("fcms_sp_data_cache");c?t=JSON.parse(c):(t=[...Ia],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)))}}catch(o){console.warn("Fallback loading SP data from cache/seed:",o);let d=localStorage.getItem("fcms_sp_data_cache");t=d?JSON.parse(d):[...Ia]}f()}function F(){return t.filter(o=>{let d=l.toLowerCase().trim(),h=!d||(o.employee_name||"").toLowerCase().includes(d)||(o.branch_name||"").toLowerCase().includes(d)||(o.division||"").toLowerCase().includes(d)||(o.sp_type||"").toLowerCase().includes(d),c=b==="all"||String(o.branch_id)===String(b),y=p==="all"||o.division===p,s=ze(o.akhir_sp),g=!0;return r==="active"?g=!s.isExpired:r==="expiring"?g=s.isExpiringSoon:r==="expired"?g=s.isExpired:r==="sp1"?g=(o.sp_type||"").toUpperCase()==="SP 1":r==="sp2"?g=(o.sp_type||"").toUpperCase()==="SP 2":r==="sp3"?g=(o.sp_type||"").toUpperCase()==="SP 3":r==="teguran"&&(g=(o.sp_type||"").toUpperCase().includes("TEGURAN")),h&&c&&y&&g})}function $(){let o=0,d=0,h=0,c=0,y=0,s=0,g=0,E=new Set;return t.forEach(C=>{C.employee_name&&E.add(C.employee_name.toLowerCase().trim());let B=ze(C.akhir_sp);B.isExpired?h++:o++,B.isExpiringSoon&&d++;let P=(C.sp_type||"").toUpperCase();P==="SP 1"?c++:P==="SP 2"?y++:P==="SP 3"?s++:P.includes("TEGURAN")&&g++}),{total:t.length,active:o,expiring:d,expired:h,sp1:c,sp2:y,sp3:s,teguran:g,uniqueEmployees:E.size}}function S(o=null){let d=!!o,h=o?.tanggal?o.tanggal.slice(0,10):new Date().toISOString().slice(0,10),c=o?.akhir_sp?o.akhir_sp.slice(0,10):"";if(!c){let E=new Date(h);E.setMonth(E.getMonth()+6),c=E.toISOString().slice(0,10)}let y=document.createElement("div");y.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${o?.employee_name?String(o.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${n.map(E=>`<option value="${E.label}">${E.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${o?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${o?.division==="SECURITY"?"selected":""}>SECURITY</option>
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
              ${i.map(E=>`<option value="${E.value}" ${String(o?.branch_id)===String(E.value)?"selected":""}>${E.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${o?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${o?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${o?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${o?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${h}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${c}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${o?.document_link?String(o.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${o?.notes||o?.description||""}</textarea>
        </div>
      </div>
    `;let s=y.querySelector("#sp-modal-tanggal"),g=y.querySelector("#sp-modal-akhir");s.addEventListener("change",()=>{if(s.value){let E=new Date(s.value);E.setMonth(E.getMonth()+6),g.value=E.toISOString().slice(0,10)}}),Y({title:d?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:y,size:"md",confirmText:d?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(E,C)=>{let B=(y.querySelector("#sp-modal-emp").value||"").trim(),P=y.querySelector("#sp-modal-div").value,D=y.querySelector("#sp-modal-branch").value,_=y.querySelector("#sp-modal-type").value,L=y.querySelector("#sp-modal-tanggal").value,M=y.querySelector("#sp-modal-akhir").value,z=(y.querySelector("#sp-modal-link").value||"").trim(),x=(y.querySelector("#sp-modal-notes").value||"").trim();if(!B){N("Nama Karyawan wajib diisi");return}if(!D){N("Cabang penempatan wajib dipilih");return}if(!L||!M){N("Tanggal SP dan Akhir SP wajib diisi");return}let T=i.find(O=>String(O.value)===String(D)),A=T?T.label:"Cabang",q={employee_name:B,division:P,branch_id:Number(D),branch_name:A,sp_type:_,tanggal:L,akhir_sp:M,document_link:z,notes:x,status:ze(M).status};try{if(d){let O=await I(`/api/sp/${o.id}`,{method:"PUT",body:JSON.stringify(q)}),G=t.findIndex(j=>j.id===o.id);G!==-1&&(t[G]={...t[G],...q}),K("Data SP berhasil diperbarui")}else{let O=await I("/api/sp",{method:"POST",body:JSON.stringify(q)}),G=O.ok&&O.data?.id?O.data.id:Date.now();t.unshift({id:G,...q}),K("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),C(),f()}catch(O){console.error("Error saving SP:",O),N("Gagal menyimpan SP: "+O.message)}}})}function v(o){ee(`Hapus data Surat Peringatan untuk karyawan <strong>${o.employee_name}</strong> (${o.sp_type})?`,async()=>{try{await I(`/api/sp/${o.id}`,{method:"DELETE"})}catch(d){console.warn("API delete error fallback to local:",d)}t=t.filter(d=>d.id!==o.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),K(`Data SP ${o.employee_name} berhasil dihapus`),f()},"Hapus Data SP")}function k(o){let d=ze(o.akhir_sp),h=Ma(o.sp_type),c=document.createElement("div");c.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${o.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${o.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${o.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${o.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${h.bg}; color:${h.color}; border:1px solid ${h.border}; font-size:0.85rem;">
              <span>${h.icon}</span>
              <span>${h.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border}; font-size:0.85rem;">
              <span>${d.icon}</span>
              <span>${d.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${ut(o.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${ut(o.akhir_sp)}</div>
          </div>
        </div>

        ${o.notes||o.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${o.notes||o.description}</div>
          </div>
        `:""}

        ${o.document_link?`
          <div>
            <a href="${o.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,Y({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:c,size:"md",cancelText:"Tutup"})}async function w(){try{let{downloadExcel:o}=await Promise.resolve().then(()=>(X(),ye)),d=t.map((h,c)=>{let y=ze(h.akhir_sp);return{No:c+1,"Nama Karyawan":h.employee_name||"",Divisi:h.division||"",Cabang:h.branch_name||"","Jenis SP":h.sp_type||"","Tanggal SP":h.tanggal||"","Akhir SP":h.akhir_sp||"","Status Masa Berlaku":y.status,"Sisa Hari":y.daysLeft>0?y.daysLeft:0,"Link BAP / Foto":h.document_link||"","Catatan / Kronologi":h.notes||h.description||""}});o(d,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),K("Berhasil mengunduh data SP ke format Excel.")}catch(o){N("Gagal ekspor: "+o.message)}}function f(){let o=F(),d=$();e.innerHTML=`
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
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <button id="sp-btn-add" class="btn btn-primary" style="background:#2563EB; border:none; border-radius:10px; padding:10px 18px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>\u2795</span> Terbitkan SP Baru
              </button>
              <button id="sp-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 16px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>\u{1F4E5}</span> Ekspor Excel
              </button>
              <button id="sp-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 14px; font-weight:600;">
                <span>\u{1F504}</span>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="sp-stats-grid">
            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(239,68,68,0.2); color:#F87171;">\u{1F534}</div>
              <div>
                <div class="sp-stat-value">${d.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${d.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${d.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${d.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${l?l.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(h=>`<option value="${h.value}" ${String(b)===String(h.value)?"selected":""}>${h.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${p==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${p==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${p==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${r==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${d.total}</span>
            </button>
            <button class="sp-pill ${r==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${d.active}</span>
            </button>
            <button class="sp-pill ${r==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${d.expiring}</span>
            </button>
            <button class="sp-pill ${r==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${d.expired}</span>
            </button>
            <button class="sp-pill ${r==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${d.sp1}</span>
            </button>
            <button class="sp-pill ${r==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${d.sp2}</span>
            </button>
            <button class="sp-pill ${r==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${d.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${o.length===0?`
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
              ${o.map(h=>{let c=ze(h.akhir_sp),y=Ma(h.sp_type);return`
                  <div class="sp-card" data-id="${h.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border};">
                          <span>${y.icon}</span>
                          <span>${y.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${c.bg}; color:${c.color}; border:1px solid ${c.border};">
                          <span>${c.icon}</span>
                          <span>${c.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${h.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${h.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${h.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${h.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${ut(h.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${ut(h.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${h.notes||h.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${h.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${h.document_link?`
                          <a href="${h.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${h.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${h.id}">
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
    `,m()}function m(){let o=e.querySelector("#sp-btn-add");o&&(o.onclick=()=>S());let d=e.querySelector("#sp-btn-excel");d&&(d.onclick=()=>w());let h=e.querySelector("#sp-btn-refresh");h&&(h.onclick=()=>u());let c=e.querySelector("#sp-btn-reset-filter");c&&(c.onclick=()=>{l="",b="all",p="all",r="all",f()});let y=e.querySelector("#sp-search-input");y&&(y.oninput=E=>{l=E.target.value,f();let C=e.querySelector("#sp-search-input");C&&(C.focus(),C.setSelectionRange(C.value.length,C.value.length))});let s=e.querySelector("#sp-filter-branch");s&&(s.onchange=E=>{b=E.target.value,f()});let g=e.querySelector("#sp-filter-division");g&&(g.onchange=E=>{p=E.target.value,f()}),e.querySelectorAll(".sp-pill").forEach(E=>{E.onclick=()=>{r=E.getAttribute("data-filter")||"all",f()}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(E=>{E.onclick=C=>{C.stopPropagation();let B=Number(E.getAttribute("data-id")),P=t.find(D=>D.id===B);P&&k(P)}}),e.querySelectorAll(".sp-edit-btn").forEach(E=>{E.onclick=C=>{C.stopPropagation();let B=Number(E.getAttribute("data-id")),P=t.find(D=>D.id===B);P&&S(P)}}),e.querySelectorAll(".sp-delete-btn").forEach(E=>{E.onclick=C=>{C.stopPropagation();let B=Number(E.getAttribute("data-id")),P=t.find(D=>D.id===B);P&&v(P)}})}await u()}V();me();var jt=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function Bn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Na(e){let t=[],i=[],n=[],r="all",b="all",p="all",l="",a=!1,F=ne()?.role!=="viewer";try{[i,n]=await Promise.all([Q(),ce()])}catch(o){console.error("Error fetching cache:",o)}async function $(){a=!0,f();try{let o=await I("/api/mutasi?limit=10000");if(o.ok){let d=o.data?.data||o.data||[];t=Array.isArray(d)&&d.length>0?d:jt}else t=jt}catch(o){console.warn("Gagal load mutasi, menggunakan data lokal:",o),t=jt}finally{a=!1,f()}}async function S(){try{let{downloadExcel:o}=await Promise.resolve().then(()=>(X(),ye)),d=t.map((h,c)=>({No:c+1,Tanggal:h.tanggal||"","Nama Karyawan":h.employee_name||"","Cabang Asal":h.from_branch_name||"","Cabang Tujuan":h.to_branch_name||"",Status:h.status||"","Link Dokumen":h.document_link||""}));o(d,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),K("Data mutasi berhasil diekspor ke Excel!")}catch(o){N("Gagal ekspor Excel: "+o.message)}}async function v(o){if(!F)return;let d=t.find(h=>h.id===o);ee({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${d?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let h=await I(`/api/mutasi/${o}`,{method:"DELETE"});h.ok?(K("Data mutasi berhasil dihapus"),t=t.filter(c=>c.id!==o),f()):N(h.data?.error||"Gagal menghapus mutasi")}catch(h){N("Gagal menghapus data mutasi: "+h.message)}}})}function k(o=null){if(!F)return;let d=!!o,h=o?.tanggal?o.tanggal.slice(0,10):new Date().toISOString().slice(0,10),c=document.createElement("div");c.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${h}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${o?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${o?.status==="Selesai"||!o?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <input list="emp-list-options" id="mutasi-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik atau pilih nama karyawan..." value="${o?.employee_name?String(o.employee_name).replace(/"/g,"&quot;"):""}" required />
          <datalist id="emp-list-options">
            ${n.map(y=>`<option value="${typeof y=="string"?y:y.label||y.name||""}">`).join("")}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${i.map(y=>`
                <option value="${y.value}" ${String(o?.from_branch_id)===String(y.value)?"selected":""}>
                  ${y.label}
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
              ${i.map(y=>`
                <option value="${y.value}" ${String(o?.to_branch_id)===String(y.value)?"selected":""}>
                  ${y.label}
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
            placeholder="https://drive.google.com/file/d/..." value="${o?.document_link?String(o.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `,Y({title:d?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:c,size:"md",confirmText:d?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(y,s)=>{let g=c.querySelector("#mutasi-modal-tanggal").value,E=(c.querySelector("#mutasi-modal-emp").value||"").trim(),C=c.querySelector("#mutasi-modal-from").value,B=c.querySelector("#mutasi-modal-to").value,P=c.querySelector("#mutasi-modal-status").value,D=(c.querySelector("#mutasi-modal-link").value||"").trim();if(!g){N("Tanggal mutasi wajib diisi");return}if(!E){N("Nama karyawan wajib diisi");return}if(!C){N("Cabang asal wajib dipilih");return}if(!B){N("Cabang tujuan wajib dipilih");return}let _=i.find(z=>String(z.value)===String(C)),L=i.find(z=>String(z.value)===String(B)),M={tanggal:g,employee_name:E,from_branch_id:Number(C),from_branch_name:_?_.label:"Cabang Asal",to_branch_id:Number(B),to_branch_name:L?L.label:"Cabang Tujuan",status:P,document_link:D};try{if(d){let z=await I(`/api/mutasi/${o.id}`,{method:"PUT",body:JSON.stringify(M)});if(z.ok){K("Data mutasi berhasil diperbarui");let x=t.findIndex(T=>T.id===o.id);x!==-1&&(t[x]={...t[x],...M}),s(),f()}else N(z.data?.error||"Gagal memperbarui data mutasi")}else{let z=await I("/api/mutasi",{method:"POST",body:JSON.stringify(M)});if(z.ok){K("Data mutasi baru berhasil ditambahkan");let x=z.data?.id||Date.now();t.unshift({id:x,...M}),s(),f()}else N(z.data?.error||"Gagal menambahkan mutasi")}}catch(z){N("Terjadi kesalahan: "+z.message)}}})}function w(){return t.filter(o=>{if(l){let d=l.toLowerCase(),h=(o.employee_name||"").toLowerCase().includes(d),c=(o.from_branch_name||"").toLowerCase().includes(d),y=(o.to_branch_name||"").toLowerCase().includes(d);if(!h&&!c&&!y)return!1}return!(b!=="all"&&String(o.from_branch_id)!==String(b)||p!=="all"&&String(o.to_branch_id)!==String(p)||r!=="all"&&o.status!==r)})}function f(){let o=w(),d=t.length,h=t.filter(C=>C.status==="Proses").length,c=t.filter(C=>C.status==="Selesai").length,y=new Date,s=y.getFullYear(),g=y.getMonth(),E=t.filter(C=>{if(!C.tanggal)return!1;let B=new Date(C.tanggal);return!isNaN(B.getTime())&&B.getFullYear()===s&&B.getMonth()===g}).length;e.innerHTML=`
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
              ${F?`
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
                <div class="mutasi-stat-val">${d}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${h}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val">${c}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val">${E}</div>
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
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${l}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${i.length})</option>
              ${i.map(C=>`<option value="${C.value}" ${String(b)===String(C.value)?"selected":""}>${C.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${i.length})</option>
              ${i.map(C=>`<option value="${C.value}" ${String(p)===String(C.value)?"selected":""}>${C.label}</option>`).join("")}
            </select>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${r==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${d}</span>
            </div>
            <div class="mutasi-pill ${r==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${h}</span>
            </div>
            <div class="mutasi-pill ${r==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${c}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${o.length===0?`
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
              ${o.map(C=>{let B=C.status==="Selesai",P=(C.employee_name||"Karyawan").split(" ").map(D=>D[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${C.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${B?"selesai":"proses"}">
                          <span>${B?"\u2705":"\u23F3"}</span>
                          <span>${C.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${Bn(C.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${P}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${C.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${C.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${C.from_branch_name||"-"}">
                            ${C.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node" style="text-align:right;">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${C.to_branch_name||"-"}">
                            ${C.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${C.document_link?`
                          <a href="${C.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${F?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${C.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${C.id}">
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
    `,m()}function m(){let o=e.querySelector("#mutasi-btn-add");o&&(o.onclick=()=>k());let d=e.querySelector("#mutasi-btn-excel");d&&(d.onclick=()=>S());let h=e.querySelector("#mutasi-btn-refresh");h&&(h.onclick=()=>$());let c=e.querySelector("#mutasi-btn-reset-filter");c&&(c.onclick=()=>{l="",b="all",p="all",r="all",f()});let y=e.querySelector("#mutasi-search-input");y&&(y.oninput=E=>{l=E.target.value,f();let C=e.querySelector("#mutasi-search-input");C&&(C.focus(),C.setSelectionRange(C.value.length,C.value.length))});let s=e.querySelector("#mutasi-filter-from");s&&(s.onchange=E=>{b=E.target.value,f()});let g=e.querySelector("#mutasi-filter-to");g&&(g.onchange=E=>{p=E.target.value,f()}),e.querySelectorAll(".mutasi-pill").forEach(E=>{E.onclick=()=>{r=E.getAttribute("data-status")||"all",f()}}),e.querySelectorAll(".mutasi-edit-btn").forEach(E=>{E.onclick=C=>{C.stopPropagation();let B=Number(E.getAttribute("data-id")),P=t.find(D=>D.id===B);P&&k(P)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(E=>{E.onclick=C=>{C.stopPropagation();let B=Number(E.getAttribute("data-id"));v(B)}})}await $()}V();me();var Ot=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],Ka=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function ja(e,t,i){if(!e||!t)return 0;try{let[n,r]=e.split(":").map(Number),[b,p]=t.split(":").map(Number),l=b*60+p-(n*60+r);l<0&&(l+=24*60);let a=l/60-(parseFloat(i)||0);return a<0&&(a=0),Number.isInteger(a)?a:parseFloat(a.toFixed(2))}catch{return 0}}function Oa(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function Ra(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function qa(e){let t=[],i=[],n=[],r="all",b="all",p="",l=!1,a=ne(),u=a?.role==="viewer",F=a?.role==="input_lembur",$=!u,S=!u,v=!u&&!F;try{[i,n]=await Promise.all([Q(),_e()])}catch(c){console.error("Error fetching cache:",c)}async function k(){l=!0,d();try{let c=await I("/api/overtime?limit=10000");if(c.ok){let y=c.data?.data||c.data||[];t=Array.isArray(y)&&y.length>0?y:Ot}else t=Ot}catch(c){console.warn("Gagal load lembur, menggunakan data lokal:",c),t=Ot}finally{l=!1,d()}}async function w(){try{let{downloadExcel:c}=await Promise.resolve().then(()=>(X(),ye)),y=t.map((s,g)=>({No:g+1,Tanggal:s.date||"","Cabang Lembur":s.branch_name||"","Nama Karyawan":s.employee_name||"","Jam Mulai":s.start_time||"","Jam Selesai":s.end_time||"","Istirahat (Jam)":s.break_hours!=null?s.break_hours:0,"Total Jam":s.total_hours!=null?s.total_hours:0,"Alasan / Keterangan":s.reason||""}));c(y,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),K("Data lembur berhasil diekspor ke Excel!")}catch(c){N("Gagal ekspor Excel: "+c.message)}}async function f(c){if(!v)return;let y=t.find(s=>s.id===c);ee({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${y?.employee_name||"karyawan ini"}" pada tanggal ${Ra(y?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let s=await I(`/api/overtime/${c}`,{method:"DELETE"});s.ok?(K("Catatan lembur berhasil dihapus"),t=t.filter(g=>g.id!==c),d()):N(s.data?.error||"Gagal menghapus lembur")}catch(s){N("Gagal menghapus catatan lembur: "+s.message)}}})}function m(c=null){if(!S)return;let y=!!c,s=c?.date?c.date.slice(0,10):new Date().toISOString().slice(0,10),g=c?.start_time||"22:00",E=c?.end_time||"08:00",C=c?.break_hours!=null?c.break_hours:1,B=c?.total_hours!=null?c.total_hours:ja(g,E,C),P=document.createElement("div");P.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${s}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(x=>`
                <option value="${x.value}" ${String(c?.branch_id)===String(x.value)?"selected":""}>
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
            ${n.map(x=>{let T=typeof x=="string"?x:x.label||x.full_name||x.name||"",A=typeof x=="object"&&x!==null?x.value!=null?x.value:x.id!=null?x.id:T:x,q=String(c?.employee_id)===String(A)||c?.employee_name===T;return`<option value="${A}" ${q?"selected":""}>${T}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${g}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${E}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${C}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${B}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${Ka.map(x=>`
              <option value="${x}" ${c?.reason===x?"selected":""}>${x}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let D=P.querySelector("#ot-modal-start"),_=P.querySelector("#ot-modal-end"),L=P.querySelector("#ot-modal-break"),M=P.querySelector("#ot-modal-total");function z(){let x=D.value,T=_.value,A=L.value;M.value=ja(x,T,A)}D.addEventListener("input",z),_.addEventListener("input",z),L.addEventListener("input",z),Y({title:y?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:P,size:"md",confirmText:y?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(x,T)=>{let A=P.querySelector("#ot-modal-date").value,q=P.querySelector("#ot-modal-branch").value,O=P.querySelector("#ot-modal-emp").value,G=P.querySelector("#ot-modal-start").value,j=P.querySelector("#ot-modal-end").value,R=parseFloat(P.querySelector("#ot-modal-break").value)||0,ie=parseFloat(P.querySelector("#ot-modal-total").value)||0,de=P.querySelector("#ot-modal-reason").value;if(!A){N("Tanggal lembur wajib diisi");return}if(!q){N("Cabang penugasan wajib dipilih");return}if(!O){N("Nama karyawan wajib dipilih");return}if(!G||!j){N("Jam mulai dan jam selesai wajib diisi");return}let ge=i.find(H=>String(H.value)===String(q)),U=n.find(H=>String(H.value)===String(O)||String(H.id)===String(O)),re=U?U.label||U.full_name||U.name||"Karyawan":typeof O=="string"&&isNaN(Number(O))?O:"Karyawan",oe={date:A,branch_id:Number(q),branch_name:ge?ge.label:"Cabang Lembur",employee_id:Number(O),employee_name:U?U.label:"Karyawan",start_time:G,end_time:j,break_hours:R,total_hours:ie,reason:de};try{if(y){let H=await I(`/api/overtime/${c.id}`,{method:"PUT",body:JSON.stringify(oe)});if(H.ok){K("Data lembur berhasil diperbarui");let pe=t.findIndex(fe=>fe.id===c.id);pe!==-1&&(t[pe]={...t[pe],...oe}),T(),d()}else N(H.data?.error||"Gagal memperbarui data lembur")}else{let H=await I("/api/overtime",{method:"POST",body:JSON.stringify(oe)});if(H.ok){K("Data lembur berhasil ditambahkan");let pe=H.data?.id||Date.now();t.unshift({id:pe,...oe}),T(),d()}else N(H.data?.error||"Gagal menambahkan lembur")}}catch(H){N("Terjadi kesalahan: "+H.message)}}})}function o(){return t.filter(c=>{if(p){let y=p.toLowerCase(),s=(c.employee_name||"").toLowerCase().includes(y),g=(c.branch_name||"").toLowerCase().includes(y),E=(c.reason||"").toLowerCase().includes(y);if(!s&&!g&&!E)return!1}return!(b!=="all"&&String(c.branch_id)!==String(b)||r!=="all"&&c.reason!==r)})}function d(){let c=o(),y=t.length,s=t.reduce((D,_)=>D+(parseFloat(_.total_hours)||0),0),g=y>0?(s/y).toFixed(1):0,E=new Date,C=E.getFullYear(),B=E.getMonth(),P=t.filter(D=>{if(!D.date)return!1;let _=new Date(D.date);return!isNaN(_.getTime())&&_.getFullYear()===C&&_.getMonth()===B}).length;e.innerHTML=`
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
              ${$?`
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
                <div class="lembur-stat-val">${s} Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val">${y}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val">${g} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${P}</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${p}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(D=>`<option value="${D.value}" ${String(b)===String(D.value)?"selected":""}>${D.label}</option>`).join("")}
            </select>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${r==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${y}</span>
            </div>
            ${Ka.map(D=>{let _=t.filter(M=>M.reason===D).length,L=Oa(D);return`
                <div class="lembur-pill ${r===D?"active":""}" data-reason="${D}">
                  <span>${L.icon}</span>
                  <span>${D}</span>
                  ${_>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${_}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${c.length===0?`
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
              ${c.map(D=>{let _=(D.employee_name||"Karyawan").split(" ").map(M=>M[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),L=Oa(D.reason);return`
                  <div class="lembur-card" data-id="${D.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${D.total_hours!=null?D.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${D.break_hours!=null?D.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${_}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${D.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${D.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${Ra(D.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${D.start_time||"00:00"} - ${D.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${L.bg}; color:${L.color}; border:1px solid ${L.border};">
                          <span>${L.icon}</span>
                          <span>${D.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${S?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${D.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        ${v?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${D.id}">
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
    `,h()}function h(){let c=e.querySelector("#lembur-btn-add");c&&(c.onclick=()=>m());let y=e.querySelector("#lembur-btn-excel");y&&(y.onclick=()=>w());let s=e.querySelector("#lembur-btn-refresh");s&&(s.onclick=()=>k());let g=e.querySelector("#lembur-btn-reset-filter");g&&(g.onclick=()=>{p="",b="all",r="all",d()});let E=e.querySelector("#lembur-search-input");E&&(E.oninput=B=>{p=B.target.value,d();let P=e.querySelector("#lembur-search-input");P&&(P.focus(),P.setSelectionRange(P.value.length,P.value.length))});let C=e.querySelector("#lembur-filter-branch");C&&(C.onchange=B=>{b=B.target.value,d()}),e.querySelectorAll(".lembur-pill").forEach(B=>{B.onclick=()=>{r=B.getAttribute("data-reason")||"all",d()}}),e.querySelectorAll(".lembur-edit-btn").forEach(B=>{B.onclick=P=>{P.stopPropagation();let D=Number(B.getAttribute("data-id")),_=t.find(L=>L.id===D);_&&m(_)}}),e.querySelectorAll(".lembur-delete-btn").forEach(B=>{B.onclick=P=>{P.stopPropagation();let D=Number(B.getAttribute("data-id"));f(D)}})}await k()}V();async function Ha(e){let t=ne();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}W({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,n)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${n.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,n)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let b=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(u=>String(u.id)===String(i));if(!b)return alert("Data tidak ditemukan");let p=u=>{if(!u)return"Tidak ada data";try{return JSON.stringify(JSON.parse(u),null,2)}catch{return u}},l=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${p(b.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${p(b.new_data)}</pre>
           </div>
         </div>
       `,{createModal:a}=await Promise.resolve().then(()=>(me(),vt));a({title:`Detail Audit Log #${i}`,content:l,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}V();me();X();var Rt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ye={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function Ja(e){e.innerHTML=`
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
  `;let t="board",i=null,n=[],r=ne(),b=r&&["superadmin","admin","manager","spv"].includes(r.role);I("/api/branches?all=1").then(m=>{m.ok&&m.data?.data&&(n=m.data.data)}).catch(()=>{});async function p(){try{let m=await I("/api/os-migrations/summary");if(!m.ok)throw new Error("Gagal memuat data migrasi");i=m.data?.data||m.data,l(i),u()}catch(m){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${m.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(m){let o=document.getElementById("migration-summary-cards");if(!o||!m)return;let d=m.total_branches||0,h=m.vendors?.AEON?.total||0,c=m.vendors?.["Triple S"]?.total||0,y=m.by_status?.completed||0,s=m.by_status?.in_progress||0,g=m.by_status?.scheduled||0,E=d>0?Math.round(y/d*100):0;o.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${d} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${h} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${c} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${E}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${E}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${y}</strong></span>
          <span>\u{1F535} Proses: <strong>${s}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${g}</strong></span>
        </div>
      </div>
    `}function a(){if(!i||!i.matrix)return[];let m=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),o=document.getElementById("filter-vendor")?.value||"",d=document.getElementById("filter-month")?.value||"",h=document.getElementById("filter-status")?.value||"";return i.matrix.filter(c=>!(o&&c.vendor!==o||d&&c.target_month!==d||h&&c.status!==h||m&&!`${c.branch_code} ${c.branch_name} ${c.branch_full_name} ${c.vendor} ${c.notes||""} ${c.pic_vendor||""}`.toLowerCase().includes(m)))}function u(){let m=a(),o=document.getElementById("filter-stats-text");o&&(o.innerHTML=`Menampilkan <strong>${m.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let d=document.getElementById("migration-view-container");t==="board"?F(d,m):$(d,m)}function F(m,o){if(!o.length){m.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let d=["AEON","Triple S"],h=["2026-09","2026-10","2026-11","2026-12"],c="";d.forEach(y=>{let s=o.filter(D=>D.vendor===y);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==y||s.length===0)return;let g=s.length,E=y==="AEON"?"#1E40AF":"#0F766E",C=y==="AEON"?"#EFF6FF":"#F0FDFA",B=y==="AEON"?"#BFDBFE":"#99F6E4",P=h.filter(D=>s.some(_=>_.target_month===D));c+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${C}; border-bottom:1px solid ${B}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${E};">${y==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${E}; border:1px solid ${B}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${g} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${P.map(D=>{let _=s.filter(z=>z.target_month===D);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Rt[D]||D).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${_.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${_.map(z=>{let x=Ye[z.status]||Ye.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${z.branch_code}.</span>${z.branch_name}
                            </div>
                            ${b?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${z.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${z.id}" data-name="${z.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${z.id}" data-status="${z.status}" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${x.icon} ${x.label} \u25BE
                              </button>
                            </div>

                            ${z.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${z.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${z.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${z.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${z.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),m.innerHTML=c,S(m)}function $(m,o){if(!o.length){m.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}m.innerHTML=`
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
                ${b?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${o.map((d,h)=>{let c=Ye[d.status]||Ye.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${h+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${d.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${d.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${d.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${d.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Rt[d.target_month]||d.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${d.id}" data-status="${d.status}" style="background:${c.bg}; color:${c.color}; border:1px solid ${c.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${c.icon} ${c.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${d.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${d.employee_count>0?`\u{1F464} ${d.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${d.notes||"-"}
                    </td>
                    ${b?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${d.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${d.id}" data-name="${d.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,S(m)}function S(m){m.querySelectorAll(".status-toggle-btn").forEach(o=>{o.addEventListener("click",d=>{d.stopPropagation();let h=o.dataset.id,c=o.dataset.status;v(o,h,c)})}),m.querySelectorAll(".edit-card-btn").forEach(o=>{o.addEventListener("click",()=>{let d=o.dataset.id,h=i?.matrix?.find(c=>String(c.id)===String(d));h&&k(h)})}),m.querySelectorAll(".delete-card-btn").forEach(o=>{o.addEventListener("click",()=>{let d=o.dataset.id,h=o.dataset.name||"Cabang";ee(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${h}</strong>?`,async()=>{try{if(!(await I(`/api/os-migrations/${d}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");K(`Jadwal migrasi ${h} berhasil dihapus`),p()}catch(c){N(c.message)}})})})}function v(m,o,d){document.querySelectorAll(".status-popup-menu").forEach(g=>g.remove());let h=m.getBoundingClientRect(),c=document.createElement("div");c.className="status-popup-menu",c.style.cssText=`
      position: fixed;
      top: ${h.bottom+window.scrollY+4}px;
      left: ${h.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let y=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];c.innerHTML=y.map(g=>`
      <button class="status-option-item" data-key="${g.key}" style="
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
        color: ${g.key===d?"#2563EB":"var(--text-1)"};
        font-weight: ${g.key===d?"700":"500"};
      ">
        ${g.label} ${g.key===d?"\u2713":""}
      </button>
    `).join("");let s=g=>{!c.contains(g.target)&&g.target!==m&&(c.remove(),document.removeEventListener("click",s))};setTimeout(()=>document.addEventListener("click",s),50),c.querySelectorAll(".status-option-item").forEach(g=>{g.addEventListener("click",async()=>{let E=g.dataset.key;if(c.remove(),document.removeEventListener("click",s),E!==d)try{if(!(await I(`/api/os-migrations/${o}/status`,{method:"PATCH",body:JSON.stringify({status:E})})).ok)throw new Error("Gagal memperbarui status");K("Status berhasil diubah"),p()}catch(C){N(C.message)}})}),document.body.appendChild(c)}function k(m=null){let o=!!m,d=o?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",c=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${n.map(y=>`
      <option value="${y.id}" ${m&&m.branch_id===y.id?"selected":""}>
        ${y.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${m?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${m?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${m&&!["AEON","Triple S"].includes(m.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${m?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${m?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${m?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${m?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${m?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${m?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${m?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${m?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${m?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${m?.notes||""}</textarea>
        </div>
      </form>
    `;Y({title:d,content:c,confirmText:o?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(y,s)=>{let g=y.querySelector("#migration-form");if(!g.checkValidity()){g.reportValidity();return}let E=new FormData(g),C={branch_id:parseInt(E.get("branch_id"),10),vendor:E.get("vendor"),target_month:E.get("target_month"),status:E.get("status"),actual_date:E.get("actual_date")||null,pic_vendor:E.get("pic_vendor")||null,notes:E.get("notes")||null};try{let B=o?`/api/os-migrations/${m.id}`:"/api/os-migrations",D=await I(B,{method:o?"PUT":"POST",body:JSON.stringify(C)});if(!D.ok){let _=await D.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan data")}K(o?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),s(),p()}catch(B){N(B.message)}}})}let w=document.getElementById("view-board-btn"),f=document.getElementById("view-table-btn");w.addEventListener("click",()=>{t="board",w.style.background="#fff",w.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),u()}),f.addEventListener("click",()=>{t="table",f.style.background="#fff",f.classList.remove("btn-ghost"),w.style.background="transparent",w.classList.add("btn-ghost"),u()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(m=>{let o=document.getElementById(m);o&&(o.addEventListener("input",u),o.addEventListener("change",u))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{k()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let m=a();if(!m.length){N("Tidak ada data untuk diexport");return}let o=m.map((d,h)=>({No:h+1,"Kode Cabang":d.branch_code,"Nama Cabang":d.branch_name,"Cabang Lengkap":d.branch_full_name,"Vendor OS":d.vendor,"Target Bulan":Rt[d.target_month]||d.target_month,Status:Ye[d.status]?.label||d.status,"Tanggal Realisasi":d.actual_date||"","Jml Personil FC":d.employee_count||0,"PIC Vendor":d.pic_vendor||"",Catatan:d.notes||""}));J(o,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),K("File Excel berhasil diunduh")}),p()}V();me();X();var Ga={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Ua(e){e.innerHTML=`
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
  `;let t="pagi",i="timeline",n=[],r=ne(),b=r&&["superadmin","admin","manager","spv"].includes(r.role);async function p(){try{let m=await I("/api/dac?all=1");if(!m.ok)throw new Error("Gagal memuat standar DAC");n=m.data?.data||m.data||[];let o=n.filter(y=>y.shift==="pagi").length,d=n.filter(y=>y.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=o,document.getElementById("badge-count-siang").textContent=d,document.getElementById("badge-count-all").textContent=n.length;let h=document.getElementById("dac-category-filter"),c=Array.from(new Set(n.map(y=>y.category))).sort();h.innerHTML='<option value="">Semua Kategori</option>'+c.map(y=>`<option value="${y}">${y}</option>`).join(""),a(),u()}catch(m){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${m.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(){let m=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),o=document.getElementById("dac-category-filter")?.value||"",d=document.getElementById("dac-period-filter")?.value||"";return n.filter(h=>{if(t!=="all"&&h.shift!==t||o&&h.category!==o)return!1;if(d==="morning"){let c=parseInt(h.start_time.split(":")[0],10);if(h.shift==="pagi"&&c>=12||h.shift==="siang"&&c>=17)return!1}else if(d==="afternoon"){let c=parseInt(h.start_time.split(":")[0],10);if(h.shift==="pagi"&&c<12||h.shift==="siang"&&c<17)return!1}return!(m&&!`${h.start_time} ${h.end_time} ${h.category} ${h.location} ${h.activity}`.toLowerCase().includes(m))})}function a(){let m=document.getElementById("dac-summary-cards");if(!m)return;let o=n.filter(s=>s.shift==="pagi"),d=n.filter(s=>s.shift==="siang"),h=t==="pagi"?o:t==="siang"?d:n,c=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",y=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";m.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${y})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${h.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${c}</div>
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
    `}function u(){let m=l(),o=document.getElementById("dac-stats-text");o&&(o.innerHTML=`Menampilkan <strong>${m.length}</strong> aktivitas`);let d=document.getElementById("dac-view-container");i==="timeline"?F(d,m):$(d,m)}function F(m,o){if(!o.length){m.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}m.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${o.map(d=>{let h=Ga[d.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},c=d.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",y=d.shift==="pagi"?"#FEF3C7":"#E0E7FF",s=d.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${d.start_time} - ${d.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${d.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${y}; color:${s}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${c}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${h.bg}; color:${h.color}; border:1px solid ${h.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${d.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${d.location}
                    </span>
                  </div>
                  ${b?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${d.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${d.id}" data-name="${d.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${h.color}; margin-top:8px;">
                  ${d.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,S(m)}function $(m,o){if(!o.length){m.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}m.innerHTML=`
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
                ${b?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${o.map(d=>{let h=Ga[d.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${d.item_no}</td>
                    ${t==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${d.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${d.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${d.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${d.start_time} - ${d.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${h.bg}; color:${h.color}; border:1px solid ${h.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${d.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${d.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${d.activity}
                    </td>
                    ${b?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${d.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${d.id}" data-name="${d.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,S(m)}function S(m){m.querySelectorAll(".edit-dac-btn").forEach(o=>{o.addEventListener("click",()=>{let d=o.dataset.id,h=n.find(c=>String(c.id)===String(d));h&&v(h)})}),m.querySelectorAll(".delete-dac-btn").forEach(o=>{o.addEventListener("click",()=>{let d=o.dataset.id,h=o.dataset.name||"item";ee(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${h}</strong>?`,async()=>{try{if(!(await I(`/api/dac/${d}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");K("Aktivitas DAC berhasil dihapus"),p()}catch(c){N(c.message)}})})})}function v(m=null){let o=!!m,d=o?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",h=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],c=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${m?.shift==="pagi"||!m&&t==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${m?.shift==="siang"||!m&&t==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${m?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${m?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${m?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${m?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${h.map(y=>`<option value="${y}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${m?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${m?.activity||""}</textarea>
        </div>
      </form>
    `;Y({title:d,content:c,confirmText:o?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(y,s)=>{let g=y.querySelector("#dac-form");if(!g.checkValidity()){g.reportValidity();return}let E=new FormData(g),C={shift:E.get("shift"),item_no:E.get("item_no")?parseInt(E.get("item_no"),10):void 0,start_time:E.get("start_time"),end_time:E.get("end_time"),category:E.get("category"),location:E.get("location"),activity:E.get("activity")};try{let B=o?`/api/dac/${m.id}`:"/api/dac",D=await I(B,{method:o?"PUT":"POST",body:JSON.stringify(C)});if(!D.ok){let _=await D.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan aktivitas")}K(o?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),s(),p()}catch(B){N(B.message)}}})}function k(){let m=l();if(!m.length){N("Tidak ada data untuk dicetak");return}let o=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",d=window.open("","_blank");if(!d){N("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}d.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${o}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${o}</div>
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
            ${m.map(h=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${h.item_no}</td>
                ${t==="all"?`<td>${h.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${h.start_time} - ${h.end_time}</td>
                <td><strong>${h.category}</strong></td>
                <td><strong>${h.location}</strong></td>
                <td>${h.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),d.document.close(),d.focus(),setTimeout(()=>{d.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(m=>{m.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(o=>{o.style.background="#F1F5F9",o.style.color="var(--text-2)"}),m.style.background="#2563EB",m.style.color="#fff",t=m.dataset.shift,a(),u()})});let w=document.getElementById("view-timeline-btn"),f=document.getElementById("view-table-btn");w.addEventListener("click",()=>{i="timeline",w.style.background="#fff",w.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),u()}),f.addEventListener("click",()=>{i="table",f.style.background="#fff",f.classList.remove("btn-ghost"),w.style.background="transparent",w.classList.add("btn-ghost"),u()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(m=>{let o=document.getElementById(m);o&&(o.addEventListener("input",u),o.addEventListener("change",u))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{v()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{k()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let m=l();if(!m.length){N("Tidak ada data untuk diexport");return}let o=m.map(d=>({No:d.item_no,Shift:d.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":d.start_time,"Jam Selesai":d.end_time,Kategori:d.category,"Lokasi / Area":d.location,"Rincian Aktivitas":d.activity}));J(o,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),K("File Excel berhasil diunduh")}),p()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],n=parseInt(t[2],10),r=parseInt(t[3],10);if(n>12&&r<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let n=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),b=n[parseInt(i[1],10)-1];return`${r} ${b} ${i[0]}`}return t};function te(e){return async t=>{if(!je()){$e("/login");return}return e(t)}}var We=null;function Pn(){We&&clearInterval(We);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),b=document.getElementById("header-clock-date");r&&(r.textContent=i),b&&(b.textContent=n)};e(),We=setInterval(e,1e3)}async function Ln(){try{let e=await I("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(n,r)=>{let b=document.getElementById(n);b&&(b.textContent=r>0?r:"",b.style.display=r>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var Ne=[];async function In(){try{let e=await I("/api/dashboard/notifications");if(!e.ok)return;Ne=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=Ne.length>0?"block":"none",t.textContent=Ne.length)}catch{}}function Mn(){if(!Ne.length){Y({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ne.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;Y({title:`Notifikasi (${Ne.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function Va(){let e=ne(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let u=new Date().getHours();return u>=4&&u<11?"Selamat Pagi":u>=11&&u<15?"Selamat Siang":u>=15&&u<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),b=document.getElementById("sidebar-close"),p=()=>{i.classList.add("open"),n.classList.add("show")},l=()=>{i.classList.remove("open"),n.classList.remove("show")};r?.addEventListener("click",p),b?.addEventListener("click",l),n?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(u=>u.addEventListener("click",l)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(u=>{u.dataset.route!=="/overtime"&&(u.style.display="none")}),document.querySelectorAll(".nav-section").forEach(u=>{Array.from(u.querySelectorAll(".nav-item")).some($=>$.style.display!=="none")||(u.style.display="none")}));function a(){let u=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(S=>{let v=S.dataset.route;S.classList.toggle("active",u===v||v!=="/dashboard"&&u.startsWith(v))});let F=document.getElementById("topbar-title"),$=document.querySelector(".nav-item.active .nav-label");F&&$&&(F.textContent=$.textContent)}window.addEventListener("hashchange",a),a(),Pn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await I("/api/auth/logout",{method:"POST"}),Oe(),We&&clearInterval(We),$e("/login")}),Ln(),In(),document.getElementById("btn-notif")?.addEventListener("click",u=>{u.preventDefault(),Mn()})}async function zn(){Z("/login",({main:t})=>Zt(t)),Z("/dashboard",te(({main:t})=>Vt(t))),Z("/calendar",te(({main:t})=>_a(t))),Z("/employees",te(({main:t,params:i})=>ea(t,i))),Z("/contracts",te(({main:t,params:i})=>lt(t,i))),Z("/sp",te(({main:t})=>za(t))),Z("/mutasi",te(({main:t})=>Na(t))),Z("/overtime",te(({main:t})=>qa(t))),Z("/timeline",te(({main:t,params:i})=>aa(t,i))),Z("/os-migrations",te(({main:t})=>Ja(t))),Z("/issues",te(({main:t,params:i})=>na(t,i))),Z("/one-on-one",te(({main:t,params:i})=>ia(t,i))),Z("/training",te(({main:t})=>ra(t))),Z("/relievers",te(({main:t,params:i})=>sa(t,i))),Z("/reports/inspection",te(({main:t})=>la(t))),Z("/reports/cleaning",te(({main:t})=>da(t))),Z("/reports/fogging",te(({main:t})=>pa(t))),Z("/reports/basecamp",te(({main:t})=>ca(t))),Z("/reports/supply",te(({main:t})=>Ca(t,"supply"))),Z("/sop",()=>{window.location.hash="#/sop-fc"}),Z("/sop-fc",te(({main:t})=>ba(t))),Z("/sop-security",te(({main:t})=>xa(t))),Z("/checklist",te(({main:t})=>Sa(t))),Z("/hygiene-standards",te(({main:t})=>Ea(t))),Z("/dac",te(({main:t})=>Ua(t))),Z("/users",te(({main:t})=>Fa(t))),Z("/branches",te(({main:t})=>$a(t))),Z("/profile",te(({main:t})=>Ta(t))),Z("/settings/import",te(({main:t})=>La(t))),Z("/audit-logs",te(({main:t})=>Ha(t)));let e=je();if(!e&&window.location.hash!=="#/login"&&$e("/login"),e){let t=await I("/api/auth/me");if(t.ok){Re(t.data.data),Va();let i=ne(),n=window.location.hash;i&&i.role==="input_lembur"&&(n==="#/dashboard"||n===""||n==="#/")&&$e("/overtime")}else Oe(),$e("/login")}window.addEventListener("fm:login",()=>{Va();let t=ne();t&&t.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),Jt()}zn();
