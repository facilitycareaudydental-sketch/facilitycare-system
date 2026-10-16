var _a=Object.defineProperty;var bt=(t,e)=>()=>(t&&(e=t(t=0)),e);var ht=(t,e)=>{for(var i in e)_a(t,i,{get:e[i],enumerable:!0})};var Ie={};ht(Ie,{API:()=>Bt,CLIENT_SIDE_MAX_ROWS:()=>De,IS_DEVELOPMENT:()=>ze,apiFetch:()=>T,clearToken:()=>qe,getToken:()=>je,getUser:()=>ee,setToken:()=>ft,setUser:()=>He});function je(){return localStorage.getItem("fm_token")}function ft(t){localStorage.setItem("fm_token",t)}function qe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ee(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function He(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function T(t,e={}){let i=je(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let r=`cb=${Date.now()}`,d=t.includes("?")?"&":"?",o=`${Bt}${t}${d}${r}`,l=await fetch(o,{...e,headers:a}),n;try{let s=await l.text();try{n=JSON.parse(s)}catch{n={error:`Server Error (${l.status}): ${s.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return l.status===401&&(qe(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:n}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var ze,$a,Bt,De,M=bt(()=>{ze=!1,$a="https://fm-operations-api.facilitycare-audydental.workers.dev",Bt=$a,De=1e4});var vt={};ht(vt,{confirmDialog:()=>ue,createModal:()=>oe});function oe({title:t,content:e,onConfirm:i,onCancel:a,confirmText:r="Simpan",cancelText:d="Batal",size:o="md",confirmClass:l="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},s=document.createElement("div");s.className="modal-overlay",s.innerHTML=`
    <div class="modal" style="max-width:${n[o]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${i?`<button class="btn ${l} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&s.querySelector(".modal-body").appendChild(e);let y=()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),250)};return s.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),y()}),s.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),y()}),i&&s.querySelector(".modal-confirm").addEventListener("click",()=>i(s,y)),s.addEventListener("click",v=>{v.target===s&&(a&&a(),y())}),document.body.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),{overlay:s,close:y}}function ue(t,e,i="Konfirmasi"){return oe({title:i,content:`<p>${t}</p>`,onConfirm:(a,r)=>{e(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ge=bt(()=>{});var le={};ht(le,{downloadExcel:()=>B,parseExcel:()=>Ge,renderExcelButtons:()=>Ca});function Ge(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=r=>{try{let d=new Uint8Array(r.target.result),o=XLSX.read(d,{type:"array"}),l=o.SheetNames[0],n=o.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let s=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),y=s.e.r-s.s.r+1,v=s.e.c-s.s.c+1;console.log(`Total Rows (including empty): ${y}`),console.log(`Total Columns: ${v}`);let g=[];for(let w=s.s.c;w<=s.e.c;++w){let f=n[XLSX.utils.encode_cell({c:w,r:s.s.r})];f&&f.v&&g.push(f.v)}console.log(`Headers Found: ${g.join(", ")}`),console.log("---------------------------");let p=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(p,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(p,"__headers",{value:g,enumerable:!1}),e(p)}catch(d){i(d)}},a.onerror=r=>i(r),a.readAsArrayBuffer(t)})}function B(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Ca(t){return`
    <div class="excel-actions-dropdown" style="position:relative; display:inline-block;">
      <button class="btn" id="btn-aksi-${t}" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('excel-menu-${t}').classList.toggle('show-aksi-menu')">
        \u22EE Aksi
      </button>
      <div id="excel-menu-${t}" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:180px; z-index:999; padding:8px 0;">
        <button class="dropdown-item" id="btn-export-${t}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
          \u{1F4E5} Export Excel
        </button>
        <button class="dropdown-item" id="btn-template-${t}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
          \u{1F4C4} Download Template
        </button>
        <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${t}">
          \u{1F4E4} Import Excel
          <input type="file" id="input-import-${t}" accept=".xlsx, .xls, .csv" style="display:none;">
        </label>
      </div>
    </div>
    <style>
      .show-aksi-menu { display: flex !important; }
      .dropdown-item:hover { background-color: #F8FAFC !important; color: #2563EB !important; }
    </style>
  `}var z=bt(()=>{});M();var yt={},Ze=null;function J(t,e){yt[t]=e}function Ce(t){window.location.hash=t}function Lt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),r=yt[i];if(!r){for(let[o,l]of Object.entries(yt))if(o.endsWith("/*")&&i.startsWith(o.slice(0,-2))){r=l;break}}Ze&&(Ze(),Ze=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let o=new URLSearchParams(a.join("?")),l=i.split("/").filter(Boolean),n=await r({path:i,params:o,segments:l,main:d});n&&(Ze=n)}else{let o=d||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Je;function Ea(){return Je||(Je=document.createElement("div"),Je.id="toast-container",document.body.appendChild(Je)),Je}function At(t,e="info",i=3500){let a=Ea(),r=document.createElement("div");r.className=`toast toast-${e}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${d[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var j=t=>At(t,"success"),O=t=>At(t,"error");ge();M();M();ge();function et({columns:t,data:e,onEdit:i,onDelete:a,onView:r,actions:d=[],emptyText:o="Tidak ada data",bulkSelect:l=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,n;let s=document.createElement("table");s.className="data-table";let y=document.createElement("thead"),v=document.createElement("tr");if(l){let p=document.createElement("th");p.style.width="40px",p.style.textAlign="center";let w=document.createElement("input");w.type="checkbox",w.id="select-all-checkbox",w.title="Pilih semua",w.addEventListener("change",()=>{e.forEach(f=>{w.checked?l.selectedIds.add(f.id):l.selectedIds.delete(f.id)}),n.querySelectorAll(".row-checkbox").forEach(f=>f.checked=w.checked),l.onToggle()}),p.appendChild(w),v.appendChild(p)}if(t.forEach(p=>{let w=document.createElement("th");w.textContent=p.label,p.width&&(w.style.width=p.width),v.appendChild(w)}),i||a||r||d.length>0){let p=document.createElement("th");p.textContent="Aksi",p.style.width="120px",v.appendChild(p)}y.appendChild(v),s.appendChild(y);let g=document.createElement("tbody");return e.forEach(p=>{let w=document.createElement("tr");if(l){let f=document.createElement("td");f.style.textAlign="center",f.style.width="40px";let b=document.createElement("input");b.type="checkbox",b.className="row-checkbox",b.checked=l.selectedIds.has(p.id),b.addEventListener("change",()=>{if(b.checked)l.selectedIds.add(p.id);else{l.selectedIds.delete(p.id);let c=document.getElementById("select-all-checkbox");c&&(c.checked=!1)}l.onToggle()}),f.appendChild(b),w.appendChild(f)}if(t.forEach(f=>{let b=document.createElement("td");if(f.render){let c=f.render(p[f.key],p);c instanceof HTMLElement?b.appendChild(c):b.innerHTML=c||""}else b.textContent=p[f.key]!==null&&p[f.key]!==void 0&&p[f.key]!==""?p[f.key]:"";f.nowrap&&(b.style.whiteSpace="nowrap"),w.appendChild(b)}),i||a||r||d.length>0){let f=document.createElement("td");f.className="actions-cell";let b=document.createElement("div");if(b.className="btn-group",r){let c=document.createElement("button");c.className="btn btn-xs btn-ghost",c.innerHTML="\u{1F441}",c.title="Lihat",c.addEventListener("click",()=>r(p)),b.appendChild(c)}if(i){let c=document.createElement("button");c.className="btn btn-xs btn-secondary",c.innerHTML="\u270F\uFE0F",c.title="Edit",c.addEventListener("click",()=>i(p)),b.appendChild(c)}d.forEach(c=>{let h=document.createElement("button");h.className=`btn btn-xs ${c.class||"btn-ghost"}`,h.innerHTML=c.icon||c.label,h.title=c.label,h.addEventListener("click",()=>c.handler(p)),b.appendChild(h)}),f.appendChild(b),w.appendChild(f)}g.appendChild(w)}),s.appendChild(g),n.appendChild(s),n}function tt({page:t,pages:e,total:i,limit:a,onPage:r}){if(e<=1)return null;let d=document.createElement("div");d.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${i} data`,d.appendChild(o);let l=document.createElement("div");l.className="pagination-btns";let n=(v,g,p=!1,w=!1)=>{let f=document.createElement("button");f.className=`btn btn-sm ${w?"btn-primary":"btn-ghost"} pagination-btn`,f.textContent=v,f.disabled=p,f.addEventListener("click",()=>r(g)),l.appendChild(f)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let s=Math.max(1,t-2),y=Math.min(e,t+2);for(let v=s;v<=y;v++)n(v,v,!1,v===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),d.appendChild(l),d}ge();function Ue(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Ue(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(e.type){case"textarea":r=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(p=>{let w=typeof p=="object"?p.value:p,f=typeof p=="object"?p.label:p,b=e.value==w?"selected":"";return`<option value="${w}" ${b}>${f}</option>`}).join("");r=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let l=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(p=>{let w=typeof p=="object"?p.value:p,f=typeof p=="object"?p.label||p.value||"":p||"";return(f==="undefined"||f==="[object Object]"||f==="null")&&(f=""),f?`<option value="${f}"></option>`:""}).join(""),s=e.value||"";if(e.value){let p=(e.options||[]).find(w=>(typeof w=="object"?w.value:w)==e.value);if(p){let w=typeof p=="object"?p.label||p.value||"":p||"";w&&w!=="undefined"&&w!=="[object Object]"&&w!=="null"&&(s=w)}}r=`
          <input type="text" name="${e.name}" list="${l}" class="form-control" value="${s}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${l}">${n}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let y=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";r=`<input type="date" name="${e.name}" class="form-control" value="${y}" ${i}>`;break;case"month":let v=e.value?String(e.value).slice(0,7):"";r=`<input type="month" name="${e.name}" class="form-control" value="${v}" ${i}>`;break;case"number":let g=e.readonly?"readonly":"";r=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i} ${g}>`;break;case"email":r=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let d=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${r}${d}</div>`}).join("")}function at(t){let e={},i=new FormData(t);for(let[a,r]of i.entries())e[a]=r===""?null:r;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function nt(t,e){e&&Object.entries(e).forEach(([i,a])=>{let r=t.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!a:r.type==="date"&&a&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(a):r.type==="month"&&a?r.value=String(a).slice(0,7):r.value=a??""))})}z();var $e={},it={on(t,e){$e[t]||($e[t]=new Set),$e[t].add(e)},off(t,e){$e[t]&&$e[t].delete(e)},emit(t,e){$e[t]&&$e[t].forEach(i=>{try{i(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys($e).forEach(t=>delete $e[t])}},Ta=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function ce(t){if(!t){it.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");it.emit("data:changed",{module:e,relevant:Ta.has(e)})}function P({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:d,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:s=!0,canEdit:y=!0,canDelete:v=!0,onBeforeSubmit:g,onAfterLoad:p,onDataLoaded:w,extraActions:f=[],initialSearch:b="",exportOptions:c=null,bulkDelete:h=!1,paginationMode:m="server"}){let x=ee();x&&x.role==="viewer"&&(s=!1,y=!1,v=!1,h=!1,c=null),x&&x.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(s=!1,y=!1,v=!1,h=!1,c=null),x&&x.role==="input_lembur"&&(a==="/api/overtime"?(v=!1,h=!1,c=null):(s=!1,y=!1,v=!1,h=!1,c=null));let u=1,S={...l};b&&(S.search=b);let _=new Set;t.innerHTML=`
    ${h?`
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
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${s?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${c?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${c.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${c.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${c.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${c.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${o&&o.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${o.filter(k=>k.type==="search").map(k=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${k.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let C=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select" name="${k.name}" id="filter-${k.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${C}</option>${(k.options||[]).map(E=>`<option value="${typeof E=="object"?E.value:E}" ${S[k.name]===(typeof E=="object"?E.value:E)?"selected":""}>${typeof E=="object"?E.label:E}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let C=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select filter-select-sheet" name="${k.name}-sheet" id="filter-sheet-${k.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${C}</option>${(k.options||[]).map(E=>`<option value="${typeof E=="object"?E.value:E}" ${S[k.name]===(typeof E=="object"?E.value:E)?"selected":""}>${typeof E=="object"?E.label:E}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let k=document.getElementById("bulk-toolbar");if(!k)return;let C=document.getElementById("bulk-count"),E=document.getElementById("btn-bulk-delete"),K=document.getElementById("btn-bulk-cancel");C.textContent=`${_.size} item dipilih`,_.size>0?(k.style.display="flex",E.disabled=!1,K.disabled=!1):(k.style.display="none",E.disabled=!0,K.disabled=!0);let R=document.getElementById("select-all-checkbox");if(R){let Q=document.querySelectorAll(".row-checkbox");if(Q.length>0){let F=[...Q].every(se=>se.checked),A=[...Q].some(se=>se.checked);R.checked=F,R.indeterminate=A&&!F}else R.checked=!1,R.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(C=>C.checked=!1);let k=document.getElementById("select-all-checkbox");k&&(k.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let k=[..._],C=document.createElement("div");C.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",C.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${k.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${k.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(C),C.querySelector("#bulk-cancel-btn").addEventListener("click",()=>C.remove()),C.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let E=C.querySelector("#bulk-confirm-btn");E.disabled=!0,E.textContent="Menghapus...";let K=await T(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:k})});C.remove(),K.ok?(j(`${k.length} ${n} berhasil dihapus.`),_.clear(),$(),ce(a),N()):O(K.data?.error||"Gagal menghapus data.")})});let D=document.getElementById("filter-search"),I;D?.addEventListener("input",k=>{clearTimeout(I),I=setTimeout(()=>{S.search=k.target.value,u=1,_.clear(),$(),N()},400)}),o?.forEach(k=>{(k.type==="select"||k.type==="combobox")&&(document.getElementById(`filter-${k.name}`)?.addEventListener("change",C=>{S[k.name]=C.target.value;let E=document.getElementById(`filter-sheet-${k.name}`);E&&(E.value=C.target.value),u=1,_.clear(),$(),N()}),document.getElementById(`filter-sheet-${k.name}`)?.addEventListener("change",C=>{S[k.name]=C.target.value;let E=document.getElementById(`filter-${k.name}`);E&&(E.value=C.target.value),u=1,_.clear(),$(),N(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={},D&&(D.value=""),o?.forEach(k=>{let C=document.getElementById(`filter-${k.name}`);C&&(C.value="");let E=document.getElementById(`filter-sheet-${k.name}`);E&&(E.value="")}),u=1,_.clear(),$(),N()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{S={},D&&(D.value=""),o?.forEach(k=>{let C=document.getElementById(`filter-${k.name}`);C&&(C.value="");let E=document.getElementById(`filter-sheet-${k.name}`);E&&(E.value="")}),u=1,_.clear(),$(),N(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ke(null)),c&&document.addEventListener("click",function(k){let C=document.getElementById("aksi-menu-main"),E=document.getElementById("btn-aksi-main");C&&E&&!E.contains(k.target)&&!C.contains(k.target)&&C.classList.remove("show-aksi-menu")});let L=document.getElementById("btn-mobile-filter"),V=document.getElementById("filter-options-wrapper"),W=document.getElementById("btn-close-filter-sheet");if(L&&V&&(L.addEventListener("click",k=>{k.preventDefault(),V.classList.add("sheet-open")}),W&&W.addEventListener("click",k=>{k.preventDefault(),V.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async C=>{let E=C.target,K=E.innerHTML;E.innerHTML="\u23F3 Loading...",E.disabled=!0;try{await c.onExport()}catch{O("Gagal export data")}finally{E.innerHTML=K,E.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let k=document.getElementById(`input-import-${c.moduleName}`);k?.addEventListener("change",async C=>{let E=C.target.files[0];if(!E)return;let K=document.getElementById(`label-import-${c.moduleName}`),R=K?K.querySelector(".import-text"):null,Q=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),K&&(K.style.pointerEvents="none"),k.disabled=!0;try{let F=await Ge(E);if(F.length===0)throw new Error("File kosong atau format salah");await c.onImport(F),j("Import berhasil!"),ce(a),N()}catch(F){O(F.message||"Gagal import data")}finally{R&&(R.innerText=Q),K&&(K.style.pointerEvents="auto"),k.disabled=!1,k.value=""}})}async function N(){$();let k=document.getElementById("table-container");if(!k)return;k.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let C=m==="client",E=C?1:u,K=C?De:20,R=new URLSearchParams({page:E,limit:K,...Object.fromEntries(Object.entries(S).filter(([,U])=>U))}),Q=await T(`${a}?${R}`);if(!Q.ok){k.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Q.data?.error||"Error"}</p></div>`;return}let F=Q.data?.data||Q.data||[],A=Q.data?.pagination,se=F.length;if(C){F=w(F);let U=F.length,X=20,ie=Math.ceil(U/X);u>ie&&ie>0&&(u=ie);let q=(u-1)*X,de=u*X;F=F.slice(q,de),A={page:u,limit:X,total:U,pages:ie}}!1,p&&p(F);let Se=et({columns:r,data:F,onEdit:y?U=>ke(U):null,actions:f.map(U=>({...U,handler:X=>U.handler(X,N)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:h?{selectedIds:_,onToggle:$}:null});k.innerHTML="",k.appendChild(Se);let me=document.getElementById("pagination-container");if(me&&(me.innerHTML="",A&&A.pages>1)){let U=tt({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:X=>{u=X,N()}});U&&me.appendChild(U)}}function xe(k){let C=typeof d=="function"?d(k):d;return Ue(C)}function ke(k){let C=!!k,E=document.createElement("form");if(E.noValidate=!0,E.innerHTML=xe(k),C){let R=typeof d=="function"?d(k):d;nt(E,k)}let{close:K}=oe({title:C?`Edit ${n}`:`Tambah ${n}`,content:E,size:"lg",confirmText:C?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(R,Q)=>{if(!E.reportValidity())return;let F=R.querySelector(".modal-confirm");F.disabled=!0,F.textContent="Menyimpan...";let A=at(E),se=typeof d=="function"?d(k):d,Se=async ie=>{for(let q of ie)if(q.type==="row")await Se(q.fields);else if(q.type==="combobox"&&A[q.name]){let de=A[q.name],_e=(q.options||[]).find(Z=>{let re=String(typeof Z=="object"?Z.value:Z),gt=String(typeof Z=="object"?Z.label:Z);return re===de||gt===de});if(_e)A[q.name]=typeof _e=="object"?_e.value:_e;else if(q.createApi){let Z={};Z[q.createApi.field]=de,q.createApi.extra&&Object.assign(Z,q.createApi.extra);let re=await T(q.createApi.path,{method:"POST",body:JSON.stringify(Z)});if(re.ok&&re.data?.id)A[q.name]=re.data.id;else if(re.ok&&!re.data?.id)A[q.name]=de;else throw new Error(`Gagal membuat master data: ${re.data?.error||"Unknown error"}`)}}};try{await Se(se)}catch(ie){O(ie.message),F.disabled=!1,F.textContent=C?"Simpan Perubahan":`Tambah ${n}`;return}g&&(A=await g(A,k));let me=C?"PUT":"POST",U=C?`${a}/${k.id}`:a,X=await T(U,{method:me,body:JSON.stringify(A)});X.ok?(j(C?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),Q(),ce(a),N()):(O(X.data?.error||"Gagal menyimpan data."),F.disabled=!1,F.textContent=C?"Simpan Perubahan":`Tambah ${n}`)}})}function ne(k){ue(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let C=await T(`${a}/${k.id}`,{method:"DELETE"});C.ok?(j(`${n} berhasil dihapus.`),ce(a),N()):O(C.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return N(),N}M();M();var Fe=null,rt=null;async function Te(t=!1){if(Fe&&!t)return console.log("Employees Raw (Cache Hit)",Fe.slice(0,5)),Fe;let e=await T(`/api/employees?limit=${De}&status=Aktif`);return Fe=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Fe.slice(0,5)),Fe}async function ae(t=!1){let i=(await Te(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function H(t=!1){return rt&&!t||(rt=((await T("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),rt}function Y(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function xt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Be(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function kt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function be(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}z();M();z();M();z();M();z();function ot(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),d=new Date(i.value),o=Math.floor((d-r)/864e5);a.value=isNaN(o)?"":o}else a.value=""}}document.body.removeEventListener("input",ot);document.body.addEventListener("input",ot);document.body.removeEventListener("change",ot);document.body.addEventListener("change",ot);M();var Ee={};function Pe(t){if(Ee[t]){try{Ee[t].destroy()}catch{}delete Ee[t]}}function Da(){Object.keys(Ee).forEach(Pe)}var pe=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Le=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function st(t,e,i=900){if(!t)return;t._animFrame&&cancelAnimationFrame(t._animFrame);let a=Math.max(0,Math.round(pe(e))),r=t.textContent?t.textContent.replace(/[^\d]/g,""):"",d=r!==""&&parseInt(r)||0;if(a===d){t.textContent=a.toLocaleString("id-ID"),t._animFrame=null;return}let o=Date.now(),l=()=>{let n=Math.min((Date.now()-o)/i,1),s=1-Math.pow(1-n,3);t.textContent=Math.round(d+s*(a-d)).toLocaleString("id-ID"),n<1?t._animFrame=requestAnimationFrame(l):(t.textContent=a.toLocaleString("id-ID"),t._animFrame=null)};t._animFrame=requestAnimationFrame(l)}var Ia={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Fa=t=>{let e=Le(t,"\u2014");return`<span class="status-pill ${Ia[e]||"pill-neutral"}">${e}</span>`};var fe={family:"Inter",size:11},we="#94A3B8",Ae="#F1F5F9",wt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Ba=()=>window.innerWidth<768;function lt(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Ba()?"bottom":"top",labels:{font:fe,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:fe,titleFont:{...fe,weight:"700"}}},scales:{x:{grid:{color:Ae},ticks:{font:fe,color:we,maxRotation:0}},y:{grid:{color:Ae},ticks:{font:fe,color:we},beginAtZero:!0}},...t}}var La=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),Aa=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Pt(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function he(t,e,i=8e3){try{let a=new AbortController,r=setTimeout(()=>a.abort(),i),d=await T(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(r),!d||!d.ok)return e;let o=d.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function Pa(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let r=document.getElementById(a);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let r=document.getElementById(a);if(r&&r.style.display==="none"){r.style.display="block";let d=r.parentElement;if(d&&!d.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",r.style.display="none",d.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Ot({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Rt({}),["table-contracts","table-issues"].forEach(a=>{let r=document.getElementById(a);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Mt(t){Da(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${La()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Aa()}</div>

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
            <div class="chart-card-title">Jadwal Hari Ini <span style="font-size:0.75rem; font-weight:normal; color:var(--text-3); margin-left:6px">${e}</span></div>
            <a href="#/calendar" class="chart-link">Lihat Kalender</a>
          </div>
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Pt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Pt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Nt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=a);let d=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");d&&(d.style.display="block",d.style.position="absolute"),o&&(o.style.display="none");let l=await he(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Kt(l)}catch(n){console.warn("ScheduleChart render:",n),ye("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,r=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",d=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");d&&(d.style.display="block",d.style.position="absolute"),o&&(o.style.display="none");let l=await he(r,{},8e3);try{zt(l)}catch(n){console.warn("InspBar render:",n),ye("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>Pa(),5e3),await Nt(t)}async function Nt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,r,d,o,l,n]=await Promise.all([he("/api/dashboard/kpi",{},8e3),he("/api/dashboard/issues-trend",{},8e3),he("/api/dashboard/issues-summary",{},8e3),he("/api/dashboard/stats",{},8e3),he("/api/dashboard/calendar",[],8e3),he("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),he(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),he("/api/dashboard/dropdown-aggregates",{},8e3)]),s=document.getElementById("filter-insp-month"),y=s?s.value:"",v=y?`/api/dashboard/inspection-bar?month=${y}`:"/api/dashboard/inspection-bar",g=await he(v,{},8e3);window.dashboardAggregates=n||{};try{Ot(e)}catch(p){console.warn("KPI render:",p)}try{Rt(e)}catch(p){console.warn("MiniStats render:",p)}try{Kt(l)}catch(p){console.warn("ScheduleChart render:",p),ye("skel-jadwal","chart-jadwal")}try{Na(Array.isArray(a?.by_category)?a.by_category:[])}catch(p){console.warn("Donut render:",p),ye("skel-donut","chart-donut")}try{Ma(i)}catch(p){console.warn("Trend render:",p),ye("skel-trend","chart-trend")}try{zt(g)}catch(p){console.warn("InspBar render:",p),ye("skel-insp","chart-insp")}try{let p=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Ra(p)}catch(p){console.warn("IssuesTable render:",p)}try{let p=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];Oa(o)}catch(p){console.warn("ContractsTable render:",p)}try{Ka(Array.isArray(d)?d:[])}catch(p){console.warn("Agenda render:",p)}try{za()}catch(p){console.warn("Quick Actions render:",p)}}function Ot(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=window.dashboardAggregates||{},r=n=>{let s=window.dashboardAggregates?.schedule_quarter_details?.[n];return s?`Insp: ${s.insp} \u2022 GC: ${s.gc} \u2022 DC: ${s.dc} \u2022 Fog: ${s.fog}`:""},d=n=>{let s=window.dashboardAggregates?.schedule_quarter_details?.[n];return s?`Rincian ${n}: Inspeksi: ${s.insp} | General Cleaning: ${s.gc} | Deep Cleaning: ${s.dc} | Fogging: ${s.fog}`:`Total Kegiatan ${n}`},o=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:a.schedule_by_quarter?.[i]??t.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:r(i),subId:"mini-jadwal-detail",tooltip:d(i)}];e.innerHTML=o.map(n=>{let s=n.val!==void 0?pe(n.val,0):pe(t[n.key]?.current,0);return n.dropdown?`
        <div class="kpi-card ${n.color}" ${n.id?`id="${n.id}"`:""} ${n.tooltip?`title="${n.tooltip}"`:""} data-href="${n.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${n.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${s}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${s}</div>
                ${n.dropdown}
              </div>
              <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${n.label}</div>
              <div class="kpi-subtitle" ${n.subId?`id="${n.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${n.tooltip||""}">${n.sub}</div>
            </div>
          </div>
        </div>`:`
      <a href="${n.href}" class="kpi-card ${n.color}" ${n.id?`id="${n.id}"`:""} ${n.tooltip?`title="${n.tooltip}"`:""} style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${n.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="kpi-value" data-target="${s}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${s}</div>
              ${n.dropdown?n.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${n.label}</div>
            <div class="kpi-subtitle" ${n.subId?`id="${n.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${n.tooltip||""}">${n.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-card[data-href]").forEach(n=>{n.addEventListener("click",s=>{s.target.closest("select")||n.dataset.href&&(window.location.hash=n.dataset.href)}),n.addEventListener("keydown",s=>{s.key==="Enter"&&!s.target.closest("select")&&n.dataset.href&&(window.location.hash=n.dataset.href)})}),e.querySelectorAll(".kpi-value").forEach(n=>{st(n,parseInt(n.dataset.target)||0)});let l=document.getElementById("dash-jadwal-period");if(l){let n=s=>{let v=(window.dashboardAggregates?.schedule_by_quarter||{})[s]||0,g=document.querySelector("#mini-jadwal .kpi-value");g&&(g.dataset.target=v,st(g,v,400));let p=document.getElementById("mini-jadwal-detail");p&&(p.textContent=r(s),p.title=d(s));let w=document.getElementById("mini-jadwal");if(w){let f=`#/timeline?dash_filter=period_${s.toLowerCase()}`;w.tagName==="A"?w.href=f:w.dataset.href=f,w.title=d(s)}};l.addEventListener("change",s=>n(s.target.value))}}function Rt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=new Date().getFullYear(),a=String(new Date().getMonth()+1).padStart(2,"0"),r=`${i}-${a}`,d=s=>`
    <select id="${s}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
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
  `,o=window.dashboardAggregates||{},l=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:d("dash-inspeksi-month"),val:o.inspeksi_by_month?.[r]??t.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${r}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:d("dash-gcdc-month"),val:o.gcdc_by_month?.[r]??t.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${r}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:d("dash-reliefer-month"),val:o.relievers_by_month?.[r]??t.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${r}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:d("dash-fogging-month"),val:o.fogging_by_month?.[r]??t.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${r}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=l.map(s=>s.dropdown?`
        <div class="mini-stat ${s.color}" style="cursor:pointer;" id="${s.id||""}" data-href="${s.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${s.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${pe(s.val)}">0</div>
              ${s.dropdown}
            </div>
            <div class="mini-stat-text">${s.label}</div>
            ${s.extra||""}
          </div>
        </div>`:`
      <a href="${s.href}" class="mini-stat ${s.color}" style="text-decoration:none" id="${s.id||""}">
        <div class="mini-stat-icon">${s.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${pe(s.val)}">0</div>
            ${s.dropdown?s.dropdown:""}
          </div>
          <div class="mini-stat-text">${s.label}</div>
          ${s.extra||""}
        </div>
      </a>`).join(""),e.querySelectorAll(".mini-stat[data-href]").forEach(s=>{s.addEventListener("click",y=>{y.target.closest("select")||s.dataset.href&&(window.location.hash=s.dataset.href)}),s.addEventListener("keydown",y=>{y.key==="Enter"&&!y.target.closest("select")&&s.dataset.href&&(window.location.hash=s.dataset.href)})}),e.querySelectorAll(".mini-stat-value").forEach(s=>st(s,parseInt(s.dataset.target)||0,700));let n=(s,y,v,g)=>{let p=document.getElementById(s);if(p){let w=f=>{let c=(window.dashboardAggregates?.[v]||{})[f]||0,h=document.querySelector(`#${y} .mini-stat-value`);h&&(h.dataset.target=c,st(h,c,400));let m=document.getElementById(y);if(m){let x=`${g}&month=${f}`;m.tagName==="A"?m.href=x:m.dataset.href=x}};p.addEventListener("change",f=>w(f.target.value))}};n("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),n("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),n("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),n("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Na(t){ye("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Pe("donut");let a=(t||[]).filter(n=>pe(n.count)>0);if(!a.length){Ve(e,"Belum ada data permasalahan");return}let r=a.map(n=>`${Le(n.category,"Lainnya")}`),d=a.map(n=>pe(n.count)),o=d.reduce((n,s)=>n+s,0);i.innerHTML=a.map((n,s)=>{let y=wt[s%wt.length],v=o>0?Math.round(n.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${y}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${v}%)</span></div>
          <div class="donut-legend-label">${r[s]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(n){let s=n.width,y=n.height,v=n.ctx;v.restore();let g=(y/80).toFixed(2);v.font="bold "+g+"em Inter",v.textBaseline="middle",v.fillStyle="#1E293B";let p=o.toString(),w=Math.round((s-v.measureText(p).width)/2),f=y/2;v.fillText(p,w,f-4),v.font="600 "+(g*.35).toFixed(2)+"em Inter",v.fillStyle="#64748B";let b="Total",c=Math.round((s-v.measureText(b).width)/2);v.fillText(b,c,f+10),v.save()}};Ee.donut=new Chart(e,{type:"doughnut",data:{labels:r,datasets:[{data:d,backgroundColor:wt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:fe,titleFont:{...fe,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function Ma(t){ye("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Pe("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[l,n]=o.split("-");return(i[Number(n)-1]||n)+" "+String(l).slice(-2)}catch{return o}}),r=(t.open||[]).map(o=>pe(o)),d=(t.closed||[]).map(o=>pe(o));if(!a.length){Ve(e,"Belum ada data trend");return}Ee.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:d,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:lt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:we,maxRotation:45,autoSkip:!0}},y:{grid:{color:Ae},ticks:{font:{family:"Inter",size:9},color:we},beginAtZero:!0}}})})}function Kt(t){ye("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Pe("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(s=>s>0))){Ve(e,"Belum ada data jadwal");return}let r=t["Inspeksi Hygiene"]||Array(12).fill(0),d=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),l=t.Fogging||Array(12).fill(0);Ee.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:d,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:lt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:we,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Ae},ticks:{font:{family:"Inter",size:9},color:we},min:0}}})})}function zt(t){ye("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Pe("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(d=>pe(d)),r=(t.spv||[]).map(d=>pe(d));if(!i.length){Ve(e,"Belum ada data inspeksi");return}Ee.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:lt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:fe,color:we,maxRotation:45,minRotation:45}},y:{grid:{color:Ae},ticks:{font:fe,color:we},min:0,max:100}}})})}function Oa(t){ye("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Pe("contractMiniBar"),t=t||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let l=o.split("-")[1];return i[l]||o}),r=(t.data||[]).map(o=>pe(o));if(!a.length){Ve(e,"Belum ada data");return}let d=e.getContext("2d");Ee.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:lt({onClick:(o,l)=>{if(l&&l.length>0){let n=l[0].index,s=(t.labels||[])[n];s&&(window.location.hash="#/contracts?month_expiry="+s)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:fe,color:we,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ae,borderDash:[4,4],drawBorder:!1},ticks:{font:fe,color:we,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Ra(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${Fa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Le(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Le(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Ka(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,d=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!d.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${d.map(o=>{let l="#3B82F6",n="#EFF6FF",s="Agenda",y=(o.title||"").toLowerCase();return y.includes("inspeksi")?(l="#10B981",n="#ECFDF5",s="Inspeksi"):y.includes("cleaning")||y.includes("gcdc")?(l="#3B82F6",n="#EFF6FF",s="Cleaning"):y.includes("reliefer")?(l="#F59E0B",n="#FFFBEB",s="Reliefer"):y.includes("fogging")&&(l="#8B5CF6",n="#F5F3FF",s="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Le(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Le(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${l}">${s}</div>
        </div>
      `}).join("")}
    </div>
  `}function za(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ye(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let r=a.parentElement;if(r){let d=r.querySelector(".chart-empty");d&&d.remove()}}}function Ve(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=e,i.appendChild(r)}}M();async function jt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),d=document.getElementById("login-password");r?.addEventListener("click",()=>{let o=d.type==="text";d.type=o?"password":"text",r.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),i.style.display="none";let l=e.username.value.trim(),n=e.password.value;if(!l||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let s=await T("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:n})});s.ok&&s.data.success?(ft(s.data.data.token),He(s.data.data.user),j("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=s.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}M();z();async function ja(){return await H()}function Ne(t){if(!t||t==="-"||String(t).trim()==="")return"-";let e=String(t).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],a=e.match(/^(\d{4})-(\d{1,2})/);if(a){let d=a[1],o=parseInt(a[2],10)-1;if(o>=0&&o<12)return`${i[o]} ${d}`}let r=e.split(/[\/\-\.]/);if(r.length===2&&r[1].length===4){let d=parseInt(r[0],10)-1;if(d>=0&&d<12)return`${i[d]} ${r[1]}`}if(r.length===3&&r[2].length===4){let d=parseInt(r[1],10)-1;if(d>=0&&d<12)return`${i[d]} ${r[2]}`}for(let d=0;d<12;d++)if(e.toLowerCase().includes(i[d].toLowerCase()))return e;return e}function qa(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function qt(t,e){let i=await ja(),a=e?e.get("dash_filter"):null;P({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>a?r.filter(d=>qa(d,a)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Be(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"target_pindah_os",label:"Target Pindah OS",render:r=>Ne(r)},{key:"target_selesai_os",label:"Target Selesai OS",render:r=>Ne(r)},{key:"status",label:"Status",render:r=>Y(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await T(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let d=r.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||"","Target Pindah OS":Ne(o.target_pindah_os)==="-"?"":Ne(o.target_pindah_os),"Target Selesai OS":Ne(o.target_selesai_os)==="-"?"":Ne(o.target_selesai_os)}));B(d,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async r=>{let d=n=>{if(!n)return null;let s=String(n||"").toLowerCase(),y=i.find(v=>String(v.label||"").toLowerCase()===s);return y?y.value:null},o=r.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:d(String(n.Cabang||"").trim()),division:String(n.Divisi||n["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||n["No. Hp"]||"").trim(),join_date:String(n["Tgl Masuk"]||n["Tanggal Masuk"]||"").trim(),status:String(n.Status||"").trim(),target_pindah_os:String(n["Target Pindah OS"]||"").trim(),target_selesai_os:String(n["Target Selesai OS"]||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),l=await T("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}M();z();var _t=[],Ht=[];async function Ha(){_t=await H(),Ht=await Te()}var St=async t=>{let e=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(M(),Ie))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let d=r.data?.data||r.data||[],o=Array.isArray(d)?d:[];if(e=e.concat(o),o.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return e};async function dt(t,e){await Ha(),P({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Be(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':xt(a)},{key:"status",label:"Status",render:a=>Y(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[d,o]=await Promise.all([St("/api/employees?status=Aktif"),St("/api/contracts")]);if(d.length>0){let l=o.filter(v=>v.status==="Aktif"),n=new Set(l.map(v=>v.employee_id)),s=d.filter(v=>!n.has(v.id)),y=`<p style="margin-bottom:12px">Data yang terbaca: <b>${d.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${s.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;s.forEach(v=>{let g=o.filter(w=>w.employee_id===v.id),p='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(g.length>0){let w=g[0];p=`Pernah ada kontrak (Status: <b style="color:#EF4444">${w.status}</b>, Selesai: ${window.formatDate(w.end_date)})`}y+=`<li style="margin-bottom:8px"><b>${v.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${v.branch_name||"-"} | ${p}</span></li>`}),y+="</ul>",Promise.resolve().then(()=>(ge(),vt)).then(v=>v.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:y,cancelText:"Tutup"}))}}catch(d){console.error(d)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Ht,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:_t,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await T(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let r=a.data.data.map(d=>({"Nama Lengkap":d.employee_name,Cabang:d.branch_name||"","Div / Bagian":d.division||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":d.end_date&&String(d.end_date).startsWith("2099")?"":d.end_date||"","Sisa Kontrak":d.end_date&&String(d.end_date).startsWith("2099")?"Tetap":d.days_remaining!==null&&d.days_remaining!==void 0?`${d.days_remaining} Hari`:"",Status:d.status||""}));B(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[r,d]=await Promise.all([T("/api/branches?limit=10000"),St("/api/employees")]),o=r.data?.data||[],l=d||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((b,c)=>{console.log(`${c+1}. ID: ${b.id}, Name: ${b.full_name}, Status: ${b.status}`)}));let n=b=>{if(!b)return null;let c=String(b||"").replace(/\s+/g," ").toLowerCase().trim(),h=o.find(m=>String(m.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===c||String(m.code||"").replace(/\s+/g," ").toLowerCase().trim()===c||String(m.name||"").replace(/\s+/g," ").toLowerCase().trim()===c);return h?h.id:null},s=(b,c)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${c}`),console.log(`Nama dari Excel : "${b}"`),!b)return console.log("Alasan gagal mapping : Nama kosong"),null;let h=String(b||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${h}"`),console.log(`Jumlah employee di database : ${l.length}`);let m=l.find(x=>String(x.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return m?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${m.id}`),m.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},y=b=>{if(!b)return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let c=String(b).trim();if(/^\d{4,5}(\.\d+)?$/.test(c)){let m=Math.floor(Number(c));if(m>2e4&&m<99999){let x=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(x.getTime())?"":x.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);let h=c.split(/[\/\-\.]/);if(h.length===3){let[m,x,u]=h.map(S=>S.trim());if(m.length===4&&x.length<=2&&u.length<=2)return`${m}-${x.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&x.length<=2&&m.length<=2)return`${u}-${x.padStart(2,"0")}-${m.padStart(2,"0")}`}return c},v=a.map((b,c)=>{let h=c+2,m=String(b["Nama Lengkap"]||"").trim(),x=b["Tanggal Mulai"],u=y(x);if(!u){let $=a.__worksheet,D=a.__headers||[],I=D.indexOf("Tanggal Mulai"),L="N/A",V="N/A",W="N/A";if(I!==-1&&$&&window.XLSX){let xe=window.XLSX.utils.encode_cell({c:I,r:h-1});W=xe;let ke=$[xe];ke?(L=ke.t||"undefined",V=ke.w||"undefined"):L="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let N="Unknown";x==null||x===""?N="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":x instanceof Date&&isNaN(x.getTime())?N="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":N="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${h}`),console.log(`Employee Name : ${m}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${I})`),console.log(`Raw Cell Value : "${x}"`),console.log(`JavaScript Type : ${typeof x}`),console.log(`SheetJS Cell Type : ${L}`),console.log(`SheetJS Formatted Value : "${V}"`),console.log(`Value After Trim : "${String(x||"").trim()}"`),console.log(`Value After Date Parser : "${u}"`),console.log(`Is Empty : ${!x}`),console.log(`Is Invalid Date : ${x instanceof Date?isNaN(x.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${N}`),console.log(`Workbook Sheet : ${$?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${W}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(b,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(D)),console.log(`==========================
`)}let S=s(m,h),_=null;return S?u||(_="Tanggal Mulai kosong atau tidak berformat tanggal"):_="Karyawan tidak ditemukan di Database",{isValid:!!(S&&u),invalidReason:_,rowNum:h,data:{employee_id:S,branch_id:n(String(b.Cabang||"").trim()),division:String(b["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:u,end_date:y(b["Tanggal Selesai"])||"2099-12-31",status:String(b.Status||"").trim(),_rawName:m}}}),g=[],p=[];if(v.forEach(b=>{b.isValid?g.push(b.data):p.push({rowNum:b.rowNum,name:b.data._rawName,reason:b.invalidReason})}),console.log(`Split Validation - Valid: ${g.length}, Invalid: ${p.length}`),g.length===0){let b=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${p.length}

Daftar Kegagalan (Contoh):
`;p.slice(0,10).forEach(c=>{b+=`- Row ${c.rowNum} | Nama: ${c.name} | Alasan: ${c.reason}
`}),p.length>10&&(b+=`- ... dan ${p.length-10} lainnya.
`),alert(b);return}let w=await T("/api/contracts/import",{method:"POST",body:JSON.stringify(g)}),f=`IMPORT SUMMARY
======================
`;f+=`Total Baris Excel : ${a.length}
`,f+=`Baris Valid       : ${g.length}
`,f+=`Baris Invalid     : ${p.length}

`,w&&w.data&&w.data.metrics?(f+=`Berhasil INSERT   : ${w.data.metrics.inserted}
`,f+=`Berhasil UPDATE   : ${w.data.metrics.updated}
`):f+=`Berhasil diproses : ${g.length}
`,p.length>0&&(f+=`
DAFTAR DATA DILEWATI:
`,p.forEach(b=>{f+=`- Row ${b.rowNum} | ${b.name} | ${b.reason}
`})),alert(f),typeof dt=="function"&&dt()}}})}M();z();var $t=[],Qe=[];function Ja(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Jt(t,e){$t=await H();let i=await ae();Qe=["Berlin Ariansyah","Ade Surahman"];let a=b=>b&&!Qe.find(c=>String(typeof c=="object"?c.value:c).toLowerCase()===String(b).toLowerCase())?[...Qe,b]:Qe,r=await T(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),d=b=>{if(!b||b==="-"||String(b).trim()==="")return"";let c=String(b).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:b},o=r.data?.data||[],l=Ja(o),n=e?e.get("dash_filter"):null,s=new Date,y=`${s.getFullYear()}-${String(s.getMonth()+1).padStart(2,"0")}`,v={},g=e&&e.get("month")?e.get("month"):null;n==="inspeksi"?v={status:"Done",activity_type:"Inspeksi Hygiene",month:g}:n==="gcdc"?v={status:"Done",activity_type:"GCDC",month:g}:n&&n.startsWith("period_")&&(v={period:n.replace("period_","").toUpperCase()});let p=new Date().getFullYear(),f=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((b,c)=>{let h=String(c+1).padStart(2,"0");return{value:`${p}-${h}`,label:`${b} ${p}`}});P({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:v,onDataLoaded:b=>b.sort((c,h)=>{let m=c.opening_date?new Date(c.opening_date).getTime():0;return(h.opening_date?new Date(h.opening_date).getTime():0)-m}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:b=>kt(b)},{key:"period",label:"Periode",render:b=>be(b)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:b=>d(b)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:b=>d(b)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:b=>d(b)},{key:"status",label:"Status",render:b=>Y(b)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:$t},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:f},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Qe}],formFields:b=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:$t,value:b?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:b?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:b?.period},{name:"pic",label:"PIC",type:"combobox",options:a(b?.pic),value:b?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:b?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:b?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:b?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:b?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:b?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let b=await T(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(b.ok){let c=b.data.data.map(h=>({Cabang:h.branch_name||"",Kegiatan:h.activity_type||"",Periode:h.period||"",PIC:h.pic||"","Tgl Opening":h.opening_date||"","Tgl Target":h.target_date||"","Tgl Selesai":h.completion_date||"",Status:h.status||""}));B(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async b=>{let h=(await T("/api/branches?all=1")).data?.data||[],m=_=>{if(!_)return null;let $=String(_||"").toLowerCase().trim(),D=$.match(/^(\d{3})/);if(D){let L=h.find(V=>String(V.code||"")===D[1]);if(L)return L.id}let I=h.find(L=>String(L.full_name||"").toLowerCase()===$||String(L.code||"").toLowerCase()===$||String(L.name||"").toLowerCase()===$);return I?I.id:null},x=_=>{if(_==null||_==="")return"";if(_ instanceof Date&&!isNaN(_.getTime()))return _.toISOString().slice(0,10);let $=String(_).trim();if($===""||$==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test($))return $.slice(0,10);if(/^\d{4,5}$/.test($)){let I=Number($);if(I>2e4&&I<99999){let L=new Date(Date.UTC(1899,11,30)+I*864e5);return isNaN(L.getTime())?"":L.toISOString().slice(0,10)}}let D=$.split(/[\/\-\.]/);if(D.length===3){let[I,L,V]=D.map(W=>W.trim());if(I.length===4&&L.length<=2&&V.length<=2)return`${I}-${L.padStart(2,"0")}-${V.padStart(2,"0")}`;if(V.length===4&&L.length<=2&&I.length<=2)return`${V}-${L.padStart(2,"0")}-${I.padStart(2,"0")}`}return $},u=b.map(_=>({branch_id:m(String(_.Cabang||"").trim()),branch_name:String(_.Cabang||"").trim(),activity_type:String(_.Kegiatan||"").trim(),period:String(_.Periode||"").trim(),pic:String(_.PIC||_.Pic||"").trim(),opening_date:x(_["Tgl Opening"]||_["Tanggal Opening"]||_["Tgl Openir"]),target_date:x(_["Tgl Target"]||_["Tanggal Target"]),completion_date:x(_["Tgl Selesai"]||_["Tanggal Selesai"]),status:String(_.Status||"").trim(),notes:String(_.Catatan||_.Keterangan||"").trim()})).filter(_=>_.activity_type&&_.period),S=await T("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!S.ok)throw new Error(S.data?.error||"Import gagal");return ce("schedule"),S.data}}})}M();z();var Et=[],ct=[];function pt(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),d=new Date(i.value),o=Math.floor((d-r)/864e5);a.value=isNaN(o)?"":o}else a.value=""}}document.body.removeEventListener("input",pt);document.body.addEventListener("input",pt);document.body.removeEventListener("change",pt);document.body.addEventListener("change",pt);function Ua(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function Ut(t,e){let i=e?e.get("dash_filter"):null;Et=await H(),ct=await ae();let a=n=>n&&!ct.find(s=>s.value===n)?[...ct,{value:n,label:n}]:ct,r=new Date().getFullYear(),d=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,s)=>{let y=String(s+1).padStart(2,"0");return{value:`${r}-${y}`,label:`${n} ${r}`}});P({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>i?n.filter(s=>Ua(s,i)):n,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"...":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"...":n||"-"}</span>`},{key:"status",label:"Status",render:n=>Y(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Day",render:n=>n??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Et},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:d}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Et,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date},{name:"day_count",label:"Day",type:"number",value:n?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await T(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let s=n.data.data.map(y=>({"Tanggal Info":y.report_date||"",Cabang:y.branch_name||"",Kategori:y.category||"","Sumber Laporan":y.source||"",Keluhan:y.complaint||"","Nama FC":y.employee_name||"","FC Spesialis":y.fc_specialist||"",Solusi:y.solution||"",Status:y.status||"","Tanggal Selesai":y.completion_date||"",Day:y.day_count!==null?y.day_count:""}));B(s,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async n=>{let y=(await T("/api/branches?all=1")).data?.data||[],v=w=>{if(!w)return null;let f=String(w||"").toLowerCase(),b=y.find(c=>String(c.full_name||"").toLowerCase()===f||String(c.code||"").toLowerCase()===f||String(c.name||"").toLowerCase()===f);return b?b.id:null},g=n.map(w=>({branch_id:v(String(w.Cabang||"").trim()),report_date:String(w["Tanggal Info"]||w.Tanggal||"").trim(),category:String(w.Kategori||"").trim(),source:String(w["Sumber Laporan"]||w.Sumber||"").trim(),complaint:String(w.Keluhan||"").trim(),employee_name:String(w["Nama FC"]||"").trim(),fc_specialist:String(w["FC Spesialis"]||"").trim(),solution:String(w.Solusi||"").trim(),completion_date:String(w["Tanggal Selesai"]||w["Tgl Selesai"]||"").trim(),status:String(w.Status||"").trim(),day_count:w.Day||w.Hari||null})).filter(w=>w.report_date&&w.complaint&&w.category),p=await T("/api/import/issues",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}M();var Me=[];function Ga(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Gt(t,e){let i=e?e.get("dash_filter"):null;Me=await H();let a=await ae(),r=["Ade","Berlin"],d=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a,o=l=>l&&!r.find(n=>(typeof n=="object"?n.value:n)===l)?[...r,l]:r;P({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>i?l.filter(n=>Ga(n,i)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>Y(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),s=await T(`/api/one-on-one?limit=10000&${n}`);if(s.ok){let y=s.data.data.map(g=>({Tanggal:g.meeting_date||"",Cabang:g.branch_name||"","Nama Karyawan":g.employee_name||"",PIC:g.pic||"",Masalah:g.problem||"",Solusi:g.solution||"",Status:g.status||"","Tgl Selesai":g.completion_date||"",Dokumen:g.document_link||""})),{downloadExcel:v}=await Promise.resolve().then(()=>(z(),le));v(y,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(z(),le));n(l,"Template_Import_OneOnOne")},onImport:async l=>{let n=g=>{if(!g)return null;let p=String(g||"").toLowerCase(),w=Me.find(f=>String(f.label||"").toLowerCase()===p);return w?w.value:null},s=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let p=String(g).trim();if(/^\d{4,5}$/.test(p)){let f=Number(p);if(f>2e4&&f<99999){let b=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);let w=p.split(/[\/\-\.]/);if(w.length===3){let[f,b,c]=w.map(h=>h.trim());if(f.length===4&&b.length<=2&&c.length<=2)return`${f}-${b.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&b.length<=2&&f.length<=2)return`${c}-${b.padStart(2,"0")}-${f.padStart(2,"0")}`}return p},y=l.map(g=>({meeting_date:s(g.Tanggal),employee_name:String(g["Nama Karyawan"]||"").trim(),branch_id:n(String(g.Cabang||"").trim()),pic:String(g.PIC||"").trim(),problem:String(g.Masalah||"").trim(),solution:String(g.Solusi||"").trim(),status:String(g.Status||"").trim(),completion_date:s(g["Tgl Selesai"]),document_link:String(g.Dokumen||"").trim()})).filter(g=>g.meeting_date&&g.employee_name&&g.branch_id),v=await T("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:y,onDuplicate:"update"})});if(!v.ok)throw new Error(v.data?.error||"Import gagal");return v.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Me.find(n=>n.value==l.branch_id)?[...Me,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:d(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}M();async function Vt(t){let e=await H(),i=await ae(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=n=>n&&!i.find(s=>s.value===n)?[...i,{value:n,label:n}]:i,d=n=>n&&!a.find(s=>(typeof s=="object"?s.value:s)===n)?[...a,n]:a,o=Array.from({length:5},(n,s)=>String(new Date().getFullYear()-s));P({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let s=JSON.parse(n);return Array.isArray(s)?s.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let s=new URLSearchParams(n||{}).toString(),y=await T(`/api/training?limit=10000&${s}`);if(y.ok){let v=y.data.data.map(p=>{let w=p.participants||"";try{let f=JSON.parse(w);w=Array.isArray(f)?f.join(", "):w}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:w,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:g}=await Promise.resolve().then(()=>(z(),le));g(v,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(z(),le));s(n,"Template_Import_Training")},onImport:async n=>{let s=p=>{if(!p)return null;let w=String(p||"").toLowerCase(),f=e.find(b=>String(b.label||"").toLowerCase()===w);return f?f.value:null},y=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let w=String(p).trim();if(/^\d{4,5}$/.test(w)){let b=Number(w);if(b>2e4&&b<99999){let c=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(w))return w.slice(0,10);let f=w.split(/[\/\-\.]/);if(f.length===3){let[b,c,h]=f.map(m=>m.trim());if(b.length===4&&c.length<=2&&h.length<=2)return`${b}-${c.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&c.length<=2&&b.length<=2)return`${h}-${c.padStart(2,"0")}-${b.padStart(2,"0")}`}return w},v=n.map(p=>({training_date:y(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:s(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),g=await T("/api/training/import",{method:"POST",body:JSON.stringify(v)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:d(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let s=JSON.parse(n?.participants);return Array.isArray(s)?s.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>n})}M();ge();z();function Qt({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:d,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:s=!0,canEdit:y=!0,canDelete:v=!0,onBeforeSubmit:g,onAfterLoad:p,onDataLoaded:w,extraActions:f=[],initialSearch:b="",exportOptions:c=null,bulkDelete:h=!1,paginationMode:m="server"}){let x=ee();x&&x.role==="viewer"&&(s=!1,y=!1,v=!1,h=!1,c=null),x&&x.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(s=!1,y=!1,v=!1,h=!1,c=null),x&&x.role==="input_lembur"&&(a==="/api/overtime"?(v=!1,h=!1,c=null):(s=!1,y=!1,v=!1,h=!1,c=null));let u=1,S={...l};b&&(S.search=b);let _=new Set;t.innerHTML=`
    ${h?`
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
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${s?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${c?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${c.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${c.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${c.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${c.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${o&&o.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${o.filter(k=>k.type==="search").map(k=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${k.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let C=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select" name="${k.name}" id="filter-${k.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${C}</option>${(k.options||[]).map(E=>`<option value="${typeof E=="object"?E.value:E}" ${S[k.name]===(typeof E=="object"?E.value:E)?"selected":""}>${typeof E=="object"?E.label:E}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let C=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select filter-select-sheet" name="${k.name}-sheet" id="filter-sheet-${k.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${C}</option>${(k.options||[]).map(E=>`<option value="${typeof E=="object"?E.value:E}" ${S[k.name]===(typeof E=="object"?E.value:E)?"selected":""}>${typeof E=="object"?E.label:E}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let k=document.getElementById("bulk-toolbar");if(!k)return;let C=document.getElementById("bulk-count"),E=document.getElementById("btn-bulk-delete"),K=document.getElementById("btn-bulk-cancel");C.textContent=`${_.size} item dipilih`,_.size>0?(k.style.display="flex",E.disabled=!1,K.disabled=!1):(k.style.display="none",E.disabled=!0,K.disabled=!0);let R=document.getElementById("select-all-checkbox");if(R){let Q=document.querySelectorAll(".row-checkbox");if(Q.length>0){let F=[...Q].every(se=>se.checked),A=[...Q].some(se=>se.checked);R.checked=F,R.indeterminate=A&&!F}else R.checked=!1,R.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(C=>C.checked=!1);let k=document.getElementById("select-all-checkbox");k&&(k.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let k=[..._],C=document.createElement("div");C.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",C.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${k.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${k.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(C),C.querySelector("#bulk-cancel-btn").addEventListener("click",()=>C.remove()),C.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let E=C.querySelector("#bulk-confirm-btn");E.disabled=!0,E.textContent="Menghapus...";let K=await T(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:k})});C.remove(),K.ok?(j(`${k.length} ${n} berhasil dihapus.`),_.clear(),$(),ce(a),N()):O(K.data?.error||"Gagal menghapus data.")})});let D=document.getElementById("filter-search"),I;D?.addEventListener("input",k=>{clearTimeout(I),I=setTimeout(()=>{S.search=k.target.value,u=1,_.clear(),$(),N()},400)}),o?.forEach(k=>{(k.type==="select"||k.type==="combobox")&&(document.getElementById(`filter-${k.name}`)?.addEventListener("change",C=>{S[k.name]=C.target.value;let E=document.getElementById(`filter-sheet-${k.name}`);E&&(E.value=C.target.value),u=1,_.clear(),$(),N()}),document.getElementById(`filter-sheet-${k.name}`)?.addEventListener("change",C=>{S[k.name]=C.target.value;let E=document.getElementById(`filter-${k.name}`);E&&(E.value=C.target.value),u=1,_.clear(),$(),N(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={},D&&(D.value=""),o?.forEach(k=>{let C=document.getElementById(`filter-${k.name}`);C&&(C.value="");let E=document.getElementById(`filter-sheet-${k.name}`);E&&(E.value="")}),u=1,_.clear(),$(),N()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{S={},D&&(D.value=""),o?.forEach(k=>{let C=document.getElementById(`filter-${k.name}`);C&&(C.value="");let E=document.getElementById(`filter-sheet-${k.name}`);E&&(E.value="")}),u=1,_.clear(),$(),N(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ke(null)),c&&document.addEventListener("click",function(k){let C=document.getElementById("aksi-menu-main"),E=document.getElementById("btn-aksi-main");C&&E&&!E.contains(k.target)&&!C.contains(k.target)&&C.classList.remove("show-aksi-menu")});let L=document.getElementById("btn-mobile-filter"),V=document.getElementById("filter-options-wrapper"),W=document.getElementById("btn-close-filter-sheet");if(L&&V&&(L.addEventListener("click",k=>{k.preventDefault(),V.classList.add("sheet-open")}),W&&W.addEventListener("click",k=>{k.preventDefault(),V.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async C=>{let E=C.target,K=E.innerHTML;E.innerHTML="\u23F3 Loading...",E.disabled=!0;try{await c.onExport()}catch{O("Gagal export data")}finally{E.innerHTML=K,E.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let k=document.getElementById(`input-import-${c.moduleName}`);k?.addEventListener("change",async C=>{let E=C.target.files[0];if(!E)return;let K=document.getElementById(`label-import-${c.moduleName}`),R=K?K.querySelector(".import-text"):null,Q=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),K&&(K.style.pointerEvents="none"),k.disabled=!0;try{let F=await Ge(E);if(F.length===0)throw new Error("File kosong atau format salah");await c.onImport(F),j("Import berhasil!"),ce(a),N()}catch(F){O(F.message||"Gagal import data")}finally{R&&(R.innerText=Q),K&&(K.style.pointerEvents="auto"),k.disabled=!1,k.value=""}})}async function N(){$();let k=document.getElementById("table-container");if(!k)return;k.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let C=m==="client",E=C?1:u,K=C?De:20,R=new URLSearchParams({page:E,limit:K,...Object.fromEntries(Object.entries(S).filter(([,U])=>U))}),Q=await T(`${a}?${R}`);if(!Q.ok){k.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Q.data?.error||"Error"}</p></div>`;return}let F=Q.data?.data||Q.data||[],A=Q.data?.pagination,se=F.length;if(C){F=w(F);let U=F.length,X=20,ie=Math.ceil(U/X);u>ie&&ie>0&&(u=ie);let q=(u-1)*X,de=u*X;F=F.slice(q,de),A={page:u,limit:X,total:U,pages:ie}}!1,p&&p(F);let Se=et({columns:r,data:F,onEdit:y?U=>ke(U):null,actions:f.map(U=>({...U,handler:X=>U.handler(X,N)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:h?{selectedIds:_,onToggle:$}:null});k.innerHTML="",k.appendChild(Se);let me=document.getElementById("pagination-container");if(me&&(me.innerHTML="",A&&A.pages>1)){let U=tt({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:X=>{u=X,N()}});U&&me.appendChild(U)}}function xe(k){let C=typeof d=="function"?d(k):d;return Ue(C)}function ke(k){let C=!!k,E=document.createElement("form");if(E.noValidate=!0,E.innerHTML=xe(k),C){let R=typeof d=="function"?d(k):d;nt(E,k)}let{close:K}=oe({title:C?`Edit ${n}`:`Tambah ${n}`,content:E,size:"lg",confirmText:C?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(R,Q)=>{if(!E.reportValidity())return;let F=R.querySelector(".modal-confirm");F.disabled=!0,F.textContent="Menyimpan...";let A=at(E),se=typeof d=="function"?d(k):d,Se=async ie=>{for(let q of ie)if(q.type==="row")await Se(q.fields);else if(q.type==="combobox"&&A[q.name]){let de=A[q.name],_e=(q.options||[]).find(Z=>{let re=String(typeof Z=="object"?Z.value:Z),gt=String(typeof Z=="object"?Z.label:Z);return re===de||gt===de});if(_e)A[q.name]=typeof _e=="object"?_e.value:_e;else if(q.createApi){let Z={};Z[q.createApi.field]=de,q.createApi.extra&&Object.assign(Z,q.createApi.extra);let re=await T(q.createApi.path,{method:"POST",body:JSON.stringify(Z)});if(re.ok&&re.data?.id)A[q.name]=re.data.id;else if(re.ok&&!re.data?.id)A[q.name]=de;else throw new Error(`Gagal membuat master data: ${re.data?.error||"Unknown error"}`)}}};try{await Se(se)}catch(ie){O(ie.message),F.disabled=!1,F.textContent=C?"Simpan Perubahan":`Tambah ${n}`;return}g&&(A=await g(A,k));let me=C?"PUT":"POST",U=C?`${a}/${k.id}`:a,X=await T(U,{method:me,body:JSON.stringify(A)});X.ok?(j(C?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),Q(),ce(a),N()):(O(X.data?.error||"Gagal menyimpan data."),F.disabled=!1,F.textContent=C?"Simpan Perubahan":`Tambah ${n}`)}})}function ne(k){ue(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let C=await T(`${a}/${k.id}`,{method:"DELETE"});C.ok?(j(`${n} berhasil dihapus.`),ce(a),N()):O(C.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return N(),N}M();z();async function Yt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await H(),a=await ae(),r=e?e.get("dash_filter"):null,d={};if(r==="reliever"){let g=new Date,p=`${g.getFullYear()}-${String(g.getMonth()+1).padStart(2,"0")}`;d={status:"Done",month:e&&e.get("month")?e.get("month"):p}}console.log("RAW",await Te()),console.log("OPTIONS",a);let o=g=>g&&!a.find(p=>p.value===g)?[...a,{value:g,label:g}]:a,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=g=>g&&!l.includes(g)?[...l,g]:l,s=new Date().getFullYear(),v=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((g,p)=>{let w=String(p+1).padStart(2,"0");return{value:`${s}-${w}`,label:`${g} ${s}`}});Qt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:d,onDataLoaded:g=>g.sort((p,w)=>{let f=p.backup_date?new Date(p.backup_date).getTime():0;return(w.backup_date?new Date(w.backup_date).getTime():0)-f}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:g=>be(g)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:g=>window.formatDate(g)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:g=>window.formatDate(g)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:g=>g?`<span class="badge badge-info">${g}</span>`:"-"},{key:"status",label:"Status",render:g=>Y(g)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:v},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:g=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:g?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:g?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(g?.original_fc_name),value:g?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(g?.reliever_name),value:g?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:g?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:g?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:g?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:g?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:g?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let g=await T(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(g.ok){let p=g.data.data.map(w=>({Cabang:w.branch_name||"","Nama Facility care":w.original_fc_name||"",Periode:w.period||"",Relifer:w.reliever_name||"","Tanggal Back Up":w.backup_date||"","Tanggal Selesai":w.completion_date||"",Keterangan:w.reason||"",Shift:w.shift||"",Status:w.status||""}));p.length===0&&p.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),B(p,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async g=>{let w=(await T("/api/branches?all=1")).data?.data||[],f=h=>{if(!h)return null;let m=String(h||"").toLowerCase(),x=w.find(u=>String(u.full_name||"").toLowerCase()===m||String(u.code||"").toLowerCase()===m||String(u.name||"").toLowerCase()===m);return x?x.id:null},b=g.map(h=>({branch_name:String(h.Cabang||"").trim(),backup_date:String(h["Tanggal Back Up"]||h["Tanggal Backup"]||"").trim(),original_fc_name:String(h["Nama Facility care"]||h["FC Digantikan"]||"").trim(),reliever_name:String(h.Relifer||h.Reliefer||"").trim(),period:String(h.Periode||"").trim(),reason:String(h.Keterangan||"").trim(),shift:String(h.Shift||"").trim(),completion_date:String(h["Tanggal Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.reliever_name&&h.backup_date),c=await T("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}M();z();async function Wt(t){let e=await H(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));P({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>be(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>Y(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),d=await T(`/api/reports/inspection?limit=10000&${r}`);if(d.ok){let o=d.data.data.map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));B(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let r=n=>{if(!n)return null;let s=String(n||"").toLowerCase(),y=e.find(v=>String(v.label||"").toLowerCase()===s);return y?y.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let s=String(n).trim();if(s===""||s==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);if(/^\d{4,5}$/.test(s)){let v=Number(s);if(v>2e4&&v<99999){let g=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let y=s.split(/[\/\-\.]/);if(y.length===3){let[v,g,p]=y.map(w=>w.trim());if(v.length===4&&g.length<=2&&p.length<=2)return`${v}-${g.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&g.length<=2&&v.length<=2)return`${p}-${g.padStart(2,"0")}-${v.padStart(2,"0")}`}return s},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:d(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),l=await T("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}M();z();async function Xt(t){let e=await H(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));P({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>be(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>Y(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),d=await T(`/api/reports/cleaning?limit=10000&${r}`);if(d.ok){let o=d.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));B(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let r=n=>{if(!n)return null;let s=String(n||"").toLowerCase(),y=e.find(v=>String(v.label||"").toLowerCase()===s);return y?y.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let s=String(n).trim();if(s===""||s==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);if(/^\d{4,5}$/.test(s)){let v=Number(s);if(v>2e4&&v<99999){let g=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let y=s.split(/[\/\-\.]/);if(y.length===3){let[v,g,p]=y.map(w=>w.trim());if(v.length===4&&g.length<=2&&p.length<=2)return`${v}-${g.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&g.length<=2&&v.length<=2)return`${p}-${g.padStart(2,"0")}-${v.padStart(2,"0")}`}return s},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:d(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),l=await T("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}M();z();async function Zt(t,e){let i=await H(),a=Array.from({length:4},(o,l)=>String(new Date().getFullYear()-l)),r=e?e.get("dash_filter"):null,d={};if(r==="fogging"){let o=new Date,l=String(o.getMonth()+1).padStart(2,"0"),n=String(o.getFullYear()),s=e?e.get("month"):null;s&&s.length===7&&(n=s.split("-")[0],l=s.split("-")[1]),d={status:"Done",month:l,year:n}}P({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:d,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>be(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>Y(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let l=new URLSearchParams(o||{}).toString(),n=await T(`/api/reports/fogging?limit=10000&${l}`);if(n.ok){let s=n.data.data.map(y=>({Cabang:y.branch_name||"",Jenis:y.activity_type||"Fogging",Periode:y.period||"",Tanggal:y.activity_date||"",Status:y.status||"","Link Dokumen":y.document_link||""}));B(s,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let l=v=>{if(!v)return null;let g=String(v||"").toLowerCase(),p=i.find(w=>String(w.label||"").toLowerCase()===g);return p?p.value:null},n=v=>{if(v==null||v==="")return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let g=String(v).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let w=Number(g);if(w>2e4&&w<99999){let f=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let p=g.split(/[\/\-\.]/);if(p.length===3){let[w,f,b]=p.map(c=>c.trim());if(w.length===4&&f.length<=2&&b.length<=2)return`${w}-${f.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&f.length<=2&&w.length<=2)return`${b}-${f.padStart(2,"0")}-${w.padStart(2,"0")}`}return g},s=o.map(v=>({branch_id:l(String(v.Cabang||"").trim()),activity_type:String(v.Jenis||v.Kegiatan||"Fogging").trim(),period:String(v.Periode||"").trim(),activity_date:n(v.Tanggal),status:String(v.Status||"").trim(),document_link:String(v["Link Dokumen"]||"").trim(),notes:String(v.Catatan||v.Keterangan||"").trim()})).filter(v=>v.branch_id&&v.period&&v.activity_date),y=await T("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(s)});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}M();z();async function ea(t){let e=await H(),i=await ae(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(l,n)=>String(new Date().getFullYear()-n)),d=l=>l&&!i.find(n=>n.value===l)?[...i,{value:l,label:l}]:i,o=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a;P({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>Y(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),s=await T(`/api/reports/basecamp?limit=10000&${n}`);if(s.ok){let y=s.data.data.map(v=>({"Tgl Info":v.info_date||"",Cabang:v.branch_name||"",Permasalahan:v.problem||"",PIC:v.pic||"","Tgl Done":v.done_date||"",Status:v.status||"",Keterangan:v.notes||""}));B(y,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let n=g=>{if(!g)return null;let p=String(g||"").toLowerCase(),w=e.find(f=>String(f.label||"").toLowerCase()===p);return w?w.value:null},s=g=>{if(g==null||g==="")return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let p=String(g).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let f=Number(p);if(f>2e4&&f<99999){let b=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let w=p.split(/[\/\-\.]/);if(w.length===3){let[f,b,c]=w.map(h=>h.trim());if(f.length===4&&b.length<=2&&c.length<=2)return`${f}-${b.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&b.length<=2&&f.length<=2)return`${c}-${b.padStart(2,"0")}-${f.padStart(2,"0")}`}return p},y=l.map(g=>({info_date:s(g["Tgl Info"]||g["Tanggal Info"]),branch_id:n(String(g.Cabang||"").trim()),problem:String(g.Permasalahan||"").trim(),pic:String(g.PIC||"").trim(),done_date:s(g["Tgl Done"]||g["Tanggal Done"]),status:String(g.Status||"").trim(),notes:String(g.Keterangan||g.Catatan||"").trim()})).filter(g=>g.info_date&&g.branch_id&&g.problem),v=await T("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(y)});if(!v.ok)throw new Error(v.data?.error||"Import gagal");return v.data}}})}async function ta(t){P({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(M(),Ie)),r=await a(`/api/sop?limit=10000&${i}`);if(r.ok){let d=r.data.data.map(l=>({"Nama SOP":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Catatan:l.notes||l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(z(),le));o(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(z(),le));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(M(),Ie)),r=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function aa(t){P({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(M(),Ie)),r=await a(`/api/checklist?limit=10000&${i}`);if(r.ok){let d=r.data.data.map(l=>({"Nama Checklist":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Deskripsi:l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(z(),le));o(d,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(z(),le));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(d=>({name:String(d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(M(),Ie)),r=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}M();ge();z();var Oe={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Ct(t){for(let[e,i]of Object.entries(Oe))if(i.rooms.includes(t))return{zoneName:e,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function na(t){t.innerHTML=`
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
        ${Object.entries(Oe).map(([m,x])=>`
          <button class="zone-tab-btn" data-zone="${m}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${x.icon}</span> ${m}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${m.replace(/[\s&]+/g,"-")}">-</span>
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
  `;let e="all",i="table",a=[],r=ee(),d=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let m=await T("/api/hygiene-standards?all=1");if(!m.ok)throw new Error("Gagal memuat standar hygiene");a=Array.isArray(m.data)?m.data:m.data?.data||[],n(),l(),y()}catch(m){console.error(m),O(m.message||"Terjadi kesalahan memuat data");let x=t.querySelector("#hygiene-view-container");x&&(x.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${m.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,x.querySelector("#btn-retry-hygiene")?.addEventListener("click",o))}}function l(){let m=t.querySelector("#badge-count-all");m&&(m.textContent=a.length),Object.entries(Oe).forEach(([S,_])=>{let $=S.replace(/[\s&]+/g,"-"),D=t.querySelector(`#badge-count-${$}`);if(D){let I=a.filter(L=>_.rooms.includes(L.room_name)).length;D.textContent=I}});let x=new Set(a.map(S=>S.room_name)),u=t.querySelector("#card-total-rooms");u&&(u.textContent=x.size)}function n(){let m=t.querySelector("#hygiene-room-filter");if(!m)return;let x=m.value,u=[];if(e==="all")u=Array.from(new Set(a.map(S=>S.room_name))).filter(Boolean).sort();else{let S=Oe[e];S&&(u=S.rooms.filter(_=>a.some($=>$.room_name===_)))}m.innerHTML=`<option value="">Semua Ruangan (${u.length})</option>`+u.map(S=>`<option value="${S}" ${S===x?"selected":""}>${S}</option>`).join("")}function s(){let m=(t.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),x=t.querySelector("#hygiene-room-filter")?.value||"";return a.filter(u=>{if(e!=="all"){let S=Oe[e];if(!S||!S.rooms.includes(u.room_name))return!1}if(x&&u.room_name!==x)return!1;if(m){let S=(u.room_name||"").toLowerCase(),_=(u.item_name||"").toLowerCase(),$=(u.cleanliness_standard||"").toLowerCase();if(!S.includes(m)&&!_.includes(m)&&!$.includes(m))return!1}return!0})}function y(){let m=s(),x=t.querySelector("#hygiene-stats-text");x&&(x.innerHTML=`Menampilkan <strong>${m.length}</strong> dari <strong>${a.length}</strong> parameter`);let u=t.querySelector("#hygiene-view-container");if(u){if(m.length===0){u.innerHTML=`
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
      `,u.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let S=t.querySelector("#hygiene-search"),_=t.querySelector("#hygiene-room-filter");S&&(S.value=""),_&&(_.value=""),e="all",t.querySelectorAll(".zone-tab-btn").forEach($=>$.classList.toggle("active",$.dataset.zone==="all")),t.querySelectorAll(".zone-tab-btn").forEach($=>{$.dataset.zone==="all"?($.style.background="#2563EB",$.style.color="#fff"):($.style.background="#F1F5F9",$.style.color="var(--text-2)")}),n(),y()});return}i==="table"?v(u,m):g(u,m)}}function v(m,x){m.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${d?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${x.map((u,S)=>{let _=Ct(u.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${S+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${_.bg}; color:${_.color}; border:1px solid ${_.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${_.icon}</span> ${u.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${_.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${u.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${u.cleanliness_standard}
                      </div>
                    </td>
                    ${d?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${u.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${u.id}" data-name="${u.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,p(m)}function g(m,x){let u={};x.forEach(S=>{let _=Ct(S.room_name);u[_.zoneName]||(u[_.zoneName]={info:_,rooms:{}}),u[_.zoneName].rooms[S.room_name]||(u[_.zoneName].rooms[S.room_name]=[]),u[_.zoneName].rooms[S.room_name].push(S)}),m.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(u).map(([S,_])=>{let $=Object.values(_.rooms).reduce((D,I)=>D+I.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${_.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${_.info.bg}; border-bottom:1px solid ${_.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${_.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${_.info.color};">
                      ${S}
                    </h3>
                    <div style="font-size:0.78rem; color:${_.info.color}; opacity:0.85;">
                      ${Object.keys(_.rooms).length} Ruangan \xB7 ${$} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${_.info.color}; border:1px solid ${_.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${$} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(_.rooms).map(([D,I])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${D}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${I.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${d?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${I.map((L,V)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${V+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${L.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${L.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${d?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${L.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${L.id}" data-name="${L.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,p(m)}function p(m){m.querySelectorAll(".btn-edit-item").forEach(x=>{x.addEventListener("click",()=>{let u=x.dataset.id,S=a.find(_=>String(_.id)===String(u));S&&w(S)})}),m.querySelectorAll(".btn-delete-item").forEach(x=>{x.addEventListener("click",()=>{let u=x.dataset.id,S=x.dataset.name;ue(`Yakin ingin menghapus parameter standar <strong>"${S}"</strong>?`,async()=>{try{let _=await T(`/api/hygiene-standards/${u}`,{method:"DELETE"});if(!_.ok)throw new Error(_.data?.error||"Gagal menghapus parameter");j("Parameter kebersihan berhasil dihapus"),o()}catch(_){O(_.message)}},"Hapus Parameter Standar")})})}function w(m=null){let x=!!m,u=Array.from(new Set(a.map(_=>_.room_name))).filter(Boolean).sort(),S=document.createElement("div");S.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${m?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${u.map(_=>`<option value="${_}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${m?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${m?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,oe({title:x?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:S,confirmText:x?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(_,$)=>{let D=S.querySelector("#modal-room-name").value.trim(),I=S.querySelector("#modal-item-name").value.trim(),L=S.querySelector("#modal-standard").value.trim();if(!D||!I||!L){O("Semua kolom wajib diisi");return}try{let V={room_name:D,item_name:I,cleanliness_standard:L},W=x?`/api/hygiene-standards/${m.id}`:"/api/hygiene-standards",xe=await T(W,{method:x?"PUT":"POST",body:JSON.stringify(V)});if(!xe.ok)throw new Error(xe.data?.error||"Gagal menyimpan parameter");j(x?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),$(),o()}catch(V){O(V.message)}}})}function f(){let m=window.open("","_blank");if(!m){O("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let u=`
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

        ${Object.entries(Oe).map(([S,_])=>{let $=a.filter(D=>_.rooms.includes(D.room_name));return $.length===0?"":`
            <div class="zone-title">${_.icon} ${S.toUpperCase()} (${$.length} PARAMETER)</div>
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
                ${$.map((D,I)=>`
                  <tr>
                    <td class="td-no">${I+1}</td>
                    <td class="td-room">${D.room_name}</td>
                    <td class="td-obj">${D.item_name}</td>
                    <td>${D.cleanliness_standard}</td>
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
    `;m.document.open(),m.document.write(u),m.document.close()}function b(){let m=a.map((u,S)=>{let _=Ct(u.room_name);return{No:S+1,"Kategori Zona":_.zoneName,"Ruangan / Lokasi":u.room_name,"Item / Objek Inspeksi":u.item_name,"Standar Kebersihan Fisik":u.cleanliness_standard}}),x=new Date().toISOString().slice(0,10);B(m,`Master_Standar_Kebersihan_FCMS_${x}`),j("Data master kebersihan berhasil diekspor ke Excel")}t.querySelector("#hygiene-search")?.addEventListener("input",()=>y()),t.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>y()),t.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let m=t.querySelector("#hygiene-search"),x=t.querySelector("#hygiene-room-filter");m&&(m.value=""),x&&(x.value=""),y()}),t.querySelectorAll(".zone-tab-btn").forEach(m=>{m.addEventListener("click",()=>{e=m.dataset.zone,t.querySelectorAll(".zone-tab-btn").forEach(x=>{let u=x.dataset.zone===e;if(x.classList.toggle("active",u),u){x.style.background="#2563EB",x.style.color="#fff";let S=x.querySelector(".badge");S&&(S.style.background="rgba(255,255,255,0.25)",S.style.color="#fff")}else{x.style.background="#F1F5F9",x.style.color="var(--text-2)";let S=x.querySelector(".badge");S&&(S.style.background="#CBD5E1",S.style.color="#334155")}}),n(),y()})});let c=t.querySelector("#view-table-btn"),h=t.querySelector("#view-zone-btn");c?.addEventListener("click",()=>{i="table",c.style.background="#fff",c.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",h.style.background="transparent",h.style.boxShadow="none",y()}),h?.addEventListener("click",()=>{i="zone",h.style.background="#fff",h.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",c.style.background="transparent",c.style.boxShadow="none",y()}),t.querySelector("#btn-print-hygiene")?.addEventListener("click",f),t.querySelector("#btn-export-hygiene")?.addEventListener("click",b),t.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>w(null)),o()}M();ge();z();async function Tt(t,e="forms"){if(e==="supply")return Qa(t);Va(t)}function Va(t){P({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await T(`/api/forms?limit=10000&${i}`);a.data?.data?B(a.data.data,"Data_Master_Form"):O("Gagal export data master form")},onImport:async e=>{let i=await T("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Qa(t){let i=((await T("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));P({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>Y(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let r=a?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let d=a?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(o=>o.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),d=await T(`/api/reports/supply?limit=10000&${r}`);if(d.ok){let o=d.data.data.map(l=>{let n=l.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let s=l.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":n||"",Chemical:s||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});B(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let d=(await T("/api/branches?all=1")).data?.data||[],o=y=>{if(!y)return null;let v=String(y||"").toLowerCase(),g=d.find(p=>String(p.full_name||"").toLowerCase()===v||String(p.code||"").toLowerCase()===v||String(p.name||"").toLowerCase()===v);return g?g.id:null},l=y=>{if(y==null||y==="")return"";if(y instanceof Date&&!isNaN(y.getTime()))return y.toISOString().slice(0,10);let v=String(y).trim();if(v===""||v==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(v))return v.slice(0,10);if(/^\d{4,5}$/.test(v)){let p=Number(v);if(p>2e4&&p<99999){let w=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let g=v.split(/[\/\-\.]/);if(g.length===3){let[p,w,f]=g.map(b=>b.trim());if(p.length===4&&w.length<=2&&f.length<=2)return`${p}-${w.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&w.length<=2&&p.length<=2)return`${f}-${w.padStart(2,"0")}-${p.padStart(2,"0")}`}return v},n=a.map(y=>({submitted_at:l(y.Waktu||y.Tanggal),submitter_name:String(y.Pengirim||"").trim(),branch_id:o(String(y.Cabang||"").trim()),tools_items:String(y["Alat/Barang"]||y.Alat||"").trim(),chemical_items:String(y.Chemical||"").trim(),additional_notes:String(y.Catatan||y.Keterangan||"").trim(),status:String(y.Status||"").trim(),processed_by:String(y["Diproses Oleh"]||y.PIC||"").trim()})).filter(y=>y.submitted_at&&y.submitter_name&&y.branch_id),s=await T("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,r)=>{let d=oe({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,l)=>{let n=o.querySelector("#supply-status").value,s=o.querySelector("#supply-processed-by").value;(await T(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:s})})).ok?(j("Status diperbarui."),l(),r()):O("Gagal update status.")}})}}]})}M();z();async function ia(t){let e=ee();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}P({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await T(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));B(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),r=await T("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}M();z();async function ra(t){P({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await T(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)B(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),a=await T("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}M();async function oa(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",r));async function a(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await T(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),i=[]}}async function r(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let l=e.getFullYear(),n=e.getMonth(),s=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),y=document.getElementById("cal-month-label");y&&(y.textContent=s);let v=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(u=>u.value)),g=i.filter(u=>v.has(u.type)),p={};g.forEach(u=>{let S=Ya(u.event_date);S&&(p[S]||(p[S]=[]),p[S].push(u))});let w=new Date(l,n,1).getDay(),f=new Date(l,n+1,0).getDate(),b=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],c=new Date().toISOString().slice(0,10),h='<div class="calendar-grid">';b.forEach(u=>{h+=`<div class="cal-day-header">${u}</div>`});for(let u=0;u<w;u++)h+='<div class="cal-cell cal-cell-empty"></div>';for(let u=1;u<=f;u++){let S=`${l}-${String(n+1).padStart(2,"0")}-${String(u).padStart(2,"0")}`,_=p[S]||[],$=S===c;h+=`
          <div class="cal-cell ${$?"cal-today":""} ${_.length?"cal-has-events":""}"
               data-date="${S}" tabindex="0" role="button" aria-label="${S}">
            <div class="cal-day-num ${$?"today-num":""}">${u}</div>
            <div class="cal-events-preview">
              ${_.slice(0,3).map(D=>`
                <div class="cal-event-dot cal-color-${D.color||"gray"}" title="${ve(D.title||D.type)}">
                  <span class="cal-event-dot-label">${Wa(D.title||D.branch_name||D.type,18)}</span>
                </div>
              `).join("")}
              ${_.length>3?`<div class="cal-more">+${_.length-3} lagi</div>`:""}
            </div>
          </div>`}let x=(w+f)%7;if(x!==0)for(let u=0;u<7-x;u++)h+='<div class="cal-cell cal-cell-empty"></div>';h+="</div>",o.innerHTML=h,o.querySelectorAll(".cal-cell[data-date]").forEach(u=>{u.addEventListener("click",()=>{let S=u.dataset.date,_=p[S]||[];if(!_.length)return;let $=document.getElementById("cal-event-list"),D=new Date(S+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=D,document.getElementById("cal-event-items").innerHTML=_.map(I=>`
            <div class="cal-event-item cal-color-border-${I.color||"gray"}">
              <div class="cal-event-type">
                ${Xa(I.type)}
                ${I.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${ve(I.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${ve(I.title||"-")}</div>
              ${I.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${ve(I.branch_name)}</div>`:""}
              ${I.pic||I.period||I.opening_date||I.target_date||I.completion_date||I.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${I.pic?`<div>\u{1F464} <b>PIC:</b> ${ve(I.pic)}</div>`:""}
                  ${I.period?`<div>\u{1F4C5} <b>Periode:</b> ${ve(I.period)}</div>`:""}
                  ${I.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${ve(Dt(I.opening_date))}</div>`:""}
                  ${I.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${ve(Dt(I.target_date))}</div>`:""}
                  ${I.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${ve(Dt(I.completion_date))}</div>`:""}
                  ${I.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${ve(I.notes)}</div>`:""}
                </div>
              `:""}
              ${I.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${I.status==="Done"?"badge-success":"badge-warning"}">${ve(I.status)}</span></div>`:""}
              ${I.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${I.days_remaining} hari</div>`:""}
            </div>
          `).join(""),$.style.display="block"})})}catch(l){console.error("[Calendar] Render error:",l),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let d=()=>{a().then(()=>r())};it.on("data:changed",d),r()}function Ya(t){if(!t||t==="-"||String(t).trim()==="")return"";let e=String(t).trim();if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);let i=e.split(/[\/\-\.]/);if(i.length===3){let[a,r,d]=i.map(o=>o.trim());if(d.length===4)return`${d}-${r.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${r.padStart(2,"0")}-${d.padStart(2,"0")}`}return e.slice(0,10)}function Dt(t){if(!t||t==="-"||String(t).trim()==="")return"";let i=String(t).trim().split(/[\/\-\.]/);if(i.length===3){let[a,r,d]=i.map(o=>o.trim());if(a.length===4)return`${d.padStart(2,"0")}-${r.padStart(2,"0")}-${a}`;if(d.length===4)return`${a.padStart(2,"0")}-${r.padStart(2,"0")}-${d}`}return t}function Wa(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function ve(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Xa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[t]||t}M();async function sa(t){let e=ee(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${r}18;color:${r};margin-top:6px">
              ${(e?.role||"viewer").toUpperCase()}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">\u{1F4E7} Email</span>
            <span class="info-value">${e?.email||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F464} Username</span>
            <span class="info-value">${e?.username||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F3AF} Role</span>
            <span class="info-value" style="color:${r};font-weight:700">${e?.role||"\u2014"}</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: Change Password -->
      <div class="chart-card">
        ${e?.role!=="viewer"?`
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
  `;let d=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(d&&o)try{let l=JSON.parse(atob(d.split(".")[1])),n=new Date(l.exp*1e3);o.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let n=document.getElementById("pwd-error"),s=document.getElementById("pwd-success"),y=document.getElementById("btn-save-pwd");n.style.display="none",s.style.display="none";let v=l.target,g=v.current_password.value,p=v.new_password.value,w=v.confirm_password.value;if(p!==w){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(p.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}y.disabled=!0,y.textContent="\u23F3 Menyimpan...";let f=await T("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:g,new_password:p})});y.disabled=!1,y.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',f.ok?(s.textContent="\u2705 Password berhasil diubah.",s.style.display="block",v.reset(),j("Password berhasil diubah.")):(n.textContent=f.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}M();var mt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let r=Number(e);if(r>2e4&&r<99999){let d=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[r,d,o]=i.map(y=>y.trim()),l=Number(r),n=Number(d),s=Number(o);if(r.length===4&&l>1900)return`${r}-${d.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&s>1900)return l>12?`${o}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`:n>12?`${o}-${r.padStart(2,"0")}-${d.padStart(2,"0")}`:`${o}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`;if(o.length===2&&!isNaN(s)){let y=s>=50?`19${o}`:`20${o}`;return l>12?`${y}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`:`${y}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function la(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Za={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||t.Divisi||"FACILITY CARE",phone:t["No. Hp"]||t["No. HP"],join_date:te(t["Tanggal Masuk"]||t["Tgl Masuk"]),status:t.Status||"",target_pindah_os:te(t["Target Pindah OS"]),target_selesai_os:te(t["Target Selesai OS"]),notes:t.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function en(t,e){let i=mt[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Za[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],d=[],o=[];return e.filter(n=>!la(n)).forEach((n,s)=>{let y=e.indexOf(n)+2,v=[];a.required.forEach(({key:p,label:w})=>{let f=n[p];if(f==null||String(f).trim()===""){let b=Object.keys(n).filter(c=>c.trim()).join(", ");v.push({column:w,originalValue:f||"",reason:`Kolom "${w}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${b.slice(0,120)}`})}});let g=a.map(n);v.length>0?d.push({row:y,data:g,raw:n,errors:v}):(r.push(n),o.push(g))}),{valid:r,errors:d,mapped:o}}function da(t){let e=[];return t.SheetNames.forEach(i=>{let a=mt[i];if(!a)return;let r=t.Sheets[i],d=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=en(i,d),l=d.filter(n=>!la(n));e.push({sheetName:i,module:a.module,label:a.label,total:l.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function ca(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,r])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(r),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function pa(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(r=>{if(!r.errors||r.errors.length===0)return;a=!0;let d=r.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(l.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([n,s])=>[n,s??""]))})),o=e.utils.json_to_sheet(d);e.utils.book_append_sheet(i,o,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var tn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ma(t){t.innerHTML=`
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
              ${Object.entries(mt).map(([f,{label:b}])=>`<span class="import-sheet-tag">\u{1F4C4} ${f} \u2192 ${b}</span>`).join("")}
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
  `;let e=null,i=null,a=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(f){Object.entries(r).forEach(([b,c])=>{c.style.display=b===f?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let f=document.getElementById("btn-backup-db");f.disabled=!0,f.textContent="\u23F3 Memproses Backup...";try{let b=await T("/api/import/backup");if(b.ok){if(!window.XLSX){O("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let c=window.XLSX,h=c.utils.book_new();Object.entries(b.data.database).forEach(([m,x])=>{let u=x.length>0?x:[{}],S=c.utils.json_to_sheet(u);c.utils.book_append_sheet(h,S,m.substring(0,31))}),c.writeFile(h,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),j("Backup berhasil diunduh!")}else O("Gagal memproses backup: "+(b.data?.error||"Unknown error"))}catch(b){O("Gagal memproses backup: "+b.message)}finally{f.disabled=!1,f.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let f=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let b=await T("/api/sync/google-sheets",{method:"POST"});b.ok?alert("Sinkronisasi Berhasil: "+(b.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(b.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=f,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{ca(),j("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",f=>{f.stopPropagation(),l.click()}),l.addEventListener("change",f=>{f.target.files[0]&&s(f.target.files[0])}),n.addEventListener("dragover",f=>{f.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",f=>{f.preventDefault(),n.classList.remove("drag-over");let b=f.dataTransfer.files[0];b&&b.name.match(/\.xlsx?$/i)?s(b):O("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",d("upload")});async function s(f){e=f,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${f.name} (${(f.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await y(f)}async function y(f){d("validating");let b=document.getElementById("validation-status"),c=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");b.textContent="Membaca file Excel...",c.style.width="20%",await Ye(200);let h=await f.arrayBuffer(),m=window.XLSX.read(h,{type:"array",cellDates:!0});b.textContent=`Memvalidasi ${m.SheetNames.length} sheet...`,c.style.width="50%",await Ye(100),i=da(m),c.style.width="100%",b.textContent="Validasi selesai!",await Ye(300),v()}catch(h){d("upload"),O("Gagal memproses file: "+h.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function v(){d("preview");let f=i.filter($=>!$.skipped).length,b=i.reduce(($,D)=>$+D.total,0),c=i.reduce(($,D)=>$+D.valid,0),h=i.reduce(($,D)=>$+D.errorCount,0),m=b>0?Math.round(c/b*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${f} sheet</span>
      <span class="badge badge-secondary">${b} baris</span>
      <span class="badge badge-success">${c} valid (${m}%)</span>
      ${h>0?`<span class="badge badge-danger">${h} error</span>`:""}
    `;let x=document.getElementById("preview-table-container");x.innerHTML=`
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
          ${i.map(($,D)=>`
            <tr class="${$.errorCount>0?"row-error":$.skipped?"row-skipped":"row-ok"}">
              <td><strong>${$.sheetName}</strong></td>
              <td>${$.label}</td>
              <td style="text-align:center">${$.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${$.valid}</span></td>
              <td style="text-align:center">${$.errorCount>0?`<span class="badge badge-danger">${$.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${$.skipped?'<span class="badge badge-neutral">Dilewati</span>':$.errorCount>0&&$.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':$.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':$.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${$.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${D}">\u{1F50D} ${$.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,x.querySelectorAll(".btn-detail-error").forEach($=>{$.addEventListener("click",()=>{let D=i[Number($.dataset.idx)];g(D)})});let u=document.getElementById("error-detail-section"),S=document.getElementById("error-detail-container");S.innerHTML="",u.style.display="none";let _=document.getElementById("btn-start-import");c===0?(_.disabled=!0,_.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(_.disabled=!1,h>0?(_.innerHTML=`\u{1F680} Import ${c} Data Valid (${h} dilewati)`,_.title="Baris error akan dilewati, baris valid tetap diimport"):_.innerHTML=`\u{1F680} Mulai Import ${c} Data`)}function g(f){let b=document.getElementById("error-detail-section"),c=document.getElementById("error-detail-container");b.style.display="";let h=f.errors.slice(0,100).map(m=>(Array.isArray(m.errors)?m.errors:[]).map(u=>{let S=typeof u=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${m.row}</span></td>
            <td><strong>${S?u.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${S&&u.originalValue!==void 0?u.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${S?u.reason:u}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${S&&u.aliases?`Gunakan salah satu nama kolom:<br><em>${u.aliases}</em>`:S&&u.hint?u.hint:""}
            </td>
          </tr>
        `}).join("")).join("");c.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${f.sheetName} \u2014 ${f.errorCount} baris error dari ${f.total} total
          ${f.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${h||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${f.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,b.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;pa(i)?j("Log error berhasil didownload."):j("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let f=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";p(f)});async function p(f){d("importing"),a=Date.now();let b=[];tn.forEach(u=>{let S=i?.find(_=>_.module===u&&_.mapped?.length>0);S&&b.push(S)});let c=document.getElementById("import-steps-list");c.innerHTML=b.map(u=>`
      <div class="import-step-item" id="step-item-${u.module}">
        <span class="step-item-icon" id="step-icon-${u.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${u.label} <span class="step-item-count">(${u.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${u.module}"></span>
      </div>
    `).join("");let h=document.getElementById("import-bar"),m=document.getElementById("import-current-status"),x={totalSheets:b.length,totalRows:b.reduce((u,S)=>u+S.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let u=0;u<b.length;u++){let S=b[u],_=document.getElementById(`step-icon-${S.module}`),$=document.getElementById(`step-status-${S.module}`);_.textContent="\u{1F504}",$.textContent="Mengimport...",m.textContent=`Mengimport ${S.label}...`,h.style.width=`${Math.round(u/b.length*100)}%`;try{let D=await T(`/api/import/${S.module}`,{method:"POST",body:JSON.stringify({rows:S.mapped,onDuplicate:f})});if(D.ok){let I=D.data;x.inserted+=I.inserted||0,x.skipped+=I.skipped||0,x.moduleResults.push({label:S.label,inserted:I.inserted||0,skipped:I.skipped||0,status:"ok"}),_.textContent="\u2705",$.innerHTML=`<span class="badge badge-success">${I.inserted||0} berhasil</span>${I.skipped>0?` <span class="badge badge-neutral">${I.skipped} skip</span>`:""}`}else x.failed++,x.moduleResults.push({label:S.label,inserted:0,skipped:0,status:"error",error:D.data?.error}),_.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(D){x.failed++,x.moduleResults.push({label:S.label,inserted:0,skipped:0,status:"error",error:D.message}),_.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ye(150)}h.style.width="100%",m.textContent="Selesai!",await Ye(400),w(x)}function w(f){d("summary");let b=((Date.now()-a)/1e3).toFixed(1),c=f.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${c?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${c?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${f.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${f.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${f.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${f.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${f.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${f.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${b}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${f.moduleResults.map(h=>`
            <tr>
              <td>${h.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${h.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${h.skipped}</span></td>
              <td style="text-align:center">
                ${h.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${h.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ye(t){return new Promise(e=>setTimeout(e,t))}M();var ut=[],ua=[];async function ga(t){ut=await H(),ua=await ae(),P({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ut}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await T(`/api/sp?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(z(),le));d(r,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(z(),le));i(e,"Template_Import_SP")},onImport:async e=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=ut.find(s=>String(s.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let s=Number(l);if(s>2e4&&s<99999){let y=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[s,y,v]=n.map(g=>g.trim());if(s.length===4&&y.length<=2&&v.length<=2)return`${s}-${y.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&y.length<=2&&s.length<=2)return`${v}-${y.padStart(2,"0")}-${s.padStart(2,"0")}`}return l},r=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:i(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),d=await T("/api/import/sp",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ua},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:ut,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}M();var Re=[],ba=[];async function ha(t){Re=await H(),ba=await ae(),P({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Re},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Re}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await T(`/api/mutasi?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(z(),le));d(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(z(),le));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=Re.find(s=>String(s.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let s=Number(l);if(s>2e4&&s<99999){let y=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[s,y,v]=n.map(g=>g.trim());if(s.length===4&&y.length<=2&&v.length<=2)return`${s}-${y.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&y.length<=2&&s.length<=2)return`${v}-${y.padStart(2,"0")}-${s.padStart(2,"0")}`}return l},r=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:i(String(o["Cabang Asal"]||"").trim()),to_branch_id:i(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),d=await T("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ba},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Re,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Re,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}M();var It=[],fa=[];async function ya(t){It=await H(),fa=await Te(),P({container:t,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:It}],exportOptions:{moduleName:"overtime_records",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await T(`/api/overtime?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.date||"","Cabang Lembur":o.branch_name||"","Nama Karyawan":o.employee_name||"","Jam Mulai":o.start_time||"","Jam Selesai":o.end_time||"","Istirahat (Jam)":o.break_hours||0,"Total Jam":o.total_hours||0,"Alasan / Ket Lembur":o.reason||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(z(),le));d(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:It},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:fa},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",t=>{if(t.target.name==="start_time"||t.target.name==="end_time"||t.target.name==="break_hours"){let e=t.target.closest("form");if(!e)return;let i=e.querySelector('[name="start_time"]'),a=e.querySelector('[name="end_time"]'),r=e.querySelector('[name="break_hours"]'),d=e.querySelector('[name="total_hours"]');if(!i||!a||!r||!d)return;let o=i.value,l=a.value,n=parseFloat(r.value)||0;if(o&&l){let[s,y]=o.split(":").map(Number),[v,g]=l.split(":").map(Number),p=v*60+g-(s*60+y);p<0&&(p+=24*60);let w=p/60-n;w<0&&(w=0),d.value=Number.isInteger(w)?w:w.toFixed(2)}}}));M();async function va(t){let e=ee();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}P({container:t,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let d=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(s=>String(s.id)===String(i));if(!d)return alert("Data tidak ditemukan");let o=s=>{if(!s)return"Tidak ada data";try{return JSON.stringify(JSON.parse(s),null,2)}catch{return s}},l=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(d.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(d.new_data)}</pre>
           </div>
         </div>
       `,{createModal:n}=await Promise.resolve().then(()=>(ge(),vt));n({title:`Detail Audit Log #${i}`,content:l,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}M();ge();z();var Ft={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},We={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function xa(t){t.innerHTML=`
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
  `;let e="board",i=null,a=[],r=ee(),d=r&&["superadmin","admin","manager","spv"].includes(r.role);T("/api/branches?all=1").then(c=>{c.ok&&c.data?.data&&(a=c.data.data)}).catch(()=>{});async function o(){try{let c=await T("/api/os-migrations/summary");if(!c.ok)throw new Error("Gagal memuat data migrasi");i=c.data?.data||c.data,l(i),s()}catch(c){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(c){let h=document.getElementById("migration-summary-cards");if(!h||!c)return;let m=c.total_branches||0,x=c.vendors?.AEON?.total||0,u=c.vendors?.["Triple S"]?.total||0,S=c.by_status?.completed||0,_=c.by_status?.in_progress||0,$=c.by_status?.scheduled||0,D=m>0?Math.round(S/m*100):0;h.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${m} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${x} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${u} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${D}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${D}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${S}</strong></span>
          <span>\u{1F535} Proses: <strong>${_}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${$}</strong></span>
        </div>
      </div>
    `}function n(){if(!i||!i.matrix)return[];let c=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),h=document.getElementById("filter-vendor")?.value||"",m=document.getElementById("filter-month")?.value||"",x=document.getElementById("filter-status")?.value||"";return i.matrix.filter(u=>!(h&&u.vendor!==h||m&&u.target_month!==m||x&&u.status!==x||c&&!`${u.branch_code} ${u.branch_name} ${u.branch_full_name} ${u.vendor} ${u.notes||""} ${u.pic_vendor||""}`.toLowerCase().includes(c)))}function s(){let c=n(),h=document.getElementById("filter-stats-text");h&&(h.innerHTML=`Menampilkan <strong>${c.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let m=document.getElementById("migration-view-container");e==="board"?y(m,c):v(m,c)}function y(c,h){if(!h.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let m=["AEON","Triple S"],x=["2026-09","2026-10","2026-11","2026-12"],u="";m.forEach(S=>{let _=h.filter(W=>W.vendor===S);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==S||_.length===0)return;let $=_.length,D=S==="AEON"?"#1E40AF":"#0F766E",I=S==="AEON"?"#EFF6FF":"#F0FDFA",L=S==="AEON"?"#BFDBFE":"#99F6E4",V=x.filter(W=>_.some(N=>N.target_month===W));u+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${I}; border-bottom:1px solid ${L}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${D};">${S==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${D}; border:1px solid ${L}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${$} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${V.map(W=>{let N=_.filter(ne=>ne.target_month===W);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Ft[W]||W).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${N.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${N.map(ne=>{let k=We[ne.status]||We.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${ne.branch_code}.</span>${ne.branch_name}
                            </div>
                            ${d?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${ne.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${ne.id}" data-name="${ne.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${ne.id}" data-status="${ne.status}" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${k.icon} ${k.label} \u25BE
                              </button>
                            </div>

                            ${ne.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${ne.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${ne.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${ne.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${ne.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),c.innerHTML=u,g(c)}function v(c,h){if(!h.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}c.innerHTML=`
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
                ${d?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${h.map((m,x)=>{let u=We[m.status]||We.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${x+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${m.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${m.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${m.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${m.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Ft[m.target_month]||m.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${m.id}" data-status="${m.status}" style="background:${u.bg}; color:${u.color}; border:1px solid ${u.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${u.icon} ${u.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${m.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${m.employee_count>0?`\u{1F464} ${m.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${m.notes||"-"}
                    </td>
                    ${d?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${m.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${m.id}" data-name="${m.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,g(c)}function g(c){c.querySelectorAll(".status-toggle-btn").forEach(h=>{h.addEventListener("click",m=>{m.stopPropagation();let x=h.dataset.id,u=h.dataset.status;p(h,x,u)})}),c.querySelectorAll(".edit-card-btn").forEach(h=>{h.addEventListener("click",()=>{let m=h.dataset.id,x=i?.matrix?.find(u=>String(u.id)===String(m));x&&w(x)})}),c.querySelectorAll(".delete-card-btn").forEach(h=>{h.addEventListener("click",()=>{let m=h.dataset.id,x=h.dataset.name||"Cabang";ue(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${x}</strong>?`,async()=>{try{if(!(await T(`/api/os-migrations/${m}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");j(`Jadwal migrasi ${x} berhasil dihapus`),o()}catch(u){O(u.message)}})})})}function p(c,h,m){document.querySelectorAll(".status-popup-menu").forEach($=>$.remove());let x=c.getBoundingClientRect(),u=document.createElement("div");u.className="status-popup-menu",u.style.cssText=`
      position: fixed;
      top: ${x.bottom+window.scrollY+4}px;
      left: ${x.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let S=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];u.innerHTML=S.map($=>`
      <button class="status-option-item" data-key="${$.key}" style="
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
        color: ${$.key===m?"#2563EB":"var(--text-1)"};
        font-weight: ${$.key===m?"700":"500"};
      ">
        ${$.label} ${$.key===m?"\u2713":""}
      </button>
    `).join("");let _=$=>{!u.contains($.target)&&$.target!==c&&(u.remove(),document.removeEventListener("click",_))};setTimeout(()=>document.addEventListener("click",_),50),u.querySelectorAll(".status-option-item").forEach($=>{$.addEventListener("click",async()=>{let D=$.dataset.key;if(u.remove(),document.removeEventListener("click",_),D!==m)try{if(!(await T(`/api/os-migrations/${h}/status`,{method:"PATCH",body:JSON.stringify({status:D})})).ok)throw new Error("Gagal memperbarui status");j("Status berhasil diubah"),o()}catch(I){O(I.message)}})}),document.body.appendChild(u)}function w(c=null){let h=!!c,m=h?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",u=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${a.map(S=>`
      <option value="${S.id}" ${c&&c.branch_id===S.id?"selected":""}>
        ${S.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${c?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${c?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${c&&!["AEON","Triple S"].includes(c.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${c?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${c?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${c?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${c?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${c?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${c?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${c?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${c?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${c?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${c?.notes||""}</textarea>
        </div>
      </form>
    `;oe({title:m,content:u,confirmText:h?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(S,_)=>{let $=S.querySelector("#migration-form");if(!$.checkValidity()){$.reportValidity();return}let D=new FormData($),I={branch_id:parseInt(D.get("branch_id"),10),vendor:D.get("vendor"),target_month:D.get("target_month"),status:D.get("status"),actual_date:D.get("actual_date")||null,pic_vendor:D.get("pic_vendor")||null,notes:D.get("notes")||null};try{let L=h?`/api/os-migrations/${c.id}`:"/api/os-migrations",W=await T(L,{method:h?"PUT":"POST",body:JSON.stringify(I)});if(!W.ok){let N=await W.json().catch(()=>({}));throw new Error(N.error||"Gagal menyimpan data")}j(h?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),_(),o()}catch(L){O(L.message)}}})}let f=document.getElementById("view-board-btn"),b=document.getElementById("view-table-btn");f.addEventListener("click",()=>{e="board",f.style.background="#fff",f.classList.remove("btn-ghost"),b.style.background="transparent",b.classList.add("btn-ghost"),s()}),b.addEventListener("click",()=>{e="table",b.style.background="#fff",b.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),s()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(c=>{let h=document.getElementById(c);h&&(h.addEventListener("input",s),h.addEventListener("change",s))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{w()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let c=n();if(!c.length){O("Tidak ada data untuk diexport");return}let h=c.map((m,x)=>({No:x+1,"Kode Cabang":m.branch_code,"Nama Cabang":m.branch_name,"Cabang Lengkap":m.branch_full_name,"Vendor OS":m.vendor,"Target Bulan":Ft[m.target_month]||m.target_month,Status:We[m.status]?.label||m.status,"Tanggal Realisasi":m.actual_date||"","Jml Personil FC":m.employee_count||0,"PIC Vendor":m.pic_vendor||"",Catatan:m.notes||""}));B(h,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),j("File Excel berhasil diunduh")}),o()}M();ge();z();var ka={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function wa(t){t.innerHTML=`
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
  `;let e="pagi",i="timeline",a=[],r=ee(),d=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let c=await T("/api/dac?all=1");if(!c.ok)throw new Error("Gagal memuat standar DAC");a=c.data?.data||c.data||[];let h=a.filter(S=>S.shift==="pagi").length,m=a.filter(S=>S.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=h,document.getElementById("badge-count-siang").textContent=m,document.getElementById("badge-count-all").textContent=a.length;let x=document.getElementById("dac-category-filter"),u=Array.from(new Set(a.map(S=>S.category))).sort();x.innerHTML='<option value="">Semua Kategori</option>'+u.map(S=>`<option value="${S}">${S}</option>`).join(""),n(),s()}catch(c){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(){let c=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),h=document.getElementById("dac-category-filter")?.value||"",m=document.getElementById("dac-period-filter")?.value||"";return a.filter(x=>{if(e!=="all"&&x.shift!==e||h&&x.category!==h)return!1;if(m==="morning"){let u=parseInt(x.start_time.split(":")[0],10);if(x.shift==="pagi"&&u>=12||x.shift==="siang"&&u>=17)return!1}else if(m==="afternoon"){let u=parseInt(x.start_time.split(":")[0],10);if(x.shift==="pagi"&&u<12||x.shift==="siang"&&u<17)return!1}return!(c&&!`${x.start_time} ${x.end_time} ${x.category} ${x.location} ${x.activity}`.toLowerCase().includes(c))})}function n(){let c=document.getElementById("dac-summary-cards");if(!c)return;let h=a.filter(_=>_.shift==="pagi"),m=a.filter(_=>_.shift==="siang"),x=e==="pagi"?h:e==="siang"?m:a,u=e==="pagi"?"08:00 \u2013 17:00":e==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",S=e==="pagi"?"Shift Pagi":e==="siang"?"Shift Siang":"Semua Shift";c.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${S})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${x.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${u}</div>
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
    `}function s(){let c=l(),h=document.getElementById("dac-stats-text");h&&(h.innerHTML=`Menampilkan <strong>${c.length}</strong> aktivitas`);let m=document.getElementById("dac-view-container");i==="timeline"?y(m,c):v(m,c)}function y(c,h){if(!h.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}c.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${h.map(m=>{let x=ka[m.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},u=m.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",S=m.shift==="pagi"?"#FEF3C7":"#E0E7FF",_=m.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${m.start_time} - ${m.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${m.item_no}
                </div>
                ${e==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${S}; color:${_}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${u}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${m.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${m.location}
                    </span>
                  </div>
                  ${d?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${m.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${m.id}" data-name="${m.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${x.color}; margin-top:8px;">
                  ${m.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,g(c)}function v(c,h){if(!h.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}c.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                ${e==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${d?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${h.map(m=>{let x=ka[m.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${m.item_no}</td>
                    ${e==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${m.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${m.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${m.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${m.start_time} - ${m.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${m.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${m.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${m.activity}
                    </td>
                    ${d?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${m.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${m.id}" data-name="${m.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,g(c)}function g(c){c.querySelectorAll(".edit-dac-btn").forEach(h=>{h.addEventListener("click",()=>{let m=h.dataset.id,x=a.find(u=>String(u.id)===String(m));x&&p(x)})}),c.querySelectorAll(".delete-dac-btn").forEach(h=>{h.addEventListener("click",()=>{let m=h.dataset.id,x=h.dataset.name||"item";ue(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${x}</strong>?`,async()=>{try{if(!(await T(`/api/dac/${m}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");j("Aktivitas DAC berhasil dihapus"),o()}catch(u){O(u.message)}})})})}function p(c=null){let h=!!c,m=h?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",x=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],u=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${c?.shift==="pagi"||!c&&e==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${c?.shift==="siang"||!c&&e==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${c?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${c?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${c?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${c?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${x.map(S=>`<option value="${S}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${c?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${c?.activity||""}</textarea>
        </div>
      </form>
    `;oe({title:m,content:u,confirmText:h?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(S,_)=>{let $=S.querySelector("#dac-form");if(!$.checkValidity()){$.reportValidity();return}let D=new FormData($),I={shift:D.get("shift"),item_no:D.get("item_no")?parseInt(D.get("item_no"),10):void 0,start_time:D.get("start_time"),end_time:D.get("end_time"),category:D.get("category"),location:D.get("location"),activity:D.get("activity")};try{let L=h?`/api/dac/${c.id}`:"/api/dac",W=await T(L,{method:h?"PUT":"POST",body:JSON.stringify(I)});if(!W.ok){let N=await W.json().catch(()=>({}));throw new Error(N.error||"Gagal menyimpan aktivitas")}j(h?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),_(),o()}catch(L){O(L.message)}}})}function w(){let c=l();if(!c.length){O("Tidak ada data untuk dicetak");return}let h=e==="pagi"?"SHIFT PAGI (08:00 - 17:00)":e==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",m=window.open("","_blank");if(!m){O("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}m.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${h}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${h}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:35px; text-align:center;">NO</th>
              ${e==="all"?'<th style="width:55px;">SHIFT</th>':""}
              <th style="width:105px;">JAM</th>
              <th style="width:130px;">KATEGORI</th>
              <th style="width:180px;">LOKASI / AREA</th>
              <th>RINCIAN AKTIVITAS</th>
            </tr>
          </thead>
          <tbody>
            ${c.map(x=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${x.item_no}</td>
                ${e==="all"?`<td>${x.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${x.start_time} - ${x.end_time}</td>
                <td><strong>${x.category}</strong></td>
                <td><strong>${x.location}</strong></td>
                <td>${x.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),m.document.close(),m.focus(),setTimeout(()=>{m.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(c=>{c.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(h=>{h.style.background="#F1F5F9",h.style.color="var(--text-2)"}),c.style.background="#2563EB",c.style.color="#fff",e=c.dataset.shift,n(),s()})});let f=document.getElementById("view-timeline-btn"),b=document.getElementById("view-table-btn");f.addEventListener("click",()=>{i="timeline",f.style.background="#fff",f.classList.remove("btn-ghost"),b.style.background="transparent",b.classList.add("btn-ghost"),s()}),b.addEventListener("click",()=>{i="table",b.style.background="#fff",b.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),s()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(c=>{let h=document.getElementById(c);h&&(h.addEventListener("input",s),h.addEventListener("change",s))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{p()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{w()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let c=l();if(!c.length){O("Tidak ada data untuk diexport");return}let h=c.map(m=>({No:m.item_no,Shift:m.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":m.start_time,"Jam Selesai":m.end_time,Kategori:m.category,"Lokasi / Area":m.location,"Rincian Aktivitas":m.activity}));B(h,`DAC_Daily_Activity_Cleaning_${e}_${new Date().toISOString().slice(0,10)}`),j("File Excel berhasil diunduh")}),o()}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let i=Math.floor(Number(t)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=t.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let i=e[1],a=parseInt(e[2],10),r=parseInt(e[3],10);if(a>12&&r<=12)return`${i}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),d=a[parseInt(i[1],10)-1];return`${r} ${d} ${i[0]}`}return e};function G(t){return async e=>{if(!je()){Ce("/login");return}return t(e)}}var Xe=null;function an(){Xe&&clearInterval(Xe);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");r&&(r.textContent=i),d&&(d.textContent=a)};t(),Xe=setInterval(t,1e3)}async function nn(){try{let t=await T("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,r)=>{let d=document.getElementById(a);d&&(d.textContent=r>0?r:"",d.style.display=r>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Ke=[];async function rn(){try{let t=await T("/api/dashboard/notifications");if(!t.ok)return;Ke=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ke.length>0?"block":"none",e.textContent=Ke.length)}catch{}}function on(){if(!Ke.length){oe({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ke.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;oe({title:`Notifikasi (${Ke.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function Sa(){let t=ee(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
            <a href="#/reports/supply" class="nav-item" data-route="/reports/supply">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 01-8 0"/></svg>
              </span>
              <span class="nav-label">Permintaan Chemical</span>
              <span class="nav-badge" id="badge-supply"></span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop" class="nav-item" data-route="/sop">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/></svg>
              </span>
              <span class="nav-label">SOP</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Checklist</span>
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
            <a href="#/forms" class="nav-item" data-route="/forms">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              </span>
              <span class="nav-label">Master Form</span>
            </a>
          </div>

          <!-- Admin -->
          ${t&&(t.role==="superadmin"||t.role==="admin")?`
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
            <div class="sidebar-avatar">${(t?.full_name||t?.username||"U")[0].toUpperCase()}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${t?.full_name||t?.username||"Guest"}</div>
              <div class="sidebar-user-role">${t?.role||"Viewer"}</div>
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
                <span class="topbar-greeting-time">${(()=>{let s=new Date().getHours();return s>=4&&s<11?"Selamat Pagi":s>=11&&s<15?"Selamat Siang":s>=15&&s<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||t?.username||"Guest"}</span> \u{1F44B}
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
              <img src="https://ui-avatars.com/api/?name=${encodeURIComponent(t?.full_name||t?.username||"Guest")}&background=2563EB&color=fff&bold=true" class="topbar-avatar" alt="Avatar" />
              <div class="topbar-user-text">
                <span class="topbar-user-name">${t?.full_name||t?.username||"Guest"}</span>
                <span class="topbar-user-role-mini">${t?.role||"Viewer"}</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--gray-400)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),o=()=>{i.classList.add("open"),a.classList.add("show")},l=()=>{i.classList.remove("open"),a.classList.remove("show")};r?.addEventListener("click",o),d?.addEventListener("click",l),a?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(s=>s.addEventListener("click",l)),t&&t.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(s=>{s.dataset.route!=="/overtime"&&(s.style.display="none")}),document.querySelectorAll(".nav-section").forEach(s=>{Array.from(s.querySelectorAll(".nav-item")).some(v=>v.style.display!=="none")||(s.style.display="none")}));function n(){let s=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(g=>{let p=g.dataset.route;g.classList.toggle("active",s===p||p!=="/dashboard"&&s.startsWith(p))});let y=document.getElementById("topbar-title"),v=document.querySelector(".nav-item.active .nav-label");y&&v&&(y.textContent=v.textContent)}window.addEventListener("hashchange",n),n(),an(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await T("/api/auth/logout",{method:"POST"}),qe(),Xe&&clearInterval(Xe),Ce("/login")}),nn(),rn(),document.getElementById("btn-notif")?.addEventListener("click",s=>{s.preventDefault(),on()})}async function sn(){J("/login",({main:e})=>jt(e)),J("/dashboard",G(({main:e})=>Mt(e))),J("/calendar",G(({main:e})=>oa(e))),J("/employees",G(({main:e,params:i})=>qt(e,i))),J("/contracts",G(({main:e,params:i})=>dt(e,i))),J("/sp",G(({main:e})=>ga(e))),J("/mutasi",G(({main:e})=>ha(e))),J("/overtime",G(({main:e})=>ya(e))),J("/timeline",G(({main:e,params:i})=>Jt(e,i))),J("/os-migrations",G(({main:e})=>xa(e))),J("/issues",G(({main:e,params:i})=>Ut(e,i))),J("/one-on-one",G(({main:e,params:i})=>Gt(e,i))),J("/training",G(({main:e})=>Vt(e))),J("/relievers",G(({main:e,params:i})=>Yt(e,i))),J("/reports/inspection",G(({main:e})=>Wt(e))),J("/reports/cleaning",G(({main:e})=>Xt(e))),J("/reports/fogging",G(({main:e})=>Zt(e))),J("/reports/basecamp",G(({main:e})=>ea(e))),J("/reports/supply",G(({main:e})=>Tt(e,"supply"))),J("/sop",G(({main:e})=>ta(e))),J("/checklist",G(({main:e})=>aa(e))),J("/hygiene-standards",G(({main:e})=>na(e))),J("/dac",G(({main:e})=>wa(e))),J("/forms",G(({main:e})=>Tt(e))),J("/users",G(({main:e})=>ia(e))),J("/branches",G(({main:e})=>ra(e))),J("/profile",G(({main:e})=>sa(e))),J("/settings/import",G(({main:e})=>ma(e))),J("/audit-logs",G(({main:e})=>va(e)));let t=je();if(!t&&window.location.hash!=="#/login"&&Ce("/login"),t){let e=await T("/api/auth/me");if(e.ok){He(e.data.data),Sa();let i=ee(),a=window.location.hash;i&&i.role==="input_lembur"&&(a==="#/dashboard"||a===""||a==="#/")&&Ce("/overtime")}else qe(),Ce("/login")}window.addEventListener("fm:login",()=>{Sa();let e=ee();e&&e.role==="input_lembur"?Ce("/overtime"):Ce("/dashboard")}),Lt()}sn();
