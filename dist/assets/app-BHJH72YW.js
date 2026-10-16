var Ja=Object.defineProperty;var gt=(e,t)=>()=>(e&&(t=e(e=0)),t);var bt=(e,t)=>{for(var i in t)Ja(e,i,{get:t[i],enumerable:!0})};var Jt={};bt(Jt,{API:()=>Ht,CLIENT_SIDE_MAX_ROWS:()=>Te,IS_DEVELOPMENT:()=>ze,apiFetch:()=>L,clearToken:()=>je,getToken:()=>Oe,getUser:()=>oe,setToken:()=>ht,setUser:()=>Re});function Oe(){return localStorage.getItem("fm_token")}function ht(e){localStorage.setItem("fm_token",e)}function je(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function oe(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Re(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function L(e,t={}){let i=Oe(),n={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let o=`cb=${Date.now()}`,c=e.includes("?")?"&":"?",r=`${Ht}${e}${c}${o}`,l=await fetch(r,{...t,headers:n}),a;try{let s=await l.text();try{a=JSON.parse(s)}catch{a={error:`Server Error (${l.status}): ${s.substring(0,80)}...`}}}catch{a={error:"Gagal membaca respon dari server"}}return l.status===401&&(je(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:a}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var ze,Ga,Ht,Te,G=gt(()=>{ze=!1,Ga="https://fm-operations-api.facilitycare-audydental.workers.dev",Ht=Ga,Te=1e4});var vt={};bt(vt,{confirmDialog:()=>ie,createModal:()=>Z});function Z({title:e,content:t,onConfirm:i,onCancel:n,confirmText:o="Simpan",cancelText:c="Batal",size:r="md",confirmClass:l="btn-primary"}){let a={sm:"400px",md:"560px",lg:"720px",xl:"900px"},s=document.createElement("div");s.className="modal-overlay",s.innerHTML=`
    <div class="modal" style="max-width:${a[r]||a.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${c}</button>
        ${i?`<button class="btn ${l} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&s.querySelector(".modal-body").appendChild(t);let k=()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),250)};return s.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),k()}),s.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),k()}),i&&s.querySelector(".modal-confirm").addEventListener("click",()=>i(s,k)),s.addEventListener("click",w=>{w.target===s&&(n&&n(),k())}),document.body.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),{overlay:s,close:k}}function ie(e,t,i="Konfirmasi"){return Z({title:i,content:`<p>${e}</p>`,onConfirm:(n,o)=>{t(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ue=gt(()=>{});var be={};bt(be,{downloadExcel:()=>R,parseExcel:()=>Je,renderExcelButtons:()=>Va});function Je(e){return new Promise((t,i)=>{let n=new FileReader;n.onload=o=>{try{let c=new Uint8Array(o.target.result),r=XLSX.read(c,{type:"array"}),l=r.SheetNames[0],a=r.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let s=XLSX.utils.decode_range(a["!ref"]||"A1:A1"),k=s.e.r-s.s.r+1,w=s.e.c-s.s.c+1;console.log(`Total Rows (including empty): ${k}`),console.log(`Total Columns: ${w}`);let v=[];for(let C=s.s.c;C<=s.e.c;++C){let g=a[XLSX.utils.encode_cell({c:C,r:s.s.r})];g&&g.v&&v.push(g.v)}console.log(`Headers Found: ${v.join(", ")}`),console.log("---------------------------");let b=XLSX.utils.sheet_to_json(a,{defval:""});Object.defineProperty(b,"__worksheet",{value:a,enumerable:!1}),Object.defineProperty(b,"__headers",{value:v,enumerable:!1}),t(b)}catch(c){i(c)}},n.onerror=o=>i(o),n.readAsArrayBuffer(e)})}function R(e,t){try{let i=XLSX.utils.json_to_sheet(e),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,i,"Data"),XLSX.writeFile(n,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Va(e){return`
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
  `}var U=gt(()=>{});G();var ft={},Ye=null;function X(e,t){ft[e]=t}function _e(e){window.location.hash=e}function Gt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...n]=t.split("?"),o=ft[i];if(!o){for(let[r,l]of Object.entries(ft))if(r.endsWith("/*")&&i.startsWith(r.slice(0,-2))){o=l;break}}Ye&&(Ye(),Ye=null);let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let r=new URLSearchParams(n.join("?")),l=i.split("/").filter(Boolean),a=await o({path:i,params:r,segments:l,main:c});a&&(Ye=a)}else{let r=c||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var qe;function Ua(){return qe||(qe=document.createElement("div"),qe.id="toast-container",document.body.appendChild(qe)),qe}function yt(e,t="info",i=3500){let n=Ua(),o=document.createElement("div");o.className=`toast toast-${t}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${c[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,n.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},i)}var K=e=>yt(e,"success"),j=e=>yt(e,"error");var xt=e=>yt(e,"info");ue();G();G();ue();function Xe({columns:e,data:t,onEdit:i,onDelete:n,onView:o,actions:c=[],emptyText:r="Tidak ada data",bulkSelect:l=null}){let a=document.createElement("div");if(a.className="table-wrapper",!t||t.length===0)return a.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,a;let s=document.createElement("table");s.className="data-table";let k=document.createElement("thead"),w=document.createElement("tr");if(l){let b=document.createElement("th");b.style.width="40px",b.style.textAlign="center";let C=document.createElement("input");C.type="checkbox",C.id="select-all-checkbox",C.title="Pilih semua",C.addEventListener("change",()=>{t.forEach(g=>{C.checked?l.selectedIds.add(g.id):l.selectedIds.delete(g.id)}),a.querySelectorAll(".row-checkbox").forEach(g=>g.checked=C.checked),l.onToggle()}),b.appendChild(C),w.appendChild(b)}if(e.forEach(b=>{let C=document.createElement("th");C.textContent=b.label,b.width&&(C.style.width=b.width),w.appendChild(C)}),i||n||o||c.length>0){let b=document.createElement("th");b.textContent="Aksi",b.style.width="120px",w.appendChild(b)}k.appendChild(w),s.appendChild(k);let v=document.createElement("tbody");return t.forEach(b=>{let C=document.createElement("tr");if(l){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let h=document.createElement("input");h.type="checkbox",h.className="row-checkbox",h.checked=l.selectedIds.has(b.id),h.addEventListener("change",()=>{if(h.checked)l.selectedIds.add(b.id);else{l.selectedIds.delete(b.id);let d=document.getElementById("select-all-checkbox");d&&(d.checked=!1)}l.onToggle()}),g.appendChild(h),C.appendChild(g)}if(e.forEach(g=>{let h=document.createElement("td");if(g.render){let d=g.render(b[g.key],b);d instanceof HTMLElement?h.appendChild(d):h.innerHTML=d||""}else h.textContent=b[g.key]!==null&&b[g.key]!==void 0&&b[g.key]!==""?b[g.key]:"";g.nowrap&&(h.style.whiteSpace="nowrap"),C.appendChild(h)}),i||n||o||c.length>0){let g=document.createElement("td");g.className="actions-cell";let h=document.createElement("div");if(h.className="btn-group",o){let d=document.createElement("button");d.className="btn btn-xs btn-ghost",d.innerHTML="\u{1F441}",d.title="Lihat",d.addEventListener("click",()=>o(b)),h.appendChild(d)}if(i){let d=document.createElement("button");d.className="btn btn-xs btn-secondary",d.innerHTML="\u270F\uFE0F",d.title="Edit",d.addEventListener("click",()=>i(b)),h.appendChild(d)}c.forEach(d=>{let y=document.createElement("button");y.className=`btn btn-xs ${d.class||"btn-ghost"}`,y.innerHTML=d.icon||d.label,y.title=d.label,y.addEventListener("click",()=>d.handler(b)),h.appendChild(y)}),g.appendChild(h),C.appendChild(g)}v.appendChild(C)}),s.appendChild(v),a.appendChild(s),a}function Ze({page:e,pages:t,total:i,limit:n,onPage:o}){if(t<=1)return null;let c=document.createElement("div");c.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${i} data`,c.appendChild(r);let l=document.createElement("div");l.className="pagination-btns";let a=(w,v,b=!1,C=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${C?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=w,g.disabled=b,g.addEventListener("click",()=>o(v)),l.appendChild(g)};a("\xAB",1,e===1),a("\u2039",e-1,e===1);let s=Math.max(1,e-2),k=Math.min(t,e+2);for(let w=s;w<=k;w++)a(w,w,!1,w===e);return a("\u203A",e+1,e===t),a("\xBB",t,e===t),c.appendChild(l),c}ue();function He(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${He(t.fields)}</div>`;let i=t.required?"required":"",n=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(t.type){case"textarea":o=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let r=(t.options||[]).map(b=>{let C=typeof b=="object"?b.value:b,g=typeof b=="object"?b.label:b,h=t.value==C?"selected":"";return`<option value="${C}" ${h}>${g}</option>`}).join("");o=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${r}</select>`;break;case"combobox":let l=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,a=(t.options||[]).map(b=>{let C=typeof b=="object"?b.value:b,g=typeof b=="object"?b.label||b.value||"":b||"";return(g==="undefined"||g==="[object Object]"||g==="null")&&(g=""),g?`<option value="${g}"></option>`:""}).join(""),s=t.value||"";if(t.value){let b=(t.options||[]).find(C=>(typeof C=="object"?C.value:C)==t.value);if(b){let C=typeof b=="object"?b.label||b.value||"":b||"";C&&C!=="undefined"&&C!=="[object Object]"&&C!=="null"&&(s=C)}}o=`
          <input type="text" name="${t.name}" list="${l}" class="form-control" value="${s}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${l}">${a}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let k=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";o=`<input type="date" name="${t.name}" class="form-control" value="${k}" ${i}>`;break;case"month":let w=t.value?String(t.value).slice(0,7):"";o=`<input type="month" name="${t.name}" class="form-control" value="${w}" ${i}>`;break;case"number":let v=t.readonly?"readonly":"";o=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${v}>`;break;case"email":o=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":o=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:o=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let c=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${n}${o}${c}</div>`}).join("")}function et(e){let t={},i=new FormData(e);for(let[n,o]of i.entries())t[n]=o===""?null:o;return e.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(t[n.name]=null)}),t}function tt(e,t){t&&Object.entries(t).forEach(([i,n])=>{let o=e.querySelector(`[name="${i}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!n:o.type==="date"&&n&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(n):o.type==="month"&&n?o.value=String(n).slice(0,7):o.value=n??""))})}U();var Ce={},at={on(e,t){Ce[e]||(Ce[e]=new Set),Ce[e].add(t)},off(e,t){Ce[e]&&Ce[e].delete(t)},emit(e,t){Ce[e]&&Ce[e].forEach(i=>{try{i(t)}catch(n){console.warn("[calendarBus] Handler error:",n)}})},clear(){Object.keys(Ce).forEach(e=>delete Ce[e])}},Wa=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function he(e){if(!e){at.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");at.emit("data:changed",{module:t,relevant:Wa.has(t)})}function J({container:e,title:t,icon:i,apiPath:n,columns:o,formFields:c,filterFields:r,defaultFilters:l={},itemLabel:a="Data",canCreate:s=!0,canEdit:k=!0,canDelete:w=!0,onBeforeSubmit:v,onAfterLoad:b,onDataLoaded:C,extraActions:g=[],initialSearch:h="",exportOptions:d=null,bulkDelete:y=!1,paginationMode:m="server"}){let S=oe();S&&S.role==="viewer"&&(s=!1,k=!1,w=!1,y=!1,d=null),S&&S.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(s=!1,k=!1,w=!1,y=!1,d=null),S&&S.role==="input_lembur"&&(n==="/api/overtime"?(w=!1,y=!1,d=null):(s=!1,k=!1,w=!1,y=!1,d=null));let f=1,E={...l};h&&(E.search=h);let p=new Set;e.innerHTML=`
    ${y?`
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
        ${s?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
        ${d?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${d.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${d.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${d.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${d.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${r&&r.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${r.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${E.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${r.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${E[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${r.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${E[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function u(){let x=document.getElementById("bulk-toolbar");if(!x)return;let T=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),H=document.getElementById("btn-bulk-cancel");T.textContent=`${p.size} item dipilih`,p.size>0?(x.style.display="flex",A.disabled=!1,H.disabled=!1):(x.style.display="none",A.disabled=!0,H.disabled=!0);let q=document.getElementById("select-all-checkbox");if(q){let Q=document.querySelectorAll(".row-checkbox");if(Q.length>0){let z=[...Q].every(ne=>ne.checked),O=[...Q].some(ne=>ne.checked);q.checked=z,q.indeterminate=O&&!z}else q.checked=!1,q.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{p.clear(),document.querySelectorAll(".row-checkbox").forEach(T=>T.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),u()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(p.size===0)return;let x=[...p],T=document.createElement("div");T.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",T.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(T),T.querySelector("#bulk-cancel-btn").addEventListener("click",()=>T.remove()),T.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=T.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let H=await L(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});T.remove(),H.ok?(K(`${x.length} ${a} berhasil dihapus.`),p.clear(),u(),he(n),F()):j(H.data?.error||"Gagal menghapus data.")})});let $=document.getElementById("filter-search"),_;$?.addEventListener("input",x=>{clearTimeout(_),_=setTimeout(()=>{E.search=x.target.value,f=1,p.clear(),u(),F()},400)}),r?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",T=>{E[x.name]=T.target.value;let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value=T.target.value),f=1,p.clear(),u(),F()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",T=>{E[x.name]=T.target.value;let A=document.getElementById(`filter-${x.name}`);A&&(A.value=T.target.value),f=1,p.clear(),u(),F(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{E={},$&&($.value=""),r?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),f=1,p.clear(),u(),F()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{E={},$&&($.value=""),r?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),f=1,p.clear(),u(),F(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>B(null)),d&&document.addEventListener("click",function(x){let T=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");T&&A&&!A.contains(x.target)&&!T.contains(x.target)&&T.classList.remove("show-aksi-menu")});let P=document.getElementById("btn-mobile-filter"),N=document.getElementById("filter-options-wrapper"),M=document.getElementById("btn-close-filter-sheet");if(P&&N&&(P.addEventListener("click",x=>{x.preventDefault(),N.classList.add("sheet-open")}),M&&M.addEventListener("click",x=>{x.preventDefault(),N.classList.remove("sheet-open")})),d){document.getElementById(`btn-export-${d.moduleName}`)?.addEventListener("click",async T=>{let A=T.target,H=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await d.onExport()}catch{j("Gagal export data")}finally{A.innerHTML=H,A.disabled=!1}}),document.getElementById(`btn-template-${d.moduleName}`)?.addEventListener("click",()=>{d.onTemplate()});let x=document.getElementById(`input-import-${d.moduleName}`);x?.addEventListener("change",async T=>{let A=T.target.files[0];if(!A)return;let H=document.getElementById(`label-import-${d.moduleName}`),q=H?H.querySelector(".import-text"):null,Q=q?q.innerText:"";q&&(q.innerText="\u231B Memproses..."),H&&(H.style.pointerEvents="none"),x.disabled=!0;try{let z=await Je(A);if(z.length===0)throw new Error("File kosong atau format salah");await d.onImport(z),K("Import berhasil!"),he(n),F()}catch(z){j(z.message||"Gagal import data")}finally{q&&(q.innerText=Q),H&&(H.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function F(){u();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let T=m==="client",A=T?1:f,H=T?Te:20,q=new URLSearchParams({page:A,limit:H,...Object.fromEntries(Object.entries(E).filter(([,Y])=>Y))}),Q=await L(`${n}?${q}`);if(!Q.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Q.data?.error||"Error"}</p></div>`;return}let z=Q.data?.data||Q.data||[],O=Q.data?.pagination,ne=z.length;if(T){z=C(z);let Y=z.length,ae=20,ce=Math.ceil(Y/ae);f>ce&&ce>0&&(f=ce);let V=(f-1)*ae,ge=f*ae;z=z.slice(V,ge),O={page:f,limit:ae,total:Y,pages:ce}}!1,b&&b(z);let de=Xe({columns:o,data:z,onEdit:k?Y=>B(Y):null,actions:g.map(Y=>({...Y,handler:ae=>Y.handler(ae,F)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:y?{selectedIds:p,onToggle:u}:null});x.innerHTML="",x.appendChild(de);let me=document.getElementById("pagination-container");if(me&&(me.innerHTML="",O&&O.pages>1)){let Y=Ze({page:O.page,pages:O.pages,total:O.total,limit:O.limit,onPage:ae=>{f=ae,F()}});Y&&me.appendChild(Y)}}function D(x){let T=typeof c=="function"?c(x):c;return He(T)}function B(x){let T=!!x,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=D(x),T){let q=typeof c=="function"?c(x):c;tt(A,x)}let{close:H}=Z({title:T?`Edit ${a}`:`Tambah ${a}`,content:A,size:"lg",confirmText:T?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(q,Q)=>{if(!A.reportValidity())return;let z=q.querySelector(".modal-confirm");z.disabled=!0,z.textContent="Menyimpan...";let O=et(A),ne=typeof c=="function"?c(x):c,de=async ce=>{for(let V of ce)if(V.type==="row")await de(V.fields);else if(V.type==="combobox"&&O[V.name]){let ge=O[V.name],Ee=(V.options||[]).find(re=>{let pe=String(typeof re=="object"?re.value:re),mt=String(typeof re=="object"?re.label:re);return pe===ge||mt===ge});if(Ee)O[V.name]=typeof Ee=="object"?Ee.value:Ee;else if(V.createApi){let re={};re[V.createApi.field]=ge,V.createApi.extra&&Object.assign(re,V.createApi.extra);let pe=await L(V.createApi.path,{method:"POST",body:JSON.stringify(re)});if(pe.ok&&pe.data?.id)O[V.name]=pe.data.id;else if(pe.ok&&!pe.data?.id)O[V.name]=ge;else throw new Error(`Gagal membuat master data: ${pe.data?.error||"Unknown error"}`)}}};try{await de(ne)}catch(ce){j(ce.message),z.disabled=!1,z.textContent=T?"Simpan Perubahan":`Tambah ${a}`;return}v&&(O=await v(O,x));let me=T?"PUT":"POST",Y=T?`${n}/${x.id}`:n,ae=await L(Y,{method:me,body:JSON.stringify(O)});ae.ok?(K(T?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),Q(),he(n),F()):(j(ae.data?.error||"Gagal menyimpan data."),z.disabled=!1,z.textContent=T?"Simpan Perubahan":`Tambah ${a}`)}})}function I(x){ie(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let T=await L(`${n}/${x.id}`,{method:"DELETE"});T.ok?(K(`${a} berhasil dihapus.`),he(n),F()):j(T.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return F(),F}G();G();var Ae=null,nt=null;async function Fe(e=!1){if(Ae&&!e)return console.log("Employees Raw (Cache Hit)",Ae.slice(0,5)),Ae;let t=await L(`/api/employees?limit=${Te}&status=Aktif`);return Ae=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Ae.slice(0,5)),Ae}async function le(e=!1){let i=(await Fe(e)).map(n=>({value:n.label,label:n.label}));return console.log("Employee Options",i.slice(0,5)),i}async function W(e=!1){return nt&&!e||(nt=((await L("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),nt}function te(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function kt(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function De(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function wt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function ye(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}U();G();U();G();U();G();U();function it(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let o=new Date(t.value),c=new Date(i.value),r=Math.floor((c-o)/864e5);n.value=isNaN(r)?"":r}else n.value=""}}document.body.removeEventListener("input",it);document.body.addEventListener("input",it);document.body.removeEventListener("change",it);document.body.addEventListener("change",it);G();var $e={};function Pe(e){if($e[e]){try{$e[e].destroy()}catch{}delete $e[e]}}function Qa(){Object.keys($e).forEach(Pe)}var fe=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},Le=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function rt(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let n=Math.max(0,Math.round(fe(t))),o=e.textContent?e.textContent.replace(/[^\d]/g,""):"",c=o!==""&&parseInt(o)||0;if(n===c){e.textContent=n.toLocaleString("id-ID"),e._animFrame=null;return}let r=Date.now(),l=()=>{let a=Math.min((Date.now()-r)/i,1),s=1-Math.pow(1-a,3);e.textContent=Math.round(c+s*(n-c)).toLocaleString("id-ID"),a<1?e._animFrame=requestAnimationFrame(l):(e.textContent=n.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(l)}var Ya={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Xa=e=>{let t=Le(e,"\u2014");return`<span class="status-pill ${Ya[t]||"pill-neutral"}">${t}</span>`};var ve={family:"Inter",size:11},Se="#94A3B8",Be="#F1F5F9",St=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Za=()=>window.innerWidth<768;function ot(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Za()?"bottom":"top",labels:{font:ve,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ve,titleFont:{...ve,weight:"700"}}},scales:{x:{grid:{color:Be},ticks:{font:ve,color:Se,maxRotation:0}},y:{grid:{color:Be},ticks:{font:ve,color:Se},beginAtZero:!0}},...e}}var en=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),tn=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ut(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function xe(e,t,i=8e3){try{let n=new AbortController,o=setTimeout(()=>n.abort(),i),c=await L(e,{signal:n.signal}).catch(()=>null);if(clearTimeout(o),!c||!c.ok)return t;let r=c.data;return r?r.data!==void 0?r.data??t:r:t}catch{return t}}function an(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(n=>{let o=document.getElementById(n);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(n=>{let o=document.getElementById(n);if(o&&o.style.display==="none"){o.style.display="block";let c=o.parentElement;if(c&&!c.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",o.style.display="none",c.appendChild(r)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Qt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Yt({}),["table-contracts","table-issues"].forEach(n=>{let o=document.getElementById(n);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Wt(e){Qa(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${en()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${tn()}</div>

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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Vt(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let n=i.target.value,o=document.getElementById("jadwal-year-label");o&&(o.textContent=n);let c=document.getElementById("skel-jadwal"),r=document.getElementById("chart-jadwal");c&&(c.style.display="block",c.style.position="absolute"),r&&(r.style.display="none");let l=await xe(`/api/dashboard/schedule-chart?year=${n}`,{},8e3);try{Xt(l)}catch(a){console.warn("ScheduleChart render:",a),ke("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let n=i.target.value,o=n?`/api/dashboard/inspection-bar?month=${n}`:"/api/dashboard/inspection-bar",c=document.getElementById("skel-insp"),r=document.getElementById("chart-insp");c&&(c.style.display="block",c.style.position="absolute"),r&&(r.style.display="none");let l=await xe(o,{},8e3);try{Zt(l)}catch(a){console.warn("InspBar render:",a),ke("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>an(),5e3),await Vt(e)}async function Vt(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,n,o,c,r,l,a]=await Promise.all([xe("/api/dashboard/kpi",{},8e3),xe("/api/dashboard/issues-trend",{},8e3),xe("/api/dashboard/issues-summary",{},8e3),xe("/api/dashboard/stats",{},8e3),xe("/api/dashboard/calendar",[],8e3),xe("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),xe(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),xe("/api/dashboard/dropdown-aggregates",{},8e3)]),s=document.getElementById("filter-insp-month"),k=s?s.value:"",w=k?`/api/dashboard/inspection-bar?month=${k}`:"/api/dashboard/inspection-bar",v=await xe(w,{},8e3);window.dashboardAggregates=a||{};try{Qt(t)}catch(b){console.warn("KPI render:",b)}try{Yt(t)}catch(b){console.warn("MiniStats render:",b)}try{Xt(l)}catch(b){console.warn("ScheduleChart render:",b),ke("skel-jadwal","chart-jadwal")}try{nn(Array.isArray(n?.by_category)?n.by_category:[])}catch(b){console.warn("Donut render:",b),ke("skel-donut","chart-donut")}try{rn(i)}catch(b){console.warn("Trend render:",b),ke("skel-trend","chart-trend")}try{Zt(v)}catch(b){console.warn("InspBar render:",b),ke("skel-insp","chart-insp")}try{let b=Array.isArray(o)?o:Array.isArray(o?.recent_issues)?o.recent_issues:[];sn(b)}catch(b){console.warn("IssuesTable render:",b)}try{let b=Array.isArray(o?.expiring_contracts)?o.expiring_contracts:[];on(r)}catch(b){console.warn("ContractsTable render:",b)}try{ln(Array.isArray(c)?c:[])}catch(b){console.warn("Agenda render:",b)}try{dn()}catch(b){console.warn("Quick Actions render:",b)}}function Qt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,n=window.dashboardAggregates||{},o=a=>{let s=window.dashboardAggregates?.schedule_quarter_details?.[a];return s?`Insp: ${s.insp} \u2022 GC: ${s.gc} \u2022 DC: ${s.dc} \u2022 Fog: ${s.fog}`:""},c=a=>{let s=window.dashboardAggregates?.schedule_quarter_details?.[a];return s?`Rincian ${a}: Inspeksi: ${s.insp} | General Cleaning: ${s.gc} | Deep Cleaning: ${s.dc} | Fogging: ${s.fog}`:`Total Kegiatan ${a}`},r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:n.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:o(i),subId:"mini-jadwal-detail",tooltip:c(i)}];t.innerHTML=r.map(a=>{let s=a.val!==void 0?fe(a.val,0):fe(e[a.key]?.current,0);return a.dropdown?`
        <div class="kpi-card ${a.color}" ${a.id?`id="${a.id}"`:""} ${a.tooltip?`title="${a.tooltip}"`:""} data-href="${a.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${s}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${s}</div>
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
              <div class="kpi-value" data-target="${s}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${s}</div>
              ${a.dropdown?a.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" ${a.subId?`id="${a.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${a.tooltip||""}">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(a=>{a.addEventListener("click",s=>{s.target.closest("select")||a.dataset.href&&(window.location.hash=a.dataset.href)}),a.addEventListener("keydown",s=>{s.key==="Enter"&&!s.target.closest("select")&&a.dataset.href&&(window.location.hash=a.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(a=>{rt(a,parseInt(a.dataset.target)||0)});let l=document.getElementById("dash-jadwal-period");if(l){let a=s=>{let w=(window.dashboardAggregates?.schedule_by_quarter||{})[s]||0,v=document.querySelector("#mini-jadwal .kpi-value");v&&(v.dataset.target=w,rt(v,w,400));let b=document.getElementById("mini-jadwal-detail");b&&(b.textContent=o(s),b.title=c(s));let C=document.getElementById("mini-jadwal");if(C){let g=`#/timeline?dash_filter=period_${s.toLowerCase()}`;C.tagName==="A"?C.href=g:C.dataset.href=g,C.title=c(s)}};l.addEventListener("change",s=>a(s.target.value))}}function Yt(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),n=String(new Date().getMonth()+1).padStart(2,"0"),o=`${i}-${n}`,c=s=>`
    <select id="${s}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${i}-01" ${o===`${i}-01`?"selected":""}>Jan</option>
      <option value="${i}-02" ${o===`${i}-02`?"selected":""}>Feb</option>
      <option value="${i}-03" ${o===`${i}-03`?"selected":""}>Mar</option>
      <option value="${i}-04" ${o===`${i}-04`?"selected":""}>Apr</option>
      <option value="${i}-05" ${o===`${i}-05`?"selected":""}>Mei</option>
      <option value="${i}-06" ${o===`${i}-06`?"selected":""}>Jun</option>
      <option value="${i}-07" ${o===`${i}-07`?"selected":""}>Jul</option>
      <option value="${i}-08" ${o===`${i}-08`?"selected":""}>Agu</option>
      <option value="${i}-09" ${o===`${i}-09`?"selected":""}>Sep</option>
      <option value="${i}-10" ${o===`${i}-10`?"selected":""}>Okt</option>
      <option value="${i}-11" ${o===`${i}-11`?"selected":""}>Nov</option>
      <option value="${i}-12" ${o===`${i}-12`?"selected":""}>Des</option>
    </select>
  `,r=window.dashboardAggregates||{},l=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:c("dash-inspeksi-month"),val:r.inspeksi_by_month?.[o]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${o}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:c("dash-gcdc-month"),val:r.gcdc_by_month?.[o]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${o}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:c("dash-reliefer-month"),val:r.relievers_by_month?.[o]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${o}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:c("dash-fogging-month"),val:r.fogging_by_month?.[o]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${o}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=l.map(s=>s.dropdown?`
        <div class="mini-stat ${s.color}" style="cursor:pointer;" id="${s.id||""}" data-href="${s.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${s.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${fe(s.val)}">0</div>
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
            <div class="mini-stat-value" data-target="${fe(s.val)}">0</div>
            ${s.dropdown?s.dropdown:""}
          </div>
          <div class="mini-stat-text">${s.label}</div>
          ${s.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(s=>{s.addEventListener("click",k=>{k.target.closest("select")||s.dataset.href&&(window.location.hash=s.dataset.href)}),s.addEventListener("keydown",k=>{k.key==="Enter"&&!k.target.closest("select")&&s.dataset.href&&(window.location.hash=s.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(s=>rt(s,parseInt(s.dataset.target)||0,700));let a=(s,k,w,v)=>{let b=document.getElementById(s);if(b){let C=g=>{let d=(window.dashboardAggregates?.[w]||{})[g]||0,y=document.querySelector(`#${k} .mini-stat-value`);y&&(y.dataset.target=d,rt(y,d,400));let m=document.getElementById(k);if(m){let S=`${v}&month=${g}`;m.tagName==="A"?m.href=S:m.dataset.href=S}};b.addEventListener("change",g=>C(g.target.value))}};a("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),a("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),a("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),a("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function nn(e){ke("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Pe("donut");let n=(e||[]).filter(a=>fe(a.count)>0);if(!n.length){Ge(t,"Belum ada data permasalahan");return}let o=n.map(a=>`${Le(a.category,"Lainnya")}`),c=n.map(a=>fe(a.count)),r=c.reduce((a,s)=>a+s,0);i.innerHTML=n.map((a,s)=>{let k=St[s%St.length],w=r>0?Math.round(a.count/r*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${k}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${a.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${w}%)</span></div>
          <div class="donut-legend-label">${o[s]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(a){let s=a.width,k=a.height,w=a.ctx;w.restore();let v=(k/80).toFixed(2);w.font="bold "+v+"em Inter",w.textBaseline="middle",w.fillStyle="#1E293B";let b=r.toString(),C=Math.round((s-w.measureText(b).width)/2),g=k/2;w.fillText(b,C,g-4),w.font="600 "+(v*.35).toFixed(2)+"em Inter",w.fillStyle="#64748B";let h="Total",d=Math.round((s-w.measureText(h).width)/2);w.fillText(h,d,g+10),w.save()}};$e.donut=new Chart(t,{type:"doughnut",data:{labels:o,datasets:[{data:c,backgroundColor:St,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ve,titleFont:{...ve,weight:"700"},callbacks:{label:a=>` ${a.label}: ${a.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function rn(e){ke("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Pe("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],n=(e.labels||[]).map(r=>{if(!r||typeof r!="string")return"";try{let[l,a]=r.split("-");return(i[Number(a)-1]||a)+" "+String(l).slice(-2)}catch{return r}}),o=(e.open||[]).map(r=>fe(r)),c=(e.closed||[]).map(r=>fe(r));if(!n.length){Ge(t,"Belum ada data trend");return}$e.trend=new Chart(t,{type:"line",data:{labels:n,datasets:[{label:"Open",data:o,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:c,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:ot({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:45,autoSkip:!0}},y:{grid:{color:Be},ticks:{font:{family:"Inter",size:9},color:Se},beginAtZero:!0}}})})}function Xt(e){ke("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Pe("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(a=>Array.isArray(a)&&a.some(s=>s>0))){Ge(t,"Belum ada data jadwal");return}let o=e["Inspeksi Hygiene"]||Array(12).fill(0),c=e["General Cleaning"]||Array(12).fill(0),r=e["Deep Cleaning"]||Array(12).fill(0),l=e.Fogging||Array(12).fill(0);$e.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:o,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:c,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:r,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:ot({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Be},ticks:{font:{family:"Inter",size:9},color:Se},min:0}}})})}function Zt(e){ke("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Pe("inspBar"),e=e||{};let i=e.labels||[],n=(e.fc||[]).map(c=>fe(c)),o=(e.spv||[]).map(c=>fe(c));if(!i.length){Ge(t,"Belum ada data inspeksi");return}$e.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:ot({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Se,maxRotation:45,minRotation:45}},y:{grid:{color:Be},ticks:{font:ve,color:Se},min:0,max:100}}})})}function on(e){ke("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Pe("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},n=(e.labels||[]).map(r=>{let l=r.split("-")[1];return i[l]||r}),o=(e.data||[]).map(r=>fe(r));if(!n.length){Ge(t,"Belum ada data");return}let c=t.getContext("2d");$e.contractMiniBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Habis",data:o,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:ot({onClick:(r,l)=>{if(l&&l.length>0){let a=l[0].index,s=(e.labels||[])[a];s&&(window.location.hash="#/contracts?month_expiry="+s)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Se,maxRotation:0,autoSkip:!1}},y:{grid:{color:Be,borderDash:[4,4],drawBorder:!1},ticks:{font:ve,color:Se,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function sn(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(n=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${Xa(n.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Le(n.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Le(n.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function ln(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,n=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,c=(e||[]).filter(r=>(r.event_date||"").startsWith(n)).slice(0,10);if(!c.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${c.map(r=>{let l="#3B82F6",a="#EFF6FF",s="Agenda",k=(r.title||"").toLowerCase();return k.includes("inspeksi")?(l="#10B981",a="#ECFDF5",s="Inspeksi"):k.includes("cleaning")||k.includes("gcdc")?(l="#3B82F6",a="#EFF6FF",s="Cleaning"):k.includes("reliefer")?(l="#F59E0B",a="#FFFBEB",s="Reliefer"):k.includes("fogging")&&(l="#8B5CF6",a="#F5F3FF",s="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Le(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Le(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${a};color:${l}">${s}</div>
        </div>
      `}).join("")}
    </div>
  `}function dn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ke(e,t){let i=document.getElementById(e),n=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),n){n.style.display="block";let o=n.parentElement;if(o){let c=o.querySelector(".chart-empty");c&&c.remove()}}}function Ge(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=t,i.appendChild(o)}}G();async function ea(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),n=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),c=document.getElementById("login-password");o?.addEventListener("click",()=>{let r=c.type==="text";c.type=r?"password":"text",o.style.color=r?"":"var(--primary)"}),t?.addEventListener("submit",async r=>{r.preventDefault(),i.style.display="none";let l=t.username.value.trim(),a=t.password.value;if(!l||!a){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let s=await L("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:a})});s.ok&&s.data.success?(ht(s.data.data.token),Re(s.data.data.user),K("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=s.data.error||"Username atau password salah.",i.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}G();U();async function cn(){return await W()}function st(e){if(!e||e==="-"||String(e).trim()==="")return"-";let t=String(e).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],n=t.match(/^(\d{4})-(\d{1,2})/);if(n){let c=n[1],r=parseInt(n[2],10)-1;if(r>=0&&r<12)return`${i[r]} ${c}`}let o=t.split(/[\/\-\.]/);if(o.length===2&&o[1].length===4){let c=parseInt(o[0],10)-1;if(c>=0&&c<12)return`${i[c]} ${o[1]}`}if(o.length===3&&o[2].length===4){let c=parseInt(o[1],10)-1;if(c>=0&&c<12)return`${i[c]} ${o[2]}`}for(let c=0;c<12;c++)if(t.toLowerCase().includes(i[c].toLowerCase()))return t;return t}function pn(e,t){let i=String(e.status||"").toLowerCase();return t==="active"?i==="aktif":t==="reliefer"?e.division==="FC - RELIEFER"&&i==="aktif":!1}async function ta(e,t){let i=await cn(),n=t?t.get("dash_filter"):null;J({container:e,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>n?o.filter(c=>pn(c,n)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>De(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>te(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:o?.target_pindah_os?String(o.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:o?.target_selesai_os?String(o.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await L(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let c=o.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||"","Target Pindah OS":st(r.target_pindah_os)==="-"?"":st(r.target_pindah_os),"Target Selesai OS":st(r.target_selesai_os)==="-"?"":st(r.target_selesai_os)}));R(c,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async o=>{let c=a=>{if(!a)return null;let s=String(a||"").toLowerCase(),k=i.find(w=>String(w.label||"").toLowerCase()===s);return k?k.value:null},r=o.map(a=>({full_name:String(a["Nama Lengkap"]||"").trim(),branch_id:c(String(a.Cabang||"").trim()),division:String(a.Divisi||a["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(a["No. HP"]||a["No. Hp"]||"").trim(),join_date:String(a["Tgl Masuk"]||a["Tanggal Masuk"]||"").trim(),status:String(a.Status||"").trim(),target_pindah_os:String(a["Target Pindah OS"]||"").trim(),target_selesai_os:String(a["Target Selesai OS"]||"").trim(),notes:String(a.Catatan||"").trim()})).filter(a=>a.full_name),l=await L("/api/import/employees",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}G();U();var Ct=[],aa=[];async function un(){Ct=await W(),aa=await Fe()}var Et=async e=>{let t=[],i=1;for(;;){let o=await(await Promise.resolve().then(()=>(G(),Jt))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${i}`);if(!o.ok)break;let c=o.data?.data||o.data||[],r=Array.isArray(c)?c:[];if(t=t.concat(r),r.length<100||o.data?.pagination&&i>=o.data.pagination.pages)break;i++}return t};async function lt(e,t){await un(),J({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:n=>n,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:n=>De(n)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:n=>window.formatDate(n)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:n=>!n||String(n).startsWith("2099")?"Tetap / PKWTT":window.formatDate(n)},{key:"days_remaining",label:"Sisa Kontrak",render:(n,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':kt(n)},{key:"status",label:"Status",render:n=>te(n)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ct},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:n=>(n.end_date||(n.end_date="2099-12-31"),n),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let n=document.createElement("button");n.id="btn-find-missing",n.className="btn btn-ghost",n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.style.marginLeft="8px",n.style.color="#EF4444",n.style.border="1px solid currentColor",n.onclick=async()=>{n.innerHTML="\u231B Mencari...",n.disabled=!0;try{let[c,r]=await Promise.all([Et("/api/employees?status=Aktif"),Et("/api/contracts")]);if(c.length>0){let l=r.filter(w=>w.status==="Aktif"),a=new Set(l.map(w=>w.employee_id)),s=c.filter(w=>!a.has(w.id)),k=`<p style="margin-bottom:12px">Data yang terbaca: <b>${c.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${s.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;s.forEach(w=>{let v=r.filter(C=>C.employee_id===w.id),b='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(v.length>0){let C=v[0];b=`Pernah ada kontrak (Status: <b style="color:#EF4444">${C.status}</b>, Selesai: ${window.formatDate(C.end_date)})`}k+=`<li style="margin-bottom:8px"><b>${w.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${w.branch_name||"-"} | ${b}</span></li>`}),k+="</ul>",Promise.resolve().then(()=>(ue(),vt)).then(w=>w.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:k,cancelText:"Tutup"}))}}catch(c){console.error(c)}n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(n)}},formFields:n=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:aa,value:n?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:Ct,value:n?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:n?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:n?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:n?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:n?.end_date&&!String(n.end_date).startsWith("2099")?n.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:n?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:n?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let n=await L(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let o=n.data.data.map(c=>({"Nama Lengkap":c.employee_name,Cabang:c.branch_name||"","Div / Bagian":c.division||"","Tanggal Mulai":c.start_date||"","Tanggal Selesai":c.end_date&&String(c.end_date).startsWith("2099")?"":c.end_date||"","Sisa Kontrak":c.end_date&&String(c.end_date).startsWith("2099")?"Tetap":c.days_remaining!==null&&c.days_remaining!==void 0?`${c.days_remaining} Hari`:"",Status:c.status||""}));R(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async n=>{let[o,c]=await Promise.all([L("/api/branches?limit=10000"),Et("/api/employees")]),r=o.data?.data||[],l=c||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((h,d)=>{console.log(`${d+1}. ID: ${h.id}, Name: ${h.full_name}, Status: ${h.status}`)}));let a=h=>{if(!h)return null;let d=String(h||"").replace(/\s+/g," ").toLowerCase().trim(),y=r.find(m=>String(m.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===d||String(m.code||"").replace(/\s+/g," ").toLowerCase().trim()===d||String(m.name||"").replace(/\s+/g," ").toLowerCase().trim()===d);return y?y.id:null},s=(h,d)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${d}`),console.log(`Nama dari Excel : "${h}"`),!h)return console.log("Alasan gagal mapping : Nama kosong"),null;let y=String(h||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${y}"`),console.log(`Jumlah employee di database : ${l.length}`);let m=l.find(S=>String(S.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return m?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${m.id}`),m.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},k=h=>{if(!h)return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let d=String(h).trim();if(/^\d{4,5}(\.\d+)?$/.test(d)){let m=Math.floor(Number(d));if(m>2e4&&m<99999){let S=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let y=d.split(/[\/\-\.]/);if(y.length===3){let[m,S,f]=y.map(E=>E.trim());if(m.length===4&&S.length<=2&&f.length<=2)return`${m}-${S.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&S.length<=2&&m.length<=2)return`${f}-${S.padStart(2,"0")}-${m.padStart(2,"0")}`}return d},w=n.map((h,d)=>{let y=d+2,m=String(h["Nama Lengkap"]||"").trim(),S=h["Tanggal Mulai"],f=k(S);if(!f){let u=n.__worksheet,$=n.__headers||[],_=$.indexOf("Tanggal Mulai"),P="N/A",N="N/A",M="N/A";if(_!==-1&&u&&window.XLSX){let D=window.XLSX.utils.encode_cell({c:_,r:y-1});M=D;let B=u[D];B?(P=B.t||"undefined",N=B.w||"undefined"):P="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let F="Unknown";S==null||S===""?F="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":S instanceof Date&&isNaN(S.getTime())?F="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":F="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${y}`),console.log(`Employee Name : ${m}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${_})`),console.log(`Raw Cell Value : "${S}"`),console.log(`JavaScript Type : ${typeof S}`),console.log(`SheetJS Cell Type : ${P}`),console.log(`SheetJS Formatted Value : "${N}"`),console.log(`Value After Trim : "${String(S||"").trim()}"`),console.log(`Value After Date Parser : "${f}"`),console.log(`Is Empty : ${!S}`),console.log(`Is Invalid Date : ${S instanceof Date?isNaN(S.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${F}`),console.log(`Workbook Sheet : ${u?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${M}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(h,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify($)),console.log(`==========================
`)}let E=s(m,y),p=null;return E?f||(p="Tanggal Mulai kosong atau tidak berformat tanggal"):p="Karyawan tidak ditemukan di Database",{isValid:!!(E&&f),invalidReason:p,rowNum:y,data:{employee_id:E,branch_id:a(String(h.Cabang||"").trim()),division:String(h["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:f,end_date:k(h["Tanggal Selesai"])||"2099-12-31",status:String(h.Status||"").trim(),_rawName:m}}}),v=[],b=[];if(w.forEach(h=>{h.isValid?v.push(h.data):b.push({rowNum:h.rowNum,name:h.data._rawName,reason:h.invalidReason})}),console.log(`Split Validation - Valid: ${v.length}, Invalid: ${b.length}`),v.length===0){let h=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${n.length}
Valid: 0
Invalid: ${b.length}

Daftar Kegagalan (Contoh):
`;b.slice(0,10).forEach(d=>{h+=`- Row ${d.rowNum} | Nama: ${d.name} | Alasan: ${d.reason}
`}),b.length>10&&(h+=`- ... dan ${b.length-10} lainnya.
`),alert(h);return}let C=await L("/api/contracts/import",{method:"POST",body:JSON.stringify(v)}),g=`IMPORT SUMMARY
======================
`;g+=`Total Baris Excel : ${n.length}
`,g+=`Baris Valid       : ${v.length}
`,g+=`Baris Invalid     : ${b.length}

`,C&&C.data&&C.data.metrics?(g+=`Berhasil INSERT   : ${C.data.metrics.inserted}
`,g+=`Berhasil UPDATE   : ${C.data.metrics.updated}
`):g+=`Berhasil diproses : ${v.length}
`,b.length>0&&(g+=`
DAFTAR DATA DILEWATI:
`,b.forEach(h=>{g+=`- Row ${h.rowNum} | ${h.name} | ${h.reason}
`})),alert(g),typeof lt=="function"&&lt()}}})}G();U();var $t=[],Ue=[];function mn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(n=>n.period===i))return i;return"Q3"}async function na(e,t){$t=await W();let i=await le();Ue=["Berlin Ariansyah","Ade Surahman"];let n=h=>h&&!Ue.find(d=>String(typeof d=="object"?d.value:d).toLowerCase()===String(h).toLowerCase())?[...Ue,h]:Ue,o=await L(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),c=h=>{if(!h||h==="-"||String(h).trim()==="")return"";let d=String(h).split("-");return d.length===3&&d[0].length===4?`${d[2]}-${d[1]}-${d[0]}`:h},r=o.data?.data||[],l=mn(r),a=t?t.get("dash_filter"):null,s=new Date,k=`${s.getFullYear()}-${String(s.getMonth()+1).padStart(2,"0")}`,w={},v=t&&t.get("month")?t.get("month"):null;a==="inspeksi"?w={status:"Done",activity_type:"Inspeksi Hygiene",month:v}:a==="gcdc"?w={status:"Done",activity_type:"GCDC",month:v}:a&&a.startsWith("period_")&&(w={period:a.replace("period_","").toUpperCase()});let b=new Date().getFullYear(),g=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((h,d)=>{let y=String(d+1).padStart(2,"0");return{value:`${b}-${y}`,label:`${h} ${b}`}});J({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:w,onDataLoaded:h=>h.sort((d,y)=>{let m=d.opening_date?new Date(d.opening_date).getTime():0;return(y.opening_date?new Date(y.opening_date).getTime():0)-m}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:h=>wt(h)},{key:"period",label:"Periode",render:h=>ye(h)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:h=>c(h)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:h=>c(h)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:h=>c(h)},{key:"status",label:"Status",render:h=>te(h)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:$t},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:g},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ue}],formFields:h=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:$t,value:h?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:h?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:h?.period},{name:"pic",label:"PIC",type:"combobox",options:n(h?.pic),value:h?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:h?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:h?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:h?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:h?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:h?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let h=await L(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(h.ok){let d=h.data.data.map(y=>({Cabang:y.branch_name||"",Kegiatan:y.activity_type||"",Periode:y.period||"",PIC:y.pic||"","Tgl Opening":y.opening_date||"","Tgl Target":y.target_date||"","Tgl Selesai":y.completion_date||"",Status:y.status||""}));R(d,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async h=>{let y=(await L("/api/branches?all=1")).data?.data||[],m=p=>{if(!p)return null;let u=String(p||"").toLowerCase().trim(),$=u.match(/^(\d{3})/);if($){let P=y.find(N=>String(N.code||"")===$[1]);if(P)return P.id}let _=y.find(P=>String(P.full_name||"").toLowerCase()===u||String(P.code||"").toLowerCase()===u||String(P.name||"").toLowerCase()===u);return _?_.id:null},S=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let _=Number(u);if(_>2e4&&_<99999){let P=new Date(Date.UTC(1899,11,30)+_*864e5);return isNaN(P.getTime())?"":P.toISOString().slice(0,10)}}let $=u.split(/[\/\-\.]/);if($.length===3){let[_,P,N]=$.map(M=>M.trim());if(_.length===4&&P.length<=2&&N.length<=2)return`${_}-${P.padStart(2,"0")}-${N.padStart(2,"0")}`;if(N.length===4&&P.length<=2&&_.length<=2)return`${N}-${P.padStart(2,"0")}-${_.padStart(2,"0")}`}return u},f=h.map(p=>({branch_id:m(String(p.Cabang||"").trim()),branch_name:String(p.Cabang||"").trim(),activity_type:String(p.Kegiatan||"").trim(),period:String(p.Periode||"").trim(),pic:String(p.PIC||p.Pic||"").trim(),opening_date:S(p["Tgl Opening"]||p["Tanggal Opening"]||p["Tgl Openir"]),target_date:S(p["Tgl Target"]||p["Tanggal Target"]),completion_date:S(p["Tgl Selesai"]||p["Tanggal Selesai"]),status:String(p.Status||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.activity_type&&p.period),E=await L("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:f,onDuplicate:"update"})});if(!E.ok)throw new Error(E.data?.error||"Import gagal");return he("schedule"),E.data}}})}G();U();var _t=[],dt=[];function ct(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let o=new Date(t.value),c=new Date(i.value),r=Math.floor((c-o)/864e5);n.value=isNaN(r)?"":r}else n.value=""}}document.body.removeEventListener("input",ct);document.body.addEventListener("input",ct);document.body.removeEventListener("change",ct);document.body.addEventListener("change",ct);function gn(e,t){let i=String(e.status||"").toLowerCase();return t==="open"?i==="open":!1}async function ia(e,t){let i=t?t.get("dash_filter"):null;_t=await W(),dt=await le();let n=a=>a&&!dt.find(s=>s.value===a)?[...dt,{value:a,label:a}]:dt,o=new Date().getFullYear(),c=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((a,s)=>{let k=String(s+1).padStart(2,"0");return{value:`${o}-${k}`,label:`${a} ${o}`}});J({container:e,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:a=>i?a.filter(s=>gn(s,i)):a,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:a=>window.formatDate(a)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:a=>`<span class="badge badge-secondary">${a}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:a=>`<span title="${a}">${a?.length>50?a.slice(0,50)+"...":a}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:a=>`<span title="${a||""}">${a?.length>40?a.slice(0,40)+"...":a||"-"}</span>`},{key:"status",label:"Status",render:a=>te(a)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:a=>window.formatDate(a)},{key:"day_count",label:"Day",render:a=>a??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:c}],formFields:a=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:a?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:_t,value:a?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:a?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:a?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:a?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:n(a?.employee_name),value:a?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:n(a?.fc_specialist),value:a?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:a?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:a?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:a?.completion_date},{name:"day_count",label:"Day",type:"number",value:a?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let a=await L(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let s=a.data.data.map(k=>({"Tanggal Info":k.report_date||"",Cabang:k.branch_name||"",Kategori:k.category||"","Sumber Laporan":k.source||"",Keluhan:k.complaint||"","Nama FC":k.employee_name||"","FC Spesialis":k.fc_specialist||"",Solusi:k.solution||"",Status:k.status||"","Tanggal Selesai":k.completion_date||"",Day:k.day_count!==null?k.day_count:""}));R(s,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async a=>{let k=(await L("/api/branches?all=1")).data?.data||[],w=C=>{if(!C)return null;let g=String(C||"").toLowerCase(),h=k.find(d=>String(d.full_name||"").toLowerCase()===g||String(d.code||"").toLowerCase()===g||String(d.name||"").toLowerCase()===g);return h?h.id:null},v=a.map(C=>({branch_id:w(String(C.Cabang||"").trim()),report_date:String(C["Tanggal Info"]||C.Tanggal||"").trim(),category:String(C.Kategori||"").trim(),source:String(C["Sumber Laporan"]||C.Sumber||"").trim(),complaint:String(C.Keluhan||"").trim(),employee_name:String(C["Nama FC"]||"").trim(),fc_specialist:String(C["FC Spesialis"]||"").trim(),solution:String(C.Solusi||"").trim(),completion_date:String(C["Tanggal Selesai"]||C["Tgl Selesai"]||"").trim(),status:String(C.Status||"").trim(),day_count:C.Day||C.Hari||null})).filter(C=>C.report_date&&C.complaint&&C.category),b=await L("/api/import/issues",{method:"POST",body:JSON.stringify({rows:v,onDuplicate:"update"})});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}G();var Ie=[];function bn(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function ra(e,t){let i=t?t.get("dash_filter"):null;Ie=await W();let n=await le(),o=["Ade","Berlin"],c=l=>l&&!n.find(a=>a.value===l)?[...n,{value:l,label:l}]:n,r=l=>l&&!o.find(a=>(typeof a=="object"?a.value:a)===l)?[...o,l]:o;J({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>i?l.filter(a=>bn(a,i)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>te(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Ie},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let a=new URLSearchParams(l||{}).toString(),s=await L(`/api/one-on-one?limit=10000&${a}`);if(s.ok){let k=s.data.data.map(v=>({Tanggal:v.meeting_date||"",Cabang:v.branch_name||"","Nama Karyawan":v.employee_name||"",PIC:v.pic||"",Masalah:v.problem||"",Solusi:v.solution||"",Status:v.status||"","Tgl Selesai":v.completion_date||"",Dokumen:v.document_link||""})),{downloadExcel:w}=await Promise.resolve().then(()=>(U(),be));w(k,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:a}=await Promise.resolve().then(()=>(U(),be));a(l,"Template_Import_OneOnOne")},onImport:async l=>{let a=v=>{if(!v)return null;let b=String(v||"").toLowerCase(),C=Ie.find(g=>String(g.label||"").toLowerCase()===b);return C?C.value:null},s=v=>{if(!v)return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let b=String(v).trim();if(/^\d{4,5}$/.test(b)){let g=Number(b);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let C=b.split(/[\/\-\.]/);if(C.length===3){let[g,h,d]=C.map(y=>y.trim());if(g.length===4&&h.length<=2&&d.length<=2)return`${g}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&h.length<=2&&g.length<=2)return`${d}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return b},k=l.map(v=>({meeting_date:s(v.Tanggal),employee_name:String(v["Nama Karyawan"]||"").trim(),branch_id:a(String(v.Cabang||"").trim()),pic:String(v.PIC||"").trim(),problem:String(v.Masalah||"").trim(),solution:String(v.Solusi||"").trim(),status:String(v.Status||"").trim(),completion_date:s(v["Tgl Selesai"]),document_link:String(v.Dokumen||"").trim()})).filter(v=>v.meeting_date&&v.employee_name&&v.branch_id),w=await L("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:k,onDuplicate:"update"})});if(!w.ok)throw new Error(w.data?.error||"Import gagal");return w.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Ie.find(a=>a.value==l.branch_id)?[...Ie,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Ie,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:c(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:r(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}G();async function oa(e){let t=await W(),i=await le(),n=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],o=a=>a&&!i.find(s=>s.value===a)?[...i,{value:a,label:a}]:i,c=a=>a&&!n.find(s=>(typeof s=="object"?s.value:s)===a)?[...n,a]:n,r=Array.from({length:5},(a,s)=>String(new Date().getFullYear()-s));J({container:e,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:a=>{try{let s=JSON.parse(a);return Array.isArray(s)?s.join(", "):a||"-"}catch{return a||"-"}}},{key:"score",label:"Nilai",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),k=await L(`/api/training?limit=10000&${s}`);if(k.ok){let w=k.data.data.map(b=>{let C=b.participants||"";try{let g=JSON.parse(C);C=Array.isArray(g)?g.join(", "):C}catch{}return{Tanggal:b.training_date||"",Batch:b.batch||"",Materi:b.subject||"",Cabang:b.branch_name||"",Trainer:b.trainer||"",Peserta:C,Nilai:b.score!==null&&b.score!==void 0?b.score:"",Dokumen:b.document_link||""}}),{downloadExcel:v}=await Promise.resolve().then(()=>(U(),be));v(w,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let a=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(U(),be));s(a,"Template_Import_Training")},onImport:async a=>{let s=b=>{if(!b)return null;let C=String(b||"").toLowerCase(),g=t.find(h=>String(h.label||"").toLowerCase()===C);return g?g.value:null},k=b=>{if(!b)return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let C=String(b).trim();if(/^\d{4,5}$/.test(C)){let h=Number(C);if(h>2e4&&h<99999){let d=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);let g=C.split(/[\/\-\.]/);if(g.length===3){let[h,d,y]=g.map(m=>m.trim());if(h.length===4&&d.length<=2&&y.length<=2)return`${h}-${d.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&d.length<=2&&h.length<=2)return`${y}-${d.padStart(2,"0")}-${h.padStart(2,"0")}`}return C},w=a.map(b=>({training_date:k(b.Tanggal),batch:String(b.Batch||"").trim(),subject:String(b.Materi||"").trim(),branch_id:s(String(b.Cabang||"").trim()),trainer:String(b.Trainer||"").trim(),participants:String(b.Peserta||"").trim(),score:b.Nilai?Number(b.Nilai):null,document_link:String(b.Dokumen||"").trim()})).filter(b=>b.training_date&&b.subject&&b.branch_id),v=await L("/api/training/import",{method:"POST",body:JSON.stringify(w)});if(!v.ok)throw new Error(v.data?.error||"Import gagal")}},formFields:a=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:a?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:a?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:a?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,value:a?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:c(a?.trainer),value:a?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let s=JSON.parse(a?.participants);return Array.isArray(s)?s.join(", "):a?.participants||""}catch{return a?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:a?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],onBeforeSubmit:async a=>a})}G();ue();U();function sa({container:e,title:t,icon:i,apiPath:n,columns:o,formFields:c,filterFields:r,defaultFilters:l={},itemLabel:a="Data",canCreate:s=!0,canEdit:k=!0,canDelete:w=!0,onBeforeSubmit:v,onAfterLoad:b,onDataLoaded:C,extraActions:g=[],initialSearch:h="",exportOptions:d=null,bulkDelete:y=!1,paginationMode:m="server"}){let S=oe();S&&S.role==="viewer"&&(s=!1,k=!1,w=!1,y=!1,d=null),S&&S.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(s=!1,k=!1,w=!1,y=!1,d=null),S&&S.role==="input_lembur"&&(n==="/api/overtime"?(w=!1,y=!1,d=null):(s=!1,k=!1,w=!1,y=!1,d=null));let f=1,E={...l};h&&(E.search=h);let p=new Set;e.innerHTML=`
    ${y?`
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
        ${s?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
        ${d?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${d.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${d.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${d.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${d.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${r&&r.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${r.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${E.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${r.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${E[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${r.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let T=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${T}</option>${(x.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${E[x.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function u(){let x=document.getElementById("bulk-toolbar");if(!x)return;let T=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),H=document.getElementById("btn-bulk-cancel");T.textContent=`${p.size} item dipilih`,p.size>0?(x.style.display="flex",A.disabled=!1,H.disabled=!1):(x.style.display="none",A.disabled=!0,H.disabled=!0);let q=document.getElementById("select-all-checkbox");if(q){let Q=document.querySelectorAll(".row-checkbox");if(Q.length>0){let z=[...Q].every(ne=>ne.checked),O=[...Q].some(ne=>ne.checked);q.checked=z,q.indeterminate=O&&!z}else q.checked=!1,q.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{p.clear(),document.querySelectorAll(".row-checkbox").forEach(T=>T.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),u()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(p.size===0)return;let x=[...p],T=document.createElement("div");T.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",T.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(T),T.querySelector("#bulk-cancel-btn").addEventListener("click",()=>T.remove()),T.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=T.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let H=await L(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});T.remove(),H.ok?(K(`${x.length} ${a} berhasil dihapus.`),p.clear(),u(),he(n),F()):j(H.data?.error||"Gagal menghapus data.")})});let $=document.getElementById("filter-search"),_;$?.addEventListener("input",x=>{clearTimeout(_),_=setTimeout(()=>{E.search=x.target.value,f=1,p.clear(),u(),F()},400)}),r?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",T=>{E[x.name]=T.target.value;let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value=T.target.value),f=1,p.clear(),u(),F()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",T=>{E[x.name]=T.target.value;let A=document.getElementById(`filter-${x.name}`);A&&(A.value=T.target.value),f=1,p.clear(),u(),F(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{E={},$&&($.value=""),r?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),f=1,p.clear(),u(),F()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{E={},$&&($.value=""),r?.forEach(x=>{let T=document.getElementById(`filter-${x.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${x.name}`);A&&(A.value="")}),f=1,p.clear(),u(),F(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>B(null)),d&&document.addEventListener("click",function(x){let T=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");T&&A&&!A.contains(x.target)&&!T.contains(x.target)&&T.classList.remove("show-aksi-menu")});let P=document.getElementById("btn-mobile-filter"),N=document.getElementById("filter-options-wrapper"),M=document.getElementById("btn-close-filter-sheet");if(P&&N&&(P.addEventListener("click",x=>{x.preventDefault(),N.classList.add("sheet-open")}),M&&M.addEventListener("click",x=>{x.preventDefault(),N.classList.remove("sheet-open")})),d){document.getElementById(`btn-export-${d.moduleName}`)?.addEventListener("click",async T=>{let A=T.target,H=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await d.onExport()}catch{j("Gagal export data")}finally{A.innerHTML=H,A.disabled=!1}}),document.getElementById(`btn-template-${d.moduleName}`)?.addEventListener("click",()=>{d.onTemplate()});let x=document.getElementById(`input-import-${d.moduleName}`);x?.addEventListener("change",async T=>{let A=T.target.files[0];if(!A)return;let H=document.getElementById(`label-import-${d.moduleName}`),q=H?H.querySelector(".import-text"):null,Q=q?q.innerText:"";q&&(q.innerText="\u231B Memproses..."),H&&(H.style.pointerEvents="none"),x.disabled=!0;try{let z=await Je(A);if(z.length===0)throw new Error("File kosong atau format salah");await d.onImport(z),K("Import berhasil!"),he(n),F()}catch(z){j(z.message||"Gagal import data")}finally{q&&(q.innerText=Q),H&&(H.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function F(){u();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let T=m==="client",A=T?1:f,H=T?Te:20,q=new URLSearchParams({page:A,limit:H,...Object.fromEntries(Object.entries(E).filter(([,Y])=>Y))}),Q=await L(`${n}?${q}`);if(!Q.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Q.data?.error||"Error"}</p></div>`;return}let z=Q.data?.data||Q.data||[],O=Q.data?.pagination,ne=z.length;if(T){z=C(z);let Y=z.length,ae=20,ce=Math.ceil(Y/ae);f>ce&&ce>0&&(f=ce);let V=(f-1)*ae,ge=f*ae;z=z.slice(V,ge),O={page:f,limit:ae,total:Y,pages:ce}}!1,b&&b(z);let de=Xe({columns:o,data:z,onEdit:k?Y=>B(Y):null,actions:g.map(Y=>({...Y,handler:ae=>Y.handler(ae,F)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:y?{selectedIds:p,onToggle:u}:null});x.innerHTML="",x.appendChild(de);let me=document.getElementById("pagination-container");if(me&&(me.innerHTML="",O&&O.pages>1)){let Y=Ze({page:O.page,pages:O.pages,total:O.total,limit:O.limit,onPage:ae=>{f=ae,F()}});Y&&me.appendChild(Y)}}function D(x){let T=typeof c=="function"?c(x):c;return He(T)}function B(x){let T=!!x,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=D(x),T){let q=typeof c=="function"?c(x):c;tt(A,x)}let{close:H}=Z({title:T?`Edit ${a}`:`Tambah ${a}`,content:A,size:"lg",confirmText:T?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(q,Q)=>{if(!A.reportValidity())return;let z=q.querySelector(".modal-confirm");z.disabled=!0,z.textContent="Menyimpan...";let O=et(A),ne=typeof c=="function"?c(x):c,de=async ce=>{for(let V of ce)if(V.type==="row")await de(V.fields);else if(V.type==="combobox"&&O[V.name]){let ge=O[V.name],Ee=(V.options||[]).find(re=>{let pe=String(typeof re=="object"?re.value:re),mt=String(typeof re=="object"?re.label:re);return pe===ge||mt===ge});if(Ee)O[V.name]=typeof Ee=="object"?Ee.value:Ee;else if(V.createApi){let re={};re[V.createApi.field]=ge,V.createApi.extra&&Object.assign(re,V.createApi.extra);let pe=await L(V.createApi.path,{method:"POST",body:JSON.stringify(re)});if(pe.ok&&pe.data?.id)O[V.name]=pe.data.id;else if(pe.ok&&!pe.data?.id)O[V.name]=ge;else throw new Error(`Gagal membuat master data: ${pe.data?.error||"Unknown error"}`)}}};try{await de(ne)}catch(ce){j(ce.message),z.disabled=!1,z.textContent=T?"Simpan Perubahan":`Tambah ${a}`;return}v&&(O=await v(O,x));let me=T?"PUT":"POST",Y=T?`${n}/${x.id}`:n,ae=await L(Y,{method:me,body:JSON.stringify(O)});ae.ok?(K(T?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),Q(),he(n),F()):(j(ae.data?.error||"Gagal menyimpan data."),z.disabled=!1,z.textContent=T?"Simpan Perubahan":`Tambah ${a}`)}})}function I(x){ie(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let T=await L(`${n}/${x.id}`,{method:"DELETE"});T.ok?(K(`${a} berhasil dihapus.`),he(n),F()):j(T.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return F(),F}G();U();async function la(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await W(),n=await le(),o=t?t.get("dash_filter"):null,c={};if(o==="reliever"){let v=new Date,b=`${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}`;c={status:"Done",month:t&&t.get("month")?t.get("month"):b}}console.log("RAW",await Fe()),console.log("OPTIONS",n);let r=v=>v&&!n.find(b=>b.value===v)?[...n,{value:v,label:v}]:n,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],a=v=>v&&!l.includes(v)?[...l,v]:l,s=new Date().getFullYear(),w=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((v,b)=>{let C=String(b+1).padStart(2,"0");return{value:`${s}-${C}`,label:`${v} ${s}`}});sa({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:c,onDataLoaded:v=>v.sort((b,C)=>{let g=b.backup_date?new Date(b.backup_date).getTime():0;return(C.backup_date?new Date(C.backup_date).getTime():0)-g}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:v=>ye(v)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:v=>window.formatDate(v)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:v=>window.formatDate(v)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:v=>v?`<span class="badge badge-info">${v}</span>`:"-"},{key:"status",label:"Status",render:v=>te(v)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:w},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:v=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:v?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:v?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:r(v?.original_fc_name),value:v?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:a(v?.reliever_name),value:v?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:v?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:v?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:v?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:v?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:v?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let v=await L(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(v.ok){let b=v.data.data.map(C=>({Cabang:C.branch_name||"","Nama Facility care":C.original_fc_name||"",Periode:C.period||"",Relifer:C.reliever_name||"","Tanggal Back Up":C.backup_date||"","Tanggal Selesai":C.completion_date||"",Keterangan:C.reason||"",Shift:C.shift||"",Status:C.status||""}));b.length===0&&b.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),R(b,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async v=>{let C=(await L("/api/branches?all=1")).data?.data||[],g=y=>{if(!y)return null;let m=String(y||"").toLowerCase(),S=C.find(f=>String(f.full_name||"").toLowerCase()===m||String(f.code||"").toLowerCase()===m||String(f.name||"").toLowerCase()===m);return S?S.id:null},h=v.map(y=>({branch_name:String(y.Cabang||"").trim(),backup_date:String(y["Tanggal Back Up"]||y["Tanggal Backup"]||"").trim(),original_fc_name:String(y["Nama Facility care"]||y["FC Digantikan"]||"").trim(),reliever_name:String(y.Relifer||y.Reliefer||"").trim(),period:String(y.Periode||"").trim(),reason:String(y.Keterangan||"").trim(),shift:String(y.Shift||"").trim(),completion_date:String(y["Tanggal Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.reliever_name&&y.backup_date),d=await L("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}G();U();async function da(e){let t=await W(),i=Array.from({length:4},(n,o)=>String(new Date().getFullYear()-o));J({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:n=>ye(n)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"fc_score",label:"Point FC",render:n=>n!=null?`<strong class="${n>=80?"text-success":n>=60?"text-warning":"text-danger"}">${n}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"status",label:"Status",render:n=>te(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:n?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:n?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:n?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async n=>{let o=new URLSearchParams(n||{}).toString(),c=await L(`/api/reports/inspection?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));R(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async n=>{let o=a=>{if(!a)return null;let s=String(a||"").toLowerCase(),k=t.find(w=>String(w.label||"").toLowerCase()===s);return k?k.value:null},c=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let s=String(a).trim();if(s===""||s==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);if(/^\d{4,5}$/.test(s)){let w=Number(s);if(w>2e4&&w<99999){let v=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}let k=s.split(/[\/\-\.]/);if(k.length===3){let[w,v,b]=k.map(C=>C.trim());if(w.length===4&&v.length<=2&&b.length<=2)return`${w}-${v.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&v.length<=2&&w.length<=2)return`${b}-${v.padStart(2,"0")}-${w.padStart(2,"0")}`}return s},r=n.map(a=>({branch_id:o(String(a.Cabang||"").trim()),period:String(a.Periode||"").trim(),inspection_date:c(a.Tanggal),fc_score:a["Point FC"]!==void 0&&a["Point FC"]!==""?Number(a["Point FC"]):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==""?Number(a["Point SPV"]):null,status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.period&&a.inspection_date),l=await L("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}G();U();async function ca(e){let t=await W(),i=Array.from({length:4},(n,o)=>String(new Date().getFullYear()-o));J({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge ${n==="Deep Cleaning"?"badge-purple":"badge-success"}">${n}</span>`},{key:"period",label:"Periode",render:n=>ye(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>te(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:n?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:n?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:n?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async n=>{let o=new URLSearchParams(n||{}).toString(),c=await L(`/api/reports/cleaning?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));R(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async n=>{let o=a=>{if(!a)return null;let s=String(a||"").toLowerCase(),k=t.find(w=>String(w.label||"").toLowerCase()===s);return k?k.value:null},c=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let s=String(a).trim();if(s===""||s==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);if(/^\d{4,5}$/.test(s)){let w=Number(s);if(w>2e4&&w<99999){let v=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}let k=s.split(/[\/\-\.]/);if(k.length===3){let[w,v,b]=k.map(C=>C.trim());if(w.length===4&&v.length<=2&&b.length<=2)return`${w}-${v.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&v.length<=2&&w.length<=2)return`${b}-${v.padStart(2,"0")}-${w.padStart(2,"0")}`}return s},r=n.map(a=>({branch_id:o(String(a.Cabang||"").trim()),activity_type:String(a.Jenis||a.Kegiatan||"").trim(),period:String(a.Periode||"").trim(),activity_date:c(a.Tanggal),status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.activity_type&&a.period&&a.activity_date),l=await L("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}G();U();async function pa(e,t){let i=await W(),n=Array.from({length:4},(r,l)=>String(new Date().getFullYear()-l)),o=t?t.get("dash_filter"):null,c={};if(o==="fogging"){let r=new Date,l=String(r.getMonth()+1).padStart(2,"0"),a=String(r.getFullYear()),s=t?t.get("month"):null;s&&s.length===7&&(a=s.split("-")[0],l=s.split("-")[1]),c={status:"Done",month:l,year:a}}J({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:c,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:r=>`<span class="badge badge-warning">${r}</span>`},{key:"period",label:"Periode",render:r=>ye(r)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>te(r)},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:r=>r||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:r?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:r?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:r?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:r?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:r?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async r=>{let l=new URLSearchParams(r||{}).toString(),a=await L(`/api/reports/fogging?limit=10000&${l}`);if(a.ok){let s=a.data.data.map(k=>({Cabang:k.branch_name||"",Jenis:k.activity_type||"Fogging",Periode:k.period||"",Tanggal:k.activity_date||"",Status:k.status||"","Link Dokumen":k.document_link||""}));R(s,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async r=>{let l=w=>{if(!w)return null;let v=String(w||"").toLowerCase(),b=i.find(C=>String(C.label||"").toLowerCase()===v);return b?b.value:null},a=w=>{if(w==null||w==="")return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let v=String(w).trim();if(v===""||v==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(v))return v.slice(0,10);if(/^\d{4,5}$/.test(v)){let C=Number(v);if(C>2e4&&C<99999){let g=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let b=v.split(/[\/\-\.]/);if(b.length===3){let[C,g,h]=b.map(d=>d.trim());if(C.length===4&&g.length<=2&&h.length<=2)return`${C}-${g.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&g.length<=2&&C.length<=2)return`${h}-${g.padStart(2,"0")}-${C.padStart(2,"0")}`}return v},s=r.map(w=>({branch_id:l(String(w.Cabang||"").trim()),activity_type:String(w.Jenis||w.Kegiatan||"Fogging").trim(),period:String(w.Periode||"").trim(),activity_date:a(w.Tanggal),status:String(w.Status||"").trim(),document_link:String(w["Link Dokumen"]||"").trim(),notes:String(w.Catatan||w.Keterangan||"").trim()})).filter(w=>w.branch_id&&w.period&&w.activity_date),k=await L("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(s)});if(!k.ok)throw new Error(k.data?.error||"Import gagal");return k.data}}})}G();U();async function ua(e){let t=await W(),i=await le(),n=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],o=Array.from({length:4},(l,a)=>String(new Date().getFullYear()-a)),c=l=>l&&!i.find(a=>a.value===l)?[...i,{value:l,label:l}]:i,r=l=>l&&!n.find(a=>a.value===l)?[...n,{value:l,label:l}]:n;J({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>te(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:r(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let a=new URLSearchParams(l||{}).toString(),s=await L(`/api/reports/basecamp?limit=10000&${a}`);if(s.ok){let k=s.data.data.map(w=>({"Tgl Info":w.info_date||"",Cabang:w.branch_name||"",Permasalahan:w.problem||"",PIC:w.pic||"","Tgl Done":w.done_date||"",Status:w.status||"",Keterangan:w.notes||""}));R(k,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let a=v=>{if(!v)return null;let b=String(v||"").toLowerCase(),C=t.find(g=>String(g.label||"").toLowerCase()===b);return C?C.value:null},s=v=>{if(v==null||v==="")return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let b=String(v).trim();if(b===""||b==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);if(/^\d{4,5}$/.test(b)){let g=Number(b);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let C=b.split(/[\/\-\.]/);if(C.length===3){let[g,h,d]=C.map(y=>y.trim());if(g.length===4&&h.length<=2&&d.length<=2)return`${g}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&h.length<=2&&g.length<=2)return`${d}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return b},k=l.map(v=>({info_date:s(v["Tgl Info"]||v["Tanggal Info"]),branch_id:a(String(v.Cabang||"").trim()),problem:String(v.Permasalahan||"").trim(),pic:String(v.PIC||"").trim(),done_date:s(v["Tgl Done"]||v["Tanggal Done"]),status:String(v.Status||"").trim(),notes:String(v.Keterangan||v.Catatan||"").trim()})).filter(v=>v.info_date&&v.branch_id&&v.problem),w=await L("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(k)});if(!w.ok)throw new Error(w.data?.error||"Import gagal");return w.data}}})}ue();var hn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],fn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],yn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Dt="fcms_sop_fc_ik_v1",Lt="fcms_sop_fc_tools_v1",Bt="fcms_sop_fc_chem_v1";function ma(){try{let e=localStorage.getItem(Dt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(hn))}function Ft(e){localStorage.setItem(Dt,JSON.stringify(e))}function ga(){try{let e=localStorage.getItem(Lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(fn))}function Tt(e){localStorage.setItem(Lt,JSON.stringify(e))}function ba(){try{let e=localStorage.getItem(Bt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(yn))}function At(e){localStorage.setItem(Bt,JSON.stringify(e))}async function ha(e){let t="ik",i="",n="all",o="tools",c="",r=ma(),l=ga(),a=ba();function s(){e.innerHTML=`
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
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto;">
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${k(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${w(t==="ik")}">${r.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${k(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${w(t==="tools")}">${l.length+a.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${k(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${k(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,v(),b()}function k(u){return u?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function w(u){return u?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function v(){e.querySelectorAll(".sop-tab-btn").forEach(P=>{P.addEventListener("click",()=>{t=P.dataset.tab,s()})});let u=e.querySelector("#btn-reset-default");u&&u.addEventListener("click",()=>{ie("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Dt),localStorage.removeItem(Lt),localStorage.removeItem(Bt),r=ma(),l=ga(),a=ba(),K("Data SOP Facility Care berhasil di-reset ke standar asli"),s()},"Konfirmasi Reset Standar")});let $=e.querySelector("#btn-print-sop");$&&$.addEventListener("click",()=>{window.print()});let _=e.querySelector("#btn-export-tools");_&&_.addEventListener("click",()=>{E()})}function b(){let u=e.querySelector("#sop-content-area");u&&(t==="ik"?C(u):t==="tools"?d(u):t==="policy"?S(u):t==="links"&&f(u))}function C(u){u.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${r.length})</button>
              <button class="cat-pill ${n==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${r.filter(D=>D.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${n==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${r.filter(D=>D.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${n==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${r.filter(D=>D.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${n==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${r.filter(D=>D.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${n==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${r.filter(D=>D.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${r.length} dari ${r.length} Instruksi Kerja
            </div>
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let $=document.createElement("style");$.innerHTML=`
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
    `,u.appendChild($);let _=u.querySelector("#ik-cards-grid-container"),P=u.querySelector("#ik-count-label");function N(){let D=r.filter(B=>{let I=n==="all"||B.category===n,x=i.toLowerCase().trim(),T=!x||B.title.toLowerCase().includes(x)||B.target.toLowerCase().includes(x)||B.tools.some(A=>A.toLowerCase().includes(x))||B.steps.some(A=>A.toLowerCase().includes(x));return I&&T});if(P&&(P.textContent=`Menampilkan ${D.length} dari ${r.length} Instruksi Kerja`),D.length===0){_.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}_.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${D.map(B=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${B.bg||"#EFF6FF"}; border-radius:10px;">${B.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${B.id}. ${B.title}
                      </h3>
                      <span class="badge" style="background:${B.bg||"#EFF6FF"}; color:${B.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${B.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${B.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${B.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${B.color||"#2563EB"};">
                  <strong>Area:</strong> ${B.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${B.tools.slice(0,3).map(I=>`<li style="margin-bottom:2px;">${I}</li>`).join("")}
                    ${B.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${B.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${B.chemical_ratio&&B.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${B.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${B.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${B.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,_.querySelectorAll(".btn-view-ik").forEach(B=>{B.addEventListener("click",()=>{let I=parseInt(B.dataset.id,10),x=r.find(T=>T.id===I);x&&g(x)})}),_.querySelectorAll(".btn-edit-ik").forEach(B=>{B.addEventListener("click",()=>{let I=parseInt(B.dataset.id,10),x=r.find(T=>T.id===I);x&&h(x)})}),_.querySelectorAll(".btn-delete-ik").forEach(B=>{B.addEventListener("click",()=>{let I=parseInt(B.dataset.id,10),x=r.find(T=>T.id===I);x&&ie(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${x.title}"?`,()=>{r=r.filter(T=>T.id!==I),Ft(r),K("Instruksi Kerja berhasil dihapus"),s()},"Hapus Instruksi Kerja")})})}N();let M=u.querySelector("#ik-search");M&&M.addEventListener("input",D=>{i=D.target.value,N()}),u.querySelectorAll(".cat-pill").forEach(D=>{D.addEventListener("click",()=>{u.querySelectorAll(".cat-pill").forEach(B=>B.classList.remove("active")),D.classList.add("active"),n=D.dataset.cat,N()})});let F=u.querySelector("#btn-add-ik");F&&F.addEventListener("click",()=>{h()})}function g(u){Z({title:`${u.icon||"\u{1F4CB}"} IK-${u.id}: ${u.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${u.bg||"#EFF6FF"}; color:${u.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${u.category}</span>
            <span>ID: IK-${String(u.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${u.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${u.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${u.tools.map($=>`<li>${$}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${u.chemical_ratio&&u.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${u.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${u.steps.map($=>`<li style="margin-bottom:6px;">${$}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${u.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${u.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:($,_)=>{_()}})}function h(u=null){let $=!!u,_=u?u.tools.join(`
`):"",P=u?u.steps.join(`
`):"",N=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${u?p(u.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${u?u.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${u?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${u?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${u?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${u?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${u?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${u?p(u.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${u?p(u.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${p(_)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${p(P)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${u?p(u.notes||""):""}</textarea>
        </div>
      </form>
    `;Z({title:$?`\u270F\uFE0F Edit IK-${u.id}: ${u.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:N,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(M,F)=>{let D=M.querySelector("#form-ik-title").value.trim(),B=M.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",I=M.querySelector("#form-ik-cat").value,x=M.querySelector("#form-ik-ratio").value.trim(),T=M.querySelector("#form-ik-target").value.trim(),A=M.querySelector("#form-ik-tools").value.trim(),H=M.querySelector("#form-ik-steps").value.trim(),q=M.querySelector("#form-ik-notes").value.trim();if(!D||!T||!A||!H){j("Harap lengkapi seluruh field wajib (*)");return}let Q=A.split(`
`).map(de=>de.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),z=H.split(`
`).map(de=>de.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),ne={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[I]||{color:"#2563EB",bg:"#EFF6FF"};if($)u.title=D,u.icon=B,u.category=I,u.chemical_ratio=x,u.target=T,u.tools=Q,u.steps=z,u.notes=q,u.color=ne.color,u.bg=ne.bg,Ft(r),K(`Instruksi Kerja "${D}" berhasil diperbarui`);else{let de=r.length>0?Math.max(...r.map(Y=>Y.id))+1:1,me={id:de,title:D,category:I,icon:B,color:ne.color,bg:ne.bg,target:T,chemical_ratio:x||"-",tools:Q,steps:z,notes:q};r.push(me),Ft(r),K(`Instruksi Kerja baru IK-${de} berhasil ditambahkan`)}F(),s()}})}function d(u){u.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${o==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${l.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${o==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${a.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${c}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${o==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let $=u.querySelector("#tool-table-container"),_=u.querySelector("#tool-search");function P(){let N=c.toLowerCase().trim();if(o==="tools"){let M=l.filter(F=>!N||F.name.toLowerCase().includes(N)||F.area.toLowerCase().includes(N));$.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${M.length} Item</span>
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
                  ${M.map((F,D)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${D%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${F.no||D+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${p(F.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${p(F.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${p(F.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${F.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${F.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,$.querySelectorAll(".btn-edit-tool").forEach(F=>{F.addEventListener("click",()=>{let D=parseInt(F.dataset.no,10),B=l.find(I=>I.no===D);B&&y(B)})}),$.querySelectorAll(".btn-delete-tool").forEach(F=>{F.addEventListener("click",()=>{let D=parseInt(F.dataset.no,10),B=l.find(I=>I.no===D);B&&ie(`Apakah Anda yakin ingin menghapus alat: "${B.name}"?`,()=>{l=l.filter(I=>I.no!==D),l.forEach((I,x)=>{I.no=x+1}),Tt(l),K("Alat kebersihan berhasil dihapus"),s()},"Hapus Alat Kebersihan")})})}else{let M=a.filter(F=>!N||F.name.toLowerCase().includes(N)||F.function.toLowerCase().includes(N));$.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${M.length} Chemical</span>
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
                  ${M.map((F,D)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${D%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${F.no||D+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${p(F.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${p(F.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${p(F.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${p(F.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${F.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${F.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,$.querySelectorAll(".btn-edit-chem").forEach(F=>{F.addEventListener("click",()=>{let D=parseInt(F.dataset.no,10),B=a.find(I=>I.no===D);B&&m(B)})}),$.querySelectorAll(".btn-delete-chem").forEach(F=>{F.addEventListener("click",()=>{let D=parseInt(F.dataset.no,10),B=a.find(I=>I.no===D);B&&ie(`Apakah Anda yakin ingin menghapus chemical: "${B.name}"?`,()=>{a=a.filter(I=>I.no!==D),a.forEach((I,x)=>{I.no=x+1}),At(a),K("Bahan chemical berhasil dihapus"),s()},"Hapus Bahan Chemical")})})}}P(),u.querySelector("#view-tools-btn").addEventListener("click",()=>{o="tools",d(u)}),u.querySelector("#view-chem-btn").addEventListener("click",()=>{o="chemical",d(u)}),_&&_.addEventListener("input",N=>{c=N.target.value,P()}),u.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{o==="tools"?y():m()})}function y(u=null){let $=!!u,_=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${u?p(u.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${u?p(u.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${u?p(u.area):""}</textarea>
        </div>
      </form>
    `;Z({title:$?`\u270F\uFE0F Edit Alat: ${u.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:_,confirmText:"Simpan",onConfirm:(P,N)=>{let M=P.querySelector("#form-tool-name").value.trim(),F=P.querySelector("#form-tool-qty").value.trim(),D=P.querySelector("#form-tool-area").value.trim();if(!M||!F||!D){j("Harap lengkapi semua kolom");return}if($)u.name=M,u.qty=F,u.area=D,Tt(l),K(`Alat "${M}" berhasil diperbarui`);else{let B=l.length>0?Math.max(...l.map(I=>I.no))+1:1;l.push({no:B,name:M,qty:F,area:D}),Tt(l),K(`Alat baru "${M}" berhasil ditambahkan`)}N(),s()}})}function m(u=null){let $=!!u,_=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${u?p(u.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${u?p(u.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${u?p(u.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${u?p(u.function):""}</textarea>
        </div>
      </form>
    `;Z({title:$?`\u270F\uFE0F Edit Chemical: ${u.name}`:"\u2795 Tambah Chemical Baru",content:_,confirmText:"Simpan",onConfirm:(P,N)=>{let M=P.querySelector("#form-chem-name").value.trim(),F=P.querySelector("#form-chem-size").value.trim(),D=P.querySelector("#form-chem-ratio").value.trim(),B=P.querySelector("#form-chem-function").value.trim();if(!M||!F||!D||!B){j("Harap lengkapi semua kolom");return}if($)u.name=M,u.size=F,u.ratio=D,u.function=B,At(a),K(`Chemical "${M}" berhasil diperbarui`);else{let I=a.length>0?Math.max(...a.map(x=>x.no))+1:1;a.push({no:I,name:M,size:F,ratio:D,function:B}),At(a),K(`Chemical baru "${M}" berhasil ditambahkan`)}N(),s()}})}function S(u){u.innerHTML=`
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
    `}function f(u){u.innerHTML=`
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
    `}async function E(){try{let{downloadExcel:u}=await Promise.resolve().then(()=>(U(),be)),$=r.map(P=>({"No IK":`IK-${P.id}`,"Judul Pekerjaan":P.title,Kategori:P.category,"Sasaran Target Area":P.target,"Rasio Chemical":P.chemical_ratio||"-","Alat & Bahan":P.tools.join(", "),"Langkah Kerja":P.steps.map((N,M)=>`${M+1}. ${N}`).join(" | "),"Catatan Safety":P.notes||"-"})),_=new Date().toISOString().slice(0,10);u($,`SOP_Facility_Care_Instruksi_Kerja_${_}`),K("Data Excel Instruksi Kerja berhasil diunduh")}catch(u){console.error(u),j("Gagal mendownload data Excel: "+u.message)}}function p(u){return u?String(u).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}s()}ue();var xn=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],vn=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],fa={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},Mt="fcms_sop_sec_ik_v1",Nt="fcms_sop_sec_assets_v1";function ya(){try{let e=localStorage.getItem(Mt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(xn))}function Pt(e){localStorage.setItem(Mt,JSON.stringify(e))}function xa(){try{let e=localStorage.getItem(Nt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(vn))}function It(e){localStorage.setItem(Nt,JSON.stringify(e))}async function va(e){let t="ik",i="",n="all",o="",c="cabang",r=ya(),l=xa();function a(){e.innerHTML=`
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
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto;">
        <button class="sec-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${s(t==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${k(t==="ik")}">${r.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="zonasi"?"active":""}" data-tab="zonasi" style="${s(t==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${t==="assets"?"active":""}" data-tab="assets" style="${s(t==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${k(t==="assets")}">${l.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${s(t==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${s(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,w(),v()}function s(p){return p?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function k(p){return p?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function w(){e.querySelectorAll(".sec-tab-btn").forEach(_=>{_.addEventListener("click",()=>{t=_.dataset.tab,a()})});let p=e.querySelector("#btn-reset-sec-default");p&&p.addEventListener("click",()=>{ie("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(Mt),localStorage.removeItem(Nt),r=ya(),l=xa(),K("Data SOP Security berhasil di-reset ke standar dokumen"),a()},"Reset ke Standar Dokumen")});let u=e.querySelector("#btn-print-sec");u&&u.addEventListener("click",()=>{window.print()});let $=e.querySelector("#btn-export-sec");$&&$.addEventListener("click",()=>{f()})}function v(){let p=e.querySelector("#sec-content-area");p&&(t==="ik"?b(p):t==="zonasi"?h(p):t==="assets"?d(p):t==="policy"?m(p):t==="links"&&S(p))}function b(p){p.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${r.length})</button>
              <button class="sec-cat-pill ${n==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${r.filter(F=>F.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${n==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${r.filter(F=>F.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${n==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${r.filter(F=>F.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${r.length} Instruksi Kerja
            </div>
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let u=document.createElement("style");u.innerHTML=`
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
    `,p.appendChild(u);let $=p.querySelector("#sec-ik-cards-grid-container"),_=p.querySelector("#sec-ik-count-label");function P(){let F=r.filter(D=>{let B=n==="all"||D.category===n,I=i.toLowerCase().trim(),x=!I||D.title.toLowerCase().includes(I)||D.target.toLowerCase().includes(I)||D.tools.some(T=>T.toLowerCase().includes(I))||D.steps.some(T=>T.toLowerCase().includes(I));return B&&x});if(_&&(_.textContent=`Menampilkan ${F.length} dari ${r.length} Instruksi Kerja`),F.length===0){$.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}$.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${F.map(D=>`
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${D.bg||"#EFF6FF"}; border-radius:10px;">${D.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(D.id).padStart(2,"0")}. ${D.title}
                      </h3>
                      <span class="badge" style="background:${D.bg||"#EFF6FF"}; color:${D.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${D.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${D.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${D.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${D.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${D.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${D.tools.slice(0,3).map(B=>`<li style="margin-bottom:2px;">${E(B)}</li>`).join("")}
                    ${D.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${D.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${D.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${D.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,$.querySelectorAll(".btn-view-sec-ik").forEach(D=>{D.addEventListener("click",()=>{let B=parseInt(D.dataset.id,10),I=r.find(x=>x.id===B);I&&C(I)})}),$.querySelectorAll(".btn-edit-sec-ik").forEach(D=>{D.addEventListener("click",()=>{let B=parseInt(D.dataset.id,10),I=r.find(x=>x.id===B);I&&g(I)})}),$.querySelectorAll(".btn-delete-sec-ik").forEach(D=>{D.addEventListener("click",()=>{let B=parseInt(D.dataset.id,10),I=r.find(x=>x.id===B);I&&ie(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${I.title}"?`,()=>{r=r.filter(x=>x.id!==B),Pt(r),K("Instruksi Kerja Security berhasil dihapus"),a()},"Hapus Instruksi Kerja")})})}P();let N=p.querySelector("#sec-ik-search");N&&N.addEventListener("input",F=>{i=F.target.value,P()}),p.querySelectorAll(".sec-cat-pill").forEach(F=>{F.addEventListener("click",()=>{p.querySelectorAll(".sec-cat-pill").forEach(D=>D.classList.remove("active")),F.classList.add("active"),n=F.dataset.cat,P()})});let M=p.querySelector("#btn-add-sec-ik");M&&M.addEventListener("click",()=>{g()})}function C(p){Z({title:`${p.icon||"\u{1F6E1}\uFE0F"} IK-${String(p.id).padStart(2,"0")}: ${p.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${p.bg||"#EFF6FF"}; color:${p.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${p.category}</span>
            <span>Kode: IK-SEC-${String(p.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${p.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${p.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${p.tools.map(u=>{if(u.includes("http")){let $=u.split(": ");return`<li>${$[0]}: <a href="${$.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${E(u)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${p.steps.map(u=>`<li style="margin-bottom:6px;">${E(u)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${p.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${E(p.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(u,$)=>{$()}})}function g(p=null){let u=!!p,$=p?p.tools.join(`
`):"",_=p?p.steps.join(`
`):"",P=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${p?E(p.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${p?p.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${p?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${p?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${p?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${p?E(p.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${E($)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${E(_)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${p?E(p.notes||""):""}</textarea>
        </div>
      </form>
    `;Z({title:u?`\u270F\uFE0F Edit IK-${p.id}: ${p.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:P,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(N,M)=>{let F=N.querySelector("#form-sec-ik-title").value.trim(),D=N.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",B=N.querySelector("#form-sec-ik-cat").value,I=N.querySelector("#form-sec-ik-target").value.trim(),x=N.querySelector("#form-sec-ik-tools").value.trim(),T=N.querySelector("#form-sec-ik-steps").value.trim(),A=N.querySelector("#form-sec-ik-notes").value.trim();if(!F||!I||!x||!T){j("Harap lengkapi semua kolom bertanda bintang (*)");return}let H=x.split(`
`).map(O=>O.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),q=T.split(`
`).map(O=>O.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),z={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[B]||{color:"#1E40AF",bg:"#EFF6FF"};if(u)p.title=F,p.icon=D,p.category=B,p.target=I,p.tools=H,p.steps=q,p.notes=A,p.color=z.color,p.bg=z.bg,Pt(r),K(`Instruksi Kerja "${F}" berhasil diperbarui`);else{let O=r.length>0?Math.max(...r.map(de=>de.id))+1:1,ne={id:O,title:F,category:B,icon:D,color:z.color,bg:z.bg,target:I,tools:H,steps:q,notes:A};r.push(ne),Pt(r),K(`Instruksi Kerja baru IK-${O} berhasil ditambahkan`)}M(),a()}})}function h(p){p.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${c==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${c==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let u=p.querySelector("#zonasi-content-container");c==="cabang"?u.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${fa.cabang.map($=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${$.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${$.color}; display:inline-block;"></span>
                  ${$.zone}
                </h3>
                <span class="badge" style="background:${$.bg}; color:${$.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${$.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${$.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${$.rooms.map(_=>`<li>${_}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:u.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${fa.hq.map($=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${$.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${$.hijau.map(_=>`<li>${_}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${$.kuning.map(_=>`<li>${_}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${$.merah.map(_=>`<li>${_}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,p.querySelector("#view-cabang-btn").addEventListener("click",()=>{c="cabang",h(p)}),p.querySelector("#view-hq-btn").addEventListener("click",()=>{c="hq",h(p)})}function d(p){p.innerHTML=`
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
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${o}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let u=p.querySelector("#sec-asset-table-container"),$=p.querySelector("#sec-asset-search");function _(){let N=o.toLowerCase().trim(),M=l.filter(F=>!N||F.item.toLowerCase().includes(N)||F.type.toLowerCase().includes(N)||F.description.toLowerCase().includes(N)||F.pic&&F.pic.toLowerCase().includes(N));u.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${M.length} Aset</span>
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
                ${M.map((F,D)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${D%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${F.no||D+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${P(F.type)}">${E(F.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${E(F.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${E(F.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${E(F.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${F.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${F.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,u.querySelectorAll(".btn-edit-asset").forEach(F=>{F.addEventListener("click",()=>{let D=parseInt(F.dataset.no,10),B=l.find(I=>I.no===D);B&&y(B)})}),u.querySelectorAll(".btn-delete-asset").forEach(F=>{F.addEventListener("click",()=>{let D=parseInt(F.dataset.no,10),B=l.find(I=>I.no===D);B&&ie(`Apakah Anda yakin ingin menghapus aset: "${B.item}"?`,()=>{l=l.filter(I=>I.no!==D),l.forEach((I,x)=>{I.no=x+1}),It(l),K("Aset berhasil dihapus"),a()},"Hapus Aset")})})}function P(N){return N==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":N==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}_(),$&&$.addEventListener("input",N=>{o=N.target.value,_()}),p.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{y()})}function y(p=null){let u=!!p,$=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${p?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${p?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${p?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${p?E(p.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${p?E(p.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${p?E(p.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;Z({title:u?`\u270F\uFE0F Edit Aset: ${p.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:$,confirmText:"Simpan",onConfirm:(_,P)=>{let N=_.querySelector("#form-asset-type").value,M=_.querySelector("#form-asset-item").value.trim(),F=_.querySelector("#form-asset-desc").value.trim(),D=_.querySelector("#form-asset-pic").value.trim();if(!M||!F||!D){j("Harap lengkapi semua kolom");return}if(u)p.type=N,p.item=M,p.description=F,p.pic=D,It(l),K(`Aset "${M}" berhasil diperbarui`);else{let B=l.length>0?Math.max(...l.map(I=>I.no))+1:1;l.push({no:B,type:N,item:M,description:F,pic:D}),It(l),K(`Aset baru "${M}" berhasil ditambahkan`)}P(),a()}})}function m(p){p.innerHTML=`
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
    `}function S(p){p.innerHTML=`
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
    `}async function f(){try{let{downloadExcel:p}=await Promise.resolve().then(()=>(U(),be)),u=r.map(_=>({"Kode IK":`IK-SEC-${String(_.id).padStart(2,"0")}`,"Judul Prosedur":_.title,Kategori:_.category,"Sasaran Target":_.target,"Alat & Perlengkapan":_.tools.join(", "),"Langkah Pelaksanaan":_.steps.map((P,N)=>`${N+1}. ${P}`).join(" | "),"Catatan Khusus":_.notes||"-"})),$=new Date().toISOString().slice(0,10);p(u,`SOP_Security_Instruksi_Kerja_${$}`),K("Data Excel SOP Security berhasil diunduh")}catch(p){console.error(p),j("Gagal mendownload data Excel: "+p.message)}}function E(p){return p?String(p).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}a()}G();ue();var ka=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Kt(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function zt(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function wa(e){let t=[],i="all",n="",o=localStorage.getItem("fcms_doc_view")||"grid",c="fcms-master-doc-styles";if(!document.getElementById(c)){let g=document.createElement("style");g.id=c,g.textContent=`
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
      .doc-view-toggle {
        display: flex;
        background: #F1F5F9;
        padding: 4px;
        border-radius: 10px;
        gap: 4px;
      }
      .doc-view-btn {
        border: none;
        background: transparent;
        padding: 6px 14px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748B;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 0.15s ease;
      }
      .doc-view-btn.active {
        background: #FFFFFF;
        color: #1E293B;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
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
    `,document.head.appendChild(g)}async function r(){try{let g=await L("/api/checklist?limit=1000");if(g.ok&&g.data?.data&&Array.isArray(g.data.data)&&g.data.data.length>0)t=g.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t));else{let h=localStorage.getItem("fcms_master_doc_cache");h?t=JSON.parse(h):(t=[...ka],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)))}}catch(g){console.warn("Fallback loading documents from cache/seed:",g);let h=localStorage.getItem("fcms_master_doc_cache");t=h?JSON.parse(h):[...ka]}b()}function l(){return t.filter(g=>{let h=i==="all"||g.category===i,d=n.toLowerCase().trim(),y=!d||(g.name||"").toLowerCase().includes(d)||(g.category||"").toLowerCase().includes(d)||(g.description||"").toLowerCase().includes(d);return h&&y})}function a(){let g={all:t.length};return t.forEach(h=>{let d=h.category||"Lainnya";g[d]=(g[d]||0)+1}),g}function s(g=null){let h=!!g,d=Array.from(new Set(t.map(m=>m.category).filter(Boolean))),y=document.createElement("div");y.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${g?.name?String(g.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${g?.category?String(g.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${d.map(m=>`<option value="${m}">${m}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${g?.document_link?String(g.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${g?.description||""}</textarea>
        </div>
      </div>
    `,Z({title:h?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:y,size:"md",confirmText:h?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(m,S)=>{let f=y.querySelector("#doc-modal-name"),E=y.querySelector("#doc-modal-category"),p=y.querySelector("#doc-modal-link"),u=y.querySelector("#doc-modal-desc"),$=(f.value||"").trim(),_=(E.value||"").trim(),P=(p.value||"").trim(),N=(u.value||"").trim();if(!$){j("Nama Dokumen wajib diisi"),f.focus();return}if(!_){j("Kategori Dokumen wajib diisi"),E.focus();return}let M={name:$,category:_,document_link:P,description:N};try{if(h)if((await L(`/api/checklist/${g.id}`,{method:"PUT",body:JSON.stringify(M)})).ok){let D=t.findIndex(B=>B.id===g.id);D!==-1&&(t[D]={...t[D],...M}),K("Dokumen berhasil diperbarui")}else{let D=t.findIndex(B=>B.id===g.id);D!==-1&&(t[D]={...t[D],...M}),K("Dokumen disimpan secara lokal")}else{let F=await L("/api/checklist",{method:"POST",body:JSON.stringify(M)});if(F.ok&&F.data?.id)t.unshift({id:F.data.id,...M}),K("Dokumen baru berhasil ditambahkan");else{let D=Date.now();t.unshift({id:D,...M}),K("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),S(),b()}catch(F){console.error("Error saving document:",F),j("Gagal menyimpan dokumen: "+F.message)}}})}function k(g){ie(`Apakah Anda yakin ingin menghapus dokumen <strong>"${g.name}"</strong>?`,async()=>{try{let h=await L(`/api/checklist/${g.id}`,{method:"DELETE"});h.ok||console.warn("API delete warning:",h),t=t.filter(d=>d.id!==g.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),K(`Dokumen "${g.name}" berhasil dihapus`),b()}catch{t=t.filter(d=>d.id!==g.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),K(`Dokumen "${g.name}" berhasil dihapus`),b()}},"Hapus Dokumen")}function w(g,h){if(!g){xt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(g).then(()=>{K(`Tautan "${h}" disalin ke clipboard!`)}).catch(()=>{xt(`Tautan: ${g}`)})}async function v(){try{let{downloadExcel:g}=await Promise.resolve().then(()=>(U(),be)),h=t.map((d,y)=>({No:y+1,"Nama Dokumen":d.name||"",Kategori:d.category||"","Platform / Format":zt(d.document_link).type,"Link Dokumen":d.document_link||"","Deskripsi / Keterangan":d.description||""}));g(h,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),K("Berhasil mengunduh data dokumen ke Excel.")}catch(g){j("Gagal mengekspor data: "+g.message)}}function b(){let g=l(),h=a(),d=Object.keys(h).filter(m=>m!=="all"),y=t.filter(m=>!!m.document_link).length;e.innerHTML=`
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
                <div class="doc-stat-value">${d.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${y} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${t.length}</span></div>
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
            <div class="doc-search-wrapper">
              <span class="doc-search-icon">\u{1F50D}</span>
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${n?n.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- View Mode Switch -->
            <div class="doc-view-toggle">
              <button id="doc-view-grid" class="doc-view-btn ${o==="grid"?"active":""}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
                <span>Grid Kartu</span>
              </button>
              <button id="doc-view-table" class="doc-view-btn ${o==="table"?"active":""}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel Data</span>
              </button>
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${i==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${h.all||0}</span>
            </button>
            ${d.map(m=>{let S=Kt(m);return`
                <button class="doc-pill ${i===m?"active":""}" data-cat="${m}">
                  <span>${S.icon}</span>
                  <span>${m}</span>
                  <span class="doc-pill-count">${h[m]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${g.length===0?`
            <div class="doc-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Dokumen Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Coba sesuaikan kata kunci pencarian atau pilih kategori lain.</p>
              <button id="doc-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Tampilkan Semua Dokumen
              </button>
            </div>
          `:o==="grid"?`
            <!-- Grid Card View -->
            <div class="doc-cards-grid">
              ${g.map(m=>{let S=Kt(m.category),f=zt(m.document_link);return`
                  <div class="doc-card" data-id="${m.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${S.bg}; color:${S.color}; border:1px solid ${S.border};">
                          <span>${S.icon}</span>
                          <span>${S.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${f.icon}</span>
                          <span>${f.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${m.name}</h3>
                      <p class="doc-card-desc">${m.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${m.document_link?`
                        <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${m.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${m.document_link}" data-title="${m.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${m.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${m.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `:`
            <!-- Table View -->
            <div class="doc-table-card">
              <table class="doc-table">
                <thead>
                  <tr>
                    <th style="width:50px; text-align:center;">No</th>
                    <th>Nama Dokumen</th>
                    <th>Kategori</th>
                    <th>Format</th>
                    <th>Tautan Dokumen</th>
                    <th>Deskripsi</th>
                    <th style="width:100px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${g.map((m,S)=>{let f=Kt(m.category),E=zt(m.document_link);return`
                      <tr>
                        <td style="text-align:center; font-weight:600; color:#94A3B8;">${S+1}</td>
                        <td>
                          <div style="font-weight:700; color:#0F172A;">${m.name}</div>
                        </td>
                        <td>
                          <span class="doc-category-badge" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border};">
                            <span>${f.icon}</span>
                            <span>${f.label}</span>
                          </span>
                        </td>
                        <td>
                          <span class="doc-platform-badge">
                            <span>${E.icon}</span>
                            <span>${E.type}</span>
                          </span>
                        </td>
                        <td>
                          ${m.document_link?`
                            <div style="display:flex; align-items:center; gap:6px;">
                              <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-primary" style="font-size:0.78rem; padding:5px 10px; border-radius:6px; display:inline-flex; align-items:center; gap:4px; text-decoration:none;">
                                <span>Buka Dokumen</span>
                                <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                              </a>
                              <button class="doc-icon-btn copy-link-btn" title="Salin Link" style="width:28px; height:28px;" data-link="${m.document_link}" data-title="${m.name}">
                                <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                              </button>
                            </div>
                          `:'<span style="color:#94A3B8; font-size:0.82rem;">-</span>'}
                        </td>
                        <td style="max-width:280px; color:#64748B; font-size:0.84rem;">
                          ${m.description||"-"}
                        </td>
                        <td style="text-align:center;">
                          <div style="display:flex; justify-content:center; gap:6px;">
                            <button class="doc-icon-btn edit-doc-btn" style="width:30px; height:30px;" title="Edit" data-id="${m.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                            <button class="doc-icon-btn delete delete-doc-btn" style="width:30px; height:30px;" title="Hapus" data-id="${m.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          </div>
                        </td>
                      </tr>
                    `}).join("")}
                </tbody>
              </table>
            </div>
          `}
        </div>

      </div>
    `,C()}function C(){let g=e.querySelector("#doc-btn-add");g&&(g.onclick=()=>s());let h=e.querySelector("#doc-btn-excel");h&&(h.onclick=()=>v());let d=e.querySelector("#doc-btn-refresh");d&&(d.onclick=()=>r());let y=e.querySelector("#doc-btn-reset-filter");y&&(y.onclick=()=>{n="",i="all",b()});let m=e.querySelector("#doc-search-input");m&&(m.oninput=E=>{n=E.target.value;let p=l();b();let u=e.querySelector("#doc-search-input");u&&(u.focus(),u.setSelectionRange(u.value.length,u.value.length))});let S=e.querySelector("#doc-view-grid"),f=e.querySelector("#doc-view-table");S&&(S.onclick=()=>{o="grid",localStorage.setItem("fcms_doc_view","grid"),b()}),f&&(f.onclick=()=>{o="table",localStorage.setItem("fcms_doc_view","table"),b()}),e.querySelectorAll(".doc-pill").forEach(E=>{E.onclick=()=>{i=E.getAttribute("data-cat")||"all",b()}}),e.querySelectorAll(".copy-link-btn").forEach(E=>{E.onclick=p=>{p.stopPropagation();let u=E.getAttribute("data-link"),$=E.getAttribute("data-title");w(u,$)}}),e.querySelectorAll(".edit-doc-btn").forEach(E=>{E.onclick=p=>{p.stopPropagation();let u=Number(E.getAttribute("data-id")),$=t.find(_=>_.id===u);$&&s($)}}),e.querySelectorAll(".delete-doc-btn").forEach(E=>{E.onclick=p=>{p.stopPropagation();let u=Number(E.getAttribute("data-id")),$=t.find(_=>_.id===u);$&&k($)}})}await r()}G();ue();U();var Me={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Ot(e){for(let[t,i]of Object.entries(Me))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function Sa(e){e.innerHTML=`
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
        ${Object.entries(Me).map(([m,S])=>`
          <button class="zone-tab-btn" data-zone="${m}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${S.icon}</span> ${m}
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
  `;let t="all",i="table",n=[],o=oe(),c=o&&["superadmin","admin","manager","spv"].includes(o.role);async function r(){try{let m=await L("/api/hygiene-standards?all=1");if(!m.ok)throw new Error("Gagal memuat standar hygiene");n=Array.isArray(m.data)?m.data:m.data?.data||[],a(),l(),k()}catch(m){console.error(m),j(m.message||"Terjadi kesalahan memuat data");let S=e.querySelector("#hygiene-view-container");S&&(S.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${m.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,S.querySelector("#btn-retry-hygiene")?.addEventListener("click",r))}}function l(){let m=e.querySelector("#badge-count-all");m&&(m.textContent=n.length),Object.entries(Me).forEach(([E,p])=>{let u=E.replace(/[\s&]+/g,"-"),$=e.querySelector(`#badge-count-${u}`);if($){let _=n.filter(P=>p.rooms.includes(P.room_name)).length;$.textContent=_}});let S=new Set(n.map(E=>E.room_name)),f=e.querySelector("#card-total-rooms");f&&(f.textContent=S.size)}function a(){let m=e.querySelector("#hygiene-room-filter");if(!m)return;let S=m.value,f=[];if(t==="all")f=Array.from(new Set(n.map(E=>E.room_name))).filter(Boolean).sort();else{let E=Me[t];E&&(f=E.rooms.filter(p=>n.some(u=>u.room_name===p)))}m.innerHTML=`<option value="">Semua Ruangan (${f.length})</option>`+f.map(E=>`<option value="${E}" ${E===S?"selected":""}>${E}</option>`).join("")}function s(){let m=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),S=e.querySelector("#hygiene-room-filter")?.value||"";return n.filter(f=>{if(t!=="all"){let E=Me[t];if(!E||!E.rooms.includes(f.room_name))return!1}if(S&&f.room_name!==S)return!1;if(m){let E=(f.room_name||"").toLowerCase(),p=(f.item_name||"").toLowerCase(),u=(f.cleanliness_standard||"").toLowerCase();if(!E.includes(m)&&!p.includes(m)&&!u.includes(m))return!1}return!0})}function k(){let m=s(),S=e.querySelector("#hygiene-stats-text");S&&(S.innerHTML=`Menampilkan <strong>${m.length}</strong> dari <strong>${n.length}</strong> parameter`);let f=e.querySelector("#hygiene-view-container");if(f){if(m.length===0){f.innerHTML=`
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
      `,f.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let E=e.querySelector("#hygiene-search"),p=e.querySelector("#hygiene-room-filter");E&&(E.value=""),p&&(p.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(u=>u.classList.toggle("active",u.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(u=>{u.dataset.zone==="all"?(u.style.background="#2563EB",u.style.color="#fff"):(u.style.background="#F1F5F9",u.style.color="var(--text-2)")}),a(),k()});return}i==="table"?w(f,m):v(f,m)}}function w(m,S){m.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${c?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${S.map((f,E)=>{let p=Ot(f.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${E+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${p.bg}; color:${p.color}; border:1px solid ${p.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${p.icon}</span> ${f.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${p.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${f.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${f.cleanliness_standard}
                      </div>
                    </td>
                    ${c?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${f.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${f.id}" data-name="${f.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,b(m)}function v(m,S){let f={};S.forEach(E=>{let p=Ot(E.room_name);f[p.zoneName]||(f[p.zoneName]={info:p,rooms:{}}),f[p.zoneName].rooms[E.room_name]||(f[p.zoneName].rooms[E.room_name]=[]),f[p.zoneName].rooms[E.room_name].push(E)}),m.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(f).map(([E,p])=>{let u=Object.values(p.rooms).reduce(($,_)=>$+_.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${p.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${p.info.bg}; border-bottom:1px solid ${p.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${p.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${p.info.color};">
                      ${E}
                    </h3>
                    <div style="font-size:0.78rem; color:${p.info.color}; opacity:0.85;">
                      ${Object.keys(p.rooms).length} Ruangan \xB7 ${u} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${p.info.color}; border:1px solid ${p.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${u} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(p.rooms).map(([$,_])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${$}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${_.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${c?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${_.map((P,N)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${N+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${P.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${P.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${c?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${P.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${P.id}" data-name="${P.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,b(m)}function b(m){m.querySelectorAll(".btn-edit-item").forEach(S=>{S.addEventListener("click",()=>{let f=S.dataset.id,E=n.find(p=>String(p.id)===String(f));E&&C(E)})}),m.querySelectorAll(".btn-delete-item").forEach(S=>{S.addEventListener("click",()=>{let f=S.dataset.id,E=S.dataset.name;ie(`Yakin ingin menghapus parameter standar <strong>"${E}"</strong>?`,async()=>{try{let p=await L(`/api/hygiene-standards/${f}`,{method:"DELETE"});if(!p.ok)throw new Error(p.data?.error||"Gagal menghapus parameter");K("Parameter kebersihan berhasil dihapus"),r()}catch(p){j(p.message)}},"Hapus Parameter Standar")})})}function C(m=null){let S=!!m,f=Array.from(new Set(n.map(p=>p.room_name))).filter(Boolean).sort(),E=document.createElement("div");E.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${m?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${f.map(p=>`<option value="${p}">`).join("")}
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
    `,Z({title:S?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:E,confirmText:S?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(p,u)=>{let $=E.querySelector("#modal-room-name").value.trim(),_=E.querySelector("#modal-item-name").value.trim(),P=E.querySelector("#modal-standard").value.trim();if(!$||!_||!P){j("Semua kolom wajib diisi");return}try{let N={room_name:$,item_name:_,cleanliness_standard:P},M=S?`/api/hygiene-standards/${m.id}`:"/api/hygiene-standards",D=await L(M,{method:S?"PUT":"POST",body:JSON.stringify(N)});if(!D.ok)throw new Error(D.data?.error||"Gagal menyimpan parameter");K(S?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),u(),r()}catch(N){j(N.message)}}})}function g(){let m=window.open("","_blank");if(!m){j("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let f=`
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

        ${Object.entries(Me).map(([E,p])=>{let u=n.filter($=>p.rooms.includes($.room_name));return u.length===0?"":`
            <div class="zone-title">${p.icon} ${E.toUpperCase()} (${u.length} PARAMETER)</div>
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
                ${u.map(($,_)=>`
                  <tr>
                    <td class="td-no">${_+1}</td>
                    <td class="td-room">${$.room_name}</td>
                    <td class="td-obj">${$.item_name}</td>
                    <td>${$.cleanliness_standard}</td>
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
    `;m.document.open(),m.document.write(f),m.document.close()}function h(){let m=n.map((f,E)=>{let p=Ot(f.room_name);return{No:E+1,"Kategori Zona":p.zoneName,"Ruangan / Lokasi":f.room_name,"Item / Objek Inspeksi":f.item_name,"Standar Kebersihan Fisik":f.cleanliness_standard}}),S=new Date().toISOString().slice(0,10);R(m,`Master_Standar_Kebersihan_FCMS_${S}`),K("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>k()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>k()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let m=e.querySelector("#hygiene-search"),S=e.querySelector("#hygiene-room-filter");m&&(m.value=""),S&&(S.value=""),k()}),e.querySelectorAll(".zone-tab-btn").forEach(m=>{m.addEventListener("click",()=>{t=m.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(S=>{let f=S.dataset.zone===t;if(S.classList.toggle("active",f),f){S.style.background="#2563EB",S.style.color="#fff";let E=S.querySelector(".badge");E&&(E.style.background="rgba(255,255,255,0.25)",E.style.color="#fff")}else{S.style.background="#F1F5F9",S.style.color="var(--text-2)";let E=S.querySelector(".badge");E&&(E.style.background="#CBD5E1",E.style.color="#334155")}}),a(),k()})});let d=e.querySelector("#view-table-btn"),y=e.querySelector("#view-zone-btn");d?.addEventListener("click",()=>{i="table",d.style.background="#fff",d.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",y.style.background="transparent",y.style.boxShadow="none",k()}),y?.addEventListener("click",()=>{i="zone",y.style.background="#fff",y.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",d.style.background="transparent",d.style.boxShadow="none",k()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",g),e.querySelector("#btn-export-hygiene")?.addEventListener("click",h),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>C(null)),r()}G();ue();U();async function Ea(e,t="supply"){return kn(e)}async function kn(e){let i=((await L("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));J({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let o=JSON.parse(n);return Array.isArray(o)?o.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let o=JSON.parse(n);return Array.isArray(o)?o.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>te(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let o=n?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let c=n?.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!i.find(r=>r.value==n.branch_id)?[...i,{value:n.branch_id,label:n.branch_name||n.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:c},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let o=new URLSearchParams(n||{}).toString(),c=await L(`/api/reports/supply?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(l=>{let a=l.tools_items;try{a=Array.isArray(JSON.parse(a))?JSON.parse(a).join(", "):a}catch{}let s=l.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":a||"",Chemical:s||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});R(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let c=(await L("/api/branches?all=1")).data?.data||[],r=k=>{if(!k)return null;let w=String(k||"").toLowerCase(),v=c.find(b=>String(b.full_name||"").toLowerCase()===w||String(b.code||"").toLowerCase()===w||String(b.name||"").toLowerCase()===w);return v?v.id:null},l=k=>{if(k==null||k==="")return"";if(k instanceof Date&&!isNaN(k.getTime()))return k.toISOString().slice(0,10);let w=String(k).trim();if(w===""||w==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(w))return w.slice(0,10);if(/^\d{4,5}$/.test(w)){let b=Number(w);if(b>2e4&&b<99999){let C=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let v=w.split(/[\/\-\.]/);if(v.length===3){let[b,C,g]=v.map(h=>h.trim());if(b.length===4&&C.length<=2&&g.length<=2)return`${b}-${C.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&C.length<=2&&b.length<=2)return`${g}-${C.padStart(2,"0")}-${b.padStart(2,"0")}`}return w},a=n.map(k=>({submitted_at:l(k.Waktu||k.Tanggal),submitter_name:String(k.Pengirim||"").trim(),branch_id:r(String(k.Cabang||"").trim()),tools_items:String(k["Alat/Barang"]||k.Alat||"").trim(),chemical_items:String(k.Chemical||"").trim(),additional_notes:String(k.Catatan||k.Keterangan||"").trim(),status:String(k.Status||"").trim(),processed_by:String(k["Diproses Oleh"]||k.PIC||"").trim()})).filter(k=>k.submitted_at&&k.submitter_name&&k.branch_id),s=await L("/api/reports/supply/import",{method:"POST",body:JSON.stringify(a)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,o)=>{let c=Z({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,l)=>{let a=r.querySelector("#supply-status").value,s=r.querySelector("#supply-processed-by").value;(await L(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:a,processed_by:s})})).ok?(K("Status diperbarui."),l(),o()):j("Gagal update status.")}})}}]})}G();U();async function Ca(e){let t=oe();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}J({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let n=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await L(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let n=i.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));R(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let n=i.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),username:String(c.Username||"").trim(),email:String(c.Email||"").trim(),role:String(c.Role||"").trim()||"viewer",password:String(c.Password||"").trim()})).filter(c=>c.username&&c.password&&c.email&&c.full_name),o=await L("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}G();U();async function $a(e){J({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await L(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)R(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),n=await L("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}G();async function _a(e){let t=new Date,i=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(r=>r.addEventListener("change",o));async function n(){try{let r=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await L(`/api/dashboard/calendar?month=${r}`)).data?.data||[]}catch(r){console.warn("[Calendar] Failed to load events, rendering empty grid:",r),i=[]}}async function o(){let r=document.getElementById("calendar-grid");if(r){r.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let l=t.getFullYear(),a=t.getMonth(),s=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),k=document.getElementById("cal-month-label");k&&(k.textContent=s);let w=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(f=>f.value)),v=i.filter(f=>w.has(f.type)),b={};v.forEach(f=>{let E=wn(f.event_date);E&&(b[E]||(b[E]=[]),b[E].push(f))});let C=new Date(l,a,1).getDay(),g=new Date(l,a+1,0).getDate(),h=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],d=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';h.forEach(f=>{y+=`<div class="cal-day-header">${f}</div>`});for(let f=0;f<C;f++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let f=1;f<=g;f++){let E=`${l}-${String(a+1).padStart(2,"0")}-${String(f).padStart(2,"0")}`,p=b[E]||[],u=E===d;y+=`
          <div class="cal-cell ${u?"cal-today":""} ${p.length?"cal-has-events":""}"
               data-date="${E}" tabindex="0" role="button" aria-label="${E}">
            <div class="cal-day-num ${u?"today-num":""}">${f}</div>
            <div class="cal-events-preview">
              ${p.slice(0,3).map($=>`
                <div class="cal-event-dot cal-color-${$.color||"gray"}" title="${we($.title||$.type)}">
                  <span class="cal-event-dot-label">${Sn($.title||$.branch_name||$.type,18)}</span>
                </div>
              `).join("")}
              ${p.length>3?`<div class="cal-more">+${p.length-3} lagi</div>`:""}
            </div>
          </div>`}let S=(C+g)%7;if(S!==0)for(let f=0;f<7-S;f++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",r.innerHTML=y,r.querySelectorAll(".cal-cell[data-date]").forEach(f=>{f.addEventListener("click",()=>{let E=f.dataset.date,p=b[E]||[];if(!p.length)return;let u=document.getElementById("cal-event-list"),$=new Date(E+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=$,document.getElementById("cal-event-items").innerHTML=p.map(_=>`
            <div class="cal-event-item cal-color-border-${_.color||"gray"}">
              <div class="cal-event-type">
                ${En(_.type)}
                ${_.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${we(_.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${we(_.title||"-")}</div>
              ${_.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${we(_.branch_name)}</div>`:""}
              ${_.pic||_.period||_.opening_date||_.target_date||_.completion_date||_.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${_.pic?`<div>\u{1F464} <b>PIC:</b> ${we(_.pic)}</div>`:""}
                  ${_.period?`<div>\u{1F4C5} <b>Periode:</b> ${we(_.period)}</div>`:""}
                  ${_.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${we(jt(_.opening_date))}</div>`:""}
                  ${_.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${we(jt(_.target_date))}</div>`:""}
                  ${_.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${we(jt(_.completion_date))}</div>`:""}
                  ${_.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${we(_.notes)}</div>`:""}
                </div>
              `:""}
              ${_.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${_.status==="Done"?"badge-success":"badge-warning"}">${we(_.status)}</span></div>`:""}
              ${_.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${_.days_remaining} hari</div>`:""}
            </div>
          `).join(""),u.style.display="block"})})}catch(l){console.error("[Calendar] Render error:",l),r&&(r.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let c=()=>{n().then(()=>o())};at.on("data:changed",c),o()}function wn(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[n,o,c]=i.map(r=>r.trim());if(c.length===4)return`${c}-${o.padStart(2,"0")}-${n.padStart(2,"0")}`;if(n.length===4)return`${n}-${o.padStart(2,"0")}-${c.padStart(2,"0")}`}return t.slice(0,10)}function jt(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[n,o,c]=i.map(r=>r.trim());if(n.length===4)return`${c.padStart(2,"0")}-${o.padStart(2,"0")}-${n}`;if(c.length===4)return`${n.padStart(2,"0")}-${o.padStart(2,"0")}-${c}`}return e}function Sn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function we(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function En(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}G();async function Fa(e){let t=oe(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${o},${o}99)">
            ${i}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${o}18;color:${o};margin-top:6px">
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
            <span class="info-value" style="color:${o};font-weight:700">${t?.role||"\u2014"}</span>
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
  `;let c=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(c&&r)try{let l=JSON.parse(atob(c.split(".")[1])),a=new Date(l.exp*1e3);r.textContent=`Berakhir: ${a.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let a=document.getElementById("pwd-error"),s=document.getElementById("pwd-success"),k=document.getElementById("btn-save-pwd");a.style.display="none",s.style.display="none";let w=l.target,v=w.current_password.value,b=w.new_password.value,C=w.confirm_password.value;if(b!==C){a.textContent="\u274C Konfirmasi password tidak cocok.",a.style.display="block";return}if(b.length<6){a.textContent="\u274C Password baru minimal 6 karakter.",a.style.display="block";return}k.disabled=!0,k.textContent="\u23F3 Menyimpan...";let g=await L("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:v,new_password:b})});k.disabled=!1,k.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(s.textContent="\u2705 Password berhasil diubah.",s.style.display="block",w.reset(),K("Password berhasil diubah.")):(a.textContent=g.data?.error||"Gagal mengubah password.",a.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}G();var pt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function se(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let o=Number(t);if(o>2e4&&o<99999){let c=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(c.getTime())?null:c.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[o,c,r]=i.map(k=>k.trim()),l=Number(o),a=Number(c),s=Number(r);if(o.length===4&&l>1900)return`${o}-${c.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&s>1900)return l>12?`${r}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:a>12?`${r}-${o.padStart(2,"0")}-${c.padStart(2,"0")}`:`${r}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`;if(r.length===2&&!isNaN(s)){let k=s>=50?`19${r}`:`20${r}`;return l>12?`${k}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:`${k}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`}}let n=new Date(t);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Ta(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Cn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:se(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:se(e["Target Pindah OS"]),target_selesai_os:se(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:se(e["Tanggal Mulai"]),end_date:se(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:se(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:se(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:se(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:se(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:se(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:se(e["Tanggal Target"]||e["Tgl Target"]),completion_date:se(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:se(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:se(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:se(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:se(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:se(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:se(e["Tanggal Back Up"]),completion_date:se(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:se(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:se(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function $n(e,t){let i=pt[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let n=Cn[i.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],c=[],r=[];return t.filter(a=>!Ta(a)).forEach((a,s)=>{let k=t.indexOf(a)+2,w=[];n.required.forEach(({key:b,label:C})=>{let g=a[b];if(g==null||String(g).trim()===""){let h=Object.keys(a).filter(d=>d.trim()).join(", ");w.push({column:C,originalValue:g||"",reason:`Kolom "${C}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${h.slice(0,120)}`})}});let v=n.map(a);w.length>0?c.push({row:k,data:v,raw:a,errors:w}):(o.push(a),r.push(v))}),{valid:o,errors:c,mapped:r}}function Aa(e){let t=[];return e.SheetNames.forEach(i=>{let n=pt[i];if(!n)return;let o=e.Sheets[i],c=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=$n(i,c),l=c.filter(a=>!Ta(a));t.push({sheetName:i,module:n.module,label:n.label,total:l.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),t}function Da(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([n,o])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(o),n)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function La(e){let t=window.XLSX,i=t.utils.book_new(),n=!1;return e.forEach(o=>{if(!o.errors||o.errors.length===0)return;n=!0;let c=o.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(a=>a.column||a).join("; "),"Alasan Error":(l.errors||[]).map(a=>a.reason||a).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([a,s])=>[a,s??""]))})),r=t.utils.json_to_sheet(c);t.utils.book_append_sheet(i,r,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var _n=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function Ba(e){e.innerHTML=`
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
              ${Object.entries(pt).map(([g,{label:h}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${h}</span>`).join("")}
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
  `;let t=null,i=null,n=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(g){Object.entries(o).forEach(([h,d])=>{d.style.display=h===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let h=await L("/api/import/backup");if(h.ok){if(!window.XLSX){j("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let d=window.XLSX,y=d.utils.book_new();Object.entries(h.data.database).forEach(([m,S])=>{let f=S.length>0?S:[{}],E=d.utils.json_to_sheet(f);d.utils.book_append_sheet(y,E,m.substring(0,31))}),d.writeFile(y,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),K("Backup berhasil diunduh!")}else j("Gagal memproses backup: "+(h.data?.error||"Unknown error"))}catch(h){j("Gagal memproses backup: "+h.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let h=await L("/api/sync/google-sheets",{method:"POST"});h.ok?alert("Sinkronisasi Berhasil: "+(h.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(h.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=g,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Da(),K("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),a=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),l.click()}),l.addEventListener("change",g=>{g.target.files[0]&&s(g.target.files[0])}),a.addEventListener("dragover",g=>{g.preventDefault(),a.classList.add("drag-over")}),a.addEventListener("dragleave",()=>a.classList.remove("drag-over")),a.addEventListener("drop",g=>{g.preventDefault(),a.classList.remove("drag-over");let h=g.dataTransfer.files[0];h&&h.name.match(/\.xlsx?$/i)?s(h):j("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,l.value="",document.getElementById("file-info").style.display="none",a.style.display="",c("upload")});async function s(g){t=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",a.style.display="none",await k(g)}async function k(g){c("validating");let h=document.getElementById("validation-status"),d=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");h.textContent="Membaca file Excel...",d.style.width="20%",await Ve(200);let y=await g.arrayBuffer(),m=window.XLSX.read(y,{type:"array",cellDates:!0});h.textContent=`Memvalidasi ${m.SheetNames.length} sheet...`,d.style.width="50%",await Ve(100),i=Aa(m),d.style.width="100%",h.textContent="Validasi selesai!",await Ve(300),w()}catch(y){c("upload"),j("Gagal memproses file: "+y.message),document.getElementById("file-info").style.display="flex",a.style.display="none"}}function w(){c("preview");let g=i.filter(u=>!u.skipped).length,h=i.reduce((u,$)=>u+$.total,0),d=i.reduce((u,$)=>u+$.valid,0),y=i.reduce((u,$)=>u+$.errorCount,0),m=h>0?Math.round(d/h*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${h} baris</span>
      <span class="badge badge-success">${d} valid (${m}%)</span>
      ${y>0?`<span class="badge badge-danger">${y} error</span>`:""}
    `;let S=document.getElementById("preview-table-container");S.innerHTML=`
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
          ${i.map((u,$)=>`
            <tr class="${u.errorCount>0?"row-error":u.skipped?"row-skipped":"row-ok"}">
              <td><strong>${u.sheetName}</strong></td>
              <td>${u.label}</td>
              <td style="text-align:center">${u.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${u.valid}</span></td>
              <td style="text-align:center">${u.errorCount>0?`<span class="badge badge-danger">${u.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${u.skipped?'<span class="badge badge-neutral">Dilewati</span>':u.errorCount>0&&u.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':u.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':u.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${u.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${$}">\u{1F50D} ${u.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(u=>{u.addEventListener("click",()=>{let $=i[Number(u.dataset.idx)];v($)})});let f=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");E.innerHTML="",f.style.display="none";let p=document.getElementById("btn-start-import");d===0?(p.disabled=!0,p.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(p.disabled=!1,y>0?(p.innerHTML=`\u{1F680} Import ${d} Data Valid (${y} dilewati)`,p.title="Baris error akan dilewati, baris valid tetap diimport"):p.innerHTML=`\u{1F680} Mulai Import ${d} Data`)}function v(g){let h=document.getElementById("error-detail-section"),d=document.getElementById("error-detail-container");h.style.display="";let y=g.errors.slice(0,100).map(m=>(Array.isArray(m.errors)?m.errors:[]).map(f=>{let E=typeof f=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${m.row}</span></td>
            <td><strong>${E?f.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${E&&f.originalValue!==void 0?f.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${E?f.reason:f}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${E&&f.aliases?`Gunakan salah satu nama kolom:<br><em>${f.aliases}</em>`:E&&f.hint?f.hint:""}
            </td>
          </tr>
        `}).join("")).join("");d.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${g.sheetName} \u2014 ${g.errorCount} baris error dari ${g.total} total
          ${g.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${y||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${g.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,h.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",a.style.display="",t=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;La(i)?K("Log error berhasil didownload."):K("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";b(g)});async function b(g){c("importing"),n=Date.now();let h=[];_n.forEach(f=>{let E=i?.find(p=>p.module===f&&p.mapped?.length>0);E&&h.push(E)});let d=document.getElementById("import-steps-list");d.innerHTML=h.map(f=>`
      <div class="import-step-item" id="step-item-${f.module}">
        <span class="step-item-icon" id="step-icon-${f.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${f.label} <span class="step-item-count">(${f.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${f.module}"></span>
      </div>
    `).join("");let y=document.getElementById("import-bar"),m=document.getElementById("import-current-status"),S={totalSheets:h.length,totalRows:h.reduce((f,E)=>f+E.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let f=0;f<h.length;f++){let E=h[f],p=document.getElementById(`step-icon-${E.module}`),u=document.getElementById(`step-status-${E.module}`);p.textContent="\u{1F504}",u.textContent="Mengimport...",m.textContent=`Mengimport ${E.label}...`,y.style.width=`${Math.round(f/h.length*100)}%`;try{let $=await L(`/api/import/${E.module}`,{method:"POST",body:JSON.stringify({rows:E.mapped,onDuplicate:g})});if($.ok){let _=$.data;S.inserted+=_.inserted||0,S.skipped+=_.skipped||0,S.moduleResults.push({label:E.label,inserted:_.inserted||0,skipped:_.skipped||0,status:"ok"}),p.textContent="\u2705",u.innerHTML=`<span class="badge badge-success">${_.inserted||0} berhasil</span>${_.skipped>0?` <span class="badge badge-neutral">${_.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:$.data?.error}),p.textContent="\u274C",u.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch($){S.failed++,S.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:$.message}),p.textContent="\u274C",u.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ve(150)}y.style.width="100%",m.textContent="Selesai!",await Ve(400),C(S)}function C(g){c("summary");let h=((Date.now()-n)/1e3).toFixed(1),d=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${d?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${d?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${g.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${g.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${g.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${g.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${g.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${g.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${h}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${g.moduleResults.map(y=>`
            <tr>
              <td>${y.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${y.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${y.skipped}</span></td>
              <td style="text-align:center">
                ${y.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${y.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,l.value="",document.getElementById("file-info").style.display="none",a.style.display="",c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ve(e){return new Promise(t=>setTimeout(t,e))}G();var ut=[],Pa=[];async function Ia(e){ut=await W(),Pa=await le(),J({container:e,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:t=>t?`<span class="badge badge-info">${t}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:t=>`<span class="badge badge-warning">${t||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ut}],exportOptions:{moduleName:"sp_data",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await L(`/api/sp?limit=10000&${i}`);if(n.ok){let o=n.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(U(),be));c(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(U(),be));i(t,"Template_Import_SP")},onImport:async t=>{let i=r=>{if(!r)return null;let l=String(r||"").toLowerCase(),a=ut.find(s=>String(s.label||"").toLowerCase()===l);return a?a.value:null},n=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let l=String(r).trim();if(/^\d{4,5}$/.test(l)){let s=Number(l);if(s>2e4&&s<99999){let k=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let a=l.split(/[\/\-\.]/);if(a.length===3){let[s,k,w]=a.map(v=>v.trim());if(s.length===4&&k.length<=2&&w.length<=2)return`${s}-${k.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&k.length<=2&&s.length<=2)return`${w}-${k.padStart(2,"0")}-${s.padStart(2,"0")}`}return l},o=t.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:i(String(r.Cabang||"").trim()),tanggal:n(r["Tanggal Sp"]),akhir_sp:n(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),c=await L("/api/import/sp",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:Pa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:ut,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}G();var Ne=[],Ma=[];async function Na(e){Ne=await W(),Ma=await le(),J({container:e,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:t=>`<span class="badge ${t==="Selesai"?"badge-success":"badge-warning"}">${t||"-"}</span>`},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Ne},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Ne}],exportOptions:{moduleName:"mutasi_data",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await L(`/api/mutasi?limit=10000&${i}`);if(n.ok){let o=n.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(U(),be));c(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(U(),be));i(t,"Template_Import_Mutasi")},onImport:async t=>{let i=r=>{if(!r)return null;let l=String(r||"").toLowerCase(),a=Ne.find(s=>String(s.label||"").toLowerCase()===l);return a?a.value:null},n=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let l=String(r).trim();if(/^\d{4,5}$/.test(l)){let s=Number(l);if(s>2e4&&s<99999){let k=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let a=l.split(/[\/\-\.]/);if(a.length===3){let[s,k,w]=a.map(v=>v.trim());if(s.length===4&&k.length<=2&&w.length<=2)return`${s}-${k.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&k.length<=2&&s.length<=2)return`${w}-${k.padStart(2,"0")}-${s.padStart(2,"0")}`}return l},o=t.map(r=>({tanggal:n(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:i(String(r["Cabang Asal"]||"").trim()),to_branch_id:i(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),c=await L("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:Ma},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Ne,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Ne,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}G();var Rt=[],Ka=[];async function za(e){Rt=await W(),Ka=await Fe(),J({container:e,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:t=>t!=null?`${t} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:t=>t!=null?`${t} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Rt}],exportOptions:{moduleName:"overtime_records",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await L(`/api/overtime?limit=10000&${i}`);if(n.ok){let o=n.data.data.map(r=>({Tanggal:r.date||"","Cabang Lembur":r.branch_name||"","Nama Karyawan":r.employee_name||"","Jam Mulai":r.start_time||"","Jam Selesai":r.end_time||"","Istirahat (Jam)":r.break_hours||0,"Total Jam":r.total_hours||0,"Alasan / Ket Lembur":r.reason||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(U(),be));c(o,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:Rt},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:Ka},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",e=>{if(e.target.name==="start_time"||e.target.name==="end_time"||e.target.name==="break_hours"){let t=e.target.closest("form");if(!t)return;let i=t.querySelector('[name="start_time"]'),n=t.querySelector('[name="end_time"]'),o=t.querySelector('[name="break_hours"]'),c=t.querySelector('[name="total_hours"]');if(!i||!n||!o||!c)return;let r=i.value,l=n.value,a=parseFloat(o.value)||0;if(r&&l){let[s,k]=r.split(":").map(Number),[w,v]=l.split(":").map(Number),b=w*60+v-(s*60+k);b<0&&(b+=24*60);let C=b/60-a;C<0&&(C=0),c.value=Number.isInteger(C)?C:C.toFixed(2)}}}));G();async function Oa(e){let t=oe();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}J({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,n)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${n.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,n)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let c=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(s=>String(s.id)===String(i));if(!c)return alert("Data tidak ditemukan");let r=s=>{if(!s)return"Tidak ada data";try{return JSON.stringify(JSON.parse(s),null,2)}catch{return s}},l=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${r(c.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${r(c.new_data)}</pre>
           </div>
         </div>
       `,{createModal:a}=await Promise.resolve().then(()=>(ue(),vt));a({title:`Detail Audit Log #${i}`,content:l,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}G();ue();U();var qt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},We={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function ja(e){e.innerHTML=`
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
  `;let t="board",i=null,n=[],o=oe(),c=o&&["superadmin","admin","manager","spv"].includes(o.role);L("/api/branches?all=1").then(d=>{d.ok&&d.data?.data&&(n=d.data.data)}).catch(()=>{});async function r(){try{let d=await L("/api/os-migrations/summary");if(!d.ok)throw new Error("Gagal memuat data migrasi");i=d.data?.data||d.data,l(i),s()}catch(d){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${d.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(d){let y=document.getElementById("migration-summary-cards");if(!y||!d)return;let m=d.total_branches||0,S=d.vendors?.AEON?.total||0,f=d.vendors?.["Triple S"]?.total||0,E=d.by_status?.completed||0,p=d.by_status?.in_progress||0,u=d.by_status?.scheduled||0,$=m>0?Math.round(E/m*100):0;y.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${m} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${S} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${f} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${$}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${$}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${E}</strong></span>
          <span>\u{1F535} Proses: <strong>${p}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${u}</strong></span>
        </div>
      </div>
    `}function a(){if(!i||!i.matrix)return[];let d=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),y=document.getElementById("filter-vendor")?.value||"",m=document.getElementById("filter-month")?.value||"",S=document.getElementById("filter-status")?.value||"";return i.matrix.filter(f=>!(y&&f.vendor!==y||m&&f.target_month!==m||S&&f.status!==S||d&&!`${f.branch_code} ${f.branch_name} ${f.branch_full_name} ${f.vendor} ${f.notes||""} ${f.pic_vendor||""}`.toLowerCase().includes(d)))}function s(){let d=a(),y=document.getElementById("filter-stats-text");y&&(y.innerHTML=`Menampilkan <strong>${d.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let m=document.getElementById("migration-view-container");t==="board"?k(m,d):w(m,d)}function k(d,y){if(!y.length){d.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let m=["AEON","Triple S"],S=["2026-09","2026-10","2026-11","2026-12"],f="";m.forEach(E=>{let p=y.filter(M=>M.vendor===E);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==E||p.length===0)return;let u=p.length,$=E==="AEON"?"#1E40AF":"#0F766E",_=E==="AEON"?"#EFF6FF":"#F0FDFA",P=E==="AEON"?"#BFDBFE":"#99F6E4",N=S.filter(M=>p.some(F=>F.target_month===M));f+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${_}; border-bottom:1px solid ${P}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${$};">${E==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${$}; border:1px solid ${P}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${u} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${N.map(M=>{let F=p.filter(I=>I.target_month===M);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(qt[M]||M).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${F.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${F.map(I=>{let x=We[I.status]||We.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${I.branch_code}.</span>${I.branch_name}
                            </div>
                            ${c?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${I.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${I.id}" data-name="${I.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${I.id}" data-status="${I.status}" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${x.icon} ${x.label} \u25BE
                              </button>
                            </div>

                            ${I.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${I.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${I.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${I.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${I.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),d.innerHTML=f,v(d)}function w(d,y){if(!y.length){d.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}d.innerHTML=`
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
                ${c?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${y.map((m,S)=>{let f=We[m.status]||We.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${S+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${m.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${m.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${m.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${m.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${qt[m.target_month]||m.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${m.id}" data-status="${m.status}" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${f.icon} ${f.label} \u25BE
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
                    ${c?`
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
    `,v(d)}function v(d){d.querySelectorAll(".status-toggle-btn").forEach(y=>{y.addEventListener("click",m=>{m.stopPropagation();let S=y.dataset.id,f=y.dataset.status;b(y,S,f)})}),d.querySelectorAll(".edit-card-btn").forEach(y=>{y.addEventListener("click",()=>{let m=y.dataset.id,S=i?.matrix?.find(f=>String(f.id)===String(m));S&&C(S)})}),d.querySelectorAll(".delete-card-btn").forEach(y=>{y.addEventListener("click",()=>{let m=y.dataset.id,S=y.dataset.name||"Cabang";ie(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${S}</strong>?`,async()=>{try{if(!(await L(`/api/os-migrations/${m}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");K(`Jadwal migrasi ${S} berhasil dihapus`),r()}catch(f){j(f.message)}})})})}function b(d,y,m){document.querySelectorAll(".status-popup-menu").forEach(u=>u.remove());let S=d.getBoundingClientRect(),f=document.createElement("div");f.className="status-popup-menu",f.style.cssText=`
      position: fixed;
      top: ${S.bottom+window.scrollY+4}px;
      left: ${S.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let E=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];f.innerHTML=E.map(u=>`
      <button class="status-option-item" data-key="${u.key}" style="
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
        color: ${u.key===m?"#2563EB":"var(--text-1)"};
        font-weight: ${u.key===m?"700":"500"};
      ">
        ${u.label} ${u.key===m?"\u2713":""}
      </button>
    `).join("");let p=u=>{!f.contains(u.target)&&u.target!==d&&(f.remove(),document.removeEventListener("click",p))};setTimeout(()=>document.addEventListener("click",p),50),f.querySelectorAll(".status-option-item").forEach(u=>{u.addEventListener("click",async()=>{let $=u.dataset.key;if(f.remove(),document.removeEventListener("click",p),$!==m)try{if(!(await L(`/api/os-migrations/${y}/status`,{method:"PATCH",body:JSON.stringify({status:$})})).ok)throw new Error("Gagal memperbarui status");K("Status berhasil diubah"),r()}catch(_){j(_.message)}})}),document.body.appendChild(f)}function C(d=null){let y=!!d,m=y?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",f=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${n.map(E=>`
      <option value="${E.id}" ${d&&d.branch_id===E.id?"selected":""}>
        ${E.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${d?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${d?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${d&&!["AEON","Triple S"].includes(d.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${d?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${d?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${d?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${d?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${d?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${d?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${d?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${d?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${d?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${d?.notes||""}</textarea>
        </div>
      </form>
    `;Z({title:m,content:f,confirmText:y?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(E,p)=>{let u=E.querySelector("#migration-form");if(!u.checkValidity()){u.reportValidity();return}let $=new FormData(u),_={branch_id:parseInt($.get("branch_id"),10),vendor:$.get("vendor"),target_month:$.get("target_month"),status:$.get("status"),actual_date:$.get("actual_date")||null,pic_vendor:$.get("pic_vendor")||null,notes:$.get("notes")||null};try{let P=y?`/api/os-migrations/${d.id}`:"/api/os-migrations",M=await L(P,{method:y?"PUT":"POST",body:JSON.stringify(_)});if(!M.ok){let F=await M.json().catch(()=>({}));throw new Error(F.error||"Gagal menyimpan data")}K(y?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),p(),r()}catch(P){j(P.message)}}})}let g=document.getElementById("view-board-btn"),h=document.getElementById("view-table-btn");g.addEventListener("click",()=>{t="board",g.style.background="#fff",g.classList.remove("btn-ghost"),h.style.background="transparent",h.classList.add("btn-ghost"),s()}),h.addEventListener("click",()=>{t="table",h.style.background="#fff",h.classList.remove("btn-ghost"),g.style.background="transparent",g.classList.add("btn-ghost"),s()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(d=>{let y=document.getElementById(d);y&&(y.addEventListener("input",s),y.addEventListener("change",s))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{C()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let d=a();if(!d.length){j("Tidak ada data untuk diexport");return}let y=d.map((m,S)=>({No:S+1,"Kode Cabang":m.branch_code,"Nama Cabang":m.branch_name,"Cabang Lengkap":m.branch_full_name,"Vendor OS":m.vendor,"Target Bulan":qt[m.target_month]||m.target_month,Status:We[m.status]?.label||m.status,"Tanggal Realisasi":m.actual_date||"","Jml Personil FC":m.employee_count||0,"PIC Vendor":m.pic_vendor||"",Catatan:m.notes||""}));R(y,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),K("File Excel berhasil diunduh")}),r()}G();ue();U();var Ra={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function qa(e){e.innerHTML=`
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
  `;let t="pagi",i="timeline",n=[],o=oe(),c=o&&["superadmin","admin","manager","spv"].includes(o.role);async function r(){try{let d=await L("/api/dac?all=1");if(!d.ok)throw new Error("Gagal memuat standar DAC");n=d.data?.data||d.data||[];let y=n.filter(E=>E.shift==="pagi").length,m=n.filter(E=>E.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=y,document.getElementById("badge-count-siang").textContent=m,document.getElementById("badge-count-all").textContent=n.length;let S=document.getElementById("dac-category-filter"),f=Array.from(new Set(n.map(E=>E.category))).sort();S.innerHTML='<option value="">Semua Kategori</option>'+f.map(E=>`<option value="${E}">${E}</option>`).join(""),a(),s()}catch(d){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${d.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(){let d=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),y=document.getElementById("dac-category-filter")?.value||"",m=document.getElementById("dac-period-filter")?.value||"";return n.filter(S=>{if(t!=="all"&&S.shift!==t||y&&S.category!==y)return!1;if(m==="morning"){let f=parseInt(S.start_time.split(":")[0],10);if(S.shift==="pagi"&&f>=12||S.shift==="siang"&&f>=17)return!1}else if(m==="afternoon"){let f=parseInt(S.start_time.split(":")[0],10);if(S.shift==="pagi"&&f<12||S.shift==="siang"&&f<17)return!1}return!(d&&!`${S.start_time} ${S.end_time} ${S.category} ${S.location} ${S.activity}`.toLowerCase().includes(d))})}function a(){let d=document.getElementById("dac-summary-cards");if(!d)return;let y=n.filter(p=>p.shift==="pagi"),m=n.filter(p=>p.shift==="siang"),S=t==="pagi"?y:t==="siang"?m:n,f=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",E=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";d.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${E})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${S.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${f}</div>
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
    `}function s(){let d=l(),y=document.getElementById("dac-stats-text");y&&(y.innerHTML=`Menampilkan <strong>${d.length}</strong> aktivitas`);let m=document.getElementById("dac-view-container");i==="timeline"?k(m,d):w(m,d)}function k(d,y){if(!y.length){d.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}d.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${y.map(m=>{let S=Ra[m.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},f=m.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",E=m.shift==="pagi"?"#FEF3C7":"#E0E7FF",p=m.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${m.start_time} - ${m.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${m.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${E}; color:${p}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${f}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${S.bg}; color:${S.color}; border:1px solid ${S.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${m.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${m.location}
                    </span>
                  </div>
                  ${c?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${m.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${m.id}" data-name="${m.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${S.color}; margin-top:8px;">
                  ${m.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,v(d)}function w(d,y){if(!y.length){d.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}d.innerHTML=`
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
                ${c?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${y.map(m=>{let S=Ra[m.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${m.item_no}</td>
                    ${t==="all"?`
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
                      <span class="badge" style="background:${S.bg}; color:${S.color}; border:1px solid ${S.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${m.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${m.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${m.activity}
                    </td>
                    ${c?`
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
    `,v(d)}function v(d){d.querySelectorAll(".edit-dac-btn").forEach(y=>{y.addEventListener("click",()=>{let m=y.dataset.id,S=n.find(f=>String(f.id)===String(m));S&&b(S)})}),d.querySelectorAll(".delete-dac-btn").forEach(y=>{y.addEventListener("click",()=>{let m=y.dataset.id,S=y.dataset.name||"item";ie(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${S}</strong>?`,async()=>{try{if(!(await L(`/api/dac/${m}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");K("Aktivitas DAC berhasil dihapus"),r()}catch(f){j(f.message)}})})})}function b(d=null){let y=!!d,m=y?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",S=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],f=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${d?.shift==="pagi"||!d&&t==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${d?.shift==="siang"||!d&&t==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${d?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${d?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${d?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${d?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${S.map(E=>`<option value="${E}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${d?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${d?.activity||""}</textarea>
        </div>
      </form>
    `;Z({title:m,content:f,confirmText:y?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(E,p)=>{let u=E.querySelector("#dac-form");if(!u.checkValidity()){u.reportValidity();return}let $=new FormData(u),_={shift:$.get("shift"),item_no:$.get("item_no")?parseInt($.get("item_no"),10):void 0,start_time:$.get("start_time"),end_time:$.get("end_time"),category:$.get("category"),location:$.get("location"),activity:$.get("activity")};try{let P=y?`/api/dac/${d.id}`:"/api/dac",M=await L(P,{method:y?"PUT":"POST",body:JSON.stringify(_)});if(!M.ok){let F=await M.json().catch(()=>({}));throw new Error(F.error||"Gagal menyimpan aktivitas")}K(y?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),p(),r()}catch(P){j(P.message)}}})}function C(){let d=l();if(!d.length){j("Tidak ada data untuk dicetak");return}let y=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",m=window.open("","_blank");if(!m){j("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}m.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${y}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${y}</div>
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
            ${d.map(S=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${S.item_no}</td>
                ${t==="all"?`<td>${S.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${S.start_time} - ${S.end_time}</td>
                <td><strong>${S.category}</strong></td>
                <td><strong>${S.location}</strong></td>
                <td>${S.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),m.document.close(),m.focus(),setTimeout(()=>{m.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(d=>{d.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(y=>{y.style.background="#F1F5F9",y.style.color="var(--text-2)"}),d.style.background="#2563EB",d.style.color="#fff",t=d.dataset.shift,a(),s()})});let g=document.getElementById("view-timeline-btn"),h=document.getElementById("view-table-btn");g.addEventListener("click",()=>{i="timeline",g.style.background="#fff",g.classList.remove("btn-ghost"),h.style.background="transparent",h.classList.add("btn-ghost"),s()}),h.addEventListener("click",()=>{i="table",h.style.background="#fff",h.classList.remove("btn-ghost"),g.style.background="transparent",g.classList.add("btn-ghost"),s()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(d=>{let y=document.getElementById(d);y&&(y.addEventListener("input",s),y.addEventListener("change",s))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{b()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{C()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let d=l();if(!d.length){j("Tidak ada data untuk diexport");return}let y=d.map(m=>({No:m.item_no,Shift:m.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":m.start_time,"Jam Selesai":m.end_time,Kategori:m.category,"Lokasi / Area":m.location,"Rincian Aktivitas":m.activity}));R(y,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),K("File Excel berhasil diunduh")}),r()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],n=parseInt(t[2],10),o=parseInt(t[3],10);if(n>12&&o<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let n=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(i[2],10),c=n[parseInt(i[1],10)-1];return`${o} ${c} ${i[0]}`}return t};function ee(e){return async t=>{if(!Oe()){_e("/login");return}return e(t)}}var Qe=null;function Fn(){Qe&&clearInterval(Qe);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),c=document.getElementById("header-clock-date");o&&(o.textContent=i),c&&(c.textContent=n)};e(),Qe=setInterval(e,1e3)}async function Tn(){try{let e=await L("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(n,o)=>{let c=document.getElementById(n);c&&(c.textContent=o>0?o:"",c.style.display=o>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var Ke=[];async function An(){try{let e=await L("/api/dashboard/notifications");if(!e.ok)return;Ke=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=Ke.length>0?"block":"none",t.textContent=Ke.length)}catch{}}function Dn(){if(!Ke.length){Z({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ke.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;Z({title:`Notifikasi (${Ke.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function Ha(){let e=oe(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let s=new Date().getHours();return s>=4&&s<11?"Selamat Pagi":s>=11&&s<15?"Selamat Siang":s>=15&&s<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),c=document.getElementById("sidebar-close"),r=()=>{i.classList.add("open"),n.classList.add("show")},l=()=>{i.classList.remove("open"),n.classList.remove("show")};o?.addEventListener("click",r),c?.addEventListener("click",l),n?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(s=>s.addEventListener("click",l)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(s=>{s.dataset.route!=="/overtime"&&(s.style.display="none")}),document.querySelectorAll(".nav-section").forEach(s=>{Array.from(s.querySelectorAll(".nav-item")).some(w=>w.style.display!=="none")||(s.style.display="none")}));function a(){let s=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(v=>{let b=v.dataset.route;v.classList.toggle("active",s===b||b!=="/dashboard"&&s.startsWith(b))});let k=document.getElementById("topbar-title"),w=document.querySelector(".nav-item.active .nav-label");k&&w&&(k.textContent=w.textContent)}window.addEventListener("hashchange",a),a(),Fn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await L("/api/auth/logout",{method:"POST"}),je(),Qe&&clearInterval(Qe),_e("/login")}),Tn(),An(),document.getElementById("btn-notif")?.addEventListener("click",s=>{s.preventDefault(),Dn()})}async function Ln(){X("/login",({main:t})=>ea(t)),X("/dashboard",ee(({main:t})=>Wt(t))),X("/calendar",ee(({main:t})=>_a(t))),X("/employees",ee(({main:t,params:i})=>ta(t,i))),X("/contracts",ee(({main:t,params:i})=>lt(t,i))),X("/sp",ee(({main:t})=>Ia(t))),X("/mutasi",ee(({main:t})=>Na(t))),X("/overtime",ee(({main:t})=>za(t))),X("/timeline",ee(({main:t,params:i})=>na(t,i))),X("/os-migrations",ee(({main:t})=>ja(t))),X("/issues",ee(({main:t,params:i})=>ia(t,i))),X("/one-on-one",ee(({main:t,params:i})=>ra(t,i))),X("/training",ee(({main:t})=>oa(t))),X("/relievers",ee(({main:t,params:i})=>la(t,i))),X("/reports/inspection",ee(({main:t})=>da(t))),X("/reports/cleaning",ee(({main:t})=>ca(t))),X("/reports/fogging",ee(({main:t})=>pa(t))),X("/reports/basecamp",ee(({main:t})=>ua(t))),X("/reports/supply",ee(({main:t})=>Ea(t,"supply"))),X("/sop",()=>{window.location.hash="#/sop-fc"}),X("/sop-fc",ee(({main:t})=>ha(t))),X("/sop-security",ee(({main:t})=>va(t))),X("/checklist",ee(({main:t})=>wa(t))),X("/hygiene-standards",ee(({main:t})=>Sa(t))),X("/dac",ee(({main:t})=>qa(t))),X("/users",ee(({main:t})=>Ca(t))),X("/branches",ee(({main:t})=>$a(t))),X("/profile",ee(({main:t})=>Fa(t))),X("/settings/import",ee(({main:t})=>Ba(t))),X("/audit-logs",ee(({main:t})=>Oa(t)));let e=Oe();if(!e&&window.location.hash!=="#/login"&&_e("/login"),e){let t=await L("/api/auth/me");if(t.ok){Re(t.data.data),Ha();let i=oe(),n=window.location.hash;i&&i.role==="input_lembur"&&(n==="#/dashboard"||n===""||n==="#/")&&_e("/overtime")}else je(),_e("/login")}window.addEventListener("fm:login",()=>{Ha();let t=oe();t&&t.role==="input_lembur"?_e("/overtime"):_e("/dashboard")}),Gt()}Ln();
