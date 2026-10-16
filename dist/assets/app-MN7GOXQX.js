var Ga=Object.defineProperty;var bt=(e,t)=>()=>(e&&(t=e(e=0)),t);var ht=(e,t)=>{for(var i in t)Ga(e,i,{get:t[i],enumerable:!0})};var Gt={};ht(Gt,{API:()=>Ut,CLIENT_SIDE_MAX_ROWS:()=>Te,IS_DEVELOPMENT:()=>je,apiFetch:()=>B,clearToken:()=>qe,getToken:()=>Re,getUser:()=>oe,setToken:()=>ft,setUser:()=>He});function Re(){return localStorage.getItem("fm_token")}function ft(e){localStorage.setItem("fm_token",e)}function qe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function oe(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function He(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function B(e,t={}){let i=Re(),n={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let r=`cb=${Date.now()}`,m=e.includes("?")?"&":"?",o=`${Ut}${e}${m}${r}`,l=await fetch(o,{...t,headers:n}),a;try{let d=await l.text();try{a=JSON.parse(d)}catch{a={error:`Server Error (${l.status}): ${d.substring(0,80)}...`}}}catch{a={error:"Gagal membaca respon dari server"}}return l.status===401&&(qe(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:a}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var je,Va,Ut,Te,U=bt(()=>{je=!1,Va="https://fm-operations-api.facilitycare-audydental.workers.dev",Ut=Va,Te=1e4});var kt={};ht(kt,{confirmDialog:()=>ae,createModal:()=>Z});function Z({title:e,content:t,onConfirm:i,onCancel:n,confirmText:r="Simpan",cancelText:m="Batal",size:o="md",confirmClass:l="btn-primary"}){let a={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
    <div class="modal" style="max-width:${a[o]||a.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${m}</button>
        ${i?`<button class="btn ${l} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(t);let S=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),S()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),S()}),i&&d.querySelector(".modal-confirm").addEventListener("click",()=>i(d,S)),d.addEventListener("click",E=>{E.target===d&&(n&&n(),S())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:S}}function ae(e,t,i="Konfirmasi"){return Z({title:i,content:`<p>${e}</p>`,onConfirm:(n,r)=>{t(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ue=bt(()=>{});var he={};ht(he,{downloadExcel:()=>q,parseExcel:()=>Ge,renderExcelButtons:()=>Qa});function Ge(e){return new Promise((t,i)=>{let n=new FileReader;n.onload=r=>{try{let m=new Uint8Array(r.target.result),o=XLSX.read(m,{type:"array"}),l=o.SheetNames[0],a=o.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let d=XLSX.utils.decode_range(a["!ref"]||"A1:A1"),S=d.e.r-d.s.r+1,E=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${S}`),console.log(`Total Columns: ${E}`);let w=[];for(let C=d.s.c;C<=d.e.c;++C){let b=a[XLSX.utils.encode_cell({c:C,r:d.s.r})];b&&b.v&&w.push(b.v)}console.log(`Headers Found: ${w.join(", ")}`),console.log("---------------------------");let h=XLSX.utils.sheet_to_json(a,{defval:""});Object.defineProperty(h,"__worksheet",{value:a,enumerable:!1}),Object.defineProperty(h,"__headers",{value:w,enumerable:!1}),t(h)}catch(m){i(m)}},n.onerror=r=>i(r),n.readAsArrayBuffer(e)})}function q(e,t){try{let i=XLSX.utils.json_to_sheet(e),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,i,"Data"),XLSX.writeFile(n,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Qa(e){return`
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
  `}var Y=bt(()=>{});U();var yt={},Ze=null;function X(e,t){yt[e]=t}function Fe(e){window.location.hash=e}function Vt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...n]=t.split("?"),r=yt[i];if(!r){for(let[o,l]of Object.entries(yt))if(o.endsWith("/*")&&i.startsWith(o.slice(0,-2))){r=l;break}}Ze&&(Ze(),Ze=null);let m=document.getElementById("main-content");if(m&&(m.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let o=new URLSearchParams(n.join("?")),l=i.split("/").filter(Boolean),a=await r({path:i,params:o,segments:l,main:m});a&&(Ze=a)}else{let o=m||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var Je;function Ya(){return Je||(Je=document.createElement("div"),Je.id="toast-container",document.body.appendChild(Je)),Je}function xt(e,t="info",i=3500){let n=Ya(),r=document.createElement("div");r.className=`toast toast-${t}`;let m={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${m[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,n.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var z=e=>xt(e,"success"),O=e=>xt(e,"error");var vt=e=>xt(e,"info");ue();U();U();ue();function et({columns:e,data:t,onEdit:i,onDelete:n,onView:r,actions:m=[],emptyText:o="Tidak ada data",bulkSelect:l=null}){let a=document.createElement("div");if(a.className="table-wrapper",!t||t.length===0)return a.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,a;let d=document.createElement("table");d.className="data-table";let S=document.createElement("thead"),E=document.createElement("tr");if(l){let h=document.createElement("th");h.style.width="40px",h.style.textAlign="center";let C=document.createElement("input");C.type="checkbox",C.id="select-all-checkbox",C.title="Pilih semua",C.addEventListener("change",()=>{t.forEach(b=>{C.checked?l.selectedIds.add(b.id):l.selectedIds.delete(b.id)}),a.querySelectorAll(".row-checkbox").forEach(b=>b.checked=C.checked),l.onToggle()}),h.appendChild(C),E.appendChild(h)}if(e.forEach(h=>{let C=document.createElement("th");C.textContent=h.label,h.width&&(C.style.width=h.width),E.appendChild(C)}),i||n||r||m.length>0){let h=document.createElement("th");h.textContent="Aksi",h.style.width="120px",E.appendChild(h)}S.appendChild(E),d.appendChild(S);let w=document.createElement("tbody");return t.forEach(h=>{let C=document.createElement("tr");if(l){let b=document.createElement("td");b.style.textAlign="center",b.style.width="40px";let f=document.createElement("input");f.type="checkbox",f.className="row-checkbox",f.checked=l.selectedIds.has(h.id),f.addEventListener("change",()=>{if(f.checked)l.selectedIds.add(h.id);else{l.selectedIds.delete(h.id);let c=document.getElementById("select-all-checkbox");c&&(c.checked=!1)}l.onToggle()}),b.appendChild(f),C.appendChild(b)}if(e.forEach(b=>{let f=document.createElement("td");if(b.render){let c=b.render(h[b.key],h);c instanceof HTMLElement?f.appendChild(c):f.innerHTML=c||""}else f.textContent=h[b.key]!==null&&h[b.key]!==void 0&&h[b.key]!==""?h[b.key]:"";b.nowrap&&(f.style.whiteSpace="nowrap"),C.appendChild(f)}),i||n||r||m.length>0){let b=document.createElement("td");b.className="actions-cell";let f=document.createElement("div");if(f.className="btn-group",r){let c=document.createElement("button");c.className="btn btn-xs btn-ghost",c.innerHTML="\u{1F441}",c.title="Lihat",c.addEventListener("click",()=>r(h)),f.appendChild(c)}if(i){let c=document.createElement("button");c.className="btn btn-xs btn-secondary",c.innerHTML="\u270F\uFE0F",c.title="Edit",c.addEventListener("click",()=>i(h)),f.appendChild(c)}m.forEach(c=>{let v=document.createElement("button");v.className=`btn btn-xs ${c.class||"btn-ghost"}`,v.innerHTML=c.icon||c.label,v.title=c.label,v.addEventListener("click",()=>c.handler(h)),f.appendChild(v)}),b.appendChild(f),C.appendChild(b)}w.appendChild(C)}),d.appendChild(w),a.appendChild(d),a}function tt({page:e,pages:t,total:i,limit:n,onPage:r}){if(t<=1)return null;let m=document.createElement("div");m.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${i} data`,m.appendChild(o);let l=document.createElement("div");l.className="pagination-btns";let a=(E,w,h=!1,C=!1)=>{let b=document.createElement("button");b.className=`btn btn-sm ${C?"btn-primary":"btn-ghost"} pagination-btn`,b.textContent=E,b.disabled=h,b.addEventListener("click",()=>r(w)),l.appendChild(b)};a("\xAB",1,e===1),a("\u2039",e-1,e===1);let d=Math.max(1,e-2),S=Math.min(t,e+2);for(let E=d;E<=S;E++)a(E,E,!1,E===e);return a("\u203A",e+1,e===t),a("\xBB",t,e===t),m.appendChild(l),m}ue();function Ue(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Ue(t.fields)}</div>`;let i=t.required?"required":"",n=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(t.type){case"textarea":r=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let o=(t.options||[]).map(h=>{let C=typeof h=="object"?h.value:h,b=typeof h=="object"?h.label:h,f=t.value==C?"selected":"";return`<option value="${C}" ${f}>${b}</option>`}).join("");r=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${o}</select>`;break;case"combobox":let l=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,a=(t.options||[]).map(h=>{let C=typeof h=="object"?h.value:h,b=typeof h=="object"?h.label||h.value||"":h||"";return(b==="undefined"||b==="[object Object]"||b==="null")&&(b=""),b?`<option value="${b}"></option>`:""}).join(""),d=t.value||"";if(t.value){let h=(t.options||[]).find(C=>(typeof C=="object"?C.value:C)==t.value);if(h){let C=typeof h=="object"?h.label||h.value||"":h||"";C&&C!=="undefined"&&C!=="[object Object]"&&C!=="null"&&(d=C)}}r=`
          <input type="text" name="${t.name}" list="${l}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${l}">${a}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let S=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";r=`<input type="date" name="${t.name}" class="form-control" value="${S}" ${i}>`;break;case"month":let E=t.value?String(t.value).slice(0,7):"";r=`<input type="month" name="${t.name}" class="form-control" value="${E}" ${i}>`;break;case"number":let w=t.readonly?"readonly":"";r=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${w}>`;break;case"email":r=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let m=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${n}${r}${m}</div>`}).join("")}function at(e){let t={},i=new FormData(e);for(let[n,r]of i.entries())t[n]=r===""?null:r;return e.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(t[n.name]=null)}),t}function nt(e,t){t&&Object.entries(t).forEach(([i,n])=>{let r=e.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!n:r.type==="date"&&n&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(n):r.type==="month"&&n?r.value=String(n).slice(0,7):r.value=n??""))})}Y();var Ce={},it={on(e,t){Ce[e]||(Ce[e]=new Set),Ce[e].add(t)},off(e,t){Ce[e]&&Ce[e].delete(t)},emit(e,t){Ce[e]&&Ce[e].forEach(i=>{try{i(t)}catch(n){console.warn("[calendarBus] Handler error:",n)}})},clear(){Object.keys(Ce).forEach(e=>delete Ce[e])}},Wa=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function be(e){if(!e){it.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");it.emit("data:changed",{module:t,relevant:Wa.has(t)})}function G({container:e,title:t,icon:i,apiPath:n,columns:r,formFields:m,filterFields:o,defaultFilters:l={},itemLabel:a="Data",canCreate:d=!0,canEdit:S=!0,canDelete:E=!0,onBeforeSubmit:w,onAfterLoad:h,onDataLoaded:C,extraActions:b=[],initialSearch:f="",exportOptions:c=null,bulkDelete:v=!1,paginationMode:s="server"}){let y=oe();y&&y.role==="viewer"&&(d=!1,S=!1,E=!1,v=!1,c=null),y&&y.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(d=!1,S=!1,E=!1,v=!1,c=null),y&&y.role==="input_lembur"&&(n==="/api/overtime"?(E=!1,v=!1,c=null):(d=!1,S=!1,E=!1,v=!1,c=null));let g=1,x={...l};f&&(x.search=f);let p=new Set;e.innerHTML=`
    ${v?`
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
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
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
        ${o.filter(k=>k.type==="search").map(k=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${k.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let T=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select" name="${k.name}" id="filter-${k.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${T}</option>${(k.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${x[k.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let T=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select filter-select-sheet" name="${k.name}-sheet" id="filter-sheet-${k.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${T}</option>${(k.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${x[k.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function u(){let k=document.getElementById("bulk-toolbar");if(!k)return;let T=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),H=document.getElementById("btn-bulk-cancel");T.textContent=`${p.size} item dipilih`,p.size>0?(k.style.display="flex",A.disabled=!1,H.disabled=!1):(k.style.display="none",A.disabled=!0,H.disabled=!0);let R=document.getElementById("select-all-checkbox");if(R){let J=document.querySelectorAll(".row-checkbox");if(J.length>0){let K=[...J].every(ie=>ie.checked),j=[...J].some(ie=>ie.checked);R.checked=K,R.indeterminate=j&&!K}else R.checked=!1,R.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{p.clear(),document.querySelectorAll(".row-checkbox").forEach(T=>T.checked=!1);let k=document.getElementById("select-all-checkbox");k&&(k.checked=!1),u()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(p.size===0)return;let k=[...p],T=document.createElement("div");T.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",T.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${k.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${k.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(T),T.querySelector("#bulk-cancel-btn").addEventListener("click",()=>T.remove()),T.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=T.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let H=await B(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:k})});T.remove(),H.ok?(z(`${k.length} ${a} berhasil dihapus.`),p.clear(),u(),be(n),_()):O(H.data?.error||"Gagal menghapus data.")})});let $=document.getElementById("filter-search"),F;$?.addEventListener("input",k=>{clearTimeout(F),F=setTimeout(()=>{x.search=k.target.value,g=1,p.clear(),u(),_()},400)}),o?.forEach(k=>{(k.type==="select"||k.type==="combobox")&&(document.getElementById(`filter-${k.name}`)?.addEventListener("change",T=>{x[k.name]=T.target.value;let A=document.getElementById(`filter-sheet-${k.name}`);A&&(A.value=T.target.value),g=1,p.clear(),u(),_()}),document.getElementById(`filter-sheet-${k.name}`)?.addEventListener("change",T=>{x[k.name]=T.target.value;let A=document.getElementById(`filter-${k.name}`);A&&(A.value=T.target.value),g=1,p.clear(),u(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={},$&&($.value=""),o?.forEach(k=>{let T=document.getElementById(`filter-${k.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${k.name}`);A&&(A.value="")}),g=1,p.clear(),u(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{x={},$&&($.value=""),o?.forEach(k=>{let T=document.getElementById(`filter-${k.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${k.name}`);A&&(A.value="")}),g=1,p.clear(),u(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>P(null)),c&&document.addEventListener("click",function(k){let T=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");T&&A&&!A.contains(k.target)&&!T.contains(k.target)&&T.classList.remove("show-aksi-menu")});let I=document.getElementById("btn-mobile-filter"),L=document.getElementById("filter-options-wrapper"),M=document.getElementById("btn-close-filter-sheet");if(I&&L&&(I.addEventListener("click",k=>{k.preventDefault(),L.classList.add("sheet-open")}),M&&M.addEventListener("click",k=>{k.preventDefault(),L.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async T=>{let A=T.target,H=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await c.onExport()}catch{O("Gagal export data")}finally{A.innerHTML=H,A.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let k=document.getElementById(`input-import-${c.moduleName}`);k?.addEventListener("change",async T=>{let A=T.target.files[0];if(!A)return;let H=document.getElementById(`label-import-${c.moduleName}`),R=H?H.querySelector(".import-text"):null,J=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),H&&(H.style.pointerEvents="none"),k.disabled=!0;try{let K=await Ge(A);if(K.length===0)throw new Error("File kosong atau format salah");await c.onImport(K),z("Import berhasil!"),be(n),_()}catch(K){O(K.message||"Gagal import data")}finally{R&&(R.innerText=J),H&&(H.style.pointerEvents="auto"),k.disabled=!1,k.value=""}})}async function _(){u();let k=document.getElementById("table-container");if(!k)return;k.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let T=s==="client",A=T?1:g,H=T?Te:20,R=new URLSearchParams({page:A,limit:H,...Object.fromEntries(Object.entries(x).filter(([,W])=>W))}),J=await B(`${n}?${R}`);if(!J.ok){k.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${J.data?.error||"Error"}</p></div>`;return}let K=J.data?.data||J.data||[],j=J.data?.pagination,ie=K.length;if(T){K=C(K);let W=K.length,ne=20,ce=Math.ceil(W/ne);g>ce&&ce>0&&(g=ce);let V=(g-1)*ne,me=g*ne;K=K.slice(V,me),j={page:g,limit:ne,total:W,pages:ce}}!1,h&&h(K);let de=et({columns:r,data:K,onEdit:S?W=>P(W):null,actions:b.map(W=>({...W,handler:ne=>W.handler(ne,_)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:v?{selectedIds:p,onToggle:u}:null});k.innerHTML="",k.appendChild(de);let ge=document.getElementById("pagination-container");if(ge&&(ge.innerHTML="",j&&j.pages>1)){let W=tt({page:j.page,pages:j.pages,total:j.total,limit:j.limit,onPage:ne=>{g=ne,_()}});W&&ge.appendChild(W)}}function D(k){let T=typeof m=="function"?m(k):m;return Ue(T)}function P(k){let T=!!k,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=D(k),T){let R=typeof m=="function"?m(k):m;nt(A,k)}let{close:H}=Z({title:T?`Edit ${a}`:`Tambah ${a}`,content:A,size:"lg",confirmText:T?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(R,J)=>{if(!A.reportValidity())return;let K=R.querySelector(".modal-confirm");K.disabled=!0,K.textContent="Menyimpan...";let j=at(A),ie=typeof m=="function"?m(k):m,de=async ce=>{for(let V of ce)if(V.type==="row")await de(V.fields);else if(V.type==="combobox"&&j[V.name]){let me=j[V.name],Ee=(V.options||[]).find(re=>{let pe=String(typeof re=="object"?re.value:re),mt=String(typeof re=="object"?re.label:re);return pe===me||mt===me});if(Ee)j[V.name]=typeof Ee=="object"?Ee.value:Ee;else if(V.createApi){let re={};re[V.createApi.field]=me,V.createApi.extra&&Object.assign(re,V.createApi.extra);let pe=await B(V.createApi.path,{method:"POST",body:JSON.stringify(re)});if(pe.ok&&pe.data?.id)j[V.name]=pe.data.id;else if(pe.ok&&!pe.data?.id)j[V.name]=me;else throw new Error(`Gagal membuat master data: ${pe.data?.error||"Unknown error"}`)}}};try{await de(ie)}catch(ce){O(ce.message),K.disabled=!1,K.textContent=T?"Simpan Perubahan":`Tambah ${a}`;return}w&&(j=await w(j,k));let ge=T?"PUT":"POST",W=T?`${n}/${k.id}`:n,ne=await B(W,{method:ge,body:JSON.stringify(j)});ne.ok?(z(T?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),J(),be(n),_()):(O(ne.data?.error||"Gagal menyimpan data."),K.disabled=!1,K.textContent=T?"Simpan Perubahan":`Tambah ${a}`)}})}function N(k){ae(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let T=await B(`${n}/${k.id}`,{method:"DELETE"});T.ok?(z(`${a} berhasil dihapus.`),be(n),_()):O(T.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return _(),_}U();U();var De=null,rt=null;async function _e(e=!1){if(De&&!e)return console.log("Employees Raw (Cache Hit)",De.slice(0,5)),De;let t=await B(`/api/employees?limit=${Te}&status=Aktif`);return De=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",De.slice(0,5)),De}async function le(e=!1){let i=(await _e(e)).map(n=>({value:n.label,label:n.label}));return console.log("Employee Options",i.slice(0,5)),i}async function Q(e=!1){return rt&&!e||(rt=((await B("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),rt}function te(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function wt(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function Be(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function St(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function ye(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}Y();U();Y();U();Y();U();Y();function ot(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let r=new Date(t.value),m=new Date(i.value),o=Math.floor((m-r)/864e5);n.value=isNaN(o)?"":o}else n.value=""}}document.body.removeEventListener("input",ot);document.body.addEventListener("input",ot);document.body.removeEventListener("change",ot);document.body.addEventListener("change",ot);U();var $e={};function Ie(e){if($e[e]){try{$e[e].destroy()}catch{}delete $e[e]}}function Xa(){Object.keys($e).forEach(Ie)}var fe=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},Pe=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function st(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let n=Math.max(0,Math.round(fe(t))),r=e.textContent?e.textContent.replace(/[^\d]/g,""):"",m=r!==""&&parseInt(r)||0;if(n===m){e.textContent=n.toLocaleString("id-ID"),e._animFrame=null;return}let o=Date.now(),l=()=>{let a=Math.min((Date.now()-o)/i,1),d=1-Math.pow(1-a,3);e.textContent=Math.round(m+d*(n-m)).toLocaleString("id-ID"),a<1?e._animFrame=requestAnimationFrame(l):(e.textContent=n.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(l)}var Za={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},en=e=>{let t=Pe(e,"\u2014");return`<span class="status-pill ${Za[t]||"pill-neutral"}">${t}</span>`};var ve={family:"Inter",size:11},Se="#94A3B8",Le="#F1F5F9",Et=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],tn=()=>window.innerWidth<768;function lt(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:tn()?"bottom":"top",labels:{font:ve,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ve,titleFont:{...ve,weight:"700"}}},scales:{x:{grid:{color:Le},ticks:{font:ve,color:Se,maxRotation:0}},y:{grid:{color:Le},ticks:{font:ve,color:Se},beginAtZero:!0}},...e}}var an=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),nn=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Yt(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function xe(e,t,i=8e3){try{let n=new AbortController,r=setTimeout(()=>n.abort(),i),m=await B(e,{signal:n.signal}).catch(()=>null);if(clearTimeout(r),!m||!m.ok)return t;let o=m.data;return o?o.data!==void 0?o.data??t:o:t}catch{return t}}function rn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(n=>{let r=document.getElementById(n);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(n=>{let r=document.getElementById(n);if(r&&r.style.display==="none"){r.style.display="block";let m=r.parentElement;if(m&&!m.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",r.style.display="none",m.appendChild(o)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Xt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Zt({}),["table-contracts","table-issues"].forEach(n=>{let r=document.getElementById(n);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Wt(e){Xa(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${an()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${nn()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Yt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Yt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Qt(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let n=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=n);let m=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");m&&(m.style.display="block",m.style.position="absolute"),o&&(o.style.display="none");let l=await xe(`/api/dashboard/schedule-chart?year=${n}`,{},8e3);try{ea(l)}catch(a){console.warn("ScheduleChart render:",a),ke("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let n=i.target.value,r=n?`/api/dashboard/inspection-bar?month=${n}`:"/api/dashboard/inspection-bar",m=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");m&&(m.style.display="block",m.style.position="absolute"),o&&(o.style.display="none");let l=await xe(r,{},8e3);try{ta(l)}catch(a){console.warn("InspBar render:",a),ke("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>rn(),5e3),await Qt(e)}async function Qt(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,n,r,m,o,l,a]=await Promise.all([xe("/api/dashboard/kpi",{},8e3),xe("/api/dashboard/issues-trend",{},8e3),xe("/api/dashboard/issues-summary",{},8e3),xe("/api/dashboard/stats",{},8e3),xe("/api/dashboard/calendar",[],8e3),xe("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),xe(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),xe("/api/dashboard/dropdown-aggregates",{},8e3)]),d=document.getElementById("filter-insp-month"),S=d?d.value:"",E=S?`/api/dashboard/inspection-bar?month=${S}`:"/api/dashboard/inspection-bar",w=await xe(E,{},8e3);window.dashboardAggregates=a||{};try{Xt(t)}catch(h){console.warn("KPI render:",h)}try{Zt(t)}catch(h){console.warn("MiniStats render:",h)}try{ea(l)}catch(h){console.warn("ScheduleChart render:",h),ke("skel-jadwal","chart-jadwal")}try{on(Array.isArray(n?.by_category)?n.by_category:[])}catch(h){console.warn("Donut render:",h),ke("skel-donut","chart-donut")}try{sn(i)}catch(h){console.warn("Trend render:",h),ke("skel-trend","chart-trend")}try{ta(w)}catch(h){console.warn("InspBar render:",h),ke("skel-insp","chart-insp")}try{let h=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];dn(h)}catch(h){console.warn("IssuesTable render:",h)}try{let h=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];ln(o)}catch(h){console.warn("ContractsTable render:",h)}try{cn(Array.isArray(m)?m:[])}catch(h){console.warn("Agenda render:",h)}try{pn()}catch(h){console.warn("Quick Actions render:",h)}}function Xt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,n=window.dashboardAggregates||{},r=a=>{let d=window.dashboardAggregates?.schedule_quarter_details?.[a];return d?`Insp: ${d.insp} \u2022 GC: ${d.gc} \u2022 DC: ${d.dc} \u2022 Fog: ${d.fog}`:""},m=a=>{let d=window.dashboardAggregates?.schedule_quarter_details?.[a];return d?`Rincian ${a}: Inspeksi: ${d.insp} | General Cleaning: ${d.gc} | Deep Cleaning: ${d.dc} | Fogging: ${d.fog}`:`Total Kegiatan ${a}`},o=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:n.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:r(i),subId:"mini-jadwal-detail",tooltip:m(i)}];t.innerHTML=o.map(a=>{let d=a.val!==void 0?fe(a.val,0):fe(e[a.key]?.current,0);return a.dropdown?`
        <div class="kpi-card ${a.color}" ${a.id?`id="${a.id}"`:""} ${a.tooltip?`title="${a.tooltip}"`:""} data-href="${a.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${d}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${d}</div>
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
              <div class="kpi-value" data-target="${d}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${d}</div>
              ${a.dropdown?a.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" ${a.subId?`id="${a.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${a.tooltip||""}">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(a=>{a.addEventListener("click",d=>{d.target.closest("select")||a.dataset.href&&(window.location.hash=a.dataset.href)}),a.addEventListener("keydown",d=>{d.key==="Enter"&&!d.target.closest("select")&&a.dataset.href&&(window.location.hash=a.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(a=>{st(a,parseInt(a.dataset.target)||0)});let l=document.getElementById("dash-jadwal-period");if(l){let a=d=>{let E=(window.dashboardAggregates?.schedule_by_quarter||{})[d]||0,w=document.querySelector("#mini-jadwal .kpi-value");w&&(w.dataset.target=E,st(w,E,400));let h=document.getElementById("mini-jadwal-detail");h&&(h.textContent=r(d),h.title=m(d));let C=document.getElementById("mini-jadwal");if(C){let b=`#/timeline?dash_filter=period_${d.toLowerCase()}`;C.tagName==="A"?C.href=b:C.dataset.href=b,C.title=m(d)}};l.addEventListener("change",d=>a(d.target.value))}}function Zt(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),n=String(new Date().getMonth()+1).padStart(2,"0"),r=`${i}-${n}`,m=d=>`
    <select id="${d}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
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
  `,o=window.dashboardAggregates||{},l=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:m("dash-inspeksi-month"),val:o.inspeksi_by_month?.[r]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${r}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:m("dash-gcdc-month"),val:o.gcdc_by_month?.[r]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${r}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:m("dash-reliefer-month"),val:o.relievers_by_month?.[r]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${r}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:m("dash-fogging-month"),val:o.fogging_by_month?.[r]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${r}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=l.map(d=>d.dropdown?`
        <div class="mini-stat ${d.color}" style="cursor:pointer;" id="${d.id||""}" data-href="${d.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${d.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${fe(d.val)}">0</div>
              ${d.dropdown}
            </div>
            <div class="mini-stat-text">${d.label}</div>
            ${d.extra||""}
          </div>
        </div>`:`
      <a href="${d.href}" class="mini-stat ${d.color}" style="text-decoration:none" id="${d.id||""}">
        <div class="mini-stat-icon">${d.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${fe(d.val)}">0</div>
            ${d.dropdown?d.dropdown:""}
          </div>
          <div class="mini-stat-text">${d.label}</div>
          ${d.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(d=>{d.addEventListener("click",S=>{S.target.closest("select")||d.dataset.href&&(window.location.hash=d.dataset.href)}),d.addEventListener("keydown",S=>{S.key==="Enter"&&!S.target.closest("select")&&d.dataset.href&&(window.location.hash=d.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(d=>st(d,parseInt(d.dataset.target)||0,700));let a=(d,S,E,w)=>{let h=document.getElementById(d);if(h){let C=b=>{let c=(window.dashboardAggregates?.[E]||{})[b]||0,v=document.querySelector(`#${S} .mini-stat-value`);v&&(v.dataset.target=c,st(v,c,400));let s=document.getElementById(S);if(s){let y=`${w}&month=${b}`;s.tagName==="A"?s.href=y:s.dataset.href=y}};h.addEventListener("change",b=>C(b.target.value))}};a("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),a("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),a("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),a("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function on(e){ke("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Ie("donut");let n=(e||[]).filter(a=>fe(a.count)>0);if(!n.length){Ve(t,"Belum ada data permasalahan");return}let r=n.map(a=>`${Pe(a.category,"Lainnya")}`),m=n.map(a=>fe(a.count)),o=m.reduce((a,d)=>a+d,0);i.innerHTML=n.map((a,d)=>{let S=Et[d%Et.length],E=o>0?Math.round(a.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${S}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${a.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${E}%)</span></div>
          <div class="donut-legend-label">${r[d]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(a){let d=a.width,S=a.height,E=a.ctx;E.restore();let w=(S/80).toFixed(2);E.font="bold "+w+"em Inter",E.textBaseline="middle",E.fillStyle="#1E293B";let h=o.toString(),C=Math.round((d-E.measureText(h).width)/2),b=S/2;E.fillText(h,C,b-4),E.font="600 "+(w*.35).toFixed(2)+"em Inter",E.fillStyle="#64748B";let f="Total",c=Math.round((d-E.measureText(f).width)/2);E.fillText(f,c,b+10),E.save()}};$e.donut=new Chart(t,{type:"doughnut",data:{labels:r,datasets:[{data:m,backgroundColor:Et,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ve,titleFont:{...ve,weight:"700"},callbacks:{label:a=>` ${a.label}: ${a.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function sn(e){ke("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Ie("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],n=(e.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[l,a]=o.split("-");return(i[Number(a)-1]||a)+" "+String(l).slice(-2)}catch{return o}}),r=(e.open||[]).map(o=>fe(o)),m=(e.closed||[]).map(o=>fe(o));if(!n.length){Ve(t,"Belum ada data trend");return}$e.trend=new Chart(t,{type:"line",data:{labels:n,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:m,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:lt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:45,autoSkip:!0}},y:{grid:{color:Le},ticks:{font:{family:"Inter",size:9},color:Se},beginAtZero:!0}}})})}function ea(e){ke("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Ie("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(a=>Array.isArray(a)&&a.some(d=>d>0))){Ve(t,"Belum ada data jadwal");return}let r=e["Inspeksi Hygiene"]||Array(12).fill(0),m=e["General Cleaning"]||Array(12).fill(0),o=e["Deep Cleaning"]||Array(12).fill(0),l=e.Fogging||Array(12).fill(0);$e.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:m,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:lt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Le},ticks:{font:{family:"Inter",size:9},color:Se},min:0}}})})}function ta(e){ke("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Ie("inspBar"),e=e||{};let i=e.labels||[],n=(e.fc||[]).map(m=>fe(m)),r=(e.spv||[]).map(m=>fe(m));if(!i.length){Ve(t,"Belum ada data inspeksi");return}$e.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:lt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Se,maxRotation:45,minRotation:45}},y:{grid:{color:Le},ticks:{font:ve,color:Se},min:0,max:100}}})})}function ln(e){ke("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Ie("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},n=(e.labels||[]).map(o=>{let l=o.split("-")[1];return i[l]||o}),r=(e.data||[]).map(o=>fe(o));if(!n.length){Ve(t,"Belum ada data");return}let m=t.getContext("2d");$e.contractMiniBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:lt({onClick:(o,l)=>{if(l&&l.length>0){let a=l[0].index,d=(e.labels||[])[a];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Se,maxRotation:0,autoSkip:!1}},y:{grid:{color:Le,borderDash:[4,4],drawBorder:!1},ticks:{font:ve,color:Se,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function dn(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(n=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${en(n.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Pe(n.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Pe(n.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function cn(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,n=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,m=(e||[]).filter(o=>(o.event_date||"").startsWith(n)).slice(0,10);if(!m.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${m.map(o=>{let l="#3B82F6",a="#EFF6FF",d="Agenda",S=(o.title||"").toLowerCase();return S.includes("inspeksi")?(l="#10B981",a="#ECFDF5",d="Inspeksi"):S.includes("cleaning")||S.includes("gcdc")?(l="#3B82F6",a="#EFF6FF",d="Cleaning"):S.includes("reliefer")?(l="#F59E0B",a="#FFFBEB",d="Reliefer"):S.includes("fogging")&&(l="#8B5CF6",a="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Pe(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Pe(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${a};color:${l}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function pn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ke(e,t){let i=document.getElementById(e),n=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),n){n.style.display="block";let r=n.parentElement;if(r){let m=r.querySelector(".chart-empty");m&&m.remove()}}}function Ve(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=t,i.appendChild(r)}}U();async function aa(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),n=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),m=document.getElementById("login-password");r?.addEventListener("click",()=>{let o=m.type==="text";m.type=o?"password":"text",r.style.color=o?"":"var(--primary)"}),t?.addEventListener("submit",async o=>{o.preventDefault(),i.style.display="none";let l=t.username.value.trim(),a=t.password.value;if(!l||!a){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let d=await B("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:a})});d.ok&&d.data.success?(ft(d.data.data.token),He(d.data.data.user),z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=d.data.error||"Username atau password salah.",i.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}U();Y();async function un(){return await Q()}function dt(e){if(!e||e==="-"||String(e).trim()==="")return"-";let t=String(e).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],n=t.match(/^(\d{4})-(\d{1,2})/);if(n){let m=n[1],o=parseInt(n[2],10)-1;if(o>=0&&o<12)return`${i[o]} ${m}`}let r=t.split(/[\/\-\.]/);if(r.length===2&&r[1].length===4){let m=parseInt(r[0],10)-1;if(m>=0&&m<12)return`${i[m]} ${r[1]}`}if(r.length===3&&r[2].length===4){let m=parseInt(r[1],10)-1;if(m>=0&&m<12)return`${i[m]} ${r[2]}`}for(let m=0;m<12;m++)if(t.toLowerCase().includes(i[m].toLowerCase()))return t;return t}function gn(e,t){let i=String(e.status||"").toLowerCase();return t==="active"?i==="aktif":t==="reliefer"?e.division==="FC - RELIEFER"&&i==="aktif":!1}async function na(e,t){let i=await un(),n=t?t.get("dash_filter"):null;G({container:e,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>n?r.filter(m=>gn(m,n)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Be(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>te(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await B(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let m=r.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||"","Target Pindah OS":dt(o.target_pindah_os)==="-"?"":dt(o.target_pindah_os),"Target Selesai OS":dt(o.target_selesai_os)==="-"?"":dt(o.target_selesai_os)}));q(m,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async r=>{let m=a=>{if(!a)return null;let d=String(a||"").toLowerCase(),S=i.find(E=>String(E.label||"").toLowerCase()===d);return S?S.value:null},o=r.map(a=>({full_name:String(a["Nama Lengkap"]||"").trim(),branch_id:m(String(a.Cabang||"").trim()),division:String(a.Divisi||a["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(a["No. HP"]||a["No. Hp"]||"").trim(),join_date:String(a["Tgl Masuk"]||a["Tanggal Masuk"]||"").trim(),status:String(a.Status||"").trim(),target_pindah_os:String(a["Target Pindah OS"]||"").trim(),target_selesai_os:String(a["Target Selesai OS"]||"").trim(),notes:String(a.Catatan||"").trim()})).filter(a=>a.full_name),l=await B("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}U();Y();var $t=[],ia=[];async function mn(){$t=await Q(),ia=await _e()}var Ct=async e=>{let t=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(U(),Gt))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let m=r.data?.data||r.data||[],o=Array.isArray(m)?m:[];if(t=t.concat(o),o.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return t};async function ct(e,t){await mn(),G({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:n=>n,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:n=>Be(n)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:n=>window.formatDate(n)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:n=>!n||String(n).startsWith("2099")?"Tetap / PKWTT":window.formatDate(n)},{key:"days_remaining",label:"Sisa Kontrak",render:(n,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':wt(n)},{key:"status",label:"Status",render:n=>te(n)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:$t},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:n=>(n.end_date||(n.end_date="2099-12-31"),n),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let n=document.createElement("button");n.id="btn-find-missing",n.className="btn btn-ghost",n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.style.marginLeft="8px",n.style.color="#EF4444",n.style.border="1px solid currentColor",n.onclick=async()=>{n.innerHTML="\u231B Mencari...",n.disabled=!0;try{let[m,o]=await Promise.all([Ct("/api/employees?status=Aktif"),Ct("/api/contracts")]);if(m.length>0){let l=o.filter(E=>E.status==="Aktif"),a=new Set(l.map(E=>E.employee_id)),d=m.filter(E=>!a.has(E.id)),S=`<p style="margin-bottom:12px">Data yang terbaca: <b>${m.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${d.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;d.forEach(E=>{let w=o.filter(C=>C.employee_id===E.id),h='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(w.length>0){let C=w[0];h=`Pernah ada kontrak (Status: <b style="color:#EF4444">${C.status}</b>, Selesai: ${window.formatDate(C.end_date)})`}S+=`<li style="margin-bottom:8px"><b>${E.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${E.branch_name||"-"} | ${h}</span></li>`}),S+="</ul>",Promise.resolve().then(()=>(ue(),kt)).then(E=>E.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:S,cancelText:"Tutup"}))}}catch(m){console.error(m)}n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(n)}},formFields:n=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:ia,value:n?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:$t,value:n?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:n?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:n?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:n?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:n?.end_date&&!String(n.end_date).startsWith("2099")?n.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:n?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:n?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let n=await B(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let r=n.data.data.map(m=>({"Nama Lengkap":m.employee_name,Cabang:m.branch_name||"","Div / Bagian":m.division||"","Tanggal Mulai":m.start_date||"","Tanggal Selesai":m.end_date&&String(m.end_date).startsWith("2099")?"":m.end_date||"","Sisa Kontrak":m.end_date&&String(m.end_date).startsWith("2099")?"Tetap":m.days_remaining!==null&&m.days_remaining!==void 0?`${m.days_remaining} Hari`:"",Status:m.status||""}));q(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async n=>{let[r,m]=await Promise.all([B("/api/branches?limit=10000"),Ct("/api/employees")]),o=r.data?.data||[],l=m||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((f,c)=>{console.log(`${c+1}. ID: ${f.id}, Name: ${f.full_name}, Status: ${f.status}`)}));let a=f=>{if(!f)return null;let c=String(f||"").replace(/\s+/g," ").toLowerCase().trim(),v=o.find(s=>String(s.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===c||String(s.code||"").replace(/\s+/g," ").toLowerCase().trim()===c||String(s.name||"").replace(/\s+/g," ").toLowerCase().trim()===c);return v?v.id:null},d=(f,c)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${c}`),console.log(`Nama dari Excel : "${f}"`),!f)return console.log("Alasan gagal mapping : Nama kosong"),null;let v=String(f||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${v}"`),console.log(`Jumlah employee di database : ${l.length}`);let s=l.find(y=>String(y.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===v);return s?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${s.id}`),s.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},S=f=>{if(!f)return"";if(f instanceof Date&&!isNaN(f.getTime()))return f.toISOString().slice(0,10);let c=String(f).trim();if(/^\d{4,5}(\.\d+)?$/.test(c)){let s=Math.floor(Number(c));if(s>2e4&&s<99999){let y=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);let v=c.split(/[\/\-\.]/);if(v.length===3){let[s,y,g]=v.map(x=>x.trim());if(s.length===4&&y.length<=2&&g.length<=2)return`${s}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&y.length<=2&&s.length<=2)return`${g}-${y.padStart(2,"0")}-${s.padStart(2,"0")}`}return c},E=n.map((f,c)=>{let v=c+2,s=String(f["Nama Lengkap"]||"").trim(),y=f["Tanggal Mulai"],g=S(y);if(!g){let u=n.__worksheet,$=n.__headers||[],F=$.indexOf("Tanggal Mulai"),I="N/A",L="N/A",M="N/A";if(F!==-1&&u&&window.XLSX){let D=window.XLSX.utils.encode_cell({c:F,r:v-1});M=D;let P=u[D];P?(I=P.t||"undefined",L=P.w||"undefined"):I="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let _="Unknown";y==null||y===""?_="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":y instanceof Date&&isNaN(y.getTime())?_="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":_="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${v}`),console.log(`Employee Name : ${s}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${F})`),console.log(`Raw Cell Value : "${y}"`),console.log(`JavaScript Type : ${typeof y}`),console.log(`SheetJS Cell Type : ${I}`),console.log(`SheetJS Formatted Value : "${L}"`),console.log(`Value After Trim : "${String(y||"").trim()}"`),console.log(`Value After Date Parser : "${g}"`),console.log(`Is Empty : ${!y}`),console.log(`Is Invalid Date : ${y instanceof Date?isNaN(y.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${_}`),console.log(`Workbook Sheet : ${u?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${M}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(f,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify($)),console.log(`==========================
`)}let x=d(s,v),p=null;return x?g||(p="Tanggal Mulai kosong atau tidak berformat tanggal"):p="Karyawan tidak ditemukan di Database",{isValid:!!(x&&g),invalidReason:p,rowNum:v,data:{employee_id:x,branch_id:a(String(f.Cabang||"").trim()),division:String(f["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:g,end_date:S(f["Tanggal Selesai"])||"2099-12-31",status:String(f.Status||"").trim(),_rawName:s}}}),w=[],h=[];if(E.forEach(f=>{f.isValid?w.push(f.data):h.push({rowNum:f.rowNum,name:f.data._rawName,reason:f.invalidReason})}),console.log(`Split Validation - Valid: ${w.length}, Invalid: ${h.length}`),w.length===0){let f=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${n.length}
Valid: 0
Invalid: ${h.length}

Daftar Kegagalan (Contoh):
`;h.slice(0,10).forEach(c=>{f+=`- Row ${c.rowNum} | Nama: ${c.name} | Alasan: ${c.reason}
`}),h.length>10&&(f+=`- ... dan ${h.length-10} lainnya.
`),alert(f);return}let C=await B("/api/contracts/import",{method:"POST",body:JSON.stringify(w)}),b=`IMPORT SUMMARY
======================
`;b+=`Total Baris Excel : ${n.length}
`,b+=`Baris Valid       : ${w.length}
`,b+=`Baris Invalid     : ${h.length}

`,C&&C.data&&C.data.metrics?(b+=`Berhasil INSERT   : ${C.data.metrics.inserted}
`,b+=`Berhasil UPDATE   : ${C.data.metrics.updated}
`):b+=`Berhasil diproses : ${w.length}
`,h.length>0&&(b+=`
DAFTAR DATA DILEWATI:
`,h.forEach(f=>{b+=`- Row ${f.rowNum} | ${f.name} | ${f.reason}
`})),alert(b),typeof ct=="function"&&ct()}}})}U();Y();var Ft=[],Ye=[];function bn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(n=>n.period===i))return i;return"Q3"}async function ra(e,t){Ft=await Q();let i=await le();Ye=["Berlin Ariansyah","Ade Surahman"];let n=f=>f&&!Ye.find(c=>String(typeof c=="object"?c.value:c).toLowerCase()===String(f).toLowerCase())?[...Ye,f]:Ye,r=await B(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),m=f=>{if(!f||f==="-"||String(f).trim()==="")return"";let c=String(f).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:f},o=r.data?.data||[],l=bn(o),a=t?t.get("dash_filter"):null,d=new Date,S=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,E={},w=t&&t.get("month")?t.get("month"):null;a==="inspeksi"?E={status:"Done",activity_type:"Inspeksi Hygiene",month:w}:a==="gcdc"?E={status:"Done",activity_type:"GCDC",month:w}:a&&a.startsWith("period_")&&(E={period:a.replace("period_","").toUpperCase()});let h=new Date().getFullYear(),b=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((f,c)=>{let v=String(c+1).padStart(2,"0");return{value:`${h}-${v}`,label:`${f} ${h}`}});G({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:E,onDataLoaded:f=>f.sort((c,v)=>{let s=c.opening_date?new Date(c.opening_date).getTime():0;return(v.opening_date?new Date(v.opening_date).getTime():0)-s}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:f=>St(f)},{key:"period",label:"Periode",render:f=>ye(f)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:f=>m(f)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:f=>m(f)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:f=>m(f)},{key:"status",label:"Status",render:f=>te(f)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Ft},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:b},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ye}],formFields:f=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ft,value:f?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:f?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:f?.period},{name:"pic",label:"PIC",type:"combobox",options:n(f?.pic),value:f?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:f?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:f?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:f?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:f?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:f?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let f=await B(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(f.ok){let c=f.data.data.map(v=>({Cabang:v.branch_name||"",Kegiatan:v.activity_type||"",Periode:v.period||"",PIC:v.pic||"","Tgl Opening":v.opening_date||"","Tgl Target":v.target_date||"","Tgl Selesai":v.completion_date||"",Status:v.status||""}));q(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async f=>{let v=(await B("/api/branches?all=1")).data?.data||[],s=p=>{if(!p)return null;let u=String(p||"").toLowerCase().trim(),$=u.match(/^(\d{3})/);if($){let I=v.find(L=>String(L.code||"")===$[1]);if(I)return I.id}let F=v.find(I=>String(I.full_name||"").toLowerCase()===u||String(I.code||"").toLowerCase()===u||String(I.name||"").toLowerCase()===u);return F?F.id:null},y=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let F=Number(u);if(F>2e4&&F<99999){let I=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN(I.getTime())?"":I.toISOString().slice(0,10)}}let $=u.split(/[\/\-\.]/);if($.length===3){let[F,I,L]=$.map(M=>M.trim());if(F.length===4&&I.length<=2&&L.length<=2)return`${F}-${I.padStart(2,"0")}-${L.padStart(2,"0")}`;if(L.length===4&&I.length<=2&&F.length<=2)return`${L}-${I.padStart(2,"0")}-${F.padStart(2,"0")}`}return u},g=f.map(p=>({branch_id:s(String(p.Cabang||"").trim()),branch_name:String(p.Cabang||"").trim(),activity_type:String(p.Kegiatan||"").trim(),period:String(p.Periode||"").trim(),pic:String(p.PIC||p.Pic||"").trim(),opening_date:y(p["Tgl Opening"]||p["Tanggal Opening"]||p["Tgl Openir"]),target_date:y(p["Tgl Target"]||p["Tanggal Target"]),completion_date:y(p["Tgl Selesai"]||p["Tanggal Selesai"]),status:String(p.Status||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.activity_type&&p.period),x=await B("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!x.ok)throw new Error(x.data?.error||"Import gagal");return be("schedule"),x.data}}})}U();Y();var _t=[],pt=[];function ut(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let r=new Date(t.value),m=new Date(i.value),o=Math.floor((m-r)/864e5);n.value=isNaN(o)?"":o}else n.value=""}}document.body.removeEventListener("input",ut);document.body.addEventListener("input",ut);document.body.removeEventListener("change",ut);document.body.addEventListener("change",ut);function hn(e,t){let i=String(e.status||"").toLowerCase();return t==="open"?i==="open":!1}async function oa(e,t){let i=t?t.get("dash_filter"):null;_t=await Q(),pt=await le();let n=a=>a&&!pt.find(d=>d.value===a)?[...pt,{value:a,label:a}]:pt,r=new Date().getFullYear(),m=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((a,d)=>{let S=String(d+1).padStart(2,"0");return{value:`${r}-${S}`,label:`${a} ${r}`}});G({container:e,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:a=>i?a.filter(d=>hn(d,i)):a,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:a=>window.formatDate(a)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:a=>`<span class="badge badge-secondary">${a}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:a=>`<span title="${a}">${a?.length>50?a.slice(0,50)+"...":a}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:a=>`<span title="${a||""}">${a?.length>40?a.slice(0,40)+"...":a||"-"}</span>`},{key:"status",label:"Status",render:a=>te(a)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:a=>window.formatDate(a)},{key:"day_count",label:"Day",render:a=>a??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:m}],formFields:a=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:a?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:_t,value:a?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:a?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:a?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:a?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:n(a?.employee_name),value:a?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:n(a?.fc_specialist),value:a?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:a?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:a?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:a?.completion_date},{name:"day_count",label:"Day",type:"number",value:a?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let a=await B(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let d=a.data.data.map(S=>({"Tanggal Info":S.report_date||"",Cabang:S.branch_name||"",Kategori:S.category||"","Sumber Laporan":S.source||"",Keluhan:S.complaint||"","Nama FC":S.employee_name||"","FC Spesialis":S.fc_specialist||"",Solusi:S.solution||"",Status:S.status||"","Tanggal Selesai":S.completion_date||"",Day:S.day_count!==null?S.day_count:""}));q(d,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async a=>{let S=(await B("/api/branches?all=1")).data?.data||[],E=C=>{if(!C)return null;let b=String(C||"").toLowerCase(),f=S.find(c=>String(c.full_name||"").toLowerCase()===b||String(c.code||"").toLowerCase()===b||String(c.name||"").toLowerCase()===b);return f?f.id:null},w=a.map(C=>({branch_id:E(String(C.Cabang||"").trim()),report_date:String(C["Tanggal Info"]||C.Tanggal||"").trim(),category:String(C.Kategori||"").trim(),source:String(C["Sumber Laporan"]||C.Sumber||"").trim(),complaint:String(C.Keluhan||"").trim(),employee_name:String(C["Nama FC"]||"").trim(),fc_specialist:String(C["FC Spesialis"]||"").trim(),solution:String(C.Solusi||"").trim(),completion_date:String(C["Tanggal Selesai"]||C["Tgl Selesai"]||"").trim(),status:String(C.Status||"").trim(),day_count:C.Day||C.Hari||null})).filter(C=>C.report_date&&C.complaint&&C.category),h=await B("/api/import/issues",{method:"POST",body:JSON.stringify({rows:w,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}U();var Me=[];function fn(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function sa(e,t){let i=t?t.get("dash_filter"):null;Me=await Q();let n=await le(),r=["Ade","Berlin"],m=l=>l&&!n.find(a=>a.value===l)?[...n,{value:l,label:l}]:n,o=l=>l&&!r.find(a=>(typeof a=="object"?a.value:a)===l)?[...r,l]:r;G({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>i?l.filter(a=>fn(a,i)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>te(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let a=new URLSearchParams(l||{}).toString(),d=await B(`/api/one-on-one?limit=10000&${a}`);if(d.ok){let S=d.data.data.map(w=>({Tanggal:w.meeting_date||"",Cabang:w.branch_name||"","Nama Karyawan":w.employee_name||"",PIC:w.pic||"",Masalah:w.problem||"",Solusi:w.solution||"",Status:w.status||"","Tgl Selesai":w.completion_date||"",Dokumen:w.document_link||""})),{downloadExcel:E}=await Promise.resolve().then(()=>(Y(),he));E(S,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:a}=await Promise.resolve().then(()=>(Y(),he));a(l,"Template_Import_OneOnOne")},onImport:async l=>{let a=w=>{if(!w)return null;let h=String(w||"").toLowerCase(),C=Me.find(b=>String(b.label||"").toLowerCase()===h);return C?C.value:null},d=w=>{if(!w)return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let h=String(w).trim();if(/^\d{4,5}$/.test(h)){let b=Number(h);if(b>2e4&&b<99999){let f=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let C=h.split(/[\/\-\.]/);if(C.length===3){let[b,f,c]=C.map(v=>v.trim());if(b.length===4&&f.length<=2&&c.length<=2)return`${b}-${f.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&f.length<=2&&b.length<=2)return`${c}-${f.padStart(2,"0")}-${b.padStart(2,"0")}`}return h},S=l.map(w=>({meeting_date:d(w.Tanggal),employee_name:String(w["Nama Karyawan"]||"").trim(),branch_id:a(String(w.Cabang||"").trim()),pic:String(w.PIC||"").trim(),problem:String(w.Masalah||"").trim(),solution:String(w.Solusi||"").trim(),status:String(w.Status||"").trim(),completion_date:d(w["Tgl Selesai"]),document_link:String(w.Dokumen||"").trim()})).filter(w=>w.meeting_date&&w.employee_name&&w.branch_id),E=await B("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:S,onDuplicate:"update"})});if(!E.ok)throw new Error(E.data?.error||"Import gagal");return E.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Me.find(a=>a.value==l.branch_id)?[...Me,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:m(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}U();async function la(e){let t=await Q(),i=await le(),n=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=a=>a&&!i.find(d=>d.value===a)?[...i,{value:a,label:a}]:i,m=a=>a&&!n.find(d=>(typeof d=="object"?d.value:d)===a)?[...n,a]:n,o=Array.from({length:5},(a,d)=>String(new Date().getFullYear()-d));G({container:e,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:a=>{try{let d=JSON.parse(a);return Array.isArray(d)?d.join(", "):a||"-"}catch{return a||"-"}}},{key:"score",label:"Nilai",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async a=>{let d=new URLSearchParams(a||{}).toString(),S=await B(`/api/training?limit=10000&${d}`);if(S.ok){let E=S.data.data.map(h=>{let C=h.participants||"";try{let b=JSON.parse(C);C=Array.isArray(b)?b.join(", "):C}catch{}return{Tanggal:h.training_date||"",Batch:h.batch||"",Materi:h.subject||"",Cabang:h.branch_name||"",Trainer:h.trainer||"",Peserta:C,Nilai:h.score!==null&&h.score!==void 0?h.score:"",Dokumen:h.document_link||""}}),{downloadExcel:w}=await Promise.resolve().then(()=>(Y(),he));w(E,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let a=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:d}=await Promise.resolve().then(()=>(Y(),he));d(a,"Template_Import_Training")},onImport:async a=>{let d=h=>{if(!h)return null;let C=String(h||"").toLowerCase(),b=t.find(f=>String(f.label||"").toLowerCase()===C);return b?b.value:null},S=h=>{if(!h)return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let C=String(h).trim();if(/^\d{4,5}$/.test(C)){let f=Number(C);if(f>2e4&&f<99999){let c=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);let b=C.split(/[\/\-\.]/);if(b.length===3){let[f,c,v]=b.map(s=>s.trim());if(f.length===4&&c.length<=2&&v.length<=2)return`${f}-${c.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&c.length<=2&&f.length<=2)return`${v}-${c.padStart(2,"0")}-${f.padStart(2,"0")}`}return C},E=a.map(h=>({training_date:S(h.Tanggal),batch:String(h.Batch||"").trim(),subject:String(h.Materi||"").trim(),branch_id:d(String(h.Cabang||"").trim()),trainer:String(h.Trainer||"").trim(),participants:String(h.Peserta||"").trim(),score:h.Nilai?Number(h.Nilai):null,document_link:String(h.Dokumen||"").trim()})).filter(h=>h.training_date&&h.subject&&h.branch_id),w=await B("/api/training/import",{method:"POST",body:JSON.stringify(E)});if(!w.ok)throw new Error(w.data?.error||"Import gagal")}},formFields:a=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:a?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:a?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:a?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,value:a?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:m(a?.trainer),value:a?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let d=JSON.parse(a?.participants);return Array.isArray(d)?d.join(", "):a?.participants||""}catch{return a?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:a?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],onBeforeSubmit:async a=>a})}U();ue();Y();function da({container:e,title:t,icon:i,apiPath:n,columns:r,formFields:m,filterFields:o,defaultFilters:l={},itemLabel:a="Data",canCreate:d=!0,canEdit:S=!0,canDelete:E=!0,onBeforeSubmit:w,onAfterLoad:h,onDataLoaded:C,extraActions:b=[],initialSearch:f="",exportOptions:c=null,bulkDelete:v=!1,paginationMode:s="server"}){let y=oe();y&&y.role==="viewer"&&(d=!1,S=!1,E=!1,v=!1,c=null),y&&y.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(d=!1,S=!1,E=!1,v=!1,c=null),y&&y.role==="input_lembur"&&(n==="/api/overtime"?(E=!1,v=!1,c=null):(d=!1,S=!1,E=!1,v=!1,c=null));let g=1,x={...l};f&&(x.search=f);let p=new Set;e.innerHTML=`
    ${v?`
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
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
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
        ${o.filter(k=>k.type==="search").map(k=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${k.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let T=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select" name="${k.name}" id="filter-${k.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${T}</option>${(k.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${x[k.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(k=>k.type!=="search").map(k=>{if(k.type==="select"||k.type==="combobox"){let T=(k.label||"").startsWith("Pilih")?k.label:`Pilih ${k.label||""}`;return`<select class="filter-select filter-select-sheet" name="${k.name}-sheet" id="filter-sheet-${k.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${T}</option>${(k.options||[]).map(A=>`<option value="${typeof A=="object"?A.value:A}" ${x[k.name]===(typeof A=="object"?A.value:A)?"selected":""}>${typeof A=="object"?A.label:A}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function u(){let k=document.getElementById("bulk-toolbar");if(!k)return;let T=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),H=document.getElementById("btn-bulk-cancel");T.textContent=`${p.size} item dipilih`,p.size>0?(k.style.display="flex",A.disabled=!1,H.disabled=!1):(k.style.display="none",A.disabled=!0,H.disabled=!0);let R=document.getElementById("select-all-checkbox");if(R){let J=document.querySelectorAll(".row-checkbox");if(J.length>0){let K=[...J].every(ie=>ie.checked),j=[...J].some(ie=>ie.checked);R.checked=K,R.indeterminate=j&&!K}else R.checked=!1,R.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{p.clear(),document.querySelectorAll(".row-checkbox").forEach(T=>T.checked=!1);let k=document.getElementById("select-all-checkbox");k&&(k.checked=!1),u()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(p.size===0)return;let k=[...p],T=document.createElement("div");T.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",T.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${k.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${k.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(T),T.querySelector("#bulk-cancel-btn").addEventListener("click",()=>T.remove()),T.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=T.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let H=await B(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:k})});T.remove(),H.ok?(z(`${k.length} ${a} berhasil dihapus.`),p.clear(),u(),be(n),_()):O(H.data?.error||"Gagal menghapus data.")})});let $=document.getElementById("filter-search"),F;$?.addEventListener("input",k=>{clearTimeout(F),F=setTimeout(()=>{x.search=k.target.value,g=1,p.clear(),u(),_()},400)}),o?.forEach(k=>{(k.type==="select"||k.type==="combobox")&&(document.getElementById(`filter-${k.name}`)?.addEventListener("change",T=>{x[k.name]=T.target.value;let A=document.getElementById(`filter-sheet-${k.name}`);A&&(A.value=T.target.value),g=1,p.clear(),u(),_()}),document.getElementById(`filter-sheet-${k.name}`)?.addEventListener("change",T=>{x[k.name]=T.target.value;let A=document.getElementById(`filter-${k.name}`);A&&(A.value=T.target.value),g=1,p.clear(),u(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={},$&&($.value=""),o?.forEach(k=>{let T=document.getElementById(`filter-${k.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${k.name}`);A&&(A.value="")}),g=1,p.clear(),u(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{x={},$&&($.value=""),o?.forEach(k=>{let T=document.getElementById(`filter-${k.name}`);T&&(T.value="");let A=document.getElementById(`filter-sheet-${k.name}`);A&&(A.value="")}),g=1,p.clear(),u(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>P(null)),c&&document.addEventListener("click",function(k){let T=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");T&&A&&!A.contains(k.target)&&!T.contains(k.target)&&T.classList.remove("show-aksi-menu")});let I=document.getElementById("btn-mobile-filter"),L=document.getElementById("filter-options-wrapper"),M=document.getElementById("btn-close-filter-sheet");if(I&&L&&(I.addEventListener("click",k=>{k.preventDefault(),L.classList.add("sheet-open")}),M&&M.addEventListener("click",k=>{k.preventDefault(),L.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async T=>{let A=T.target,H=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await c.onExport()}catch{O("Gagal export data")}finally{A.innerHTML=H,A.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let k=document.getElementById(`input-import-${c.moduleName}`);k?.addEventListener("change",async T=>{let A=T.target.files[0];if(!A)return;let H=document.getElementById(`label-import-${c.moduleName}`),R=H?H.querySelector(".import-text"):null,J=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),H&&(H.style.pointerEvents="none"),k.disabled=!0;try{let K=await Ge(A);if(K.length===0)throw new Error("File kosong atau format salah");await c.onImport(K),z("Import berhasil!"),be(n),_()}catch(K){O(K.message||"Gagal import data")}finally{R&&(R.innerText=J),H&&(H.style.pointerEvents="auto"),k.disabled=!1,k.value=""}})}async function _(){u();let k=document.getElementById("table-container");if(!k)return;k.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let T=s==="client",A=T?1:g,H=T?Te:20,R=new URLSearchParams({page:A,limit:H,...Object.fromEntries(Object.entries(x).filter(([,W])=>W))}),J=await B(`${n}?${R}`);if(!J.ok){k.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${J.data?.error||"Error"}</p></div>`;return}let K=J.data?.data||J.data||[],j=J.data?.pagination,ie=K.length;if(T){K=C(K);let W=K.length,ne=20,ce=Math.ceil(W/ne);g>ce&&ce>0&&(g=ce);let V=(g-1)*ne,me=g*ne;K=K.slice(V,me),j={page:g,limit:ne,total:W,pages:ce}}!1,h&&h(K);let de=et({columns:r,data:K,onEdit:S?W=>P(W):null,actions:b.map(W=>({...W,handler:ne=>W.handler(ne,_)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:v?{selectedIds:p,onToggle:u}:null});k.innerHTML="",k.appendChild(de);let ge=document.getElementById("pagination-container");if(ge&&(ge.innerHTML="",j&&j.pages>1)){let W=tt({page:j.page,pages:j.pages,total:j.total,limit:j.limit,onPage:ne=>{g=ne,_()}});W&&ge.appendChild(W)}}function D(k){let T=typeof m=="function"?m(k):m;return Ue(T)}function P(k){let T=!!k,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=D(k),T){let R=typeof m=="function"?m(k):m;nt(A,k)}let{close:H}=Z({title:T?`Edit ${a}`:`Tambah ${a}`,content:A,size:"lg",confirmText:T?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(R,J)=>{if(!A.reportValidity())return;let K=R.querySelector(".modal-confirm");K.disabled=!0,K.textContent="Menyimpan...";let j=at(A),ie=typeof m=="function"?m(k):m,de=async ce=>{for(let V of ce)if(V.type==="row")await de(V.fields);else if(V.type==="combobox"&&j[V.name]){let me=j[V.name],Ee=(V.options||[]).find(re=>{let pe=String(typeof re=="object"?re.value:re),mt=String(typeof re=="object"?re.label:re);return pe===me||mt===me});if(Ee)j[V.name]=typeof Ee=="object"?Ee.value:Ee;else if(V.createApi){let re={};re[V.createApi.field]=me,V.createApi.extra&&Object.assign(re,V.createApi.extra);let pe=await B(V.createApi.path,{method:"POST",body:JSON.stringify(re)});if(pe.ok&&pe.data?.id)j[V.name]=pe.data.id;else if(pe.ok&&!pe.data?.id)j[V.name]=me;else throw new Error(`Gagal membuat master data: ${pe.data?.error||"Unknown error"}`)}}};try{await de(ie)}catch(ce){O(ce.message),K.disabled=!1,K.textContent=T?"Simpan Perubahan":`Tambah ${a}`;return}w&&(j=await w(j,k));let ge=T?"PUT":"POST",W=T?`${n}/${k.id}`:n,ne=await B(W,{method:ge,body:JSON.stringify(j)});ne.ok?(z(T?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),J(),be(n),_()):(O(ne.data?.error||"Gagal menyimpan data."),K.disabled=!1,K.textContent=T?"Simpan Perubahan":`Tambah ${a}`)}})}function N(k){ae(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let T=await B(`${n}/${k.id}`,{method:"DELETE"});T.ok?(z(`${a} berhasil dihapus.`),be(n),_()):O(T.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return _(),_}U();Y();async function ca(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await Q(),n=await le(),r=t?t.get("dash_filter"):null,m={};if(r==="reliever"){let w=new Date,h=`${w.getFullYear()}-${String(w.getMonth()+1).padStart(2,"0")}`;m={status:"Done",month:t&&t.get("month")?t.get("month"):h}}console.log("RAW",await _e()),console.log("OPTIONS",n);let o=w=>w&&!n.find(h=>h.value===w)?[...n,{value:w,label:w}]:n,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],a=w=>w&&!l.includes(w)?[...l,w]:l,d=new Date().getFullYear(),E=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((w,h)=>{let C=String(h+1).padStart(2,"0");return{value:`${d}-${C}`,label:`${w} ${d}`}});da({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:m,onDataLoaded:w=>w.sort((h,C)=>{let b=h.backup_date?new Date(h.backup_date).getTime():0;return(C.backup_date?new Date(C.backup_date).getTime():0)-b}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:w=>ye(w)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:w=>window.formatDate(w)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:w=>window.formatDate(w)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:w=>w?`<span class="badge badge-info">${w}</span>`:"-"},{key:"status",label:"Status",render:w=>te(w)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:E},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:w=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:w?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:w?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(w?.original_fc_name),value:w?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:a(w?.reliever_name),value:w?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:w?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:w?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:w?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:w?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:w?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let w=await B(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(w.ok){let h=w.data.data.map(C=>({Cabang:C.branch_name||"","Nama Facility care":C.original_fc_name||"",Periode:C.period||"",Relifer:C.reliever_name||"","Tanggal Back Up":C.backup_date||"","Tanggal Selesai":C.completion_date||"",Keterangan:C.reason||"",Shift:C.shift||"",Status:C.status||""}));h.length===0&&h.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),q(h,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async w=>{let C=(await B("/api/branches?all=1")).data?.data||[],b=v=>{if(!v)return null;let s=String(v||"").toLowerCase(),y=C.find(g=>String(g.full_name||"").toLowerCase()===s||String(g.code||"").toLowerCase()===s||String(g.name||"").toLowerCase()===s);return y?y.id:null},f=w.map(v=>({branch_name:String(v.Cabang||"").trim(),backup_date:String(v["Tanggal Back Up"]||v["Tanggal Backup"]||"").trim(),original_fc_name:String(v["Nama Facility care"]||v["FC Digantikan"]||"").trim(),reliever_name:String(v.Relifer||v.Reliefer||"").trim(),period:String(v.Periode||"").trim(),reason:String(v.Keterangan||"").trim(),shift:String(v.Shift||"").trim(),completion_date:String(v["Tanggal Selesai"]||"").trim(),status:String(v.Status||"").trim()})).filter(v=>v.reliever_name&&v.backup_date),c=await B("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:f,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}U();Y();async function pa(e){let t=await Q(),i=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r));G({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:n=>ye(n)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"fc_score",label:"Point FC",render:n=>n!=null?`<strong class="${n>=80?"text-success":n>=60?"text-warning":"text-danger"}">${n}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"status",label:"Status",render:n=>te(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:n?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:n?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:n?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),m=await B(`/api/reports/inspection?limit=10000&${r}`);if(m.ok){let o=m.data.data.map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));q(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async n=>{let r=a=>{if(!a)return null;let d=String(a||"").toLowerCase(),S=t.find(E=>String(E.label||"").toLowerCase()===d);return S?S.value:null},m=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let d=String(a).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let E=Number(d);if(E>2e4&&E<99999){let w=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let S=d.split(/[\/\-\.]/);if(S.length===3){let[E,w,h]=S.map(C=>C.trim());if(E.length===4&&w.length<=2&&h.length<=2)return`${E}-${w.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&w.length<=2&&E.length<=2)return`${h}-${w.padStart(2,"0")}-${E.padStart(2,"0")}`}return d},o=n.map(a=>({branch_id:r(String(a.Cabang||"").trim()),period:String(a.Periode||"").trim(),inspection_date:m(a.Tanggal),fc_score:a["Point FC"]!==void 0&&a["Point FC"]!==""?Number(a["Point FC"]):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==""?Number(a["Point SPV"]):null,status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.period&&a.inspection_date),l=await B("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}U();Y();async function ua(e){let t=await Q(),i=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r));G({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge ${n==="Deep Cleaning"?"badge-purple":"badge-success"}">${n}</span>`},{key:"period",label:"Periode",render:n=>ye(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>te(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:n?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:n?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:n?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),m=await B(`/api/reports/cleaning?limit=10000&${r}`);if(m.ok){let o=m.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));q(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async n=>{let r=a=>{if(!a)return null;let d=String(a||"").toLowerCase(),S=t.find(E=>String(E.label||"").toLowerCase()===d);return S?S.value:null},m=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let d=String(a).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let E=Number(d);if(E>2e4&&E<99999){let w=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let S=d.split(/[\/\-\.]/);if(S.length===3){let[E,w,h]=S.map(C=>C.trim());if(E.length===4&&w.length<=2&&h.length<=2)return`${E}-${w.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&w.length<=2&&E.length<=2)return`${h}-${w.padStart(2,"0")}-${E.padStart(2,"0")}`}return d},o=n.map(a=>({branch_id:r(String(a.Cabang||"").trim()),activity_type:String(a.Jenis||a.Kegiatan||"").trim(),period:String(a.Periode||"").trim(),activity_date:m(a.Tanggal),status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.activity_type&&a.period&&a.activity_date),l=await B("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}U();Y();async function ga(e,t){let i=await Q(),n=Array.from({length:4},(o,l)=>String(new Date().getFullYear()-l)),r=t?t.get("dash_filter"):null,m={};if(r==="fogging"){let o=new Date,l=String(o.getMonth()+1).padStart(2,"0"),a=String(o.getFullYear()),d=t?t.get("month"):null;d&&d.length===7&&(a=d.split("-")[0],l=d.split("-")[1]),m={status:"Done",month:l,year:a}}G({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:m,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>ye(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>te(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let l=new URLSearchParams(o||{}).toString(),a=await B(`/api/reports/fogging?limit=10000&${l}`);if(a.ok){let d=a.data.data.map(S=>({Cabang:S.branch_name||"",Jenis:S.activity_type||"Fogging",Periode:S.period||"",Tanggal:S.activity_date||"",Status:S.status||"","Link Dokumen":S.document_link||""}));q(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let l=E=>{if(!E)return null;let w=String(E||"").toLowerCase(),h=i.find(C=>String(C.label||"").toLowerCase()===w);return h?h.value:null},a=E=>{if(E==null||E==="")return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let w=String(E).trim();if(w===""||w==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(w))return w.slice(0,10);if(/^\d{4,5}$/.test(w)){let C=Number(w);if(C>2e4&&C<99999){let b=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let h=w.split(/[\/\-\.]/);if(h.length===3){let[C,b,f]=h.map(c=>c.trim());if(C.length===4&&b.length<=2&&f.length<=2)return`${C}-${b.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&b.length<=2&&C.length<=2)return`${f}-${b.padStart(2,"0")}-${C.padStart(2,"0")}`}return w},d=o.map(E=>({branch_id:l(String(E.Cabang||"").trim()),activity_type:String(E.Jenis||E.Kegiatan||"Fogging").trim(),period:String(E.Periode||"").trim(),activity_date:a(E.Tanggal),status:String(E.Status||"").trim(),document_link:String(E["Link Dokumen"]||"").trim(),notes:String(E.Catatan||E.Keterangan||"").trim()})).filter(E=>E.branch_id&&E.period&&E.activity_date),S=await B("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!S.ok)throw new Error(S.data?.error||"Import gagal");return S.data}}})}U();Y();async function ma(e){let t=await Q(),i=await le(),n=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(l,a)=>String(new Date().getFullYear()-a)),m=l=>l&&!i.find(a=>a.value===l)?[...i,{value:l,label:l}]:i,o=l=>l&&!n.find(a=>a.value===l)?[...n,{value:l,label:l}]:n;G({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>te(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let a=new URLSearchParams(l||{}).toString(),d=await B(`/api/reports/basecamp?limit=10000&${a}`);if(d.ok){let S=d.data.data.map(E=>({"Tgl Info":E.info_date||"",Cabang:E.branch_name||"",Permasalahan:E.problem||"",PIC:E.pic||"","Tgl Done":E.done_date||"",Status:E.status||"",Keterangan:E.notes||""}));q(S,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let a=w=>{if(!w)return null;let h=String(w||"").toLowerCase(),C=t.find(b=>String(b.label||"").toLowerCase()===h);return C?C.value:null},d=w=>{if(w==null||w==="")return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let h=String(w).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let b=Number(h);if(b>2e4&&b<99999){let f=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let C=h.split(/[\/\-\.]/);if(C.length===3){let[b,f,c]=C.map(v=>v.trim());if(b.length===4&&f.length<=2&&c.length<=2)return`${b}-${f.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&f.length<=2&&b.length<=2)return`${c}-${f.padStart(2,"0")}-${b.padStart(2,"0")}`}return h},S=l.map(w=>({info_date:d(w["Tgl Info"]||w["Tanggal Info"]),branch_id:a(String(w.Cabang||"").trim()),problem:String(w.Permasalahan||"").trim(),pic:String(w.PIC||"").trim(),done_date:d(w["Tgl Done"]||w["Tanggal Done"]),status:String(w.Status||"").trim(),notes:String(w.Keterangan||w.Catatan||"").trim()})).filter(w=>w.info_date&&w.branch_id&&w.problem),E=await B("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(S)});if(!E.ok)throw new Error(E.data?.error||"Import gagal");return E.data}}})}ue();var yn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],xn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],vn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Bt="fcms_sop_fc_ik_v1",Pt="fcms_sop_fc_tools_v1",Lt="fcms_sop_fc_chem_v1";function ba(){try{let e=localStorage.getItem(Bt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(yn))}function Tt(e){localStorage.setItem(Bt,JSON.stringify(e))}function ha(){try{let e=localStorage.getItem(Pt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(xn))}function At(e){localStorage.setItem(Pt,JSON.stringify(e))}function fa(){try{let e=localStorage.getItem(Lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(vn))}function Dt(e){localStorage.setItem(Lt,JSON.stringify(e))}async function ya(e){let t="ik",i="",n="all",r="tools",m="",o=ba(),l=ha(),a=fa();function d(){e.innerHTML=`
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
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${S(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${E(t==="ik")}">${o.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${S(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${E(t==="tools")}">${l.length+a.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${S(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${S(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,w(),h()}function S(u){return u?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function E(u){return u?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function w(){e.querySelectorAll(".sop-tab-btn").forEach(I=>{I.addEventListener("click",()=>{t=I.dataset.tab,d()})});let u=e.querySelector("#btn-reset-default");u&&u.addEventListener("click",()=>{ae("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Bt),localStorage.removeItem(Pt),localStorage.removeItem(Lt),o=ba(),l=ha(),a=fa(),z("Data SOP Facility Care berhasil di-reset ke standar asli"),d()},"Konfirmasi Reset Standar")});let $=e.querySelector("#btn-print-sop");$&&$.addEventListener("click",()=>{window.print()});let F=e.querySelector("#btn-export-tools");F&&F.addEventListener("click",()=>{x()})}function h(){let u=e.querySelector("#sop-content-area");u&&(t==="ik"?C(u):t==="tools"?c(u):t==="policy"?y(u):t==="links"&&g(u))}function C(u){u.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${o.length})</button>
              <button class="cat-pill ${n==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${o.filter(D=>D.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${n==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${o.filter(D=>D.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${n==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${o.filter(D=>D.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${n==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${o.filter(D=>D.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${n==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${o.filter(D=>D.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${o.length} dari ${o.length} Instruksi Kerja
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
    `,u.appendChild($);let F=u.querySelector("#ik-cards-grid-container"),I=u.querySelector("#ik-count-label");function L(){let D=o.filter(P=>{let N=n==="all"||P.category===n,k=i.toLowerCase().trim(),T=!k||P.title.toLowerCase().includes(k)||P.target.toLowerCase().includes(k)||P.tools.some(A=>A.toLowerCase().includes(k))||P.steps.some(A=>A.toLowerCase().includes(k));return N&&T});if(I&&(I.textContent=`Menampilkan ${D.length} dari ${o.length} Instruksi Kerja`),D.length===0){F.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}F.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${D.map(P=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${P.bg||"#EFF6FF"}; border-radius:10px;">${P.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${P.id}. ${P.title}
                      </h3>
                      <span class="badge" style="background:${P.bg||"#EFF6FF"}; color:${P.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${P.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${P.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${P.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${P.color||"#2563EB"};">
                  <strong>Area:</strong> ${P.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${P.tools.slice(0,3).map(N=>`<li style="margin-bottom:2px;">${N}</li>`).join("")}
                    ${P.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${P.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${P.chemical_ratio&&P.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${P.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${P.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${P.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,F.querySelectorAll(".btn-view-ik").forEach(P=>{P.addEventListener("click",()=>{let N=parseInt(P.dataset.id,10),k=o.find(T=>T.id===N);k&&b(k)})}),F.querySelectorAll(".btn-edit-ik").forEach(P=>{P.addEventListener("click",()=>{let N=parseInt(P.dataset.id,10),k=o.find(T=>T.id===N);k&&f(k)})}),F.querySelectorAll(".btn-delete-ik").forEach(P=>{P.addEventListener("click",()=>{let N=parseInt(P.dataset.id,10),k=o.find(T=>T.id===N);k&&ae(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${k.title}"?`,()=>{o=o.filter(T=>T.id!==N),Tt(o),z("Instruksi Kerja berhasil dihapus"),d()},"Hapus Instruksi Kerja")})})}L();let M=u.querySelector("#ik-search");M&&M.addEventListener("input",D=>{i=D.target.value,L()}),u.querySelectorAll(".cat-pill").forEach(D=>{D.addEventListener("click",()=>{u.querySelectorAll(".cat-pill").forEach(P=>P.classList.remove("active")),D.classList.add("active"),n=D.dataset.cat,L()})});let _=u.querySelector("#btn-add-ik");_&&_.addEventListener("click",()=>{f()})}function b(u){Z({title:`${u.icon||"\u{1F4CB}"} IK-${u.id}: ${u.title}`,size:"lg",content:`
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
      `,confirmText:"Tutup Panduan",onConfirm:($,F)=>{F()}})}function f(u=null){let $=!!u,F=u?u.tools.join(`
`):"",I=u?u.steps.join(`
`):"",L=`
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
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${p(F)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${p(I)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${u?p(u.notes||""):""}</textarea>
        </div>
      </form>
    `;Z({title:$?`\u270F\uFE0F Edit IK-${u.id}: ${u.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:L,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(M,_)=>{let D=M.querySelector("#form-ik-title").value.trim(),P=M.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",N=M.querySelector("#form-ik-cat").value,k=M.querySelector("#form-ik-ratio").value.trim(),T=M.querySelector("#form-ik-target").value.trim(),A=M.querySelector("#form-ik-tools").value.trim(),H=M.querySelector("#form-ik-steps").value.trim(),R=M.querySelector("#form-ik-notes").value.trim();if(!D||!T||!A||!H){O("Harap lengkapi seluruh field wajib (*)");return}let J=A.split(`
`).map(de=>de.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),K=H.split(`
`).map(de=>de.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),ie={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[N]||{color:"#2563EB",bg:"#EFF6FF"};if($)u.title=D,u.icon=P,u.category=N,u.chemical_ratio=k,u.target=T,u.tools=J,u.steps=K,u.notes=R,u.color=ie.color,u.bg=ie.bg,Tt(o),z(`Instruksi Kerja "${D}" berhasil diperbarui`);else{let de=o.length>0?Math.max(...o.map(W=>W.id))+1:1,ge={id:de,title:D,category:N,icon:P,color:ie.color,bg:ie.bg,target:T,chemical_ratio:k||"-",tools:J,steps:K,notes:R};o.push(ge),Tt(o),z(`Instruksi Kerja baru IK-${de} berhasil ditambahkan`)}_(),d()}})}function c(u){u.innerHTML=`
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
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${m}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${r==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let $=u.querySelector("#tool-table-container"),F=u.querySelector("#tool-search");function I(){let L=m.toLowerCase().trim();if(r==="tools"){let M=l.filter(_=>!L||_.name.toLowerCase().includes(L)||_.area.toLowerCase().includes(L));$.innerHTML=`
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
                  ${M.map((_,D)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${D%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||D+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${p(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${p(_.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${p(_.area)}</td>
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
        `,$.querySelectorAll(".btn-edit-tool").forEach(_=>{_.addEventListener("click",()=>{let D=parseInt(_.dataset.no,10),P=l.find(N=>N.no===D);P&&v(P)})}),$.querySelectorAll(".btn-delete-tool").forEach(_=>{_.addEventListener("click",()=>{let D=parseInt(_.dataset.no,10),P=l.find(N=>N.no===D);P&&ae(`Apakah Anda yakin ingin menghapus alat: "${P.name}"?`,()=>{l=l.filter(N=>N.no!==D),l.forEach((N,k)=>{N.no=k+1}),At(l),z("Alat kebersihan berhasil dihapus"),d()},"Hapus Alat Kebersihan")})})}else{let M=a.filter(_=>!L||_.name.toLowerCase().includes(L)||_.function.toLowerCase().includes(L));$.innerHTML=`
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
                  ${M.map((_,D)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${D%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||D+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${p(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${p(_.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${p(_.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${p(_.function)}</td>
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
        `,$.querySelectorAll(".btn-edit-chem").forEach(_=>{_.addEventListener("click",()=>{let D=parseInt(_.dataset.no,10),P=a.find(N=>N.no===D);P&&s(P)})}),$.querySelectorAll(".btn-delete-chem").forEach(_=>{_.addEventListener("click",()=>{let D=parseInt(_.dataset.no,10),P=a.find(N=>N.no===D);P&&ae(`Apakah Anda yakin ingin menghapus chemical: "${P.name}"?`,()=>{a=a.filter(N=>N.no!==D),a.forEach((N,k)=>{N.no=k+1}),Dt(a),z("Bahan chemical berhasil dihapus"),d()},"Hapus Bahan Chemical")})})}}I(),u.querySelector("#view-tools-btn").addEventListener("click",()=>{r="tools",c(u)}),u.querySelector("#view-chem-btn").addEventListener("click",()=>{r="chemical",c(u)}),F&&F.addEventListener("input",L=>{m=L.target.value,I()}),u.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{r==="tools"?v():s()})}function v(u=null){let $=!!u,F=`
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
    `;Z({title:$?`\u270F\uFE0F Edit Alat: ${u.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:F,confirmText:"Simpan",onConfirm:(I,L)=>{let M=I.querySelector("#form-tool-name").value.trim(),_=I.querySelector("#form-tool-qty").value.trim(),D=I.querySelector("#form-tool-area").value.trim();if(!M||!_||!D){O("Harap lengkapi semua kolom");return}if($)u.name=M,u.qty=_,u.area=D,At(l),z(`Alat "${M}" berhasil diperbarui`);else{let P=l.length>0?Math.max(...l.map(N=>N.no))+1:1;l.push({no:P,name:M,qty:_,area:D}),At(l),z(`Alat baru "${M}" berhasil ditambahkan`)}L(),d()}})}function s(u=null){let $=!!u,F=`
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
    `;Z({title:$?`\u270F\uFE0F Edit Chemical: ${u.name}`:"\u2795 Tambah Chemical Baru",content:F,confirmText:"Simpan",onConfirm:(I,L)=>{let M=I.querySelector("#form-chem-name").value.trim(),_=I.querySelector("#form-chem-size").value.trim(),D=I.querySelector("#form-chem-ratio").value.trim(),P=I.querySelector("#form-chem-function").value.trim();if(!M||!_||!D||!P){O("Harap lengkapi semua kolom");return}if($)u.name=M,u.size=_,u.ratio=D,u.function=P,Dt(a),z(`Chemical "${M}" berhasil diperbarui`);else{let N=a.length>0?Math.max(...a.map(k=>k.no))+1:1;a.push({no:N,name:M,size:_,ratio:D,function:P}),Dt(a),z(`Chemical baru "${M}" berhasil ditambahkan`)}L(),d()}})}function y(u){u.innerHTML=`
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
    `}function g(u){u.innerHTML=`
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
    `}async function x(){try{let{downloadExcel:u}=await Promise.resolve().then(()=>(Y(),he)),$=o.map(I=>({"No IK":`IK-${I.id}`,"Judul Pekerjaan":I.title,Kategori:I.category,"Sasaran Target Area":I.target,"Rasio Chemical":I.chemical_ratio||"-","Alat & Bahan":I.tools.join(", "),"Langkah Kerja":I.steps.map((L,M)=>`${M+1}. ${L}`).join(" | "),"Catatan Safety":I.notes||"-"})),F=new Date().toISOString().slice(0,10);u($,`SOP_Facility_Care_Instruksi_Kerja_${F}`),z("Data Excel Instruksi Kerja berhasil diunduh")}catch(u){console.error(u),O("Gagal mendownload data Excel: "+u.message)}}function p(u){return u?String(u).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}d()}ue();var kn=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],wn=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],xa={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},Nt="fcms_sop_sec_ik_v1",zt="fcms_sop_sec_assets_v1";function va(){try{let e=localStorage.getItem(Nt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(kn))}function It(e){localStorage.setItem(Nt,JSON.stringify(e))}function ka(){try{let e=localStorage.getItem(zt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(wn))}function Mt(e){localStorage.setItem(zt,JSON.stringify(e))}async function wa(e){let t="ik",i="",n="all",r="",m="cabang",o=va(),l=ka();function a(){e.innerHTML=`
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
        <button class="sec-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${d(t==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${S(t==="ik")}">${o.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="zonasi"?"active":""}" data-tab="zonasi" style="${d(t==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${t==="assets"?"active":""}" data-tab="assets" style="${d(t==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${S(t==="assets")}">${l.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${d(t==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${d(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,E(),w()}function d(p){return p?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function S(p){return p?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function E(){e.querySelectorAll(".sec-tab-btn").forEach(F=>{F.addEventListener("click",()=>{t=F.dataset.tab,a()})});let p=e.querySelector("#btn-reset-sec-default");p&&p.addEventListener("click",()=>{ae("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(Nt),localStorage.removeItem(zt),o=va(),l=ka(),z("Data SOP Security berhasil di-reset ke standar dokumen"),a()},"Reset ke Standar Dokumen")});let u=e.querySelector("#btn-print-sec");u&&u.addEventListener("click",()=>{window.print()});let $=e.querySelector("#btn-export-sec");$&&$.addEventListener("click",()=>{g()})}function w(){let p=e.querySelector("#sec-content-area");p&&(t==="ik"?h(p):t==="zonasi"?f(p):t==="assets"?c(p):t==="policy"?s(p):t==="links"&&y(p))}function h(p){p.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${o.length})</button>
              <button class="sec-cat-pill ${n==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${o.filter(_=>_.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${n==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${o.filter(_=>_.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${n==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${o.filter(_=>_.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${o.length} Instruksi Kerja
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
    `,p.appendChild(u);let $=p.querySelector("#sec-ik-cards-grid-container"),F=p.querySelector("#sec-ik-count-label");function I(){let _=o.filter(D=>{let P=n==="all"||D.category===n,N=i.toLowerCase().trim(),k=!N||D.title.toLowerCase().includes(N)||D.target.toLowerCase().includes(N)||D.tools.some(T=>T.toLowerCase().includes(N))||D.steps.some(T=>T.toLowerCase().includes(N));return P&&k});if(F&&(F.textContent=`Menampilkan ${_.length} dari ${o.length} Instruksi Kerja`),_.length===0){$.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}$.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${_.map(D=>`
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
                    ${D.tools.slice(0,3).map(P=>`<li style="margin-bottom:2px;">${x(P)}</li>`).join("")}
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
      `,$.querySelectorAll(".btn-view-sec-ik").forEach(D=>{D.addEventListener("click",()=>{let P=parseInt(D.dataset.id,10),N=o.find(k=>k.id===P);N&&C(N)})}),$.querySelectorAll(".btn-edit-sec-ik").forEach(D=>{D.addEventListener("click",()=>{let P=parseInt(D.dataset.id,10),N=o.find(k=>k.id===P);N&&b(N)})}),$.querySelectorAll(".btn-delete-sec-ik").forEach(D=>{D.addEventListener("click",()=>{let P=parseInt(D.dataset.id,10),N=o.find(k=>k.id===P);N&&ae(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${N.title}"?`,()=>{o=o.filter(k=>k.id!==P),It(o),z("Instruksi Kerja Security berhasil dihapus"),a()},"Hapus Instruksi Kerja")})})}I();let L=p.querySelector("#sec-ik-search");L&&L.addEventListener("input",_=>{i=_.target.value,I()}),p.querySelectorAll(".sec-cat-pill").forEach(_=>{_.addEventListener("click",()=>{p.querySelectorAll(".sec-cat-pill").forEach(D=>D.classList.remove("active")),_.classList.add("active"),n=_.dataset.cat,I()})});let M=p.querySelector("#btn-add-sec-ik");M&&M.addEventListener("click",()=>{b()})}function C(p){Z({title:`${p.icon||"\u{1F6E1}\uFE0F"} IK-${String(p.id).padStart(2,"0")}: ${p.title}`,size:"lg",content:`
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
                ${p.tools.map(u=>{if(u.includes("http")){let $=u.split(": ");return`<li>${$[0]}: <a href="${$.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${x(u)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${p.steps.map(u=>`<li style="margin-bottom:6px;">${x(u)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${p.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${x(p.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(u,$)=>{$()}})}function b(p=null){let u=!!p,$=p?p.tools.join(`
`):"",F=p?p.steps.join(`
`):"",I=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${p?x(p.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
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
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${p?x(p.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${x($)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${x(F)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${p?x(p.notes||""):""}</textarea>
        </div>
      </form>
    `;Z({title:u?`\u270F\uFE0F Edit IK-${p.id}: ${p.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:I,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(L,M)=>{let _=L.querySelector("#form-sec-ik-title").value.trim(),D=L.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",P=L.querySelector("#form-sec-ik-cat").value,N=L.querySelector("#form-sec-ik-target").value.trim(),k=L.querySelector("#form-sec-ik-tools").value.trim(),T=L.querySelector("#form-sec-ik-steps").value.trim(),A=L.querySelector("#form-sec-ik-notes").value.trim();if(!_||!N||!k||!T){O("Harap lengkapi semua kolom bertanda bintang (*)");return}let H=k.split(`
`).map(j=>j.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),R=T.split(`
`).map(j=>j.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),K={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[P]||{color:"#1E40AF",bg:"#EFF6FF"};if(u)p.title=_,p.icon=D,p.category=P,p.target=N,p.tools=H,p.steps=R,p.notes=A,p.color=K.color,p.bg=K.bg,It(o),z(`Instruksi Kerja "${_}" berhasil diperbarui`);else{let j=o.length>0?Math.max(...o.map(de=>de.id))+1:1,ie={id:j,title:_,category:P,icon:D,color:K.color,bg:K.bg,target:N,tools:H,steps:R,notes:A};o.push(ie),It(o),z(`Instruksi Kerja baru IK-${j} berhasil ditambahkan`)}M(),a()}})}function f(p){p.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${m==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${m==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let u=p.querySelector("#zonasi-content-container");m==="cabang"?u.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${xa.cabang.map($=>`
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
                  ${$.rooms.map(F=>`<li>${F}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:u.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${xa.hq.map($=>`
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
                    ${$.hijau.map(F=>`<li>${F}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${$.kuning.map(F=>`<li>${F}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${$.merah.map(F=>`<li>${F}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,p.querySelector("#view-cabang-btn").addEventListener("click",()=>{m="cabang",f(p)}),p.querySelector("#view-hq-btn").addEventListener("click",()=>{m="hq",f(p)})}function c(p){p.innerHTML=`
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
    `;let u=p.querySelector("#sec-asset-table-container"),$=p.querySelector("#sec-asset-search");function F(){let L=r.toLowerCase().trim(),M=l.filter(_=>!L||_.item.toLowerCase().includes(L)||_.type.toLowerCase().includes(L)||_.description.toLowerCase().includes(L)||_.pic&&_.pic.toLowerCase().includes(L));u.innerHTML=`
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
                ${M.map((_,D)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${D%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||D+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${I(_.type)}">${x(_.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${x(_.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${x(_.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${x(_.pic||"-")}</td>
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
      `,u.querySelectorAll(".btn-edit-asset").forEach(_=>{_.addEventListener("click",()=>{let D=parseInt(_.dataset.no,10),P=l.find(N=>N.no===D);P&&v(P)})}),u.querySelectorAll(".btn-delete-asset").forEach(_=>{_.addEventListener("click",()=>{let D=parseInt(_.dataset.no,10),P=l.find(N=>N.no===D);P&&ae(`Apakah Anda yakin ingin menghapus aset: "${P.item}"?`,()=>{l=l.filter(N=>N.no!==D),l.forEach((N,k)=>{N.no=k+1}),Mt(l),z("Aset berhasil dihapus"),a()},"Hapus Aset")})})}function I(L){return L==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":L==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}F(),$&&$.addEventListener("input",L=>{r=L.target.value,F()}),p.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{v()})}function v(p=null){let u=!!p,$=`
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
          <input type="text" id="form-asset-item" class="form-control" required value="${p?x(p.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${p?x(p.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${p?x(p.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;Z({title:u?`\u270F\uFE0F Edit Aset: ${p.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:$,confirmText:"Simpan",onConfirm:(F,I)=>{let L=F.querySelector("#form-asset-type").value,M=F.querySelector("#form-asset-item").value.trim(),_=F.querySelector("#form-asset-desc").value.trim(),D=F.querySelector("#form-asset-pic").value.trim();if(!M||!_||!D){O("Harap lengkapi semua kolom");return}if(u)p.type=L,p.item=M,p.description=_,p.pic=D,Mt(l),z(`Aset "${M}" berhasil diperbarui`);else{let P=l.length>0?Math.max(...l.map(N=>N.no))+1:1;l.push({no:P,type:L,item:M,description:_,pic:D}),Mt(l),z(`Aset baru "${M}" berhasil ditambahkan`)}I(),a()}})}function s(p){p.innerHTML=`
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
    `}function y(p){p.innerHTML=`
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
    `}async function g(){try{let{downloadExcel:p}=await Promise.resolve().then(()=>(Y(),he)),u=o.map(F=>({"Kode IK":`IK-SEC-${String(F.id).padStart(2,"0")}`,"Judul Prosedur":F.title,Kategori:F.category,"Sasaran Target":F.target,"Alat & Perlengkapan":F.tools.join(", "),"Langkah Pelaksanaan":F.steps.map((I,L)=>`${L+1}. ${I}`).join(" | "),"Catatan Khusus":F.notes||"-"})),$=new Date().toISOString().slice(0,10);p(u,`SOP_Security_Instruksi_Kerja_${$}`),z("Data Excel SOP Security berhasil diunduh")}catch(p){console.error(p),O("Gagal mendownload data Excel: "+p.message)}}function x(p){return p?String(p).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}a()}U();ue();var Sa=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Kt(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Ot(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function Ea(e){let t=[],i="all",n="",r=localStorage.getItem("fcms_doc_view")||"grid",m="fcms-master-doc-styles";if(!document.getElementById(m)){let b=document.createElement("style");b.id=m,b.textContent=`
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
    `,document.head.appendChild(b)}async function o(){try{let b=await B("/api/checklist?limit=1000");if(b.ok&&b.data?.data&&Array.isArray(b.data.data)&&b.data.data.length>0)t=b.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t));else{let f=localStorage.getItem("fcms_master_doc_cache");f?t=JSON.parse(f):(t=[...Sa],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)))}}catch(b){console.warn("Fallback loading documents from cache/seed:",b);let f=localStorage.getItem("fcms_master_doc_cache");t=f?JSON.parse(f):[...Sa]}h()}function l(){return t.filter(b=>{let f=i==="all"||b.category===i,c=n.toLowerCase().trim(),v=!c||(b.name||"").toLowerCase().includes(c)||(b.category||"").toLowerCase().includes(c)||(b.description||"").toLowerCase().includes(c);return f&&v})}function a(){let b={all:t.length};return t.forEach(f=>{let c=f.category||"Lainnya";b[c]=(b[c]||0)+1}),b}function d(b=null){let f=!!b,c=Array.from(new Set(t.map(s=>s.category).filter(Boolean))),v=document.createElement("div");v.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${b?.name?String(b.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${b?.category?String(b.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${c.map(s=>`<option value="${s}">${s}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${b?.document_link?String(b.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${b?.description||""}</textarea>
        </div>
      </div>
    `,Z({title:f?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:v,size:"md",confirmText:f?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(s,y)=>{let g=v.querySelector("#doc-modal-name"),x=v.querySelector("#doc-modal-category"),p=v.querySelector("#doc-modal-link"),u=v.querySelector("#doc-modal-desc"),$=(g.value||"").trim(),F=(x.value||"").trim(),I=(p.value||"").trim(),L=(u.value||"").trim();if(!$){O("Nama Dokumen wajib diisi"),g.focus();return}if(!F){O("Kategori Dokumen wajib diisi"),x.focus();return}let M={name:$,category:F,document_link:I,description:L};try{if(f)if((await B(`/api/checklist/${b.id}`,{method:"PUT",body:JSON.stringify(M)})).ok){let D=t.findIndex(P=>P.id===b.id);D!==-1&&(t[D]={...t[D],...M}),z("Dokumen berhasil diperbarui")}else{let D=t.findIndex(P=>P.id===b.id);D!==-1&&(t[D]={...t[D],...M}),z("Dokumen disimpan secara lokal")}else{let _=await B("/api/checklist",{method:"POST",body:JSON.stringify(M)});if(_.ok&&_.data?.id)t.unshift({id:_.data.id,...M}),z("Dokumen baru berhasil ditambahkan");else{let D=Date.now();t.unshift({id:D,...M}),z("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),y(),h()}catch(_){console.error("Error saving document:",_),O("Gagal menyimpan dokumen: "+_.message)}}})}function S(b){ae(`Apakah Anda yakin ingin menghapus dokumen <strong>"${b.name}"</strong>?`,async()=>{try{let f=await B(`/api/checklist/${b.id}`,{method:"DELETE"});f.ok||console.warn("API delete warning:",f),t=t.filter(c=>c.id!==b.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),z(`Dokumen "${b.name}" berhasil dihapus`),h()}catch{t=t.filter(c=>c.id!==b.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),z(`Dokumen "${b.name}" berhasil dihapus`),h()}},"Hapus Dokumen")}function E(b,f){if(!b){vt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(b).then(()=>{z(`Tautan "${f}" disalin ke clipboard!`)}).catch(()=>{vt(`Tautan: ${b}`)})}async function w(){try{let{downloadExcel:b}=await Promise.resolve().then(()=>(Y(),he)),f=t.map((c,v)=>({No:v+1,"Nama Dokumen":c.name||"",Kategori:c.category||"","Platform / Format":Ot(c.document_link).type,"Link Dokumen":c.document_link||"","Deskripsi / Keterangan":c.description||""}));b(f,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),z("Berhasil mengunduh data dokumen ke Excel.")}catch(b){O("Gagal mengekspor data: "+b.message)}}function h(){let b=l(),f=a(),c=Object.keys(f).filter(s=>s!=="all"),v=t.filter(s=>!!s.document_link).length;e.innerHTML=`
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
                <div class="doc-stat-value">${c.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${v} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${t.length}</span></div>
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
              <button id="doc-view-grid" class="doc-view-btn ${r==="grid"?"active":""}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
                <span>Grid Kartu</span>
              </button>
              <button id="doc-view-table" class="doc-view-btn ${r==="table"?"active":""}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel Data</span>
              </button>
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${i==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${f.all||0}</span>
            </button>
            ${c.map(s=>{let y=Kt(s);return`
                <button class="doc-pill ${i===s?"active":""}" data-cat="${s}">
                  <span>${y.icon}</span>
                  <span>${s}</span>
                  <span class="doc-pill-count">${f[s]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${b.length===0?`
            <div class="doc-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Dokumen Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Coba sesuaikan kata kunci pencarian atau pilih kategori lain.</p>
              <button id="doc-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Tampilkan Semua Dokumen
              </button>
            </div>
          `:r==="grid"?`
            <!-- Grid Card View -->
            <div class="doc-cards-grid">
              ${b.map(s=>{let y=Kt(s.category),g=Ot(s.document_link);return`
                  <div class="doc-card" data-id="${s.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border};">
                          <span>${y.icon}</span>
                          <span>${y.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${g.icon}</span>
                          <span>${g.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${s.name}</h3>
                      <p class="doc-card-desc">${s.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${s.document_link?`
                        <a href="${s.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${s.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${s.document_link}" data-title="${s.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${s.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${s.id}">
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
                  ${b.map((s,y)=>{let g=Kt(s.category),x=Ot(s.document_link);return`
                      <tr>
                        <td style="text-align:center; font-weight:600; color:#94A3B8;">${y+1}</td>
                        <td>
                          <div style="font-weight:700; color:#0F172A;">${s.name}</div>
                        </td>
                        <td>
                          <span class="doc-category-badge" style="background:${g.bg}; color:${g.color}; border:1px solid ${g.border};">
                            <span>${g.icon}</span>
                            <span>${g.label}</span>
                          </span>
                        </td>
                        <td>
                          <span class="doc-platform-badge">
                            <span>${x.icon}</span>
                            <span>${x.type}</span>
                          </span>
                        </td>
                        <td>
                          ${s.document_link?`
                            <div style="display:flex; align-items:center; gap:6px;">
                              <a href="${s.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-primary" style="font-size:0.78rem; padding:5px 10px; border-radius:6px; display:inline-flex; align-items:center; gap:4px; text-decoration:none;">
                                <span>Buka Dokumen</span>
                                <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                              </a>
                              <button class="doc-icon-btn copy-link-btn" title="Salin Link" style="width:28px; height:28px;" data-link="${s.document_link}" data-title="${s.name}">
                                <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                              </button>
                            </div>
                          `:'<span style="color:#94A3B8; font-size:0.82rem;">-</span>'}
                        </td>
                        <td style="max-width:280px; color:#64748B; font-size:0.84rem;">
                          ${s.description||"-"}
                        </td>
                        <td style="text-align:center;">
                          <div style="display:flex; justify-content:center; gap:6px;">
                            <button class="doc-icon-btn edit-doc-btn" style="width:30px; height:30px;" title="Edit" data-id="${s.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                            <button class="doc-icon-btn delete delete-doc-btn" style="width:30px; height:30px;" title="Hapus" data-id="${s.id}">
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
    `,C()}function C(){let b=e.querySelector("#doc-btn-add");b&&(b.onclick=()=>d());let f=e.querySelector("#doc-btn-excel");f&&(f.onclick=()=>w());let c=e.querySelector("#doc-btn-refresh");c&&(c.onclick=()=>o());let v=e.querySelector("#doc-btn-reset-filter");v&&(v.onclick=()=>{n="",i="all",h()});let s=e.querySelector("#doc-search-input");s&&(s.oninput=x=>{n=x.target.value;let p=l();h();let u=e.querySelector("#doc-search-input");u&&(u.focus(),u.setSelectionRange(u.value.length,u.value.length))});let y=e.querySelector("#doc-view-grid"),g=e.querySelector("#doc-view-table");y&&(y.onclick=()=>{r="grid",localStorage.setItem("fcms_doc_view","grid"),h()}),g&&(g.onclick=()=>{r="table",localStorage.setItem("fcms_doc_view","table"),h()}),e.querySelectorAll(".doc-pill").forEach(x=>{x.onclick=()=>{i=x.getAttribute("data-cat")||"all",h()}}),e.querySelectorAll(".copy-link-btn").forEach(x=>{x.onclick=p=>{p.stopPropagation();let u=x.getAttribute("data-link"),$=x.getAttribute("data-title");E(u,$)}}),e.querySelectorAll(".edit-doc-btn").forEach(x=>{x.onclick=p=>{p.stopPropagation();let u=Number(x.getAttribute("data-id")),$=t.find(F=>F.id===u);$&&d($)}}),e.querySelectorAll(".delete-doc-btn").forEach(x=>{x.onclick=p=>{p.stopPropagation();let u=Number(x.getAttribute("data-id")),$=t.find(F=>F.id===u);$&&S($)}})}await o()}U();ue();Y();var Ne={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function jt(e){for(let[t,i]of Object.entries(Ne))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function Ca(e){e.innerHTML=`
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
        ${Object.entries(Ne).map(([s,y])=>`
          <button class="zone-tab-btn" data-zone="${s}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${y.icon}</span> ${s}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${s.replace(/[\s&]+/g,"-")}">-</span>
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
  `;let t="all",i="table",n=[],r=oe(),m=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let s=await B("/api/hygiene-standards?all=1");if(!s.ok)throw new Error("Gagal memuat standar hygiene");n=Array.isArray(s.data)?s.data:s.data?.data||[],a(),l(),S()}catch(s){console.error(s),O(s.message||"Terjadi kesalahan memuat data");let y=e.querySelector("#hygiene-view-container");y&&(y.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${s.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,y.querySelector("#btn-retry-hygiene")?.addEventListener("click",o))}}function l(){let s=e.querySelector("#badge-count-all");s&&(s.textContent=n.length),Object.entries(Ne).forEach(([x,p])=>{let u=x.replace(/[\s&]+/g,"-"),$=e.querySelector(`#badge-count-${u}`);if($){let F=n.filter(I=>p.rooms.includes(I.room_name)).length;$.textContent=F}});let y=new Set(n.map(x=>x.room_name)),g=e.querySelector("#card-total-rooms");g&&(g.textContent=y.size)}function a(){let s=e.querySelector("#hygiene-room-filter");if(!s)return;let y=s.value,g=[];if(t==="all")g=Array.from(new Set(n.map(x=>x.room_name))).filter(Boolean).sort();else{let x=Ne[t];x&&(g=x.rooms.filter(p=>n.some(u=>u.room_name===p)))}s.innerHTML=`<option value="">Semua Ruangan (${g.length})</option>`+g.map(x=>`<option value="${x}" ${x===y?"selected":""}>${x}</option>`).join("")}function d(){let s=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),y=e.querySelector("#hygiene-room-filter")?.value||"";return n.filter(g=>{if(t!=="all"){let x=Ne[t];if(!x||!x.rooms.includes(g.room_name))return!1}if(y&&g.room_name!==y)return!1;if(s){let x=(g.room_name||"").toLowerCase(),p=(g.item_name||"").toLowerCase(),u=(g.cleanliness_standard||"").toLowerCase();if(!x.includes(s)&&!p.includes(s)&&!u.includes(s))return!1}return!0})}function S(){let s=d(),y=e.querySelector("#hygiene-stats-text");y&&(y.innerHTML=`Menampilkan <strong>${s.length}</strong> dari <strong>${n.length}</strong> parameter`);let g=e.querySelector("#hygiene-view-container");if(g){if(s.length===0){g.innerHTML=`
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
      `,g.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let x=e.querySelector("#hygiene-search"),p=e.querySelector("#hygiene-room-filter");x&&(x.value=""),p&&(p.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(u=>u.classList.toggle("active",u.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(u=>{u.dataset.zone==="all"?(u.style.background="#2563EB",u.style.color="#fff"):(u.style.background="#F1F5F9",u.style.color="var(--text-2)")}),a(),S()});return}i==="table"?E(g,s):w(g,s)}}function E(s,y){s.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${m?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${y.map((g,x)=>{let p=jt(g.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${x+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${p.bg}; color:${p.color}; border:1px solid ${p.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${p.icon}</span> ${g.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${p.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${g.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${g.cleanliness_standard}
                      </div>
                    </td>
                    ${m?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${g.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${g.id}" data-name="${g.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,h(s)}function w(s,y){let g={};y.forEach(x=>{let p=jt(x.room_name);g[p.zoneName]||(g[p.zoneName]={info:p,rooms:{}}),g[p.zoneName].rooms[x.room_name]||(g[p.zoneName].rooms[x.room_name]=[]),g[p.zoneName].rooms[x.room_name].push(x)}),s.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(g).map(([x,p])=>{let u=Object.values(p.rooms).reduce(($,F)=>$+F.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${p.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${p.info.bg}; border-bottom:1px solid ${p.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${p.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${p.info.color};">
                      ${x}
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
                ${Object.entries(p.rooms).map(([$,F])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${$}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${F.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${m?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${F.map((I,L)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${L+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${I.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${I.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${m?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${I.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${I.id}" data-name="${I.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,h(s)}function h(s){s.querySelectorAll(".btn-edit-item").forEach(y=>{y.addEventListener("click",()=>{let g=y.dataset.id,x=n.find(p=>String(p.id)===String(g));x&&C(x)})}),s.querySelectorAll(".btn-delete-item").forEach(y=>{y.addEventListener("click",()=>{let g=y.dataset.id,x=y.dataset.name;ae(`Yakin ingin menghapus parameter standar <strong>"${x}"</strong>?`,async()=>{try{let p=await B(`/api/hygiene-standards/${g}`,{method:"DELETE"});if(!p.ok)throw new Error(p.data?.error||"Gagal menghapus parameter");z("Parameter kebersihan berhasil dihapus"),o()}catch(p){O(p.message)}},"Hapus Parameter Standar")})})}function C(s=null){let y=!!s,g=Array.from(new Set(n.map(p=>p.room_name))).filter(Boolean).sort(),x=document.createElement("div");x.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${s?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${g.map(p=>`<option value="${p}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${s?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${s?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,Z({title:y?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:x,confirmText:y?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(p,u)=>{let $=x.querySelector("#modal-room-name").value.trim(),F=x.querySelector("#modal-item-name").value.trim(),I=x.querySelector("#modal-standard").value.trim();if(!$||!F||!I){O("Semua kolom wajib diisi");return}try{let L={room_name:$,item_name:F,cleanliness_standard:I},M=y?`/api/hygiene-standards/${s.id}`:"/api/hygiene-standards",D=await B(M,{method:y?"PUT":"POST",body:JSON.stringify(L)});if(!D.ok)throw new Error(D.data?.error||"Gagal menyimpan parameter");z(y?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),u(),o()}catch(L){O(L.message)}}})}function b(){let s=window.open("","_blank");if(!s){O("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let g=`
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

        ${Object.entries(Ne).map(([x,p])=>{let u=n.filter($=>p.rooms.includes($.room_name));return u.length===0?"":`
            <div class="zone-title">${p.icon} ${x.toUpperCase()} (${u.length} PARAMETER)</div>
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
                ${u.map(($,F)=>`
                  <tr>
                    <td class="td-no">${F+1}</td>
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
    `;s.document.open(),s.document.write(g),s.document.close()}function f(){let s=n.map((g,x)=>{let p=jt(g.room_name);return{No:x+1,"Kategori Zona":p.zoneName,"Ruangan / Lokasi":g.room_name,"Item / Objek Inspeksi":g.item_name,"Standar Kebersihan Fisik":g.cleanliness_standard}}),y=new Date().toISOString().slice(0,10);q(s,`Master_Standar_Kebersihan_FCMS_${y}`),z("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>S()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>S()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let s=e.querySelector("#hygiene-search"),y=e.querySelector("#hygiene-room-filter");s&&(s.value=""),y&&(y.value=""),S()}),e.querySelectorAll(".zone-tab-btn").forEach(s=>{s.addEventListener("click",()=>{t=s.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(y=>{let g=y.dataset.zone===t;if(y.classList.toggle("active",g),g){y.style.background="#2563EB",y.style.color="#fff";let x=y.querySelector(".badge");x&&(x.style.background="rgba(255,255,255,0.25)",x.style.color="#fff")}else{y.style.background="#F1F5F9",y.style.color="var(--text-2)";let x=y.querySelector(".badge");x&&(x.style.background="#CBD5E1",x.style.color="#334155")}}),a(),S()})});let c=e.querySelector("#view-table-btn"),v=e.querySelector("#view-zone-btn");c?.addEventListener("click",()=>{i="table",c.style.background="#fff",c.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",v.style.background="transparent",v.style.boxShadow="none",S()}),v?.addEventListener("click",()=>{i="zone",v.style.background="#fff",v.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",c.style.background="transparent",c.style.boxShadow="none",S()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",b),e.querySelector("#btn-export-hygiene")?.addEventListener("click",f),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>C(null)),o()}U();ue();Y();async function $a(e,t="supply"){return Sn(e)}async function Sn(e){let i=((await B("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));G({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>te(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let r=n?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let m=n?.chemical_items;try{m=Array.isArray(JSON.parse(m))?JSON.parse(m).join(", "):m}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!i.find(o=>o.value==n.branch_id)?[...i,{value:n.branch_id,label:n.branch_name||n.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:m},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),m=await B(`/api/reports/supply?limit=10000&${r}`);if(m.ok){let o=m.data.data.map(l=>{let a=l.tools_items;try{a=Array.isArray(JSON.parse(a))?JSON.parse(a).join(", "):a}catch{}let d=l.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":a||"",Chemical:d||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});q(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let m=(await B("/api/branches?all=1")).data?.data||[],o=S=>{if(!S)return null;let E=String(S||"").toLowerCase(),w=m.find(h=>String(h.full_name||"").toLowerCase()===E||String(h.code||"").toLowerCase()===E||String(h.name||"").toLowerCase()===E);return w?w.id:null},l=S=>{if(S==null||S==="")return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let E=String(S).trim();if(E===""||E==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(E))return E.slice(0,10);if(/^\d{4,5}$/.test(E)){let h=Number(E);if(h>2e4&&h<99999){let C=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let w=E.split(/[\/\-\.]/);if(w.length===3){let[h,C,b]=w.map(f=>f.trim());if(h.length===4&&C.length<=2&&b.length<=2)return`${h}-${C.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&C.length<=2&&h.length<=2)return`${b}-${C.padStart(2,"0")}-${h.padStart(2,"0")}`}return E},a=n.map(S=>({submitted_at:l(S.Waktu||S.Tanggal),submitter_name:String(S.Pengirim||"").trim(),branch_id:o(String(S.Cabang||"").trim()),tools_items:String(S["Alat/Barang"]||S.Alat||"").trim(),chemical_items:String(S.Chemical||"").trim(),additional_notes:String(S.Catatan||S.Keterangan||"").trim(),status:String(S.Status||"").trim(),processed_by:String(S["Diproses Oleh"]||S.PIC||"").trim()})).filter(S=>S.submitted_at&&S.submitter_name&&S.branch_id),d=await B("/api/reports/supply/import",{method:"POST",body:JSON.stringify(a)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,r)=>{let m=Z({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,l)=>{let a=o.querySelector("#supply-status").value,d=o.querySelector("#supply-processed-by").value;(await B(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:a,processed_by:d})})).ok?(z("Status diperbarui."),l(),r()):O("Gagal update status.")}})}}]})}U();Y();async function Fa(e){let t=oe();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}G({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let n=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await B(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let n=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));q(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let n=i.map(m=>({full_name:String(m["Nama Lengkap"]||"").trim(),username:String(m.Username||"").trim(),email:String(m.Email||"").trim(),role:String(m.Role||"").trim()||"viewer",password:String(m.Password||"").trim()})).filter(m=>m.username&&m.password&&m.email&&m.full_name),r=await B("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}U();Y();async function _a(e){G({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await B(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)q(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{q([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),n=await B("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}U();async function Ta(e){let t=new Date,i=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",r));async function n(){try{let o=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await B(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),i=[]}}async function r(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let l=t.getFullYear(),a=t.getMonth(),d=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),S=document.getElementById("cal-month-label");S&&(S.textContent=d);let E=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(g=>g.value)),w=i.filter(g=>E.has(g.type)),h={};w.forEach(g=>{let x=En(g.event_date);x&&(h[x]||(h[x]=[]),h[x].push(g))});let C=new Date(l,a,1).getDay(),b=new Date(l,a+1,0).getDate(),f=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],c=new Date().toISOString().slice(0,10),v='<div class="calendar-grid">';f.forEach(g=>{v+=`<div class="cal-day-header">${g}</div>`});for(let g=0;g<C;g++)v+='<div class="cal-cell cal-cell-empty"></div>';for(let g=1;g<=b;g++){let x=`${l}-${String(a+1).padStart(2,"0")}-${String(g).padStart(2,"0")}`,p=h[x]||[],u=x===c;v+=`
          <div class="cal-cell ${u?"cal-today":""} ${p.length?"cal-has-events":""}"
               data-date="${x}" tabindex="0" role="button" aria-label="${x}">
            <div class="cal-day-num ${u?"today-num":""}">${g}</div>
            <div class="cal-events-preview">
              ${p.slice(0,3).map($=>`
                <div class="cal-event-dot cal-color-${$.color||"gray"}" title="${we($.title||$.type)}">
                  <span class="cal-event-dot-label">${Cn($.title||$.branch_name||$.type,18)}</span>
                </div>
              `).join("")}
              ${p.length>3?`<div class="cal-more">+${p.length-3} lagi</div>`:""}
            </div>
          </div>`}let y=(C+b)%7;if(y!==0)for(let g=0;g<7-y;g++)v+='<div class="cal-cell cal-cell-empty"></div>';v+="</div>",o.innerHTML=v,o.querySelectorAll(".cal-cell[data-date]").forEach(g=>{g.addEventListener("click",()=>{let x=g.dataset.date,p=h[x]||[];if(!p.length)return;let u=document.getElementById("cal-event-list"),$=new Date(x+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=$,document.getElementById("cal-event-items").innerHTML=p.map(F=>`
            <div class="cal-event-item cal-color-border-${F.color||"gray"}">
              <div class="cal-event-type">
                ${$n(F.type)}
                ${F.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${we(F.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${we(F.title||"-")}</div>
              ${F.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${we(F.branch_name)}</div>`:""}
              ${F.pic||F.period||F.opening_date||F.target_date||F.completion_date||F.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${F.pic?`<div>\u{1F464} <b>PIC:</b> ${we(F.pic)}</div>`:""}
                  ${F.period?`<div>\u{1F4C5} <b>Periode:</b> ${we(F.period)}</div>`:""}
                  ${F.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${we(Rt(F.opening_date))}</div>`:""}
                  ${F.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${we(Rt(F.target_date))}</div>`:""}
                  ${F.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${we(Rt(F.completion_date))}</div>`:""}
                  ${F.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${we(F.notes)}</div>`:""}
                </div>
              `:""}
              ${F.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${F.status==="Done"?"badge-success":"badge-warning"}">${we(F.status)}</span></div>`:""}
              ${F.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${F.days_remaining} hari</div>`:""}
            </div>
          `).join(""),u.style.display="block"})})}catch(l){console.error("[Calendar] Render error:",l),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let m=()=>{n().then(()=>r())};it.on("data:changed",m),r()}function En(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[n,r,m]=i.map(o=>o.trim());if(m.length===4)return`${m}-${r.padStart(2,"0")}-${n.padStart(2,"0")}`;if(n.length===4)return`${n}-${r.padStart(2,"0")}-${m.padStart(2,"0")}`}return t.slice(0,10)}function Rt(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[n,r,m]=i.map(o=>o.trim());if(n.length===4)return`${m.padStart(2,"0")}-${r.padStart(2,"0")}-${n}`;if(m.length===4)return`${n.padStart(2,"0")}-${r.padStart(2,"0")}-${m}`}return e}function Cn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function we(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function $n(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}U();async function Aa(e){let t=oe(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
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
  `;let m=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(m&&o)try{let l=JSON.parse(atob(m.split(".")[1])),a=new Date(l.exp*1e3);o.textContent=`Berakhir: ${a.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let a=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),S=document.getElementById("btn-save-pwd");a.style.display="none",d.style.display="none";let E=l.target,w=E.current_password.value,h=E.new_password.value,C=E.confirm_password.value;if(h!==C){a.textContent="\u274C Konfirmasi password tidak cocok.",a.style.display="block";return}if(h.length<6){a.textContent="\u274C Password baru minimal 6 karakter.",a.style.display="block";return}S.disabled=!0,S.textContent="\u23F3 Menyimpan...";let b=await B("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:w,new_password:h})});S.disabled=!1,S.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',b.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",E.reset(),z("Password berhasil diubah.")):(a.textContent=b.data?.error||"Gagal mengubah password.",a.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}U();var gt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function se(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let r=Number(t);if(r>2e4&&r<99999){let m=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(m.getTime())?null:m.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[r,m,o]=i.map(S=>S.trim()),l=Number(r),a=Number(m),d=Number(o);if(r.length===4&&l>1900)return`${r}-${m.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&d>1900)return l>12?`${o}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`:a>12?`${o}-${r.padStart(2,"0")}-${m.padStart(2,"0")}`:`${o}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`;if(o.length===2&&!isNaN(d)){let S=d>=50?`19${o}`:`20${o}`;return l>12?`${S}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`:`${S}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`}}let n=new Date(t);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Da(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Fn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:se(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:se(e["Target Pindah OS"]),target_selesai_os:se(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:se(e["Tanggal Mulai"]),end_date:se(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:se(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:se(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:se(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:se(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:se(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:se(e["Tanggal Target"]||e["Tgl Target"]),completion_date:se(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:se(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:se(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:se(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:se(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:se(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:se(e["Tanggal Back Up"]),completion_date:se(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:se(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:se(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function _n(e,t){let i=gt[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let n=Fn[i.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],m=[],o=[];return t.filter(a=>!Da(a)).forEach((a,d)=>{let S=t.indexOf(a)+2,E=[];n.required.forEach(({key:h,label:C})=>{let b=a[h];if(b==null||String(b).trim()===""){let f=Object.keys(a).filter(c=>c.trim()).join(", ");E.push({column:C,originalValue:b||"",reason:`Kolom "${C}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${f.slice(0,120)}`})}});let w=n.map(a);E.length>0?m.push({row:S,data:w,raw:a,errors:E}):(r.push(a),o.push(w))}),{valid:r,errors:m,mapped:o}}function Ba(e){let t=[];return e.SheetNames.forEach(i=>{let n=gt[i];if(!n)return;let r=e.Sheets[i],m=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=_n(i,m),l=m.filter(a=>!Da(a));t.push({sheetName:i,module:n.module,label:n.label,total:l.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),t}function Pa(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([n,r])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(r),n)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function La(e){let t=window.XLSX,i=t.utils.book_new(),n=!1;return e.forEach(r=>{if(!r.errors||r.errors.length===0)return;n=!0;let m=r.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(a=>a.column||a).join("; "),"Alasan Error":(l.errors||[]).map(a=>a.reason||a).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([a,d])=>[a,d??""]))})),o=t.utils.json_to_sheet(m);t.utils.book_append_sheet(i,o,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Tn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function Ia(e){e.innerHTML=`
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
              ${Object.entries(gt).map(([b,{label:f}])=>`<span class="import-sheet-tag">\u{1F4C4} ${b} \u2192 ${f}</span>`).join("")}
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
  `;let t=null,i=null,n=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function m(b){Object.entries(r).forEach(([f,c])=>{c.style.display=f===b?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let b=document.getElementById("btn-backup-db");b.disabled=!0,b.textContent="\u23F3 Memproses Backup...";try{let f=await B("/api/import/backup");if(f.ok){if(!window.XLSX){O("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let c=window.XLSX,v=c.utils.book_new();Object.entries(f.data.database).forEach(([s,y])=>{let g=y.length>0?y:[{}],x=c.utils.json_to_sheet(g);c.utils.book_append_sheet(v,x,s.substring(0,31))}),c.writeFile(v,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),z("Backup berhasil diunduh!")}else O("Gagal memproses backup: "+(f.data?.error||"Unknown error"))}catch(f){O("Gagal memproses backup: "+f.message)}finally{b.disabled=!1,b.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let b=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let f=await B("/api/sync/google-sheets",{method:"POST"});f.ok?alert("Sinkronisasi Berhasil: "+(f.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(f.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=b,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Pa(),z("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),a=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",b=>{b.stopPropagation(),l.click()}),l.addEventListener("change",b=>{b.target.files[0]&&d(b.target.files[0])}),a.addEventListener("dragover",b=>{b.preventDefault(),a.classList.add("drag-over")}),a.addEventListener("dragleave",()=>a.classList.remove("drag-over")),a.addEventListener("drop",b=>{b.preventDefault(),a.classList.remove("drag-over");let f=b.dataTransfer.files[0];f&&f.name.match(/\.xlsx?$/i)?d(f):O("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,l.value="",document.getElementById("file-info").style.display="none",a.style.display="",m("upload")});async function d(b){t=b,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${b.name} (${(b.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",a.style.display="none",await S(b)}async function S(b){m("validating");let f=document.getElementById("validation-status"),c=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");f.textContent="Membaca file Excel...",c.style.width="20%",await Qe(200);let v=await b.arrayBuffer(),s=window.XLSX.read(v,{type:"array",cellDates:!0});f.textContent=`Memvalidasi ${s.SheetNames.length} sheet...`,c.style.width="50%",await Qe(100),i=Ba(s),c.style.width="100%",f.textContent="Validasi selesai!",await Qe(300),E()}catch(v){m("upload"),O("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",a.style.display="none"}}function E(){m("preview");let b=i.filter(u=>!u.skipped).length,f=i.reduce((u,$)=>u+$.total,0),c=i.reduce((u,$)=>u+$.valid,0),v=i.reduce((u,$)=>u+$.errorCount,0),s=f>0?Math.round(c/f*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${b} sheet</span>
      <span class="badge badge-secondary">${f} baris</span>
      <span class="badge badge-success">${c} valid (${s}%)</span>
      ${v>0?`<span class="badge badge-danger">${v} error</span>`:""}
    `;let y=document.getElementById("preview-table-container");y.innerHTML=`
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
    `,y.querySelectorAll(".btn-detail-error").forEach(u=>{u.addEventListener("click",()=>{let $=i[Number(u.dataset.idx)];w($)})});let g=document.getElementById("error-detail-section"),x=document.getElementById("error-detail-container");x.innerHTML="",g.style.display="none";let p=document.getElementById("btn-start-import");c===0?(p.disabled=!0,p.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(p.disabled=!1,v>0?(p.innerHTML=`\u{1F680} Import ${c} Data Valid (${v} dilewati)`,p.title="Baris error akan dilewati, baris valid tetap diimport"):p.innerHTML=`\u{1F680} Mulai Import ${c} Data`)}function w(b){let f=document.getElementById("error-detail-section"),c=document.getElementById("error-detail-container");f.style.display="";let v=b.errors.slice(0,100).map(s=>(Array.isArray(s.errors)?s.errors:[]).map(g=>{let x=typeof g=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${s.row}</span></td>
            <td><strong>${x?g.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${x&&g.originalValue!==void 0?g.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${x?g.reason:g}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${x&&g.aliases?`Gunakan salah satu nama kolom:<br><em>${g.aliases}</em>`:x&&g.hint?g.hint:""}
            </td>
          </tr>
        `}).join("")).join("");c.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${b.sheetName} \u2014 ${b.errorCount} baris error dari ${b.total} total
          ${b.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${v||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${b.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,f.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{m("upload"),document.getElementById("file-info").style.display="none",a.style.display="",t=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;La(i)?z("Log error berhasil didownload."):z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let b=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";h(b)});async function h(b){m("importing"),n=Date.now();let f=[];Tn.forEach(g=>{let x=i?.find(p=>p.module===g&&p.mapped?.length>0);x&&f.push(x)});let c=document.getElementById("import-steps-list");c.innerHTML=f.map(g=>`
      <div class="import-step-item" id="step-item-${g.module}">
        <span class="step-item-icon" id="step-icon-${g.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${g.label} <span class="step-item-count">(${g.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${g.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),s=document.getElementById("import-current-status"),y={totalSheets:f.length,totalRows:f.reduce((g,x)=>g+x.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let g=0;g<f.length;g++){let x=f[g],p=document.getElementById(`step-icon-${x.module}`),u=document.getElementById(`step-status-${x.module}`);p.textContent="\u{1F504}",u.textContent="Mengimport...",s.textContent=`Mengimport ${x.label}...`,v.style.width=`${Math.round(g/f.length*100)}%`;try{let $=await B(`/api/import/${x.module}`,{method:"POST",body:JSON.stringify({rows:x.mapped,onDuplicate:b})});if($.ok){let F=$.data;y.inserted+=F.inserted||0,y.skipped+=F.skipped||0,y.moduleResults.push({label:x.label,inserted:F.inserted||0,skipped:F.skipped||0,status:"ok"}),p.textContent="\u2705",u.innerHTML=`<span class="badge badge-success">${F.inserted||0} berhasil</span>${F.skipped>0?` <span class="badge badge-neutral">${F.skipped} skip</span>`:""}`}else y.failed++,y.moduleResults.push({label:x.label,inserted:0,skipped:0,status:"error",error:$.data?.error}),p.textContent="\u274C",u.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch($){y.failed++,y.moduleResults.push({label:x.label,inserted:0,skipped:0,status:"error",error:$.message}),p.textContent="\u274C",u.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Qe(150)}v.style.width="100%",s.textContent="Selesai!",await Qe(400),C(y)}function C(b){m("summary");let f=((Date.now()-n)/1e3).toFixed(1),c=b.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${c?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${c?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${b.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${b.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${b.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${b.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${b.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${b.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
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
          ${b.moduleResults.map(v=>`
            <tr>
              <td>${v.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${v.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${v.skipped}</span></td>
              <td style="text-align:center">
                ${v.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${v.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,l.value="",document.getElementById("file-info").style.display="none",a.style.display="",m("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Qe(e){return new Promise(t=>setTimeout(t,e))}U();ue();var Ma=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function Ae(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let i=new Date(e);i.setHours(0,0,0,0);let n=i.getTime()-t.getTime(),r=Math.ceil(n/(1e3*60*60*24));return r<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(r)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:r,isExpired:!0,isExpiringSoon:!1}:r<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${r} hari)`,shortLabel:`Sisa ${r} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:r,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${r} hari)`,shortLabel:`Aktif (${r} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:r,isExpired:!1,isExpiringSoon:!1}}function qt(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function ze(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Na(e){let t=[],i=[],n=[],r="all",m="all",o="all",l="",a=localStorage.getItem("fcms_sp_view")||"table",d="fcms-sp-page-styles";if(!document.getElementById(d)){let s=document.createElement("style");s.id=d,s.textContent=`
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
      .sp-view-toggle {
        display: flex;
        background: #F1F5F9;
        padding: 4px;
        border-radius: 10px;
        gap: 4px;
      }
      .sp-view-btn {
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
      .sp-view-btn.active {
        background: #FFFFFF;
        color: #1E293B;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
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
    `,document.head.appendChild(s)}async function S(){try{let[s,y]=await Promise.all([Q(),le()]);i=s||[],n=y||[];let g=await B("/api/sp?limit=1000");if(g.ok&&g.data?.data&&Array.isArray(g.data.data)&&g.data.data.length>0)t=g.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t));else{let x=localStorage.getItem("fcms_sp_data_cache");x?t=JSON.parse(x):(t=[...Ma],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)))}}catch(s){console.warn("Fallback loading SP data from cache/seed:",s);let y=localStorage.getItem("fcms_sp_data_cache");t=y?JSON.parse(y):[...Ma]}c()}function E(){return t.filter(s=>{let y=l.toLowerCase().trim(),g=!y||(s.employee_name||"").toLowerCase().includes(y)||(s.branch_name||"").toLowerCase().includes(y)||(s.division||"").toLowerCase().includes(y)||(s.sp_type||"").toLowerCase().includes(y),x=m==="all"||String(s.branch_id)===String(m),p=o==="all"||s.division===o,u=Ae(s.akhir_sp),$=!0;return r==="active"?$=!u.isExpired:r==="expiring"?$=u.isExpiringSoon:r==="expired"?$=u.isExpired:r==="sp1"?$=(s.sp_type||"").toUpperCase()==="SP 1":r==="sp2"?$=(s.sp_type||"").toUpperCase()==="SP 2":r==="sp3"?$=(s.sp_type||"").toUpperCase()==="SP 3":r==="teguran"&&($=(s.sp_type||"").toUpperCase().includes("TEGURAN")),g&&x&&p&&$})}function w(){let s=0,y=0,g=0,x=0,p=0,u=0,$=0,F=new Set;return t.forEach(I=>{I.employee_name&&F.add(I.employee_name.toLowerCase().trim());let L=Ae(I.akhir_sp);L.isExpired?g++:s++,L.isExpiringSoon&&y++;let M=(I.sp_type||"").toUpperCase();M==="SP 1"?x++:M==="SP 2"?p++:M==="SP 3"?u++:M.includes("TEGURAN")&&$++}),{total:t.length,active:s,expiring:y,expired:g,sp1:x,sp2:p,sp3:u,teguran:$,uniqueEmployees:F.size}}function h(s=null){let y=!!s,g=s?.tanggal?s.tanggal.slice(0,10):new Date().toISOString().slice(0,10),x=s?.akhir_sp?s.akhir_sp.slice(0,10):"";if(!x){let F=new Date(g);F.setMonth(F.getMonth()+6),x=F.toISOString().slice(0,10)}let p=document.createElement("div");p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${s?.employee_name?String(s.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${n.map(F=>`<option value="${F.label}">${F.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${s?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${s?.division==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">Pilih Cabang...</option>
              ${i.map(F=>`<option value="${F.value}" ${String(s?.branch_id)===String(F.value)?"selected":""}>${F.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${s?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${s?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${s?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${s?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${g}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${x}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${s?.document_link?String(s.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${s?.notes||s?.description||""}</textarea>
        </div>
      </div>
    `;let u=p.querySelector("#sp-modal-tanggal"),$=p.querySelector("#sp-modal-akhir");u.addEventListener("change",()=>{if(u.value){let F=new Date(u.value);F.setMonth(F.getMonth()+6),$.value=F.toISOString().slice(0,10)}}),Z({title:y?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:p,size:"md",confirmText:y?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(F,I)=>{let L=(p.querySelector("#sp-modal-emp").value||"").trim(),M=p.querySelector("#sp-modal-div").value,_=p.querySelector("#sp-modal-branch").value,D=p.querySelector("#sp-modal-type").value,P=p.querySelector("#sp-modal-tanggal").value,N=p.querySelector("#sp-modal-akhir").value,k=(p.querySelector("#sp-modal-link").value||"").trim(),T=(p.querySelector("#sp-modal-notes").value||"").trim();if(!L){O("Nama Karyawan wajib diisi");return}if(!_){O("Cabang penempatan wajib dipilih");return}if(!P||!N){O("Tanggal SP dan Akhir SP wajib diisi");return}let A=i.find(J=>String(J.value)===String(_)),H=A?A.label:"Cabang",R={employee_name:L,division:M,branch_id:Number(_),branch_name:H,sp_type:D,tanggal:P,akhir_sp:N,document_link:k,notes:T,status:Ae(N).status};try{if(y){let J=await B(`/api/sp/${s.id}`,{method:"PUT",body:JSON.stringify(R)}),K=t.findIndex(j=>j.id===s.id);K!==-1&&(t[K]={...t[K],...R}),z("Data SP berhasil diperbarui")}else{let J=await B("/api/sp",{method:"POST",body:JSON.stringify(R)}),K=J.ok&&J.data?.id?J.data.id:Date.now();t.unshift({id:K,...R}),z("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),I(),c()}catch(J){console.error("Error saving SP:",J),O("Gagal menyimpan SP: "+J.message)}}})}function C(s){ae(`Hapus data Surat Peringatan untuk karyawan <strong>${s.employee_name}</strong> (${s.sp_type})?`,async()=>{try{await B(`/api/sp/${s.id}`,{method:"DELETE"})}catch(y){console.warn("API delete error fallback to local:",y)}t=t.filter(y=>y.id!==s.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),z(`Data SP ${s.employee_name} berhasil dihapus`),c()},"Hapus Data SP")}function b(s){let y=Ae(s.akhir_sp),g=qt(s.sp_type),x=document.createElement("div");x.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${s.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${s.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${s.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${s.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${g.bg}; color:${g.color}; border:1px solid ${g.border}; font-size:0.85rem;">
              <span>${g.icon}</span>
              <span>${g.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; font-size:0.85rem;">
              <span>${y.icon}</span>
              <span>${y.label}</span>
            </span>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${ze(s.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${ze(s.akhir_sp)}</div>
          </div>
        </div>

        ${s.notes||s.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${s.notes||s.description}</div>
          </div>
        `:""}

        ${s.document_link?`
          <div>
            <a href="${s.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,Z({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:x,size:"md",cancelText:"Tutup"})}async function f(){try{let{downloadExcel:s}=await Promise.resolve().then(()=>(Y(),he)),y=t.map((g,x)=>{let p=Ae(g.akhir_sp);return{No:x+1,"Nama Karyawan":g.employee_name||"",Divisi:g.division||"",Cabang:g.branch_name||"","Jenis SP":g.sp_type||"","Tanggal SP":g.tanggal||"","Akhir SP":g.akhir_sp||"","Status Masa Berlaku":p.status,"Sisa Hari":p.daysLeft>0?p.daysLeft:0,"Link BAP / Foto":g.document_link||"","Catatan / Kronologi":g.notes||g.description||""}});s(y,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),z("Berhasil mengunduh data SP ke format Excel.")}catch(s){O("Gagal ekspor: "+s.message)}}function c(){let s=E(),y=w();e.innerHTML=`
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
                <div class="sp-stat-value">${y.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${y.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${y.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${y.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              ${i.map(g=>`<option value="${g.value}" ${String(m)===String(g.value)?"selected":""}>${g.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${o==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${o==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${o==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <!-- View Switch -->
            <div class="sp-view-toggle">
              <button id="sp-view-table-btn" class="sp-view-btn ${a==="table"?"active":""}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
              <button id="sp-view-grid-btn" class="sp-view-btn ${a==="grid"?"active":""}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
                <span>Kartu</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${r==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${y.total}</span>
            </button>
            <button class="sp-pill ${r==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${y.active}</span>
            </button>
            <button class="sp-pill ${r==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${y.expiring}</span>
            </button>
            <button class="sp-pill ${r==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${y.expired}</span>
            </button>
            <button class="sp-pill ${r==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${y.sp1}</span>
            </button>
            <button class="sp-pill ${r==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${y.sp2}</span>
            </button>
            <button class="sp-pill ${r==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${y.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${s.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:a==="table"?`
            <!-- Modern Table View -->
            <div class="sp-table-card">
              <table class="sp-table">
                <thead>
                  <tr>
                    <th style="width:40px; text-align:center;">No</th>
                    <th>Nama Karyawan</th>
                    <th>Divisi</th>
                    <th>Cabang</th>
                    <th>Tingkat SP</th>
                    <th>Tanggal Terbit</th>
                    <th>Akhir SP</th>
                    <th>Status Masa Berlaku</th>
                    <th>Dokumen</th>
                    <th style="width:100px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${s.map((g,x)=>{let p=Ae(g.akhir_sp),u=qt(g.sp_type);return`
                      <tr>
                        <td style="text-align:center; font-weight:600; color:#94A3B8;">${x+1}</td>
                        <td>
                          <div style="display:flex; align-items:center; gap:10px;">
                            <div class="sp-avatar-box">
                              ${g.employee_name.slice(0,2).toUpperCase()}
                            </div>
                            <div>
                              <div style="font-weight:700; color:#0F172A; cursor:pointer;" class="sp-view-detail-link" data-id="${g.id}">
                                ${g.employee_name}
                              </div>
                              <div style="font-size:0.75rem; color:#64748B;">ID: #${g.id}</div>
                            </div>
                          </div>
                        </td>
                        <td>
                          <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700; font-size:0.74rem;">
                            ${g.division||"FACILITY CARE"}
                          </span>
                        </td>
                        <td>
                          <span style="font-weight:600; color:#334155;">${g.branch_name||"-"}</span>
                        </td>
                        <td>
                          <span class="sp-type-badge" style="background:${u.bg}; color:${u.color}; border:1px solid ${u.border};">
                            <span>${u.icon}</span>
                            <span>${u.label}</span>
                          </span>
                        </td>
                        <td style="font-weight:600; color:#475569;">
                          ${ze(g.tanggal)}
                        </td>
                        <td style="font-weight:700; color:#0F172A;">
                          ${ze(g.akhir_sp)}
                        </td>
                        <td>
                          <span class="sp-status-badge" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border};">
                            <span>${p.icon}</span>
                            <span>${p.shortLabel}</span>
                          </span>
                        </td>
                        <td>
                          ${g.document_link?`
                            <a href="${g.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4C4} Buka</span>
                            </a>
                          `:'<span style="color:#94A3B8; font-size:0.8rem;">-</span>'}
                        </td>
                        <td style="text-align:center;">
                          <div style="display:flex; justify-content:center; gap:6px;">
                            <button class="sp-icon-btn sp-detail-btn" title="Lihat Detail BAP" data-id="${g.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                            </button>
                            <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${g.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                            <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${g.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          </div>
                        </td>
                      </tr>
                    `}).join("")}
                </tbody>
              </table>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="sp-cards-grid">
              ${s.map(g=>{let x=Ae(g.akhir_sp),p=qt(g.sp_type);return`
                  <div class="sp-card" data-id="${g.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border};">
                          <span>${p.icon}</span>
                          <span>${p.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border};">
                          <span>${x.icon}</span>
                          <span>${x.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${g.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${g.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${g.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${g.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${ze(g.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${ze(g.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${g.notes||g.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${g.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${g.document_link?`
                          <a href="${g.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${g.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${g.id}">
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
    `,v()}function v(){let s=e.querySelector("#sp-btn-add");s&&(s.onclick=()=>h());let y=e.querySelector("#sp-btn-excel");y&&(y.onclick=()=>f());let g=e.querySelector("#sp-btn-refresh");g&&(g.onclick=()=>S());let x=e.querySelector("#sp-btn-reset-filter");x&&(x.onclick=()=>{l="",m="all",o="all",r="all",c()});let p=e.querySelector("#sp-search-input");p&&(p.oninput=L=>{l=L.target.value,c();let M=e.querySelector("#sp-search-input");M&&(M.focus(),M.setSelectionRange(M.value.length,M.value.length))});let u=e.querySelector("#sp-filter-branch");u&&(u.onchange=L=>{m=L.target.value,c()});let $=e.querySelector("#sp-filter-division");$&&($.onchange=L=>{o=L.target.value,c()});let F=e.querySelector("#sp-view-table-btn"),I=e.querySelector("#sp-view-grid-btn");F&&(F.onclick=()=>{a="table",localStorage.setItem("fcms_sp_view","table"),c()}),I&&(I.onclick=()=>{a="grid",localStorage.setItem("fcms_sp_view","grid"),c()}),e.querySelectorAll(".sp-pill").forEach(L=>{L.onclick=()=>{r=L.getAttribute("data-filter")||"all",c()}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(L=>{L.onclick=M=>{M.stopPropagation();let _=Number(L.getAttribute("data-id")),D=t.find(P=>P.id===_);D&&b(D)}}),e.querySelectorAll(".sp-edit-btn").forEach(L=>{L.onclick=M=>{M.stopPropagation();let _=Number(L.getAttribute("data-id")),D=t.find(P=>P.id===_);D&&h(D)}}),e.querySelectorAll(".sp-delete-btn").forEach(L=>{L.onclick=M=>{M.stopPropagation();let _=Number(L.getAttribute("data-id")),D=t.find(P=>P.id===_);D&&C(D)}})}await S()}U();var Ke=[],za=[];async function Ka(e){Ke=await Q(),za=await le(),G({container:e,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:t=>`<span class="badge ${t==="Selesai"?"badge-success":"badge-warning"}">${t||"-"}</span>`},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Ke},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Ke}],exportOptions:{moduleName:"mutasi_data",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await B(`/api/mutasi?limit=10000&${i}`);if(n.ok){let r=n.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:m}=await Promise.resolve().then(()=>(Y(),he));m(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(Y(),he));i(t,"Template_Import_Mutasi")},onImport:async t=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),a=Ke.find(d=>String(d.label||"").toLowerCase()===l);return a?a.value:null},n=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let d=Number(l);if(d>2e4&&d<99999){let S=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let a=l.split(/[\/\-\.]/);if(a.length===3){let[d,S,E]=a.map(w=>w.trim());if(d.length===4&&S.length<=2&&E.length<=2)return`${d}-${S.padStart(2,"0")}-${E.padStart(2,"0")}`;if(E.length===4&&S.length<=2&&d.length<=2)return`${E}-${S.padStart(2,"0")}-${d.padStart(2,"0")}`}return l},r=t.map(o=>({tanggal:n(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:i(String(o["Cabang Asal"]||"").trim()),to_branch_id:i(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),m=await B("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:za},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Ke,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Ke,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}U();var Ht=[],Oa=[];async function ja(e){Ht=await Q(),Oa=await _e(),G({container:e,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:t=>t!=null?`${t} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:t=>t!=null?`${t} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Ht}],exportOptions:{moduleName:"overtime_records",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await B(`/api/overtime?limit=10000&${i}`);if(n.ok){let r=n.data.data.map(o=>({Tanggal:o.date||"","Cabang Lembur":o.branch_name||"","Nama Karyawan":o.employee_name||"","Jam Mulai":o.start_time||"","Jam Selesai":o.end_time||"","Istirahat (Jam)":o.break_hours||0,"Total Jam":o.total_hours||0,"Alasan / Ket Lembur":o.reason||""})),{downloadExcel:m}=await Promise.resolve().then(()=>(Y(),he));m(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:Ht},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:Oa},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",e=>{if(e.target.name==="start_time"||e.target.name==="end_time"||e.target.name==="break_hours"){let t=e.target.closest("form");if(!t)return;let i=t.querySelector('[name="start_time"]'),n=t.querySelector('[name="end_time"]'),r=t.querySelector('[name="break_hours"]'),m=t.querySelector('[name="total_hours"]');if(!i||!n||!r||!m)return;let o=i.value,l=n.value,a=parseFloat(r.value)||0;if(o&&l){let[d,S]=o.split(":").map(Number),[E,w]=l.split(":").map(Number),h=E*60+w-(d*60+S);h<0&&(h+=24*60);let C=h/60-a;C<0&&(C=0),m.value=Number.isInteger(C)?C:C.toFixed(2)}}}));U();async function Ra(e){let t=oe();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}G({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,n)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${n.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,n)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let m=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(d=>String(d.id)===String(i));if(!m)return alert("Data tidak ditemukan");let o=d=>{if(!d)return"Tidak ada data";try{return JSON.stringify(JSON.parse(d),null,2)}catch{return d}},l=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(m.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(m.new_data)}</pre>
           </div>
         </div>
       `,{createModal:a}=await Promise.resolve().then(()=>(ue(),kt));a({title:`Detail Audit Log #${i}`,content:l,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}U();ue();Y();var Jt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},We={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function qa(e){e.innerHTML=`
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
  `;let t="board",i=null,n=[],r=oe(),m=r&&["superadmin","admin","manager","spv"].includes(r.role);B("/api/branches?all=1").then(c=>{c.ok&&c.data?.data&&(n=c.data.data)}).catch(()=>{});async function o(){try{let c=await B("/api/os-migrations/summary");if(!c.ok)throw new Error("Gagal memuat data migrasi");i=c.data?.data||c.data,l(i),d()}catch(c){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(c){let v=document.getElementById("migration-summary-cards");if(!v||!c)return;let s=c.total_branches||0,y=c.vendors?.AEON?.total||0,g=c.vendors?.["Triple S"]?.total||0,x=c.by_status?.completed||0,p=c.by_status?.in_progress||0,u=c.by_status?.scheduled||0,$=s>0?Math.round(x/s*100):0;v.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${s} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${y} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${g} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
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
          <span>\u{1F7E2} Selesai: <strong>${x}</strong></span>
          <span>\u{1F535} Proses: <strong>${p}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${u}</strong></span>
        </div>
      </div>
    `}function a(){if(!i||!i.matrix)return[];let c=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),v=document.getElementById("filter-vendor")?.value||"",s=document.getElementById("filter-month")?.value||"",y=document.getElementById("filter-status")?.value||"";return i.matrix.filter(g=>!(v&&g.vendor!==v||s&&g.target_month!==s||y&&g.status!==y||c&&!`${g.branch_code} ${g.branch_name} ${g.branch_full_name} ${g.vendor} ${g.notes||""} ${g.pic_vendor||""}`.toLowerCase().includes(c)))}function d(){let c=a(),v=document.getElementById("filter-stats-text");v&&(v.innerHTML=`Menampilkan <strong>${c.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let s=document.getElementById("migration-view-container");t==="board"?S(s,c):E(s,c)}function S(c,v){if(!v.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let s=["AEON","Triple S"],y=["2026-09","2026-10","2026-11","2026-12"],g="";s.forEach(x=>{let p=v.filter(M=>M.vendor===x);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==x||p.length===0)return;let u=p.length,$=x==="AEON"?"#1E40AF":"#0F766E",F=x==="AEON"?"#EFF6FF":"#F0FDFA",I=x==="AEON"?"#BFDBFE":"#99F6E4",L=y.filter(M=>p.some(_=>_.target_month===M));g+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${F}; border-bottom:1px solid ${I}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${$};">${x==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${$}; border:1px solid ${I}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${u} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${L.map(M=>{let _=p.filter(N=>N.target_month===M);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Jt[M]||M).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${_.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${_.map(N=>{let k=We[N.status]||We.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${N.branch_code}.</span>${N.branch_name}
                            </div>
                            ${m?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${N.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${N.id}" data-name="${N.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${N.id}" data-status="${N.status}" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${k.icon} ${k.label} \u25BE
                              </button>
                            </div>

                            ${N.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${N.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${N.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${N.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${N.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),c.innerHTML=g,w(c)}function E(c,v){if(!v.length){c.innerHTML=`
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
                ${m?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${v.map((s,y)=>{let g=We[s.status]||We.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${y+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${s.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${s.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${s.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${s.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Jt[s.target_month]||s.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${s.id}" data-status="${s.status}" style="background:${g.bg}; color:${g.color}; border:1px solid ${g.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${g.icon} ${g.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${s.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${s.employee_count>0?`\u{1F464} ${s.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${s.notes||"-"}
                    </td>
                    ${m?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${s.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${s.id}" data-name="${s.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,w(c)}function w(c){c.querySelectorAll(".status-toggle-btn").forEach(v=>{v.addEventListener("click",s=>{s.stopPropagation();let y=v.dataset.id,g=v.dataset.status;h(v,y,g)})}),c.querySelectorAll(".edit-card-btn").forEach(v=>{v.addEventListener("click",()=>{let s=v.dataset.id,y=i?.matrix?.find(g=>String(g.id)===String(s));y&&C(y)})}),c.querySelectorAll(".delete-card-btn").forEach(v=>{v.addEventListener("click",()=>{let s=v.dataset.id,y=v.dataset.name||"Cabang";ae(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${y}</strong>?`,async()=>{try{if(!(await B(`/api/os-migrations/${s}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");z(`Jadwal migrasi ${y} berhasil dihapus`),o()}catch(g){O(g.message)}})})})}function h(c,v,s){document.querySelectorAll(".status-popup-menu").forEach(u=>u.remove());let y=c.getBoundingClientRect(),g=document.createElement("div");g.className="status-popup-menu",g.style.cssText=`
      position: fixed;
      top: ${y.bottom+window.scrollY+4}px;
      left: ${y.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let x=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];g.innerHTML=x.map(u=>`
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
        color: ${u.key===s?"#2563EB":"var(--text-1)"};
        font-weight: ${u.key===s?"700":"500"};
      ">
        ${u.label} ${u.key===s?"\u2713":""}
      </button>
    `).join("");let p=u=>{!g.contains(u.target)&&u.target!==c&&(g.remove(),document.removeEventListener("click",p))};setTimeout(()=>document.addEventListener("click",p),50),g.querySelectorAll(".status-option-item").forEach(u=>{u.addEventListener("click",async()=>{let $=u.dataset.key;if(g.remove(),document.removeEventListener("click",p),$!==s)try{if(!(await B(`/api/os-migrations/${v}/status`,{method:"PATCH",body:JSON.stringify({status:$})})).ok)throw new Error("Gagal memperbarui status");z("Status berhasil diubah"),o()}catch(F){O(F.message)}})}),document.body.appendChild(g)}function C(c=null){let v=!!c,s=v?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",g=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${n.map(x=>`
      <option value="${x.id}" ${c&&c.branch_id===x.id?"selected":""}>
        ${x.full_name}
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
    `;Z({title:s,content:g,confirmText:v?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(x,p)=>{let u=x.querySelector("#migration-form");if(!u.checkValidity()){u.reportValidity();return}let $=new FormData(u),F={branch_id:parseInt($.get("branch_id"),10),vendor:$.get("vendor"),target_month:$.get("target_month"),status:$.get("status"),actual_date:$.get("actual_date")||null,pic_vendor:$.get("pic_vendor")||null,notes:$.get("notes")||null};try{let I=v?`/api/os-migrations/${c.id}`:"/api/os-migrations",M=await B(I,{method:v?"PUT":"POST",body:JSON.stringify(F)});if(!M.ok){let _=await M.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan data")}z(v?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),p(),o()}catch(I){O(I.message)}}})}let b=document.getElementById("view-board-btn"),f=document.getElementById("view-table-btn");b.addEventListener("click",()=>{t="board",b.style.background="#fff",b.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),d()}),f.addEventListener("click",()=>{t="table",f.style.background="#fff",f.classList.remove("btn-ghost"),b.style.background="transparent",b.classList.add("btn-ghost"),d()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(c=>{let v=document.getElementById(c);v&&(v.addEventListener("input",d),v.addEventListener("change",d))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{C()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let c=a();if(!c.length){O("Tidak ada data untuk diexport");return}let v=c.map((s,y)=>({No:y+1,"Kode Cabang":s.branch_code,"Nama Cabang":s.branch_name,"Cabang Lengkap":s.branch_full_name,"Vendor OS":s.vendor,"Target Bulan":Jt[s.target_month]||s.target_month,Status:We[s.status]?.label||s.status,"Tanggal Realisasi":s.actual_date||"","Jml Personil FC":s.employee_count||0,"PIC Vendor":s.pic_vendor||"",Catatan:s.notes||""}));q(v,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),z("File Excel berhasil diunduh")}),o()}U();ue();Y();var Ha={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Ja(e){e.innerHTML=`
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
  `;let t="pagi",i="timeline",n=[],r=oe(),m=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let c=await B("/api/dac?all=1");if(!c.ok)throw new Error("Gagal memuat standar DAC");n=c.data?.data||c.data||[];let v=n.filter(x=>x.shift==="pagi").length,s=n.filter(x=>x.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=v,document.getElementById("badge-count-siang").textContent=s,document.getElementById("badge-count-all").textContent=n.length;let y=document.getElementById("dac-category-filter"),g=Array.from(new Set(n.map(x=>x.category))).sort();y.innerHTML='<option value="">Semua Kategori</option>'+g.map(x=>`<option value="${x}">${x}</option>`).join(""),a(),d()}catch(c){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(){let c=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),v=document.getElementById("dac-category-filter")?.value||"",s=document.getElementById("dac-period-filter")?.value||"";return n.filter(y=>{if(t!=="all"&&y.shift!==t||v&&y.category!==v)return!1;if(s==="morning"){let g=parseInt(y.start_time.split(":")[0],10);if(y.shift==="pagi"&&g>=12||y.shift==="siang"&&g>=17)return!1}else if(s==="afternoon"){let g=parseInt(y.start_time.split(":")[0],10);if(y.shift==="pagi"&&g<12||y.shift==="siang"&&g<17)return!1}return!(c&&!`${y.start_time} ${y.end_time} ${y.category} ${y.location} ${y.activity}`.toLowerCase().includes(c))})}function a(){let c=document.getElementById("dac-summary-cards");if(!c)return;let v=n.filter(p=>p.shift==="pagi"),s=n.filter(p=>p.shift==="siang"),y=t==="pagi"?v:t==="siang"?s:n,g=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",x=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";c.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${x})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${y.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${g}</div>
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
    `}function d(){let c=l(),v=document.getElementById("dac-stats-text");v&&(v.innerHTML=`Menampilkan <strong>${c.length}</strong> aktivitas`);let s=document.getElementById("dac-view-container");i==="timeline"?S(s,c):E(s,c)}function S(c,v){if(!v.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}c.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${v.map(s=>{let y=Ha[s.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},g=s.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",x=s.shift==="pagi"?"#FEF3C7":"#E0E7FF",p=s.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${s.start_time} - ${s.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${s.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${x}; color:${p}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${g}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${s.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${s.location}
                    </span>
                  </div>
                  ${m?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${s.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${s.id}" data-name="${s.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${y.color}; margin-top:8px;">
                  ${s.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,w(c)}function E(c,v){if(!v.length){c.innerHTML=`
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
                ${t==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${m?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${v.map(s=>{let y=Ha[s.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${s.item_no}</td>
                    ${t==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${s.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${s.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${s.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${s.start_time} - ${s.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${s.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${s.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${s.activity}
                    </td>
                    ${m?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${s.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${s.id}" data-name="${s.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,w(c)}function w(c){c.querySelectorAll(".edit-dac-btn").forEach(v=>{v.addEventListener("click",()=>{let s=v.dataset.id,y=n.find(g=>String(g.id)===String(s));y&&h(y)})}),c.querySelectorAll(".delete-dac-btn").forEach(v=>{v.addEventListener("click",()=>{let s=v.dataset.id,y=v.dataset.name||"item";ae(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${y}</strong>?`,async()=>{try{if(!(await B(`/api/dac/${s}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");z("Aktivitas DAC berhasil dihapus"),o()}catch(g){O(g.message)}})})})}function h(c=null){let v=!!c,s=v?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",y=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],g=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${c?.shift==="pagi"||!c&&t==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${c?.shift==="siang"||!c&&t==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
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
            ${y.map(x=>`<option value="${x}">`).join("")}
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
    `;Z({title:s,content:g,confirmText:v?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(x,p)=>{let u=x.querySelector("#dac-form");if(!u.checkValidity()){u.reportValidity();return}let $=new FormData(u),F={shift:$.get("shift"),item_no:$.get("item_no")?parseInt($.get("item_no"),10):void 0,start_time:$.get("start_time"),end_time:$.get("end_time"),category:$.get("category"),location:$.get("location"),activity:$.get("activity")};try{let I=v?`/api/dac/${c.id}`:"/api/dac",M=await B(I,{method:v?"PUT":"POST",body:JSON.stringify(F)});if(!M.ok){let _=await M.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan aktivitas")}z(v?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),p(),o()}catch(I){O(I.message)}}})}function C(){let c=l();if(!c.length){O("Tidak ada data untuk dicetak");return}let v=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",s=window.open("","_blank");if(!s){O("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}s.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${v}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${v}</div>
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
            ${c.map(y=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${y.item_no}</td>
                ${t==="all"?`<td>${y.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${y.start_time} - ${y.end_time}</td>
                <td><strong>${y.category}</strong></td>
                <td><strong>${y.location}</strong></td>
                <td>${y.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),s.document.close(),s.focus(),setTimeout(()=>{s.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(c=>{c.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(v=>{v.style.background="#F1F5F9",v.style.color="var(--text-2)"}),c.style.background="#2563EB",c.style.color="#fff",t=c.dataset.shift,a(),d()})});let b=document.getElementById("view-timeline-btn"),f=document.getElementById("view-table-btn");b.addEventListener("click",()=>{i="timeline",b.style.background="#fff",b.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),d()}),f.addEventListener("click",()=>{i="table",f.style.background="#fff",f.classList.remove("btn-ghost"),b.style.background="transparent",b.classList.add("btn-ghost"),d()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(c=>{let v=document.getElementById(c);v&&(v.addEventListener("input",d),v.addEventListener("change",d))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{h()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{C()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let c=l();if(!c.length){O("Tidak ada data untuk diexport");return}let v=c.map(s=>({No:s.item_no,Shift:s.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":s.start_time,"Jam Selesai":s.end_time,Kategori:s.category,"Lokasi / Area":s.location,"Rincian Aktivitas":s.activity}));q(v,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),z("File Excel berhasil diunduh")}),o()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],n=parseInt(t[2],10),r=parseInt(t[3],10);if(n>12&&r<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let n=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),m=n[parseInt(i[1],10)-1];return`${r} ${m} ${i[0]}`}return t};function ee(e){return async t=>{if(!Re()){Fe("/login");return}return e(t)}}var Xe=null;function An(){Xe&&clearInterval(Xe);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),m=document.getElementById("header-clock-date");r&&(r.textContent=i),m&&(m.textContent=n)};e(),Xe=setInterval(e,1e3)}async function Dn(){try{let e=await B("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(n,r)=>{let m=document.getElementById(n);m&&(m.textContent=r>0?r:"",m.style.display=r>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var Oe=[];async function Bn(){try{let e=await B("/api/dashboard/notifications");if(!e.ok)return;Oe=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=Oe.length>0?"block":"none",t.textContent=Oe.length)}catch{}}function Pn(){if(!Oe.length){Z({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Oe.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;Z({title:`Notifikasi (${Oe.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function Ua(){let e=oe(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let d=new Date().getHours();return d>=4&&d<11?"Selamat Pagi":d>=11&&d<15?"Selamat Siang":d>=15&&d<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),m=document.getElementById("sidebar-close"),o=()=>{i.classList.add("open"),n.classList.add("show")},l=()=>{i.classList.remove("open"),n.classList.remove("show")};r?.addEventListener("click",o),m?.addEventListener("click",l),n?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",l)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(d=>{d.dataset.route!=="/overtime"&&(d.style.display="none")}),document.querySelectorAll(".nav-section").forEach(d=>{Array.from(d.querySelectorAll(".nav-item")).some(E=>E.style.display!=="none")||(d.style.display="none")}));function a(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(w=>{let h=w.dataset.route;w.classList.toggle("active",d===h||h!=="/dashboard"&&d.startsWith(h))});let S=document.getElementById("topbar-title"),E=document.querySelector(".nav-item.active .nav-label");S&&E&&(S.textContent=E.textContent)}window.addEventListener("hashchange",a),a(),An(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await B("/api/auth/logout",{method:"POST"}),qe(),Xe&&clearInterval(Xe),Fe("/login")}),Dn(),Bn(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),Pn()})}async function Ln(){X("/login",({main:t})=>aa(t)),X("/dashboard",ee(({main:t})=>Wt(t))),X("/calendar",ee(({main:t})=>Ta(t))),X("/employees",ee(({main:t,params:i})=>na(t,i))),X("/contracts",ee(({main:t,params:i})=>ct(t,i))),X("/sp",ee(({main:t})=>Na(t))),X("/mutasi",ee(({main:t})=>Ka(t))),X("/overtime",ee(({main:t})=>ja(t))),X("/timeline",ee(({main:t,params:i})=>ra(t,i))),X("/os-migrations",ee(({main:t})=>qa(t))),X("/issues",ee(({main:t,params:i})=>oa(t,i))),X("/one-on-one",ee(({main:t,params:i})=>sa(t,i))),X("/training",ee(({main:t})=>la(t))),X("/relievers",ee(({main:t,params:i})=>ca(t,i))),X("/reports/inspection",ee(({main:t})=>pa(t))),X("/reports/cleaning",ee(({main:t})=>ua(t))),X("/reports/fogging",ee(({main:t})=>ga(t))),X("/reports/basecamp",ee(({main:t})=>ma(t))),X("/reports/supply",ee(({main:t})=>$a(t,"supply"))),X("/sop",()=>{window.location.hash="#/sop-fc"}),X("/sop-fc",ee(({main:t})=>ya(t))),X("/sop-security",ee(({main:t})=>wa(t))),X("/checklist",ee(({main:t})=>Ea(t))),X("/hygiene-standards",ee(({main:t})=>Ca(t))),X("/dac",ee(({main:t})=>Ja(t))),X("/users",ee(({main:t})=>Fa(t))),X("/branches",ee(({main:t})=>_a(t))),X("/profile",ee(({main:t})=>Aa(t))),X("/settings/import",ee(({main:t})=>Ia(t))),X("/audit-logs",ee(({main:t})=>Ra(t)));let e=Re();if(!e&&window.location.hash!=="#/login"&&Fe("/login"),e){let t=await B("/api/auth/me");if(t.ok){He(t.data.data),Ua();let i=oe(),n=window.location.hash;i&&i.role==="input_lembur"&&(n==="#/dashboard"||n===""||n==="#/")&&Fe("/overtime")}else qe(),Fe("/login")}window.addEventListener("fm:login",()=>{Ua();let t=oe();t&&t.role==="input_lembur"?Fe("/overtime"):Fe("/dashboard")}),Vt()}Ln();
