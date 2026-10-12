var ka=Object.defineProperty;var gt=(t,e)=>()=>(t&&(e=t(t=0)),e);var bt=(t,e)=>{for(var i in e)ka(t,i,{get:e[i],enumerable:!0})};var Ie={};bt(Ie,{API:()=>It,CLIENT_SIDE_MAX_ROWS:()=>Te,IS_DEVELOPMENT:()=>Ke,apiFetch:()=>_,clearToken:()=>qe,getToken:()=>je,getUser:()=>re,setToken:()=>ht,setUser:()=>He});function je(){return localStorage.getItem("fm_token")}function ht(t){localStorage.setItem("fm_token",t)}function qe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function re(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function He(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function _(t,e={}){let i=je(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let r=`cb=${Date.now()}`,l=t.includes("?")?"&":"?",o=`${It}${t}${l}${r}`,s=await fetch(o,{...e,headers:a}),n;try{let d=await s.text();try{n=JSON.parse(d)}catch{n={error:`Server Error (${s.status}): ${d.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return s.status===401&&(qe(),window.location.hash="#/login"),{ok:s.ok,status:s.status,data:n}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var Ke,xa,It,Te,A=gt(()=>{Ke=!1,xa="https://fm-operations-api.facilitycare-audydental.workers.dev",It=xa,Te=1e4});var yt={};bt(yt,{confirmDialog:()=>De,createModal:()=>de});function de({title:t,content:e,onConfirm:i,onCancel:a,confirmText:r="Simpan",cancelText:l="Batal",size:o="md",confirmClass:s="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
    <div class="modal" style="max-width:${n[o]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${l}</button>
        ${i?`<button class="btn ${s} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(e);let f=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),f()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),f()}),i&&d.querySelector(".modal-confirm").addEventListener("click",()=>i(d,f)),d.addEventListener("click",g=>{g.target===d&&(a&&a(),f())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:f}}function De(t,e,i="Konfirmasi"){return de({title:i,content:`<p>${t}</p>`,onConfirm:(a,r)=>{e(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var xe=gt(()=>{});var le={};bt(le,{downloadExcel:()=>L,parseExcel:()=>Ue,renderExcelButtons:()=>wa});function Ue(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=r=>{try{let l=new Uint8Array(r.target.result),o=XLSX.read(l,{type:"array"}),s=o.SheetNames[0],n=o.Sheets[s];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${s}`);let d=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),f=d.e.r-d.s.r+1,g=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${f}`),console.log(`Total Columns: ${g}`);let u=[];for(let b=d.s.c;b<=d.e.c;++b){let h=n[XLSX.utils.encode_cell({c:b,r:d.s.r})];h&&h.v&&u.push(h.v)}console.log(`Headers Found: ${u.join(", ")}`),console.log("---------------------------");let c=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(c,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(c,"__headers",{value:u,enumerable:!1}),e(c)}catch(l){i(l)}},a.onerror=r=>i(r),a.readAsArrayBuffer(t)})}function L(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function wa(t){return`
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
  `}var R=gt(()=>{});A();var ft={},Xe=null;function J(t,e){ft[t]=e}function _e(t){window.location.hash=t}function Bt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),r=ft[i];if(!r){for(let[o,s]of Object.entries(ft))if(o.endsWith("/*")&&i.startsWith(o.slice(0,-2))){r=s;break}}Xe&&(Xe(),Xe=null);let l=document.getElementById("main-content");if(l&&(l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let o=new URLSearchParams(a.join("?")),s=i.split("/").filter(Boolean),n=await r({path:i,params:o,segments:s,main:l});n&&(Xe=n)}else{let o=l||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Je;function Sa(){return Je||(Je=document.createElement("div"),Je.id="toast-container",document.body.appendChild(Je)),Je}function Lt(t,e="info",i=3500){let a=Sa(),r=document.createElement("div");r.className=`toast toast-${e}`;let l={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${l[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var Q=t=>Lt(t,"success"),z=t=>Lt(t,"error");xe();A();A();xe();function Ze({columns:t,data:e,onEdit:i,onDelete:a,onView:r,actions:l=[],emptyText:o="Tidak ada data",bulkSelect:s=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,n;let d=document.createElement("table");d.className="data-table";let f=document.createElement("thead"),g=document.createElement("tr");if(s){let c=document.createElement("th");c.style.width="40px",c.style.textAlign="center";let b=document.createElement("input");b.type="checkbox",b.id="select-all-checkbox",b.title="Pilih semua",b.addEventListener("change",()=>{e.forEach(h=>{b.checked?s.selectedIds.add(h.id):s.selectedIds.delete(h.id)}),n.querySelectorAll(".row-checkbox").forEach(h=>h.checked=b.checked),s.onToggle()}),c.appendChild(b),g.appendChild(c)}if(t.forEach(c=>{let b=document.createElement("th");b.textContent=c.label,c.width&&(b.style.width=c.width),g.appendChild(b)}),i||a||r||l.length>0){let c=document.createElement("th");c.textContent="Aksi",c.style.width="120px",g.appendChild(c)}f.appendChild(g),d.appendChild(f);let u=document.createElement("tbody");return e.forEach(c=>{let b=document.createElement("tr");if(s){let h=document.createElement("td");h.style.textAlign="center",h.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=s.selectedIds.has(c.id),m.addEventListener("change",()=>{if(m.checked)s.selectedIds.add(c.id);else{s.selectedIds.delete(c.id);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1)}s.onToggle()}),h.appendChild(m),b.appendChild(h)}if(t.forEach(h=>{let m=document.createElement("td");if(h.render){let p=h.render(c[h.key],c);p instanceof HTMLElement?m.appendChild(p):m.innerHTML=p||""}else m.textContent=c[h.key]!==null&&c[h.key]!==void 0&&c[h.key]!==""?c[h.key]:"";h.nowrap&&(m.style.whiteSpace="nowrap"),b.appendChild(m)}),i||a||r||l.length>0){let h=document.createElement("td");h.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",r){let p=document.createElement("button");p.className="btn btn-xs btn-ghost",p.innerHTML="\u{1F441}",p.title="Lihat",p.addEventListener("click",()=>r(c)),m.appendChild(p)}if(i){let p=document.createElement("button");p.className="btn btn-xs btn-secondary",p.innerHTML="\u270F\uFE0F",p.title="Edit",p.addEventListener("click",()=>i(c)),m.appendChild(p)}l.forEach(p=>{let v=document.createElement("button");v.className=`btn btn-xs ${p.class||"btn-ghost"}`,v.innerHTML=p.icon||p.label,v.title=p.label,v.addEventListener("click",()=>p.handler(c)),m.appendChild(v)}),h.appendChild(m),b.appendChild(h)}u.appendChild(b)}),d.appendChild(u),n.appendChild(d),n}function et({page:t,pages:e,total:i,limit:a,onPage:r}){if(e<=1)return null;let l=document.createElement("div");l.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${i} data`,l.appendChild(o);let s=document.createElement("div");s.className="pagination-btns";let n=(g,u,c=!1,b=!1)=>{let h=document.createElement("button");h.className=`btn btn-sm ${b?"btn-primary":"btn-ghost"} pagination-btn`,h.textContent=g,h.disabled=c,h.addEventListener("click",()=>r(u)),s.appendChild(h)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let d=Math.max(1,t-2),f=Math.min(e,t+2);for(let g=d;g<=f;g++)n(g,g,!1,g===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),l.appendChild(s),l}xe();function ze(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${ze(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(e.type){case"textarea":r=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(c=>{let b=typeof c=="object"?c.value:c,h=typeof c=="object"?c.label:c,m=e.value==b?"selected":"";return`<option value="${b}" ${m}>${h}</option>`}).join("");r=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let s=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(c=>{let b=typeof c=="object"?c.value:c,h=typeof c=="object"?c.label||c.value||"":c||"";return(h==="undefined"||h==="[object Object]"||h==="null")&&(h=""),h?`<option value="${h}"></option>`:""}).join(""),d=e.value||"";if(e.value){let c=(e.options||[]).find(b=>(typeof b=="object"?b.value:b)==e.value);if(c){let b=typeof c=="object"?c.label||c.value||"":c||"";b&&b!=="undefined"&&b!=="[object Object]"&&b!=="null"&&(d=b)}}r=`
          <input type="text" name="${e.name}" list="${s}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${s}">${n}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let f=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";r=`<input type="date" name="${e.name}" class="form-control" value="${f}" ${i}>`;break;case"month":let g=e.value?String(e.value).slice(0,7):"";r=`<input type="month" name="${e.name}" class="form-control" value="${g}" ${i}>`;break;case"number":let u=e.readonly?"readonly":"";r=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i} ${u}>`;break;case"email":r=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let l=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${r}${l}</div>`}).join("")}function tt(t){let e={},i=new FormData(t);for(let[a,r]of i.entries())e[a]=r===""?null:r;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function at(t,e){e&&Object.entries(e).forEach(([i,a])=>{let r=t.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!a:r.type==="date"&&a&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(a):r.type==="month"&&a?r.value=String(a).slice(0,7):r.value=a??""))})}R();var Se={},nt={on(t,e){Se[t]||(Se[t]=new Set),Se[t].add(e)},off(t,e){Se[t]&&Se[t].delete(e)},emit(t,e){Se[t]&&Se[t].forEach(i=>{try{i(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(Se).forEach(t=>delete Se[t])}},_a=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function ce(t){if(!t){nt.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");nt.emit("data:changed",{module:e,relevant:_a.has(e)})}function P({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:l,filterFields:o,defaultFilters:s={},itemLabel:n="Data",canCreate:d=!0,canEdit:f=!0,canDelete:g=!0,onBeforeSubmit:u,onAfterLoad:c,onDataLoaded:b,extraActions:h=[],initialSearch:m="",exportOptions:p=null,bulkDelete:v=!1,paginationMode:w="server"}){let T=re();T&&T.role==="viewer"&&(d=!1,f=!1,g=!1,v=!1,p=null),T&&T.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(d=!1,f=!1,g=!1,v=!1,p=null),T&&T.role==="input_lembur"&&(a==="/api/overtime"?(g=!1,v=!1,p=null):(d=!1,f=!1,g=!1,v=!1,p=null));let k=1,E={...s};m&&(E.search=m);let C=new Set;t.innerHTML=`
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
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${p?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${p.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${p.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${p.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${p.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
        ${o.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${E.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let S=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${S}</option>${(y.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${E[y.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let S=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${S}</option>${(y.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${E[y.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let y=document.getElementById("bulk-toolbar");if(!y)return;let S=document.getElementById("bulk-count"),x=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");S.textContent=`${C.size} item dipilih`,C.size>0?(y.style.display="flex",x.disabled=!1,O.disabled=!1):(y.style.display="none",x.disabled=!0,O.disabled=!0);let N=document.getElementById("select-all-checkbox");if(N){let U=document.querySelectorAll(".row-checkbox");if(U.length>0){let B=[...U].every(oe=>oe.checked),F=[...U].some(oe=>oe.checked);N.checked=B,N.indeterminate=F&&!B}else N.checked=!1,N.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{C.clear(),document.querySelectorAll(".row-checkbox").forEach(S=>S.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(C.size===0)return;let y=[...C],S=document.createElement("div");S.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",S.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(S),S.querySelector("#bulk-cancel-btn").addEventListener("click",()=>S.remove()),S.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let x=S.querySelector("#bulk-confirm-btn");x.disabled=!0,x.textContent="Menghapus...";let O=await _(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});S.remove(),O.ok?(Q(`${y.length} ${n} berhasil dihapus.`),C.clear(),$(),ce(a),M()):z(O.data?.error||"Gagal menghapus data.")})});let I=document.getElementById("filter-search"),D;I?.addEventListener("input",y=>{clearTimeout(D),D=setTimeout(()=>{E.search=y.target.value,k=1,C.clear(),$(),M()},400)}),o?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",S=>{E[y.name]=S.target.value;let x=document.getElementById(`filter-sheet-${y.name}`);x&&(x.value=S.target.value),k=1,C.clear(),$(),M()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",S=>{E[y.name]=S.target.value;let x=document.getElementById(`filter-${y.name}`);x&&(x.value=S.target.value),k=1,C.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{E={},I&&(I.value=""),o?.forEach(y=>{let S=document.getElementById(`filter-${y.name}`);S&&(S.value="");let x=document.getElementById(`filter-sheet-${y.name}`);x&&(x.value="")}),k=1,C.clear(),$(),M()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{E={},I&&(I.value=""),o?.forEach(y=>{let S=document.getElementById(`filter-${y.name}`);S&&(S.value="");let x=document.getElementById(`filter-sheet-${y.name}`);x&&(x.value="")}),k=1,C.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>Y(null)),p&&document.addEventListener("click",function(y){let S=document.getElementById("aksi-menu-main"),x=document.getElementById("btn-aksi-main");S&&x&&!x.contains(y.target)&&!S.contains(y.target)&&S.classList.remove("show-aksi-menu")});let q=document.getElementById("btn-mobile-filter"),W=document.getElementById("filter-options-wrapper"),ae=document.getElementById("btn-close-filter-sheet");if(q&&W&&(q.addEventListener("click",y=>{y.preventDefault(),W.classList.add("sheet-open")}),ae&&ae.addEventListener("click",y=>{y.preventDefault(),W.classList.remove("sheet-open")})),p){document.getElementById(`btn-export-${p.moduleName}`)?.addEventListener("click",async S=>{let x=S.target,O=x.innerHTML;x.innerHTML="\u23F3 Loading...",x.disabled=!0;try{await p.onExport()}catch{z("Gagal export data")}finally{x.innerHTML=O,x.disabled=!1}}),document.getElementById(`btn-template-${p.moduleName}`)?.addEventListener("click",()=>{p.onTemplate()});let y=document.getElementById(`input-import-${p.moduleName}`);y?.addEventListener("change",async S=>{let x=S.target.files[0];if(!x)return;let O=document.getElementById(`label-import-${p.moduleName}`),N=O?O.querySelector(".import-text"):null,U=N?N.innerText:"";N&&(N.innerText="\u231B Memproses..."),O&&(O.style.pointerEvents="none"),y.disabled=!0;try{let B=await Ue(x);if(B.length===0)throw new Error("File kosong atau format salah");await p.onImport(B),Q("Import berhasil!"),ce(a),M()}catch(B){z(B.message||"Gagal import data")}finally{N&&(N.innerText=U),O&&(O.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function M(){$();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let S=w==="client",x=S?1:k,O=S?Te:20,N=new URLSearchParams({page:x,limit:O,...Object.fromEntries(Object.entries(E).filter(([,H])=>H))}),U=await _(`${a}?${N}`);if(!U.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let B=U.data?.data||U.data||[],F=U.data?.pagination,oe=B.length;if(S){B=b(B);let H=B.length,X=20,ne=Math.ceil(H/X);k>ne&&ne>0&&(k=ne);let K=(k-1)*X,se=k*X;B=B.slice(K,se),F={page:k,limit:X,total:H,pages:ne}}!1,c&&c(B);let ve=Ze({columns:r,data:B,onEdit:f?H=>Y(H):null,actions:h.map(H=>({...H,handler:X=>H.handler(X,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:v?{selectedIds:C,onToggle:$}:null});y.innerHTML="",y.appendChild(ve);let pe=document.getElementById("pagination-container");if(pe&&(pe.innerHTML="",F&&F.pages>1)){let H=et({page:F.page,pages:F.pages,total:F.total,limit:F.limit,onPage:X=>{k=X,M()}});H&&pe.appendChild(H)}}function $e(y){let S=typeof l=="function"?l(y):l;return ze(S)}function Y(y){let S=!!y,x=document.createElement("form");if(x.noValidate=!0,x.innerHTML=$e(y),S){let N=typeof l=="function"?l(y):l;at(x,y)}let{close:O}=de({title:S?`Edit ${n}`:`Tambah ${n}`,content:x,size:"lg",confirmText:S?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(N,U)=>{if(!x.reportValidity())return;let B=N.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let F=tt(x),oe=typeof l=="function"?l(y):l,ve=async ne=>{for(let K of ne)if(K.type==="row")await ve(K.fields);else if(K.type==="combobox"&&F[K.name]){let se=F[K.name],ke=(K.options||[]).find(Z=>{let ie=String(typeof Z=="object"?Z.value:Z),ut=String(typeof Z=="object"?Z.label:Z);return ie===se||ut===se});if(ke)F[K.name]=typeof ke=="object"?ke.value:ke;else if(K.createApi){let Z={};Z[K.createApi.field]=se,K.createApi.extra&&Object.assign(Z,K.createApi.extra);let ie=await _(K.createApi.path,{method:"POST",body:JSON.stringify(Z)});if(ie.ok&&ie.data?.id)F[K.name]=ie.data.id;else if(ie.ok&&!ie.data?.id)F[K.name]=se;else throw new Error(`Gagal membuat master data: ${ie.data?.error||"Unknown error"}`)}}};try{await ve(oe)}catch(ne){z(ne.message),B.disabled=!1,B.textContent=S?"Simpan Perubahan":`Tambah ${n}`;return}u&&(F=await u(F,y));let pe=S?"PUT":"POST",H=S?`${a}/${y.id}`:a,X=await _(H,{method:pe,body:JSON.stringify(F)});X.ok?(Q(S?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),ce(a),M()):(z(X.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=S?"Simpan Perubahan":`Tambah ${n}`)}})}function Ee(y){De(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let S=await _(`${a}/${y.id}`,{method:"DELETE"});S.ok?(Q(`${n} berhasil dihapus.`),ce(a),M()):z(S.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}A();A();var Be=null,it=null;async function Ce(t=!1){if(Be&&!t)return console.log("Employees Raw (Cache Hit)",Be.slice(0,5)),Be;let e=await _(`/api/employees?limit=${Te}&status=Aktif`);return Be=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Be.slice(0,5)),Be}async function te(t=!1){let i=(await Ce(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function j(t=!1){return it&&!t||(it=((await _("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),it}function G(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function vt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Le(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function kt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function me(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}R();A();R();A();R();A();R();function rt(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),l=new Date(i.value),o=Math.floor((l-r)/864e5);a.value=isNaN(o)?"":o}else a.value=""}}document.body.removeEventListener("input",rt);document.body.addEventListener("input",rt);document.body.removeEventListener("change",rt);document.body.addEventListener("change",rt);A();var we={};function Ae(t){if(we[t]){try{we[t].destroy()}catch{}delete we[t]}}function Ca(){Object.keys(we).forEach(Ae)}var fe=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Pe=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function ot(t,e,i=900){if(!t)return;t._animFrame&&cancelAnimationFrame(t._animFrame);let a=Math.max(0,Math.round(fe(e))),r=t.textContent?t.textContent.replace(/[^\d]/g,""):"",l=r!==""&&parseInt(r)||0;if(a===l){t.textContent=a.toLocaleString("id-ID"),t._animFrame=null;return}let o=Date.now(),s=()=>{let n=Math.min((Date.now()-o)/i,1),d=1-Math.pow(1-n,3);t.textContent=Math.round(l+d*(a-l)).toLocaleString("id-ID"),n<1?t._animFrame=requestAnimationFrame(s):(t.textContent=a.toLocaleString("id-ID"),t._animFrame=null)};t._animFrame=requestAnimationFrame(s)}var $a={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Ea=t=>{let e=Pe(t,"\u2014");return`<span class="status-pill ${$a[e]||"pill-neutral"}">${e}</span>`};var ge={family:"Inter",size:11},ye="#94A3B8",Fe="#F1F5F9",xt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Ta=()=>window.innerWidth<768;function lt(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Ta()?"bottom":"top",labels:{font:ge,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ge,titleFont:{...ge,weight:"700"}}},scales:{x:{grid:{color:Fe},ticks:{font:ge,color:ye,maxRotation:0}},y:{grid:{color:Fe},ticks:{font:ge,color:ye},beginAtZero:!0}},...t}}var Da=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),Ia=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Pt(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ue(t,e,i=8e3){try{let a=new AbortController,r=setTimeout(()=>a.abort(),i),l=await _(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(r),!l||!l.ok)return e;let o=l.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function Ba(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let r=document.getElementById(a);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let r=document.getElementById(a);if(r&&r.style.display==="none"){r.style.display="block";let l=r.parentElement;if(l&&!l.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",r.style.display="none",l.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Nt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Mt({}),["table-contracts","table-issues"].forEach(a=>{let r=document.getElementById(a);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function At(t){Ca(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Da()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Ia()}</div>

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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ft(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=a);let l=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");l&&(l.style.display="block",l.style.position="absolute"),o&&(o.style.display="none");let s=await ue(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Ot(s)}catch(n){console.warn("ScheduleChart render:",n),be("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,r=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",l=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");l&&(l.style.display="block",l.style.position="absolute"),o&&(o.style.display="none");let s=await ue(r,{},8e3);try{Rt(s)}catch(n){console.warn("InspBar render:",n),be("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>Ba(),5e3),await Ft(t)}async function Ft(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,r,l,o,s,n]=await Promise.all([ue("/api/dashboard/kpi",{},8e3),ue("/api/dashboard/issues-trend",{},8e3),ue("/api/dashboard/issues-summary",{},8e3),ue("/api/dashboard/stats",{},8e3),ue("/api/dashboard/calendar",[],8e3),ue("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ue(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ue("/api/dashboard/dropdown-aggregates",{},8e3)]),d=document.getElementById("filter-insp-month"),f=d?d.value:"",g=f?`/api/dashboard/inspection-bar?month=${f}`:"/api/dashboard/inspection-bar",u=await ue(g,{},8e3);window.dashboardAggregates=n||{};try{Nt(e)}catch(c){console.warn("KPI render:",c)}try{Mt(e)}catch(c){console.warn("MiniStats render:",c)}try{Ot(s)}catch(c){console.warn("ScheduleChart render:",c),be("skel-jadwal","chart-jadwal")}try{La(Array.isArray(a?.by_category)?a.by_category:[])}catch(c){console.warn("Donut render:",c),be("skel-donut","chart-donut")}try{Pa(i)}catch(c){console.warn("Trend render:",c),be("skel-trend","chart-trend")}try{Rt(u)}catch(c){console.warn("InspBar render:",c),be("skel-insp","chart-insp")}try{let c=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Aa(c)}catch(c){console.warn("IssuesTable render:",c)}try{let c=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];Fa(o)}catch(c){console.warn("ContractsTable render:",c)}try{Na(Array.isArray(l)?l:[])}catch(c){console.warn("Agenda render:",c)}try{Ma()}catch(c){console.warn("Quick Actions render:",c)}}function Nt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let r=fe(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${r}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${r}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{ot(a,parseInt(a.dataset.target)||0)})}function Mt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),r=String(new Date().getMonth()+1).padStart(2,"0"),l=`${a}-${r}`,o=c=>`
    <select id="${c}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
      <option value="${a}-01" ${l===`${a}-01`?"selected":""}>Jan</option>
      <option value="${a}-02" ${l===`${a}-02`?"selected":""}>Feb</option>
      <option value="${a}-03" ${l===`${a}-03`?"selected":""}>Mar</option>
      <option value="${a}-04" ${l===`${a}-04`?"selected":""}>Apr</option>
      <option value="${a}-05" ${l===`${a}-05`?"selected":""}>Mei</option>
      <option value="${a}-06" ${l===`${a}-06`?"selected":""}>Jun</option>
      <option value="${a}-07" ${l===`${a}-07`?"selected":""}>Jul</option>
      <option value="${a}-08" ${l===`${a}-08`?"selected":""}>Agu</option>
      <option value="${a}-09" ${l===`${a}-09`?"selected":""}>Sep</option>
      <option value="${a}-10" ${l===`${a}-10`?"selected":""}>Okt</option>
      <option value="${a}-11" ${l===`${a}-11`?"selected":""}>Nov</option>
      <option value="${a}-12" ${l===`${a}-12`?"selected":""}>Des</option>
    </select>
  `,s=window.dashboardAggregates||{},n=c=>{let b=window.dashboardAggregates?.schedule_quarter_details?.[c];return b?`Insp: ${b.insp} \u2022 GC: ${b.gc} \u2022 DC: ${b.dc} \u2022 Fog: ${b.fog}`:""},d=c=>{let b=window.dashboardAggregates?.schedule_quarter_details?.[c];return b?`Rincian ${c}: Inspeksi: ${b.insp} | General Cleaning: ${b.gc} | Deep Cleaning: ${b.dc} | Fogging: ${b.fog}`:`Total Kegiatan ${c}`},f=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:s.schedule_by_quarter?.[i]??t.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"mini-blue",extra:`<div id="mini-jadwal-detail" style="font-size:0.6rem; color:var(--text-3); font-weight:500; margin-top:2px; line-height:1.2; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${d(i)}">${n(i)}</div>`},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:o("dash-inspeksi-month"),val:s.inspeksi_by_month?.[l]??t.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${l}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:o("dash-gcdc-month"),val:s.gcdc_by_month?.[l]??t.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${l}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:o("dash-reliefer-month"),val:s.relievers_by_month?.[l]??t.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${l}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:o("dash-fogging-month"),val:s.fogging_by_month?.[l]??t.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${l}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=f.map(c=>`
    <a href="${c.href}" class="mini-stat ${c.color}" style="text-decoration:none" id="${c.id||""}">
      <div class="mini-stat-icon">${c.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${fe(c.val)}">0</div>
          ${c.dropdown?c.dropdown:""}
        </div>
        <div class="mini-stat-text">${c.label}</div>
        ${c.extra||""}
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(c=>ot(c,parseInt(c.dataset.target)||0,700));let g=document.getElementById("dash-jadwal-period");if(g){let c=b=>{let m=(window.dashboardAggregates?.schedule_by_quarter||{})[b]||0,p=document.querySelector("#mini-jadwal .mini-stat-value");p&&(p.dataset.target=m,ot(p,m,400));let v=document.getElementById("mini-jadwal-detail");v&&(v.textContent=n(b),v.title=d(b));let w=document.getElementById("mini-jadwal");w&&(w.href=`#/timeline?dash_filter=period_${b.toLowerCase()}`,w.title=d(b))};g.addEventListener("change",b=>c(b.target.value))}let u=(c,b,h,m)=>{let p=document.getElementById(c);if(p){let v=w=>{let k=(window.dashboardAggregates?.[h]||{})[w]||0,E=document.querySelector(`#${b} .mini-stat-value`);E&&(E.dataset.target=k,ot(E,k,400));let C=document.getElementById(b);C&&(C.href=`${m}&month=${w}`)};p.addEventListener("change",w=>v(w.target.value))}};u("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),u("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),u("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),u("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function La(t){be("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Ae("donut");let a=(t||[]).filter(n=>fe(n.count)>0);if(!a.length){Ge(e,"Belum ada data permasalahan");return}let r=a.map(n=>`${Pe(n.category,"Lainnya")}`),l=a.map(n=>fe(n.count)),o=l.reduce((n,d)=>n+d,0);i.innerHTML=a.map((n,d)=>{let f=xt[d%xt.length],g=o>0?Math.round(n.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${f}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${g}%)</span></div>
          <div class="donut-legend-label">${r[d]}</div>
        </div>
      </div>
    `}).join("");let s={id:"centerText",beforeDraw:function(n){let d=n.width,f=n.height,g=n.ctx;g.restore();let u=(f/80).toFixed(2);g.font="bold "+u+"em Inter",g.textBaseline="middle",g.fillStyle="#1E293B";let c=o.toString(),b=Math.round((d-g.measureText(c).width)/2),h=f/2;g.fillText(c,b,h-4),g.font="600 "+(u*.35).toFixed(2)+"em Inter",g.fillStyle="#64748B";let m="Total",p=Math.round((d-g.measureText(m).width)/2);g.fillText(m,p,h+10),g.save()}};we.donut=new Chart(e,{type:"doughnut",data:{labels:r,datasets:[{data:l,backgroundColor:xt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ge,titleFont:{...ge,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[s]})}function Pa(t){be("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Ae("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[s,n]=o.split("-");return(i[Number(n)-1]||n)+" "+String(s).slice(-2)}catch{return o}}),r=(t.open||[]).map(o=>fe(o)),l=(t.closed||[]).map(o=>fe(o));if(!a.length){Ge(e,"Belum ada data trend");return}we.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:l,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:lt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ye,maxRotation:45,autoSkip:!0}},y:{grid:{color:Fe},ticks:{font:{family:"Inter",size:9},color:ye},beginAtZero:!0}}})})}function Ot(t){be("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Ae("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(d=>d>0))){Ge(e,"Belum ada data jadwal");return}let r=t["Inspeksi Hygiene"]||Array(12).fill(0),l=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),s=t.Fogging||Array(12).fill(0);we.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:l,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:s,backgroundColor:"#EF4444"}]},options:lt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ye,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Fe},ticks:{font:{family:"Inter",size:9},color:ye},min:0}}})})}function Rt(t){be("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Ae("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(l=>fe(l)),r=(t.spv||[]).map(l=>fe(l));if(!i.length){Ge(e,"Belum ada data inspeksi");return}we.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:lt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ge,color:ye,maxRotation:45,minRotation:45}},y:{grid:{color:Fe},ticks:{font:ge,color:ye},min:0,max:100}}})})}function Fa(t){be("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Ae("contractMiniBar"),t=t||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let s=o.split("-")[1];return i[s]||o}),r=(t.data||[]).map(o=>fe(o));if(!a.length){Ge(e,"Belum ada data");return}let l=e.getContext("2d");we.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:lt({onClick:(o,s)=>{if(s&&s.length>0){let n=s[0].index,d=(t.labels||[])[n];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ge,color:ye,maxRotation:0,autoSkip:!1}},y:{grid:{color:Fe,borderDash:[4,4],drawBorder:!1},ticks:{font:ge,color:ye,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Aa(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${Ea(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Pe(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Pe(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Na(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,l=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!l.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${l.map(o=>{let s="#3B82F6",n="#EFF6FF",d="Agenda",f=(o.title||"").toLowerCase();return f.includes("inspeksi")?(s="#10B981",n="#ECFDF5",d="Inspeksi"):f.includes("cleaning")||f.includes("gcdc")?(s="#3B82F6",n="#EFF6FF",d="Cleaning"):f.includes("reliefer")?(s="#F59E0B",n="#FFFBEB",d="Reliefer"):f.includes("fogging")&&(s="#8B5CF6",n="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${s};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Pe(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Pe(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${s}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ma(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function be(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let r=a.parentElement;if(r){let l=r.querySelector(".chart-empty");l&&l.remove()}}}function Ge(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=e,i.appendChild(r)}}A();async function Kt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),l=document.getElementById("login-password");r?.addEventListener("click",()=>{let o=l.type==="text";l.type=o?"password":"text",r.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),i.style.display="none";let s=e.username.value.trim(),n=e.password.value;if(!s||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let d=await _("/api/auth/login",{method:"POST",body:JSON.stringify({username:s,password:n})});d.ok&&d.data.success?(ht(d.data.data.token),He(d.data.data.user),Q("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=d.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}A();R();async function Oa(){return await j()}function Ne(t){if(!t||t==="-"||String(t).trim()==="")return"-";let e=String(t).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],a=e.match(/^(\d{4})-(\d{1,2})/);if(a){let l=a[1],o=parseInt(a[2],10)-1;if(o>=0&&o<12)return`${i[o]} ${l}`}let r=e.split(/[\/\-\.]/);if(r.length===2&&r[1].length===4){let l=parseInt(r[0],10)-1;if(l>=0&&l<12)return`${i[l]} ${r[1]}`}if(r.length===3&&r[2].length===4){let l=parseInt(r[1],10)-1;if(l>=0&&l<12)return`${i[l]} ${r[2]}`}for(let l=0;l<12;l++)if(e.toLowerCase().includes(i[l].toLowerCase()))return e;return e}function Ra(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function jt(t,e){let i=await Oa(),a=e?e.get("dash_filter"):null;P({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>a?r.filter(l=>Ra(l,a)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Le(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"target_pindah_os",label:"Target Pindah OS",render:r=>Ne(r)},{key:"target_selesai_os",label:"Target Selesai OS",render:r=>Ne(r)},{key:"status",label:"Status",render:r=>G(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await _(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let l=r.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||"","Target Pindah OS":Ne(o.target_pindah_os)==="-"?"":Ne(o.target_pindah_os),"Target Selesai OS":Ne(o.target_selesai_os)==="-"?"":Ne(o.target_selesai_os)}));L(l,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async r=>{let l=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),f=i.find(g=>String(g.label||"").toLowerCase()===d);return f?f.value:null},o=r.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:l(String(n.Cabang||"").trim()),division:String(n.Divisi||n["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||n["No. Hp"]||"").trim(),join_date:String(n["Tgl Masuk"]||n["Tanggal Masuk"]||"").trim(),status:String(n.Status||"").trim(),target_pindah_os:String(n["Target Pindah OS"]||"").trim(),target_selesai_os:String(n["Target Selesai OS"]||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),s=await _("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}A();R();var wt=[],qt=[];async function Ka(){wt=await j(),qt=await Ce()}var St=async t=>{let e=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(A(),Ie))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let l=r.data?.data||r.data||[],o=Array.isArray(l)?l:[];if(e=e.concat(o),o.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return e};async function st(t,e){await Ka(),P({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Le(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':vt(a)},{key:"status",label:"Status",render:a=>G(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:wt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[l,o]=await Promise.all([St("/api/employees?status=Aktif"),St("/api/contracts")]);if(l.length>0){let s=o.filter(g=>g.status==="Aktif"),n=new Set(s.map(g=>g.employee_id)),d=l.filter(g=>!n.has(g.id)),f=`<p style="margin-bottom:12px">Data yang terbaca: <b>${l.length}</b> Karyawan Aktif, dan <b>${s.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${d.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;d.forEach(g=>{let u=o.filter(b=>b.employee_id===g.id),c='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(u.length>0){let b=u[0];c=`Pernah ada kontrak (Status: <b style="color:#EF4444">${b.status}</b>, Selesai: ${window.formatDate(b.end_date)})`}f+=`<li style="margin-bottom:8px"><b>${g.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${g.branch_name||"-"} | ${c}</span></li>`}),f+="</ul>",Promise.resolve().then(()=>(xe(),yt)).then(g=>g.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:f,cancelText:"Tutup"}))}}catch(l){console.error(l)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:qt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:wt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await _(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let r=a.data.data.map(l=>({"Nama Lengkap":l.employee_name,Cabang:l.branch_name||"","Div / Bagian":l.division||"","Tanggal Mulai":l.start_date||"","Tanggal Selesai":l.end_date&&String(l.end_date).startsWith("2099")?"":l.end_date||"","Sisa Kontrak":l.end_date&&String(l.end_date).startsWith("2099")?"Tetap":l.days_remaining!==null&&l.days_remaining!==void 0?`${l.days_remaining} Hari`:"",Status:l.status||""}));L(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[r,l]=await Promise.all([_("/api/branches?limit=10000"),St("/api/employees")]),o=r.data?.data||[],s=l||[];console.log(`Total employee yang berhasil dimuat dari database : ${s.length}`),s.length>0&&(console.log("Contoh 5 employee pertama:"),s.slice(0,5).forEach((m,p)=>{console.log(`${p+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let n=m=>{if(!m)return null;let p=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),v=o.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===p||String(w.code||"").replace(/\s+/g," ").toLowerCase().trim()===p||String(w.name||"").replace(/\s+/g," ").toLowerCase().trim()===p);return v?v.id:null},d=(m,p)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${p}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let v=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${v}"`),console.log(`Jumlah employee di database : ${s.length}`);let w=s.find(T=>String(T.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===v);return w?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${w.id}`),w.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},f=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let p=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(p)){let w=Math.floor(Number(p));if(w>2e4&&w<99999){let T=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);let v=p.split(/[\/\-\.]/);if(v.length===3){let[w,T,k]=v.map(E=>E.trim());if(w.length===4&&T.length<=2&&k.length<=2)return`${w}-${T.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&T.length<=2&&w.length<=2)return`${k}-${T.padStart(2,"0")}-${w.padStart(2,"0")}`}return p},g=a.map((m,p)=>{let v=p+2,w=String(m["Nama Lengkap"]||"").trim(),T=m["Tanggal Mulai"],k=f(T);if(!k){let $=a.__worksheet,I=a.__headers||[],D=I.indexOf("Tanggal Mulai"),q="N/A",W="N/A",ae="N/A";if(D!==-1&&$&&window.XLSX){let $e=window.XLSX.utils.encode_cell({c:D,r:v-1});ae=$e;let Y=$[$e];Y?(q=Y.t||"undefined",W=Y.w||"undefined"):q="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let M="Unknown";T==null||T===""?M="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":T instanceof Date&&isNaN(T.getTime())?M="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":M="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${v}`),console.log(`Employee Name : ${w}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${D})`),console.log(`Raw Cell Value : "${T}"`),console.log(`JavaScript Type : ${typeof T}`),console.log(`SheetJS Cell Type : ${q}`),console.log(`SheetJS Formatted Value : "${W}"`),console.log(`Value After Trim : "${String(T||"").trim()}"`),console.log(`Value After Date Parser : "${k}"`),console.log(`Is Empty : ${!T}`),console.log(`Is Invalid Date : ${T instanceof Date?isNaN(T.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${M}`),console.log(`Workbook Sheet : ${$?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${ae}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(I)),console.log(`==========================
`)}let E=d(w,v),C=null;return E?k||(C="Tanggal Mulai kosong atau tidak berformat tanggal"):C="Karyawan tidak ditemukan di Database",{isValid:!!(E&&k),invalidReason:C,rowNum:v,data:{employee_id:E,branch_id:n(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:k,end_date:f(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:w}}}),u=[],c=[];if(g.forEach(m=>{m.isValid?u.push(m.data):c.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${u.length}, Invalid: ${c.length}`),u.length===0){let m=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${c.length}

Daftar Kegagalan (Contoh):
`;c.slice(0,10).forEach(p=>{m+=`- Row ${p.rowNum} | Nama: ${p.name} | Alasan: ${p.reason}
`}),c.length>10&&(m+=`- ... dan ${c.length-10} lainnya.
`),alert(m);return}let b=await _("/api/contracts/import",{method:"POST",body:JSON.stringify(u)}),h=`IMPORT SUMMARY
======================
`;h+=`Total Baris Excel : ${a.length}
`,h+=`Baris Valid       : ${u.length}
`,h+=`Baris Invalid     : ${c.length}

`,b&&b.data&&b.data.metrics?(h+=`Berhasil INSERT   : ${b.data.metrics.inserted}
`,h+=`Berhasil UPDATE   : ${b.data.metrics.updated}
`):h+=`Berhasil diproses : ${u.length}
`,c.length>0&&(h+=`
DAFTAR DATA DILEWATI:
`,c.forEach(m=>{h+=`- Row ${m.rowNum} | ${m.name} | ${m.reason}
`})),alert(h),typeof st=="function"&&st()}}})}A();R();var _t=[],Ve=[];function ja(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Ht(t,e){_t=await j();let i=await te();Ve=["Berlin Ariansyah","Ade Surahman"];let a=m=>m&&!Ve.find(p=>String(typeof p=="object"?p.value:p).toLowerCase()===String(m).toLowerCase())?[...Ve,m]:Ve,r=await _(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),l=m=>{if(!m||m==="-"||String(m).trim()==="")return"";let p=String(m).split("-");return p.length===3&&p[0].length===4?`${p[2]}-${p[1]}-${p[0]}`:m},o=r.data?.data||[],s=ja(o),n=e?e.get("dash_filter"):null,d=new Date,f=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,g={},u=e&&e.get("month")?e.get("month"):null;n==="inspeksi"?g={status:"Done",activity_type:"Inspeksi Hygiene",month:u}:n==="gcdc"?g={status:"Done",activity_type:"GCDC",month:u}:n&&n.startsWith("period_")&&(g={period:n.replace("period_","").toUpperCase()});let c=new Date().getFullYear(),h=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((m,p)=>{let v=String(p+1).padStart(2,"0");return{value:`${c}-${v}`,label:`${m} ${c}`}});P({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:g,onDataLoaded:m=>m.sort((p,v)=>{let w=p.opening_date?new Date(p.opening_date).getTime():0;return(v.opening_date?new Date(v.opening_date).getTime():0)-w}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:m=>kt(m)},{key:"period",label:"Periode",render:m=>me(m)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:m=>l(m)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:m=>l(m)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:m=>l(m)},{key:"status",label:"Status",render:m=>G(m)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:h},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ve}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:_t,value:m?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:m?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period},{name:"pic",label:"PIC",type:"combobox",options:a(m?.pic),value:m?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:m?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:m?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:m?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:m?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let m=await _(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(m.ok){let p=m.data.data.map(v=>({Cabang:v.branch_name||"",Kegiatan:v.activity_type||"",Periode:v.period||"",PIC:v.pic||"","Tgl Opening":v.opening_date||"","Tgl Target":v.target_date||"","Tgl Selesai":v.completion_date||"",Status:v.status||""}));L(p,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async m=>{let v=(await _("/api/branches?all=1")).data?.data||[],w=C=>{if(!C)return null;let $=String(C||"").toLowerCase().trim(),I=$.match(/^(\d{3})/);if(I){let q=v.find(W=>String(W.code||"")===I[1]);if(q)return q.id}let D=v.find(q=>String(q.full_name||"").toLowerCase()===$||String(q.code||"").toLowerCase()===$||String(q.name||"").toLowerCase()===$);return D?D.id:null},T=C=>{if(C==null||C==="")return"";if(C instanceof Date&&!isNaN(C.getTime()))return C.toISOString().slice(0,10);let $=String(C).trim();if($===""||$==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test($))return $.slice(0,10);if(/^\d{4,5}$/.test($)){let D=Number($);if(D>2e4&&D<99999){let q=new Date(Date.UTC(1899,11,30)+D*864e5);return isNaN(q.getTime())?"":q.toISOString().slice(0,10)}}let I=$.split(/[\/\-\.]/);if(I.length===3){let[D,q,W]=I.map(ae=>ae.trim());if(D.length===4&&q.length<=2&&W.length<=2)return`${D}-${q.padStart(2,"0")}-${W.padStart(2,"0")}`;if(W.length===4&&q.length<=2&&D.length<=2)return`${W}-${q.padStart(2,"0")}-${D.padStart(2,"0")}`}return $},k=m.map(C=>({branch_id:w(String(C.Cabang||"").trim()),branch_name:String(C.Cabang||"").trim(),activity_type:String(C.Kegiatan||"").trim(),period:String(C.Periode||"").trim(),pic:String(C.PIC||C.Pic||"").trim(),opening_date:T(C["Tgl Opening"]||C["Tanggal Opening"]||C["Tgl Openir"]),target_date:T(C["Tgl Target"]||C["Tanggal Target"]),completion_date:T(C["Tgl Selesai"]||C["Tanggal Selesai"]),status:String(C.Status||"").trim(),notes:String(C.Catatan||C.Keterangan||"").trim()})).filter(C=>C.activity_type&&C.period),E=await _("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:k,onDuplicate:"update"})});if(!E.ok)throw new Error(E.data?.error||"Import gagal");return ce("schedule"),E.data}}})}A();R();var Ct=[],dt=[];function ct(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),l=new Date(i.value),o=Math.floor((l-r)/864e5);a.value=isNaN(o)?"":o}else a.value=""}}document.body.removeEventListener("input",ct);document.body.addEventListener("input",ct);document.body.removeEventListener("change",ct);document.body.addEventListener("change",ct);function qa(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function Jt(t,e){let i=e?e.get("dash_filter"):null;Ct=await j(),dt=await te();let a=n=>n&&!dt.find(d=>d.value===n)?[...dt,{value:n,label:n}]:dt,r=new Date().getFullYear(),l=["2025","2026","2027","2028","2029","2030"],s=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,d)=>{let f=String(d+1).padStart(2,"0");return{value:`${r}-${f}`,label:`${n} ${r}`}});P({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>i?n.filter(d=>qa(d,i)):n,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"...":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"...":n||"-"}</span>`},{key:"status",label:"Status",render:n=>G(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Day",render:n=>n??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Ct},{type:"select",name:"month",label:"Bulan",options:s},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ct,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date},{name:"day_count",label:"Day",type:"number",value:n?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await _(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let d=n.data.data.map(f=>({"Tanggal Info":f.report_date||"",Cabang:f.branch_name||"",Kategori:f.category||"","Sumber Laporan":f.source||"",Keluhan:f.complaint||"","Nama FC":f.employee_name||"","FC Spesialis":f.fc_specialist||"",Solusi:f.solution||"",Status:f.status||"","Tanggal Selesai":f.completion_date||"",Day:f.day_count!==null?f.day_count:""}));L(d,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async n=>{let f=(await _("/api/branches?all=1")).data?.data||[],g=b=>{if(!b)return null;let h=String(b||"").toLowerCase(),m=f.find(p=>String(p.full_name||"").toLowerCase()===h||String(p.code||"").toLowerCase()===h||String(p.name||"").toLowerCase()===h);return m?m.id:null},u=n.map(b=>({branch_id:g(String(b.Cabang||"").trim()),report_date:String(b["Tanggal Info"]||b.Tanggal||"").trim(),category:String(b.Kategori||"").trim(),source:String(b["Sumber Laporan"]||b.Sumber||"").trim(),complaint:String(b.Keluhan||"").trim(),employee_name:String(b["Nama FC"]||"").trim(),fc_specialist:String(b["FC Spesialis"]||"").trim(),solution:String(b.Solusi||"").trim(),completion_date:String(b["Tanggal Selesai"]||b["Tgl Selesai"]||"").trim(),status:String(b.Status||"").trim(),day_count:b.Day||b.Hari||null})).filter(b=>b.report_date&&b.complaint&&b.category),c=await _("/api/import/issues",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}A();var Me=[];function Ha(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function zt(t,e){let i=e?e.get("dash_filter"):null;Me=await j();let a=await te(),r=["Ade","Berlin"],l=s=>s&&!a.find(n=>n.value===s)?[...a,{value:s,label:s}]:a,o=s=>s&&!r.find(n=>(typeof n=="object"?n.value:n)===s)?[...r,s]:r;P({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:s=>i?s.filter(n=>Ha(n,i)):s,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:s=>`<span title="${s||""}">${s?.length>50?s.slice(0,50)+"\u2026":s||"-"}</span>`},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>G(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),d=await _(`/api/one-on-one?limit=10000&${n}`);if(d.ok){let f=d.data.data.map(u=>({Tanggal:u.meeting_date||"",Cabang:u.branch_name||"","Nama Karyawan":u.employee_name||"",PIC:u.pic||"",Masalah:u.problem||"",Solusi:u.solution||"",Status:u.status||"","Tgl Selesai":u.completion_date||"",Dokumen:u.document_link||""})),{downloadExcel:g}=await Promise.resolve().then(()=>(R(),le));g(f,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),le));n(s,"Template_Import_OneOnOne")},onImport:async s=>{let n=u=>{if(!u)return null;let c=String(u||"").toLowerCase(),b=Me.find(h=>String(h.label||"").toLowerCase()===c);return b?b.value:null},d=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let c=String(u).trim();if(/^\d{4,5}$/.test(c)){let h=Number(c);if(h>2e4&&h<99999){let m=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);let b=c.split(/[\/\-\.]/);if(b.length===3){let[h,m,p]=b.map(v=>v.trim());if(h.length===4&&m.length<=2&&p.length<=2)return`${h}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&h.length<=2)return`${p}-${m.padStart(2,"0")}-${h.padStart(2,"0")}`}return c},f=s.map(u=>({meeting_date:d(u.Tanggal),employee_name:String(u["Nama Karyawan"]||"").trim(),branch_id:n(String(u.Cabang||"").trim()),pic:String(u.PIC||"").trim(),problem:String(u.Masalah||"").trim(),solution:String(u.Solusi||"").trim(),status:String(u.Status||"").trim(),completion_date:d(u["Tgl Selesai"]),document_link:String(u.Dokumen||"").trim()})).filter(u=>u.meeting_date&&u.employee_name&&u.branch_id),g=await _("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:f,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}},formFields:s=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:s?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:s?.branch_id&&!Me.find(n=>n.value==s.branch_id)?[...Me,{value:s.branch_id,label:s.branch_name||s.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:l(s?.employee_name),value:s?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(s?.pic),createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:s?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:s?.document_link}]})}A();async function Ut(t){let e=await j(),i=await te(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=n=>n&&!i.find(d=>d.value===n)?[...i,{value:n,label:n}]:i,l=n=>n&&!a.find(d=>(typeof d=="object"?d.value:d)===n)?[...a,n]:a,o=Array.from({length:5},(n,d)=>String(new Date().getFullYear()-d));P({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let d=JSON.parse(n);return Array.isArray(d)?d.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let d=new URLSearchParams(n||{}).toString(),f=await _(`/api/training?limit=10000&${d}`);if(f.ok){let g=f.data.data.map(c=>{let b=c.participants||"";try{let h=JSON.parse(b);b=Array.isArray(h)?h.join(", "):b}catch{}return{Tanggal:c.training_date||"",Batch:c.batch||"",Materi:c.subject||"",Cabang:c.branch_name||"",Trainer:c.trainer||"",Peserta:b,Nilai:c.score!==null&&c.score!==void 0?c.score:"",Dokumen:c.document_link||""}}),{downloadExcel:u}=await Promise.resolve().then(()=>(R(),le));u(g,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:d}=await Promise.resolve().then(()=>(R(),le));d(n,"Template_Import_Training")},onImport:async n=>{let d=c=>{if(!c)return null;let b=String(c||"").toLowerCase(),h=e.find(m=>String(m.label||"").toLowerCase()===b);return h?h.value:null},f=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let b=String(c).trim();if(/^\d{4,5}$/.test(b)){let m=Number(b);if(m>2e4&&m<99999){let p=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let h=b.split(/[\/\-\.]/);if(h.length===3){let[m,p,v]=h.map(w=>w.trim());if(m.length===4&&p.length<=2&&v.length<=2)return`${m}-${p.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&p.length<=2&&m.length<=2)return`${v}-${p.padStart(2,"0")}-${m.padStart(2,"0")}`}return b},g=n.map(c=>({training_date:f(c.Tanggal),batch:String(c.Batch||"").trim(),subject:String(c.Materi||"").trim(),branch_id:d(String(c.Cabang||"").trim()),trainer:String(c.Trainer||"").trim(),participants:String(c.Peserta||"").trim(),score:c.Nilai?Number(c.Nilai):null,document_link:String(c.Dokumen||"").trim()})).filter(c=>c.training_date&&c.subject&&c.branch_id),u=await _("/api/training/import",{method:"POST",body:JSON.stringify(g)});if(!u.ok)throw new Error(u.data?.error||"Import gagal")}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:l(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let d=JSON.parse(n?.participants);return Array.isArray(d)?d.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>n})}A();xe();R();function Gt({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:l,filterFields:o,defaultFilters:s={},itemLabel:n="Data",canCreate:d=!0,canEdit:f=!0,canDelete:g=!0,onBeforeSubmit:u,onAfterLoad:c,onDataLoaded:b,extraActions:h=[],initialSearch:m="",exportOptions:p=null,bulkDelete:v=!1,paginationMode:w="server"}){let T=re();T&&T.role==="viewer"&&(d=!1,f=!1,g=!1,v=!1,p=null),T&&T.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(d=!1,f=!1,g=!1,v=!1,p=null),T&&T.role==="input_lembur"&&(a==="/api/overtime"?(g=!1,v=!1,p=null):(d=!1,f=!1,g=!1,v=!1,p=null));let k=1,E={...s};m&&(E.search=m);let C=new Set;t.innerHTML=`
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
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${p?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${p.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${p.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${p.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${p.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
        ${o.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${E.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let S=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${S}</option>${(y.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${E[y.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let S=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${S}</option>${(y.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${E[y.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let y=document.getElementById("bulk-toolbar");if(!y)return;let S=document.getElementById("bulk-count"),x=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");S.textContent=`${C.size} item dipilih`,C.size>0?(y.style.display="flex",x.disabled=!1,O.disabled=!1):(y.style.display="none",x.disabled=!0,O.disabled=!0);let N=document.getElementById("select-all-checkbox");if(N){let U=document.querySelectorAll(".row-checkbox");if(U.length>0){let B=[...U].every(oe=>oe.checked),F=[...U].some(oe=>oe.checked);N.checked=B,N.indeterminate=F&&!B}else N.checked=!1,N.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{C.clear(),document.querySelectorAll(".row-checkbox").forEach(S=>S.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(C.size===0)return;let y=[...C],S=document.createElement("div");S.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",S.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(S),S.querySelector("#bulk-cancel-btn").addEventListener("click",()=>S.remove()),S.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let x=S.querySelector("#bulk-confirm-btn");x.disabled=!0,x.textContent="Menghapus...";let O=await _(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});S.remove(),O.ok?(Q(`${y.length} ${n} berhasil dihapus.`),C.clear(),$(),ce(a),M()):z(O.data?.error||"Gagal menghapus data.")})});let I=document.getElementById("filter-search"),D;I?.addEventListener("input",y=>{clearTimeout(D),D=setTimeout(()=>{E.search=y.target.value,k=1,C.clear(),$(),M()},400)}),o?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",S=>{E[y.name]=S.target.value;let x=document.getElementById(`filter-sheet-${y.name}`);x&&(x.value=S.target.value),k=1,C.clear(),$(),M()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",S=>{E[y.name]=S.target.value;let x=document.getElementById(`filter-${y.name}`);x&&(x.value=S.target.value),k=1,C.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{E={},I&&(I.value=""),o?.forEach(y=>{let S=document.getElementById(`filter-${y.name}`);S&&(S.value="");let x=document.getElementById(`filter-sheet-${y.name}`);x&&(x.value="")}),k=1,C.clear(),$(),M()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{E={},I&&(I.value=""),o?.forEach(y=>{let S=document.getElementById(`filter-${y.name}`);S&&(S.value="");let x=document.getElementById(`filter-sheet-${y.name}`);x&&(x.value="")}),k=1,C.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>Y(null)),p&&document.addEventListener("click",function(y){let S=document.getElementById("aksi-menu-main"),x=document.getElementById("btn-aksi-main");S&&x&&!x.contains(y.target)&&!S.contains(y.target)&&S.classList.remove("show-aksi-menu")});let q=document.getElementById("btn-mobile-filter"),W=document.getElementById("filter-options-wrapper"),ae=document.getElementById("btn-close-filter-sheet");if(q&&W&&(q.addEventListener("click",y=>{y.preventDefault(),W.classList.add("sheet-open")}),ae&&ae.addEventListener("click",y=>{y.preventDefault(),W.classList.remove("sheet-open")})),p){document.getElementById(`btn-export-${p.moduleName}`)?.addEventListener("click",async S=>{let x=S.target,O=x.innerHTML;x.innerHTML="\u23F3 Loading...",x.disabled=!0;try{await p.onExport()}catch{z("Gagal export data")}finally{x.innerHTML=O,x.disabled=!1}}),document.getElementById(`btn-template-${p.moduleName}`)?.addEventListener("click",()=>{p.onTemplate()});let y=document.getElementById(`input-import-${p.moduleName}`);y?.addEventListener("change",async S=>{let x=S.target.files[0];if(!x)return;let O=document.getElementById(`label-import-${p.moduleName}`),N=O?O.querySelector(".import-text"):null,U=N?N.innerText:"";N&&(N.innerText="\u231B Memproses..."),O&&(O.style.pointerEvents="none"),y.disabled=!0;try{let B=await Ue(x);if(B.length===0)throw new Error("File kosong atau format salah");await p.onImport(B),Q("Import berhasil!"),ce(a),M()}catch(B){z(B.message||"Gagal import data")}finally{N&&(N.innerText=U),O&&(O.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function M(){$();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let S=w==="client",x=S?1:k,O=S?Te:20,N=new URLSearchParams({page:x,limit:O,...Object.fromEntries(Object.entries(E).filter(([,H])=>H))}),U=await _(`${a}?${N}`);if(!U.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let B=U.data?.data||U.data||[],F=U.data?.pagination,oe=B.length;if(S){B=b(B);let H=B.length,X=20,ne=Math.ceil(H/X);k>ne&&ne>0&&(k=ne);let K=(k-1)*X,se=k*X;B=B.slice(K,se),F={page:k,limit:X,total:H,pages:ne}}!1,c&&c(B);let ve=Ze({columns:r,data:B,onEdit:f?H=>Y(H):null,actions:h.map(H=>({...H,handler:X=>H.handler(X,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:v?{selectedIds:C,onToggle:$}:null});y.innerHTML="",y.appendChild(ve);let pe=document.getElementById("pagination-container");if(pe&&(pe.innerHTML="",F&&F.pages>1)){let H=et({page:F.page,pages:F.pages,total:F.total,limit:F.limit,onPage:X=>{k=X,M()}});H&&pe.appendChild(H)}}function $e(y){let S=typeof l=="function"?l(y):l;return ze(S)}function Y(y){let S=!!y,x=document.createElement("form");if(x.noValidate=!0,x.innerHTML=$e(y),S){let N=typeof l=="function"?l(y):l;at(x,y)}let{close:O}=de({title:S?`Edit ${n}`:`Tambah ${n}`,content:x,size:"lg",confirmText:S?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(N,U)=>{if(!x.reportValidity())return;let B=N.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let F=tt(x),oe=typeof l=="function"?l(y):l,ve=async ne=>{for(let K of ne)if(K.type==="row")await ve(K.fields);else if(K.type==="combobox"&&F[K.name]){let se=F[K.name],ke=(K.options||[]).find(Z=>{let ie=String(typeof Z=="object"?Z.value:Z),ut=String(typeof Z=="object"?Z.label:Z);return ie===se||ut===se});if(ke)F[K.name]=typeof ke=="object"?ke.value:ke;else if(K.createApi){let Z={};Z[K.createApi.field]=se,K.createApi.extra&&Object.assign(Z,K.createApi.extra);let ie=await _(K.createApi.path,{method:"POST",body:JSON.stringify(Z)});if(ie.ok&&ie.data?.id)F[K.name]=ie.data.id;else if(ie.ok&&!ie.data?.id)F[K.name]=se;else throw new Error(`Gagal membuat master data: ${ie.data?.error||"Unknown error"}`)}}};try{await ve(oe)}catch(ne){z(ne.message),B.disabled=!1,B.textContent=S?"Simpan Perubahan":`Tambah ${n}`;return}u&&(F=await u(F,y));let pe=S?"PUT":"POST",H=S?`${a}/${y.id}`:a,X=await _(H,{method:pe,body:JSON.stringify(F)});X.ok?(Q(S?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),ce(a),M()):(z(X.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=S?"Simpan Perubahan":`Tambah ${n}`)}})}function Ee(y){De(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let S=await _(`${a}/${y.id}`,{method:"DELETE"});S.ok?(Q(`${n} berhasil dihapus.`),ce(a),M()):z(S.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}A();R();async function Vt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await j(),a=await te(),r=e?e.get("dash_filter"):null,l={};if(r==="reliever"){let u=new Date,c=`${u.getFullYear()}-${String(u.getMonth()+1).padStart(2,"0")}`;l={status:"Done",month:e&&e.get("month")?e.get("month"):c}}console.log("RAW",await Ce()),console.log("OPTIONS",a);let o=u=>u&&!a.find(c=>c.value===u)?[...a,{value:u,label:u}]:a,s=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=u=>u&&!s.includes(u)?[...s,u]:s,d=new Date().getFullYear(),g=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((u,c)=>{let b=String(c+1).padStart(2,"0");return{value:`${d}-${b}`,label:`${u} ${d}`}});Gt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:l,onDataLoaded:u=>u.sort((c,b)=>{let h=c.backup_date?new Date(c.backup_date).getTime():0;return(b.backup_date?new Date(b.backup_date).getTime():0)-h}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:u=>me(u)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:u=>window.formatDate(u)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:u=>window.formatDate(u)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:u=>u?`<span class="badge badge-info">${u}</span>`:"-"},{key:"status",label:"Status",render:u=>G(u)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:s},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:g},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:u?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:u?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(u?.original_fc_name),value:u?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(u?.reliever_name),value:u?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:u?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:u?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:u?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:u?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let u=await _(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(u.ok){let c=u.data.data.map(b=>({Cabang:b.branch_name||"","Nama Facility care":b.original_fc_name||"",Periode:b.period||"",Relifer:b.reliever_name||"","Tanggal Back Up":b.backup_date||"","Tanggal Selesai":b.completion_date||"",Keterangan:b.reason||"",Shift:b.shift||"",Status:b.status||""}));c.length===0&&c.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),L(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async u=>{let b=(await _("/api/branches?all=1")).data?.data||[],h=v=>{if(!v)return null;let w=String(v||"").toLowerCase(),T=b.find(k=>String(k.full_name||"").toLowerCase()===w||String(k.code||"").toLowerCase()===w||String(k.name||"").toLowerCase()===w);return T?T.id:null},m=u.map(v=>({branch_name:String(v.Cabang||"").trim(),backup_date:String(v["Tanggal Back Up"]||v["Tanggal Backup"]||"").trim(),original_fc_name:String(v["Nama Facility care"]||v["FC Digantikan"]||"").trim(),reliever_name:String(v.Relifer||v.Reliefer||"").trim(),period:String(v.Periode||"").trim(),reason:String(v.Keterangan||"").trim(),shift:String(v.Shift||"").trim(),completion_date:String(v["Tanggal Selesai"]||"").trim(),status:String(v.Status||"").trim()})).filter(v=>v.reliever_name&&v.backup_date),p=await _("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}A();R();async function Qt(t){let e=await j(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));P({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>me(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>G(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),l=await _(`/api/reports/inspection?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||""}));L(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let r=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),f=e.find(g=>String(g.label||"").toLowerCase()===d);return f?f.value:null},l=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let d=String(n).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let u=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let f=d.split(/[\/\-\.]/);if(f.length===3){let[g,u,c]=f.map(b=>b.trim());if(g.length===4&&u.length<=2&&c.length<=2)return`${g}-${u.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&u.length<=2&&g.length<=2)return`${c}-${u.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:l(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),s=await _("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}A();R();async function Yt(t){let e=await j(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));P({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>me(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>G(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),l=await _(`/api/reports/cleaning?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));L(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let r=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),f=e.find(g=>String(g.label||"").toLowerCase()===d);return f?f.value:null},l=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let d=String(n).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let u=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let f=d.split(/[\/\-\.]/);if(f.length===3){let[g,u,c]=f.map(b=>b.trim());if(g.length===4&&u.length<=2&&c.length<=2)return`${g}-${u.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&u.length<=2&&g.length<=2)return`${c}-${u.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:l(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),s=await _("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}A();R();async function Wt(t,e){let i=await j(),a=Array.from({length:4},(o,s)=>String(new Date().getFullYear()-s)),r=e?e.get("dash_filter"):null,l={};if(r==="fogging"){let o=new Date,s=String(o.getMonth()+1).padStart(2,"0"),n=String(o.getFullYear()),d=e?e.get("month"):null;d&&d.length===7&&(n=d.split("-")[0],s=d.split("-")[1]),l={status:"Done",month:s,year:n}}P({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:l,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>me(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>G(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let s=new URLSearchParams(o||{}).toString(),n=await _(`/api/reports/fogging?limit=10000&${s}`);if(n.ok){let d=n.data.data.map(f=>({Cabang:f.branch_name||"",Jenis:f.activity_type||"Fogging",Periode:f.period||"",Tanggal:f.activity_date||"",Status:f.status||"","Link Dokumen":f.document_link||""}));L(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let s=g=>{if(!g)return null;let u=String(g||"").toLowerCase(),c=i.find(b=>String(b.label||"").toLowerCase()===u);return c?c.value:null},n=g=>{if(g==null||g==="")return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let u=String(g).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let b=Number(u);if(b>2e4&&b<99999){let h=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let c=u.split(/[\/\-\.]/);if(c.length===3){let[b,h,m]=c.map(p=>p.trim());if(b.length===4&&h.length<=2&&m.length<=2)return`${b}-${h.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&h.length<=2&&b.length<=2)return`${m}-${h.padStart(2,"0")}-${b.padStart(2,"0")}`}return u},d=o.map(g=>({branch_id:s(String(g.Cabang||"").trim()),activity_type:String(g.Jenis||g.Kegiatan||"Fogging").trim(),period:String(g.Periode||"").trim(),activity_date:n(g.Tanggal),status:String(g.Status||"").trim(),document_link:String(g["Link Dokumen"]||"").trim(),notes:String(g.Catatan||g.Keterangan||"").trim()})).filter(g=>g.branch_id&&g.period&&g.activity_date),f=await _("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}}})}A();R();async function Xt(t){let e=await j(),i=await te(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(s,n)=>String(new Date().getFullYear()-n)),l=s=>s&&!i.find(n=>n.value===s)?[...i,{value:s,label:s}]:i,o=s=>s&&!a.find(n=>n.value===s)?[...a,{value:s,label:s}]:a;P({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:s=>`<span title="${s||""}">${s?.length>60?s.slice(0,60)+"\u2026":s||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:s=>window.formatDate(s)},{key:"status",label:"Status",render:s=>G(s)},{key:"notes",label:"Keterangan",render:s=>s?.length>40?s.slice(0,40)+"\u2026":s||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:s?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(s?.pic),value:s?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:s?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:s?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:s?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:s?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),d=await _(`/api/reports/basecamp?limit=10000&${n}`);if(d.ok){let f=d.data.data.map(g=>({"Tgl Info":g.info_date||"",Cabang:g.branch_name||"",Permasalahan:g.problem||"",PIC:g.pic||"","Tgl Done":g.done_date||"",Status:g.status||"",Keterangan:g.notes||""}));L(f,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async s=>{let n=u=>{if(!u)return null;let c=String(u||"").toLowerCase(),b=e.find(h=>String(h.label||"").toLowerCase()===c);return b?b.value:null},d=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let c=String(u).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let h=Number(c);if(h>2e4&&h<99999){let m=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let b=c.split(/[\/\-\.]/);if(b.length===3){let[h,m,p]=b.map(v=>v.trim());if(h.length===4&&m.length<=2&&p.length<=2)return`${h}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&h.length<=2)return`${p}-${m.padStart(2,"0")}-${h.padStart(2,"0")}`}return c},f=s.map(u=>({info_date:d(u["Tgl Info"]||u["Tanggal Info"]),branch_id:n(String(u.Cabang||"").trim()),problem:String(u.Permasalahan||"").trim(),pic:String(u.PIC||"").trim(),done_date:d(u["Tgl Done"]||u["Tanggal Done"]),status:String(u.Status||"").trim(),notes:String(u.Keterangan||u.Catatan||"").trim()})).filter(u=>u.info_date&&u.branch_id&&u.problem),g=await _("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(f)});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}async function Zt(t){P({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ie)),r=await a(`/api/sop?limit=10000&${i}`);if(r.ok){let l=r.data.data.map(s=>({"Nama SOP":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Catatan:s.notes||s.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(R(),le));o(l,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),le));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(l=>({name:String(l["Nama SOP"]||"").trim(),category:String(l.Kategori||"").trim(),document_link:String(l.Dokumen||"").trim(),description:String(l.Catatan||"").trim()})).filter(l=>l.name),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ie)),r=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function ea(t){P({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ie)),r=await a(`/api/checklist?limit=10000&${i}`);if(r.ok){let l=r.data.data.map(s=>({"Nama Checklist":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Deskripsi:s.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(R(),le));o(l,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),le));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(l=>({name:String(l["Nama Checklist"]||"").trim(),category:String(l.Kategori||"").trim(),document_link:String(l.Dokumen||"").trim(),description:String(l.Deskripsi||"").trim()})).filter(l=>l.name),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ie)),r=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}A();R();var Ja=["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas","Teras","Lobby 1","Receptionist","Snack Corner","DU 101","Ronsen","Toilet 1","Janitor","Ruang Limbah","Ruang Steril","Area Tangga","Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Dokter","Ruang Team","Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Lift","Area Genset","Area Tangga Exit","Taman","Pos Security","Toilet Security","Halaman Parkir"];async function ta(t){let e=Ja;try{let i=await _("/api/hygiene-standards/rooms");i.ok&&Array.isArray(i.data)&&i.data.length>0&&(e=i.data)}catch{}P({container:t,title:"Master Hygiene",icon:"\u2728",apiPath:"/api/hygiene-standards",bulkDelete:!1,enableMobileFilterSheet:!0,itemLabel:"Parameter Kebersihan",columns:[{key:"room_name",label:"Ruangan",render:i=>`<span class="badge badge-info" style="font-weight:600; white-space:nowrap;">${i||"-"}</span>`},{key:"item_name",label:"Item / Objek",render:i=>`<strong style="color:var(--text-1);">${i||"-"}</strong>`},{key:"cleanliness_standard",label:"Standar Kebersihan",render:i=>`<span style="line-height:1.5; color:var(--text-2);">${i||"-"}</span>`}],filterFields:[{type:"search",placeholder:"Cari ruangan, item, atau standar kebersihan..."},{type:"select",name:"room_name",label:"Ruangan",options:e}],formFields:i=>[{type:"row",fields:[{name:"room_name",label:"Ruangan / Kategori",type:"select",required:!0,options:e,value:i?.room_name||""},{name:"item_name",label:"Item / Objek",required:!0,placeholder:"Nama item/objek pembersihan",value:i?.item_name||""}]},{name:"cleanliness_standard",label:"Standar Kebersihan Fisik",type:"textarea",required:!0,rows:3,placeholder:"Kondisi fisik yang dipersyaratkan (misal: Bebas debu, tidak lengket, kering...)",value:i?.cleanliness_standard||""}],exportOptions:{moduleName:"hygiene_standards",onExport:async i=>{let a=new URLSearchParams(i||{}).toString(),r=await _(`/api/hygiene-standards?all=1&${a}`);if(r.ok){let o=(r.data?.data||r.data||[]).map(s=>({Ruangan:s.room_name||"","Item/Objek":s.item_name||"","Standar Kebersihan":s.cleanliness_standard||""}));L(o,`Master_Hygiene_FCMS_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data master hygiene")},onTemplate:async()=>{L([{Ruangan:"Teras","Item/Objek":"Sudut/Pojok","Standar Kebersihan":"Bebas noda, kerak lumut, puntung rokok, dan tumpukan kotoran kering"},{Ruangan:"Lobby 1","Item/Objek":"Lantai Lobby","Standar Kebersihan":"Mengkilap bersih, bebas minyak, tidak licin, nat ubin cerah"}],"Template_Master_Hygiene_FCMS")},onImport:async i=>{let a=i.map(l=>({room_name:String(l.Ruangan||l.room_name||"").trim(),item_name:String(l["Item/Objek"]||l.Item||l.item_name||"").trim(),cleanliness_standard:String(l["Standar Kebersihan"]||l.Standar||l.cleanliness_standard||"").trim()})).filter(l=>l.room_name&&l.item_name&&l.cleanliness_standard);if(a.length===0)throw new Error("Tidak ada baris valid untuk diimpor. Pastikan header: Ruangan, Item/Objek, Standar Kebersihan");let r=await _("/api/hygiene-standards/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}A();xe();R();async function $t(t,e="forms"){if(e==="supply")return Ua(t);za(t)}function za(t){P({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await _(`/api/forms?limit=10000&${i}`);a.data?.data?L(a.data.data,"Data_Master_Form"):z("Gagal export data master form")},onImport:async e=>{let i=await _("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ua(t){let i=((await _("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));P({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>G(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let r=a?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let l=a?.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(o=>o.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:l},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),l=await _(`/api/reports/supply?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(s=>{let n=s.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let d=s.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:s.submitted_at||"",Pengirim:s.submitter_name||"",Cabang:s.branch_name_ref||s.branch_name||"","Alat/Barang":n||"",Chemical:d||"",Catatan:s.additional_notes||"",Status:s.status||"","Diproses Oleh":s.processed_by||""}});L(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let l=(await _("/api/branches?all=1")).data?.data||[],o=f=>{if(!f)return null;let g=String(f||"").toLowerCase(),u=l.find(c=>String(c.full_name||"").toLowerCase()===g||String(c.code||"").toLowerCase()===g||String(c.name||"").toLowerCase()===g);return u?u.id:null},s=f=>{if(f==null||f==="")return"";if(f instanceof Date&&!isNaN(f.getTime()))return f.toISOString().slice(0,10);let g=String(f).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let c=Number(g);if(c>2e4&&c<99999){let b=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let u=g.split(/[\/\-\.]/);if(u.length===3){let[c,b,h]=u.map(m=>m.trim());if(c.length===4&&b.length<=2&&h.length<=2)return`${c}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&b.length<=2&&c.length<=2)return`${h}-${b.padStart(2,"0")}-${c.padStart(2,"0")}`}return g},n=a.map(f=>({submitted_at:s(f.Waktu||f.Tanggal),submitter_name:String(f.Pengirim||"").trim(),branch_id:o(String(f.Cabang||"").trim()),tools_items:String(f["Alat/Barang"]||f.Alat||"").trim(),chemical_items:String(f.Chemical||"").trim(),additional_notes:String(f.Catatan||f.Keterangan||"").trim(),status:String(f.Status||"").trim(),processed_by:String(f["Diproses Oleh"]||f.PIC||"").trim()})).filter(f=>f.submitted_at&&f.submitter_name&&f.branch_id),d=await _("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,r)=>{let l=de({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,s)=>{let n=o.querySelector("#supply-status").value,d=o.querySelector("#supply-processed-by").value;(await _(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:d})})).ok?(Q("Status diperbarui."),s(),r()):z("Gagal update status.")}})}}]})}A();R();async function aa(t){let e=re();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}P({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await _(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));L(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(l=>({full_name:String(l["Nama Lengkap"]||"").trim(),username:String(l.Username||"").trim(),email:String(l.Email||"").trim(),role:String(l.Role||"").trim()||"viewer",password:String(l.Password||"").trim()})).filter(l=>l.username&&l.password&&l.email&&l.full_name),r=await _("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}A();R();async function na(t){P({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await _(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)L(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),a=await _("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}A();async function ia(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",r));async function a(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await _(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),i=[]}}async function r(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let s=e.getFullYear(),n=e.getMonth(),d=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),f=document.getElementById("cal-month-label");f&&(f.textContent=d);let g=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),u=i.filter(k=>g.has(k.type)),c={};u.forEach(k=>{let E=Ga(k.event_date);E&&(c[E]||(c[E]=[]),c[E].push(k))});let b=new Date(s,n,1).getDay(),h=new Date(s,n+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],p=new Date().toISOString().slice(0,10),v='<div class="calendar-grid">';m.forEach(k=>{v+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<b;k++)v+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=h;k++){let E=`${s}-${String(n+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,C=c[E]||[],$=E===p;v+=`
          <div class="cal-cell ${$?"cal-today":""} ${C.length?"cal-has-events":""}"
               data-date="${E}" tabindex="0" role="button" aria-label="${E}">
            <div class="cal-day-num ${$?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${C.slice(0,3).map(I=>`
                <div class="cal-event-dot cal-color-${I.color||"gray"}" title="${he(I.title||I.type)}">
                  <span class="cal-event-dot-label">${Va(I.title||I.branch_name||I.type,18)}</span>
                </div>
              `).join("")}
              ${C.length>3?`<div class="cal-more">+${C.length-3} lagi</div>`:""}
            </div>
          </div>`}let T=(b+h)%7;if(T!==0)for(let k=0;k<7-T;k++)v+='<div class="cal-cell cal-cell-empty"></div>';v+="</div>",o.innerHTML=v,o.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let E=k.dataset.date,C=c[E]||[];if(!C.length)return;let $=document.getElementById("cal-event-list"),I=new Date(E+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=I,document.getElementById("cal-event-items").innerHTML=C.map(D=>`
            <div class="cal-event-item cal-color-border-${D.color||"gray"}">
              <div class="cal-event-type">
                ${Qa(D.type)}
                ${D.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${he(D.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${he(D.title||"-")}</div>
              ${D.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${he(D.branch_name)}</div>`:""}
              ${D.pic||D.period||D.opening_date||D.target_date||D.completion_date||D.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${D.pic?`<div>\u{1F464} <b>PIC:</b> ${he(D.pic)}</div>`:""}
                  ${D.period?`<div>\u{1F4C5} <b>Periode:</b> ${he(D.period)}</div>`:""}
                  ${D.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${he(Et(D.opening_date))}</div>`:""}
                  ${D.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${he(Et(D.target_date))}</div>`:""}
                  ${D.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${he(Et(D.completion_date))}</div>`:""}
                  ${D.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${he(D.notes)}</div>`:""}
                </div>
              `:""}
              ${D.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${D.status==="Done"?"badge-success":"badge-warning"}">${he(D.status)}</span></div>`:""}
              ${D.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${D.days_remaining} hari</div>`:""}
            </div>
          `).join(""),$.style.display="block"})})}catch(s){console.error("[Calendar] Render error:",s),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let l=()=>{a().then(()=>r())};nt.on("data:changed",l),r()}function Ga(t){if(!t||t==="-"||String(t).trim()==="")return"";let e=String(t).trim();if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);let i=e.split(/[\/\-\.]/);if(i.length===3){let[a,r,l]=i.map(o=>o.trim());if(l.length===4)return`${l}-${r.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${r.padStart(2,"0")}-${l.padStart(2,"0")}`}return e.slice(0,10)}function Et(t){if(!t||t==="-"||String(t).trim()==="")return"";let i=String(t).trim().split(/[\/\-\.]/);if(i.length===3){let[a,r,l]=i.map(o=>o.trim());if(a.length===4)return`${l.padStart(2,"0")}-${r.padStart(2,"0")}-${a}`;if(l.length===4)return`${a.padStart(2,"0")}-${r.padStart(2,"0")}-${l}`}return t}function Va(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function he(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Qa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[t]||t}A();async function ra(t){let e=re(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let l=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(l&&o)try{let s=JSON.parse(atob(l.split(".")[1])),n=new Date(s.exp*1e3);o.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let n=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),f=document.getElementById("btn-save-pwd");n.style.display="none",d.style.display="none";let g=s.target,u=g.current_password.value,c=g.new_password.value,b=g.confirm_password.value;if(c!==b){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(c.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}f.disabled=!0,f.textContent="\u23F3 Menyimpan...";let h=await _("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:u,new_password:c})});f.disabled=!1,f.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',h.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",g.reset(),Q("Password berhasil diubah.")):(n.textContent=h.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}A();var pt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ee(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let r=Number(e);if(r>2e4&&r<99999){let l=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(l.getTime())?null:l.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[r,l,o]=i.map(f=>f.trim()),s=Number(r),n=Number(l),d=Number(o);if(r.length===4&&s>1900)return`${r}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&d>1900)return s>12?`${o}-${l.padStart(2,"0")}-${r.padStart(2,"0")}`:n>12?`${o}-${r.padStart(2,"0")}-${l.padStart(2,"0")}`:`${o}-${l.padStart(2,"0")}-${r.padStart(2,"0")}`;if(o.length===2&&!isNaN(d)){let f=d>=50?`19${o}`:`20${o}`;return s>12?`${f}-${l.padStart(2,"0")}-${r.padStart(2,"0")}`:`${f}-${l.padStart(2,"0")}-${r.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function oa(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ya={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||t.Divisi||"FACILITY CARE",phone:t["No. Hp"]||t["No. HP"],join_date:ee(t["Tanggal Masuk"]||t["Tgl Masuk"]),status:t.Status||"",target_pindah_os:ee(t["Target Pindah OS"]),target_selesai_os:ee(t["Target Selesai OS"]),notes:t.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:ee(t["Tanggal Mulai"]),end_date:ee(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:ee(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:ee(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:ee(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:ee(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:ee(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:ee(t["Tanggal Target"]||t["Tgl Target"]),completion_date:ee(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:ee(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:ee(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:ee(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:ee(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:ee(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:ee(t["Tanggal Back Up"]),completion_date:ee(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:ee(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:ee(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Wa(t,e){let i=pt[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ya[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],l=[],o=[];return e.filter(n=>!oa(n)).forEach((n,d)=>{let f=e.indexOf(n)+2,g=[];a.required.forEach(({key:c,label:b})=>{let h=n[c];if(h==null||String(h).trim()===""){let m=Object.keys(n).filter(p=>p.trim()).join(", ");g.push({column:b,originalValue:h||"",reason:`Kolom "${b}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let u=a.map(n);g.length>0?l.push({row:f,data:u,raw:n,errors:g}):(r.push(n),o.push(u))}),{valid:r,errors:l,mapped:o}}function la(t){let e=[];return t.SheetNames.forEach(i=>{let a=pt[i];if(!a)return;let r=t.Sheets[i],l=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=Wa(i,l),s=l.filter(n=>!oa(n));e.push({sheetName:i,module:a.module,label:a.label,total:s.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function sa(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,r])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(r),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function da(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(r=>{if(!r.errors||r.errors.length===0)return;a=!0;let l=r.errors.map(s=>({"No. Baris":s.row,"Kolom Gagal":(s.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(s.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(s.data||{}).map(([n,d])=>[n,d??""]))})),o=e.utils.json_to_sheet(l);e.utils.book_append_sheet(i,o,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Xa=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ca(t){t.innerHTML=`
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
              ${Object.entries(pt).map(([h,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${h} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,i=null,a=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function l(h){Object.entries(r).forEach(([m,p])=>{p.style.display=m===h?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let h=document.getElementById("btn-backup-db");h.disabled=!0,h.textContent="\u23F3 Memproses Backup...";try{let m=await _("/api/import/backup");if(m.ok){if(!window.XLSX){z("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let p=window.XLSX,v=p.utils.book_new();Object.entries(m.data.database).forEach(([w,T])=>{let k=T.length>0?T:[{}],E=p.utils.json_to_sheet(k);p.utils.book_append_sheet(v,E,w.substring(0,31))}),p.writeFile(v,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Q("Backup berhasil diunduh!")}else z("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){z("Gagal memproses backup: "+m.message)}finally{h.disabled=!1,h.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let h=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let m=await _("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=h,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{sa(),Q("Template Excel berhasil didownload!")});let s=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",h=>{h.stopPropagation(),s.click()}),s.addEventListener("change",h=>{h.target.files[0]&&d(h.target.files[0])}),n.addEventListener("dragover",h=>{h.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",h=>{h.preventDefault(),n.classList.remove("drag-over");let m=h.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?d(m):z("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",l("upload")});async function d(h){e=h,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${h.name} (${(h.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await f(h)}async function f(h){l("validating");let m=document.getElementById("validation-status"),p=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",p.style.width="20%",await Qe(200);let v=await h.arrayBuffer(),w=window.XLSX.read(v,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${w.SheetNames.length} sheet...`,p.style.width="50%",await Qe(100),i=la(w),p.style.width="100%",m.textContent="Validasi selesai!",await Qe(300),g()}catch(v){l("upload"),z("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function g(){l("preview");let h=i.filter($=>!$.skipped).length,m=i.reduce(($,I)=>$+I.total,0),p=i.reduce(($,I)=>$+I.valid,0),v=i.reduce(($,I)=>$+I.errorCount,0),w=m>0?Math.round(p/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${h} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${p} valid (${w}%)</span>
      ${v>0?`<span class="badge badge-danger">${v} error</span>`:""}
    `;let T=document.getElementById("preview-table-container");T.innerHTML=`
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
          ${i.map(($,I)=>`
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
                ${$.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${I}">\u{1F50D} ${$.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,T.querySelectorAll(".btn-detail-error").forEach($=>{$.addEventListener("click",()=>{let I=i[Number($.dataset.idx)];u(I)})});let k=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");E.innerHTML="",k.style.display="none";let C=document.getElementById("btn-start-import");p===0?(C.disabled=!0,C.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(C.disabled=!1,v>0?(C.innerHTML=`\u{1F680} Import ${p} Data Valid (${v} dilewati)`,C.title="Baris error akan dilewati, baris valid tetap diimport"):C.innerHTML=`\u{1F680} Mulai Import ${p} Data`)}function u(h){let m=document.getElementById("error-detail-section"),p=document.getElementById("error-detail-container");m.style.display="";let v=h.errors.slice(0,100).map(w=>(Array.isArray(w.errors)?w.errors:[]).map(k=>{let E=typeof k=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${w.row}</span></td>
            <td><strong>${E?k.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${E&&k.originalValue!==void 0?k.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${E?k.reason:k}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${E&&k.aliases?`Gunakan salah satu nama kolom:<br><em>${k.aliases}</em>`:E&&k.hint?k.hint:""}
            </td>
          </tr>
        `}).join("")).join("");p.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${h.sheetName} \u2014 ${h.errorCount} baris error dari ${h.total} total
          ${h.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${h.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{l("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,s.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;da(i)?Q("Log error berhasil didownload."):Q("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let h=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";c(h)});async function c(h){l("importing"),a=Date.now();let m=[];Xa.forEach(k=>{let E=i?.find(C=>C.module===k&&C.mapped?.length>0);E&&m.push(E)});let p=document.getElementById("import-steps-list");p.innerHTML=m.map(k=>`
      <div class="import-step-item" id="step-item-${k.module}">
        <span class="step-item-icon" id="step-icon-${k.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${k.label} <span class="step-item-count">(${k.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${k.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),w=document.getElementById("import-current-status"),T={totalSheets:m.length,totalRows:m.reduce((k,E)=>k+E.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let k=0;k<m.length;k++){let E=m[k],C=document.getElementById(`step-icon-${E.module}`),$=document.getElementById(`step-status-${E.module}`);C.textContent="\u{1F504}",$.textContent="Mengimport...",w.textContent=`Mengimport ${E.label}...`,v.style.width=`${Math.round(k/m.length*100)}%`;try{let I=await _(`/api/import/${E.module}`,{method:"POST",body:JSON.stringify({rows:E.mapped,onDuplicate:h})});if(I.ok){let D=I.data;T.inserted+=D.inserted||0,T.skipped+=D.skipped||0,T.moduleResults.push({label:E.label,inserted:D.inserted||0,skipped:D.skipped||0,status:"ok"}),C.textContent="\u2705",$.innerHTML=`<span class="badge badge-success">${D.inserted||0} berhasil</span>${D.skipped>0?` <span class="badge badge-neutral">${D.skipped} skip</span>`:""}`}else T.failed++,T.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:I.data?.error}),C.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(I){T.failed++,T.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:I.message}),C.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Qe(150)}v.style.width="100%",w.textContent="Selesai!",await Qe(400),b(T)}function b(h){l("summary");let m=((Date.now()-a)/1e3).toFixed(1),p=h.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${p?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${p?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${h.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${h.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${h.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${h.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${h.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${h.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${m}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${h.moduleResults.map(v=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",l("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Qe(t){return new Promise(e=>setTimeout(e,t))}A();var mt=[],pa=[];async function ma(t){mt=await j(),pa=await te(),P({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:mt}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await _(`/api/sp?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(r,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),le));i(e,"Template_Import_SP")},onImport:async e=>{let i=o=>{if(!o)return null;let s=String(o||"").toLowerCase(),n=mt.find(d=>String(d.label||"").toLowerCase()===s);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let s=String(o).trim();if(/^\d{4,5}$/.test(s)){let d=Number(s);if(d>2e4&&d<99999){let f=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[d,f,g]=n.map(u=>u.trim());if(d.length===4&&f.length<=2&&g.length<=2)return`${d}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&f.length<=2&&d.length<=2)return`${g}-${f.padStart(2,"0")}-${d.padStart(2,"0")}`}return s},r=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:i(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),l=await _("/api/import/sp",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:pa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:mt,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}A();var Oe=[],ua=[];async function ga(t){Oe=await j(),ua=await te(),P({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Oe},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Oe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await _(`/api/mutasi?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),le));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=o=>{if(!o)return null;let s=String(o||"").toLowerCase(),n=Oe.find(d=>String(d.label||"").toLowerCase()===s);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let s=String(o).trim();if(/^\d{4,5}$/.test(s)){let d=Number(s);if(d>2e4&&d<99999){let f=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[d,f,g]=n.map(u=>u.trim());if(d.length===4&&f.length<=2&&g.length<=2)return`${d}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&f.length<=2&&d.length<=2)return`${g}-${f.padStart(2,"0")}-${d.padStart(2,"0")}`}return s},r=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:i(String(o["Cabang Asal"]||"").trim()),to_branch_id:i(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),l=await _("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ua},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Oe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Oe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}A();var Tt=[],ba=[];async function ha(t){Tt=await j(),ba=await Ce(),P({container:t,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Tt}],exportOptions:{moduleName:"overtime_records",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await _(`/api/overtime?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.date||"","Cabang Lembur":o.branch_name||"","Nama Karyawan":o.employee_name||"","Jam Mulai":o.start_time||"","Jam Selesai":o.end_time||"","Istirahat (Jam)":o.break_hours||0,"Total Jam":o.total_hours||0,"Alasan / Ket Lembur":o.reason||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:Tt},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:ba},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",t=>{if(t.target.name==="start_time"||t.target.name==="end_time"||t.target.name==="break_hours"){let e=t.target.closest("form");if(!e)return;let i=e.querySelector('[name="start_time"]'),a=e.querySelector('[name="end_time"]'),r=e.querySelector('[name="break_hours"]'),l=e.querySelector('[name="total_hours"]');if(!i||!a||!r||!l)return;let o=i.value,s=a.value,n=parseFloat(r.value)||0;if(o&&s){let[d,f]=o.split(":").map(Number),[g,u]=s.split(":").map(Number),c=g*60+u-(d*60+f);c<0&&(c+=24*60);let b=c/60-n;b<0&&(b=0),l.value=Number.isInteger(b)?b:b.toFixed(2)}}}));A();async function fa(t){let e=re();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}P({container:t,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let l=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(d=>String(d.id)===String(i));if(!l)return alert("Data tidak ditemukan");let o=d=>{if(!d)return"Tidak ada data";try{return JSON.stringify(JSON.parse(d),null,2)}catch{return d}},s=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(l.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(l.new_data)}</pre>
           </div>
         </div>
       `,{createModal:n}=await Promise.resolve().then(()=>(xe(),yt));n({title:`Detail Audit Log #${i}`,content:s,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}A();xe();R();var Dt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ye={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function ya(t){t.innerHTML=`
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
  `;let e="board",i=null,a=[],r=re(),l=r&&["superadmin","admin","manager","spv"].includes(r.role);_("/api/branches?all=1").then(p=>{p.ok&&p.data?.data&&(a=p.data.data)}).catch(()=>{});async function o(){try{let p=await _("/api/os-migrations/summary");if(!p.ok)throw new Error("Gagal memuat data migrasi");i=p.data?.data||p.data,s(i),d()}catch(p){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${p.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function s(p){let v=document.getElementById("migration-summary-cards");if(!v||!p)return;let w=p.total_branches||0,T=p.vendors?.AEON?.total||0,k=p.vendors?.["Triple S"]?.total||0,E=p.by_status?.completed||0,C=p.by_status?.in_progress||0,$=p.by_status?.scheduled||0,I=w>0?Math.round(E/w*100):0;v.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${w} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${T} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${k} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${I}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${I}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${E}</strong></span>
          <span>\u{1F535} Proses: <strong>${C}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${$}</strong></span>
        </div>
      </div>
    `}function n(){if(!i||!i.matrix)return[];let p=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),v=document.getElementById("filter-vendor")?.value||"",w=document.getElementById("filter-month")?.value||"",T=document.getElementById("filter-status")?.value||"";return i.matrix.filter(k=>!(v&&k.vendor!==v||w&&k.target_month!==w||T&&k.status!==T||p&&!`${k.branch_code} ${k.branch_name} ${k.branch_full_name} ${k.vendor} ${k.notes||""} ${k.pic_vendor||""}`.toLowerCase().includes(p)))}function d(){let p=n(),v=document.getElementById("filter-stats-text");v&&(v.innerHTML=`Menampilkan <strong>${p.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let w=document.getElementById("migration-view-container");e==="board"?f(w,p):g(w,p)}function f(p,v){if(!v.length){p.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let w=["AEON","Triple S"],T=["2026-09","2026-10","2026-11","2026-12"],k="";w.forEach(E=>{let C=v.filter(W=>W.vendor===E);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==E)return;let $=C.length,I=E==="AEON"?"#1E40AF":"#0F766E",D=E==="AEON"?"#EFF6FF":"#F0FDFA",q=E==="AEON"?"#BFDBFE":"#99F6E4";k+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${D}; border-bottom:1px solid ${q}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${I};">${E==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${I}; border:1px solid ${q}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${$} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months 4-Column Grid -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:16px; padding:18px;">
            ${T.map(W=>{let ae=C.filter(Y=>Y.target_month===W);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Dt[W]||W).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${ae.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${ae.length===0?`
                      <div style="text-align:center; padding:30px 10px; color:var(--text-3); font-size:0.8rem; font-style:italic;">
                        Tidak ada jadwal
                      </div>
                    `:ae.map(Y=>{let Ee=Ye[Y.status]||Ye.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${Y.branch_code}.</span>${Y.branch_name}
                            </div>
                            ${l?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${Y.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${Y.id}" data-name="${Y.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${Y.id}" data-status="${Y.status}" style="background:${Ee.bg}; color:${Ee.color}; border:1px solid ${Ee.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${Ee.icon} ${Ee.label} \u25BE
                              </button>
                            </div>

                            ${Y.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${Y.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${Y.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${Y.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${Y.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),p.innerHTML=k,u(p)}function g(p,v){if(!v.length){p.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}p.innerHTML=`
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
                ${l?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${v.map((w,T)=>{let k=Ye[w.status]||Ye.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${T+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${w.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${w.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${w.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${w.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Dt[w.target_month]||w.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${w.id}" data-status="${w.status}" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${k.icon} ${k.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${w.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${w.employee_count>0?`\u{1F464} ${w.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${w.notes||"-"}
                    </td>
                    ${l?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${w.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${w.id}" data-name="${w.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,u(p)}function u(p){p.querySelectorAll(".status-toggle-btn").forEach(v=>{v.addEventListener("click",w=>{w.stopPropagation();let T=v.dataset.id,k=v.dataset.status;c(v,T,k)})}),p.querySelectorAll(".edit-card-btn").forEach(v=>{v.addEventListener("click",()=>{let w=v.dataset.id,T=i?.matrix?.find(k=>String(k.id)===String(w));T&&b(T)})}),p.querySelectorAll(".delete-card-btn").forEach(v=>{v.addEventListener("click",()=>{let w=v.dataset.id,T=v.dataset.name||"Cabang";De(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${T}</strong>?`,async()=>{try{if(!(await _(`/api/os-migrations/${w}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");Q(`Jadwal migrasi ${T} berhasil dihapus`),o()}catch(k){z(k.message)}})})})}function c(p,v,w){document.querySelectorAll(".status-popup-menu").forEach($=>$.remove());let T=p.getBoundingClientRect(),k=document.createElement("div");k.className="status-popup-menu",k.style.cssText=`
      position: fixed;
      top: ${T.bottom+window.scrollY+4}px;
      left: ${T.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let E=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];k.innerHTML=E.map($=>`
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
        color: ${$.key===w?"#2563EB":"var(--text-1)"};
        font-weight: ${$.key===w?"700":"500"};
      ">
        ${$.label} ${$.key===w?"\u2713":""}
      </button>
    `).join("");let C=$=>{!k.contains($.target)&&$.target!==p&&(k.remove(),document.removeEventListener("click",C))};setTimeout(()=>document.addEventListener("click",C),50),k.querySelectorAll(".status-option-item").forEach($=>{$.addEventListener("click",async()=>{let I=$.dataset.key;if(k.remove(),document.removeEventListener("click",C),I!==w)try{if(!(await _(`/api/os-migrations/${v}/status`,{method:"PATCH",body:JSON.stringify({status:I})})).ok)throw new Error("Gagal memperbarui status");Q("Status berhasil diubah"),o()}catch(D){z(D.message)}})}),document.body.appendChild(k)}function b(p=null){let v=!!p,w=v?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",k=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${a.map(E=>`
      <option value="${E.id}" ${p&&p.branch_id===E.id?"selected":""}>
        ${E.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${p?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${p?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${p&&!["AEON","Triple S"].includes(p.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${p?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${p?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${p?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${p?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${p?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${p?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${p?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${p?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${p?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${p?.notes||""}</textarea>
        </div>
      </form>
    `;de({title:w,content:k,confirmText:v?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(E,C)=>{let $=E.querySelector("#migration-form");if(!$.checkValidity()){$.reportValidity();return}let I=new FormData($),D={branch_id:parseInt(I.get("branch_id"),10),vendor:I.get("vendor"),target_month:I.get("target_month"),status:I.get("status"),actual_date:I.get("actual_date")||null,pic_vendor:I.get("pic_vendor")||null,notes:I.get("notes")||null};try{let q=v?`/api/os-migrations/${p.id}`:"/api/os-migrations",ae=await _(q,{method:v?"PUT":"POST",body:JSON.stringify(D)});if(!ae.ok){let M=await ae.json().catch(()=>({}));throw new Error(M.error||"Gagal menyimpan data")}Q(v?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),C(),o()}catch(q){z(q.message)}}})}let h=document.getElementById("view-board-btn"),m=document.getElementById("view-table-btn");h.addEventListener("click",()=>{e="board",h.style.background="#fff",h.classList.remove("btn-ghost"),m.style.background="transparent",m.classList.add("btn-ghost"),d()}),m.addEventListener("click",()=>{e="table",m.style.background="#fff",m.classList.remove("btn-ghost"),h.style.background="transparent",h.classList.add("btn-ghost"),d()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(p=>{let v=document.getElementById(p);v&&(v.addEventListener("input",d),v.addEventListener("change",d))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{b()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let p=n();if(!p.length){z("Tidak ada data untuk diexport");return}let v=p.map((w,T)=>({No:T+1,"Kode Cabang":w.branch_code,"Nama Cabang":w.branch_name,"Cabang Lengkap":w.branch_full_name,"Vendor OS":w.vendor,"Target Bulan":Dt[w.target_month]||w.target_month,Status:Ye[w.status]?.label||w.status,"Tanggal Realisasi":w.actual_date||"","Jml Personil FC":w.employee_count||0,"PIC Vendor":w.pic_vendor||"",Catatan:w.notes||""}));L(v,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),Q("File Excel berhasil diunduh")}),o()}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let i=Math.floor(Number(t)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=t.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let i=e[1],a=parseInt(e[2],10),r=parseInt(e[3],10);if(a>12&&r<=12)return`${i}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),l=a[parseInt(i[1],10)-1];return`${r} ${l} ${i[0]}`}return e};function V(t){return async e=>{if(!je()){_e("/login");return}return t(e)}}var We=null;function Za(){We&&clearInterval(We);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),l=document.getElementById("header-clock-date");r&&(r.textContent=i),l&&(l.textContent=a)};t(),We=setInterval(t,1e3)}async function en(){try{let t=await _("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,r)=>{let l=document.getElementById(a);l&&(l.textContent=r>0?r:"",l.style.display=r>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Re=[];async function tn(){try{let t=await _("/api/dashboard/notifications");if(!t.ok)return;Re=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Re.length>0?"block":"none",e.textContent=Re.length)}catch{}}function an(){if(!Re.length){de({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Re.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;de({title:`Notifikasi (${Re.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function va(){let t=re(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
            <a href="#/relievers" class="nav-item" data-route="/relievers">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 014-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 01-4 4H3"/></svg>
              </span>
              <span class="nav-label">Jadwal Reliefer</span>
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
            <a href="#/one-on-one" class="nav-item" data-route="/one-on-one">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/></svg>
              </span>
              <span class="nav-label">One on One</span>
              <span class="nav-badge badge-warning" id="badge-oo1"></span>
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
                <span class="topbar-greeting-time">${(()=>{let d=new Date().getHours();return d>=4&&d<11?"Selamat Pagi":d>=11&&d<15?"Selamat Siang":d>=15&&d<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||t?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),l=document.getElementById("sidebar-close"),o=()=>{i.classList.add("open"),a.classList.add("show")},s=()=>{i.classList.remove("open"),a.classList.remove("show")};r?.addEventListener("click",o),l?.addEventListener("click",s),a?.addEventListener("click",s),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",s)),t&&t.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(d=>{d.dataset.route!=="/overtime"&&(d.style.display="none")}),document.querySelectorAll(".nav-section").forEach(d=>{Array.from(d.querySelectorAll(".nav-item")).some(g=>g.style.display!=="none")||(d.style.display="none")}));function n(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(u=>{let c=u.dataset.route;u.classList.toggle("active",d===c||c!=="/dashboard"&&d.startsWith(c))});let f=document.getElementById("topbar-title"),g=document.querySelector(".nav-item.active .nav-label");f&&g&&(f.textContent=g.textContent)}window.addEventListener("hashchange",n),n(),Za(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await _("/api/auth/logout",{method:"POST"}),qe(),We&&clearInterval(We),_e("/login")}),en(),tn(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),an()})}async function nn(){J("/login",({main:e})=>Kt(e)),J("/dashboard",V(({main:e})=>At(e))),J("/calendar",V(({main:e})=>ia(e))),J("/employees",V(({main:e,params:i})=>jt(e,i))),J("/contracts",V(({main:e,params:i})=>st(e,i))),J("/sp",V(({main:e})=>ma(e))),J("/mutasi",V(({main:e})=>ga(e))),J("/overtime",V(({main:e})=>ha(e))),J("/timeline",V(({main:e,params:i})=>Ht(e,i))),J("/os-migrations",V(({main:e})=>ya(e))),J("/issues",V(({main:e,params:i})=>Jt(e,i))),J("/one-on-one",V(({main:e,params:i})=>zt(e,i))),J("/training",V(({main:e})=>Ut(e))),J("/relievers",V(({main:e,params:i})=>Vt(e,i))),J("/reports/inspection",V(({main:e})=>Qt(e))),J("/reports/cleaning",V(({main:e})=>Yt(e))),J("/reports/fogging",V(({main:e})=>Wt(e))),J("/reports/basecamp",V(({main:e})=>Xt(e))),J("/reports/supply",V(({main:e})=>$t(e,"supply"))),J("/sop",V(({main:e})=>Zt(e))),J("/checklist",V(({main:e})=>ea(e))),J("/hygiene-standards",V(({main:e})=>ta(e))),J("/forms",V(({main:e})=>$t(e))),J("/users",V(({main:e})=>aa(e))),J("/branches",V(({main:e})=>na(e))),J("/profile",V(({main:e})=>ra(e))),J("/settings/import",V(({main:e})=>ca(e))),J("/audit-logs",V(({main:e})=>fa(e)));let t=je();if(!t&&window.location.hash!=="#/login"&&_e("/login"),t){let e=await _("/api/auth/me");if(e.ok){He(e.data.data),va();let i=re(),a=window.location.hash;i&&i.role==="input_lembur"&&(a==="#/dashboard"||a===""||a==="#/")&&_e("/overtime")}else qe(),_e("/login")}window.addEventListener("fm:login",()=>{va();let e=re();e&&e.role==="input_lembur"?_e("/overtime"):_e("/dashboard")}),Bt()}nn();
